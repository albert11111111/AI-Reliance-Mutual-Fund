"""把半年原始收益率与 6 个模型的半年平均 alpha 接到 fundid × date 面板。

半年归属：1–6 月 → YYYY0630，7–12 月 → YYYY1231；半年 6 个月齐全才填值，否则留空。
- raw_ret：月度净值 AdjustedNAVGrowth 的半年几何平均月收益
      raw_ret = [Π_{6 个月}(1 + AdjustedNAVGrowth/100)]^(1/6) − 1
  样本规则与滚动回归一致（rolling_alpha.mark_sample）：面板成立月起算、剔除建仓月与
  缺月后首行；但不需要 24 个月估计窗口。
- *_alpha：滚动 alpha 的半年算术平均（月度 alpha 均值，小数制，不年化）。

合并键：面板 fundid = 月度净值 Symbol。
说明书中 Symbol 是份额代码（非主代码），因此先用份额类别信息表核对：
净值份额 (Symbol, FundClassID) 对应的 FundID 必须包含面板的 基金ID。

输入：基金名单_fundid_date_TotalEquity_Holder_BuySell_Fee_Manager.xlsx、月度净值.xlsx、
      滚动alpha_*.xlsx（均不改）。
输出：基金名单_fundid_date_TotalEquity_Holder_BuySell_Fee_Manager_Alpha.xlsx。
"""

import zipfile
from pathlib import Path

import numpy as np
import pandas as pd

from excel_style import style_sheet
from merge_fund_manager import TEXT_COLS
from rolling_alpha import mark_sample

ROOT = Path(__file__).resolve().parent
PANEL = ROOT / "基金名单_fundid_date_TotalEquity_Holder_BuySell_Fee_Manager.xlsx"
NAV = ROOT / "月度净值.xlsx"
OUT = ROOT / "基金名单_fundid_date_TotalEquity_Holder_BuySell_Fee_Manager_Alpha.xlsx"
UNIT_CLASS_DIR = ROOT / "CSMAR原始数据" / "基金份额信息"

# (面板列名, 月度 alpha 文件)
ALPHAS = [
    ("capm_alpha", "滚动alpha_CAPM.xlsx"),
    ("ff3_alpha", "滚动alpha_三因子.xlsx"),
    ("car4_alpha", "滚动alpha_四因子.xlsx"),
    ("ff5_alpha", "滚动alpha_五因子.xlsx"),
    ("ch3_alpha", "滚动alpha_CH3.xlsx"),
    ("ch4_alpha", "滚动alpha_CH4.xlsx"),
]
MONTHS_PER_HALF = 6


def half_date(year_month: pd.Series) -> pd.Series:
    """YYYY-MM → 所属半年末 YYYY0630 / YYYY1231。"""
    month = year_month.str[5:7].astype(int)
    return year_month.str[:4] + np.where(month <= 6, "0630", "1231")


def half_year_raw_ret() -> pd.DataFrame:
    """每个 (Symbol, 半年) 的有效月数与几何平均月收益；不足 6 个月置空。"""
    nav = pd.read_excel(NAV, dtype=str, usecols=["TradingDate", "Symbol", "AdjustedNAVGrowth"]).iloc[2:]
    nav = nav.apply(lambda s: s.str.strip())
    nav["year_month"] = nav["TradingDate"].str[:7]
    nav["AdjustedNAVGrowth"] = pd.to_numeric(nav["AdjustedNAVGrowth"], errors="coerce")
    nav = mark_sample(nav.sort_values(["Symbol", "year_month"], ignore_index=True))

    s = nav.loc[nav["sample"]].copy()
    s["date"] = half_date(s["year_month"])
    s["gross"] = 1 + s["AdjustedNAVGrowth"] / 100
    g = s.groupby(["Symbol", "date"])["gross"].agg(["prod", "count"])
    g["value"] = (g["prod"] ** (1 / MONTHS_PER_HALF) - 1).where(g["count"].eq(MONTHS_PER_HALF))
    return g[["value", "count"]]


def load_alpha(path: Path) -> pd.DataFrame:
    a = pd.read_excel(
        path,
        dtype={"Symbol": str, "FundClassID": str, "year_month": str},
        usecols=["Symbol", "FundClassID", "year_month", "alpha"],
    )
    a["date"] = half_date(a["year_month"])
    return a


def half_year_alpha(a: pd.DataFrame) -> pd.DataFrame:
    """每个 (Symbol, 半年) 的 alpha 月数与均值；不足 6 个月置空。"""
    g = a.groupby(["Symbol", "date"])["alpha"].agg(["mean", "count"])
    g["value"] = g["mean"].where(g["count"].eq(MONTHS_PER_HALF))
    return g[["value", "count"]]


def check_share_class(panel: pd.DataFrame, alpha: pd.DataFrame) -> None:
    """核对面板 (fundid, 基金ID) 与净值份额 (Symbol, FundClassID) 指向同一基金主体。"""
    zpath = next(UNIT_CLASS_DIR.glob("*.zip"))
    with zipfile.ZipFile(zpath) as z:
        unit = pd.read_csv(
            z.open("Fund_EVL_UnitClassInfo.csv"), dtype=str, usecols=["Symbol", "FundClassID", "FundID"]
        )
    unit = unit.apply(lambda s: s.str.strip())
    # 同一份额类别可能因转型有多段历史 → 多个 FundID，任一与面板一致即可
    share = alpha[["Symbol", "FundClassID"]].drop_duplicates().merge(unit, on=["Symbol", "FundClassID"])
    pairs = panel[["fundid", "基金ID"]].drop_duplicates()
    pairs = pairs.loc[pairs["fundid"].isin(alpha["Symbol"])]
    hit = pairs.merge(share, left_on=["fundid", "基金ID"], right_on=["Symbol", "FundID"], how="left", indicator=True)
    bad = hit.loc[hit["_merge"].eq("left_only"), ["fundid", "基金ID"]]
    if not bad.empty:
        raise SystemExit(f"净值份额与面板 基金ID 不一致：\n{bad.to_string(index=False)}")
    print(f"份额核对通过：{len(pairs)} 个有 alpha 的 fundid，净值份额 FundID 均与面板 基金ID 一致")


def merge_half(out: pd.DataFrame, g: pd.DataFrame, name: str) -> pd.DataFrame:
    """把半年值左连接到面板，并打印有值与留空原因。"""
    n = len(out)
    out = out.merge(
        g["value"].rename(name), left_on=["fundid", "date"], right_index=True, how="left", validate="one_to_one"
    )
    if len(out) != n:
        raise SystemExit(f"{name} 合并后行数变化")

    # 留空原因：6 个月不全 / 该半年无数据 / 整只基金无数据
    cnt = out[["fundid", "date"]].merge(g["count"], left_on=["fundid", "date"], right_index=True, how="left")["count"]
    fund_has = out["fundid"].isin(g.index.get_level_values("Symbol"))
    print(
        f"{name}: 有值={int(out[name].notna().sum())} 留空={int(out[name].isna().sum())}"
        f"（6个月不全={int(cnt.between(1, 5).sum())} 该半年无数据={int((cnt.isna() & fund_has).sum())}"
        f" 整只基金无数据={int((~fund_has).sum())}）"
    )
    return out


def build() -> None:
    panel = pd.read_excel(PANEL, dtype={c: str for c in TEXT_COLS})
    out = merge_half(panel, half_year_raw_ret(), "raw_ret")

    for i, (name, fname) in enumerate(ALPHAS):
        a = load_alpha(ROOT / fname)
        if i == 0:
            check_share_class(panel, a)
        out = merge_half(out, half_year_alpha(a), name)

    new_cols = ["raw_ret", *(name for name, _ in ALPHAS)]
    with pd.ExcelWriter(OUT, engine="openpyxl") as writer:
        out.to_excel(writer, index=False, sheet_name="fundid_date")
        style_sheet(
            writer.sheets["fundid_date"],
            out.columns,
            text_cols=TEXT_COLS,
            number_formats={
                "TotalEquity": "0.00",
                "InstitutionRatio": "0.0000",
                "IndividualRatio": "0.0000",
                "PractitionersRatio": "0.0000",
                "BuyValue": "0.00",
                "SellValue": "0.00",
                "TotalOperatingCost": "0.00",
                "Team": "0",
                "Female": "0.0000",
                "Phd": "0.0000",
                "midAge": "0.00",
                "largeAge": "0.00",
                **{c: "0.000000" for c in new_cols},
            },
            widths={c: 12 for c in new_cols},
        )
    print(f"面板行={len(out)}，写出 {OUT.name}")


if __name__ == "__main__":
    build()
