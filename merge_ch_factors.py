"""把鸣石 CH3 / CH4 月度因子接到月度净值_rf。

源表：CH因子_鸣石_20261103/CH3_factors_monthly_202608.xlsx、
      CH4_factors_monthly_202608.xlsx（不改）。
输入：月度净值_rf.xlsx（不改）。
合并键：净值 year_month = 因子 mnthdt 的 YYYY-MM。

数值为小数制（0.01=1%）。mktrf 已是市场减无风险利率，回归时不要再减。
CH 自带 rf_mon（一年期存款折月）与面板 rf_monthly（三月定存）不同，两列都保留。
CH3 与 CH4 的 SMB 构造不同，各自文件用各自 SMB。

写出：月度净值_rf_CH3.xlsx、月度净值_rf_CH4.xlsx。
"""

from pathlib import Path

import pandas as pd

from excel_style import style_sheet

ROOT = Path(__file__).resolve().parent
NAV = ROOT / "月度净值_rf.xlsx"
CH_DIR = ROOT / "CH因子_鸣石_20261103"

SPECS = [
    {
        "label": "CH3",
        "src": CH_DIR / "CH3_factors_monthly_202608.xlsx",
        "out": ROOT / "月度净值_rf_CH3.xlsx",
        "factors": ["mktrf", "SMB", "VMG"],
    },
    {
        "label": "CH4",
        "src": CH_DIR / "CH4_factors_monthly_202608.xlsx",
        "out": ROOT / "月度净值_rf_CH4.xlsx",
        "factors": ["mktrf", "SMB", "VMG", "PMO"],
    },
]

NAV_NUM = (
    "NAV",
    "AccumulativeNAV",
    "AdjustedNAV",
    "AccumulativeNAVGrowth",
    "AdjustedNAVGrowth",
    "q_annual_pct",
    "rf_monthly",
)
TEXT_COLS = ("TradingDate", "Symbol", "FundClassID", "year_month")


def load_ch(path: Path, factors: list[str]) -> pd.DataFrame:
    raw = pd.read_excel(path, dtype=str)
    for col in raw.columns:
        raw[col] = raw[col].str.strip()
    out = raw.copy()
    out["year_month"] = pd.to_datetime(out["mnthdt"], format="%Y%m%d", errors="coerce").dt.strftime("%Y-%m")
    if out["year_month"].isna().any():
        raise SystemExit(f"{path.name} 存在无法解析的 mnthdt")
    if out["year_month"].duplicated().any():
        raise SystemExit(f"{path.name} 同一 year_month 有多条")

    keep = ["year_month", "rf_mon", *factors]
    out = out.loc[:, keep].rename(columns={"rf_mon": "CH_rf_mon"})
    for col in ["CH_rf_mon", *factors]:
        out[col] = pd.to_numeric(out[col], errors="coerce")
    if out[["CH_rf_mon", *factors]].isna().any().any():
        raise SystemExit(f"{path.name} 因子存在缺失")
    return out


def load_nav() -> pd.DataFrame:
    nav = pd.read_excel(NAV, dtype=str)
    for col in nav.columns:
        if nav[col].dtype == object:
            nav[col] = nav[col].str.strip()
    if "year_month" not in nav.columns:
        raise SystemExit("月度净值_rf 缺少 year_month")
    for col in NAV_NUM:
        if col in nav.columns:
            nav[col] = pd.to_numeric(nav[col], errors="coerce")
    return nav


def merge_one(nav: pd.DataFrame, spec: dict) -> None:
    fac = load_ch(spec["src"], spec["factors"])
    n = len(nav)
    out = nav.merge(fac, on="year_month", how="left", validate="many_to_one")
    if len(out) != n:
        raise SystemExit(f"{spec['label']} 合并后行数变化")

    fac_cols = ["CH_rf_mon", *spec["factors"]]
    out = out[[c for c in nav.columns] + fac_cols]

    matched = out[fac_cols].notna().all(axis=1)
    print(
        f"{spec['label']}: rows={len(out)} 因子齐全={int(matched.sum())} "
        f"未匹配={int((~matched).sum())} 因子月={fac['year_month'].min()}..{fac['year_month'].max()}"
    )
    if (~matched).any():
        miss_ym = sorted(out.loc[~matched, "year_month"].unique())
        print(f"  未匹配 year_month 数={len(miss_ym)} 起止 {miss_ym[0]}..{miss_ym[-1]}")

    with pd.ExcelWriter(spec["out"], engine="openpyxl") as writer:
        out.to_excel(writer, index=False, sheet_name="nav_ch")
        style_sheet(
            writer.sheets["nav_ch"],
            out.columns,
            text_cols=TEXT_COLS,
            number_formats={
                **{c: "0.0000" for c in ("NAV", "AccumulativeNAV", "AdjustedNAV") if c in out.columns},
                **{
                    c: "0.000000"
                    for c in ("AccumulativeNAVGrowth", "AdjustedNAVGrowth", "rf_monthly", "CH_rf_mon")
                    if c in out.columns
                },
                **{c: "0.000000" for c in spec["factors"]},
            },
            widths={
                "TradingDate": 12,
                "Symbol": 10,
                "FullName": 36,
                "year_month": 12,
                "CH_rf_mon": 12,
                **{c: 12 for c in spec["factors"]},
            },
        )
    print(f"  写出 {spec['out'].name}")


def build() -> None:
    nav = load_nav()
    print(f"净值行={len(nav)} year_month={nav['year_month'].min()}..{nav['year_month'].max()}")
    for spec in SPECS:
        merge_one(nav, spec)


if __name__ == "__main__":
    build()
