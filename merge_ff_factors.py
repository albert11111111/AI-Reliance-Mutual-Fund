"""把 CSMAR Fama-French / Carhart 月度因子接到月度净值_rf。

口径（说明书）：
- 市场：P9714 沪深 A 股和创业板和科创板（沪深全 A 市场）
- 加权：流通市值加权（字段后缀 1：RiskPremium1、SMB1、…）
- 五因子投资组合：Portfolios=1（2×3 划分，与经典三因子一致）

输入：月度净值_rf.xlsx（不改）；三个因子源表（不改）。
按 year_month = TradingMonth 左连接，各出一文件，供后续滚动回归。
"""

from pathlib import Path

import pandas as pd

from excel_style import style_sheet

ROOT = Path(__file__).resolve().parent
NAV = ROOT / "月度净值_rf.xlsx"
MARKET = "P9714"  # 沪深 A + 创业板 + 科创板
PORTFOLIO_5 = "1"  # 2*3

SPECS = [
    {
        "label": "三因子",
        "src": ROOT / "三因子模型指标(月)133509346(仅供上海交通大学使用)" / "STK_MKT_THRFACMONTH.xlsx",
        "out": ROOT / "月度净值_rf_三因子.xlsx",
        "factors": ["RiskPremium1", "SMB1", "HML1"],
        "extra_filters": {},
    },
    {
        "label": "四因子",
        "src": ROOT / "Carhart四因子080749357(仅供上海交通大学使用)" / "STK_MKT_CARHARTFOURFACTORS.xlsx",
        "out": ROOT / "月度净值_rf_四因子.xlsx",
        "factors": ["RiskPremium1", "SMB1", "HML1", "UMD1"],
        "extra_filters": {},
    },
    {
        "label": "五因子",
        "src": ROOT / "五因子模型指标(月)133455709(仅供上海交通大学使用)" / "STK_MKT_FIVEFACMONTH.xlsx",
        "out": ROOT / "月度净值_rf_五因子.xlsx",
        "factors": ["RiskPremium1", "SMB1", "HML1", "RMW1", "CMA1"],
        "extra_filters": {"Portfolios": PORTFOLIO_5},
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
TEXT_COLS = ("TradingDate", "Symbol", "FundClassID", "year_month", "FactorMarkettypeID")


def load_factors(path: Path, factors: list[str], extra_filters: dict) -> pd.DataFrame:
    raw = pd.read_excel(path, dtype=str).iloc[2:].copy()
    for col in raw.columns:
        raw[col] = raw[col].str.strip()
    out = raw.loc[raw["MarkettypeID"].eq(MARKET)].copy()
    for col, val in extra_filters.items():
        out = out.loc[out[col].eq(val)].copy()
    if out.empty:
        raise SystemExit(f"{path.name} 在 MarkettypeID={MARKET} 下无数据")
    if out["TradingMonth"].duplicated().any():
        raise SystemExit(f"{path.name} 同一 TradingMonth 有多条（检查 Portfolios 过滤）")

    keep = ["TradingMonth", "MarkettypeID", *factors]
    out = out.loc[:, keep].rename(
        columns={
            "TradingMonth": "year_month",
            "MarkettypeID": "FactorMarkettypeID",
        }
    )
    for col in factors:
        out[col] = pd.to_numeric(out[col], errors="coerce")
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
    fac = load_factors(spec["src"], spec["factors"], spec["extra_filters"])
    n = len(nav)
    out = nav.merge(fac, on="year_month", how="left", validate="many_to_one")
    if len(out) != n:
        raise SystemExit(f"{spec['label']} 合并后行数变化")

    # 因子列放在利率后面，便于回归取用
    base_cols = [c for c in nav.columns]
    extra = ["FactorMarkettypeID", *spec["factors"]]
    out = out[base_cols + extra]

    matched = out[spec["factors"]].notna().all(axis=1)
    any_miss = out[spec["factors"]].isna().any(axis=1)
    print(
        f"{spec['label']}: rows={len(out)} 因子齐全={int(matched.sum())} "
        f"有缺失={int(any_miss.sum())} 因子月={fac['year_month'].min()}..{fac['year_month'].max()}"
    )
    if any_miss.any():
        # 滚动回归前需知情：通常是早期 Carhart 动量未成形等
        miss_by_month = (
            out.loc[any_miss]
            .groupby("year_month", sort=True)[spec["factors"]]
            .apply(lambda d: int(d.isna().any(axis=1).sum()))
        )
        print(f"  有缺失的 year_month 数={miss_by_month.shape[0]} 样例:")
        print(miss_by_month.head(8).to_string())

    with pd.ExcelWriter(spec["out"], engine="openpyxl") as writer:
        out.to_excel(writer, index=False, sheet_name="nav_ff")
        style_sheet(
            writer.sheets["nav_ff"],
            out.columns,
            text_cols=TEXT_COLS,
            number_formats={
                **{c: "0.0000" for c in ("NAV", "AccumulativeNAV", "AdjustedNAV") if c in out.columns},
                **{
                    c: "0.000000"
                    for c in ("AccumulativeNAVGrowth", "AdjustedNAVGrowth", "rf_monthly")
                    if c in out.columns
                },
                **{c: "0.000000" for c in spec["factors"]},
            },
            widths={
                "TradingDate": 12,
                "Symbol": 10,
                "FullName": 36,
                "year_month": 12,
                "FactorMarkettypeID": 16,
                **{c: 14 for c in spec["factors"]},
            },
        )
    print(f"  写出 {spec['out'].name}")


def build() -> None:
    print(f"口径 MarkettypeID={MARKET} 流通市值加权(*1)；五因子 Portfolios={PORTFOLIO_5}")
    nav = load_nav()
    print(f"净值行={len(nav)} year_month={nav['year_month'].min()}..{nav['year_month'].max()}")
    for spec in SPECS:
        merge_one(nav, spec)


if __name__ == "__main__":
    build()
