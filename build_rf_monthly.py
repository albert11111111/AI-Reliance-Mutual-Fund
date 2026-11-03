"""把整存整取三月期年化利率换成月度无风险利率，并接到月度净值。

源表：无风险利率.xlsx（Irfid02，单位 %）、月度净值.xlsx。均不修改。
公式：rf_月 = (1 + q/100)^(1/12) - 1。
变更日之间的日历月沿用上一档利率（point-in-time：取该月末及以前最近一次变更）。

与净值对齐：净值 TradingDate 多为该月最后一个交易日，不一定等于日历月末，
故用 TradingDate 的日历年月 year_month 左连接到月度利率，不用精确日匹配。

写出：
- 无风险利率_月度.xlsx
- 月度净值_rf.xlsx
"""

from pathlib import Path

import pandas as pd

from excel_style import style_sheet

ROOT = Path(__file__).resolve().parent
RF_SRC = ROOT / "无风险利率.xlsx"
NAV_SRC = ROOT / "月度净值.xlsx"
RF_OUT = ROOT / "无风险利率_月度.xlsx"
NAV_OUT = ROOT / "月度净值_rf.xlsx"


def load_rate_changes() -> pd.DataFrame:
    raw = pd.read_excel(RF_SRC, dtype=str).iloc[2:].copy()
    for col in raw.columns:
        raw[col] = raw[col].str.strip()
    out = pd.DataFrame(
        {
            "RateChangeDate": pd.to_datetime(raw["Datesgn"], errors="coerce"),
            "q_annual_pct": pd.to_numeric(raw["Irfid02"], errors="coerce"),
        }
    )
    if out["RateChangeDate"].isna().any() or out["q_annual_pct"].isna().any():
        raise SystemExit("无风险利率存在无法解析的日期或利率")
    if out["RateChangeDate"].duplicated().any():
        raise SystemExit("无风险利率同一变更日有多条记录")
    return out.sort_values("RateChangeDate").reset_index(drop=True)


def build_monthly_rf(changes: pd.DataFrame, end_month: pd.Timestamp) -> pd.DataFrame:
    """扩成连续日历月；每月用月末及以前最近一次变更利率。"""
    start_period = changes["RateChangeDate"].iloc[0].to_period("M")
    end_period = pd.Timestamp(end_month).to_period("M")
    periods = pd.period_range(start_period, end_period, freq="M")
    monthly = pd.DataFrame(
        {
            "year_month": periods.astype(str),
            "month_end": periods.to_timestamp(how="end").normalize(),
        }
    )

    ch = changes.rename(columns={"RateChangeDate": "asof_date"}).sort_values("asof_date")
    merged = pd.merge_asof(
        monthly.sort_values("month_end"),
        ch,
        left_on="month_end",
        right_on="asof_date",
        direction="backward",
    )
    if merged["q_annual_pct"].isna().any():
        raise SystemExit("月度利率存在未能向前填充的月份")

    merged["rf_monthly"] = (1.0 + merged["q_annual_pct"] / 100.0) ** (1.0 / 12.0) - 1.0
    return merged.rename(columns={"asof_date": "RateChangeDate"})[
        ["year_month", "month_end", "RateChangeDate", "q_annual_pct", "rf_monthly"]
    ]


def load_nav() -> pd.DataFrame:
    raw = pd.read_excel(NAV_SRC, dtype=str).iloc[2:].copy()
    for col in raw.columns:
        raw[col] = raw[col].str.strip()
    out = raw.copy()
    traded = pd.to_datetime(out["TradingDate"], errors="coerce")
    if traded.isna().any():
        raise SystemExit("月度净值存在无法解析的 TradingDate")
    out["TradingDate"] = traded.dt.strftime("%Y-%m-%d")
    out["year_month"] = traded.dt.strftime("%Y-%m")
    for col in ("NAV", "AccumulativeNAV", "AdjustedNAV", "AccumulativeNAVGrowth", "AdjustedNAVGrowth"):
        out[col] = pd.to_numeric(out[col], errors="coerce")
    return out


def build() -> None:
    changes = load_rate_changes()
    nav = load_nav()
    end_month = max(pd.to_datetime(nav["TradingDate"]).max(), changes["RateChangeDate"].max())
    rf_m = build_monthly_rf(changes, end_month)

    sample = rf_m.loc[rf_m["year_month"] == "2015-10"].iloc[0]
    expect = (1.0 + 1.1 / 100.0) ** (1.0 / 12.0) - 1.0
    if abs(float(sample["rf_monthly"]) - expect) > 1e-15:
        raise SystemExit("2015-10 月度利率换算不对")
    gap = rf_m.loc[rf_m["year_month"] == "2016-01"].iloc[0]
    if float(gap["q_annual_pct"]) != 1.1:
        raise SystemExit("间隔月未沿用上一档利率")

    with pd.ExcelWriter(RF_OUT, engine="openpyxl") as writer:
        rf_m.to_excel(writer, index=False, sheet_name="rf_monthly")
        style_sheet(
            writer.sheets["rf_monthly"],
            rf_m.columns,
            text_cols=("year_month",),
            number_formats={
                "q_annual_pct": "0.00",
                "rf_monthly": "0.00000000",
            },
            widths={
                "year_month": 12,
                "month_end": 12,
                "RateChangeDate": 14,
                "q_annual_pct": 14,
                "rf_monthly": 14,
            },
        )

    merged = nav.merge(
        rf_m[["year_month", "RateChangeDate", "q_annual_pct", "rf_monthly"]],
        on="year_month",
        how="left",
    )
    if len(merged) != len(nav):
        raise SystemExit("净值合并后行数变化")
    miss = merged["rf_monthly"].isna()
    if miss.any():
        print("警告：净值有未匹配利率的行", int(miss.sum()))
        print(merged.loc[miss, ["TradingDate", "Symbol", "year_month"]].head().to_string(index=False))

    nav_cols = [c for c in nav.columns if c != "year_month"]
    out = merged[nav_cols + ["year_month", "RateChangeDate", "q_annual_pct", "rf_monthly"]]
    out = out.sort_values(["Symbol", "TradingDate"], kind="mergesort").reset_index(drop=True)

    with pd.ExcelWriter(NAV_OUT, engine="openpyxl") as writer:
        out.to_excel(writer, index=False, sheet_name="nav_rf")
        style_sheet(
            writer.sheets["nav_rf"],
            out.columns,
            text_cols=("TradingDate", "Symbol", "FundClassID", "year_month"),
            number_formats={
                "NAV": "0.0000",
                "AccumulativeNAV": "0.0000",
                "AdjustedNAV": "0.0000",
                "AccumulativeNAVGrowth": "0.000000",
                "AdjustedNAVGrowth": "0.000000",
                "q_annual_pct": "0.00",
                "rf_monthly": "0.00000000",
            },
            widths={
                "TradingDate": 12,
                "Symbol": 10,
                "FullName": 36,
                "year_month": 12,
                "rf_monthly": 14,
            },
        )

    print(
        f"利率变更={len(changes)} 月度行={len(rf_m)} "
        f"year_month={rf_m['year_month'].iloc[0]}..{rf_m['year_month'].iloc[-1]}"
    )
    print(
        f"净值行={len(out)} 匹配利率={int(out['rf_monthly'].notna().sum())} "
        f"未匹配={int(out['rf_monthly'].isna().sum())}"
    )
    print(f"写出 {RF_OUT.name} ; {NAV_OUT.name}")


if __name__ == "__main__":
    build()
