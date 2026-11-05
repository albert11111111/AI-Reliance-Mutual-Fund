"""第一步：生成半年基础面板、月度无风险利率及五组回归输入。"""
from datetime import date, datetime
from pathlib import Path
import pandas as pd

ROOT = Path(__file__).resolve().parent
RAW = ROOT / "原始数据"
WORK = ROOT / "中间数据"
RF_SRC = RAW / "无风险利率.xlsx"
NAV_SRC = RAW / "月度净值.xlsx"
MARKET = "P9714"
PORTFOLIO_5 = "1"

# %% 1. 基金名单与半年基础面板
SAMPLE_END = date(2026, 6, 30)


KEEP = [
    "fundid",
    "short_name",
    "基金类型ifind",
    "基金经理",
    "成立时间_earliest",
    "成立时间_latest",
    "到期日",
    "基金ID",
    "基金管理公司ID",
    "基金管理公司名称",
    "发生转型代码重复",
    "发生过转型",
]


def parse_day(value):
    if value is None or (isinstance(value, float) and pd.isna(value)) or pd.isna(value):
        return None
    text = str(value).strip()
    if not text or text.lower() in {"nan", "none", "0"}:
        return None
    return datetime.strptime(text[:10], "%Y-%m-%d").date()


def semi_dates(inception: date, maturity: date | None):
    """起始日之后、到期日之前（含到期日当天若恰为半年末）的 0630/1231，不超过 SAMPLE_END。"""
    cap = SAMPLE_END if maturity is None or maturity > SAMPLE_END else maturity
    out = []
    year = inception.year
    while year <= cap.year:
        for month, day in ((6, 30), (12, 31)):
            dt = date(year, month, day)
            if dt <= inception:
                continue
            if dt > cap:
                return out
            out.append(f"{dt:%Y%m%d}")
        year += 1
    return out


def build_panel():
    src = pd.read_excel(RAW / "基金名单.xlsx", dtype=str)
    src = src.sort_values(["fundid", "成立时间_earliest", "成立时间_latest"],
                          ascending=[True, True, False], kind="mergesort").drop_duplicates("fundid")
    rows = []
    for _, rec in src.iterrows():
        for end in semi_dates(parse_day(rec["成立时间_earliest"]), parse_day(rec["到期日"])):
            rows.append({"fundid": rec["fundid"], "date": end,
                         **{c: rec[c] for c in KEEP if c != "fundid"}})
    panel = pd.DataFrame(rows).sort_values(["fundid", "date"], kind="mergesort").reset_index(drop=True)
    for c in ("发生转型代码重复", "发生过转型"):
        panel[c] = panel[c].astype(int)
    if panel.duplicated(["fundid", "date"]).any():
        raise ValueError("半年面板键不唯一")
    panel.to_excel(WORK / "基金名单_fundid_date.xlsx", index=False, sheet_name="fundid_date")
    print(f"基础面板：{len(panel):,} 行")

# %% 2. 日历月无风险利率；保持旧版月末向后匹配口径
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

# %% 3. CSMAR 与 CH 因子；CH3、CH4 使用各自的 SMB
FF_SPECS = [
    {
        "label": "三因子",
        "src": RAW / "三因子模型指标(月)133509346(仅供上海交通大学使用)" / "STK_MKT_THRFACMONTH.xlsx",
        "out": WORK / "月度净值_rf_三因子.xlsx",
        "factors": ["RiskPremium1", "SMB1", "HML1"],
        "extra_filters": {},
    },
    {
        "label": "四因子",
        "src": RAW / "Carhart四因子080749357(仅供上海交通大学使用)" / "STK_MKT_CARHARTFOURFACTORS.xlsx",
        "out": WORK / "月度净值_rf_四因子.xlsx",
        "factors": ["RiskPremium1", "SMB1", "HML1", "UMD1"],
        "extra_filters": {},
    },
    {
        "label": "五因子",
        "src": RAW / "五因子模型指标(月)133455709(仅供上海交通大学使用)" / "STK_MKT_FIVEFACMONTH.xlsx",
        "out": WORK / "月度净值_rf_五因子.xlsx",
        "factors": ["RiskPremium1", "SMB1", "HML1", "RMW1", "CMA1"],
        "extra_filters": {"Portfolios": PORTFOLIO_5},
    },
]

CH_DIR = RAW / "CH因子_鸣石_20261103"
CH_SPECS = [
    {
        "label": "CH3",
        "src": CH_DIR / "CH3_factors_monthly_202608.xlsx",
        "out": WORK / "月度净值_rf_CH3.xlsx",
        "factors": ["mktrf", "SMB", "VMG"],
    },
    {
        "label": "CH4",
        "src": CH_DIR / "CH4_factors_monthly_202608.xlsx",
        "out": WORK / "月度净值_rf_CH4.xlsx",
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

# %% 4. 顺序执行；保留 Excel 落盘精度，以复现原有链路

def main():
    WORK.mkdir(exist_ok=True)
    build_panel()
    changes, nav = load_rate_changes(), load_nav()
    rf = build_monthly_rf(changes, max(pd.to_datetime(nav["TradingDate"]).max(), changes["RateChangeDate"].max()))
    rf.to_excel(WORK / "无风险利率_月度.xlsx", index=False, sheet_name="rf_monthly")
    nav = nav.merge(rf[["year_month", "RateChangeDate", "q_annual_pct", "rf_monthly"]],
                    on="year_month", how="left", validate="many_to_one")
    nav = nav.sort_values(["Symbol", "TradingDate"], kind="mergesort").reset_index(drop=True)
    nav.to_excel(WORK / "月度净值_rf.xlsx", index=False, sheet_name="nav_rf")
    nav = pd.read_excel(WORK / "月度净值_rf.xlsx", dtype=str)
    nav = nav.apply(lambda s: s.str.strip())
    for c in NAV_NUM:
        nav[c] = pd.to_numeric(nav[c], errors="coerce")
    for spec in [*FF_SPECS, *CH_SPECS]:
        fac = (load_factors(spec["src"], spec["factors"], spec["extra_filters"])
               if "extra_filters" in spec else load_ch(spec["src"], spec["factors"]))
        out = nav.merge(fac, on="year_month", how="left", validate="many_to_one")
        out.to_excel(spec["out"], index=False, sheet_name="nav_ff" if "extra_filters" in spec else "nav_ch")
        print(f"{spec['label']}：{len(out):,} 行，因子不全 {out[spec['factors']].isna().any(axis=1).sum():,}")


if __name__ == "__main__":
    main()
