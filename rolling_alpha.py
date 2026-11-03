"""滚动回归计算基金月度 alpha（CAPM / 三因子 / 四因子 / 五因子 / CH3 / CH4 各出一表）。

对基金 j、评价月 t：
    历史回归（含截距）  R_{j,s} = a + Σ_k β_k F_{k,s} + ε,  s = t-24, …, t-1
    当月 alpha          α_{j,t} = R_{j,t} − Σ_k β̂_k F_{k,t}

四个细节：
1. 历史回归包含截距；
2. 计算当月 alpha 时不减历史截距 â（输出中 a_hat 仅作参考）；
3. beta 每个月重新估计（滚动窗口 24 个月）；
4. 估计窗口截止上个月，不包含待评价月份（point in time）。

被解释变量：ExcessRet = AdjustedNAVGrowth / 100 − rf（CSMAR 因子用 rf_monthly，CH 因子用 CH_rf_mon）。

样本规则：
- 样本从面板成立月（基金名单_fundid_date 的 成立时间_earliest）起算，之前的净值
  （如转型前旧基金）不进入回归；
- 成立月起第一个非空增长率视为建仓期月份，剔除；
- 缺月之后的第一行增长率跨越多个月，也剔除；
- CH 因子从 2000-01 起，此前月份截断；
- 只在日历连续的有效月份段内滚动，窗口不跨缺月；段内满 24 个月后，下一个月起出 alpha。

输入：月度净值_rf_{三,四,五}因子.xlsx、月度净值_rf_CH{3,4}.xlsx（不改；CAPM 取三因子表的 RiskPremium1）；
输出：滚动alpha_{CAPM,三因子,四因子,五因子,CH3,CH4}.xlsx。
"""

from pathlib import Path

import pandas as pd
import statsmodels.api as sm
from statsmodels.regression.rolling import RollingOLS

from excel_style import style_sheet

ROOT = Path(__file__).resolve().parent
WINDOW = 24  # 估计窗口月数

# rf：被解释变量减去的无风险利率，须与因子自身口径一致
#   CSMAR 因子 → rf_monthly（三月定存）；鸣石 CH 因子 → CH_rf_mon（mktrf 已减该 rf）
# start：样本起始月（含），之前的月份不进入回归；空串表示不截断
SPECS = [
    {
        # CAPM：只用三因子表中的市场因子
        "label": "CAPM",
        "src": ROOT / "月度净值_rf_三因子.xlsx",
        "out": ROOT / "滚动alpha_CAPM.xlsx",
        "factors": ["RiskPremium1"],
        "rf": "rf_monthly",
        "start": "",
    },
    {
        "label": "三因子",
        "src": ROOT / "月度净值_rf_三因子.xlsx",
        "out": ROOT / "滚动alpha_三因子.xlsx",
        "factors": ["RiskPremium1", "SMB1", "HML1"],
        "rf": "rf_monthly",
        "start": "",
    },
    {
        "label": "四因子",
        "src": ROOT / "月度净值_rf_四因子.xlsx",
        "out": ROOT / "滚动alpha_四因子.xlsx",
        "factors": ["RiskPremium1", "SMB1", "HML1", "UMD1"],
        "rf": "rf_monthly",
        "start": "",
    },
    {
        "label": "五因子",
        "src": ROOT / "月度净值_rf_五因子.xlsx",
        "out": ROOT / "滚动alpha_五因子.xlsx",
        "factors": ["RiskPremium1", "SMB1", "HML1", "RMW1", "CMA1"],
        "rf": "rf_monthly",
        "start": "",
    },
    {
        # 鸣石 CH 因子从 2000-01 起，之前月份截断
        "label": "CH3",
        "src": ROOT / "月度净值_rf_CH3.xlsx",
        "out": ROOT / "滚动alpha_CH3.xlsx",
        "factors": ["mktrf", "SMB", "VMG"],
        "rf": "CH_rf_mon",
        "start": "2000-01",
    },
    {
        "label": "CH4",
        "src": ROOT / "月度净值_rf_CH4.xlsx",
        "out": ROOT / "滚动alpha_CH4.xlsx",
        "factors": ["mktrf", "SMB", "VMG", "PMO"],
        "rf": "CH_rf_mon",
        "start": "2000-01",
    },
]

ID_COLS = ["Symbol", "FundClassID", "FullName", "TradingDate", "year_month"]
FUND_LIST = ROOT / "基金名单_fundid_date.xlsx"  # 取各 fundid 的面板成立日


def load_inception() -> pd.Series:
    """面板各基金成立月（YYYY-MM），索引为 fundid（= 净值主代码份额 Symbol）。

    转型基金的面板成立日是转型日，而月度净值从转型前旧基金开始，
    因此回归样本须从面板成立月起算。
    """
    f = pd.read_excel(FUND_LIST, dtype=str, usecols=["fundid", "成立时间_earliest"])
    f = f.drop_duplicates()
    if f["fundid"].duplicated().any():
        raise SystemExit("基金名单_fundid_date 同一 fundid 有多个成立日")
    return f.set_index("fundid")["成立时间_earliest"].str[:7]


INCEPTION = load_inception()


def mark_sample(df: pd.DataFrame) -> pd.DataFrame:
    """标出可用的月度收益观测（df 须已按 Symbol, year_month 排序）。

    新增列：month_no（月序号）、pre_inception（成立月前有收益的行）、
    sample（成立月起、非建仓月、非缺月后首行、增长率非空）。
    """
    # 月序号：用于判断相邻两行是否为连续的日历月
    df["month_no"] = pd.PeriodIndex(df["year_month"], freq="M").astype("int64")
    gap = df.groupby("Symbol")["month_no"].diff()

    # 面板成立月之前的净值（转型前旧基金等）不用；不在面板的份额不限制
    start_m = df["Symbol"].map(INCEPTION).fillna("")
    after_start = df["year_month"].ge(start_m)
    df["pre_inception"] = ~after_start & df["AdjustedNAVGrowth"].notna()

    # 建仓月 = 面板成立月起第一个非空增长率；不受因子起始月截断影响
    has_growth = df["AdjustedNAVGrowth"].notna() & after_start
    first_ret = has_growth & (has_growth.astype(int).groupby(df["Symbol"]).cumsum() == 1)
    after_gap = gap.gt(1)  # 缺月后的首行，增长率跨多个月
    df["sample"] = has_growth & ~first_ret & ~after_gap
    return df


def load_panel(spec: dict) -> pd.DataFrame:
    """读入月度净值+因子，构造超额收益，并标出可进入回归的有效观测。"""
    factors, rf = spec["factors"], spec["rf"]
    df = pd.read_excel(spec["src"], dtype=str, usecols=[*ID_COLS, "AdjustedNAVGrowth", rf, *factors])
    for col in ID_COLS:
        df[col] = df[col].str.strip()
    for col in ["AdjustedNAVGrowth", rf, *factors]:
        df[col] = pd.to_numeric(df[col], errors="coerce")
    df = df.sort_values(["Symbol", "year_month"], ignore_index=True)

    df["ExcessRet"] = df["AdjustedNAVGrowth"] / 100 - df[rf]
    df = mark_sample(df)
    in_sample = df["year_month"].ge(spec["start"])  # 因子起始月之前截断
    df["valid"] = df["sample"] & df["ExcessRet"].notna() & df[factors].notna().all(axis=1) & in_sample
    return df


def rolling_alpha(df: pd.DataFrame, factors: list[str]) -> pd.DataFrame:
    """在每个“基金 × 连续有效段”内做 24 个月滚动 OLS，按上月窗口的 beta 计算当月 alpha。"""
    v = df.loc[df["valid"]].copy()
    # 同一基金内与上一有效行不是相邻月份，则开新段（被剔除的行与缺月都会切断）
    new_seg = v.groupby("Symbol")["month_no"].diff().ne(1)
    v["seg"] = new_seg.cumsum()

    parts = []
    for _, seg in v.groupby("seg", sort=False):
        if len(seg) <= WINDOW:  # 不足 25 个月，出不了 alpha
            continue
        exog = sm.add_constant(seg[factors], has_constant="add")
        params = RollingOLS(seg["ExcessRet"], exog, window=WINDOW).fit(params_only=True).params
        # RollingOLS 第 i 行用的是 i-23..i（含当月）；后移一行 → 第 i 行用 i-24..i-1
        coef = params.shift(1)

        out = seg[[*ID_COLS, "ExcessRet"]].copy()
        # 当月 alpha：只减 β̂·F_t，不减截距
        out["alpha"] = seg["ExcessRet"] - (seg[factors] * coef[factors]).sum(axis=1, min_count=len(factors))
        out["a_hat"] = coef["const"]
        for f in factors:
            out[f"beta_{f}"] = coef[f]
        out["WindowStart"] = seg["year_month"].shift(WINDOW)
        out["WindowEnd"] = seg["year_month"].shift(1)
        parts.append(out.loc[out["alpha"].notna()])

    return pd.concat(parts, ignore_index=True)


def write_out(res: pd.DataFrame, spec: dict) -> None:
    beta_cols = [f"beta_{f}" for f in spec["factors"]]
    with pd.ExcelWriter(spec["out"], engine="openpyxl") as writer:
        res.to_excel(writer, index=False, sheet_name="alpha")
        style_sheet(
            writer.sheets["alpha"],
            res.columns,
            text_cols=("Symbol", "FundClassID", "TradingDate", "year_month", "WindowStart", "WindowEnd"),
            number_formats={c: "0.000000" for c in ["ExcessRet", "alpha", "a_hat", *beta_cols]},
            widths={
                "Symbol": 10,
                "FundClassID": 14,
                "FullName": 36,
                "TradingDate": 12,
                "year_month": 11,
                "WindowStart": 12,
                "WindowEnd": 12,
                **{c: 14 for c in ["ExcessRet", "alpha", "a_hat", *beta_cols]},
            },
        )


def run(spec: dict) -> None:
    df = load_panel(spec)
    res = rolling_alpha(df, spec["factors"])

    all_funds = set(df["Symbol"])
    gap_funds = set(df.loc[df.groupby("Symbol")["month_no"].diff().gt(1), "Symbol"])
    no_alpha = all_funds - set(res["Symbol"])
    print(
        f"{spec['label']}: 净值行={len(df)} 基金={len(all_funds)} | alpha 行={len(res)} "
        f"基金={res['Symbol'].nunique()} 月份={res['year_month'].min()}..{res['year_month'].max()}"
    )
    pre = df.loc[df["pre_inception"]]
    print(
        f"  无 alpha 基金={len(no_alpha)}（其中有缺月={len(no_alpha & gap_funds)}）；"
        f"有缺月基金={len(gap_funds)}；成立月前净值被剔除的基金={pre['Symbol'].nunique()}"
        f"（{pre.groupby('Symbol').size().to_dict()}）"
    )
    write_out(res, spec)
    print(f"  写出 {spec['out'].name}")


if __name__ == "__main__":
    for spec in SPECS:
        run(spec)
