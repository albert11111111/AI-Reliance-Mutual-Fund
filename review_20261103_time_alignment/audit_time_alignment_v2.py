"""只读核查时间窗口；仅在本目录写审查结果，不重跑或覆盖生产文件。

运行：fund-manager 环境的 python audit_time_alignment.py。
核对现有输出与独立计算，并保存异常数量和可追溯样例。
"""

import json
from pathlib import Path
import sys

import numpy as np
import pandas as pd
import statsmodels.api as sm

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))
sys.dont_write_bytecode = True
import rolling_alpha as ra
import merge_alpha as ma
import merge_fee as mf
import merge_buy_sell_zero as mb

HERE = Path(__file__).resolve().parent
RESULT = {}


def record(name, value):
    """逐项落盘，长核查中断时也保留已经完成的证据。"""
    RESULT[name] = value
    (HERE / "audit_results.json").write_text(
        json.dumps(RESULT, ensure_ascii=False, indent=2, default=str), encoding="utf-8"
    )
    print(name, json.dumps(value, ensure_ascii=False, default=str), flush=True)


def compare(panel, expected, col):
    """缺失与数值分别核对，避免把两边都缺失误报为数值匹配。"""
    joined = panel[["fundid", "date", col]].merge(
        expected.rename("expected"), left_on=["fundid", "date"],
        right_index=True, how="left", validate="one_to_one",
    )
    actual, expect = joined[col], joined["expected"]
    both = actual.notna() & expect.notna()
    return {
        "有值行数": int(actual.notna().sum()),
        "缺失不一致行数": int((actual.isna() != expect.isna()).sum()),
        "数值不一致行数": int((~np.isclose(actual[both], expect[both], atol=1e-12, rtol=1e-10)).sum()),
        "最大绝对误差": float((actual[both] - expect[both]).abs().max()),
    }


def main():
    panel = pd.read_excel(ma.OUT, dtype={"fundid": str, "date": str, "基金ID": str})
    record("面板", {
        "行数": len(panel), "基金数": panel.fundid.nunique(),
        "日期范围": [panel.date.min(), panel.date.max()],
        "重复键": int(panel.duplicated(["fundid", "date"]).sum()),
        "2010年前行数": int(panel.date.lt("20100101").sum()),
        "2010年前有alpha行数": int((panel.date.lt("20100101") & panel.capm_alpha.notna()).sum()),
        "H2起点早于7月且买卖有值": int((panel.date.str.endswith("1231") & panel.BuyValue.notna()
            & panel.BuySellStartdate.fillna("9999-12-31").lt(panel.date.str[:4] + "-07-01")).sum()),
    })
    nav = pd.read_excel(ma.NAV, dtype=str).iloc[2:].copy()
    nav = nav.apply(lambda s: s.str.strip())
    nav["year_month"] = nav.TradingDate.str[:7]
    nav["AdjustedNAVGrowth"] = pd.to_numeric(nav.AdjustedNAVGrowth, errors="coerce")
    nav = nav.sort_values(["Symbol", "year_month"], ignore_index=True)
    nav = ra.mark_sample(nav)
    record("净值键", {
        "行数": len(nav), "基金数": nav.Symbol.nunique(),
        "重复基金月份": int(nav.duplicated(["Symbol", "year_month"]).sum()),
        "同代码多个份额ID": int(nav.groupby("Symbol").FundClassID.nunique().gt(1).sum()),
        "增长率小于等于负100百分比": int(nav.AdjustedNAVGrowth.le(-100).sum()),
        "首个有效增长率样例": nav.loc[nav.Symbol.eq("000006"),
            ["Symbol", "year_month", "AdjustedNAVGrowth", "sample"]].head(4).to_dict("records"),
    })
    s = nav.loc[nav["sample"]].copy()
    s["date"] = ma.half_date(s.year_month)
    s["gross"] = 1 + s.AdjustedNAVGrowth / 100
    g = s.groupby(["Symbol", "date"]).agg(
        product=("gross", "prod"), n=("gross", "count"), months=("year_month", "nunique")
    )
    expected = (g["product"] ** (1 / 6) - 1).where(g.n.eq(6))
    record("raw_ret全量核对", compare(panel, expected, "raw_ret"))
    sample = panel.loc[panel.fundid.eq("000006") & panel.date.eq("20240630")].iloc[0]
    record("六个月收益样例", {
        "fundid": "000006", "date": "20240630", "raw_ret": sample.raw_ret,
        "六个月累计收益": (1 + sample.raw_ret) ** 6 - 1,
        "底层月份": s.loc[s.Symbol.eq("000006") & s.date.eq("20240630"),
            ["year_month", "AdjustedNAVGrowth"]].to_dict("records"),
        "计数等于6但月份不唯一的分组": int((g.n.eq(6) & g.months.ne(6)).sum()),
    })
    rf = pd.read_excel(ROOT / "无风险利率_月度.xlsx")
    record("无风险利率", {
        "变更日晚于月末": int((pd.to_datetime(rf.RateChangeDate) > pd.to_datetime(rf.month_end)).sum()),
        "月内发生利率变更月份": rf.loc[
            pd.to_datetime(rf.RateChangeDate).dt.strftime("%Y-%m").eq(rf.year_month)
            & pd.to_datetime(rf.RateChangeDate).dt.day.gt(1),
            ["year_month", "RateChangeDate", "q_annual_pct", "rf_monthly"]
        ].to_dict("records"),
    })

    # 用源因子与原始收益独立重算少量回归，不依赖已合并的因子大表。
    from merge_ff_factors import SPECS as FF_SPECS, load_factors
    from merge_ch_factors import SPECS as CH_SPECS, load_ch
    for name, fname in ma.ALPHAS:
        print("读取", fname, flush=True)
        alpha = pd.read_excel(ROOT / fname, dtype={"Symbol": str, "FundClassID": str,
            "year_month": str, "WindowStart": str, "WindowEnd": str})
        alpha["date"] = ma.half_date(alpha.year_month)
        grouped = alpha.groupby(["Symbol", "date"]).alpha.agg(["mean", "count"])
        check = compare(panel, grouped["mean"].where(grouped["count"].eq(6)), name)
        months = pd.PeriodIndex(alpha.year_month, freq="M")
        check.update({
            "月度行数": len(alpha),
            "重复基金月份": int(alpha.duplicated(["Symbol", "year_month"]).sum()),
            "窗口起点非t减24": int((alpha.WindowStart.to_numpy() != (months - 24).astype(str)).sum()),
            "窗口终点非t减1": int((alpha.WindowEnd.to_numpy() != (months - 1).astype(str)).sum()),
        })
        spec = next(x for x in ra.SPECS if x["out"].name == fname)
        factors = spec["factors"]
        if spec["label"] in {"CH3", "CH4"}:
            src = next(x for x in CH_SPECS if x["label"] == spec["label"])
            fac = load_ch(src["src"], src["factors"])
        else:
            label = "三因子" if spec["label"] == "CAPM" else spec["label"]
            src = next(x for x in FF_SPECS if x["label"] == label)
            fac = load_factors(src["src"], src["factors"], src["extra_filters"])
            fac = fac.merge(rf[["year_month", "rf_monthly"]], on="year_month", validate="one_to_one")
        data = nav.merge(fac, on="year_month", how="left", validate="many_to_one")
        manual = []
        for fund, month in [("000006", "2021-05"), ("000006", "2024-06"), ("003853", "2018-06")]:
            observed = alpha.loc[alpha.Symbol.eq(fund) & alpha.year_month.eq(month)]
            if observed.empty:
                continue
            t = pd.Period(month, freq="M")
            h = data.loc[data.Symbol.eq(fund) & data.year_month.between(str(t - 24), str(t - 1))]
            now = data.loc[data.Symbol.eq(fund) & data.year_month.eq(month)].iloc[0]
            if len(h) != 24 or not h["sample"].all():
                raise ValueError(f"独立回归样例历史窗口不完整：{fund} {month}")
            fit = sm.OLS(h.AdjustedNAVGrowth / 100 - h[spec["rf"]],
                sm.add_constant(h[factors], has_constant="add")).fit()
            value = now.AdjustedNAVGrowth / 100 - now[spec["rf"]] - np.dot(fit.params[factors], now[factors])
            manual.append({"fundid": fund, "评价月": month, "历史起点": h.year_month.min(),
                "历史终点": h.year_month.max(), "历史月数": len(h),
                "alpha误差": float(value - observed.iloc[0].alpha),
                "截距误差": float(fit.params["const"] - observed.iloc[0].a_hat)})
        check["独立OLS样例"] = manual
        record(name, check)
        if name == "capm_alpha":
            record("转型跨界样例", alpha.loc[alpha.Symbol.isin(["003853", "000646", "163406"])
                & alpha.year_month.isin(["2018-06", "2019-12", "2021-06"]),
                ["Symbol", "FullName", "year_month", "WindowStart", "WindowEnd", "alpha"]].to_dict("records"))

    for label, loader, value_cols in [
        ("费用", mf.load_fee, ["TotalOperatingCost"]),
        ("买卖", mb.load_turnover, ["BuyValue", "SellValue"]),
    ]:
        reports = loader()
        annual = reports.loc[reports.ReportTypeID.eq("6")].copy()
        annual["year"] = annual.EndDate.str[:4]
        semi = reports.loc[reports.ReportTypeID.eq("5")].copy()
        semi["year"] = semi.EndDate.str[:4]
        paired = annual.merge(semi[["MasterFundCode", "FundID", "year"]],
            on=["MasterFundCode", "FundID", "year"], how="left", indicator=True)
        bad = paired.loc[paired._merge.eq("left_only") & paired.Startdate.lt(paired.year + "-07-01")]
        record(label + "无半年报且年报覆盖早于7月", {
            "原始报告数": len(bad), "样例": bad[["MasterFundCode", "FundID", "Startdate", "EndDate", *value_cols]].head(8).to_dict("records")
        })

    record("完成", True)


if __name__ == "__main__":
    main()
