"""独立核验输出、精确半年匹配和原始行情复算；不修改原始文件。"""

import json
from pathlib import Path

import numpy as np
import pandas as pd


ROOT = Path(__file__).resolve().parent


def main():
    p = pd.read_parquet(ROOT / "结果/stock_halfyear.parquet")
    controls = pd.read_parquet(ROOT / "结果/stock_holding_controls.parquet")
    seed = pd.read_parquet(ROOT / "中间数据/20091231初始快照.parquet")
    source = pd.concat([seed.drop(columns="is_formal_sample"), p], ignore_index=True)
    keys = ["stock_code", "date"]
    extra = ["pe_ttm", "ep_ttm", "pb", "roe_ttm", "leverage", "asset_growth_yoy",
             "revenue_yoy", "nonoperating_profit_share_ttm", "rd_expense_to_revenue_ttm", "tradable_cap_ratio", "loss_ttm"]
    core = ["size", "bm", "roa", "ret6m", "rec_change", *extra]
    assert p.stock_code.str.fullmatch(r"\d{6}").all()
    assert not p.duplicated(keys).any() and not controls.duplicated(keys).any()
    assert p.date.min() == 20100630 and p.date.max() == 20260630 and p.date.nunique() == 33
    assert p["asof"].dt.month.isin([6, 12]).all()
    assert len(controls) == len(p)
    assert not np.isinf(p.select_dtypes(include="number").to_numpy()).any()
    for col, dt in [("bm", "book_source_available"), ("roa", "roa_source_available"), ("rec_consensus", "rec_latest_publication")]:
        active = p[col].notna()
        assert p.loc[active, dt].notna().all()
        assert (p.loc[active, dt] <= p.loc[active, "asof"]).all()
    for col, dt in [("pe_ttm", "earnings_ttm_available"), ("ep_ttm", "earnings_ttm_available"),
                    ("pb", "book_source_available"), ("roe_ttm", "roe_source_available"),
                    ("leverage", "book_source_available"), ("asset_growth_yoy", "balance_yoy_available"),
                    ("revenue_yoy", "revenue_yoy_available"), ("nonoperating_profit_share_ttm", "earnings_ttm_available"),
                    ("rd_expense_to_revenue_ttm", "earnings_ttm_available")]:
        active = p[col].notna()
        assert p.loc[active, dt].notna().all()
        assert (p.loc[active, dt] <= p.loc[active, "asof"]).all()
    # 不调用生产函数：用公历明确构造期初键，再与快照逐值对照。
    expected_date = np.where(controls.date % 10000 == 630,
                             (controls.date // 10000 - 1) * 10000 + 1231,
                             controls.date // 10000 * 10000 + 630)
    left = controls.assign(expected_info_date=expected_date)
    expected = left.merge(source[keys + core], left_on=["stock_code", "expected_info_date"],
                          right_on=keys, how="left", validate="m:1")
    for col in core:
        np.testing.assert_allclose(expected[col + "_lag"], expected[col], rtol=0, atol=0, equal_nan=True)
    nonmissing = controls.info_date.notna()
    np.testing.assert_array_equal(controls.loc[nonmissing, "info_date"].astype(int), expected_date[nonmissing])
    assert p.loc[p.date.eq(20260630), "ret6m_fwd"].isna().all()
    next_keys = p[keys + ["ret6m_fwd"]].copy()
    next_keys["target"] = np.where(next_keys.date % 10000 == 630,
                                 next_keys.date // 10000 * 10000 + 1231,
                                 (next_keys.date // 10000 + 1) * 10000 + 630)
    forward = next_keys.merge(p[keys + ["ret6m"]], left_on=["stock_code", "target"], right_on=keys, how="left", validate="m:1")
    np.testing.assert_allclose(forward.ret6m_fwd, forward.ret6m, rtol=0, atol=0, equal_nan=True)
    # 从原始月表缓存直接复算，不引用生产流程算好的半年中间值。
    m = pd.read_parquet(ROOT / "中间数据/月个股回报率文件114059067.parquet")
    examples = p.loc[p.market_cap.notna()].sample(60, random_state=20261105)
    for row in examples.itertuples(index=False):
        end = row.asof.strftime("%Y-%m")
        start = f"{row.asof.year}-{'01' if row.asof.month == 6 else '07'}"
        month = m.loc[m.Stkcd.eq(row.stock_code) & m.Trdmnt.between(start, end)]
        endpoint = month.loc[month.Trdmnt.eq(end)].iloc[0]
        np.testing.assert_allclose(float(endpoint.Msmvttl) * 1000, row.market_cap, rtol=0, atol=0)
        returns = pd.to_numeric(month.Mretwd, errors="coerce")
        if len(month) == 6 and returns.notna().all():
            np.testing.assert_allclose(np.prod(1 + returns) - 1, row.ret6m, rtol=1e-12, atol=1e-14)
        else:
            assert pd.isna(row.ret6m)
    bm_rows = p.bm.notna()
    np.testing.assert_allclose(p.loc[bm_rows, "bm"], p.loc[bm_rows, "book_equity"] / p.loc[bm_rows, "market_cap"], rtol=0, atol=0)
    roa_rows = p.roa.notna()
    np.testing.assert_allclose(p.loc[roa_rows, "roa"], p.loc[roa_rows, "income_ttm"] / p.loc[roa_rows, "roa_average_assets"], rtol=0, atol=0)
    assert p.loc[p.parent_income_ttm.le(0), "pe_ttm"].isna().all()
    assert p.loc[p.parent_income_ttm.isna(), "loss_ttm"].isna().all()
    for col, numerator, denominator in [("pe_ttm", "market_cap", "parent_income_ttm"),
            ("ep_ttm", "parent_income_ttm", "market_cap"), ("pb", "market_cap", "book_equity"),
            ("roe_ttm", "parent_income_ttm", "roe_average_equity"),
            ("rd_expense_to_revenue_ttm", "rd_expense_ttm", "revenue_ttm"),
            ("tradable_cap_ratio", "float_market_cap", "market_cap")]:
        valid = p[col].notna()
        np.testing.assert_allclose(p.loc[valid, col], p.loc[valid, numerator] / p.loc[valid, denominator], rtol=0, atol=0)
    assert p.loc[p.fin_report_period.lt(pd.Timestamp("2018-12-31")), "rd_expense_to_revenue_ttm"].isna().all()
    # 归母利润直接从输入报表复算，避免把合并净利润误当作 PE 分母。
    inc = pd.read_parquet(ROOT / "中间数据/利润表114149331.parquet")
    inc = inc.loc[inc.Typrep.eq("A")].set_index(["Stkcd", "Accper"])
    earnings_examples = p.loc[p.pe_ttm.notna()].sample(60, random_state=20261106)
    for row in earnings_examples.itertuples(index=False):
        period = row.fin_report_period
        now = float(inc.loc[(row.stock_code, period.strftime("%Y-%m-%d")), "B002000101"])
        if period.month != 12:
            old = float(inc.loc[(row.stock_code, (period - pd.DateOffset(years=1)).strftime("%Y-%m-%d")), "B002000101"])
            annual = float(inc.loc[(row.stock_code, f"{period.year - 1}-12-31"), "B002000101"])
            now += annual - old
        np.testing.assert_allclose(row.parent_income_ttm, now, rtol=1e-12, atol=1e-6)
        np.testing.assert_allclose(row.pe_ttm, row.market_cap / now, rtol=1e-12, atol=1e-12)
    # CSV 与 Parquet 的键、核心数值完全一致，且读入时保留前导零。
    csv = pd.read_csv(ROOT / "结果/stock_halfyear.csv", dtype={"stock_code": str, "capital_change_type": str}, float_precision="round_trip")
    assert csv[keys].astype(str).equals(p[keys].astype(str))
    assert csv.capital_change_type.fillna("").equals(p.capital_change_type.fillna(""))
    for col in core:
        np.testing.assert_allclose(csv[col], p[col], rtol=0, atol=0, equal_nan=True)
    bridge = pd.read_csv(ROOT / "结果/stock_code_bridge.csv", dtype=str)
    assert not bridge.holding_code.duplicated().any()
    assert bridge.loc[bridge.holding_code.eq("832000"), "stock_code"].tolist() == ["920000"]
    # 模拟两只基金持有同一只股票：m:1 合并必须保持持仓行数和基金身份。
    example = controls.loc[controls.controls_complete_lag.eq(True)].head(2)[keys]
    holdings = pd.concat([example.assign(fundid="000001"), example.assign(fundid="000002")], ignore_index=True)
    joined = holdings.merge(controls, on=keys, how="left", validate="m:1", indicator=True)
    assert len(joined) == len(holdings) and joined._merge.eq("both").all()
    result = {"status": "PASS", "rows": len(p), "stocks": int(p.stock_code.nunique()),
              "core_values_csv_parquet_equal": True, "all_lag_keys_and_values_exact": True,
              "all_forward_return_keys_exact": True, "source_monthly_spot_checks": len(examples),
              "source_parent_earnings_pe_spot_checks": len(earnings_examples), "extra_metric_checks": extra,
              "availability_checks": "PASS", "mock_holdings_merge": "PASS"}
    (ROOT / "核验/独立验收.json").write_text(json.dumps(result, ensure_ascii=False, indent=2))
    print(json.dumps(result, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
