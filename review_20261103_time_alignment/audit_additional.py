"""补充核对费用期间与经理从业年限快照；仅新增审查结果。"""

import json
from pathlib import Path
import sys

import pandas as pd

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))
sys.dont_write_bytecode = True
import merge_fee as mf


def main():
    panel = pd.read_excel(ROOT / "基金名单_fundid_date_TotalEquity_Holder_BuySell_Fee_Manager_Alpha.xlsx",
        dtype={"fundid": str, "基金ID": str, "date": str})
    fee = mf.load_fee()
    fee["year"] = fee.EndDate.str[:4]
    annual = fee.loc[fee.ReportTypeID.eq("6")]
    semi = fee.loc[fee.ReportTypeID.eq("5"), ["MasterFundCode", "FundID", "year"]]
    paired = annual.merge(semi, on=["MasterFundCode", "FundID", "year"], how="left", indicator=True)
    bad = paired.loc[paired._merge.eq("left_only") & paired.Startdate.lt(paired.year + "-07-01")].copy()
    bad["date"] = bad.year + "1231"
    hit = panel.merge(bad[["MasterFundCode", "FundID", "date", "Startdate"]],
        left_on=["fundid", "基金ID", "date"], right_on=["MasterFundCode", "FundID", "date"],
        how="inner", validate="one_to_one")
    buy_bad = panel.date.str.endswith("1231") & panel.BuyValue.notna() & (
        panel.BuySellStartdate.fillna("9999-12-31") < panel.date.str[:4] + "-07-01")
    mgr = pd.read_excel(ROOT / "Fund_FundManager.xlsx", dtype=str).iloc[2:]
    result = {
        "费用覆盖早于7月的面板行数": len(hit),
        "其中2010年后": int(hit.date.ge("20100101").sum()),
        "费用异常样例": hit[["fundid", "date", "Startdate", "TotalOperatingCost", "raw_ret"]].head(6).to_dict("records"),
        "买卖覆盖早于7月且2010年后": int((buy_bad & panel.date.ge("20100101")).sum()),
        "转型起点样例": panel.loc[panel.fundid.isin(["003853", "000646", "163406"]),
            ["fundid", "成立时间_earliest", "成立时间_latest"]].drop_duplicates().to_dict("records"),
        "从业年限源表样例": mgr.loc[mgr.MasterFundCode.eq("163406"),
            ["FullName", "ServiceStartDate", "ServiceEndDate", "BusinessDuration"]].to_dict("records"),
        "从业年限面板样例": panel.loc[panel.fundid.eq("163406") & panel.date.isin(["20130630", "20151231", "20181231", "20201231"]),
            ["fundid", "date", "Team", "midAge", "largeAge"]].to_dict("records"),
    }
    target = Path(__file__).with_name("additional_results.json")
    target.write_text(json.dumps(result, ensure_ascii=False, indent=2, default=str), encoding="utf-8")
    print(json.dumps(result, ensure_ascii=False, indent=2, default=str))


if __name__ == "__main__":
    main()
