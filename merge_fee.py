"""在 BuySell 面板上接利润表费用合计 TotalOperatingCost。

来源：说明书表五十七 Fund_FIN_Income。
输入：基金名单_fundid_date_TotalEquity_Holder_BuySell.xlsx（不改）。
输出：基金名单_fundid_date_TotalEquity_Holder_BuySell_Fee.xlsx。
不改 费用.xlsx。

费用合计是该半年的发生额：
- 0630：半年报原值；有报告行但原值为空视为 0
- 1231：有同年半年报且起算日相同则 年报 − 半年报（缺侧按 0）；
  无同年半年报则直接用年报（缺侧按 0）；
  起算日不一致，或减出负值：该期留空
"""

from pathlib import Path

import pandas as pd

from excel_style import style_sheet

ROOT = Path(__file__).resolve().parent
PANEL = ROOT / "基金名单_fundid_date_TotalEquity_Holder_BuySell.xlsx"
FEE = ROOT / "费用.xlsx"
OUT = ROOT / "基金名单_fundid_date_TotalEquity_Holder_BuySell_Fee.xlsx"
TEXT_COLS = (
    "fundid",
    "date",
    "成立时间_earliest",
    "成立时间_latest",
    "到期日",
    "基金ID",
    "基金管理公司ID",
)
NEG_TOL = -0.01


def load_fee() -> pd.DataFrame:
    """只保留本期口径的 06-30 半年报与 12-31 年报。"""
    raw = pd.read_excel(FEE, dtype=str).iloc[2:].copy()
    for col in raw.columns:
        raw[col] = raw[col].str.strip()

    # 说明书：DATA_TYPE_ID 3＝本期
    keep = (raw["DATA_TYPE_ID"] == "3") & (
        ((raw["ReportTypeID"] == "5") & raw["EndDate"].str.endswith("06-30"))
        | ((raw["ReportTypeID"] == "6") & raw["EndDate"].str.endswith("12-31"))
    )
    out = raw.loc[keep, ["MasterFundCode", "FundID", "ReportTypeID", "EndDate", "Startdate", "TotalOperatingCost"]]
    if out.duplicated(["MasterFundCode", "FundID", "EndDate", "ReportTypeID"]).any():
        raise SystemExit("费用表同一基金、同一报告期末有多条记录")
    out["TotalOperatingCost"] = pd.to_numeric(out["TotalOperatingCost"], errors="coerce")
    return out


def half_year_fee(fee: pd.DataFrame) -> tuple[pd.DataFrame, dict]:
    """整理成每个半年末对应的费用合计。"""
    semi = fee.loc[fee["ReportTypeID"] == "5"].copy()
    annual = fee.loc[fee["ReportTypeID"] == "6"].copy()

    h1 = pd.DataFrame(
        {
            "MasterFundCode": semi["MasterFundCode"],
            "FundID": semi["FundID"],
            "ReportTypeID": semi["ReportTypeID"],
            "EndDate": semi["EndDate"],
            "TotalOperatingCost": semi["TotalOperatingCost"].fillna(0.0),
        }
    )

    left = annual.copy()
    left["year"] = left["EndDate"].str[:4]
    right = semi.rename(columns={"Startdate": "semi_start", "TotalOperatingCost": "fee_h1"})
    right["year"] = right["EndDate"].str[:4]
    paired = left.merge(
        right[["MasterFundCode", "FundID", "year", "semi_start", "fee_h1"]],
        on=["MasterFundCode", "FundID", "year"],
        how="left",
    )

    no_semi = paired["semi_start"].isna()
    same_start = paired["Startdate"].eq(paired["semi_start"])
    fee_h2 = paired["TotalOperatingCost"].fillna(0.0) - paired["fee_h1"].fillna(0.0)
    neg = same_start & (fee_h2 < NEG_TOL)
    ok_sub = same_start & ~neg

    h2 = pd.DataFrame(
        {
            "MasterFundCode": paired["MasterFundCode"],
            "FundID": paired["FundID"],
            "ReportTypeID": paired["ReportTypeID"],
            "EndDate": paired["EndDate"],
            "TotalOperatingCost": pd.NA,
        }
    )
    h2.loc[ok_sub, "TotalOperatingCost"] = fee_h2[ok_sub].clip(lower=0.0)
    h2.loc[no_semi, "TotalOperatingCost"] = paired.loc[no_semi, "TotalOperatingCost"].fillna(0.0)

    stats = {
        "h2_no_semi": int(no_semi.sum()),
        "h2_start_mismatch": int((paired["semi_start"].notna() & ~same_start).sum()),
        "h2_neg": int(neg.sum()),
        "h2_sub_ok": int(ok_sub.sum()),
        "h1_filled_na": int(semi["TotalOperatingCost"].isna().sum()),
    }
    return pd.concat([h1, h2], ignore_index=True), stats


def build() -> None:
    panel = pd.read_excel(PANEL, dtype=str)
    n = len(panel)
    panel["end"] = panel["date"].str[:4] + "-" + panel["date"].str[4:6] + "-" + panel["date"].str[6:8]
    panel["want_type"] = panel["date"].str[-4:].map({"0630": "5", "1231": "6"})
    if panel["want_type"].isna().any():
        raise SystemExit("面板 date 不是 0630 或 1231")

    amounts, stats = half_year_fee(load_fee())
    merged = panel.merge(
        amounts,
        left_on=["基金ID", "fundid", "end", "want_type"],
        right_on=["FundID", "MasterFundCode", "EndDate", "ReportTypeID"],
        how="left",
        validate="one_to_one",
    )
    if len(merged) != n:
        raise SystemExit("合并后行数变化，键不唯一")

    out = merged.drop(columns=["end", "want_type", "MasterFundCode", "FundID", "ReportTypeID", "EndDate"])
    out["TotalOperatingCost"] = pd.to_numeric(out["TotalOperatingCost"], errors="coerce")
    for col in ("BuyValue", "SellValue", "TotalEquity"):
        out[col] = pd.to_numeric(out[col], errors="coerce")
    for col in ("InstitutionRatio", "IndividualRatio", "PractitionersRatio"):
        out[col] = pd.to_numeric(out[col], errors="coerce")
    for col in ("发生转型代码重复", "发生过转型"):
        out[col] = out[col].astype(int)

    # 费用列放在买卖金额后面
    cols = [c for c in out.columns if c != "TotalOperatingCost"]
    out = out[cols + ["TotalOperatingCost"]]

    with pd.ExcelWriter(OUT, engine="openpyxl") as writer:
        out.to_excel(writer, index=False, sheet_name="fundid_date")
        style_sheet(
            writer.sheets["fundid_date"],
            out.columns,
            text_cols=TEXT_COLS,
            number_formats={
                "BuyValue": "0.00",
                "SellValue": "0.00",
                "TotalOperatingCost": "0.00",
            },
            widths={
                "BuySellPeriod": 12,
                "BuySellStartdate": 16,
                "BuyValue": 18,
                "SellValue": 18,
                "TotalOperatingCost": 18,
            },
        )

    hit = out["TotalOperatingCost"].notna()
    print(f"rows={len(out)} fee有数={int(hit.sum())} fee空={int((~hit).sum())}")
    print(
        "处理："
        f"H1原缺填0={stats['h1_filled_na']}；"
        f"H2无半年报={stats['h2_no_semi']} 相减成功={stats['h2_sub_ok']} "
        f"起算不一致留空={stats['h2_start_mismatch']} "
        f"减负留空={stats['h2_neg']}"
    )
    print(f"写出 {OUT.name}")


if __name__ == "__main__":
    build()
