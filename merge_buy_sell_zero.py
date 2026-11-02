"""在 Holder 半年面板上接股票买卖金额；缺侧按 0，一侧减负则买卖都空。

来源：说明书表五十一 Fund_Ptf_BSStockTotal。
输入：基金名单_fundid_date_TotalEquity_Holder.xlsx（不改）。
输出：基金名单_fundid_date_TotalEquity_Holder_BuySell.xlsx（新文件）。
不改 换手率计算.xlsx。

BuyValue / SellValue 是该半年的发生额：
- 0630：半年报原值；缺买或缺卖视为 0
- 1231：有同年半年报且起算日相同则 年报 − 半年报（缺侧按 0）；
  无同年半年报则直接用年报（缺侧按 0）；
  起算日不一致，或任一侧减出负值：该期买卖都留空
"""

from pathlib import Path

import pandas as pd

from excel_style import style_sheet

ROOT = Path(__file__).resolve().parent
PANEL = ROOT / "基金名单_fundid_date_TotalEquity_Holder.xlsx"
TURNOVER = ROOT / "换手率计算.xlsx"
OUT = ROOT / "基金名单_fundid_date_TotalEquity_Holder_BuySell.xlsx"
TEXT_COLS = (
    "fundid",
    "date",
    "成立时间_earliest",
    "成立时间_latest",
    "到期日",
    "基金ID",
    "基金管理公司ID",
)
NEG_TOL = -0.01  # 小于此视为减出负值，不是舍入


def load_turnover() -> pd.DataFrame:
    """只保留 06-30 半年报与 12-31 年报。"""
    raw = pd.read_excel(TURNOVER, dtype=str).iloc[2:].copy()
    for col in raw.columns:
        raw[col] = raw[col].str.strip()

    keep = ((raw["ReportTypeID"] == "5") & raw["EndDate"].str.endswith("06-30")) | (
        (raw["ReportTypeID"] == "6") & raw["EndDate"].str.endswith("12-31")
    )
    out = raw.loc[keep, ["MasterFundCode", "FundID", "ReportTypeID", "EndDate", "Startdate", "BuyValue", "SellValue"]]
    if out.duplicated(["MasterFundCode", "FundID", "EndDate", "ReportTypeID"]).any():
        raise SystemExit("买卖金额表同一基金、同一报告期末有多条记录")
    out["BuyValue"] = pd.to_numeric(out["BuyValue"], errors="coerce")
    out["SellValue"] = pd.to_numeric(out["SellValue"], errors="coerce")
    return out


def half_year_amounts(turnover: pd.DataFrame) -> tuple[pd.DataFrame, dict]:
    """整理成每个半年末对应的买卖发生额；有报告行时买卖成对出现。"""
    semi = turnover.loc[turnover["ReportTypeID"] == "5"].copy()
    annual = turnover.loc[turnover["ReportTypeID"] == "6"].copy()

    # 上半年：有半年报行则缺侧填 0，买卖同时有数
    h1 = pd.DataFrame(
        {
            "MasterFundCode": semi["MasterFundCode"],
            "FundID": semi["FundID"],
            "ReportTypeID": semi["ReportTypeID"],
            "EndDate": semi["EndDate"],
            "BuySellPeriod": "H1",
            "BuySellStartdate": semi["Startdate"],
            "BuyValue": semi["BuyValue"].fillna(0.0),
            "SellValue": semi["SellValue"].fillna(0.0),
        }
    )

    left = annual.copy()
    left["year"] = left["EndDate"].str[:4]
    right = semi.rename(
        columns={
            "Startdate": "semi_start",
            "BuyValue": "buy_h1",
            "SellValue": "sell_h1",
        }
    )
    right["year"] = right["EndDate"].str[:4]
    paired = left.merge(
        right[["MasterFundCode", "FundID", "year", "semi_start", "buy_h1", "sell_h1"]],
        on=["MasterFundCode", "FundID", "year"],
        how="left",
    )

    no_semi = paired["semi_start"].isna()
    same_start = paired["Startdate"].eq(paired["semi_start"])
    # 缺侧按 0 再减；任一侧减成负 → 该期买卖都空
    buy_h2 = paired["BuyValue"].fillna(0.0) - paired["buy_h1"].fillna(0.0)
    sell_h2 = paired["SellValue"].fillna(0.0) - paired["sell_h1"].fillna(0.0)
    neg_either = same_start & ((buy_h2 < NEG_TOL) | (sell_h2 < NEG_TOL))
    ok_sub = same_start & ~neg_either

    h2 = pd.DataFrame(
        {
            "MasterFundCode": paired["MasterFundCode"],
            "FundID": paired["FundID"],
            "ReportTypeID": paired["ReportTypeID"],
            "EndDate": paired["EndDate"],
            "BuySellPeriod": "H2",
            "BuySellStartdate": pd.NA,
            "BuyValue": pd.NA,
            "SellValue": pd.NA,
        }
    )
    h2.loc[same_start, "BuySellStartdate"] = paired.loc[same_start, "EndDate"].str[:4] + "-07-01"
    h2.loc[no_semi, "BuySellStartdate"] = paired.loc[no_semi, "Startdate"]
    h2.loc[ok_sub, "BuyValue"] = buy_h2[ok_sub].clip(lower=0.0)
    h2.loc[ok_sub, "SellValue"] = sell_h2[ok_sub].clip(lower=0.0)
    # 无同年半年报：年报缺侧按 0，买卖成对
    h2.loc[no_semi, "BuyValue"] = paired.loc[no_semi, "BuyValue"].fillna(0.0)
    h2.loc[no_semi, "SellValue"] = paired.loc[no_semi, "SellValue"].fillna(0.0)

    stats = {
        "h2_no_semi": int(no_semi.sum()),
        "h2_start_mismatch": int((paired["semi_start"].notna() & ~same_start).sum()),
        "h2_neg_either": int(neg_either.sum()),
        "h2_sub_ok": int(ok_sub.sum()),
        "h1_filled_buy": int(semi["BuyValue"].isna().sum()),
        "h1_filled_sell": int(semi["SellValue"].isna().sum()),
    }
    return pd.concat([h1, h2], ignore_index=True), stats


def build() -> None:
    panel = pd.read_excel(PANEL, dtype=str)
    n = len(panel)
    panel["end"] = panel["date"].str[:4] + "-" + panel["date"].str[4:6] + "-" + panel["date"].str[6:8]
    panel["want_type"] = panel["date"].str[-4:].map({"0630": "5", "1231": "6"})
    if panel["want_type"].isna().any():
        raise SystemExit("面板 date 不是 0630 或 1231")

    amounts, stats = half_year_amounts(load_turnover())
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
    out["BuyValue"] = pd.to_numeric(out["BuyValue"], errors="coerce")
    out["SellValue"] = pd.to_numeric(out["SellValue"], errors="coerce")
    out["TotalEquity"] = pd.to_numeric(out["TotalEquity"], errors="coerce")
    for col in ("InstitutionRatio", "IndividualRatio", "PractitionersRatio"):
        out[col] = pd.to_numeric(out[col], errors="coerce")
    for col in ("发生转型代码重复", "发生过转型"):
        out[col] = out[col].astype(int)

    # 有报告期不应出现买卖一边有一边空
    one_side = out["BuyValue"].isna() ^ out["SellValue"].isna()
    if one_side.any():
        raise SystemExit(f"仍有买卖单边空 {int(one_side.sum())} 行")

    cols = [c for c in out.columns if c not in ("BuySellPeriod", "BuySellStartdate", "BuyValue", "SellValue")]
    out = out[cols + ["BuySellPeriod", "BuySellStartdate", "BuyValue", "SellValue"]]

    with pd.ExcelWriter(OUT, engine="openpyxl") as writer:
        out.to_excel(writer, index=False, sheet_name="fundid_date")
        style_sheet(
            writer.sheets["fundid_date"],
            out.columns,
            text_cols=TEXT_COLS,
            number_formats={"BuyValue": "0.00", "SellValue": "0.00"},
            widths={
                "BuySellPeriod": 12,
                "BuySellStartdate": 16,
                "BuyValue": 18,
                "SellValue": 18,
            },
        )

    both = out["BuyValue"].notna() & out["SellValue"].notna()
    neither = out["BuyValue"].isna() & out["SellValue"].isna()
    print(
        f"rows={len(out)} "
        f"H1={int(out['BuySellPeriod'].eq('H1').sum())} "
        f"H2={int(out['BuySellPeriod'].eq('H2').sum())} "
        f"买卖都有={int(both.sum())} 买卖都空={int(neither.sum())}"
    )
    print(
        "处理："
        f"H1原缺买填0={stats['h1_filled_buy']} 缺卖填0={stats['h1_filled_sell']}；"
        f"H2无半年报={stats['h2_no_semi']} 相减成功={stats['h2_sub_ok']} "
        f"起算不一致留空={stats['h2_start_mismatch']} "
        f"任一侧减负整期留空={stats['h2_neg_either']}"
    )
    print(f"写出 {OUT.name}")


if __name__ == "__main__":
    build()
