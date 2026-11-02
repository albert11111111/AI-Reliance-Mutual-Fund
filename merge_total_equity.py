"""把资产负债表 TotalEquity 接到半年面板。不改基金名单_fundid_date.xlsx 和 基金size.xlsx。"""

from pathlib import Path

import pandas as pd

ROOT = Path(__file__).resolve().parent
PANEL = ROOT / "基金名单_fundid_date.xlsx"
SIZE = ROOT / "基金size.xlsx"
OUT = ROOT / "基金名单_fundid_date_TotalEquity.xlsx"
# 与 build 面板文本列一致；TotalEquity 在 N 列（多了到期日）
TEXT_COLS = ("A", "B", "F", "G", "H", "I", "J")


def load_size():
    size = pd.read_excel(SIZE, dtype=str).iloc[2:].copy()
    for col in size.columns:
        size[col] = size[col].str.strip()
    # 说明书表五十六：5 半年度、6 年度；StateTypeCode 2 为期末。只接与面板半年末同一天的报告。
    keep = (size["StateTypeCode"] == "2") & (
        ((size["ReportTypeID"] == "5") & size["EndDate"].str.endswith("06-30"))
        | ((size["ReportTypeID"] == "6") & size["EndDate"].str.endswith("12-31"))
    )
    size = size.loc[keep, ["MasterFundCode", "FundID", "ReportTypeID", "EndDate", "TotalEquity"]]
    if size.duplicated(["MasterFundCode", "FundID", "EndDate", "ReportTypeID"]).any():
        raise SystemExit("规模表同一基金、同一报告期末有多条 TotalEquity")
    return size


def build():
    panel = pd.read_excel(PANEL, dtype=str)
    n = len(panel)
    panel["end"] = panel["date"].str[:4] + "-" + panel["date"].str[4:6] + "-" + panel["date"].str[6:8]
    panel["want_type"] = panel["date"].str[-4:].map({"0630": "5", "1231": "6"})
    if panel["want_type"].isna().any():
        raise SystemExit("面板 date 不是 0630 或 1231")

    merged = panel.merge(
        load_size(),
        left_on=["基金ID", "fundid", "end", "want_type"],
        right_on=["FundID", "MasterFundCode", "EndDate", "ReportTypeID"],
        how="left",
        validate="one_to_one",
    )
    if len(merged) != n:
        raise SystemExit("合并后行数变化，键不唯一")
    merged["TotalEquity"] = pd.to_numeric(merged["TotalEquity"], errors="coerce")
    out = merged.drop(columns=["end", "want_type", "MasterFundCode", "FundID", "ReportTypeID", "EndDate"])
    for col in ("发生转型代码重复", "发生过转型"):
        out[col] = out[col].astype(int)

    hit = out["TotalEquity"].notna()
    assert int(hit.sum()) == 44507
    assert int((~hit).sum()) == 1422
    assert not out.duplicated(["fundid", "date"]).any()
    sample = out[(out.fundid == "000006") & (out.date == "20190630")]
    assert len(sample) == 1
    assert round(float(sample["TotalEquity"].iloc[0]), 2) == 137005458.29
    june = out[(out.fundid == "519714") & (out.date == "20150630")]
    assert round(float(june["TotalEquity"].iloc[0]), 2) == 9343854.14

    with pd.ExcelWriter(OUT, engine="openpyxl") as writer:
        out.to_excel(writer, index=False, sheet_name="fundid_date")
        ws = writer.sheets["fundid_date"]
        ws.freeze_panes = "A2"
        ws.auto_filter.ref = ws.dimensions
        for col in TEXT_COLS:
            for cell in ws[col]:
                cell.number_format = "@"
        for cell in ws["N"][1:]:
            cell.number_format = "0.00"
        ws.column_dimensions["N"].width = 18

    print(f"rows={len(out)} matched={int(hit.sum())} missing={int((~hit).sum())}")


if __name__ == "__main__":
    build()
