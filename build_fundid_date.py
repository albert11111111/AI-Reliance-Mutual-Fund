"""把基金名单展开成 fundid × 半年末 date 面板。规则见 README.md。不改基金名单.xlsx。"""

from datetime import date, datetime
from pathlib import Path

import pandas as pd

ROOT = Path(__file__).resolve().parent
SRC = ROOT / "基金名单.xlsx"
OUT = ROOT / "基金名单_fundid_date.xlsx"
# 运行日 2026-10-30 时，最近一个已经到来的半年末。下一次更新样本截止日时改这里，并改 README。
SAMPLE_END = date(2026, 6, 30)
KEEP = [
    "fundid",
    "short_name",
    "基金类型ifind",
    "基金经理",
    "成立时间_earliest",
    "成立时间_latest",
    "基金ID",
    "基金管理公司ID",
    "基金管理公司名称",
    "发生转型代码重复",
    "发生过转型",
]
TEXT_COLS = ("A", "B", "F", "G", "H", "I")


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


def build():
    src = pd.read_excel(SRC, dtype=str)
    missing = [c for c in KEEP + ["到期日"] if c not in src.columns]
    if missing:
        raise SystemExit(f"基金名单缺少列: {missing}")
    n_raw = len(src)

    # 同一 fundid 只留一行：成立时间_earliest 更早的那行决定 date。
    # 起始日相同则留 成立时间_latest 更晚的一行；其余保留列在重复行之间相同。
    src = src.sort_values(
        ["fundid", "成立时间_earliest", "成立时间_latest"],
        ascending=[True, True, False],
        kind="mergesort",
    ).drop_duplicates("fundid", keep="first")

    rows = []
    skipped = []
    for rec in src.itertuples(index=False):
        inception = parse_day(rec.成立时间_earliest)
        maturity = parse_day(rec.到期日)
        dates = semi_dates(inception, maturity)
        if not dates:
            skipped.append(rec.fundid)
            continue
        base = {k: getattr(rec, k) for k in KEEP}
        for dt in dates:
            rows.append({"fundid": base["fundid"], "date": dt, **{k: base[k] for k in KEEP if k != "fundid"}})

    order = ["fundid", "date", *[c for c in KEEP if c != "fundid"]]
    panel = pd.DataFrame(rows)[order]
    for col in ("发生转型代码重复", "发生过转型"):
        panel[col] = panel[col].astype(int)
    panel = panel.sort_values(["fundid", "date"], kind="mergesort").reset_index(drop=True)

    assert not panel.duplicated(["fundid", "date"]).any()
    assert panel["date"].str[-4:].isin(["0630", "1231"]).all()
    dt = pd.to_datetime(panel["date"], format="%Y%m%d")
    inc = pd.to_datetime(panel["成立时间_earliest"])
    assert (dt > inc).all()
    assert (dt <= pd.Timestamp(SAMPLE_END)).all()
    assert panel.loc[panel.fundid == "000006", "date"].iloc[0] == "20190630"
    assert panel.loc[panel.fundid == "000209", "date"].iloc[0] == "20131231"
    assert panel.loc[panel.fundid == "002708", "date"].iloc[0] == "20161231"
    row_163406 = panel.loc[panel.fundid == "163406"]
    assert row_163406["成立时间_earliest"].nunique() == 1
    assert row_163406["date"].iloc[0] == "20100630"
    assert row_163406["成立时间_latest"].iloc[0] == "2021-01-01"
    row_646 = panel.loc[panel.fundid == "000646"]
    assert row_646["成立时间_earliest"].iloc[0] == "2014-08-18"
    assert row_646["成立时间_latest"].iloc[0] == "2019-09-04"
    row_162107 = panel.loc[panel.fundid == "162107"]
    assert row_162107["成立时间_earliest"].iloc[0] == "2016-08-24"
    assert row_162107["成立时间_latest"].iloc[0] == "2016-08-25"

    with pd.ExcelWriter(OUT, engine="openpyxl") as writer:
        panel.to_excel(writer, index=False, sheet_name="fundid_date")
        ws = writer.sheets["fundid_date"]
        ws.freeze_panes = "A2"
        ws.auto_filter.ref = ws.dimensions
        for col in TEXT_COLS:
            for cell in ws[col]:
                cell.number_format = "@"
        widths = {
            "A": 12, "B": 12, "C": 28, "D": 16, "E": 36, "F": 20,
            "G": 20, "H": 14, "I": 16, "J": 22, "K": 18, "L": 12,
        }
        for col, width in widths.items():
            ws.column_dimensions[col].width = width

    print(f"source_rows={n_raw}")
    print(f"funds_after_dedup={len(src)}")
    print(f"panel_rows={len(panel)}")
    print(f"funds_in_panel={panel['fundid'].nunique()}")
    print(f"rows_without_date={len(skipped)}")
    print(f"date_min={panel['date'].min()} date_max={panel['date'].max()}")
    print(f"dup_fundid_date={int(panel.duplicated(['fundid', 'date']).sum())}")
    print("skipped=" + ",".join(skipped))


if __name__ == "__main__":
    build()
