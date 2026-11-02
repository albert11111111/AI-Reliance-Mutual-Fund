"""把持有人结构比例接到已含 TotalEquity 的半年面板。不改源面板和机构持有占比表。"""

from pathlib import Path

import pandas as pd

from excel_style import style_sheet

ROOT = Path(__file__).resolve().parent
PANEL = ROOT / "基金名单_fundid_date_TotalEquity.xlsx"
HOLDER = ROOT / "机构持有占比.xlsx"
OUT = ROOT / "基金名单_fundid_date_TotalEquity_Holder.xlsx"
TEXT_COLS = (
    "fundid",
    "date",
    "成立时间_earliest",
    "成立时间_latest",
    "到期日",
    "基金ID",
    "基金管理公司ID",
)
KEEP_FROM_HOLDER = [
    "Symbol",
    "FundID",
    "ReportTypeID",
    "EndDate",
    "RelationCode",
    "InstitutionRatio",
    "IndividualRatio",
    "PractitionersRatio",
    "TotalInstitutionRatio",
    "TotalIndividualRatio",
    "TotalPractitionersRatio",
]


def nonempty(series: pd.Series) -> pd.Series:
    return series.notna() & ~series.isin(["", "nan", "None"])


def load_holder():
    holder = pd.read_excel(HOLDER, dtype=str).iloc[2:].copy()
    for col in holder.columns:
        holder[col] = holder[col].str.strip()
    # 说明书表四十六：5 半年度、6 年度。只接与面板半年末同一天的报告。
    keep = ((holder["ReportTypeID"] == "5") & holder["EndDate"].str.endswith("06-30")) | (
        (holder["ReportTypeID"] == "6") & holder["EndDate"].str.endswith("12-31")
    )
    holder = holder.loc[keep, KEEP_FROM_HOLDER]
    if holder.duplicated(["Symbol", "FundID", "EndDate", "ReportTypeID"]).any():
        raise SystemExit("持有人表同一份额、同一报告期末有多条记录")
    return holder


def choose_ratios(df: pd.DataFrame) -> pd.DataFrame:
    """多份额用总产品比例，单份额用本份额比例。是否多份额看该期是否披露了任一总产品比例。"""
    multi = (
        nonempty(df["TotalInstitutionRatio"])
        | nonempty(df["TotalIndividualRatio"])
        | nonempty(df["TotalPractitionersRatio"])
    )
    matched = nonempty(df["Symbol"])
    out = pd.DataFrame(index=df.index)
    out["InstitutionRatio"] = pd.NA
    out["IndividualRatio"] = pd.NA
    out["PractitionersRatio"] = pd.NA
    out["HolderRatioScope"] = pd.NA

    out.loc[matched & multi, "InstitutionRatio"] = df.loc[matched & multi, "TotalInstitutionRatio"]
    out.loc[matched & multi, "IndividualRatio"] = df.loc[matched & multi, "TotalIndividualRatio"]
    out.loc[matched & multi, "PractitionersRatio"] = df.loc[matched & multi, "TotalPractitionersRatio"]
    out.loc[matched & multi, "HolderRatioScope"] = "总产品"

    single = matched & ~multi
    out.loc[single, "InstitutionRatio"] = df.loc[single, "InstitutionRatio"]
    out.loc[single, "IndividualRatio"] = df.loc[single, "IndividualRatio"]
    out.loc[single, "PractitionersRatio"] = df.loc[single, "PractitionersRatio"]
    out.loc[single, "HolderRatioScope"] = "本份额"

    for col in ("InstitutionRatio", "IndividualRatio", "PractitionersRatio"):
        out[col] = pd.to_numeric(out[col], errors="coerce")
    return out


def build():
    panel = pd.read_excel(PANEL, dtype=str)
    n = len(panel)
    panel["end"] = panel["date"].str[:4] + "-" + panel["date"].str[4:6] + "-" + panel["date"].str[6:8]
    panel["want_type"] = panel["date"].str[-4:].map({"0630": "5", "1231": "6"})
    if panel["want_type"].isna().any():
        raise SystemExit("面板 date 不是 0630 或 1231")

    merged = panel.merge(
        load_holder(),
        left_on=["基金ID", "fundid", "end", "want_type"],
        right_on=["FundID", "Symbol", "EndDate", "ReportTypeID"],
        how="left",
        validate="one_to_one",
    )
    if len(merged) != n:
        raise SystemExit("合并后行数变化，键不唯一")

    chosen = choose_ratios(merged)
    out = pd.concat(
        [
            merged.drop(
                columns=[
                    "end",
                    "want_type",
                    "Symbol",
                    "FundID",
                    "ReportTypeID",
                    "EndDate",
                    "RelationCode",
                    "InstitutionRatio",
                    "IndividualRatio",
                    "PractitionersRatio",
                    "TotalInstitutionRatio",
                    "TotalIndividualRatio",
                    "TotalPractitionersRatio",
                ]
            ),
            chosen,
        ],
        axis=1,
    )
    for col in ("发生转型代码重复", "发生过转型"):
        out[col] = out[col].astype(int)
    out["TotalEquity"] = pd.to_numeric(out["TotalEquity"], errors="coerce")

    matched = out["HolderRatioScope"].notna()
    multi = out["HolderRatioScope"].eq("总产品")
    single = out["HolderRatioScope"].eq("本份额")
    assert int(matched.sum()) == 44484
    assert int((~matched).sum()) == 1445
    assert not out.duplicated(["fundid", "date"]).any()

    # 000006：早期单份额用本份额，后期多份额用总产品
    row_201906 = out[(out.fundid == "000006") & (out.date == "20190630")].iloc[0]
    assert row_201906["HolderRatioScope"] == "本份额"
    assert round(float(row_201906["InstitutionRatio"]), 2) == 7.24
    assert round(float(row_201906["IndividualRatio"]), 2) == 92.76
    row_202106 = out[(out.fundid == "000006") & (out.date == "20210630")].iloc[0]
    assert row_202106["HolderRatioScope"] == "总产品"
    assert round(float(row_202106["InstitutionRatio"]), 2) == 73.42
    assert round(float(row_202106["IndividualRatio"]), 2) == 26.58

    with pd.ExcelWriter(OUT, engine="openpyxl") as writer:
        out.to_excel(writer, index=False, sheet_name="fundid_date")
        style_sheet(
            writer.sheets["fundid_date"],
            out.columns,
            text_cols=TEXT_COLS,
            number_formats={
                "TotalEquity": "0.00",
                "InstitutionRatio": "0.0000",
                "IndividualRatio": "0.0000",
                "PractitionersRatio": "0.0000",
            },
            widths={
                "TotalEquity": 18,
                "InstitutionRatio": 18,
                "IndividualRatio": 18,
                "PractitionersRatio": 18,
                "HolderRatioScope": 12,
            },
        )

    print(f"rows={len(out)} matched={int(matched.sum())} missing={int((~matched).sum())}")
    print(f"scope_总产品={int(multi.sum())} scope_本份额={int(single.sum())}")
    print(
        "filled "
        f"Institution={int(out['InstitutionRatio'].notna().sum())} "
        f"Individual={int(out['IndividualRatio'].notna().sum())} "
        f"Practitioners={int(out['PractitionersRatio'].notna().sum())}"
    )


if __name__ == "__main__":
    build()
