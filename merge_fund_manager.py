"""在 Fee 面板上按时点接基金经理团队特征（宽表，不扩行）。

来源：说明书表十 Fund_FundManager。
输入：基金名单_fundid_date_TotalEquity_Holder_BuySell_Fee.xlsx（不改）。
输出：基金名单_fundid_date_TotalEquity_Holder_BuySell_Fee_Manager.xlsx。
不改 Fund_FundManager.xlsx。

在职判定（point-in-time）：ServiceStartDate ≤ date，且离职日为空或 ≥ date。
PositionState 只做与离职日的交叉验证，不参与历史过滤。
多名经理时：Team / Female / Phd / midAge / largeAge 按团队聚合；
midAge、largeAge 来自证券从业年限 BusinessDuration，不是出生年龄。

名单空窗补丁：CSMAR 任期与名单文案不一致的缺口，按名单强制纳入指定经理。
"""

from pathlib import Path

import pandas as pd

from excel_style import style_sheet

ROOT = Path(__file__).resolve().parent
PANEL = ROOT / "基金名单_fundid_date_TotalEquity_Holder_BuySell_Fee.xlsx"
MGR = ROOT / "Fund_FundManager.xlsx"
OUT = ROOT / "基金名单_fundid_date_TotalEquity_Holder_BuySell_Fee_Manager.xlsx"
TEXT_COLS = (
    "fundid",
    "date",
    "成立时间_earliest",
    "成立时间_latest",
    "到期日",
    "基金ID",
    "基金管理公司ID",
)
KEEP_MGR = [
    "MasterFundCode",
    "FundID",
    "FullName",
    "Gender",
    "ServiceStartDate",
    "ServiceEndDate",
    "PositionState",
    "DegreeID",
    "Degree",
    "BusinessDuration",
]
# (fundid, date, FullName)：CSMAR 空窗、按名单文案补在职
LIST_GAP_FILLS = [
    ("007449", "20201231", "董理"),
    ("007449", "20210630", "董理"),
    ("540007", "20110630", "廖志峰"),
]


def load_managers() -> pd.DataFrame:
    """读经理任期，并校验 PositionState 与离职日一致。"""
    raw = pd.read_excel(MGR, dtype=str).iloc[2:].copy()
    for col in raw.columns:
        raw[col] = raw[col].str.strip()
    out = raw.loc[:, KEEP_MGR].copy()
    out["start"] = pd.to_datetime(out["ServiceStartDate"], errors="coerce")
    out["end"] = pd.to_datetime(out["ServiceEndDate"], errors="coerce")
    out["BusinessDuration"] = pd.to_numeric(out["BusinessDuration"], errors="coerce")
    if out["start"].isna().any():
        raise SystemExit("经理表存在无法解析的任职日期")

    end_open = out["end"].isna()
    bad_open = end_open & out["PositionState"].ne("1")
    bad_closed = (~end_open) & out["PositionState"].ne("2")
    if bad_open.any() or bad_closed.any():
        bad = out.loc[bad_open | bad_closed, KEEP_MGR]
        print(bad.to_string(index=False))
        raise SystemExit(
            f"PositionState 与离职日不一致：开区间非现任 {int(bad_open.sum())}，"
            f"有离职日非前任 {int(bad_closed.sum())}"
        )

    n_acting = int(out["PositionState"].eq("3").sum())
    if n_acting:
        print(f"警告：代任 PositionState=3 共 {n_acting} 条，仍按任职区间纳入在职集合")

    if out.duplicated(["MasterFundCode", "FundID", "FullName", "ServiceStartDate"]).any():
        raise SystemExit("经理表同一基金、同一人、同一任职日有多条记录")
    return out


def audit_business_duration(mgr: pd.DataFrame) -> None:
    """检查同一经理证券从业年限是否随任职开始日非降。"""
    sub = mgr.loc[mgr["BusinessDuration"].notna(), ["MasterFundCode", "FullName", "start", "BusinessDuration"]]
    # 同一基金同一姓名：多段任期
    fund_name_dec = 0
    fund_name_multi = 0
    for _, g in sub.groupby(["MasterFundCode", "FullName"], sort=False):
        if len(g) < 2:
            continue
        fund_name_multi += 1
        ordered = g.sort_values("start")["BusinessDuration"]
        if (ordered.diff().dropna() < 0).any():
            fund_name_dec += 1

    # 仅按姓名（可能同名不同人）
    name_multi = 0
    name_dec = 0
    name_flat_or_up = 0
    for _, g in sub.groupby("FullName", sort=False):
        if len(g) < 2:
            continue
        name_multi += 1
        ordered = g.sort_values("start")["BusinessDuration"]
        d = ordered.diff().dropna()
        if (d < 0).any():
            name_dec += 1
        else:
            name_flat_or_up += 1

    print(
        "从业年限检查："
        f"同基金同名多任期={fund_name_multi} 其中出现下降={fund_name_dec}；"
        f"仅按姓名多任期={name_multi} 出现下降={name_dec} 不降={name_flat_or_up}"
    )
    print(
        "说明：BusinessDuration 是 CSMAR 行内快照，约三成缺失；"
        "仅按姓名下降多因同名不同人，不能当作同一人年限倒退。"
    )


def active_team_stats(panel: pd.DataFrame, mgr: pd.DataFrame) -> pd.DataFrame:
    """按 fundid×date 聚合在职团队特征。"""
    keys = panel[["fundid", "基金ID", "date"]].copy()
    keys["asof"] = pd.to_datetime(keys["date"], format="%Y%m%d")

    paired = keys.merge(
        mgr,
        left_on=["基金ID", "fundid"],
        right_on=["FundID", "MasterFundCode"],
        how="left",
    )
    active = paired["start"].notna() & (paired["start"] <= paired["asof"]) & (
        paired["end"].isna() | (paired["end"] >= paired["asof"])
    )
    # 名单空窗：强制把指定经理算作该期在职
    gap = pd.Series(False, index=paired.index)
    for fundid, date, name in LIST_GAP_FILLS:
        hit = paired["fundid"].eq(fundid) & paired["date"].eq(date) & paired["FullName"].eq(name)
        if not hit.any():
            raise SystemExit(f"空窗补丁找不到经理行: {fundid} {date} {name}")
        gap |= hit
    n_gap = int((gap & ~active).sum())
    active = active | gap

    detail = paired.loc[active].copy()
    if detail.empty:
        return pd.DataFrame(columns=["fundid", "date", "Team", "Female", "Phd", "midAge", "largeAge"])

    detail["is_female"] = detail["Gender"].eq("女")
    detail["is_phd"] = detail["DegreeID"].eq("P0101") | detail["Degree"].fillna("").str.contains(
        "博士研究生", regex=False
    )

    stats = (
        detail.groupby(["fundid", "date"], sort=False)
        .agg(
            Team=("FullName", "size"),
            Female=("is_female", "mean"),
            Phd=("is_phd", "mean"),
            midAge=("BusinessDuration", "median"),
            largeAge=("BusinessDuration", "max"),
        )
        .reset_index()
    )
    print(f"名单空窗补丁纳入 {n_gap} 条经理×期（共 {len(LIST_GAP_FILLS)} 个指定空窗）")
    return stats


def build() -> None:
    panel = pd.read_excel(PANEL, dtype=str)
    n = len(panel)
    mgr = load_managers()
    audit_business_duration(mgr)
    stats = active_team_stats(panel, mgr)

    out = panel.merge(stats, on=["fundid", "date"], how="left", validate="one_to_one")
    if len(out) != n:
        raise SystemExit("合并后行数变化，键不唯一")

    for col in ("Team", "Female", "Phd", "midAge", "largeAge"):
        out[col] = pd.to_numeric(out[col], errors="coerce")
    for col in ("BuyValue", "SellValue", "TotalEquity", "TotalOperatingCost"):
        if col in out.columns:
            out[col] = pd.to_numeric(out[col], errors="coerce")
    for col in ("InstitutionRatio", "IndividualRatio", "PractitionersRatio"):
        if col in out.columns:
            out[col] = pd.to_numeric(out[col], errors="coerce")
    for col in ("发生转型代码重复", "发生过转型"):
        out[col] = out[col].astype(int)

    mgr_cols = ["Team", "Female", "Phd", "midAge", "largeAge"]
    cols = [c for c in out.columns if c not in mgr_cols]
    out = out[cols + mgr_cols]

    with pd.ExcelWriter(OUT, engine="openpyxl") as writer:
        out.to_excel(writer, index=False, sheet_name="fundid_date")
        style_sheet(
            writer.sheets["fundid_date"],
            out.columns,
            text_cols=TEXT_COLS,
            number_formats={
                "Team": "0",
                "Female": "0.0000",
                "Phd": "0.0000",
                "midAge": "0.00",
                "largeAge": "0.00",
                "BuyValue": "0.00",
                "SellValue": "0.00",
                "TotalOperatingCost": "0.00",
            },
            widths={
                "Team": 8,
                "Female": 10,
                "Phd": 10,
                "midAge": 10,
                "largeAge": 10,
            },
        )

    has_team = out["Team"].notna()
    no_team = ~has_team
    mid_miss = has_team & out["midAge"].isna()
    print(
        f"rows={len(out)} 有团队={int(has_team.sum())} 无人在职={int(no_team.sum())} "
        f"有团队但从业年限全缺={int(mid_miss.sum())}"
    )
    if has_team.any():
        vc = out.loc[has_team, "Team"].value_counts().sort_index()
        print("Team分布 " + " ".join(f"{int(k)}:{int(v)}" for k, v in vc.items()))
    s = out.loc[(out["fundid"] == "000006") & (out["date"] == "20190630")].iloc[0]
    if float(s["Team"]) != 1:
        raise SystemExit("000006 20190630 应为单人团队")
    for fundid, date, _name in LIST_GAP_FILLS:
        row = out.loc[(out["fundid"] == fundid) & (out["date"] == date)].iloc[0]
        if pd.isna(row["Team"]):
            raise SystemExit(f"空窗补丁后仍无团队: {fundid} {date}")
        print(
            f"核对补丁{fundid}@{date} Team={row['Team']} Female={row['Female']} "
            f"Phd={row['Phd']} midAge={row['midAge']} largeAge={row['largeAge']}"
        )
    print(f"写出 {OUT.name}")


if __name__ == "__main__":
    build()
