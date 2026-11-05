"""第三步：合并财报、经理、收益与评价指标，生成与旧版一致的半年面板。

按 # %% 分段；顺序运行 main。原始数据只读，中间结果独立保存。
本次为等价重构：保留旧版报告期、经理补丁、缺失值和预样本处理。
"""
import hashlib
import zipfile
from concurrent.futures import ProcessPoolExecutor
from pathlib import Path
import numpy as np
import pandas as pd
from openpyxl import load_workbook
from excel_style import style_sheet
from rolling_alpha import mark_sample

ROOT = Path(__file__).resolve().parent
RAW = ROOT / "原始数据"
WORK = ROOT / "中间数据"
RESULT = ROOT / "结果"
SIZE = RAW / "基金size.xlsx"
HOLDER = RAW / "机构持有占比.xlsx"
TURNOVER = RAW / "换手率计算.xlsx"
FEE = RAW / "费用.xlsx"
MGR = RAW / "Fund_FundManager.xlsx"
NAV = RAW / "月度净值.xlsx"
UNIT_CLASS_DIR = RAW / "CSMAR原始数据" / "基金份额信息"
NEG_TOL = -0.01

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


LIST_GAP_FILLS = [
    ("007449", "20201231", "董理"),
    ("007449", "20210630", "董理"),
    ("540007", "20110630", "廖志峰"),
]

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

ALPHAS = [
    ("capm_alpha", "滚动alpha_CAPM.xlsx"),
    ("ff3_alpha", "滚动alpha_三因子.xlsx"),
    ("car4_alpha", "滚动alpha_四因子.xlsx"),
    ("ff5_alpha", "滚动alpha_五因子.xlsx"),
    ("ch3_alpha", "滚动alpha_CH3.xlsx"),
    ("ch4_alpha", "滚动alpha_CH4.xlsx"),
]


MONTHS_PER_HALF = 6


NEW_COLS = ["raw_ret", *(name for name, _ in ALPHAS)]


PANEL_FORMATS = {
    "TotalEquity": "0.00",
    "InstitutionRatio": "0.0000",
    "IndividualRatio": "0.0000",
    "PractitionersRatio": "0.0000",
    "BuyValue": "0.00",
    "SellValue": "0.00",
    "TotalOperatingCost": "0.00",
    "Team": "0",
    "Female": "0.0000",
    "Phd": "0.0000",
    "midAge": "0.00",
    "largeAge": "0.00",
    **{c: "0.000000" for c in NEW_COLS},
}

CONTROLS = ["Size", "Age", "Turnover", "Expense", "Flow"]

SRC_GLOB = "基金日净值近半年评价文件*/FUND_EVL_DAYAHAY.xlsx"


ID_COLS = ["TradingDate", "Symbol", "FullName", "FundClassID"]


EVAL_TEXT = ["EvalDate", "CategoryID", "MktPortfolioID", "MktPortfolio"]


HEADER_ROWS = 2

# %% 1. 公共读写与报告期匹配

def checkpoint(frame, label, dtype):
    """保留旧链路每步的 Excel 序列化精度；省掉它会造成浮点尾数差异。"""
    path = WORK / f"面板_{label}.xlsx"
    frame.to_excel(path, index=False, sheet_name="fundid_date")
    print(f"{label}：{len(frame):,} 行")
    return pd.read_excel(path, dtype=dtype)


def merge_report(panel, report, code="MasterFundCode"):
    panel = panel.assign(end=pd.to_datetime(panel["date"]).dt.strftime("%Y-%m-%d"),
                         want_type=panel["date"].str[-4:].map({"0630": "5", "1231": "6"}))
    if panel["want_type"].isna().any():
        raise ValueError("面板日期不是半年末")
    out = panel.merge(report, left_on=["基金ID", "fundid", "end", "want_type"],
                      right_on=["FundID", code, "EndDate", "ReportTypeID"],
                      how="left", validate="one_to_one")
    return out.drop(columns=["end", "want_type", "FundID", "EndDate", "ReportTypeID"])


def numeric_panel(frame):
    for c in ("TotalEquity", "InstitutionRatio", "IndividualRatio", "PractitionersRatio",
              "BuyValue", "SellValue", "TotalOperatingCost"):
        if c in frame:
            frame[c] = pd.to_numeric(frame[c], errors="coerce")
    for c in ("发生转型代码重复", "发生过转型"):
        frame[c] = frame[c].astype(int)
    return frame


# %% 规模与持有人
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

# %% 半年买卖发生额：缺侧补零、起算不同或任一侧减负则整期空
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

# %% 半年费用：本期口径、年报减半年报
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

# %% 经理在职区间与团队聚合；保留旧版三个人工空窗补丁
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

# %% 半年收益与六模型 alpha：六个月齐全才计算
def half_date(year_month: pd.Series) -> pd.Series:
    """YYYY-MM → 所属半年末 YYYY0630 / YYYY1231。"""
    month = year_month.str[5:7].astype(int)
    return year_month.str[:4] + np.where(month <= 6, "0630", "1231")


def half_year_raw_ret() -> pd.DataFrame:
    """每个 (Symbol, 半年) 的有效月数与几何平均月收益；不足 6 个月置空。"""
    nav = pd.read_excel(NAV, dtype=str, usecols=["TradingDate", "Symbol", "AdjustedNAVGrowth"]).iloc[2:]
    nav = nav.apply(lambda s: s.str.strip())
    nav["year_month"] = nav["TradingDate"].str[:7]
    nav["AdjustedNAVGrowth"] = pd.to_numeric(nav["AdjustedNAVGrowth"], errors="coerce")
    nav = mark_sample(nav.sort_values(["Symbol", "year_month"], ignore_index=True))

    s = nav.loc[nav["sample"]].copy()
    s["date"] = half_date(s["year_month"])
    s["gross"] = 1 + s["AdjustedNAVGrowth"] / 100
    g = s.groupby(["Symbol", "date"])["gross"].agg(["prod", "count"])
    g["value"] = (g["prod"] ** (1 / MONTHS_PER_HALF) - 1).where(g["count"].eq(MONTHS_PER_HALF))
    return g[["value", "count"]]


def load_alpha(path: Path) -> pd.DataFrame:
    a = pd.read_excel(
        path,
        dtype={"Symbol": str, "FundClassID": str, "year_month": str},
        usecols=["Symbol", "FundClassID", "year_month", "alpha"],
    )
    a["date"] = half_date(a["year_month"])
    return a


def half_year_alpha(a: pd.DataFrame) -> pd.DataFrame:
    """每个 (Symbol, 半年) 的 alpha 月数与均值；不足 6 个月置空。"""
    g = a.groupby(["Symbol", "date"])["alpha"].agg(["mean", "count"])
    g["value"] = g["mean"].where(g["count"].eq(MONTHS_PER_HALF))
    return g[["value", "count"]]


def check_alpha_share(panel: pd.DataFrame, alpha: pd.DataFrame) -> None:
    """核对面板 (fundid, 基金ID) 与净值份额 (Symbol, FundClassID) 指向同一基金主体。"""
    zpath = next(UNIT_CLASS_DIR.glob("*.zip"))
    with zipfile.ZipFile(zpath) as z:
        unit = pd.read_csv(
            z.open("Fund_EVL_UnitClassInfo.csv"), dtype=str, usecols=["Symbol", "FundClassID", "FundID"]
        )
    unit = unit.apply(lambda s: s.str.strip())
    # 同一份额类别可能因转型有多段历史 → 多个 FundID，任一与面板一致即可
    share = alpha[["Symbol", "FundClassID"]].drop_duplicates().merge(unit, on=["Symbol", "FundClassID"])
    pairs = panel[["fundid", "基金ID"]].drop_duplicates()
    pairs = pairs.loc[pairs["fundid"].isin(alpha["Symbol"])]
    hit = pairs.merge(share, left_on=["fundid", "基金ID"], right_on=["Symbol", "FundID"], how="left", indicator=True)
    bad = hit.loc[hit["_merge"].eq("left_only"), ["fundid", "基金ID"]]
    if not bad.empty:
        raise SystemExit(f"净值份额与面板 基金ID 不一致：\n{bad.to_string(index=False)}")
    print(f"份额核对通过：{len(pairs)} 个有 alpha 的 fundid，净值份额 FundID 均与面板 基金ID 一致")


def merge_half(out: pd.DataFrame, g: pd.DataFrame, name: str) -> pd.DataFrame:
    """把半年值左连接到面板，并打印有值与留空原因。"""
    n = len(out)
    out = out.merge(
        g["value"].rename(name), left_on=["fundid", "date"], right_index=True, how="left", validate="one_to_one"
    )
    if len(out) != n:
        raise SystemExit(f"{name} 合并后行数变化")

    # 留空原因：6 个月不全 / 该半年无数据 / 整只基金无数据
    cnt = out[["fundid", "date"]].merge(g["count"], left_on=["fundid", "date"], right_index=True, how="left")["count"]
    fund_has = out["fundid"].isin(g.index.get_level_values("Symbol"))
    print(
        f"{name}: 有值={int(out[name].notna().sum())} 留空={int(out[name].isna().sum())}"
        f"（6个月不全={int(cnt.between(1, 5).sum())} 该半年无数据={int((cnt.isna() & fund_has).sum())}"
        f" 整只基金无数据={int((~fund_has).sum())}）"
    )
    return out

# %% 控制变量：相邻半年净资产与本期发生额
def prev_half_end(date: pd.Series) -> pd.Series:
    """YYYY0630 → 上年 1231；YYYY1231 → 同年 0630。"""
    year = date.str[:4].astype(int)
    return np.where(date.str.endswith("0630"), (year - 1).astype(str) + "1231", year.astype(str) + "0630")


def month_diff(date: pd.Series, start: pd.Series) -> pd.Series:
    """date（YYYYMMDD）与 start（YYYY-MM-DD）的自然月差。"""
    d, s = pd.to_datetime(date, format="%Y%m%d"), pd.to_datetime(start)
    return (d.dt.year - s.dt.year) * 12 + (d.dt.month - s.dt.month)


def add_controls(panel: pd.DataFrame) -> pd.DataFrame:
    out = panel.sort_values(["fundid", "date"], ignore_index=True)
    te = out["TotalEquity"]

    # 期初净资产：同一基金上一行，且必须是前一个半年末
    prev = out.groupby("fundid")[["date", "TotalEquity"]].shift()
    out["_has_prev"] = prev["date"].eq(prev_half_end(out["date"]))
    te0 = prev["TotalEquity"].where(out["_has_prev"])
    avg_te = (te0 + te) / 2

    out["Size"] = np.log(te)
    out["Age"] = np.log1p(month_diff(out["date"], out["成立时间_earliest"]))

    # 1231 行买卖起算日早于 7 月：分子是全年发生额，不能当下半年
    out["_annual_as_h2"] = out["date"].str.endswith("1231") & out["BuySellStartdate"].str[5:7].lt("07")
    out["Turnover"] = ((out["BuyValue"] + out["SellValue"]) / 2 / avg_te).mask(out["_annual_as_h2"])
    out["Expense"] = out["TotalOperatingCost"] / avg_te

    r = (1 + out["raw_ret"]) ** 6 - 1
    out["Flow"] = (te - (1 + r) * te0) / te
    out["_te0"] = te0
    return out

# %% CSMAR 评价：全批次去重、半年最后工作日
def file_md5(path: Path) -> str:
    h = hashlib.md5()
    with open(path, "rb") as f:
        for chunk in iter(lambda: f.read(1 << 20), b""):
            h.update(chunk)
    return h.hexdigest()


def read_half_end_rows(path: Path) -> pd.DataFrame:
    """流式读取一个评价文件，只保留 6 月、12 月下旬的行（半年最后一个交易日必在其中）。"""
    book = load_workbook(path, read_only=True)
    ws = book.active
    ws.reset_dimensions()  # 文件自带的表格范围不准，须重置后整行读取
    rows = ws.iter_rows(values_only=True)
    header = next(rows)
    date_index = header.index("TradingDate")
    for _ in range(HEADER_ROWS):
        next(rows)
    keep = []
    for row in rows:
        d = str(row[date_index])[:10]
        if d[5:7] in ("06", "12") and d[8:10] >= "20":
            keep.append(row)
    book.close()
    out = pd.DataFrame(keep, columns=header)
    out["TradingDate"] = out["TradingDate"].astype(str).str[:10]
    return out


def load_evaluation() -> pd.DataFrame:
    """读全部评价文件，取每个半年最后一个交易日的行。"""
    paths, seen = [], set()
    for p in sorted(RAW.glob(SRC_GLOB)):
        md5 = file_md5(p)
        if md5 in seen:
            print(f"跳过重复副本：{p.parent.name}")
            continue
        seen.add(md5)
        paths.append(p)

    with ProcessPoolExecutor(max_workers=6) as pool:
        parts = list(pool.map(read_half_end_rows, paths))
    ev = pd.concat(parts, ignore_index=True)
    for col in ("Symbol", "FundClassID", "CategoryID", "MktPortfolioID"):
        ev[col] = ev[col].astype(str).str.strip()

    # 半年最后一个交易日：该半年末月份内源数据出现的最大「工作日」。
    # CSMAR 在 6/30、12/31 落在周末时仍会给部分基金记一行（半年末披露净值），这些不是交易日，
    # 且只覆盖部分基金，用它会让排名不在同一截面，故剔除周末。
    # shortcut: 用工作日近似交易日；2009–2026 年各半年末无工作日休市，若扩展样本需改用交易所日历
    ev["date"] = ev["TradingDate"].str[:4] + ev["TradingDate"].str[5:7].map({"06": "0630", "12": "1231"})
    weekday = pd.to_datetime(ev["TradingDate"]).dt.dayofweek < 5
    last_day = ev["date"].map(ev.loc[weekday].groupby("date")["TradingDate"].max())
    ev = ev.loc[ev["TradingDate"].eq(last_day)]

    # 不同批次可能重叠：完全相同的行去重，同键仍多行则报错
    ev = ev.drop_duplicates()
    dup = ev.duplicated(["Symbol", "date"], keep=False)
    if dup.any():
        raise SystemExit(f"同一 Symbol × 半年末有不一致的多行：\n{ev.loc[dup].head(10).to_string()}")
    return ev.rename(columns={"TradingDate": "EvalDate"})


def check_eval_share(ev: pd.DataFrame) -> None:
    """评价文件的 (Symbol, FundClassID) 必须与月度净值一致。"""
    nav = pd.read_excel(NAV, dtype=str, usecols=["Symbol", "FundClassID"]).iloc[HEADER_ROWS:]
    nav = nav.apply(lambda s: s.str.strip()).drop_duplicates()
    m = ev[["Symbol", "FundClassID"]].drop_duplicates().merge(nav, on="Symbol", how="inner", suffixes=("", "_nav"))
    bad = m.loc[m["FundClassID"].ne(m["FundClassID_nav"])]
    if not bad.empty:
        raise SystemExit(f"评价文件与月度净值的 FundClassID 不一致：\n{bad.to_string(index=False)}")
    print(f"份额核对通过：{len(m)} 个 Symbol 的 FundClassID 与月度净值一致")

# %% 运行顺序：规模 → 持有人 → 买卖 → 费用 → 经理 → alpha → 控制 → 评价

def main():
    WORK.mkdir(exist_ok=True)
    RESULT.mkdir(exist_ok=True)
    panel = pd.read_excel(WORK / "基金名单_fundid_date.xlsx", dtype=str)
    out = merge_report(panel, load_size()).drop(columns="MasterFundCode")
    panel = checkpoint(numeric_panel(out), "规模", str)

    out = merge_report(panel, load_holder(), code="Symbol")
    chosen = choose_ratios(out)
    out = pd.concat([out.drop(columns=KEEP_FROM_HOLDER, errors="ignore"), chosen], axis=1)
    panel = checkpoint(numeric_panel(out), "持有人", str)

    amounts, stats = half_year_amounts(load_turnover())
    print("买卖口径异常及处理：", stats)
    out = merge_report(panel, amounts).drop(columns="MasterFundCode")
    out = numeric_panel(out)
    if (out["BuyValue"].isna() ^ out["SellValue"].isna()).any():
        raise ValueError("买卖金额出现单边缺失")
    panel = checkpoint(out, "买卖", str)

    fees, stats = half_year_fee(load_fee())
    print("费用口径异常及处理：", stats)
    out = merge_report(panel, fees).drop(columns="MasterFundCode")
    panel = checkpoint(numeric_panel(out), "费用", str)

    out = panel.merge(active_team_stats(panel, load_managers()), on=["fundid", "date"],
                      how="left", validate="one_to_one")
    panel = checkpoint(numeric_panel(out), "经理", {c: str for c in TEXT_COLS})

    out = merge_half(panel, half_year_raw_ret(), "raw_ret")
    for i, (name, filename) in enumerate(ALPHAS):
        alpha = load_alpha(WORK / filename)
        if i == 0:
            check_alpha_share(panel, alpha)
        out = merge_half(out, half_year_alpha(alpha), name)
    types = {**{c: str for c in TEXT_COLS}, "BuySellStartdate": str}
    panel = checkpoint(out, "Alpha", types)
    out = add_controls(panel)[[*panel.columns, *CONTROLS]]
    panel = checkpoint(out, "Controls", types)

    ev = load_evaluation()
    check_eval_share(ev)
    ev.to_excel(WORK / "评价_半年末.xlsx", index=False)
    cols = [c for c in ev if c not in (*ID_COLS, "date")]
    out = panel.merge(ev[["Symbol", "date", *cols]], left_on=["fundid", "date"],
                      right_on=["Symbol", "date"], how="left", validate="one_to_one").drop(columns="Symbol")
    path = RESULT / "基金名单_fundid_date_TotalEquity_Holder_BuySell_Fee_Manager_Alpha_Controls_Eval.xlsx"
    with pd.ExcelWriter(path, engine="openpyxl") as writer:
        out.to_excel(writer, index=False, sheet_name="fundid_date")
        style_sheet(writer.sheets["fundid_date"], out.columns, text_cols=(*TEXT_COLS, *EVAL_TEXT),
                    number_formats={**PANEL_FORMATS, **{c: "0.000000" for c in CONTROLS},
                                    **{c: "0.0000" for c in cols if c not in EVAL_TEXT}},
                    widths={"EvalDate": 12, "MktPortfolio": 28})
    print(f"完成：{out.shape}；评价有值 {out['EvalDate'].notna().sum():,}；{path}")


if __name__ == "__main__":
    main()
