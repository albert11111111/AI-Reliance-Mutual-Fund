"""从原始 CSMAR 压缩包构建股票半年面板；原包只读，口径见 README.md。"""

from hashlib import sha256
from io import BytesIO
import json
from pathlib import Path
import re
from zipfile import ZipFile

from lxml import etree
import numpy as np
import pandas as pd


HERE = Path(__file__).resolve().parent
RAW = HERE.parent
CACHE = HERE / "中间数据"
OUT = HERE / "结果"
QA = HERE / "核验"
END = pd.Timestamp("2026-06-30")
SEED = pd.Timestamp("2009-12-31")
A_MARKETS = [1, 4, 16, 32, 64]
EXTRA_METRICS = ["pe_ttm", "ep_ttm", "pb", "roe_ttm", "leverage", "asset_growth_yoy",
                 "revenue_yoy", "nonoperating_profit_share_ttm", "rd_expense_to_revenue_ttm",
                 "tradable_cap_ratio", "loss_ttm"]
EXTRA_AVAILABILITY = {"pe_ttm": "earnings_ttm_available", "ep_ttm": "earnings_ttm_available",
    "loss_ttm": "earnings_ttm_available", "pb": "book_source_available", "roe_ttm": "roe_source_available",
    "leverage": "book_source_available", "asset_growth_yoy": "balance_yoy_available",
    "revenue_yoy": "revenue_yoy_available", "nonoperating_profit_share_ttm": "earnings_ttm_available",
    "rd_expense_to_revenue_ttm": "earnings_ttm_available"}
SOURCES = []
AUDIT = {}
NS = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"


def unique(df, keys, label):
    if df.duplicated(keys).any():
        raise ValueError(f"{label} 存在重复键 {keys}")


def read_xlsx(data, columns):
    # 原始 Excel 的 dimension 错写成 A 列；直接流式读取 XML，避免截列及百万行内存膨胀。
    frames = []
    with ZipFile(BytesIO(data)) as z:
        shared = []
        if "xl/sharedStrings.xml" in z.namelist():
            with z.open("xl/sharedStrings.xml") as f:
                for _, node in etree.iterparse(f, events=("end",), tag=NS + "si"):
                    shared.append("".join(node.itertext()))
                    node.clear()
                    while node.getprevious() is not None:
                        del node.getparent()[0]
        for sheet in sorted(n for n in z.namelist() if re.fullmatch(r"xl/worksheets/sheet\d+\.xml", n)):
            rows, positions = [], None
            with z.open(sheet) as f:
                for _, node in etree.iterparse(f, events=("end",), tag=NS + "row"):
                    values = {}
                    for cell in node:
                        letter = re.sub(r"\d", "", cell.get("r"))
                        if positions is not None and letter not in positions:
                            continue
                        if cell.get("t") == "inlineStr":
                            value = "".join(cell.itertext())
                        else:
                            value = cell.findtext(NS + "v")
                            if value is not None and cell.get("t") == "s":
                                value = shared[int(value)]
                        values[letter] = value
                    if positions is None:
                        positions = {k: v for k, v in values.items() if v in columns}
                        if set(positions.values()) != set(columns):
                            raise ValueError(f"Excel 缺列：{set(columns) - set(positions.values())}")
                    elif int(node.get("r")) > 3:
                        rows.append({name: values.get(letter) for letter, name in positions.items()})
                    node.clear()
                    while node.getprevious() is not None:
                        del node.getparent()[0]
            frames.append(pd.DataFrame(rows, columns=columns))
    return pd.concat(frames, ignore_index=True)


def load(prefix, columns):
    paths = list(RAW.glob("0*/*" + prefix + "*.zip"))
    if len(paths) != 1:
        raise ValueError(f"原始包匹配不唯一：{prefix}")
    path = paths[0]
    digest = sha256(path.read_bytes()).hexdigest()
    meta = {"source": str(path.relative_to(RAW)), "sha256": digest,
            "columns": columns, "reader_version": 1}
    cache = CACHE / (prefix + ".parquet")
    meta_path = cache.with_suffix(".json")
    if cache.exists() and meta_path.exists() and json.loads(meta_path.read_text()) == meta:
        df = pd.read_parquet(cache)
    else:
        print(f"读取 {prefix}", flush=True)
        with ZipFile(path) as z:
            parts = sorted(n for n in z.namelist() if n.endswith((".csv", ".xlsx")))
            frames = []
            for name in parts:
                if name.endswith(".xlsx"):
                    frames.append(read_xlsx(z.read(name), columns))
                else:
                    with z.open(name) as f:
                        frames.append(pd.read_csv(f, usecols=columns, dtype=str, encoding="utf-8-sig"))
            df = pd.concat(frames, ignore_index=True)
        df = df.replace("", np.nan)
        df.to_parquet(cache, index=False)
        meta_path.write_text(json.dumps(meta, ensure_ascii=False, indent=2))
    SOURCES.append({**meta, "rows": len(df)})
    return df.copy()


def dates(values):
    return pd.to_datetime(values, errors="coerce", format="%Y-%m-%d")


def codes(values):
    return values.astype("string").str.strip().str.replace(r"\.0$", "", regex=True).str.zfill(6)


def half_end(values):
    return pd.to_datetime(values.dt.year.astype(str) + np.where(values.dt.month <= 6, "-06-30", "-12-31"))


def lag_date(values):
    return values - pd.offsets.MonthEnd(6)


def numeric(df, columns):
    for c in columns:
        df[c] = pd.to_numeric(df[c], errors="coerce")
    return df


def export(df, name, directory=OUT):
    df.to_parquet(directory / (name + ".parquet"), index=False)
    df.to_csv(directory / (name + ".csv"), index=False, encoding="utf-8-sig", date_format="%Y-%m-%d")


def market_panel():
    company = load("公司文件111337729", ["Stkcd", "Listdt", "Statco", "Statdt", "Markettype", "FormerCode", "Crcd", "Commnt"])
    company["stock_code"] = codes(company.Stkcd)
    unique(company, ["stock_code"], "公司快照")
    # 当前公司状态/行业不回填历史；曾用代码仅作为需要核验的匹配线索。
    aliases = company.loc[company.FormerCode.notna(), ["stock_code", "FormerCode", "Listdt", "Markettype"]]
    aliases.to_csv(QA / "曾用代码待核验.csv", index=False, encoding="utf-8-sig")
    m = load("月个股回报率文件114059067", ["Stkcd", "Trdmnt", "Clsdt", "Mclsprc", "Msmvttl", "Msmvosd", "Mretwd", "Mretnd", "Markettype", "Ndaytrd"])
    m["stock_code"] = codes(m.Stkcd)
    m["month"] = pd.to_datetime(m.Trdmnt, format="%Y-%m") + pd.offsets.MonthEnd(0)
    numeric(m, ["Mclsprc", "Msmvttl", "Msmvosd", "Mretwd", "Mretnd", "Markettype", "Ndaytrd"])
    AUDIT["market_rows_by_type"] = {str(k): int(v) for k, v in m.Markettype.value_counts().items()}
    before_bse = m.Markettype.eq(64) & m.month.lt(pd.Timestamp("2021-11-01"))
    AUDIT["excluded_pre_bse_months"] = int(before_bse.sum())
    m = m.loc[m.Markettype.isin(A_MARKETS) & ~before_bse & (m.month <= END)].copy()
    unique(m, ["stock_code", "month"], "A股月行情")
    if m.Mretwd.dropna().lt(-1).any():
        raise ValueError("月收益低于 -100%，需检查供应商缺失值编码")
    m["asof"] = half_end(m.month)
    bounds = m.groupby("stock_code")["asof"].agg(["min", "max"])
    grid = pd.DataFrame({"asof": pd.date_range(SEED, END, freq="2QE-DEC")})
    panel = bounds.reset_index().merge(grid, how="cross")
    panel = panel.loc[panel["asof"].between(panel["min"], panel["max"]), ["stock_code", "asof"]]
    # 不用最后一个有交易月份冒充半年末市值；半年度内缺月/缺回报时不计算完整半年收益。
    endpoint = m.loc[m.month.eq(m["asof"]), ["stock_code", "asof", "Mclsprc", "Msmvttl", "Msmvosd", "Markettype", "Clsdt", "Ndaytrd"]]
    endpoint = endpoint.rename(columns={"Mclsprc": "price", "Markettype": "market_type", "Ndaytrd": "end_month_trade_days"})
    endpoint["market_cap"] = endpoint.pop("Msmvttl") * 1000
    endpoint["float_market_cap"] = endpoint.pop("Msmvosd") * 1000
    endpoint["price_date"] = dates(endpoint["asof"].dt.strftime("%Y-%m-") + endpoint.pop("Clsdt").str.zfill(2))
    m["gross_return"] = 1 + m.Mretwd
    ret = m.groupby(["stock_code", "asof"]).agg(return_months=("Mretwd", "count"),
        observed_months=("month", "size"), ret6m=("gross_return", "prod")).reset_index()
    ret["ret6m"] = (ret.ret6m - 1).where(ret.return_months.eq(6))
    panel = panel.merge(endpoint, how="left", on=["stock_code", "asof"], validate="1:1")
    panel = panel.merge(ret, how="left", on=["stock_code", "asof"], validate="1:1")
    panel["size"] = np.log1p(panel.market_cap.where(panel.market_cap > 0) / 1_000_000)
    panel["ret6m_fwd"] = panel.merge(
        panel[["stock_code", "asof", "ret6m"]].assign(asof=lambda x: lag_date(x["asof"])),
        on=["stock_code", "asof"], how="left", suffixes=("", "_next"), validate="1:1")["ret6m_next"].to_numpy()
    company["cross_listed_snapshot"] = company.Crcd.notna() | company.Commnt.notna()
    panel = panel.merge(company[["stock_code", "cross_listed_snapshot"]], how="left", on="stock_code", validate="m:1")
    changed_codes = set(aliases.stock_code) | set(aliases.FormerCode)
    panel["code_change_review"] = panel.stock_code.isin(changed_codes)
    # 北交所换号是标识规范化；只映射原行情中已不存在的旧码，不推断换号生效日。
    bridge = pd.DataFrame({"holding_code": sorted(m.stock_code.unique())})
    bridge["stock_code"] = bridge.holding_code
    bridge["mapping_basis"] = "原始行情同代码"
    bse = aliases.loc[aliases.stock_code.str.startswith("920") & aliases.Markettype.eq("64") &
        aliases.FormerCode.str.fullmatch(r"(?:43|83|87)\d{4}") & ~aliases.FormerCode.isin(m.stock_code)]
    bse = bse.rename(columns={"FormerCode": "holding_code"})[["holding_code", "stock_code"]]
    bse["mapping_basis"] = "CSMAR曾用代码；北交所920换号；不代表历史生效日期"
    bridge = pd.concat([bridge, bse], ignore_index=True)
    unique(bridge, ["holding_code"], "持仓代码桥")
    bridge.to_csv(OUT / "stock_code_bridge.csv", index=False, encoding="utf-8-sig")
    AUDIT["bse_aliases_for_matching"] = len(bse)
    AUDIT["monthly_a_stocks"] = int(m.stock_code.nunique())
    m.drop(columns="gross_return").to_parquet(CACHE / "股票月度行情.parquet", index=False)
    return panel.reset_index(drop=True)


def prepare_financial(prefix, value_map):
    cols = ["Stkcd", "Accper", "Typrep", "IfCorrect", "DeclareDate", *value_map]
    f = load(prefix, cols)
    f["stock_code"] = codes(f.Stkcd)
    f["period"] = dates(f.Accper)
    keep = f.Typrep.eq("A") & f.period.dt.strftime("%m-%d").isin(["03-31", "06-30", "09-30", "12-31"])
    AUDIT[prefix + "_excluded_nonconsolidated_or_nonquarter"] = int((~keep).sum())
    f = f.loc[keep].rename(columns=value_map).copy()
    numeric(f, list(value_map.values()))
    # 当前下载只有一个数值版本；更正后数值只能在最后一次更正公告之后使用。
    def last_correction(s):
        if pd.isna(s):
            return pd.NaT
        parts = [p.strip() for p in str(s).split(",")]
        try:
            return max(pd.Timestamp(p) for p in parts)
        except ValueError:
            return pd.NaT
    f["correction_date"] = f.DeclareDate.map(last_correction)
    f["corrected"] = f.IfCorrect.eq("1") | f.DeclareDate.notna()
    unique(f, ["stock_code", "period"], prefix)
    return f[["stock_code", "period", "corrected", "correction_date", *value_map.values()]]


def financial_panel(panel):
    reports = load("年、中、季报基本情况文件114653443", ["Stkcd", "Reptyp", "Accper", "Annodt"])
    reports["stock_code"] = codes(reports.Stkcd)
    reports["period"] = dates(reports.Accper)
    reports["announcement"] = dates(reports.Annodt)
    valid = reports.period.notna() & reports.announcement.notna() & (reports.announcement >= reports.period)
    valid &= reports.Reptyp.eq(reports.period.dt.month.div(3).astype("Int64").astype(str))
    AUDIT["invalid_report_publication"] = int((~valid).sum())
    reports.loc[~valid].to_csv(QA / "财报公告日期异常.csv", index=False, encoding="utf-8-sig")
    reports = reports.loc[valid, ["stock_code", "period", "announcement"]]
    unique(reports, ["stock_code", "period"], "财报公告")
    bs = prepare_financial("资产负债表114132694", {"A001000000": "assets", "A003100000": "parent_equity", "A003000000": "total_equity"})
    inc = prepare_financial("利润表114149331", {"B002000000": "net_income", "B002000101": "parent_income",
        "B001100000": "revenue", "B001300000": "operating_profit", "B001000000": "total_profit", "B001216000": "rd_expense"})
    all_fin = {}
    for name, df in [("bs", bs), ("inc", inc)]:
        df = df.merge(reports, how="left", on=["stock_code", "period"], validate="1:1")
        df["available"] = df[["announcement", "correction_date"]].max(axis=1)
        invalid = df.announcement.isna() | (df.corrected & df.correction_date.isna())
        df.loc[invalid, "available"] = pd.NaT
        all_fin[name] = df
        df.to_parquet(CACHE / (name + "_财报可用时间.parquet"), index=False)
        AUDIT[name + "_corrections"] = int(df.corrected.sum())
    # 选择截至时点已经公告的最新报告期，不因该期数据缺失而偷偷改选旧报表。
    reports = reports.sort_values(["stock_code", "announcement", "period"])
    reports["latest_period"] = reports.groupby("stock_code").period.cummax()
    reports = reports.loc[reports.period.eq(reports.latest_period)].drop_duplicates(["stock_code", "announcement"], keep="last")
    selected = pd.merge_asof(panel[["stock_code", "asof"]].sort_values("asof"),
        reports.sort_values("announcement"), by="stock_code", left_on="asof", right_on="announcement", direction="backward")
    selected = selected.drop(columns="latest_period")

    def attach(frame, source, period, suffix):
        right = all_fin[source].drop(columns="announcement").rename(columns={c: c + suffix for c in all_fin[source] if c not in ["stock_code", "announcement"]})
        return frame.merge(right, left_on=["stock_code", period], right_on=["stock_code", "period" + suffix], how="left", validate="m:1")

    f = attach(selected, "bs", "period", "_bs")
    f = attach(f, "inc", "period", "_inc")
    f["period_yearago"] = f.period - pd.DateOffset(years=1)
    f["period_prevyear"] = pd.to_datetime((f.period.dt.year - 1).astype("Int64").astype(str) + "-12-31", errors="coerce")
    f = attach(f, "bs", "period_yearago", "_bs_old")
    f = attach(f, "inc", "period_yearago", "_inc_old")
    f = attach(f, "inc", "period_prevyear", "_inc_annual")
    valid_bs = f.available_bs.notna() & (f.available_bs <= f["asof"])
    annual = f.period.dt.month.eq(12)
    f["income_ttm"] = np.where(annual, f.net_income_inc,
        f.net_income_inc + f.net_income_inc_annual - f.net_income_inc_old)
    f["roa_source_available"] = f[["available_bs", "available_bs_old", "available_inc", "available_inc_old", "available_inc_annual"]].max(axis=1)
    f.loc[annual, "roa_source_available"] = f.loc[annual, ["available_bs", "available_bs_old", "available_inc"]].max(axis=1)
    required = ["available_bs", "available_bs_old", "available_inc"]
    valid_roa = f[required].notna().all(axis=1)
    valid_roa &= annual | f[["available_inc_old", "available_inc_annual"]].notna().all(axis=1)
    valid_roa &= f.roa_source_available.le(f["asof"])
    valid_roa &= f.assets_bs.gt(0) & f.assets_bs_old.gt(0)
    f["roa_average_assets"] = ((f.assets_bs + f.assets_bs_old) / 2).where(valid_roa)
    f["roa"] = (f.income_ttm / f.roa_average_assets).where(valid_roa)
    f["income_ttm"] = f.income_ttm.where(valid_roa)
    f["book_equity"] = f.parent_equity_bs.where(valid_bs)
    f["assets"] = f.assets_bs.where(valid_bs)
    f["bs_correction_blocked"] = f.corrected_bs.eq(True) & (~valid_bs)
    f["roa_correction_blocked"] = False
    for suffix in ["_bs", "_bs_old", "_inc", "_inc_old", "_inc_annual"]:
        needed = ~annual if suffix in ["_inc_old", "_inc_annual"] else pd.Series(True, index=f.index)
        f["roa_correction_blocked"] |= needed & f["corrected" + suffix].eq(True) & (
            f["available" + suffix].isna() | f["available" + suffix].gt(f["asof"]))
    # PE 的归母利润不依赖资产数据，不能沿用 ROA 对净利润施加的资产可得性筛选。
    income_dates = ["available_inc", "available_inc_old", "available_inc_annual"]
    f["earnings_ttm_available"] = f[income_dates].max(axis=1)
    f.loc[annual, "earnings_ttm_available"] = f.loc[annual, "available_inc"]
    valid_income = f.available_inc.notna() & (annual | f[income_dates].notna().all(axis=1))
    valid_income &= f.earnings_ttm_available.le(f["asof"])
    ttm_fields = ["parent_income", "revenue", "operating_profit", "total_profit", "rd_expense"]
    for c in ttm_fields:
        value = pd.Series(np.where(annual, f[c + "_inc"],
            f[c + "_inc"] + f[c + "_inc_annual"] - f[c + "_inc_old"]), index=f.index)
        f[c + "_ttm"] = value.where(valid_income)
    # 研发费用自 2018 年起单列；不从早期空值反推零研发。
    f["rd_expense_ttm"] = f.rd_expense_ttm.where(f.period.ge(pd.Timestamp("2018-12-31")))
    f["balance_yoy_available"] = f[["available_bs", "available_bs_old"]].max(axis=1)
    valid_balance_yoy = f[["available_bs", "available_bs_old"]].notna().all(axis=1) & f.balance_yoy_available.le(f["asof"])
    f["roe_source_available"] = f[["balance_yoy_available", "earnings_ttm_available"]].max(axis=1)
    valid_roe = valid_balance_yoy & valid_income & f.parent_equity_bs.gt(0) & f.parent_equity_bs_old.gt(0)
    f["roe_average_equity"] = ((f.parent_equity_bs + f.parent_equity_bs_old) / 2).where(valid_roe)
    f["roe_ttm"] = f.parent_income_ttm / f.roe_average_equity
    f["leverage"] = (1 - f.total_equity_bs / f.assets_bs).where(valid_bs & f.assets_bs.gt(0))
    f["asset_growth_yoy"] = (f.assets_bs / f.assets_bs_old - 1).where(valid_balance_yoy & f.assets_bs_old.gt(0))
    f["revenue_yoy_available"] = f[["available_inc", "available_inc_old"]].max(axis=1)
    valid_yoy = f[["available_inc", "available_inc_old"]].notna().all(axis=1) & f.revenue_yoy_available.le(f["asof"])
    f["revenue_yoy"] = (f.revenue_inc / f.revenue_inc_old - 1).where(valid_yoy & f.revenue_inc_old.gt(0))
    f["nonoperating_profit_share_ttm"] = (f.total_profit_ttm - f.operating_profit_ttm) / f.total_profit_ttm.abs().where(f.total_profit_ttm.ne(0))
    f["rd_expense_to_revenue_ttm"] = f.rd_expense_ttm.where(f.rd_expense_ttm.ge(0)) / f.revenue_ttm.where(f.revenue_ttm.gt(0))
    selected_cols = ["stock_code", "asof", "period", "announcement", "available_bs", "roa_source_available",
        "book_equity", "assets", "income_ttm", "roa_average_assets", "roa", "bs_correction_blocked", "roa_correction_blocked"]
    selected_cols += [c + "_ttm" for c in ttm_fields] + ["roe_average_equity", "earnings_ttm_available",
        "balance_yoy_available", "revenue_yoy_available", "roe_source_available", "roe_ttm", "leverage",
        "asset_growth_yoy", "revenue_yoy", "nonoperating_profit_share_ttm", "rd_expense_to_revenue_ttm"]
    f = f[selected_cols].rename(columns={"period": "fin_report_period", "announcement": "fin_announcement", "available_bs": "book_source_available"})
    panel = panel.merge(f, how="left", on=["stock_code", "asof"], validate="1:1")
    panel["bm"] = panel.book_equity / panel.market_cap.where(panel.market_cap > 0)
    market_cap = panel.market_cap.where(panel.market_cap > 0)
    panel["pe_ttm"] = market_cap / panel.parent_income_ttm.where(panel.parent_income_ttm > 0)
    panel["ep_ttm"] = panel.parent_income_ttm / market_cap
    panel["pb"] = market_cap / panel.book_equity.where(panel.book_equity > 0)
    panel["loss_ttm"] = panel.parent_income_ttm.lt(0).astype(float).where(panel.parent_income_ttm.notna())
    panel["negative_book_equity"] = panel.book_equity.lt(0)
    panel["fin_age_days"] = (panel["asof"] - panel.fin_report_period).dt.days
    panel["financial_extreme_review"] = panel.roa.abs().gt(1) | panel.bm.abs().gt(10)
    AUDIT["book_blocked_by_correction"] = int(panel.bs_correction_blocked.sum())
    AUDIT["roa_blocked_by_correction"] = int(panel.roa_correction_blocked.sum())
    return panel


def rating_panel(panel):
    cols = ["Stkcd", "ReportID", "Rptdt", "DeclareDate", "InstitutionID", "Brokern", "Stdrank"]
    prefixes = ["分析师评级及变动114629474", "分析师评级及变动114638471", "分析师评级及变动115810968", "分析师评级及变动114701530"]
    r = pd.concat([load(p, cols) for p in prefixes], ignore_index=True)
    AUDIT["rating_raw_rows"] = len(r)
    r["stock_code"] = codes(r.Stkcd)
    r["report_date"] = dates(r.Rptdt)
    r["publication"] = dates(r.DeclareDate)
    r["available"] = r[["report_date", "publication"]].max(axis=1)
    r["score"] = r.Stdrank.map({"买入": 5, "增持": 4, "中性": 3, "减持": 2, "卖出": 1})
    AUDIT["rating_missing_publication_rows"] = int(r.publication.isna().sum())
    AUDIT["rating_nonstandard_or_missing_score_rows"] = int(r.score.isna().sum())
    names = r[["Brokern", "InstitutionID"]].dropna().drop_duplicates()
    names = names.loc[~names.Brokern.duplicated(keep=False)].set_index("Brokern").InstitutionID
    r["broker"] = "id:" + r.InstitutionID.fillna(r.Brokern.map(names)).astype("string")
    r["broker"] = r.broker.fillna("name:" + r.Brokern.astype("string"))
    usable = r.publication.notna() & r.score.notna() & r.broker.notna() & r.ReportID.notna()
    r = r.loc[usable, ["stock_code", "ReportID", "broker", "report_date", "publication", "available", "score"]].drop_duplicates()
    # 重复的预测年度不重复投票；同一报告出现冲突则整份报告退出主口径。
    conflict = r.duplicated(["stock_code", "ReportID"], keep=False)
    r.loc[conflict].to_csv(QA / "评级报告冲突.csv", index=False, encoding="utf-8-sig")
    AUDIT["rating_conflicting_report_rows"] = int(conflict.sum())
    r = r.loc[~conflict].sort_values(["available", "report_date", "ReportID"])
    r.to_parquet(CACHE / "去重评级报告.parquet", index=False)
    AUDIT["rating_distinct_usable_reports"] = len(r)
    AUDIT["rating_publication_before_report"] = int(r.publication.lt(r.report_date).sum())
    # 同券商同日多份报告先取最新撰写日；仍有多份时取中位评级，不以报告 ID 大小猜先后。
    r = r.loc[r.report_date.eq(r.groupby(["stock_code", "broker", "available"]).report_date.transform("max"))]
    AUDIT["rating_broker_day_score_conflicts"] = int(r.groupby(["stock_code", "broker", "available"]).score.nunique().gt(1).sum())
    r = r.groupby(["stock_code", "broker", "available", "report_date"], as_index=False).score.median().sort_values("available")
    snapshots = []
    cutoffs = sorted(set(panel["asof"]) | {SEED - pd.offsets.MonthEnd(6)})
    for cutoff in cutoffs:
        cutoff = pd.Timestamp(cutoff)
        recent = r.loc[r.available.between(cutoff - pd.Timedelta(days=180), cutoff) & r.report_date.ge(cutoff - pd.Timedelta(days=180))]
        latest = recent.drop_duplicates(["stock_code", "broker"], keep="last")
        snap = latest.groupby("stock_code").agg(rec_consensus=("score", "median"), rec_brokers=("broker", "size"),
            rec_latest_publication=("available", "max")).reset_index()
        snap["asof"] = cutoff
        snapshots.append(snap)
    snapshot = pd.concat(snapshots, ignore_index=True)
    previous = snapshot[["stock_code", "asof", "rec_consensus"]].copy()
    previous["asof"] += pd.offsets.MonthEnd(6)
    previous = previous.rename(columns={"rec_consensus": "rec_consensus_previous"})
    snapshot = snapshot.merge(previous, on=["stock_code", "asof"], how="left", validate="1:1")
    snapshot["rec_change"] = snapshot.rec_consensus - snapshot.pop("rec_consensus_previous")
    panel = panel.merge(snapshot, on=["stock_code", "asof"], how="left", validate="1:1")
    panel["rec_brokers"] = panel.rec_brokers.fillna(0).astype(int)
    return panel


def distribution_panel(panel):
    d = load("分配文件114014091", ["Stkcd", "Disttyp", "Exdistdt", "Annodt", "Paydt", "Amount", "Roprc", "Markettype"])
    d["stock_code"] = codes(d.Stkcd)
    for raw, new in [("Exdistdt", "ex_date"), ("Annodt", "announcement"), ("Paydt", "share_list_date")]:
        d[new] = dates(d[raw])
    numeric(d, ["Amount", "Roprc", "Markettype"])
    d = d.loc[d.Markettype.isin(A_MARKETS)].drop_duplicates().copy()
    # 只自动处理 SD 每股新增股数；DS、股改对价及配股参与不能从价格因子推断。
    d["supported_bonus"] = d.Disttyp.eq("SD") & d.Amount.ge(0) & d.ex_date.notna()
    duplicate_sd = d.Disttyp.eq("SD") & d.duplicated(["stock_code", "Disttyp", "ex_date"], keep=False)
    if duplicate_sd.any():
        raise ValueError("同股票同除权日存在多条送股记录，需先确认是否同一分配方案")
    d["bonus_multiplier"] = (1 + d.Amount).where(d.supported_bonus)
    d.to_parquet(CACHE / "分配事件.parquet", index=False)
    d.loc[d.Disttyp.isin(["DS", "GQ", "RO", "SN"])].to_csv(QA / "需核验的特殊股数事件.csv", index=False, encoding="utf-8-sig")
    events = d.loc[d.ex_date.notna()].copy()
    events["asof"] = half_end(events.ex_date)
    events["sd_multiplier"] = events.bonus_multiplier.fillna(1)
    events["sd_count"] = events.Disttyp.eq("SD").astype(int)
    events["unresolved_event"] = events.Disttyp.isin(["DS", "GQ", "RO", "SN"]) | (events.Disttyp.eq("SD") & ~events.supported_bonus)
    events["publication_issue"] = ~events.Disttyp.eq("CA") & (events.announcement.isna() | events.announcement.gt(events.ex_date))
    events["bonus_listing_cross_half"] = events.Disttyp.eq("SD") & (
        events.share_list_date.isna() | events.share_list_date.gt(events["asof"]))
    summary = events.groupby(["stock_code", "asof"]).agg(
        bonus_share_factor_6m=("sd_multiplier", "prod"), bonus_event_count=("sd_count", "sum"),
        special_share_event_count=("unresolved_event", "sum"), event_publication_issue=("publication_issue", "max"),
        bonus_listing_cross_half=("bonus_listing_cross_half", "max")).reset_index()
    panel = panel.merge(summary, on=["stock_code", "asof"], how="left", validate="1:1")
    for c in ["bonus_event_count", "special_share_event_count"]:
        panel[c] = panel[c].fillna(0).astype(int)
    panel["bonus_share_factor_6m"] = panel.bonus_share_factor_6m.fillna(1)
    for c in ["event_publication_issue", "bonus_listing_cross_half"]:
        panel[c] = panel[c].eq(True)
    # 除权和新增股份上市跨半年时，两端持仓期都需要核对基金是否计入待上市股份。
    crossing = events.loc[events.bonus_listing_cross_half & events.share_list_date.notna(), ["stock_code", "share_list_date"]].copy()
    crossing["asof"] = half_end(crossing.share_list_date)
    crossing_keys = pd.MultiIndex.from_frame(crossing[["stock_code", "asof"]])
    panel["bonus_listing_cross_half"] |= pd.MultiIndex.from_frame(panel[["stock_code", "asof"]]).isin(crossing_keys)
    panel["share_adjustment_review"] = panel.special_share_event_count.gt(0) | panel.event_publication_issue | panel.bonus_listing_cross_half
    AUDIT["share_events_missing_ex_date"] = int((~d.Disttyp.eq("CA") & d.ex_date.isna()).sum())
    d.loc[~d.Disttyp.eq("CA") & d.ex_date.isna()].to_csv(QA / "无法按除权日调度的股数事件.csv", index=False, encoding="utf-8-sig")
    # 价格复权因子仅留作核对；明确不用于持股数量调整。
    p = load("股票价格复权因子表(日)114928367", ["Symbol", "TradingDate", "CumulateBwardFactor"])
    p["stock_code"] = codes(p.Symbol)
    p["factor_date"] = dates(p.TradingDate)
    p["price_factor_bwd"] = pd.to_numeric(p.CumulateBwardFactor, errors="coerce")
    unique(p, ["stock_code", "factor_date"], "价格因子")
    match = pd.merge_asof(panel[["stock_code", "asof"]].sort_values("asof"),
        p[["stock_code", "factor_date", "price_factor_bwd"]].dropna(subset="factor_date").sort_values("factor_date"),
        left_on="asof", right_on="factor_date", by="stock_code", direction="backward")
    panel = panel.merge(match, on=["stock_code", "asof"], how="left", validate="1:1")
    return panel


def capital_and_status(panel):
    c = load("股本变动文件114008558", ["Stkcd", "Shrchgdt", "Shrtyp", "Nshrttl", "Nshra"])
    c["stock_code"] = codes(c.Stkcd)
    c["capital_date"] = dates(c.Shrchgdt)
    numeric(c, ["Nshrttl", "Nshra"])
    c = c.dropna(subset="capital_date").rename(columns={"Nshrttl": "total_shares", "Nshra": "float_a_shares", "Shrtyp": "capital_change_type"})
    unique(c, ["stock_code", "capital_date"], "股本事件")
    match = pd.merge_asof(panel[["stock_code", "asof"]].sort_values("asof"),
        c[["stock_code", "capital_date", "total_shares", "float_a_shares", "capital_change_type"]].sort_values("capital_date"),
        left_on="asof", right_on="capital_date", by="stock_code", direction="backward")
    panel = panel.merge(match, on=["stock_code", "asof"], how="left", validate="1:1")
    s = load("上市公司上市状态变更表091453528", ["Symbol", "DeclareDate", "ChangeDate", "ChangedItem", "ValueBefore", "ValueAfter", "ReasonID"])
    s["stock_code"] = codes(s.Symbol)
    s["status_date"] = dates(s.ChangeDate)
    s["status_announcement"] = dates(s.DeclareDate)
    s["status_available"] = s[["status_date", "status_announcement"]].max(axis=1)
    s = s.loc[s.ChangedItem.eq("上市状态") & s.status_date.notna() & s.status_announcement.notna()].copy()
    s = s.rename(columns={"ValueAfter": "last_status_event"})
    s = s.sort_values(["stock_code", "status_available", "status_date"])
    s = s.loc[s.status_date.eq(s.groupby("stock_code").status_date.cummax())]
    cols = ["stock_code", "status_date", "status_announcement", "status_available", "last_status_event"]
    s = s[cols].drop_duplicates()
    conflict = s.duplicated(["stock_code", "status_available"], keep=False)
    s.loc[conflict].to_csv(QA / "上市状态冲突.csv", index=False, encoding="utf-8-sig")
    s = s.loc[~conflict]
    match = pd.merge_asof(panel[["stock_code", "asof"]].sort_values("asof"), s.sort_values("status_available"),
        left_on="asof", right_on="status_available", by="stock_code", direction="backward")
    return panel.merge(match, on=["stock_code", "asof"], how="left", validate="1:1")


def finish(panel):
    panel = panel.sort_values(["stock_code", "asof"]).reset_index(drop=True)
    keys = ["stock_code", "asof"]
    core = ["size", "bm", "roa", "ret6m", "rec_change"]
    panel["controls_complete"] = panel[core].notna().all(axis=1)
    panel["date"] = panel["asof"].dt.strftime("%Y%m%d").astype(int)
    panel["is_formal_sample"] = panel["asof"].gt(SEED)
    # 供应商流通市值占比不是剔除战略持股后的自由流通比例。
    panel["tradable_cap_ratio"] = panel.float_market_cap / panel.market_cap.where(panel.market_cap > 0)
    unique(panel, keys, "最终面板")
    for value, date in [("bm", "book_source_available"), ("roa", "roa_source_available"), ("rec_consensus", "rec_latest_publication")]:
        if (panel[value].notna() & (panel[date].isna() | panel[date].gt(panel["asof"]))).any():
            raise ValueError(f"{value} 违反信息可用日期约束")
    for value, date in EXTRA_AVAILABILITY.items():
        if (panel[value].notna() & (panel[date].isna() | panel[date].gt(panel["asof"]))).any():
            raise ValueError(f"{value} 违反信息可用日期约束")
    for date in ["capital_date", "factor_date", "status_available"]:
        if panel[date].gt(panel["asof"]).any():
            raise ValueError(f"{date} 错配未来记录")
    # 持仓期 h 的交易控制变量取 h-1 快照；使用精确日期连接，不能跨缺期 shift。
    lagcols = core + ["rec_consensus", "rec_brokers", "fin_report_period", "fin_announcement", "book_source_available", "roa_source_available", "controls_complete"]
    lagcols += EXTRA_METRICS + ["parent_income_ttm", "earnings_ttm_available", "balance_yoy_available", "revenue_yoy_available", "roe_source_available"]
    lag = panel[[*keys, *lagcols]].rename(columns={c: c + "_lag" for c in lagcols})
    lag["info_date"] = lag["asof"].dt.strftime("%Y%m%d").astype(int)
    lag["asof"] += pd.offsets.MonthEnd(6)
    holdingcols = [*keys, "date", "bonus_share_factor_6m", "bonus_event_count", "special_share_event_count", "share_adjustment_review", "code_change_review"]
    controls = panel[holdingcols].merge(lag, on=keys, how="left", validate="1:1")
    formal = panel.loc[panel.is_formal_sample].drop(columns="is_formal_sample")
    controls = controls.loc[controls["asof"].gt(SEED)].copy()
    controls["info_date"] = controls.info_date.astype("Int64")
    export(formal, "stock_halfyear")
    export(controls, "stock_holding_controls")
    formal.loc[formal.financial_extreme_review].to_csv(QA / "财务极端值待核验.csv", index=False, encoding="utf-8-sig", date_format="%Y-%m-%d")
    panel.loc[~panel.is_formal_sample].to_parquet(CACHE / "20091231初始快照.parquet", index=False)
    coverage = formal.groupby("date").agg(stocks=("stock_code", "size"), complete=("controls_complete", "sum"),
        returns=("ret6m", "count"), bm=("bm", "count"), roa=("roa", "count"), rec_change=("rec_change", "count"),
        share_event_review=("share_adjustment_review", "sum"), code_review=("code_change_review", "sum"))
    coverage.to_csv(QA / "分期覆盖率.csv", encoding="utf-8-sig")
    formal.groupby("date")[EXTRA_METRICS].count().to_csv(QA / "新增基本面分期覆盖.csv", encoding="utf-8-sig")
    extended_review = formal.pe_ttm.gt(1000) | formal.pb.gt(100) | formal.roe_ttm.abs().gt(1)
    extended_review |= formal.nonoperating_profit_share_ttm.abs().gt(1) | formal.rd_expense_to_revenue_ttm.gt(1)
    extended_review |= formal.tradable_cap_ratio.lt(0) | formal.tradable_cap_ratio.gt(1)
    extended_review |= formal.leverage.lt(0) | formal.leverage.gt(1) | formal.ep_ttm.abs().gt(1)
    extended_review |= formal.asset_growth_yoy.abs().gt(10) | formal.revenue_yoy.abs().gt(10)
    formal.loc[extended_review, ["stock_code", "date", *EXTRA_METRICS, "parent_income_ttm", "fin_report_period"]].to_csv(
        QA / "新增估值与基本面极端值.csv", index=False, encoding="utf-8-sig", date_format="%Y-%m-%d")
    AUDIT.update({"rows": len(formal), "stocks": int(formal.stock_code.nunique()), "periods": int(formal.date.nunique()),
        "duplicate_keys": int(formal.duplicated(["stock_code", "date"]).sum()),
        "complete_contemporaneous_controls": int(formal.controls_complete.sum()),
        "complete_lagged_controls": int(controls.controls_complete_lag.fillna(False).sum()),
        "financial_extreme_review_rows": int(formal.financial_extreme_review.sum()),
        "extra_metric_nonmissing": {c: int(formal[c].notna().sum()) for c in EXTRA_METRICS},
        "extended_review_rows": int(extended_review.sum()),
        "source_manifest": SOURCES})
    (QA / "构建核验.json").write_text(json.dumps(AUDIT, ensure_ascii=False, indent=2))
    print(json.dumps({k: v for k, v in AUDIT.items() if k != "source_manifest"}, ensure_ascii=False, indent=2), flush=True)


def main():
    for directory in [CACHE, OUT, QA]:
        directory.mkdir(parents=True, exist_ok=True)
    panel = market_panel()
    print(f"行情骨架 {len(panel):,} 行", flush=True)
    panel = financial_panel(panel)
    print("财报时点及 BM / ROA 完成", flush=True)
    panel = rating_panel(panel)
    print("评级共识及变动完成", flush=True)
    panel = distribution_panel(panel)
    panel = capital_and_status(panel)
    finish(panel)


if __name__ == "__main__":
    main()
