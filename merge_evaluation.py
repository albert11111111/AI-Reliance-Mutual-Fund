"""把 CSMAR「基金日净值近半年评价文件」（FUND_EVL_DAYAHAY）的全部指标接到 Controls 面板末尾。

源文件：每个「基金日净值近半年评价文件*」文件夹一个 FUND_EVL_DAYAHAY.xlsx，按年分批下载，
每行是 Symbol × TradingDate 的近半年滚动评价指标（日频）。
同一批次 123209348 有三份完全相同的副本，按文件内容去重只读一次。

日期匹配：面板 date 为半年末（0630 / 1231），取该半年最后一个交易日
（源数据该半年内出现的最大工作日 TradingDate）的评价指标，记在 EvalDate 列。
代码匹配：面板 fundid = 评价文件 Symbol（份额代码），并核对其 FundClassID 与月度净值中
该代码的 FundClassID 一致（月度净值的 Symbol 已核实为面板基金的主代码份额）。

输入：基金名单_fundid_date_TotalEquity_Holder_BuySell_Fee_Manager_Alpha_Controls.xlsx、
      基金日净值近半年评价文件*/FUND_EVL_DAYAHAY.xlsx、月度净值.xlsx（均不改）。
输出：基金名单_fundid_date_TotalEquity_Holder_BuySell_Fee_Manager_Alpha_Controls_Eval.xlsx。
"""

import hashlib
from concurrent.futures import ProcessPoolExecutor
from pathlib import Path

import pandas as pd
from openpyxl import load_workbook

from build_controls import CONTROLS
from excel_style import style_sheet
from merge_alpha import PANEL_FORMATS
from merge_fund_manager import TEXT_COLS

ROOT = Path(__file__).resolve().parent
PANEL = ROOT / "基金名单_fundid_date_TotalEquity_Holder_BuySell_Fee_Manager_Alpha_Controls.xlsx"
NAV = ROOT / "月度净值.xlsx"
OUT = ROOT / "基金名单_fundid_date_TotalEquity_Holder_BuySell_Fee_Manager_Alpha_Controls_Eval.xlsx"
SRC_GLOB = "基金日净值近半年评价文件*/FUND_EVL_DAYAHAY.xlsx"

ID_COLS = ["TradingDate", "Symbol", "FullName", "FundClassID"]
EVAL_TEXT = ["EvalDate", "CategoryID", "MktPortfolioID", "MktPortfolio"]
HEADER_ROWS = 2  # 字段名之后还有「中文名」「单位」两行


def file_md5(path: Path) -> str:
    h = hashlib.md5()
    with open(path, "rb") as f:
        for chunk in iter(lambda: f.read(1 << 20), b""):
            h.update(chunk)
    return h.hexdigest()


def read_half_end_rows(path: Path) -> pd.DataFrame:
    """流式读取一个评价文件，只保留 6 月、12 月下旬的行（半年最后一个交易日必在其中）。"""
    ws = load_workbook(path, read_only=True).active
    ws.reset_dimensions()  # 文件自带的表格范围不准，须重置后整行读取
    rows = ws.iter_rows(values_only=True)
    header = next(rows)
    for _ in range(HEADER_ROWS):
        next(rows)
    keep = []
    for row in rows:
        d = str(row[0])[:10]
        if d[5:7] in ("06", "12") and d[8:10] >= "20":
            keep.append(row)
    out = pd.DataFrame(keep, columns=header)
    out["TradingDate"] = out["TradingDate"].astype(str).str[:10]
    return out


def load_evaluation() -> pd.DataFrame:
    """读全部评价文件，取每个半年最后一个交易日的行。"""
    paths, seen = [], set()
    for p in sorted(ROOT.glob(SRC_GLOB)):
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
    # ponytail: 用工作日近似交易日；2009–2026 年各半年末无工作日休市，若扩展样本需改用交易所日历
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


def check_share_class(ev: pd.DataFrame) -> None:
    """评价文件的 (Symbol, FundClassID) 必须与月度净值一致。"""
    nav = pd.read_excel(NAV, dtype=str, usecols=["Symbol", "FundClassID"]).iloc[HEADER_ROWS:]
    nav = nav.apply(lambda s: s.str.strip()).drop_duplicates()
    m = ev[["Symbol", "FundClassID"]].drop_duplicates().merge(nav, on="Symbol", how="inner", suffixes=("", "_nav"))
    bad = m.loc[m["FundClassID"].ne(m["FundClassID_nav"])]
    if not bad.empty:
        raise SystemExit(f"评价文件与月度净值的 FundClassID 不一致：\n{bad.to_string(index=False)}")
    print(f"份额核对通过：{len(m)} 个 Symbol 的 FundClassID 与月度净值一致")


def build() -> None:
    panel = pd.read_excel(PANEL, dtype={**{c: str for c in TEXT_COLS}, "BuySellStartdate": str})
    ev = load_evaluation()
    check_share_class(ev)
    print(
        f"评价文件半年末行={len(ev)} 基金={ev['Symbol'].nunique()} "
        f"半年末={ev['date'].min()}..{ev['date'].max()}（{ev['date'].nunique()} 个）"
    )
    print("各半年末对应的最后交易日：", ev.groupby("date")["EvalDate"].first().to_dict())

    eval_cols = [c for c in ev.columns if c not in (*ID_COLS, "Symbol", "date")]
    out = panel.merge(
        ev[["Symbol", "date", *eval_cols]],
        left_on=["fundid", "date"],
        right_on=["Symbol", "date"],
        how="left",
        validate="one_to_one",
    ).drop(columns="Symbol")
    if len(out) != len(panel):
        raise SystemExit("合并后行数变化")

    # 未匹配原因：评价文件起点之前 / 整只基金不在评价文件 / 早于该基金首次评价（新基金不足半年）/ 其他
    miss = out["EvalDate"].isna()
    first, last = ev.groupby("Symbol")["date"].min(), ev.groupby("Symbol")["date"].max()
    f, l = out["fundid"].map(first), out["fundid"].map(last)
    absent = f.isna()
    f, l = f.fillna("99999999"), l.fillna("00000000")  # 缺失基金已归入「整只基金无」，填哨兵值避免误判
    reasons = {
        "评价文件起点前": out["date"].lt(ev["date"].min()),
        "整只基金无": absent,
        "早于首次评价": out["date"].lt(f),
        "晚于末次评价": out["date"].gt(l),
    }
    rest = miss.copy()
    for k, v in reasons.items():
        reasons[k] = v & rest
        rest &= ~v
    why = " ".join(f"{k}={int(v.sum())}" for k, v in reasons.items())
    print(f"面板行={len(out)} 匹配={int((~miss).sum())} 未匹配={int(miss.sum())}（{why} 中间缺失={int(rest.sum())}）")
    if rest.any():
        print(out.loc[rest, ["fundid", "date"]].to_string(index=False))

    num_cols = [c for c in eval_cols if c not in EVAL_TEXT]
    with pd.ExcelWriter(OUT, engine="openpyxl") as writer:
        out.to_excel(writer, index=False, sheet_name="fundid_date")
        style_sheet(
            writer.sheets["fundid_date"],
            out.columns,
            text_cols=(*TEXT_COLS, *EVAL_TEXT),
            number_formats={**PANEL_FORMATS, **{c: "0.000000" for c in CONTROLS}, **{c: "0.0000" for c in num_cols}},
            widths={"EvalDate": 12, "MktPortfolio": 28},
        )
    print(f"写出 {OUT.name}")


if __name__ == "__main__":
    build()
