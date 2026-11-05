"""从旧版提取已核验的计算函数，生成独立复现目录；仅用于本次迁移留痕。"""
import ast
from pathlib import Path
import shutil
import hashlib
import json

OLD = Path(__file__).resolve().parents[2]
NEW = OLD / 'replication'

def extract(file, names):
    src = (OLD / file).read_text()
    tree = ast.parse(src)
    return '\n\n\n'.join(ast.get_source_segment(src, n) for n in tree.body
        if (isinstance(n, ast.FunctionDef) and n.name in names)
        or (isinstance(n, ast.Assign) and any(isinstance(t, ast.Name) and t.id in names for t in n.targets)))

def write(name, text):
    (NEW / name).write_text(text.strip() + '\n')

# 原始文件复制，保留旧工作区；不用链接，目录可独立迁移。
paths = [OLD / n for n in ['基金名单.xlsx','基金size.xlsx','机构持有占比.xlsx','换手率计算.xlsx','费用.xlsx','Fund_FundManager.xlsx','无风险利率.xlsx','月度净值.xlsx']]
for pat in ['三因子模型指标*/STK*.xlsx','Carhart四因子*/STK*.xlsx','五因子模型指标*/STK*.xlsx','CH因子_鸣石_20261103/*.xlsx','基金日净值近半年评价文件*/FUND_EVL_DAYAHAY.xlsx','CSMAR原始数据/基金份额信息/*.zip']:
    paths.extend(sorted(OLD.glob(pat)))
manifest=[]
for p in paths:
    dst=NEW/'原始数据'/p.relative_to(OLD)
    dst.parent.mkdir(parents=True,exist_ok=True)
    if not dst.exists(): shutil.copy2(p,dst)
    digest=lambda f: hashlib.file_digest(f.open('rb'),'sha256').hexdigest()
    a,b=digest(p),digest(dst)
    if a!=b: raise ValueError(f'复制不一致: {p}')
    manifest.append({'path':str(p.relative_to(OLD)),'bytes':p.stat().st_size,'sha256':a})
write('核验/原始数据清单.json',json.dumps(manifest,ensure_ascii=False,indent=2))
shutil.copy2(OLD/'excel_style.py',NEW/'excel_style.py')
rolling=(OLD/'rolling_alpha.py').read_text().replace('ROOT = Path(__file__).resolve().parent','ROOT = Path(__file__).resolve().parent / "中间数据"')
write('rolling_alpha.py',rolling)

head='''"""第一步：生成半年基础面板、月度无风险利率及五组回归输入。"""
from datetime import date, datetime
from pathlib import Path
import pandas as pd

ROOT = Path(__file__).resolve().parent
RAW = ROOT / "原始数据"
WORK = ROOT / "中间数据"
RF_SRC = RAW / "无风险利率.xlsx"
NAV_SRC = RAW / "月度净值.xlsx"
MARKET = "P9714"
PORTFOLIO_5 = "1"
'''
text=head+'\n# %% 1. 基金名单与半年基础面板\n'+extract('build_fundid_date.py',['SAMPLE_END','KEEP','parse_day','semi_dates'])
text+='''


def build_panel():
    src = pd.read_excel(RAW / "基金名单.xlsx", dtype=str)
    src = src.sort_values(["fundid", "成立时间_earliest", "成立时间_latest"],
                          ascending=[True, True, False], kind="mergesort").drop_duplicates("fundid")
    rows = []
    for _, rec in src.iterrows():
        for end in semi_dates(parse_day(rec["成立时间_earliest"]), parse_day(rec["到期日"])):
            rows.append({"fundid": rec["fundid"], "date": end,
                         **{c: rec[c] for c in KEEP if c != "fundid"}})
    panel = pd.DataFrame(rows).sort_values(["fundid", "date"], kind="mergesort").reset_index(drop=True)
    for c in ("发生转型代码重复", "发生过转型"):
        panel[c] = panel[c].astype(int)
    if panel.duplicated(["fundid", "date"]).any():
        raise ValueError("半年面板键不唯一")
    panel.to_excel(WORK / "基金名单_fundid_date.xlsx", index=False, sheet_name="fundid_date")
    print(f"基础面板：{len(panel):,} 行")
'''
text+='\n# %% 2. 日历月无风险利率；保持旧版月末向后匹配口径\n'+extract('build_rf_monthly.py',['load_rate_changes','build_monthly_rf','load_nav'])
text+='\n\n# %% 3. CSMAR 与 CH 因子；CH3、CH4 使用各自的 SMB\n'
text+=extract('merge_ff_factors.py',['SPECS']).replace('SPECS =','FF_SPECS =').replace('"src": ROOT','"src": RAW').replace('"out": ROOT','"out": WORK')
text+='\n\nCH_DIR = RAW / "CH因子_鸣石_20261103"\n'+extract('merge_ch_factors.py',['SPECS']).replace('SPECS =','CH_SPECS =').replace('"out": ROOT','"out": WORK')
text+='\n\n'+extract('merge_ff_factors.py',['NAV_NUM','load_factors'])+'\n\n'+extract('merge_ch_factors.py',['load_ch'])
text+='''

# %% 4. 顺序执行；保留 Excel 落盘精度，以复现原有链路

def main():
    WORK.mkdir(exist_ok=True)
    build_panel()
    changes, nav = load_rate_changes(), load_nav()
    rf = build_monthly_rf(changes, max(pd.to_datetime(nav["TradingDate"]).max(), changes["RateChangeDate"].max()))
    rf.to_excel(WORK / "无风险利率_月度.xlsx", index=False, sheet_name="rf_monthly")
    nav = nav.merge(rf[["year_month", "RateChangeDate", "q_annual_pct", "rf_monthly"]],
                    on="year_month", how="left", validate="many_to_one")
    nav = nav.sort_values(["Symbol", "TradingDate"], kind="mergesort").reset_index(drop=True)
    nav.to_excel(WORK / "月度净值_rf.xlsx", index=False, sheet_name="nav_rf")
    nav = pd.read_excel(WORK / "月度净值_rf.xlsx", dtype=str)
    nav = nav.apply(lambda s: s.str.strip())
    for c in NAV_NUM:
        nav[c] = pd.to_numeric(nav[c], errors="coerce")
    for spec in [*FF_SPECS, *CH_SPECS]:
        fac = (load_factors(spec["src"], spec["factors"], spec["extra_filters"])
               if "extra_filters" in spec else load_ch(spec["src"], spec["factors"]))
        out = nav.merge(fac, on="year_month", how="left", validate="many_to_one")
        out.to_excel(spec["out"], index=False, sheet_name="nav_ff" if "extra_filters" in spec else "nav_ch")
        print(f"{spec['label']}：{len(out):,} 行，因子不全 {out[spec['factors']].isna().any(axis=1).sum():,}")


if __name__ == "__main__":
    main()
'''
write('build_inputs.py',text)

# 主合并文件：移植业务函数，去掉各脚本重复的读写、格式与调试断言。
text='''"""第三步：合并财报、经理、收益与评价指标，生成与旧版一致的半年面板。

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
'''
text+='\n'+extract('merge_fund_manager.py',['TEXT_COLS','KEEP_MGR','LIST_GAP_FILLS'])
text+='\n\n'+extract('merge_holder_ratio.py',['KEEP_FROM_HOLDER'])
text+='\n\n'+extract('merge_alpha.py',['ALPHAS','MONTHS_PER_HALF','NEW_COLS','PANEL_FORMATS'])
text+='\n\n'+extract('build_controls.py',['CONTROLS'])
text+='\n\n'+extract('merge_evaluation.py',['SRC_GLOB','ID_COLS','EVAL_TEXT','HEADER_ROWS'])
text+='''

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
'''
for title,file,names in [
('规模与持有人','merge_total_equity.py',['load_size']),
('', 'merge_holder_ratio.py',['nonempty','load_holder','choose_ratios']),
('半年买卖发生额：缺侧补零、起算不同或任一侧减负则整期空','merge_buy_sell_zero.py',['load_turnover','half_year_amounts']),
('半年费用：本期口径、年报减半年报','merge_fee.py',['load_fee','half_year_fee']),
('经理在职区间与团队聚合；保留旧版三个人工空窗补丁','merge_fund_manager.py',['load_managers','active_team_stats']),
('半年收益与六模型 alpha：六个月齐全才计算','merge_alpha.py',['half_date','half_year_raw_ret','load_alpha','half_year_alpha','check_share_class','merge_half']),
('控制变量：相邻半年净资产与本期发生额','build_controls.py',['prev_half_end','month_diff','add_controls']),
('CSMAR 评价：全批次去重、半年最后工作日','merge_evaluation.py',['file_md5','read_half_end_rows','load_evaluation','check_share_class'])]:
    fragment=extract(file,names)
    if file=='merge_alpha.py': fragment=fragment.replace('check_share_class','check_alpha_share')
    if file=='merge_evaluation.py':
        fragment=fragment.replace('check_share_class','check_eval_share').replace('ROOT.glob','RAW.glob')
        fragment=fragment.replace('ws = load_workbook(path, read_only=True).active','book = load_workbook(path, read_only=True)\n    ws = book.active')
        fragment=fragment.replace('header = next(rows)','header = next(rows)\n    date_index = header.index("TradingDate")').replace('str(row[0])','str(row[date_index])').replace('out = pd.DataFrame(keep, columns=header)','book.close()\n    out = pd.DataFrame(keep, columns=header)')
        fragment=fragment.replace('# ponytail:', '# shortcut:')
    text+='\n\n# %% '+title+'\n'+fragment if title else '\n\n'+fragment
text+='''

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
'''
write('merge_panel.py',text)
print('生成完毕，原始文件',len(paths))
