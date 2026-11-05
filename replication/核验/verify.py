"""独立验收：逐单元格严格比较新旧数据，检查回归算法未改及原始复制完整。"""
from pathlib import Path
import ast
import hashlib
import importlib.metadata
import json
import sys
import pandas as pd

ROOT = Path(__file__).resolve().parents[1]
OLD = ROOT.parent
WORK = ROOT / '中间数据'
REPORT = ROOT / '核验' / 'verification.json'
FINAL = '基金名单_fundid_date_TotalEquity_Holder_BuySell_Fee_Manager_Alpha_Controls_Eval.xlsx'

def compare(old, new):
    # 全列按 object 读取，保留文本代码、数值和缺失；容差为零。
    a = pd.read_excel(old, dtype=object)
    b = pd.read_excel(new, dtype=object)
    pd.testing.assert_frame_equal(a, b, check_exact=True, check_dtype=True)
    print(f'一致：{new.name} {a.shape}', flush=True)
    return {'file': new.name, 'rows': len(a), 'columns': len(a.columns), 'exact': True}


def main():
    items = []
    inputs = ['基金名单_fundid_date.xlsx','无风险利率_月度.xlsx','月度净值_rf.xlsx']
    inputs += [f'月度净值_rf_{m}.xlsx' for m in ['三因子','四因子','五因子','CH3','CH4']]
    inputs += [f'滚动alpha_{m}.xlsx' for m in ['CAPM','三因子','四因子','五因子','CH3','CH4']]
    stage = sys.argv[1] if len(sys.argv) > 1 else 'all'
    if stage not in {'all', 'upstream', 'final'}:
        raise ValueError('参数仅支持 all / upstream / final')
    if stage == 'final':
        inputs = []
    for name in inputs:
        items.append(compare(OLD / name, WORK / name))
    if stage != 'upstream':
        items.append(compare(OLD / FINAL, ROOT / '结果' / FINAL))
    # 核心回归及样本函数 AST 完全不变；仅调整输入输出目录。
    def functions(path):
        return {n.name: ast.dump(n, include_attributes=False) for n in ast.parse(path.read_text()).body
                if isinstance(n, ast.FunctionDef)}
    if functions(OLD/'rolling_alpha.py') != functions(ROOT/'rolling_alpha.py'):
        raise ValueError('rolling_alpha 函数发生变化')
    files=json.loads((ROOT/'核验/原始数据清单.json').read_text())
    for row in files:
        for p in [OLD/row['path'], ROOT/'原始数据'/row['path']]:
            with p.open('rb') as f:
                if hashlib.file_digest(f,'sha256').hexdigest() != row['sha256']:
                    raise ValueError(f'原始文件发生变化：{p}')
    result={'status':'PASS','comparison':'exact; no numeric tolerance; row and column order included',
            'files':items,'raw_files_sha256_verified':len(files),'rolling_functions_unchanged':True,
            'python':sys.version,'packages':{n:importlib.metadata.version(n) for n in ['pandas','numpy','openpyxl','statsmodels']}}
    report = REPORT if stage == 'all' else REPORT.with_name(f'verification_{stage}.json')
    report.write_text(json.dumps(result,ensure_ascii=False,indent=2)+'\n')
    print('全部核验通过',flush=True)

if __name__ == '__main__':
    main()
