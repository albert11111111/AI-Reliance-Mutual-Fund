from pathlib import Path
import re
import pandas as pd

BASE = Path(__file__).resolve().parents[1]
OUT = Path(__file__).resolve().parent
source = pd.read_excel(BASE / '基金样本名单.xlsx', dtype=str).fillna('')
source.insert(0, '名单行号', range(2, len(source) + 2))
shares = pd.read_parquet(BASE / '面板数据整合/基金明细/基金份额.parquet').fillna('')
masters = pd.read_parquet(BASE / '面板数据整合/基金明细/基金主体.parquet').fillna('')

def normalize(value):
    return re.sub(r'_\d+$', '', str(value).strip().split('.')[0])

source['匹配用份额代码'] = source['证券代码'].map(normalize)
source['匹配用主体代码'] = source['基金主代码(官方)'].map(normalize)
source['名单历史版本标记'] = source['证券代码'].str.contains(r'_\d+(?:\.|$)', regex=True).map({True:'是', False:'否'})
assert len(source) == 7683 and source['证券代码'].is_unique

share_fields = ['Symbol', 'ShortName', 'FundID', 'MasterFundCode', 'FundClassID', 'Relation', 'InceptionDate', 'MaturityDate', 'TransitionStartDate', 'AfterTransitionSymbol', 'TransExplanation', '_source_member', '_source_row']
candidates = source.merge(shares[share_fields], how='left', left_on='匹配用份额代码', right_on='Symbol', validate='many_to_many')
master_id = masters.groupby('MasterFundCode')['FundID'].agg(lambda x: sorted(set(v for v in x if v))).to_dict()
choices = []
extra = []
for _, row in source.iterrows():
    sub = shares[shares.Symbol == row['匹配用份额代码']]
    method = '份额代码'
    if sub.empty:
        sub = shares[(shares.FrontendCode == row['匹配用份额代码']) | (shares.BackendCode == row['匹配用份额代码'])]
        method = '前端或后端代码'
    if sub.empty:
        names = {str(row[x]).strip() for x in ['证券名称', '基金简称(官方)']} - {''}
        sub = shares[shares.ShortName.isin(names) | shares.FullName.isin(names) | shares.TradingShortName.isin(names) | shares.ExpandShortName.isin(names)]
        method = '名称完全一致'
    if sub.empty:
        method = '未匹配'
    elif method != '份额代码':
        for _, c in sub.iterrows():
            extra.append({'名单行号':row['名单行号'], '证券代码':row['证券代码'], '证券名称':row['证券名称'], '基金简称(官方)':row['基金简称(官方)'], '候选依据':method, **{k:c[k] for k in share_fields}, 'FrontendCode':c.FrontendCode, 'BackendCode':c.BackendCode})
    ids = sorted(set(v for v in sub.FundID if v))
    mids = master_id.get(row['匹配用主体代码'], [])
    common = sorted(set(ids) & set(mids))
    chosen = common[0] if len(common) == 1 else ids[0] if len(ids) == 1 else ''
    picked = sub[sub.FundID == chosen]
    classids = sorted(set(v for v in picked.FundClassID if v))
    classid = classids[0] if len(classids) == 1 else ''
    if not ids:
        status = '份额代码未匹配'
    elif not chosen or not classid:
        status = '多候选待核查'
    elif mids and chosen not in mids:
        status = '主体代码冲突待核查'
    elif method == '名称完全一致':
        status = '名称一致待核查'
    elif method == '前端或后端代码':
        status = '前后端代码匹配待核查'
    elif row['名单历史版本标记'] == '是':
        status = '历史版本待核查'
    else:
        status = '份额代码唯一匹配' if len(ids) == 1 else '主体代码辅助匹配待核查'
    names = {str(row[x]).strip() for x in ['证券名称', '基金简称(官方)']} - {''}
    cname = set(picked.ShortName) | set(picked.FullName) | set(picked.TradingShortName) | set(picked.ExpandShortName)
    choices.append({'CSMAR基金ID': chosen, 'CSMAR份额类别ID': classid,
                    'CSMAR份额代码': ';'.join(sorted(set(v for v in picked.Symbol if v))), '匹配依据':method,
                    '名称核验': '名称完全一致' if names & cname else '名称不同或未找到，可能改名，待核查',
                    'CSMAR候选基金ID': ';'.join(ids), '名单主体代码候选基金ID': ';'.join(mids),
                    '主体代码核验': '未找到主体代码，仅据份额代码' if not mids else ('与份额代码一致' if chosen in mids else '需核查'),
                    'CSMAR份额名称': ';'.join(sorted(set(v for v in picked.ShortName if v))),
                    'CSMAR主体代码': ';'.join(sorted(set(v for v in picked.MasterFundCode if v))),
                    'CSMAR份额成立日': ';'.join(sorted(set(v for v in picked.InceptionDate if v))),
                    'CSMAR份额终止日': ';'.join(sorted(set(v for v in picked.MaturityDate if v))),
                    '匹配状态': status})
mapping = pd.concat([source.reset_index(drop=True), pd.DataFrame(choices)], axis=1)
assert len(mapping) == len(source) and mapping['名单行号'].is_unique
mapping.to_csv(OUT / '01_名单逐条匹配结果.csv', index=False, encoding='utf-8-sig')
candidates.to_csv(OUT / '02_名单份额全部CSMAR候选.csv', index=False, encoding='utf-8-sig')
pd.DataFrame(extra).to_csv(OUT / '06_名称和前后端代码补充候选.csv', index=False, encoding='utf-8-sig')

matched = mapping[mapping['CSMAR基金ID'] != '']
share_ids = set(matched['CSMAR基金ID'])
master_ids = {v for values in mapping['名单主体代码候选基金ID'] for v in values.split(';') if v}
ids = share_ids | master_ids
allshares = shares[shares.FundID.isin(ids)].copy()
listed = {(row['CSMAR基金ID'], symbol, row['CSMAR份额类别ID']) for _, row in matched.iterrows() for symbol in row['CSMAR份额代码'].split(';') if symbol}
allshares.insert(0, '名单中出现该份额', [('是' if (a,b,c) in listed else '否') for a,b,c in zip(allshares.FundID, allshares.Symbol, allshares.FundClassID)])
allshares.insert(1, '主体候选来源', ['份额、前后端代码或名称匹配' if v in share_ids else '仅名单主体代码匹配待核查' for v in allshares.FundID])
allshares.to_csv(OUT / '03_匹配主体的全部CSMAR份额.csv', index=False, encoding='utf-8-sig')
pd.DataFrame([{'CSMAR基金ID':v, '候选来源':'份额、前后端代码或名称匹配' if v in share_ids else '仅名单主体代码匹配待核查'} for v in sorted(ids)]).to_csv(OUT / '04_名单对应基金主体ID候选.csv', index=False, encoding='utf-8-sig')
mapping[(mapping['匹配状态'] != '份额代码唯一匹配') | (mapping['主体代码核验'] != '与份额代码一致') | (mapping['名称核验'] != '名称完全一致')].to_csv(OUT / '05_需要核查的名单记录.csv', index=False, encoding='utf-8-sig')
assert set(allshares.FundID).issubset(ids) and len(matched) + len(mapping[mapping['CSMAR基金ID'] == '']) == len(source)
print('名单',len(mapping),'主体候选',len(ids),'其中仅主体代码候选',len(ids - share_ids),'全部份额行',len(allshares))
print(mapping['匹配状态'].value_counts().to_string())
print('名单对应主体的所有份额中，名单未列出的行数',sum(allshares['名单中出现该份额']=='否'))
