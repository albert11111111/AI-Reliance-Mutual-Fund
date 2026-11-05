"""汇总本次分开执行的上游与最终文件验收结果。"""
import json
from pathlib import Path

root = Path(__file__).resolve().parent
upstream = json.loads((root / 'verification_upstream.json').read_text())
final = json.loads((root / 'verification_final.json').read_text())
if upstream['status'] != 'PASS' or final['status'] != 'PASS':
    raise ValueError('存在未通过的核验')
result = {**upstream, 'files': upstream['files'] + final['files']}
(root / 'verification.json').write_text(json.dumps(result, ensure_ascii=False, indent=2) + '\n')
