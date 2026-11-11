#!/usr/bin/env python3
"""只读核验原始压缩包，结果写至标准输出。"""

import csv
import hashlib
import io
import json
import re
import sys
from datetime import date
from pathlib import Path
from zipfile import ZipFile
from xml.etree import ElementTree


ROOT = Path(__file__).resolve().parent
EXPECTED = {
    "公司文件111337729": 6057,
    "股本变动文件114008558": 144451,
    "分配文件114014091": 72494,
    "月个股回报率文件114059067": 959548,
    "股票价格复权因子表(日)114928367": 62652,
    "上市公司上市状态变更表102253551": 4452,
    "上市公司基本信息变更表115010098": 179546,
    "上市公司基本信息年度表115029827": 74354,
    "资产负债表114132694": 709197,
    "利润表114149331": 710893,
    "年、中、季报基本情况文件114653443": 314660,
    "分析师评级及变动114629474": 452,
    "分析师评级及变动114638471": 239495,
    "分析师评级及变动114655331": 969770,
    "分析师评级及变动114701530": 1074541,
}
DATE_FIELDS = {
    "TRD_Co": ["Listdt", "Statdt"],
    "TRD_Capchg": ["Shrchgdt"],
    "TRD_Cptl": ["Exdistdt", "Annodt"],
    "TRD_Mnth": ["Trdmnt"],
    "TRD_AdjustFactor": ["TradingDate"],
    "STK_ITEMCHANGE": ["DeclareDate", "ChangeDate"],
    "STK_LISTEDCOINFOCHG": ["AnnouncementDate", "ImplementDate"],
    "STK_LISTEDCOINFOANL": ["EndDate"],
    "FS_Combas": ["Accper", "DeclareDate"],
    "FS_Comins": ["Accper", "DeclareDate"],
    "IAR_Rept": ["Accper", "Annodt"],
    "AF_Bench": ["Rptdt", "DeclareDate"],
}
XLSX_EXPECTED = {
    "上市公司上市状态变更表091453528": (4452, "STK_ITEMCHANGE"),
    "上市公司基本信息年度表100107917": (74354, "STK_LISTEDCOINFOANL"),
    "上市公司基本信息变更表102817958": (179546, "STK_LISTEDCOINFOCHG"),
    "分析师评级及变动115810968": (969770, "AF_Bench"),
}

csv.field_size_limit(10_000_000)
results = []
headers = {}
for stem, expected in EXPECTED.items():
    matches = list(ROOT.rglob(stem + "*.zip"))
    if len(matches) != 1:
        raise RuntimeError(f"未找到唯一文件：{stem}，找到 {len(matches)} 个")
    path = matches[0]
    with path.open("rb") as source:
        sha256 = hashlib.file_digest(source, "sha256").hexdigest()
    result = {"文件": str(path.relative_to(ROOT)), "SHA256": sha256,
              "页面记录数": expected, "CSV记录数": 0, "列数不符行数": 0, "分片": [], "日期": {}}
    with ZipFile(path) as archive:
        parts = [item for item in archive.infolist() if item.filename.lower().endswith(".csv")]
        if not parts:
            raise RuntimeError(f"压缩包中没有 CSV：{path.name}")
        for part in parts:
            table = Path(part.filename).stem.rstrip("0123456789")
            with archive.open(part) as data:
                reader = csv.reader(io.TextIOWrapper(data, encoding="utf-8-sig", newline=""))
                header = next(reader)
                if table in headers and headers[table] != header:
                    raise RuntimeError(f"同表分片的字段不一致：{part.filename}")
                headers[table] = header
                fields = DATE_FIELDS[table]
                positions = {field: header.index(field) for field in fields}
                for field in fields:
                    result["日期"].setdefault(field, {"最早": None, "最晚": None, "缺失": 0, "多日期记录": 0, "无效": 0, "无效示例": []})
                count = 0
                for row in reader:
                    count += 1
                    if len(row) != len(header):
                        result["列数不符行数"] += 1
                        continue
                    for field, index in positions.items():
                        value = row[index].strip()
                        stat = result["日期"][field]
                        if not value:
                            stat["缺失"] += 1
                            continue
                        values = value.split(",") if table in {"FS_Combas", "FS_Comins"} and field == "DeclareDate" else [value]
                        stat["多日期记录"] += len(values) > 1
                        for item in values:
                            item = item.strip()
                            try:
                                date.fromisoformat(item + "-01" if len(item) == 7 else item)
                            except ValueError:
                                stat["无效"] += 1
                                if item not in stat["无效示例"] and len(stat["无效示例"]) < 3:
                                    stat["无效示例"].append(item)
                                continue
                            stat["最早"] = min(stat["最早"] or item, item)
                            stat["最晚"] = max(stat["最晚"] or item, item)
                result["CSV记录数"] += count
                result["分片"].append({"文件": part.filename, "记录数": count, "字段": header})
    result["记录数一致"] = result["CSV记录数"] == expected
    results.append(result)

for stem, (expected, table) in XLSX_EXPECTED.items():
    matches = list(ROOT.rglob(stem + "*.zip"))
    if len(matches) != 1:
        raise RuntimeError(f"未找到唯一文件：{stem}，找到 {len(matches)} 个")
    path = matches[0]
    with path.open("rb") as source:
        sha256 = hashlib.file_digest(source, "sha256").hexdigest()
    result = {"文件": str(path.relative_to(ROOT)), "SHA256": sha256,
              "页面记录数": expected, "Excel记录数": 0, "列数不符行数": None, "工作表": []}
    with ZipFile(path) as archive:
        workbooks = [item for item in archive.infolist() if item.filename.lower().endswith(".xlsx")]
        if not workbooks:
            raise RuntimeError(f"压缩包中没有 Excel 文件：{path.name}")
        for workbook in workbooks:
            with ZipFile(io.BytesIO(archive.read(workbook))) as xlsx:
                sheets = [name for name in xlsx.namelist() if re.fullmatch(r"xl/worksheets/sheet\d+\.xml", name)]
                if not sheets:
                    raise RuntimeError(f"Excel 文件没有工作表：{workbook.filename}")
                for sheet in sheets:
                    with xlsx.open(sheet) as data:
                        first = data.read(1_000_000)
                        dimension = re.search(rb'<dimension ref="([^"]+)"', first)
                        header = re.search(rb'<row\b[^>]*>.*?</row>', first, re.DOTALL)
                        if not dimension or not header:
                            raise RuntimeError(f"无法读取工作表范围或字段：{sheet}")
                        columns = ["".join(cell.itertext()) for cell in ElementTree.fromstring(header.group())]
                        if columns != headers[table]:
                            raise RuntimeError(f"Excel 与 CSV 字段不一致：{workbook.filename}")
                        # 单元格文本中的尖括号会被 XML 转义，故可按行标签计数。
                        count = first.count(b"<row ")
                        tail = first[-4:]
                        while chunk := data.read(1_000_000):
                            block = tail + chunk
                            count += block.count(b"<row ")
                            tail = block[-4:]
                    result["Excel记录数"] += count - 3
                    result["工作表"].append({"文件": workbook.filename, "工作表": sheet,
                                       "标记范围": dimension.group(1).decode(), "记录数": count - 3, "字段": columns})
    result["记录数一致"] = result["Excel记录数"] == expected
    results.append(result)

print(json.dumps(results, ensure_ascii=False, indent=2))
if any(not result["记录数一致"] or result["列数不符行数"] for result in results):
    sys.exit(1)
