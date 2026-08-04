#!/usr/bin/env python3
"""
料号翻译工具 - 将周报中的新料号翻译为"旧料号-新料号"格式
通过直接修改 XLSX 内部 XML 实现，完美保留原始格式、图表、图片等。
"""

import re
import sys
import os
import shutil
import tempfile
import zipfile
from pathlib import Path
import xml.etree.ElementTree as ET

import pandas as pd


def load_mapping(mapping_path: str) -> dict:
    """从 mapping.xlsx 构建 new→old 反向映射。"""
    df = pd.read_excel(mapping_path, header=None, skiprows=2)
    df.columns = range(df.shape[1])

    mapping = df[[9, 10]].dropna()
    mapping = mapping[mapping[9].astype(str).str.strip() != ""]
    mapping = mapping[mapping[10].astype(str).str.strip() != ""]

    new_to_old = {}
    for _, row in mapping.iterrows():
        old_val = str(row[9]).strip()
        new_val = str(row[10]).strip()
        if new_val.lower() != "nan" and old_val.lower() != "nan":
            new_to_old[new_val] = old_val

    return new_to_old


def is_part_number(value: str) -> bool:
    """判断是否为料号格式（A/B 开头 + 字母数字）。"""
    if not value:
        return False
    return bool(re.match(r"^[AB][A-Z0-9]{3,}(\.[ABT])?$", value.strip()))


def translate_xlsx_via_xml(report_path: str, mapping_path: str, output_path: str) -> dict:
    """
    通过直接修改 XLSX 内部 XML 来翻译料号。
    
    原理：
    1. XLSX 本质是一个 ZIP 文件
    2. 料号以共享字符串的形式存储在 xl/sharedStrings.xml 中
    3. 我们直接修改 sharedStrings.xml 中的料号文本
    4. 重新打包 ZIP，完美保留所有原始特性
    """
    new_to_old = load_mapping(mapping_path)
    print(f"  Loaded {len(new_to_old)} mapping pairs")

    # XLSX 命名空间
    NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
    ET.register_namespace('', NS)
    ET.register_namespace('r', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships')

    tmpdir = Path(tempfile.mkdtemp(prefix='xlsx_translate_'))

    try:
        # 1. 解压原始 XLSX
        with zipfile.ZipFile(report_path, 'r') as z:
            z.extractall(tmpdir)

        # 2. 修改 sharedStrings.xml
        ss_path = tmpdir / 'xl' / 'sharedStrings.xml'
        if not ss_path.exists():
            print("  Warning: no sharedStrings.xml found")
            shutil.copy2(report_path, output_path)
            return {}

        tree = ET.parse(ss_path)
        root = tree.getroot()

        # 收集所有 <si><t> 元素
        stats = {"translated": 0, "unchanged": 0, "not_found": 0}
        si_elements = root.findall(f'{{{NS}}}si')

        for si in si_elements:
            t_elem = si.find(f'{{{NS}}}t')
            if t_elem is None or t_elem.text is None:
                continue

            original = t_elem.text.strip()
            if not is_part_number(original):
                continue

            if original in new_to_old:
                old_val = new_to_old[original]
                if old_val != original:
                    # 同时保留空格属性（xml:space="preserve"）
                    t_elem.text = f"{old_val}-{original}"
                    t_elem.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
                    stats["translated"] += 1
                else:
                    stats["unchanged"] += 1
            else:
                stats["not_found"] += 1

        # 写入修改后的 sharedStrings.xml
        tree.write(ss_path, xml_declaration=True, encoding='UTF-8')

        print(f"  Translated: {stats['translated']}, "
              f"Unchanged: {stats['unchanged']}, "
              f"Not found: {stats['not_found']}")

        # 3. 重新打包为 XLSX
        with zipfile.ZipFile(output_path, 'w', zipfile.ZIP_DEFLATED) as zout:
            for fpath in sorted(tmpdir.rglob('*')):
                if fpath.is_file():
                    arcname = str(fpath.relative_to(tmpdir))
                    zout.write(fpath, arcname)

        print(f"  Saved to: {output_path}")
        return stats

    finally:
        shutil.rmtree(tmpdir, ignore_errors=True)


def translate_report(
    report_path: str,
    mapping_path: str,
    output_path: str,
    sheets_to_translate: list = None,
) -> dict:
    """主翻译入口。"""
    print(f"Processing: {os.path.basename(report_path)}")
    stats = translate_xlsx_via_xml(report_path, mapping_path, output_path)
    # 返回兼容格式，让 app.py 能正常统计
    return {
        "All Sheets": {
            "translated": stats.get("translated", 0),
            "unchanged": stats.get("unchanged", 0),
            "not_found": stats.get("not_found", 0),
        }
    }


# ── CLI ────────────────────────────────────────────────────────────
if __name__ == "__main__":
    if len(sys.argv) < 3:
        print("Usage: python translate.py <weekly_report.xlsx> <mapping.xlsx> [output.xlsx]")
        sys.exit(1)

    report = sys.argv[1]
    mapping = sys.argv[2]
    output = sys.argv[3] if len(sys.argv) > 3 else report.replace(".xlsx", "_translated.xlsx")

    for p, label in [(report, "Report"), (mapping, "Mapping")]:
        if not os.path.exists(p):
            print(f"Error: {label} file not found: {p}")
            sys.exit(1)

    translate_report(report, mapping, output)
    print("Done!")
