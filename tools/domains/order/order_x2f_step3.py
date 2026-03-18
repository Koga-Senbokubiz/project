# -*- coding: utf-8 -*-
"""
order_x2f_step3.py

Step3: 変換元レイアウトXML作成
- Step1の正規化XML
- Step2の差異一覧xlsx
- 基本形テンプレートXML(TEMPLATE_FILE)
を入力にして、変換元レイアウトXMLを生成する。

想定:
- Step2のB列に「〇」が入っている行を採用対象とする
- Step2には顧客XML側のパス情報を持つ列がある
- 列名が多少変わっても拾えるようにヘッダ別名対応あり

出力XMLは「from layout」の中間表現として、後続工程(Step5/6)で使いやすい形を意識している。
"""

from __future__ import annotations

import argparse
import sys
from pathlib import Path
from typing import Dict, List, Optional, Tuple
import xml.etree.ElementTree as ET

from openpyxl import load_workbook


VERSION = "1.1.0"


DEFAULT_OUTPUT_ROOT_TAG = "from_layout"
DEFAULT_RECORD_TAG = "record"
DEFAULT_FIELD_TAG = "field"

SELECT_MARKS = {"〇", "○", "o", "O", "1", "Y", "y", "yes", "YES", "TRUE", "true"}

# Step2列名ゆらぎ吸収
HEADER_ALIASES = {
    "selected": [
        "採用",
        "採用マッピング",
        "採用可否",
        "採用フラグ",
        "use",
        "selected",
    ],
    "source_path": [
        "顧客XPath",
        "顧客XMLパス",
        "顧客パス",
        "顧客項目パス",
        "source_path",
        "xml_path",
        "path",
        "XPath",
    ],
    "source_name": [
        "顧客項目名",
        "顧客項目",
        "source_name",
        "xml_name",
        "項目名",
    ],
    "bms_path": [
        "BMSXPath",
        "BMSパス",
        "流通BMSパス",
        "target_path",
        "bms_path",
    ],
    "bms_name": [
        "BMS項目名",
        "流通BMS項目名",
        "target_name",
        "bms_name",
    ],
    "note": [
        "備考",
        "メモ",
        "note",
    ],
}


def log(msg: str) -> None:
    print(msg, flush=True)


def normalize_text(value) -> str:
    if value is None:
        return ""
    return str(value).strip()


def is_selected(value) -> bool:
    return normalize_text(value) in SELECT_MARKS


def strip_namespace(tag: str) -> str:
    if "}" in tag:
        return tag.split("}", 1)[1]
    return tag


def detect_header_map(header_row: List[str]) -> Dict[str, int]:
    """
    ヘッダ名から列番号(0始まり)を推定
    """
    header_map: Dict[str, int] = {}
    normalized_headers = [normalize_text(x) for x in header_row]

    for logical_name, aliases in HEADER_ALIASES.items():
        for idx, header in enumerate(normalized_headers):
            if header in aliases:
                header_map[logical_name] = idx
                break

    return header_map


def find_source_path_col_by_scan(rows: List[List[str]], header_map: Dict[str, int]) -> Optional[int]:
    """
    列名が不明でも、中身に XMLパスらしき値が多い列を推定する。
    """
    if "source_path" in header_map:
        return header_map["source_path"]

    if not rows:
        return None

    best_idx = None
    best_score = -1
    col_count = max(len(r) for r in rows)

    for col_idx in range(col_count):
        score = 0
        for row in rows[1:min(len(rows), 30)]:
            if col_idx >= len(row):
                continue
            val = normalize_text(row[col_idx])
            if "/" in val and len(val) >= 3:
                score += 1
        if score > best_score:
            best_score = score
            best_idx = col_idx

    if best_score <= 0:
        return None

    return best_idx


def load_step2_rows(step2_xlsx: Path) -> Tuple[List[Dict[str, str]], Dict[str, int]]:
    """
    Step2 xlsx 読み込み
    - selected列があればそれを採用判定列とする
    - なければ B列(1) を採用判定列とする
    """
    wb = load_workbook(step2_xlsx, data_only=True)
    ws = wb[wb.sheetnames[0]]

    raw_rows: List[List[str]] = []
    for row in ws.iter_rows(values_only=True):
        raw_rows.append([normalize_text(c) for c in row])

    if not raw_rows:
        raise ValueError("Step2 xlsx にデータがありません。")

    header = raw_rows[0]
    data_rows = raw_rows[1:]
    header_map = detect_header_map(header)

    source_path_idx = find_source_path_col_by_scan(raw_rows, header_map)
    if source_path_idx is None:
        raise ValueError("Step2 xlsx から顧客XMLパス列を特定できませんでした。")

    if "source_path" not in header_map:
        header_map["source_path"] = source_path_idx

    # selected列がなければ B列
    if "selected" not in header_map:
        header_map["selected"] = 1

    rows: List[Dict[str, str]] = []

    for excel_row_no, row in enumerate(data_rows, start=2):
        selected_val = row[header_map["selected"]] if header_map["selected"] < len(row) else ""
        if not is_selected(selected_val):
            continue

        source_path = row[header_map["source_path"]] if header_map["source_path"] < len(row) else ""
        if not source_path:
            continue

        item = {
            "excel_row_no": str(excel_row_no),
            "selected": selected_val,
            "source_path": source_path,
            "source_name": row[header_map["source_name"]] if "source_name" in header_map and header_map["source_name"] < len(row) else "",
            "bms_path": row[header_map["bms_path"]] if "bms_path" in header_map and header_map["bms_path"] < len(row) else "",
            "bms_name": row[header_map["bms_name"]] if "bms_name" in header_map and header_map["bms_name"] < len(row) else "",
            "note": row[header_map["note"]] if "note" in header_map and header_map["note"] < len(row) else "",
        }
        rows.append(item)

    return rows, header_map


def build_xml_leaf_map(xml_path: Path) -> Dict[str, Dict[str, str]]:
    """
    Step1 XML を解析し、leaf要素のパス辞書を返す。
    戻り値:
        {
          "/Order/Header/OrderNo": {
              "tag": "OrderNo",
              "text": "12345",
              "parent": "/Order/Header",
          },
          ...
        }
    """
    tree = ET.parse(xml_path)
    root = tree.getroot()

    leaf_map: Dict[str, Dict[str, str]] = {}

    def walk(elem: ET.Element, current_path: str) -> None:
        children = list(elem)
        tag_name = strip_namespace(elem.tag)
        next_path = f"{current_path}/{tag_name}" if current_path else f"/{tag_name}"

        if not children:
            leaf_map[next_path] = {
                "tag": tag_name,
                "text": normalize_text(elem.text),
                "parent": current_path if current_path else "/",
            }
            return

        for child in children:
            walk(child, next_path)

    walk(root, "")
    return leaf_map


def read_template_root_info(template_file: Path) -> Dict[str, str]:
    """
    テンプレートXMLの最低限のメタ情報を取得
    """
    try:
        tree = ET.parse(template_file)
        root = tree.getroot()
        return {
            "root_tag": strip_namespace(root.tag),
            "template_name": template_file.name,
        }
    except Exception:
        return {
            "root_tag": "",
            "template_name": template_file.name,
        }


def indent_xml(elem: ET.Element, level: int = 0) -> None:
    i = "\n" + level * "  "
    if len(elem):
        if not elem.text or not elem.text.strip():
            elem.text = i + "  "
        for child in elem:
            indent_xml(child, level + 1)
        if not child.tail or not child.tail.strip():
            child.tail = i
    else:
        if level and (not elem.tail or not elem.tail.strip()):
            elem.tail = i


def build_from_layout_xml(
    input_xml: Path,
    step2_xlsx: Path,
    template_file: Path,
    selected_rows: List[Dict[str, str]],
    xml_leaf_map: Dict[str, Dict[str, str]],
) -> ET.ElementTree:
    """
    from_layout XML の中間表現を作る
    """
    template_info = read_template_root_info(template_file)

    root = ET.Element(DEFAULT_OUTPUT_ROOT_TAG)
    root.set("version", VERSION)

    meta = ET.SubElement(root, "meta")
    meta.set("input_xml", str(input_xml))
    meta.set("mapping_xlsx", str(step2_xlsx))
    meta.set("template_file", str(template_file))
    meta.set("template_root_tag", template_info["root_tag"])
    meta.set("template_name", template_info["template_name"])

    summary = ET.SubElement(root, "summary")
    summary.set("selected_count", str(len(selected_rows)))
    summary.set("xml_leaf_count", str(len(xml_leaf_map)))

    records_elem = ET.SubElement(root, "records")

    # まずは単一 record として作る
    record_elem = ET.SubElement(records_elem, DEFAULT_RECORD_TAG)
    record_elem.set("name", "customer_source")
    record_elem.set("path", "/")
    record_elem.set("occurs", "1")

    used_paths = set()

    for seq, row in enumerate(selected_rows, start=1):
        src_path = normalize_text(row["source_path"])
        if not src_path:
            continue
        if src_path in used_paths:
            continue

        field = ET.SubElement(record_elem, DEFAULT_FIELD_TAG)
        field.set("seq", str(seq))
        field.set("source_path", src_path)
        field.set("source_name", row["source_name"] or Path(src_path).name)
        field.set("bms_path", row["bms_path"])
        field.set("bms_name", row["bms_name"])
        field.set("excel_row_no", row["excel_row_no"])

        xml_info = xml_leaf_map.get(src_path)
        if xml_info:
            field.set("exists_in_step1_xml", "true")
            field.set("xml_tag", xml_info["tag"])
            field.set("sample_value", xml_info["text"])
            field.set("parent_path", xml_info["parent"])
        else:
            field.set("exists_in_step1_xml", "false")
            field.set("xml_tag", "")
            field.set("sample_value", "")
            field.set("parent_path", "")

        note = row["note"]
        if note:
            note_elem = ET.SubElement(field, "note")
            note_elem.text = note

        used_paths.add(src_path)

    warnings = ET.SubElement(root, "warnings")
    for row in selected_rows:
        src_path = normalize_text(row["source_path"])
        if src_path and src_path not in xml_leaf_map:
            warn = ET.SubElement(warnings, "warning")
            warn.set("type", "missing_source_path_in_step1_xml")
            warn.set("excel_row_no", row["excel_row_no"])
            warn.text = src_path

    return ET.ElementTree(root)


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Step3: 変換元レイアウトXML作成")
    parser.add_argument("--input-xml", required=True, help="Step1出力XML (order_x2f_step1.xml)")
    parser.add_argument("--mapping-xlsx", required=True, help="Step2出力xlsx (order_x2f_step2.xlsx)")
    parser.add_argument("--template-file", required=True, help="基本形テンプレートXML")
    parser.add_argument("--output-xml", required=True, help="Step3出力XML (order_x2f_step3_from_layout.xml)")
    return parser.parse_args()


def validate_inputs(input_xml: Path, mapping_xlsx: Path, template_file: Path) -> None:
    if not input_xml.exists():
        raise FileNotFoundError(f"入力XMLが見つかりません: {input_xml}")
    if not mapping_xlsx.exists():
        raise FileNotFoundError(f"Step2 xlsx が見つかりません: {mapping_xlsx}")
    if not template_file.exists():
        raise FileNotFoundError(f"テンプレートXMLが見つかりません: {template_file}")


def main() -> int:
    args = parse_args()

    input_xml = Path(args.input_xml)
    mapping_xlsx = Path(args.mapping_xlsx)
    template_file = Path(args.template_file)
    output_xml = Path(args.output_xml)

    try:
        log(f"[INFO] order_x2f_step3.py VERSION={VERSION}")

        validate_inputs(input_xml, mapping_xlsx, template_file)

        log("[INFO] Step1 XML 解析開始")
        xml_leaf_map = build_xml_leaf_map(input_xml)
        log(f"[INFO] Step1 XML leaf数: {len(xml_leaf_map)}")

        log("[INFO] Step2 xlsx 読み込み開始")
        selected_rows, header_map = load_step2_rows(mapping_xlsx)
        log(f"[INFO] 採用行数: {len(selected_rows)}")
        log(f"[INFO] 検出ヘッダ: {header_map}")

        log("[INFO] Step3 from_layout XML 生成開始")
        tree = build_from_layout_xml(
            input_xml=input_xml,
            step2_xlsx=mapping_xlsx,
            template_file=template_file,
            selected_rows=selected_rows,
            xml_leaf_map=xml_leaf_map,
        )

        output_xml.parent.mkdir(parents=True, exist_ok=True)
        indent_xml(tree.getroot())
        tree.write(output_xml, encoding="utf-8", xml_declaration=True)

        log(f"[INFO] 出力完了: {output_xml}")
        return 0

    except Exception as e:
        log(f"[ERROR] Step3失敗: {e}")
        return 1


if __name__ == "__main__":
    sys.exit(main())