# -*- coding: utf-8 -*-
"""
order_x2f_step4.py
Step4: 変換先レイアウトXML（固定長/JCA128）を生成する

方針:
- Step4 は変換先テンプレートXMLをそのまま採用する
- ただし工程の整合確認のため、以下3入力を正式引数として受ける
  1) order_x2f_step1.xml
  2) order_bigboss_dictionary.xlsx
  3) 基本形1_3：発注JCA128.xml
- input_xml / customer_dictionary は主に存在確認・件数確認用
- 実際のStep4出力は template_file を複製して保存する
"""

from __future__ import annotations

import argparse
import copy
import os
import sys
import traceback
import xml.etree.ElementTree as ET
from pathlib import Path
from typing import Optional

import pandas as pd


def log(msg: str) -> None:
    print(msg, flush=True)


def normalize_text(value: object) -> str:
    if value is None:
        return ""
    try:
        if pd.isna(value):
            return ""
    except Exception:
        pass
    return str(value).strip()


def indent_xml(elem: ET.Element, level: int = 0) -> None:
    i = "\n" + level * "  "
    if len(elem):
        if not elem.text or not elem.text.strip():
            elem.text = i + "  "
        last_child = None
        for child in elem:
            indent_xml(child, level + 1)
            last_child = child
        if last_child is not None and (not last_child.tail or not last_child.tail.strip()):
            last_child.tail = i
    elif level and (not elem.tail or not elem.tail.strip()):
        elem.tail = i


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Order-X2F Step4")

    # 正式引数
    parser.add_argument("-i", "--input-xml", required=True, help="order_x2f_step1.xml")
    parser.add_argument("-d", "--customer-dictionary", required=True, help="order_bigboss_dictionary.xlsx")
    parser.add_argument("-t", "--template-file", required=True, help="基本形1_3：発注JCA128.xml")
    parser.add_argument("-o", "--output-xml", required=True, help="Step4 output layout xml")

    # 互換用
    parser.add_argument("--mapping-xlsx", required=False, help="Compatibility option (unused)")
    parser.add_argument("-s", "--sheet-name", default=None, help="顧客辞書の対象シート名")
    return parser.parse_args()


def validate_args(args: argparse.Namespace) -> None:
    if not os.path.isfile(args.input_xml):
        raise FileNotFoundError(f"INPUT_XML が見つかりません: {args.input_xml}")
    if not os.path.isfile(args.customer_dictionary):
        raise FileNotFoundError(f"DICTIONARY が見つかりません: {args.customer_dictionary}")
    if not os.path.isfile(args.template_file):
        raise FileNotFoundError(f"TEMPLATE_FILE が見つかりません: {args.template_file}")


def count_xml_leaf_nodes(xml_path: str) -> int:
    tree = ET.parse(xml_path)
    root = tree.getroot()
    count = 0
    for elem in root.iter():
        if len(list(elem)) == 0:
            count += 1
    return count


def load_customer_dictionary(xlsx_path: str, sheet_name: Optional[str] = None) -> pd.DataFrame:
    target_sheet = sheet_name if sheet_name else 0
    df = pd.read_excel(xlsx_path, sheet_name=target_sheet, header=0, dtype=str).fillna("")

    if isinstance(df, dict):
        if not df:
            raise ValueError(f"顧客辞書の読み込み結果が空です: {xlsx_path}")
        df = next(iter(df.values()))

    if len(df) > 0:
        first_row_values = [normalize_text(v) for v in df.iloc[0].tolist()]
        if "状態" in first_row_values or "顧客ID" in first_row_values:
            df = df.iloc[1:].reset_index(drop=True)

    required_cols = ["status", "customer_path", "customer_tag", "candidate_bms_field_id", "confirmed_bms_field_id"]
    for col in required_cols:
        if col not in df.columns:
            df[col] = ""

    return df


def count_step4_targets(df: pd.DataFrame) -> int:
    work = df.copy()
    def is_target(row) -> bool:
        status = normalize_text(row.get("status", ""))
        customer_path = normalize_text(row.get("customer_path", ""))
        customer_tag = normalize_text(row.get("customer_tag", ""))
        if not customer_path or not customer_tag:
            return False
        if status == "標準項目":
            return True
        if normalize_text(row.get("confirmed_bms_field_id", "")) or normalize_text(row.get("candidate_bms_field_id", "")):
            return True
        return False

    return int(work.apply(is_target, axis=1).sum())


def load_template(template_file: str) -> ET.ElementTree:
    try:
        tree = ET.parse(template_file)
        root = tree.getroot()
        if root.tag != "マッピングレイアウト":
            raise ValueError(f"テンプレートXMLのルートが想定外です: {root.tag}")
        return tree
    except Exception as e:
        raise ValueError(f"テンプレートXMLの解析に失敗しました: {template_file} / {e}")


def count_template_nodes(root: ET.Element) -> tuple[int, int, int]:
    layout_count = len(root.findall(".//レイアウト"))
    record_count = len(root.findall(".//レコード"))
    item_count = len(root.findall(".//項目"))
    return layout_count, record_count, item_count


def save_xml(tree: ET.ElementTree, output_xml: str) -> None:
    out_dir = os.path.dirname(output_xml)
    if out_dir:
        os.makedirs(out_dir, exist_ok=True)

    root = tree.getroot()
    indent_xml(root)
    tree.write(output_xml, encoding="utf-8", xml_declaration=True)


def main() -> int:
    try:
        args = parse_args()
        validate_args(args)

        log("==================================================")
        log("Order-X2F Step4 start")
        log("==================================================")
        log(f"INPUT_XML     = {args.input_xml}")
        log(f"DICTIONARY    = {args.customer_dictionary}")
        log(f"TEMPLATE_FILE = {args.template_file}")
        log(f"OUTPUT_XML    = {args.output_xml}")
        if args.mapping_xlsx:
            log(f"MAPPING_XLSX  = {args.mapping_xlsx} (unused)")
        log("")

        log("[1/4] 入力XML確認...")
        leaf_count = count_xml_leaf_nodes(args.input_xml)
        log(f"  input xml leaf count = {leaf_count}")

        log("[2/4] 顧客辞書確認...")
        dict_df = load_customer_dictionary(args.customer_dictionary, sheet_name=args.sheet_name)
        dict_count = len(dict_df)
        target_count = count_step4_targets(dict_df)
        log(f"  dictionary row count   = {dict_count}")
        log(f"  step4 target row count = {target_count}")

        log("[3/4] テンプレート読込...")
        template_tree = load_template(args.template_file)
        root = template_tree.getroot()
        layout_count, record_count, item_count = count_template_nodes(root)
        log(f"  root tag     = {root.tag}")
        log(f"  layout count = {layout_count}")
        log(f"  record count = {record_count}")
        log(f"  item count   = {item_count}")

        log("[4/4] テンプレート複製・保存...")
        output_tree = copy.deepcopy(template_tree)
        save_xml(output_tree, args.output_xml)

        log("")
        log("==================================================")
        log("Order-X2F Step4 finished")
        log(f"OUTPUT_XML = {args.output_xml}")
        log("==================================================")
        return 0

    except Exception as e:
        log("")
        log("[ERROR] Step4 failed")
        log(str(e))
        log("")
        traceback.print_exc()
        return 1


if __name__ == "__main__":
    sys.exit(main())
