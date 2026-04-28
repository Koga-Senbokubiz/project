# -*- coding: utf-8 -*-
"""
order_x2f_step4.py
Ver. 2026-03-18 final-template-copy edition

目的:
- 変換先テンプレートレイアウトXML（固定長/JCA128）を
  Step4 出力XMLとしてそのまま採用する。

考え方:
- 変換元XMLは実データに合わせて削減する価値がある
- 変換先固定長テンプレートは、JCA128の完成レイアウトなので削減しない
- 実際の採用/不採用は Step5 のマッピング作成で扱う

想定引数:
  --template-file
  --output-xml

互換用:
  --input-xml
  --mapping-xlsx
  が渡されても無視可能
"""

from __future__ import annotations

import argparse
import copy
import os
import sys
import traceback
import xml.etree.ElementTree as ET


def log(msg: str) -> None:
    print(msg, flush=True)


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


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Order-X2F Step4")
    parser.add_argument("--template-file", required=True, help="Target template layout xml")
    parser.add_argument("--output-xml", required=True, help="Step4 output layout xml")

    # 互換用（batから渡されても止まらないようにする）
    parser.add_argument("--input-xml", required=False, help="Compatibility option (unused)")
    parser.add_argument("--mapping-xlsx", required=False, help="Compatibility option (unused)")
    return parser.parse_args()


def validate_args(args: argparse.Namespace) -> None:
    if not os.path.isfile(args.template_file):
        raise FileNotFoundError(f"TEMPLATE_FILE が見つかりません: {args.template_file}")


def load_template(template_file: str) -> ET.ElementTree:
    try:
        tree = ET.parse(template_file)
        root = tree.getroot()
        if root.tag != "マッピングレイアウト":
            raise ValueError(f"テンプレートXMLのルートが想定外です: {root.tag}")
        return tree
    except Exception as e:
        raise ValueError(f"テンプレートXMLの解析に失敗しました: {template_file} / {e}")


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
        log(f"TEMPLATE_FILE = {args.template_file}")
        log(f"OUTPUT_XML    = {args.output_xml}")
        if args.input_xml:
            log(f"INPUT_XML     = {args.input_xml} (unused)")
        if args.mapping_xlsx:
            log(f"MAPPING_XLSX  = {args.mapping_xlsx} (unused)")
        log("")

        log("[1/3] テンプレート読込...")
        template_tree = load_template(args.template_file)
        root = template_tree.getroot()
        log(f"  root tag = {root.tag}")

        layout_count = len(root.findall("レイアウト"))
        record_count = len(root.findall("レコード"))
        item_count = len(root.findall("項目"))
        log(f"  layout count = {layout_count}")
        log(f"  record count = {record_count}")
        log(f"  item count   = {item_count}")

        log("[2/3] テンプレート複製...")
        output_tree = copy.deepcopy(template_tree)

        log("[3/3] 保存...")
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