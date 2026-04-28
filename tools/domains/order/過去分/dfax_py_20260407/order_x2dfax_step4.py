#!/usr/bin/env python
# -*- coding: utf-8 -*-

"""
order_x2dfax_step5.py

目的:
    EasyExchange の Setting 配下に置くマッピング XML を生成する。

前提:
    テンプレート XML は PatternInfo 形式で、
    <マッピング ... 変換元レイアウト="..." 変換先レイアウト="..." ロジック="..." ... />
    を持つ。

役割:
    -sx     : 変換元レイアウト XML 名
    -tx     : 変換先レイアウト XML 名
    -logic  : Logic 配下の実ロジック XML 名
    -base   : Setting 用テンプレート XML
    -o      : 出力する Setting XML

使用例:
    python order_x2dfax_step5.py ^
      -sx "bbord_dfax_xml.xml" ^
      -tx "bbord_dfax_fax.xml" ^
      -logic "bbord_dfax_bbord_dfax_xml_bbord_dfax_fax.xml" ^
      -base "基本形1_3：発注Ver1_3（XML→JCA128）.xml" ^
      -o "bbord_dfax.xml"
"""

from __future__ import annotations

import argparse
import sys
from pathlib import Path
from typing import Tuple
import xml.etree.ElementTree as ET


def eprint(*args, **kwargs) -> None:
    print(*args, file=sys.stderr, **kwargs)


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="EasyExchange Setting 配下のマッピング XML を生成する"
    )
    parser.add_argument("-sx", "--source-xml", required=True, help="変換元レイアウト XML 名")
    parser.add_argument("-tx", "--target-xml", required=True, help="変換先レイアウト XML 名")
    parser.add_argument("-logic", "--logic-xml", required=True, help="Logic 配下の実ロジック XML 名")
    parser.add_argument("-base", "--base-xml", required=True, help="テンプレート Setting XML")
    parser.add_argument("-o", "--output-xml", required=True, help="出力する Setting XML")
    return parser.parse_args()


def detect_encoding_and_read(path: Path) -> Tuple[str, str]:
    raw = path.read_bytes()
    candidates = [
        "utf-8-sig",
        "utf-8",
        "cp932",
        "shift_jis",
    ]
    for enc in candidates:
        try:
            return raw.decode(enc), enc
        except UnicodeDecodeError:
            continue
    raise ValueError(f"文字コードを判定できませんでした: {path}")


def write_text(path: Path, text: str, encoding: str) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(text, encoding=encoding, newline="")


def find_mapping_element(root: ET.Element) -> ET.Element:
    """
    PatternInfo 配下の <マッピング> を取得する。
    名前空間なし前提だが、タグ名末尾判定で多少ゆるく探す。
    """
    for elem in root.iter():
        tag = elem.tag.split("}")[-1]
        if tag == "マッピング":
            return elem
    raise ValueError("テンプレート XML 内に <マッピング> 要素が見つかりません。")


def indent_xml(elem: ET.Element, level: int = 0) -> None:
    """
    見やすい形で出力するための簡易インデント。
    """
    i = "\n" + level * "  "
    if len(elem):
        if not elem.text or not elem.text.strip():
            elem.text = i + "  "
        for child in elem:
            indent_xml(child, level + 1)
        if not child.tail or not child.tail.strip():
            child.tail = i
    if level and (not elem.tail or not elem.tail.strip()):
        elem.tail = i


def main() -> int:
    args = parse_args()

    source_xml = Path(args.source_xml).name
    target_xml = Path(args.target_xml).name
    logic_xml = Path(args.logic_xml).name
    base_xml_path = Path(args.base_xml)
    output_xml_path = Path(args.output_xml)

    if not base_xml_path.exists():
        eprint(f"[ERROR] テンプレート XML が存在しません: {base_xml_path}")
        return 1

    try:
        base_text, encoding = detect_encoding_and_read(base_xml_path)
    except Exception as ex:
        eprint(f"[ERROR] テンプレート XML の読み込みに失敗しました: {base_xml_path}")
        eprint(f"        {ex}")
        return 1

    try:
        root = ET.fromstring(base_text)
    except ET.ParseError as ex:
        eprint(f"[ERROR] テンプレート XML の解析に失敗しました: {base_xml_path}")
        eprint(f"        {ex}")
        return 1

    try:
        mapping = find_mapping_element(root)
    except Exception as ex:
        eprint(f"[ERROR] {ex}")
        return 1

    before_source = mapping.get("変換元レイアウト", "")
    before_target = mapping.get("変換先レイアウト", "")
    before_logic = mapping.get("ロジック", "")

    mapping.set("変換元レイアウト", source_xml)
    mapping.set("変換先レイアウト", target_xml)
    mapping.set("ロジック", logic_xml)

    indent_xml(root)

    xml_text = ET.tostring(root, encoding="unicode")

    if base_text.lstrip().startswith('<?xml'):
        xml_text = '<?xml version="1.0" standalone="yes"?>\n' + xml_text
    else:
        xml_text = '<?xml version="1.0"?>\n' + xml_text

    try:
        write_text(output_xml_path, xml_text, encoding)
    except Exception as ex:
        eprint(f"[ERROR] 出力 XML の書き込みに失敗しました: {output_xml_path}")
        eprint(f"        {ex}")
        return 1

    print("[INFO] Step5 Setting XML を作成しました")
    print(f"       TEMPLATE : {base_xml_path}")
    print(f"       OUTPUT   : {output_xml_path}")
    print(f"       ENCODING : {encoding}")
    print("[INFO] マッピング属性更新")
    print(f"       変換元レイアウト : {before_source} -> {source_xml}")
    print(f"       変換先レイアウト : {before_target} -> {target_xml}")
    print(f"       ロジック         : {before_logic} -> {logic_xml}")

    if mapping.get("変換元レイアウト") != source_xml:
        eprint("[WARN] 変換元レイアウトの更新結果が想定と一致しません。")
    if mapping.get("変換先レイアウト") != target_xml:
        eprint("[WARN] 変換先レイアウトの更新結果が想定と一致しません。")
    if mapping.get("ロジック") != logic_xml:
        eprint("[WARN] ロジックの更新結果が想定と一致しません。")

    return 0


if __name__ == "__main__":
    sys.exit(main())