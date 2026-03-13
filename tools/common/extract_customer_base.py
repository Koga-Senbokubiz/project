#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
Step1: 顧客データXMLから「基準XML（1件抽出版）」を生成する。

目的:
- データ内容ではなく「タグ構造」を比較するための基準XMLを作成する
- 同一親配下にある同一タグの繰り返し要素は、先頭1件のみ残す
- 正規化（normalize）や業務的な意味付けは行わない

Usage:
  python extract_customer_data_xml.py input.xml output.xml
"""

import argparse
import os
import sys
import xml.etree.ElementTree as ET
from collections import defaultdict


def remove_duplicate_children_keep_first(parent: ET.Element) -> None:
    """
    親要素直下の子要素について、
    同一タグ（namespace込み）が複数ある場合は先頭1件のみ残す
    """
    children = list(parent)
    if not children:
        return

    seen = defaultdict(int)
    to_remove = []

    for child in children:
        key = child.tag  # {namespace}local-name
        seen[key] += 1
        if seen[key] > 1:
            to_remove.append(child)

    for child in to_remove:
        parent.remove(child)


def shrink_xml_to_single_record(root: ET.Element) -> None:
    """
    XML全体を走査し、すべての階層で繰り返し要素を1件に縮退する
    """
    stack = [root]
    while stack:
        node = stack.pop()
        remove_duplicate_children_keep_first(node)
        for child in list(node):
            stack.append(child)


def main() -> int:
    parser = argparse.ArgumentParser(
        description="Step1: 顧客データXMLを1件抽出版に変換する"
    )
    parser.add_argument("input_xml", help="入力XMLファイル")
    parser.add_argument("output_xml", help="出力XMLファイル")
    args = parser.parse_args()

    if not os.path.exists(args.input_xml):
        print(f"[ERROR] Input XML not found: {args.input_xml}", file=sys.stderr)
        return 2

    try:
        tree = ET.parse(args.input_xml)
        root = tree.getroot()

        shrink_xml_to_single_record(root)

        out_dir = os.path.dirname(os.path.abspath(args.output_xml))
        if out_dir and not os.path.exists(out_dir):
            os.makedirs(out_dir, exist_ok=True)

        tree.write(
            args.output_xml,
            encoding="utf-8",
            xml_declaration=True
        )

        print(f"[OK] Step1 base XML created: {args.output_xml}")
        return 0

    except ET.ParseError as e:
        print(f"[ERROR] XML parse error: {e}", file=sys.stderr)
        return 3
    except Exception as e:
        print(f"[ERROR] Unexpected error: {e}", file=sys.stderr)
        return 9


if __name__ == "__main__":
    sys.exit(main())
