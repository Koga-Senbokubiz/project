#!/usr/bin/env python3
# -*- coding: utf-8 -*-
from __future__ import annotations

import argparse
from pathlib import Path
import xml.etree.ElementTree as ET

SKIP_ITEM_IDS = {"record_type"}
BINARY_ITEM_ATTR = "3"   # B:8ビット単位のビット列
BINARY_CHARSET = "9"     # バイナリ
GROUP_SIZE = 3            # PHP互換: 3項目で1セット

ITEM_FIXED_TEMPLATE = {
    "文字コード": BINARY_CHARSET,
    "桁数": "1",
    "小数部桁数": "0",
    "位置": "0",
    "パディング": "1",
    "変換エラー": "0",
    "置換コード": "",
    "SISO": "0",
    "EBCDICコードタイプ": "1",
    "JEF_KEIS漢字コードタイプ": "1",
    "外字使用": "False",
    "SISOを出力する": "True",
    "全角空白のパディングコード": "0",
}

ITEM_AUX_TEMPLATE = {
    "全角": "False",
    "大文字": "False",
    "暦書式": "",
    "暦エラータイプ": "0",
    "レコード出力タイプ": "0",
    "最終出力タイプ": "1",
    "最終出力値": "",
    "項目出力": "True",
    "全角半角変換": "True",
    "大文字小文字変換": "True",
}


def indent(elem: ET.Element, level: int = 0) -> None:
    i = "\n" + level * "  "
    if len(elem):
        if not elem.text or not elem.text.strip():
            elem.text = i + "  "
        for child in elem:
            indent(child, level + 1)
        if not elem[-1].tail or not elem[-1].tail.strip():
            elem[-1].tail = i
    if level and (not elem.tail or not elem.tail.strip()):
        elem.tail = i


def max_numeric_attr(elements, attr_name: str) -> int:
    nums = []
    for e in elements:
        v = e.get(attr_name)
        if v and v.isdigit():
            nums.append(int(v))
    return max(nums) if nums else 0


def build_bundle(
    item_id_num: int,
    layout_id_num: int,
    parent_layout_id: str,
    name: str,
    value_hex: str,
):
    layout = ET.Element("レイアウト", {
        "ID": str(layout_id_num),
        "Name": name,
        "Type": "5",
        "PropertyID": str(item_id_num),
        "ParentsID": str(parent_layout_id),
    })

    item = ET.Element("項目", {
        "ID": str(item_id_num),
        "項目名": name,
        "項目ID": name,
        "属性": BINARY_ITEM_ATTR,
        "属性チェック": "0",
    })

    fixed = ET.Element("項目Fixed", {
        "ID": str(item_id_num),
        **ITEM_FIXED_TEMPLATE
    })

    aux = ET.Element("項目補助情報", {
        "ID": str(item_id_num),
        **{**ITEM_AUX_TEMPLATE, "最終出力値": value_hex}
    })

    return layout, item, fixed, aux


class BundleFactory:
    def __init__(self, next_item_id: int, next_layout_id: int):
        self.next_item_id = next_item_id
        self.next_layout_id = next_layout_id
        self.ext_seq = 1
        self.new_items: list[ET.Element] = []
        self.new_fixeds: list[ET.Element] = []
        self.new_auxs: list[ET.Element] = []

    def create(self, parent_id: str, value_hex: str) -> ET.Element:
        name = f"EXT{self.ext_seq:03d}"
        layout, item, fixed, aux = build_bundle(
            self.next_item_id,
            self.next_layout_id,
            parent_id,
            name,
            value_hex,
        )
        self.new_items.append(item)
        self.new_fixeds.append(fixed)
        self.new_auxs.append(aux)
        self.next_item_id += 1
        self.next_layout_id += 1
        self.ext_seq += 1
        return layout


def is_target_child(root: ET.Element, child: ET.Element) -> bool:
    prop_id = child.get("PropertyID", "")
    item_def = root.find(f"./項目[@ID='{prop_id}']")
    if item_def is None:
        return False
    return item_def.get("項目ID", "") not in SKIP_ITEM_IDS


def main() -> None:
    ap = argparse.ArgumentParser(
        description="EasyExchange XMLへPHP互換の区切り項目(0x01/0x02)を挿入"
    )
    ap.add_argument("input_xml")
    ap.add_argument("-o", "--output")
    args = ap.parse_args()

    input_path = Path(args.input_xml)
    output_path = (
        Path(args.output)
        if args.output
        else input_path.with_name(input_path.stem + "_php_compatible.xml")
    )

    tree = ET.parse(input_path)
    root = tree.getroot()

    all_layouts = root.findall("./レイアウト")
    all_items = root.findall("./項目")
    all_fixeds = root.findall("./項目Fixed")
    all_auxs = root.findall("./項目補助情報")

    if not (all_layouts and all_items and all_fixeds and all_auxs):
        raise RuntimeError("必要セクションが不足しています。")

    record_layouts = [e for e in all_layouts if e.get("Type") == "3"]
    if not record_layouts:
        raise RuntimeError("Type=3 のレコードが見つかりません。")

    factory = BundleFactory(
        next_item_id=max_numeric_attr(all_items, "ID") + 1,
        next_layout_id=max_numeric_attr(all_layouts, "ID") + 1,
    )

    before_map: dict[str, list[ET.Element]] = {}
    after_map: dict[str, list[ET.Element]] = {}

    for record_layout in record_layouts:
        parent_id = record_layout.get("ID")
        child_layouts = [
            e for e in all_layouts
            if e.get("Type") == "5" and e.get("ParentsID") == parent_id
        ]
        if not child_layouts:
            continue

        target_children = [child for child in child_layouts if is_target_child(root, child)]
        if not target_children:
            continue

        if len(target_children) % GROUP_SIZE != 0:
            raise RuntimeError(
                f"レコードID={parent_id} の対象項目数が{GROUP_SIZE}の倍数ではありません: "
                f"{len(target_children)}項目"
            )

        for idx, child in enumerate(target_children):
            group_pos = idx % GROUP_SIZE
            is_group_end = group_pos == GROUP_SIZE - 1

            # 2項目目・3項目目の前に 0x01 を挿入
            if group_pos in (1, 2):
                before_map.setdefault(child.get("ID", ""), []).append(
                    factory.create(parent_id, "01")
                )

            # 3項目目の後ろに 0x01, 0x02 を挿入
            # その後 insert_0xff.py の 0x01 0x02 -> 0xFF で PHP互換になる
            if is_group_end:
                after_map.setdefault(child.get("ID", ""), []).append(
                    factory.create(parent_id, "01")
                )
                after_map.setdefault(child.get("ID", ""), []).append(
                    factory.create(parent_id, "02")
                )

    children = list(root)
    root.clear()

    layout_section: list[ET.Element] = []
    item_section: list[ET.Element] = []
    fixed_section: list[ET.Element] = []
    aux_section: list[ET.Element] = []

    for elem in children:
        if elem.tag == "レイアウト":
            layout_section.extend(before_map.get(elem.get("ID", ""), []))
            layout_section.append(elem)
            layout_section.extend(after_map.get(elem.get("ID", ""), []))
        elif elem.tag == "項目":
            item_section.append(elem)
        elif elem.tag == "項目Fixed":
            fixed_section.append(elem)
        elif elem.tag == "項目補助情報":
            aux_section.append(elem)

    emitted_layout = False
    emitted_item = False
    emitted_fixed = False
    emitted_aux = False

    for elem in children:
        tag = elem.tag

        if tag == "レイアウト":
            if not emitted_layout:
                for x in layout_section:
                    root.append(x)
                emitted_layout = True
            continue

        if tag == "項目":
            if not emitted_item:
                for x in item_section:
                    root.append(x)
                for x in factory.new_items:
                    root.append(x)
                emitted_item = True
            continue

        if tag == "項目Fixed":
            if not emitted_fixed:
                for x in fixed_section:
                    root.append(x)
                for x in factory.new_fixeds:
                    root.append(x)
                emitted_fixed = True
            continue

        if tag == "項目補助情報":
            if not emitted_aux:
                for x in aux_section:
                    root.append(x)
                for x in factory.new_auxs:
                    root.append(x)
                emitted_aux = True
            continue

        root.append(elem)

    indent(root)
    tree.write(output_path, encoding="utf-8", xml_declaration=True)
    print(f"出力完了: {output_path}")


if __name__ == "__main__":
    main()
