# -*- coding: utf-8 -*-
"""
order_x2f_step5.py
Step5: Step3/Step4レイアウトXMLとロジックテンプレートXMLを使って、
       EasyExchange 用ロジックXMLを生成する（根本解決版）

方針:
- order_x2f_step1.xml は補助入力（存在確認・件数確認）
- 主入力は以下
  1) order_bigboss_dictionary.xlsx
  2) order_x2f_step3_from_layout.xml   （変換元レイアウト）
  3) order_x2f_step4_to_layout.xml     （変換先レイアウト）
  4) ロジックテンプレートXML           （EasyExchange のロジックXMLそのもの）
- ロジックテンプレートXMLの「線引き」「レコード接続」等は ID ベースで保持されるため、
  Step3/Step4レイアウトXMLから有効な項目ID/レコードIDを抽出して絞り込む
- 顧客辞書は「有効な変換元項目」を判定するために使う
"""

from __future__ import annotations

import argparse
import os
import sys
from pathlib import Path
from typing import Dict, List, Optional, Set, Tuple
import xml.etree.ElementTree as ET

import pandas as pd


CHILD_TYPE_MAP = {
    "0": "レコード接続",
    "1": "線引き",
    "2": "ロジックパラメータ",
    "3": "CSVテーブル検索",
    "4": "DBテーブル検索",
    "5": "COMメソッド呼び出し",
    "6": "VBScript呼び出し",
    "7": "パックゾーン入力",
    "8": "パックゾーン出力",
    "9": "結合",
    "10": "文字列操作",
    "11": "文字列操作コマンド",
    "12": "連番",
    "13": "件数",
    "14": "合計",
    "15": "CIIテーブル検索",
    "16": "固定値",
}

ITEM_ID_ATTRS = ["ID", "id", "項目ID", "itemId", "fieldId"]
RECORD_ID_ATTRS = ["ID", "id", "レコードID", "recordId"]

PATH_ATTRS = ["sourcePath", "xmlPath", "srcPath"]
SOURCE_TAG_ATTRS = ["sourceTag", "xmlTag", "srcTag", "physicalName", "xmlName", "tag"]
BMS_FIELD_ID_ATTRS = ["bmsFieldId", "fieldId"]
BMS_TAG_ATTRS = ["bmsTag", "targetTag", "name", "itemName", "fieldName", "label"]
ITEM_TEXT_HINT_ATTRS = ["name", "itemName", "fieldName", "label", "physicalName", "xmlName", "tag", "bmsTag", "targetTag"]


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


def normalize_key(value: str) -> str:
    s = normalize_text(value).lower()
    for ch in [" ", "　", "_", "-", ".", "/", "\\", "(", ")", "[", "]", "{", "}", ":", "："]:
        s = s.replace(ch, "")
    return s


def ensure_dir(path_text: str) -> None:
    d = os.path.dirname(path_text)
    if d:
        os.makedirs(d, exist_ok=True)


def local_name(tag: str) -> str:
    if not tag:
        return ""
    if "}" in tag:
        return tag.split("}", 1)[1]
    if ":" in tag:
        return tag.split(":", 1)[1]
    return tag


def indent_xml(elem: ET.Element, level: int = 0) -> None:
    i = "\n" + level * "  "
    if len(elem):
        if not elem.text or not elem.text.strip():
            elem.text = i + "  "
        for child in elem:
            indent_xml(child, level + 1)
        if child is not None and (not child.tail or not child.tail.strip()):
            child.tail = i
    if level and (not elem.tail or not elem.tail.strip()):
        elem.tail = i


def first_attr(elem: ET.Element, names: List[str]) -> str:
    lower_map = {k.lower(): v for k, v in elem.attrib.items()}
    for name in names:
        if name.lower() in lower_map:
            return normalize_text(lower_map[name.lower()])
    return ""


def to_set(values) -> Set[str]:
    return {normalize_text(v) for v in values if normalize_text(v)}


def parse_input_xml_leaf_count(input_xml: Path) -> int:
    tree = ET.parse(input_xml)
    root = tree.getroot()
    count = 0
    for elem in root.iter():
        if len(list(elem)) == 0:
            count += 1
    return count


def load_customer_dictionary(xlsx_path: Path, sheet_name: Optional[str] = None) -> pd.DataFrame:
    if not xlsx_path.exists():
        raise FileNotFoundError(f"顧客辞書が存在しません: {xlsx_path}")

    read_sheet = sheet_name if sheet_name else 0
    df = pd.read_excel(xlsx_path, sheet_name=read_sheet, header=0, dtype=str)
    if isinstance(df, dict):
        first_key = next(iter(df))
        df = df[first_key]
    df = df.fillna("")

    if len(df) > 0:
        first_row_values = [normalize_text(v) for v in df.iloc[0].tolist()]
        if "状態" in first_row_values or "顧客ID" in first_row_values:
            df = df.iloc[1:].reset_index(drop=True)

    required_cols = [
        "status", "customer_field_id", "customer_tag", "candidate_bms_field_id",
        "candidate_bms_tag", "normalized_customer_tag", "repeat_group", "data_type",
        "length_max", "required_flag", "repeat_flag", "customer_path", "parent_path",
        "field_class", "confirmed_bms_field_id", "match_method", "sample_value", "remarks"
    ]
    for col in required_cols:
        if col not in df.columns:
            df[col] = ""

    return df


def is_step5_target_row(row: pd.Series, include_customer_items: bool = False) -> bool:
    customer_path = normalize_text(row.get("customer_path", ""))
    customer_tag = normalize_text(row.get("customer_tag", ""))
    if not customer_path or not customer_tag:
        return False

    if include_customer_items:
        return True

    status = normalize_text(row.get("status", ""))
    if status == "標準項目":
        return True

    if normalize_text(row.get("confirmed_bms_field_id", "")) or normalize_text(row.get("candidate_bms_field_id", "")):
        return True

    return False


def build_active_dictionary_keys(df: pd.DataFrame) -> Dict[str, Set[str]]:
    active_paths = set()
    active_tags = set()
    active_norm_tags = set()
    active_bms_field_ids = set()
    active_bms_tags = set()

    for _, row in df.iterrows():
        active_paths.add(normalize_text(row.get("customer_path", "")))
        active_tags.add(normalize_key(row.get("customer_tag", "")))
        active_norm_tags.add(normalize_key(row.get("normalized_customer_tag", "")))
        active_bms_field_ids.add(normalize_text(row.get("confirmed_bms_field_id", "")) or normalize_text(row.get("candidate_bms_field_id", "")))
        active_bms_tags.add(normalize_key(row.get("candidate_bms_tag", "")))

    return {
        "paths": {x for x in active_paths if x},
        "tags": {x for x in active_tags if x},
        "norm_tags": {x for x in active_norm_tags if x},
        "bms_field_ids": {x for x in active_bms_field_ids if x},
        "bms_tags": {x for x in active_bms_tags if x},
    }


def is_record_elem(elem: ET.Element) -> bool:
    lname = local_name(elem.tag)
    if lname in {"レコード", "record"}:
        return True
    attrs = {k.lower() for k in elem.attrib.keys()}
    return "recordid" in attrs or "レコードid" in attrs


def is_item_elem(elem: ET.Element) -> bool:
    lname = local_name(elem.tag)
    if lname in {"項目", "item", "field"}:
        return True
    attrs = {k.lower() for k in elem.attrib.keys()}
    hint_attrs = {"sourcepath", "xmlpath", "srcpath", "sourcetag", "xmltag", "srctag", "bmsfieldid", "bmstag"}
    return any(x in attrs for x in hint_attrs)


def get_record_id(elem: ET.Element) -> str:
    return first_attr(elem, RECORD_ID_ATTRS)


def get_item_id(elem: ET.Element) -> str:
    return first_attr(elem, ITEM_ID_ATTRS)


def parse_layout_xml(layout_xml: Path) -> Dict[str, dict]:
    tree = ET.parse(layout_xml)
    root = tree.getroot()

    items_by_id: Dict[str, dict] = {}
    records_by_id: Dict[str, dict] = {}

    def walk(elem: ET.Element, current_record_id: str = "") -> None:
        nonlocal items_by_id, records_by_id

        record_id_here = current_record_id
        if is_record_elem(elem):
            rid = get_record_id(elem)
            if rid:
                record_id_here = rid
                records_by_id[rid] = {
                    "id": rid,
                    "tag": local_name(elem.tag),
                    "element": elem,
                }

        if is_item_elem(elem):
            item_id = get_item_id(elem)
            if item_id:
                source_path = first_attr(elem, PATH_ATTRS)
                source_tag = first_attr(elem, SOURCE_TAG_ATTRS)
                bms_field_id = first_attr(elem, BMS_FIELD_ID_ATTRS)
                bms_tag = first_attr(elem, BMS_TAG_ATTRS)
                text_hints = []
                for a in ITEM_TEXT_HINT_ATTRS:
                    v = first_attr(elem, [a])
                    if v:
                        text_hints.append(v)

                items_by_id[item_id] = {
                    "id": item_id,
                    "record_id": record_id_here,
                    "tag": local_name(elem.tag),
                    "source_path": source_path,
                    "source_tag": source_tag,
                    "bms_field_id": bms_field_id,
                    "bms_tag": bms_tag,
                    "text_hints": text_hints,
                    "element": elem,
                }

        for child in list(elem):
            walk(child, record_id_here)

    walk(root, "")
    return {
        "tree": tree,
        "root": root,
        "items_by_id": items_by_id,
        "records_by_id": records_by_id,
    }


def is_active_source_item(item_info: dict, active_keys: Dict[str, Set[str]]) -> bool:
    source_path = normalize_text(item_info.get("source_path", ""))
    source_tag = normalize_key(item_info.get("source_tag", ""))
    bms_field_id = normalize_text(item_info.get("bms_field_id", ""))
    bms_tag = normalize_key(item_info.get("bms_tag", ""))
    text_hints = [normalize_key(x) for x in item_info.get("text_hints", []) if normalize_text(x)]

    if source_path and source_path in active_keys["paths"]:
        return True
    if bms_field_id and bms_field_id in active_keys["bms_field_ids"]:
        return True
    if source_tag and (source_tag in active_keys["tags"] or source_tag in active_keys["norm_tags"] or source_tag in active_keys["bms_tags"]):
        return True
    if bms_tag and (bms_tag in active_keys["bms_tags"] or bms_tag in active_keys["tags"] or bms_tag in active_keys["norm_tags"]):
        return True
    for hint in text_hints:
        if hint in active_keys["bms_tags"] or hint in active_keys["tags"] or hint in active_keys["norm_tags"]:
            return True
    return False


def collect_active_source_ids(from_layout_info: Dict[str, dict], active_keys: Dict[str, Set[str]]) -> Tuple[Set[str], Set[str]]:
    active_item_ids: Set[str] = set()
    active_record_ids: Set[str] = set()

    for item_id, item_info in from_layout_info["items_by_id"].items():
        if is_active_source_item(item_info, active_keys):
            active_item_ids.add(item_id)
            rid = normalize_text(item_info.get("record_id", ""))
            if rid:
                active_record_ids.add(rid)

    return active_item_ids, active_record_ids


def split_id_list(value: str) -> List[str]:
    s = normalize_text(value)
    if not s:
        return []
    for sep in [";", "|", "、"]:
        s = s.replace(sep, ",")
    return [x.strip() for x in s.split(",") if x.strip()]


def filter_id_list(value: str, valid_ids: Set[str]) -> str:
    kept = [x for x in split_id_list(value) if x in valid_ids]
    return ",".join(kept)


def child_exists_by_id(root: ET.Element, tag_name: str, child_id: str) -> bool:
    for elem in root.findall(tag_name):
        if normalize_text(elem.get("ID", "")) == normalize_text(child_id):
            return True
    return False


def filter_logic_tree(
    logic_tree: ET.ElementTree,
    active_source_item_ids: Set[str],
    active_source_record_ids: Set[str],
    valid_target_item_ids: Set[str],
    valid_target_record_ids: Set[str],
) -> Dict[str, int]:
    root = logic_tree.getroot()

    removed_line_ids: Set[str] = set()
    removed_record_connection_ids: Set[str] = set()
    removed_other_nodes: Dict[str, Set[str]] = {tag: set() for tag in CHILD_TYPE_MAP.values()}

    # 1) 線引き
    for elem in list(root.findall("線引き")):
        src_id = normalize_text(elem.get("変換元項目ID", ""))
        tgt_id = normalize_text(elem.get("変換先項目ID", ""))
        keep = bool(src_id and tgt_id and src_id in active_source_item_ids and tgt_id in valid_target_item_ids)
        if not keep:
            removed_line_ids.add(normalize_text(elem.get("ID", "")))
            root.remove(elem)

    # 2) レコード接続
    for elem in list(root.findall("レコード接続")):
        src_id = normalize_text(elem.get("変換元レコードID", ""))
        tgt_id = normalize_text(elem.get("変換先レコードID", ""))
        keep = bool(src_id and tgt_id and src_id in active_source_record_ids and tgt_id in valid_target_record_ids)
        if not keep:
            removed_record_connection_ids.add(normalize_text(elem.get("ID", "")))
            root.remove(elem)

    # 3) 結合
    for elem in list(root.findall("結合")):
        src_ids = split_id_list(elem.get("変換元項目ID配列", ""))
        tgt_id = normalize_text(elem.get("変換先項目ID", ""))
        kept_src = [x for x in src_ids if x in active_source_item_ids]
        keep = bool(kept_src and tgt_id in valid_target_item_ids)
        if keep:
            elem.set("変換元項目ID配列", ",".join(kept_src))
        else:
            removed_other_nodes["結合"].add(normalize_text(elem.get("ID", "")))
            root.remove(elem)

    # 4) 文字列操作
    for elem in list(root.findall("文字列操作")):
        src_id = normalize_text(elem.get("変換元項目ID", ""))
        tgt_id = normalize_text(elem.get("変換先項目ID", ""))
        keep = bool(src_id in active_source_item_ids and tgt_id in valid_target_item_ids)
        if not keep:
            removed_other_nodes["文字列操作"].add(normalize_text(elem.get("ID", "")))
            root.remove(elem)

    # 5) 合計
    for elem in list(root.findall("合計")):
        src_id = normalize_text(elem.get("変換元項目ID", ""))
        tgt_id = normalize_text(elem.get("変換先項目ID", ""))
        base_record_id = normalize_text(elem.get("基準レコードID", ""))
        keep = bool(src_id in active_source_item_ids and tgt_id in valid_target_item_ids and (not base_record_id or base_record_id in valid_target_record_ids))
        if not keep:
            removed_other_nodes["合計"].add(normalize_text(elem.get("ID", "")))
            root.remove(elem)

    # 6) 固定値
    for elem in list(root.findall("固定値")):
        tgt_id = normalize_text(elem.get("変換先項目ID", ""))
        keep = bool(tgt_id in valid_target_item_ids)
        if not keep:
            removed_other_nodes["固定値"].add(normalize_text(elem.get("ID", "")))
            root.remove(elem)

    # 7) 件数
    for elem in list(root.findall("件数")):
        rec_id = normalize_text(elem.get("変換先レコードID", ""))
        tgt_id = normalize_text(elem.get("変換先項目ID", ""))
        keep = bool(rec_id in valid_target_record_ids and tgt_id in valid_target_item_ids)
        if not keep:
            removed_other_nodes["件数"].add(normalize_text(elem.get("ID", "")))
            root.remove(elem)

    # 8) 連番
    for elem in list(root.findall("連番")):
        rec_id = normalize_text(elem.get("変換先レコードID", ""))
        tgt_ids = split_id_list(elem.get("変換先項目ID配列", ""))
        kept_tgt_ids = [x for x in tgt_ids if x in valid_target_item_ids]
        keep = bool(rec_id in valid_target_record_ids and kept_tgt_ids)
        if keep:
            elem.set("変換先項目ID配列", ",".join(kept_tgt_ids))
        else:
            removed_other_nodes["連番"].add(normalize_text(elem.get("ID", "")))
            root.remove(elem)

    # 9) パックゾーン入力
    for elem in list(root.findall("パックゾーン入力")):
        in_id = normalize_text(elem.get("バイナリ入力項目ID", ""))
        out_num_id = normalize_text(elem.get("数値出力項目ID", ""))
        out_sign_id = normalize_text(elem.get("符号出力項目ID", ""))
        keep = bool(
            (not in_id or in_id in active_source_item_ids) and
            (not out_num_id or out_num_id in valid_target_item_ids) and
            (not out_sign_id or out_sign_id in valid_target_item_ids)
        )
        if not keep:
            removed_other_nodes["パックゾーン入力"].add(normalize_text(elem.get("ID", "")))
            root.remove(elem)

    # 10) パックゾーン出力
    for elem in list(root.findall("パックゾーン出力")):
        in_num_id = normalize_text(elem.get("数値入力項目ID", ""))
        in_sign_id = normalize_text(elem.get("符号入力項目ID", ""))
        out_id = normalize_text(elem.get("バイナリ出力項目ID", ""))
        keep = bool(
            (not in_num_id or in_num_id in active_source_item_ids) and
            (not in_sign_id or in_sign_id in active_source_item_ids) and
            (not out_id or out_id in valid_target_item_ids)
        )
        if not keep:
            removed_other_nodes["パックゾーン出力"].add(normalize_text(elem.get("ID", "")))
            root.remove(elem)

    # 11) ロジック本体（子が消えたものを消す）
    removed_logic_ids: Set[str] = set()
    for elem in list(root.findall("ロジック")):
        child_type = normalize_text(elem.get("ChildType", ""))
        child_id = normalize_text(elem.get("ChildID", ""))
        child_tag = CHILD_TYPE_MAP.get(child_type)
        keep = True
        if child_tag:
            keep = child_exists_by_id(root, child_tag, child_id)
        if not keep:
            removed_logic_ids.add(normalize_text(elem.get("ID", "")))
            root.remove(elem)

    return {
        "kept_line_count": len(root.findall("線引き")),
        "kept_record_connection_count": len(root.findall("レコード接続")),
        "kept_logic_count": len(root.findall("ロジック")),
        "removed_line_count": len(removed_line_ids),
        "removed_record_connection_count": len(removed_record_connection_ids),
        "removed_logic_count": len(removed_logic_ids),
        "removed_other_count": sum(len(v) for v in removed_other_nodes.values()),
    }


def save_xml(tree: ET.ElementTree, output_xml: Path) -> None:
    ensure_dir(str(output_xml))
    root = tree.getroot()
    indent_xml(root)
    tree.write(output_xml, encoding="utf-8", xml_declaration=True)


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Order-X2F Step5 根本解決版")
    parser.add_argument("-i", "--input-xml", required=True, help="order_x2f_step1.xml（補助入力）")
    parser.add_argument("-d", "--customer-dictionary", required=True, help="order_bigboss_dictionary.xlsx")
    parser.add_argument("-f", "--from-layout-xml", required=True, help="order_x2f_step3_from_layout.xml")
    parser.add_argument("-g", "--to-layout-xml", required=True, help="order_x2f_step4_to_layout.xml")
    parser.add_argument("-t", "--template-file", required=True, help="ロジックテンプレートXML")
    parser.add_argument("-o", "--output-xml", required=True, help="order_x2f_step5_logic.xml")
    parser.add_argument("-s", "--sheet-name", default=None, help="顧客辞書の対象シート名")
    parser.add_argument("--include-customer-items", action="store_true", help="顧客項目も含めて対象にする")
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    try:
        input_xml = Path(args.input_xml)
        customer_dictionary = Path(args.customer_dictionary)
        from_layout_xml = Path(args.from_layout_xml)
        to_layout_xml = Path(args.to_layout_xml)
        template_file = Path(args.template_file)
        output_xml = Path(args.output_xml)

        for p, label in [
            (input_xml, "INPUT_XML"),
            (customer_dictionary, "DICTIONARY"),
            (from_layout_xml, "FROM_LAYOUT_XML"),
            (to_layout_xml, "TO_LAYOUT_XML"),
            (template_file, "TEMPLATE_FILE"),
        ]:
            if not p.exists():
                raise FileNotFoundError(f"{label} が見つかりません: {p}")

        log("==================================================")
        log("Order-X2F Step5 start")
        log("==================================================")
        log(f"INPUT_XML       = {input_xml}")
        log(f"DICTIONARY      = {customer_dictionary}")
        log(f"FROM_LAYOUT_XML = {from_layout_xml}")
        log(f"TO_LAYOUT_XML   = {to_layout_xml}")
        log(f"TEMPLATE_FILE   = {template_file}")
        log(f"OUTPUT_XML      = {output_xml}")
        log("")

        log("[1/6] 入力XML確認...")
        leaf_count = parse_input_xml_leaf_count(input_xml)
        log(f"  input xml leaf count = {leaf_count}")

        log("[2/6] 顧客辞書読込...")
        dict_df = load_customer_dictionary(customer_dictionary, sheet_name=args.sheet_name)
        work_df = dict_df[dict_df.apply(lambda r: is_step5_target_row(r, include_customer_items=args.include_customer_items), axis=1)].copy()
        work_df = work_df.sort_values(by=["customer_path", "customer_field_id"], kind="stable").reset_index(drop=True)
        active_keys = build_active_dictionary_keys(work_df)
        log(f"  dictionary row count    = {len(dict_df)}")
        log(f"  step5 target row count  = {len(work_df)}")
        log(f"  active path count       = {len(active_keys['paths'])}")
        log(f"  active bms_field_id cnt = {len(active_keys['bms_field_ids'])}")

        log("[3/6] Step3/Step4 レイアウト読込...")
        from_layout_info = parse_layout_xml(from_layout_xml)
        to_layout_info = parse_layout_xml(to_layout_xml)
        log(f"  from layout item count   = {len(from_layout_info['items_by_id'])}")
        log(f"  from layout record count = {len(from_layout_info['records_by_id'])}")
        log(f"  to layout item count     = {len(to_layout_info['items_by_id'])}")
        log(f"  to layout record count   = {len(to_layout_info['records_by_id'])}")

        active_source_item_ids, active_source_record_ids = collect_active_source_ids(from_layout_info, active_keys)
        valid_target_item_ids = set(to_layout_info["items_by_id"].keys())
        valid_target_record_ids = set(to_layout_info["records_by_id"].keys())
        log(f"  active source item ids   = {len(active_source_item_ids)}")
        log(f"  active source record ids = {len(active_source_record_ids)}")

        log("[4/6] ロジックテンプレート読込...")
        logic_tree = ET.parse(template_file)
        logic_root = logic_tree.getroot()
        if local_name(logic_root.tag) != "LogicInfo":
            raise ValueError(f"ロジックテンプレートXMLのルートが想定外です: {logic_root.tag}")
        log(f"  root tag = {local_name(logic_root.tag)}")
        log(f"  template line count              = {len(logic_root.findall('線引き'))}")
        log(f"  template record connection count = {len(logic_root.findall('レコード接続'))}")
        log(f"  template logic count             = {len(logic_root.findall('ロジック'))}")

        log("[5/6] ロジック絞り込み...")
        stats = filter_logic_tree(
            logic_tree=logic_tree,
            active_source_item_ids=active_source_item_ids,
            active_source_record_ids=active_source_record_ids,
            valid_target_item_ids=valid_target_item_ids,
            valid_target_record_ids=valid_target_record_ids,
        )
        log(f"  kept line count              = {stats['kept_line_count']}")
        log(f"  kept record connection count = {stats['kept_record_connection_count']}")
        log(f"  kept logic count             = {stats['kept_logic_count']}")
        log(f"  removed line count           = {stats['removed_line_count']}")
        log(f"  removed record conn count    = {stats['removed_record_connection_count']}")
        log(f"  removed logic count          = {stats['removed_logic_count']}")
        log(f"  removed other node count     = {stats['removed_other_count']}")

        log("[6/6] 保存...")
        save_xml(logic_tree, output_xml)

        log("")
        log("==================================================")
        log("Order-X2F Step5 finished")
        log(f"OUTPUT_XML = {output_xml}")
        log("==================================================")
        return 0

    except Exception as e:
        log("")
        log("[ERROR] Step5 failed")
        log(str(e))
        return 1


if __name__ == "__main__":
    sys.exit(main())
