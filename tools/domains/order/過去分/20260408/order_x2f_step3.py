# -*- coding: utf-8 -*-
"""
order_x2f_step3.py
Step3: 差異表 + Step1 XML を使って変換元レイアウトXMLを生成する

対応内容
- Step1 XMLノード検索による customer_path 実在確認
- 差異表の判定列ゆれ吸収（判定 / judgement）
- 2行見出し対応
- order をレコード(Type=3)として出力
- 中間ノードは項目グループ(Type=4)
- 葉ノードは項目(Type=5)
- 生成path一覧 / 判定一覧のダンプ出力対応
"""

from __future__ import annotations

import argparse
import copy
import os
import sys
from dataclasses import dataclass
from pathlib import Path
from typing import Dict, List, Optional, Tuple
import xml.etree.ElementTree as ET

import pandas as pd


ALLOWED_JUDGEMENTS = {"変換可", "対象外"}
TARGET_JUDGEMENT = "変換可"

TYPE_CODE_MAP = {
    "xsd:string": "0",
    "text": "0",
    "text_max20": "0",
    "text_max25": "0",
    "text_max40": "0",
    "text_max45": "0",
    "text_max60": "0",
    "text_max80": "0",
    "identifier_alnum_max4": "0",
    "identifier_alnum_max10": "0",
    "identifier_num_max13": "2",
    "identifier_num_max14": "2",
    "code_num_2": "2",
    "code_num_3": "2",
    "code_num_5": "2",
    "numeric_4": "2",
    "quantity_6": "2",
    "amount_10": "2",
    "xsd:date": "1",
    "datetype": "1",
    "xsd:datetime": "1",
    "numeric_2_1": "3",
    "quantity_6_1": "3",
}


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
    s = str(value)
    s = s.replace("\u3000", " ")  # 全角空白→半角
    return s.strip()


def local_name(tag: str) -> str:
    if not tag:
        return ""
    if "}" in tag:
        return tag.split("}", 1)[1]
    if ":" in tag:
        return tag.split(":", 1)[1]
    return tag


def split_path(path_text: str) -> List[str]:
    return [p for p in normalize_text(path_text).split("/") if p]


def ensure_dir(path_text: str) -> None:
    parent = os.path.dirname(path_text)
    if parent:
        os.makedirs(parent, exist_ok=True)


def safe_int_text(value: object, default: str = "0") -> str:
    s = normalize_text(value)
    if not s:
        return default
    try:
        return str(int(float(s)))
    except Exception:
        return default


def indent_xml(elem: ET.Element, level: int = 0) -> None:
    i = "\n" + level * "  "
    if len(elem):
        if not elem.text or not elem.text.strip():
            elem.text = i + "  "
        last = None
        for child in elem:
            indent_xml(child, level + 1)
            last = child
        if last is not None and (not last.tail or not last.tail.strip()):
            last.tail = i
    if level and (not elem.tail or not elem.tail.strip()):
        elem.tail = i


def derive_type_code(data_type: str) -> str:
    return TYPE_CODE_MAP.get(normalize_text(data_type).lower(), "0")


def derive_decimal_places(data_type: str) -> str:
    return "1" if normalize_text(data_type).lower() in {"numeric_2_1", "quantity_6_1"} else "0"


def derive_required(value: str) -> str:
    return "1" if normalize_text(value).upper() in {"Y", "YES", "1", "TRUE", "必須"} else "0"


def find_column(df: pd.DataFrame, candidates: List[str]) -> Optional[str]:
    normalized = {normalize_text(c).lower(): c for c in df.columns}
    for cand in candidates:
        key = normalize_text(cand).lower()
        if key in normalized:
            return normalized[key]
    return None


def load_step2_diff(xlsx_path: Path, sheet_name: Optional[str] = None) -> pd.DataFrame:
    if not xlsx_path.exists():
        raise FileNotFoundError(f"差異表が存在しません: {xlsx_path}")

    read_sheet = sheet_name if sheet_name else 0
    df = pd.read_excel(xlsx_path, sheet_name=read_sheet, header=0, dtype=str)
    if isinstance(df, dict):
        if not df:
            raise RuntimeError(f"差異表にシートがありません: {xlsx_path}")
        df = next(iter(df.values()))
    df = df.fillna("")

    # 2行見出し対応
    if len(df) > 0:
        first_row_values = [normalize_text(v) for v in df.iloc[0].tolist()]
        if any(v in ("判定", "judgement", "顧客タグ名", "customer_tag", "customer_path") for v in first_row_values):
            df = df.iloc[1:].reset_index(drop=True)

    # 列名ゆれ吸収
    colmap = {}

    judgement_col = find_column(df, ["judgement", "判定"])
    customer_field_id_col = find_column(df, ["customer_field_id", "顧客項目ID"])
    customer_tag_col = find_column(df, ["customer_tag", "顧客タグ名"])
    candidate_bms_field_id_col = find_column(df, ["candidate_bms_field_id", "候補BMS項目ID"])
    candidate_bms_tag_col = find_column(df, ["candidate_bms_tag", "候補BMSタグ"])
    confirmed_bms_field_id_col = find_column(df, ["confirmed_bms_field_id", "確定BMS項目ID"])
    data_type_col = find_column(df, ["data_type", "データ型"])
    length_max_col = find_column(df, ["length_max", "最大桁数", "桁数"])
    required_flag_col = find_column(df, ["required_flag", "必須"])
    repeat_flag_col = find_column(df, ["repeat_flag", "繰返し"])
    customer_path_col = find_column(df, ["customer_path", "顧客パス"])
    sample_value_col = find_column(df, ["sample_value", "サンプル値"])
    remarks_col = find_column(df, ["remarks", "備考"])

    colmap["judgement"] = judgement_col
    colmap["customer_field_id"] = customer_field_id_col
    colmap["customer_tag"] = customer_tag_col
    colmap["candidate_bms_field_id"] = candidate_bms_field_id_col
    colmap["candidate_bms_tag"] = candidate_bms_tag_col
    colmap["confirmed_bms_field_id"] = confirmed_bms_field_id_col
    colmap["data_type"] = data_type_col
    colmap["length_max"] = length_max_col
    colmap["required_flag"] = required_flag_col
    colmap["repeat_flag"] = repeat_flag_col
    colmap["customer_path"] = customer_path_col
    colmap["sample_value"] = sample_value_col
    colmap["remarks"] = remarks_col

    out = pd.DataFrame()
    for key, src in colmap.items():
        out[key] = df[src] if src else ""

    out = out.fillna("")

    # 判定の正規化
    out["judgement"] = out["judgement"].map(normalize_text)
    out["customer_path"] = out["customer_path"].map(normalize_text)
    out["customer_tag"] = out["customer_tag"].map(normalize_text)

    return out


def validate_diff(df: pd.DataFrame) -> None:
    errors: List[str] = []
    for i, row in df.iterrows():
        excel_row = i + 3
        path = normalize_text(row.get("customer_path", ""))
        tag = normalize_text(row.get("customer_tag", ""))
        judgement = normalize_text(row.get("judgement", ""))

        if not path and not tag:
            continue

        if judgement not in ALLOWED_JUDGEMENTS:
            errors.append(
                f"row={excel_row} judgement='{judgement}' path='{path}' tag='{tag}'"
            )

    if errors:
        raise RuntimeError(
            "差異表の判定に未確定値があります。"
            f" Step3に進めるのは {sorted(ALLOWED_JUDGEMENTS)} のみです。\n"
            + "\n".join(errors[:50])
        )


@dataclass
class XmlContext:
    root: ET.Element
    parent_map: Dict[int, ET.Element]
    path_to_elements: Dict[str, List[ET.Element]]


def build_xml_context(step1_xml: Path) -> XmlContext:
    if not step1_xml.exists():
        raise FileNotFoundError(f"Step1 XML が存在しません: {step1_xml}")

    root = ET.parse(step1_xml).getroot()
    parent_map: Dict[int, ET.Element] = {}
    path_to_elements: Dict[str, List[ET.Element]] = {}

    def walk(elem: ET.Element, ancestors: List[str]) -> None:
        current_name = local_name(elem.tag)
        current_path = "/" + "/".join(ancestors + [current_name])
        path_to_elements.setdefault(current_path, []).append(elem)

        for child in list(elem):
            parent_map[id(child)] = elem
            walk(child, ancestors + [current_name])

    walk(root, [])
    return XmlContext(root=root, parent_map=parent_map, path_to_elements=path_to_elements)


def get_actual_ancestor_paths(elem: ET.Element, ctx: XmlContext) -> List[str]:
    names: List[str] = []
    current: Optional[ET.Element] = elem
    while current is not None:
        names.append(local_name(current.tag))
        current = ctx.parent_map.get(id(current))
    names.reverse()

    result: List[str] = []
    current_parts: List[str] = []
    for name in names:
        current_parts.append(name)
        result.append("/" + "/".join(current_parts))
    return result


@dataclass
class RowDecision:
    row_index: int
    customer_path: str
    verdict: str
    reason: str


def decide_target_rows(df: pd.DataFrame, ctx: XmlContext) -> Tuple[pd.DataFrame, List[RowDecision], List[str]]:
    decisions: List[RowDecision] = []
    errors: List[str] = []
    selected_indices: List[int] = []

    for i, row in df.iterrows():
        excel_row = i + 3
        judgement = normalize_text(row.get("judgement", ""))
        customer_path = normalize_text(row.get("customer_path", ""))
        customer_tag = normalize_text(row.get("customer_tag", ""))

        if not customer_path and not customer_tag:
            continue

        if judgement == "対象外":
            decisions.append(RowDecision(excel_row, customer_path, "SKIP", "judgement=対象外"))
            continue

        if judgement != TARGET_JUDGEMENT:
            errors.append(f"row={excel_row} judgement='{judgement}' path='{customer_path}'")
            continue

        elems = ctx.path_to_elements.get(customer_path, [])
        if not elems:
            errors.append(
                f"row={excel_row} customer_path が Step1 XML に存在しません: {customer_path}"
            )
            continue

        selected_indices.append(i)
        decisions.append(RowDecision(excel_row, customer_path, "TAKE", f"xml-hit={len(elems)}"))

    selected_df = df.loc[selected_indices].copy().reset_index(drop=True)
    return selected_df, decisions, errors


@dataclass
class Node:
    path: str
    name: str
    parent_path: str
    is_leaf: bool
    row: Optional[dict]


def build_tree_from_selected_rows(rows: List[dict], ctx: XmlContext) -> Tuple[str, Dict[str, Node], List[str]]:
    if not rows:
        raise RuntimeError("Step3採用件数が0件です。")

    node_map: Dict[str, Node] = {}
    generated_paths: List[str] = []

    root_name = local_name(ctx.root.tag)
    root_path = "/" + root_name
    node_map[root_path] = Node(root_path, root_name, "", False, None)
    generated_paths.append(root_path)

    for row in rows:
        customer_path = normalize_text(row["customer_path"])
        elems = ctx.path_to_elements.get(customer_path, [])
        if not elems:
            raise RuntimeError(f"Step1 XML に存在しない customer_path が混入しました: {customer_path}")

        actual_paths = get_actual_ancestor_paths(elems[0], ctx)

        for idx, path in enumerate(actual_paths):
            if path in node_map:
                continue
            parts = split_path(path)
            name = parts[-1]
            parent_path = actual_paths[idx - 1] if idx > 0 else ""
            is_leaf = (idx == len(actual_paths) - 1)
            node_map[path] = Node(path, name, parent_path, is_leaf, None)
            generated_paths.append(path)

        leaf_path = actual_paths[-1]
        node_map[leaf_path].row = row
        node_map[leaf_path].is_leaf = True

    return root_name, node_map, sorted(generated_paths, key=lambda p: (len(split_path(p)), p))


def build_output_root(template_path: Path, root_tag_name: str) -> ET.Element:
    tmpl_root = ET.parse(template_path).getroot()
    if local_name(tmpl_root.tag) != "マッピングレイアウト":
        raise RuntimeError("テンプレートXMLのルートが マッピングレイアウト ではありません。")

    out_root = ET.Element("マッピングレイアウト", tmpl_root.attrib)

    for child in tmpl_root:
        if local_name(child.tag) == "schema":
            out_root.append(copy.deepcopy(child))
            break

    out_root.append(ET.Element("レイアウト", {
        "ID": "1", "Name": "データストア", "Type": "1", "PropertyID": "1", "ParentsID": "0"
    }))
    out_root.append(ET.Element("データストア", {
        "ID": "1", "フォーマット": "6"
    }))
    out_root.append(ET.Element("データストアXML", {
        "ID": "1",
        "ルートタグ名": root_tag_name,
        "XMLスキーマ指定": "0",
        "XMLスキーマファイルパス": "",
        "桁あふれ": "1",
        "名前空間指定": "0",
        "名前空間定義": "",
        "文字コード": "12",
        "改行処理": "0",
        "インデント処理": "1",
        "インデント文字コード": "",
        "SchemaLocation": "",
    }))
    return out_root


def add_elem(root: ET.Element, tag: str, attrs: Dict[str, str]) -> None:
    elem = ET.Element(tag)
    for k, v in attrs.items():
        elem.set(k, normalize_text(v))
    root.append(elem)


def write_lines(path: Optional[Path], lines: List[str]) -> None:
    if path is None:
        return
    ensure_dir(str(path))
    path.write_text("\n".join(lines) + "\n", encoding="utf-8")


def generate_layout(
    step1_xml: Path,
    step2_diff_xlsx: Path,
    template_file: Path,
    output_xml: Path,
    diff_sheet_name: Optional[str],
    path_dump_file: Optional[Path],
    decision_dump_file: Optional[Path],
) -> None:
    log("[INFO] Step1 XML 読み込み開始")
    ctx = build_xml_context(step1_xml)
    log(f"[INFO] Step1 XML path件数={len(ctx.path_to_elements)}")

    log("[INFO] 差異表読み込み開始")
    diff_df = load_step2_diff(step2_diff_xlsx, diff_sheet_name)
    validate_diff(diff_df)

    selected_df, decisions, decision_errors = decide_target_rows(diff_df, ctx)
    if decision_errors:
        raise RuntimeError("新判定ロジックでエラーが発生しました。\n" + "\n".join(decision_errors[:50]))

    log(f"[INFO] Step3採用件数={len(selected_df)}")
    rows = [r.to_dict() for _, r in selected_df.iterrows()]

    root_name, node_map, generated_paths = build_tree_from_selected_rows(rows, ctx)
    out_root = build_output_root(template_file, root_name)

    next_layout_id = 2
    next_group_id = 1
    next_record_id = 1
    next_item_id = 1

    path_to_layout_id: Dict[str, str] = {f"/{root_name}": "1"}
    ordered_paths = sorted(node_map.keys(), key=lambda p: (len(split_path(p)), p))

    for path in ordered_paths:
        node = node_map[path]
        if path == f"/{root_name}":
            continue

        parent_layout_id = path_to_layout_id[node.parent_path]

        # order をレコード化
        if node.name == "order" and not node.is_leaf:
            layout_id = str(next_layout_id)
            record_id = str(next_record_id)
            next_layout_id += 1
            next_record_id += 1

            parent_record_name = split_path(node.parent_path)[-1] if node.parent_path else ""

            add_elem(out_root, "レイアウト", {
                "ID": layout_id,
                "Name": node.name,
                "Type": "3",
                "PropertyID": record_id,
                "ParentsID": parent_layout_id,
            })
            add_elem(out_root, "レコード", {
                "ID": record_id,
                "レコード名": node.name,
            })
            add_elem(out_root, "レコードXML", {
                "ID": record_id,
                "階層番号": "1",
                "必須": "1",
                "親レコード名": parent_record_name,
            })
            add_elem(out_root, "レコード補助情報", {
                "ID": record_id,
                "レコード出力タイプ": "0",
                "項目ID配列": "",
                "評価方法": "0",
                "評価順序": "0",
                "with空削除": "False",
            })
            path_to_layout_id[path] = layout_id

        elif not node.is_leaf:
            layout_id = str(next_layout_id)
            group_id = str(next_group_id)
            next_layout_id += 1
            next_group_id += 1

            add_elem(out_root, "レイアウト", {
                "ID": layout_id,
                "Name": node.name,
                "Type": "4",
                "PropertyID": group_id,
                "ParentsID": parent_layout_id,
            })
            add_elem(out_root, "項目グループ", {
                "ID": group_id,
                "項目グループ名": node.name,
                "項目グループの繰り返し回数": "1",
                "項目グループ種別": "0",
                "必須": "1",
                "タグの出現": "0",
            })
            path_to_layout_id[path] = layout_id

        else:
            row = node.row or {}
            layout_id = str(next_layout_id)
            item_id = str(next_item_id)
            next_layout_id += 1
            next_item_id += 1

            bms_field_id = (
                normalize_text(row.get("confirmed_bms_field_id", ""))
                or normalize_text(row.get("candidate_bms_field_id", ""))
                or normalize_text(row.get("customer_field_id", ""))
                or f"auto_{item_id}"
            )

            data_type = normalize_text(row.get("data_type", ""))
            type_code = derive_type_code(data_type)
            decimal_places = derive_decimal_places(data_type)
            cal_format = ""
            if normalize_text(data_type).lower() == "xsd:datetime":
                cal_format = "yyyy-MM-ddTHH:mm:ss"
            elif type_code == "1":
                cal_format = "yyyy-MM-dd"

            add_elem(out_root, "レイアウト", {
                "ID": layout_id,
                "Name": node.name,
                "Type": "5",
                "PropertyID": item_id,
                "ParentsID": parent_layout_id,
            })
            add_elem(out_root, "項目", {
                "ID": item_id,
                "項目名": node.name,
                "項目ID": bms_field_id,
                "属性": "1",
                "属性チェック": "0",
            })
            add_elem(out_root, "項目XML", {
                "ID": item_id,
                "必須": derive_required(row.get("required_flag", "")),
                "データ型": type_code,
                "全体の桁数": safe_int_text(row.get("length_max", ""), "0"),
                "小数部桁数": decimal_places,
                "パディングを行う": "False",
                "パディング": "2",
                "位置": "1",
            })
            add_elem(out_root, "項目補助情報", {
                "ID": item_id,
                "全角": "False",
                "大文字": "False",
                "暦書式": cal_format,
                "暦エラータイプ": "0",
                "レコード出力タイプ": "0",
                "最終出力タイプ": "0",
                "最終出力値": "",
            })
            path_to_layout_id[path] = layout_id

    indent_xml(out_root)
    ensure_dir(str(output_xml))
    ET.ElementTree(out_root).write(output_xml, encoding="utf-8", xml_declaration=True)

    write_lines(path_dump_file, generated_paths)
    decision_lines = [
        f"row={d.row_index}\tverdict={d.verdict}\treason={d.reason}\tpath={d.customer_path}"
        for d in decisions
    ]
    write_lines(decision_dump_file, decision_lines)

    log(f"[INFO] 生成path件数={len(generated_paths)}")
    if path_dump_file:
        log(f"[INFO] 生成path一覧出力={path_dump_file}")
    if decision_dump_file:
        log(f"[INFO] 判定結果一覧出力={decision_dump_file}")
    log(f"[INFO] 出力完了: {output_xml}")


def parse_args() -> argparse.Namespace:
    p = argparse.ArgumentParser(description="Step3 新判定ロジック版")
    p.add_argument("-i", "--step1-xml", required=True, help="order_x2f_step1.xml")
    p.add_argument("-d", "--step2-diff", required=True, help="order_x2f_step2.xlsx / order_x2f_step2_diff_report.xlsx")
    p.add_argument("-t", "--template-file", required=True, help="基本形1_3：発注Ver1_3.xml")
    p.add_argument("-o", "--output-xml", required=True, help="出力する変換元レイアウトXML")
    p.add_argument("-s", "--sheet-name", default=None, help="差異表の対象シート名")
    p.add_argument("-p", "--path-dump", default=None, help="生成path一覧出力txt")
    p.add_argument("-r", "--decision-dump", default=None, help="判定結果一覧出力txt")
    return p.parse_args()


def main() -> int:
    args = parse_args()
    try:
        generate_layout(
            step1_xml=Path(args.step1_xml),
            step2_diff_xlsx=Path(args.step2_diff),
            template_file=Path(args.template_file),
            output_xml=Path(args.output_xml),
            diff_sheet_name=args.sheet_name,
            path_dump_file=Path(args.path_dump) if args.path_dump else None,
            decision_dump_file=Path(args.decision_dump) if args.decision_dump else None,
        )
        return 0
    except Exception as e:
        log(f"[ERROR] {e}")
        return 1


if __name__ == "__main__":
    sys.exit(main())