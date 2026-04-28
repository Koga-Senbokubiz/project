
#!/usr/bin/env python
# -*- coding: utf-8 -*-

"""
order_x2dfax_step3.py

目的:
    EasyExchange 用 Logic XML を生成する（Order XML -> DFAX 固定長）。

今回修正版の要点:
    1. Logic XML の schema を EasyExchange 既存形式に合わせたフル定義で出力する
    2. 線引きの target 側は「レイアウトID」ではなく「項目ID」を使う
    3. 変換先が固定値出力項目（項目補助情報.最終出力タイプ=1）の場合は線を引かない
    4. レコード接続は実データ・実レイアウトに合わせて安定生成する
"""

from __future__ import annotations

import argparse
import sys
from pathlib import Path
from typing import Dict, List, Tuple
import xml.etree.ElementTree as ET


# =========================================================
# 変換ルール
# =========================================================

# 変換元レコード名 -> 変換先レコード名
RECORD_RULES: List[Tuple[str, str]] = [
    ("common:message", "D1001"),
    ("order", "D1004"),
    ("order", "D1006"),
    ("lineItem", "D1009"),
    ("lineItem", "D1011"),
]

# 変換元 項目ID（業務名） -> 変換先 項目名
FIELD_RULES: List[Tuple[str, str]] = [
    # D1001
    ("送信者ステーションアドレス", "D1001_002_fax_no"),
    ("取引先コード", "D1001_002_seller_code"),
    ("取引先名称", "D1001_002_seller_name"),

    # D1004
    ("取引番号（発注・返品）", "D1004_006_trade_number"),
    ("発注日", "D1004_012_order_date_yy_mm_dd"),
    ("最終納品先納品日", "D1004_014_delivery_date_yy_mm_dd"),
    ("発注者名称", "D1004_003_buyer_name"),

    # D1006
    ("最終納品先名称", "D1006_003_receiver_name"),
    ("最終納品先コード", "D1006_005_receiver_code"),
    ("商品分類（中）", "D1006_006_sub_major_category"),
    ("処理種別", "D1006_007_trade_type_code"),
    ("取引先コード", "D1006_008_seller_code"),
    ("取引先名称", "D1006_010_seller_name"),

    # D1009
    ("商品名", "D1009_002_item_name"),

    # D1011
    ("商品規格：規格", "D1011_004_spec"),
    ("商品コード（発注用）", "D1011_005_order_item_code"),
    ("発注単位", "D1011_007_unit_multiple"),
    ("発注数量（発注単位数）", "D1011_008_num_order_units"),
    ("発注単位コード", "D1011_009_unit_of_measure_display"),
    ("発注数量（バラ）", "D1011_010_quantity"),
    ("原単価", "D1011_012_unit_price"),
    ("原価金額", "D1011_014_net_amount"),
]


# =========================================================
# 汎用
# =========================================================
def eprint(*args, **kwargs) -> None:
    print(*args, file=sys.stderr, **kwargs)


def detect_encoding_and_read(path: Path) -> Tuple[str, str]:
    raw = path.read_bytes()
    for enc in ("utf-8-sig", "utf-8", "cp932", "shift_jis"):
        try:
            return raw.decode(enc), enc
        except UnicodeDecodeError:
            continue
    raise ValueError(f"文字コードを判定できませんでした: {path}")


def write_text(path: Path, text: str, encoding: str = "utf-8") -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(text, encoding=encoding, newline="")


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Order-X2DFAX Step3: EasyExchange Logic XML を生成する"
    )
    parser.add_argument("-f", "--from-layout", required=True, help="変換元レイアウトXML")
    parser.add_argument("-t", "--to-layout", required=True, help="変換先レイアウトXML")
    parser.add_argument("-o", "--output", required=True, help="出力Logic XML")
    return parser.parse_args()


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
    if level and (not elem.tail or not elem.tail.strip()):
        elem.tail = i


# =========================================================
# レイアウト解析
# =========================================================
class SourceLayout:
    def __init__(self, xml_path: Path):
        text, self.encoding = detect_encoding_and_read(xml_path)
        self.root = ET.fromstring(text)

    def get_records(self) -> Dict[str, str]:
        result: Dict[str, str] = {}
        for elem in self.root.findall(".//レコード"):
            rec_id = elem.get("ID")
            rec_name = elem.get("レコード名")
            if rec_id and rec_name:
                result[rec_name] = rec_id
        return result

    def get_items_by_business_id(self) -> Dict[str, str]:
        """
        key = 項目ID（業務名）
        value = 項目.ID
        """
        result: Dict[str, str] = {}
        for elem in self.root.findall(".//項目"):
            item_id = elem.get("ID")
            business_id = elem.get("項目ID")
            if item_id and business_id:
                result[business_id] = item_id
        return result


class TargetLayout:
    def __init__(self, xml_path: Path):
        text, self.encoding = detect_encoding_and_read(xml_path)
        self.root = ET.fromstring(text)

    def get_records(self) -> Dict[str, str]:
        result: Dict[str, str] = {}
        for elem in self.root.findall(".//レコード"):
            rec_id = elem.get("ID")
            rec_name = elem.get("レコード名")
            if rec_id and rec_name:
                result[rec_name] = rec_id
        return result

    def get_items_by_name(self) -> Dict[str, str]:
        """
        key = 項目名
        value = 項目.ID
        """
        result: Dict[str, str] = {}
        for elem in self.root.findall(".//項目"):
            item_id = elem.get("ID")
            item_name = elem.get("項目名")
            if item_id and item_name:
                result[item_name] = item_id
        return result

    def get_fixed_item_names(self) -> set[str]:
        """
        項目補助情報.最終出力タイプ = 1 の項目名を返す。
        これらは固定値出力項目なので、Step3 の線引き対象から除外する。
        """
        item_name_by_id: Dict[str, str] = {}
        for elem in self.root.findall(".//項目"):
            item_id = elem.get("ID")
            item_name = elem.get("項目名")
            if item_id and item_name:
                item_name_by_id[item_id] = item_name

        fixed_names: set[str] = set()
        for elem in self.root.findall(".//項目補助情報"):
            item_id = elem.get("ID")
            out_type = elem.get("最終出力タイプ")
            if item_id and out_type == "1":
                item_name = item_name_by_id.get(item_id)
                if item_name:
                    fixed_names.add(item_name)
        return fixed_names


# =========================================================
# Logic schema（EasyExchange用フル版）
# =========================================================
def build_logic_root_with_full_schema() -> ET.Element:
    root = ET.Element(
        "LogicInfo",
        {
            "xmlns:ns1": "urn:schemas-microsoft-com:xml-msdata",
            "xmlns:xs": "http://www.w3.org/2001/XMLSchema",
        },
    )

    schema = ET.SubElement(root, "xs:schema", {"id": "LogicInfo"})
    xs_element = ET.SubElement(
        schema,
        "xs:element",
        {
            "name": "LogicInfo",
            "ns1:IsDataSet": "true",
            "ns1:UseCurrentLocale": "true",
        },
    )
    complex_type = ET.SubElement(xs_element, "xs:complexType")
    choice = ET.SubElement(
        complex_type,
        "xs:choice",
        {"minOccurs": "0", "maxOccurs": "unbounded"},
    )

    def add_element(name: str, attrs: List[Tuple[str, Dict[str, str]]]) -> None:
        elem = ET.SubElement(choice, "xs:element", {"name": name})
        ct = ET.SubElement(elem, "xs:complexType")
        for attr_name, attr_def in attrs:
            d = {"name": attr_name}
            d.update(attr_def)
            ET.SubElement(ct, "xs:attribute", d)

    add_element("ロジック", [
        ("ID", {"ns1:AutoIncrement": "true", "ns1:AutoIncrementSeed": "1", "type": "xs:int", "use": "required"}),
        ("ChildType", {"type": "xs:string"}),
        ("ChildID", {"type": "xs:string"}),
        ("表示インデックス", {"type": "xs:string"}),
    ])
    add_element("レコード接続", [
        ("ID", {"ns1:AutoIncrement": "true", "ns1:AutoIncrementSeed": "1", "type": "xs:int", "use": "required"}),
        ("変換元レコードID", {"type": "xs:string"}),
        ("変換先レコードID", {"type": "xs:string"}),
    ])
    add_element("線引き", [
        ("ID", {"ns1:AutoIncrement": "true", "ns1:AutoIncrementSeed": "1", "type": "xs:int", "use": "required"}),
        ("変換元項目ID", {"type": "xs:string"}),
        ("変換先項目ID", {"type": "xs:string"}),
    ])
    add_element("ロジックパラメータ", [
        ("ID", {"ns1:AutoIncrement": "true", "ns1:AutoIncrementSeed": "1", "type": "xs:int", "use": "required"}),
        ("ParentID", {"type": "xs:string"}),
        ("入出力", {"type": "xs:string"}),
        ("引数名", {"type": "xs:string"}),
        ("引数値", {"type": "xs:string"}),
        ("項目ID", {"type": "xs:string"}),
        ("初期値", {"type": "xs:string"}),
    ])
    add_element("CSVテーブル検索", [
        ("ID", {"ns1:AutoIncrement": "true", "ns1:AutoIncrementSeed": "1", "type": "xs:int", "use": "required"}),
        ("ParentID", {"type": "xs:string"}),
        ("CSVファイル", {"type": "xs:string"}),
        ("中断する", {"type": "xs:string"}),
    ])
    add_element("DBテーブル検索", [
        ("ID", {"ns1:AutoIncrement": "true", "ns1:AutoIncrementSeed": "1", "type": "xs:int", "use": "required"}),
        ("ParentID", {"type": "xs:string"}),
        ("接続文字列", {"type": "xs:string"}),
        ("テーブル", {"type": "xs:string"}),
        ("中断する", {"type": "xs:string"}),
    ])
    add_element("COMメソッド呼び出し", [
        ("ID", {"ns1:AutoIncrement": "true", "ns1:AutoIncrementSeed": "1", "type": "xs:int", "use": "required"}),
        ("ParentID", {"type": "xs:string"}),
        ("モジュール", {"type": "xs:string"}),
        ("プログラムID", {"type": "xs:string"}),
        ("メソッド", {"type": "xs:string"}),
        ("引数", {"type": "xs:string"}),
    ])
    add_element("VBScript呼び出し", [
        ("ID", {"ns1:AutoIncrement": "true", "ns1:AutoIncrementSeed": "1", "type": "xs:int", "use": "required"}),
        ("ParentID", {"type": "xs:string"}),
        ("スクリプト", {"type": "xs:string"}),
        ("関数", {"type": "xs:string"}),
        ("引数", {"type": "xs:string"}),
    ])
    add_element("パックゾーン入力", [
        ("ID", {"ns1:AutoIncrement": "true", "ns1:AutoIncrementSeed": "1", "type": "xs:int", "use": "required"}),
        ("バイナリタイプ", {"type": "xs:string"}),
        ("符号付加", {"type": "xs:string"}),
        ("正符号省略", {"type": "xs:string"}),
        ("符号位置", {"type": "xs:string"}),
        ("バイナリ入力項目ID", {"type": "xs:string"}),
        ("数値出力項目ID", {"type": "xs:string"}),
        ("符号出力項目ID", {"type": "xs:string"}),
        ("小数部桁数", {"type": "xs:string"}),
    ])
    add_element("パックゾーン出力", [
        ("ID", {"ns1:AutoIncrement": "true", "ns1:AutoIncrementSeed": "1", "type": "xs:int", "use": "required"}),
        ("バイナリタイプ", {"type": "xs:string"}),
        ("符号入力", {"type": "xs:string"}),
        ("正符号省略", {"type": "xs:string"}),
        ("数値入力項目ID", {"type": "xs:string"}),
        ("符号入力項目ID", {"type": "xs:string"}),
        ("バイナリ出力項目ID", {"type": "xs:string"}),
        ("小数部桁数", {"type": "xs:string"}),
    ])
    add_element("結合", [
        ("ID", {"ns1:AutoIncrement": "true", "ns1:AutoIncrementSeed": "1", "type": "xs:int", "use": "required"}),
        ("変換元項目ID配列", {"type": "xs:string"}),
        ("変換先項目ID", {"type": "xs:string"}),
    ])
    add_element("文字列操作", [
        ("ID", {"ns1:AutoIncrement": "true", "ns1:AutoIncrementSeed": "1", "type": "xs:int", "use": "required"}),
        ("ParentID", {"type": "xs:string"}),
        ("変換元項目ID", {"type": "xs:string"}),
        ("変換先項目ID", {"type": "xs:string"}),
        ("テスト文字列", {"type": "xs:string"}),
    ])
    add_element("文字列操作コマンド", [
        ("ID", {"ns1:AutoIncrement": "true", "ns1:AutoIncrementSeed": "1", "type": "xs:int", "use": "required"}),
        ("ParentID", {"type": "xs:string"}),
        ("コマンド", {"type": "xs:string"}),
        ("パラメータ1", {"type": "xs:string"}),
        ("パラメータ2", {"type": "xs:string"}),
        ("パラメータ3", {"type": "xs:string"}),
        ("パラメータ4", {"type": "xs:string"}),
        ("パラメータ5", {"type": "xs:string"}),
    ])
    add_element("連番", [
        ("ID", {"ns1:AutoIncrement": "true", "ns1:AutoIncrementSeed": "1", "type": "xs:int", "use": "required"}),
        ("変換先レコードID", {"type": "xs:string"}),
        ("変換先項目ID配列", {"type": "xs:string"}),
        ("初期値", {"type": "xs:string"}),
        ("増加量", {"type": "xs:string"}),
        ("レコード整形前", {"type": "xs:string"}),
    ])
    add_element("件数", [
        ("ID", {"ns1:AutoIncrement": "true", "ns1:AutoIncrementSeed": "1", "type": "xs:int", "use": "required"}),
        ("変換先レコードID", {"type": "xs:string"}),
        ("変換先項目ID", {"type": "xs:string"}),
    ])
    add_element("合計", [
        ("ID", {"ns1:AutoIncrement": "true", "ns1:AutoIncrementSeed": "1", "type": "xs:int", "use": "required"}),
        ("変換元項目ID", {"type": "xs:string"}),
        ("変換先項目ID", {"type": "xs:string"}),
        ("基準レコードID", {"type": "xs:string"}),
    ])
    add_element("CIIテーブル検索", [
        ("ID", {"ns1:AutoIncrement": "true", "ns1:AutoIncrementSeed": "1", "type": "xs:int", "use": "required"}),
        ("ParentID", {"type": "xs:string"}),
    ])
    add_element("固定値", [
        ("ID", {"ns1:AutoIncrement": "true", "ns1:AutoIncrementSeed": "1", "type": "xs:int", "use": "required"}),
        ("固定入力値", {"type": "xs:string"}),
        ("変換先項目ID", {"type": "xs:string"}),
    ])

    def add_unique(name: str, selector: str, with_constraint_name: bool = True) -> None:
        attrs = {"name": name, "ns1:PrimaryKey": "true"}
        if with_constraint_name:
            attrs["ns1:ConstraintName"] = "Constraint1"
        uniq = ET.SubElement(xs_element, "xs:unique", attrs)
        ET.SubElement(uniq, "xs:selector", {"xpath": selector})
        ET.SubElement(uniq, "xs:field", {"xpath": "@ID"})

    add_unique("Constraint1", ".//ロジック", with_constraint_name=False)
    add_unique("レコード接続_Constraint1", ".//レコード接続")
    add_unique("線引き_Constraint1", ".//線引き")
    add_unique("ロジックパラメータ_Constraint1", ".//ロジックパラメータ")
    add_unique("CSVテーブル検索_Constraint1", ".//CSVテーブル検索")
    add_unique("DBテーブル検索_Constraint1", ".//DBテーブル検索")
    add_unique("COMメソッド呼び出し_Constraint1", ".//COMメソッド呼び出し")
    add_unique("VBScript呼び出し_Constraint1", ".//VBScript呼び出し")
    add_unique("パックゾーン入力_Constraint1", ".//パックゾーン入力")
    add_unique("パックゾーン出力_Constraint1", ".//パックゾーン出力")
    add_unique("結合_Constraint1", ".//結合")
    add_unique("文字列操作_Constraint1", ".//文字列操作")
    add_unique("文字列操作コマンド_Constraint1", ".//文字列操作コマンド")
    add_unique("連番_Constraint1", ".//連番")
    add_unique("件数_Constraint1", ".//件数")
    add_unique("合計_Constraint1", ".//合計")
    add_unique("CIIテーブル検索_Constraint1", ".//CIIテーブル検索")
    add_unique("固定値_Constraint1", ".//固定値")

    return root


# =========================================================
# 主処理
# =========================================================
def main() -> int:
    args = parse_args()

    from_path = Path(args.from_layout)
    to_path = Path(args.to_layout)
    out_path = Path(args.output)

    if not from_path.exists():
        eprint(f"[ERROR] 変換元レイアウト XML が存在しません: {from_path}")
        return 1
    if not to_path.exists():
        eprint(f"[ERROR] 変換先レイアウト XML が存在しません: {to_path}")
        return 1

    try:
        src = SourceLayout(from_path)
        tgt = TargetLayout(to_path)
    except Exception as ex:
        eprint(f"[ERROR] レイアウト XML の解析に失敗しました: {ex}")
        return 1

    src_records = src.get_records()
    src_items = src.get_items_by_business_id()
    tgt_records = tgt.get_records()
    tgt_items = tgt.get_items_by_name()
    tgt_fixed_items = tgt.get_fixed_item_names()

    root = build_logic_root_with_full_schema()

    record_conn_id = 1
    line_id = 1
    logic_id = 1

    skipped_records: List[str] = []
    skipped_fields: List[str] = []
    skipped_fixed_targets: List[str] = []

    # レコード接続
    for src_name, tgt_name in RECORD_RULES:
        src_id = src_records.get(src_name)
        tgt_id = tgt_records.get(tgt_name)

        if not src_id or not tgt_id:
            skipped_records.append(f"{src_name} -> {tgt_name}")
            continue

        ET.SubElement(root, "レコード接続", {
            "ID": str(record_conn_id),
            "変換元レコードID": str(src_id),
            "変換先レコードID": str(tgt_id),
        })
        ET.SubElement(root, "ロジック", {
            "ID": str(logic_id),
            "ChildType": "1",
            "ChildID": str(record_conn_id),
            "表示インデックス": "0",
        })
        record_conn_id += 1
        logic_id += 1

    # 線引き（固定値出力項目は除外）
    for src_business_id, tgt_item_name in FIELD_RULES:
        if tgt_item_name in tgt_fixed_items:
            skipped_fixed_targets.append(f"{src_business_id} -> {tgt_item_name}")
            continue

        src_item_id = src_items.get(src_business_id)
        tgt_item_id = tgt_items.get(tgt_item_name)

        if not src_item_id or not tgt_item_id:
            skipped_fields.append(f"{src_business_id} -> {tgt_item_name}")
            continue

        ET.SubElement(root, "線引き", {
            "ID": str(line_id),
            "変換元項目ID": str(src_item_id),
            "変換先項目ID": str(tgt_item_id),
        })
        ET.SubElement(root, "ロジック", {
            "ID": str(logic_id),
            "ChildType": "2",
            "ChildID": str(line_id),
            "表示インデックス": "0",
        })
        line_id += 1
        logic_id += 1

    indent_xml(root)
    xml_text = ET.tostring(root, encoding="unicode")
    xml_text = "<?xml version='1.0' encoding='utf-8'?>\n" + xml_text

    try:
        write_text(out_path, xml_text, encoding="utf-8")
    except Exception as ex:
        eprint(f"[ERROR] Logic XML の書き込みに失敗しました: {out_path}")
        eprint(f"        {ex}")
        return 1

    print("[INFO] Step3 Logic XML を作成しました")
    print(f"       FROM   : {from_path}")
    print(f"       TO     : {to_path}")
    print(f"       OUTPUT : {out_path}")

    print("[INFO] レコード接続")
    for src_name, tgt_name in RECORD_RULES:
        src_id = src_records.get(src_name)
        tgt_id = tgt_records.get(tgt_name)
        if src_id and tgt_id:
            print(f"       {src_name}({src_id}) -> {tgt_name}({tgt_id})")

    print("[INFO] 線引き")
    for src_business_id, tgt_item_name in FIELD_RULES:
        if tgt_item_name in tgt_fixed_items:
            continue
        src_item_id = src_items.get(src_business_id)
        tgt_item_id = tgt_items.get(tgt_item_name)
        if src_item_id and tgt_item_id:
            print(f"       {src_business_id}({src_item_id}) -> {tgt_item_name}({tgt_item_id})")

    if skipped_records:
        print("[WARN] スキップしたレコード接続")
        for row in skipped_records:
            print(f"       {row}")

    if skipped_fixed_targets:
        print("[INFO] 固定値出力のため線引きしなかった項目")
        for row in skipped_fixed_targets:
            print(f"       {row}")

    if skipped_fields:
        print("[WARN] スキップした線引き")
        for row in skipped_fields:
            print(f"       {row}")

    print(f"[INFO] record_connections = {record_conn_id - 1}")
    print(f"[INFO] line_connections   = {line_id - 1}")

    return 0


if __name__ == "__main__":
    sys.exit(main())
