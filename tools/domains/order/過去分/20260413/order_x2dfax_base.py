from __future__ import annotations

import argparse
import datetime as dt
from dataclasses import dataclass, field
from decimal import Decimal, InvalidOperation
from pathlib import Path
from typing import Iterable, List, Optional
import xml.etree.ElementTree as ET

SOH = b"\x01"
FF = b"\xff"
CRLF = b"\r\n"


def local_name(tag: str) -> str:
    return tag.rsplit("}", 1)[-1] if "}" in tag else tag


def iter_children(elem: ET.Element, name: str) -> Iterable[ET.Element]:
    for child in list(elem):
        if local_name(child.tag) == name:
            yield child


def find_child(elem: Optional[ET.Element], name: str) -> Optional[ET.Element]:
    if elem is None:
        return None
    for child in iter_children(elem, name):
        return child
    return None


def find_path(elem: Optional[ET.Element], *names: str) -> Optional[ET.Element]:
    cur = elem
    for name in names:
        cur = find_child(cur, name)
        if cur is None:
            return None
    return cur


def text_at(elem: Optional[ET.Element], *names: str) -> str:
    node = find_path(elem, *names)
    if node is None or node.text is None:
        return ""
    return node.text.strip()


def first_desc(elem: Optional[ET.Element], name: str) -> Optional[ET.Element]:
    if elem is None:
        return None
    for node in elem.iter():
        if local_name(node.tag) == name:
            return node
    return None


def safe_decimal(value: str) -> Decimal:
    s = str(value or "").strip()
    if not s:
        return Decimal("0")
    try:
        return Decimal(s)
    except InvalidOperation:
        cleaned = "".join(ch for ch in s if ch.isdigit() or ch in ".-")
        return Decimal(cleaned or "0")


def ascii_trim(text: str) -> str:
    return str(text or "").replace("\r", " ").replace("\n", " ").strip()


def cf_str(val: str, length: int) -> str:
    text = ascii_trim(val)
    if len(text) > length:
        text = text[:length]
    return text


def cf_int(val: str, length: int) -> str:
    text = ascii_trim(val)
    if text == "":
        return ""
    digits = "".join(ch for ch in text if ch.isdigit())
    if len(digits) > length:
        digits = digits[-length:]
    return digits


def cf_ymd_yy_mm_dd(val: str) -> str:
    s = ascii_trim(val)
    if not s:
        return ""
    digits = "".join(ch for ch in s if ch.isdigit())
    if len(digits) >= 8:
        digits = digits[-8:] if len(digits) > 8 else digits
        return digits[2:4] + "." + digits[4:6] + "." + digits[6:8]
    if len(digits) == 6:
        return digits[0:2] + "." + digits[2:4] + "." + digits[4:6]
    return s


def cf_jpn(text: str, max_bytes: int) -> str:
    s = str(text or "").strip()
    encoded = s.encode("shift_jis", errors="ignore")
    if len(encoded) <= max_bytes:
        return s
    cut = encoded[:max_bytes]
    return cut.decode("shift_jis", errors="ignore")


@dataclass
class ItemModel:
    item_name: str = ""
    spec: str = ""
    order_item_code: str = ""
    unit_multiple: str = ""
    num_of_order_units: str = ""
    unit_of_measure: str = ""
    quantity: str = ""
    item_net_price: str = ""
    net_amount: str = ""


@dataclass
class OrderModel:
    trade_number: str = ""
    seller_code: str = ""
    seller_name: str = ""
    buyer_name: str = ""
    receiver_code: str = ""
    receiver_name: str = ""
    sub_major_category: str = ""
    trade_type_code: str = ""
    order_date: str = ""
    delivery_date: str = ""
    note_text: str = ""
    note_text_sbcs: str = ""
    fax_no: str = ""
    items: List[ItemModel] = field(default_factory=list)


@dataclass
class D1Field:
    row_no: int
    col_no: int
    value: str
    option: str = ""

    def to_bytes(self) -> bytes:
        body = bytearray()
        body.extend(f"{self.row_no:03d}".encode("ascii"))
        body.extend(SOH)
        body.extend(f"{self.col_no:03d}".encode("ascii"))
        body.extend(SOH)
        body.extend(str(self.value).encode("shift_jis", errors="ignore"))
        body.extend(SOH)
        body.extend(str(self.option).encode("ascii", errors="ignore"))
        body.extend(FF)
        return bytes(body)


class OrderXmlNormalizer:
    def __init__(self, xml_path: Path):
        self.tree = ET.parse(xml_path)
        self.root = self.tree.getroot()

    def normalize(self, bbfax: str) -> List[OrderModel]:
        orders: List[OrderModel] = []
        message = first_desc(self.root, "message")
        list_of_orders = first_desc(message, "listOfOrders") if message is not None else None
        if list_of_orders is None:
            return orders

        buyer = find_child(list_of_orders, "buyer")
        buyer_name = self._best_name(buyer)

        for order_elem in iter_children(list_of_orders, "order"):
            order = OrderModel()
            order.buyer_name = buyer_name
            order.trade_number = text_at(order_elem, "tradeID", "tradeNumber")
            order.order_date = text_at(order_elem, "tradeSummary", "dates", "orderDate")
            order.delivery_date = text_at(order_elem, "tradeSummary", "dates", "deliveryDateToReceiver")
            order.sub_major_category = text_at(order_elem, "tradeSummary", "goodsMajorCategory", "subMajorCategory")
            order.trade_type_code = text_at(order_elem, "tradeSummary", "instructions", "tradeTypeCode")
            order.note_text = text_at(order_elem, "tradeSummary", "note", "text")
            order.note_text_sbcs = text_at(order_elem, "tradeSummary", "note", "text_sbcs")

            order.receiver_code = text_at(order_elem, "parties", "receiver", "code")
            order.receiver_name = text_at(order_elem, "parties", "receiver", "name")
            order.seller_code = text_at(order_elem, "parties", "seller", "code")
            order.seller_name = text_at(order_elem, "parties", "seller", "name")

            order.fax_no = self._extract_fax(order.note_text_sbcs) or bbfax

            for line_item in iter_children(order_elem, "lineItem"):
                item = ItemModel()
                item.item_name = text_at(line_item, "itemID", "name")
                item.spec = text_at(line_item, "itemInfo", "itemSpec", "spec")
                item.order_item_code = text_at(line_item, "itemID", "orderItemCode")
                item.unit_multiple = text_at(line_item, "quantities", "unitMultiple")
                item.num_of_order_units = text_at(line_item, "quantities", "orderQuantity", "numOfOrderUnits")
                item.unit_of_measure = text_at(line_item, "quantities", "unitOfMeasure")
                item.quantity = text_at(line_item, "quantities", "orderQuantity", "quantity")
                item.item_net_price = text_at(line_item, "amounts", "itemNetPrice")
                item.net_amount = text_at(line_item, "amounts", "itemNetPrice")
                order.items.append(item)

            orders.append(order)
        return orders

    @staticmethod
    def _best_name(elem: Optional[ET.Element]) -> str:
        if elem is None:
            return ""
        for key in ("name", "name_sbcs"):
            txt = text_at(elem, key)
            if txt:
                return txt
        return ""

    @staticmethod
    def _extract_fax(text_sbcs: str) -> str:
        temp = str(text_sbcs or "").strip()
        if len(temp) >= 3 and temp[:3].upper() == "FAX":
            return temp[4:].strip()
        return ""


class DfaxBuilder:
    def __init__(self, bbfax: str, now_text: Optional[str] = None):
        self.bbfax = bbfax
        self.now_text = now_text or dt.datetime.now().strftime("%y.%m.%d %H:%M")

    def build_lines(self, orders: List[OrderModel]) -> List[bytes]:
        lines: List[bytes] = []
        current_fax = None
        current_seller = None
        count_page = 0
        count_line = 0

        for order in orders:
            fax = order.fax_no or self.bbfax
            if fax != current_fax or order.seller_code != current_seller:
                current_fax = fax
                current_seller = order.seller_code
                count_page += 1
                count_line = 0
                lines.append(self.render_d1([
                    D1Field(1, 2, cf_str(fax, 20)),
                    D1Field(2, 2, cf_str(cf_str(order.seller_code, 4) + ":" + cf_jpn(order.seller_name, 40), 44)),
                    D1Field(2, 11, cf_str(self.now_text, 14)),
                    D1Field(2, 15, str(count_page)),
                ]))

            count_line += 4
            lines.append(self.render_d1([
                D1Field(count_line, 6, cf_str(order.trade_number, 9)),
                D1Field(count_line, 12, cf_ymd_yy_mm_dd(order.order_date)),
                D1Field(count_line, 14, cf_ymd_yy_mm_dd(order.delivery_date)),
                D1Field(count_line + 1, 3, cf_jpn(order.buyer_name, 40)),
                D1Field(count_line + 2, 3, cf_jpn(order.receiver_name, 40)),
                D1Field(count_line + 2, 5, cf_str(order.receiver_code, 6)),
                D1Field(count_line + 2, 6, cf_str(order.sub_major_category, 10)),
                D1Field(count_line + 2, 7, cf_str(order.trade_type_code, 4)),
            ]))
            lines.append(self.render_d1([
                D1Field(count_line + 2, 8, cf_str(order.seller_code, 4)),
                D1Field(count_line + 2, 10, cf_jpn(order.seller_name, 40)),
            ]))

            qty_total = Decimal("0")
            amt_total = Decimal("0")
            for item in order.items:
                count_line += 1
                lines.append(self.render_d1([
                    D1Field(count_line, 2, cf_jpn(item.item_name, 50)),
                ]))
                count_line += 1
                unit_code = "" if item.unit_of_measure == "00" else item.unit_of_measure
                lines.append(self.render_d1([
                    D1Field(count_line, 2, ""),
                    D1Field(count_line, 4, cf_jpn(item.spec, 20)),
                    D1Field(count_line, 5, cf_str(item.order_item_code, 13)),
                    D1Field(count_line, 7, cf_int(item.unit_multiple, 4)),
                    D1Field(count_line, 8, cf_str(item.num_of_order_units, 6)),
                    D1Field(count_line, 9, cf_str(unit_code, 4)),
                    D1Field(count_line, 10, cf_str(item.quantity, 6)),
                    D1Field(count_line, 12, cf_str(item.item_net_price, 10)),
                    D1Field(count_line, 14, cf_str(item.net_amount, 10)),
                ]))
                qty_total += safe_decimal(item.quantity)
                amt_total += safe_decimal(item.net_amount)

            count_line += ((9 - len(order.items)) * 2) + 1
            lines.append(self.render_d1([
                D1Field(count_line, 10, cf_int(str(int(qty_total)) if qty_total == qty_total.to_integral_value() else str(qty_total), 6)),
                D1Field(count_line, 14, cf_int(str(int(amt_total)) if amt_total == amt_total.to_integral_value() else str(amt_total), 10)),
            ]))

            count_line += 1
            lines.append(self.render_d1([
                D1Field(count_line, 3, cf_jpn(order.note_text, 20)),
            ]))

        return lines

    @staticmethod
    def render_d1(fields: List[D1Field]) -> bytes:
        body = bytearray(b"D1")
        for field in fields:
            body.extend(field.to_bytes())
        return bytes(body) + CRLF


def main() -> None:
    parser = argparse.ArgumentParser(description="order-x2dfax_base.py 置き換え用全体版")
    parser.add_argument("xml", type=Path, help="input order xml")
    parser.add_argument("--out", type=Path, default=Path("dfax_d1_preview.txt"), help="output txt")
    parser.add_argument("--bbfax", default="0000000000", help="default fax when XML note/text_sbcs has no FAX")
    parser.add_argument("--now", default="", help='override now text like "26.04.05 22:55"')
    args = parser.parse_args()

    normalizer = OrderXmlNormalizer(args.xml)
    orders = normalizer.normalize(args.bbfax)
    builder = DfaxBuilder(bbfax=args.bbfax, now_text=args.now or None)
    lines = builder.build_lines(orders)

    with args.out.open("wb") as fp:
        for line in lines:
            fp.write(line)

    print(f"output: {args.out}")
    print(f"orders: {len(orders)}")
    print(f"lines: {len(lines)}")


if __name__ == "__main__":
    main()
