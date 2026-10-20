from __future__ import annotations

from contextlib import redirect_stdout, redirect_stderr
from decimal import Decimal, InvalidOperation
from io import BytesIO
from pathlib import Path
from tempfile import TemporaryDirectory
import logging
import os
import re as regex

import requests as re
from pypdf import PdfReader


def clean_name(value: str) -> str:
    # Ignore punctuation and spaces when comparing names and headings.
    return "".join(character for character in value.casefold() if character.isalnum())


def get_number(value: object) -> Decimal:
    text = str(value).strip().replace("$", "").replace(",", "")
    try:
        number = Decimal(text)
    except InvalidOperation as error:
        raise ValueError(f"Cannot read number {value!r}; check the PDF.") from error
    if not number.is_finite() or number < 0:
        raise ValueError(f"Invalid quantity or price: {value!r}.")
    return number


def get_project_details(text: dict) -> dict[str, str]:
    required = ["Airport Name", "TxDOT CSJ#", "Project Scope", "Bid Opening Date_af_date"]
    for key in required:
        if not text.get(key):
            raise ValueError(f"Missing {key}. Expected the TxDOT aviation summary form.")
    return {
        "Airport name": text["Airport Name"],
        "TxDOT CSJ#": text["TxDOT CSJ#"],
        "Project Scope": text["Project Scope"],
        "Bid Opening Date": text["Bid Opening Date_af_date"],
    }


def get_fields(reader: PdfReader) -> list[dict]:
    fields = []
    for page_number, page in enumerate(reader.pages, start=1):
        for annotation in page.get("/Annots", []):
            box = annotation.get_object()
            if box.get("/Subtype") != "/Widget":
                continue
            field = box
            while "/V" not in field and "/Parent" in field:
                field = field["/Parent"].get_object()
            value = field.get("/V")
            position = box.get("/Rect")
            if value is not None and position is not None and str(value).strip():
                fields.append({
                    "page": page_number,
                    "value": " ".join(str(value).split()),
                    "position": list(position),
                })
    return fields


def get_contractors(reader: PdfReader) -> list[dict]:
    fields = get_fields(reader)
    header_pairs = []
    for page_number in range(1, len(reader.pages) + 1):
        page_fields = [field for field in fields if field["page"] == page_number]
        names = [field for field in page_fields if clean_name(field["value"]) == "bidder"]
        bids = [field for field in page_fields if clean_name(field["value"]) == "basebid"]
        if names and bids:
            if len(names) != 1 or len(bids) != 1:
                raise ValueError("Ambiguous summary column headings.")
            header_pairs.append((page_fields, names[0], bids[0]))
    if len(header_pairs) != 1:
        raise ValueError("Expected one editable bidder summary in the PDF.")
    page_fields, bidder_header, bid_header = header_pairs[0]

    def center_y(field: dict) -> float:
        return (float(field["position"][1]) + float(field["position"][3])) / 2

    def under_header(field: dict, header: dict) -> bool:
        x1, y1, x2, y2 = map(float, field["position"])
        hx1, hy1, hx2, hy2 = map(float, header["position"])
        return min(hx1, hx2) <= (x1 + x2) / 2 <= max(hx1, hx2) and max(y1, y2) < min(hy1, hy2)

    names = [field for field in page_fields if under_header(field, bidder_header)]
    amounts = [field for field in page_fields if under_header(field, bid_header)]
    contractors = []
    used = set()
    for name in sorted(names, key=center_y, reverse=True):
        matches = [index for index, amount in enumerate(amounts) if abs(center_y(name) - center_y(amount)) < 3]
        if len(matches) != 1 or matches[0] in used:
            raise ValueError(f"Cannot match a base bid to {name['value']}.")
        used.add(matches[0])
        contractors.append({"name": name["value"], "base_bid": get_number(amounts[matches[0]]["value"])})
    if not contractors or len(used) != len(amounts):
        raise ValueError("Some summary names or amounts could not be matched.")
    if len({clean_name(contractor["name"]) for contractor in contractors}) != len(contractors):
        raise ValueError("Duplicate contractor names in the summary.")
    return contractors


def get_summary(response: re.Response) -> dict:
    response.raise_for_status()
    reader = PdfReader(BytesIO(response.content))
    fields = reader.get_form_text_fields()
    if not fields:
        raise ValueError("No editable summary fields found. This PDF format is not supported.")
    return {"project": get_project_details(fields), "contractors": get_contractors(reader)}


def read_item_rows(rows: list[dict], pages: list[int]) -> list[dict]:
    items = []
    for row in rows:
        code = " ".join(str(row["Item"]).split())
        if not code or clean_name(code) in {"item", "nan", "none"}:
            continue
        if clean_name(code).startswith(("total", "subtotal")):
            continue
        if not regex.fullmatch(r"[A-Za-z]{1,4}-?\d[\w.\-]*", code):
            raise ValueError(f"Unclear item code {code!r} on pages {pages}; check the OCR.")
        item = {
            "Item": code,
            "Qty": get_number(row["Qty"]),
            "Unit": str(row["Unit"]).strip().upper(),
            "Description": " ".join(str(row["Description"]).split()),
            "Unit Price": get_number(row["Unit Price"]),
            "Total Price": get_number(row["Total Price"]),
            "Pages": pages,
        }
        if item["Qty"] <= 0 or item["Description"].casefold() in {"", "nan", "none"} or item["Unit"] in {"", "NAN", "NONE"}:
            raise ValueError(f"Incomplete item {code} on pages {pages}.")
        items.append(item)
    return items


def validate_bids(contractors: list[dict]) -> None:
    reference = None
    for contractor in contractors:
        if not contractor["items"]:
            raise ValueError(f"No detailed items found for {contractor['name']}.")
        details = {}
        total = Decimal("0")
        for item in contractor["items"]:
            code = clean_name(item["Item"])
            if code in details:
                raise ValueError(f"Repeated item {item['Item']} for {contractor['name']}.")
            details[code] = (item["Qty"], item["Unit"])
            calculated = (item["Qty"] * item["Unit Price"]).quantize(Decimal("0.01"))
            if abs(calculated - item["Total Price"]) > Decimal("0.01"):
                raise ValueError(f"Price mismatch for {contractor['name']}, item {item['Item']}; check the PDF.")
            total += item["Total Price"]
        if abs(total - contractor["base_bid"]) > Decimal("0.01"):
            raise ValueError(f"Item total {total} does not match base bid {contractor['base_bid']} for {contractor['name']}.")
        if reference is not None and details != reference:
            raise ValueError(f"Items, quantities or units differ for {contractor['name']}; review before comparing.")
        reference = details


def match_bid_tables(summary: dict, tables: list[dict], page_text: dict[int, str]) -> list[dict]:
    contractors = [dict(contractor, items=[]) for contractor in summary["contractors"]]
    sections = []
    for page, text in sorted(page_text.items()):
        normalized = clean_name(text)
        if "nameofbidder" not in normalized and "bidby" not in normalized:
            continue
        matches = [index for index, contractor in enumerate(contractors) if clean_name(contractor["name"]) in normalized]
        if len(matches) != 1:
            raise ValueError(f"Cannot uniquely identify the contractor on page {page}.")
        sections.append((page, matches[0]))
    for table in sorted(tables, key=lambda value: min(value["pages"])):
        first_page, last_page = min(table["pages"]), max(table["pages"])
        previous = [section for section in sections if section[0] <= first_page]
        if previous:
            contractor_index = previous[-1][1]
            if any(first_page < page <= last_page and index != contractor_index for page, index in sections):
                raise ValueError("A table spans two different contractor sections.")
        else:
            # Fallback only when a complete table matches ONE summary total.
            total = sum((item["Total Price"] for item in table["items"]), Decimal("0"))
            matches = [index for index, contractor in enumerate(contractors) if contractor["base_bid"] == total]
            if len(matches) != 1:
                raise ValueError(f"Cannot identify the contractor for pages {table['pages']}.")
            contractor_index = matches[0]
        contractors[contractor_index]["items"].extend(table["items"])
    validate_bids(contractors)
    return contractors


def get_bid_details(response: re.Response, summary: dict) -> list[dict]:
    response.raise_for_status()
    required = ["Item", "Qty", "Unit", "Description", "Unit Price", "Total Price"]
    aliases = {clean_name(column): column for column in required}
    aliases.update({"quantity": "Qty", "itemnumber": "Item", "itemno": "Item"})
    tables = []
    page_text = {}
    with TemporaryDirectory() as folder:
        pdf_path = Path(folder) / "project.pdf"
        pdf_path.write_bytes(response.content)
        previous_logging_level = logging.root.manager.disable
        try:
            logging.disable(logging.CRITICAL)
            with open(os.devnull, "w") as quiet, redirect_stdout(quiet), redirect_stderr(quiet):
                from docling.document_converter import DocumentConverter
                result = DocumentConverter().convert(pdf_path)
        finally:
            logging.disable(previous_logging_level)
        if result.status.value != "success":
            raise ValueError(f"OCR did not fully succeed: {result.status}.")
        document = result.document
        for text in document.texts:
            for location in text.prov:
                page_text[location.page_no] = page_text.get(location.page_no, "") + "\n" + text.text
        previous_columns = None
        previous_page = 0
        for table in document.tables:
            pages = sorted({location.page_no for location in table.prov})
            if not pages:
                raise ValueError("OCR returned a table without page information.")
            df = table.export_to_dataframe(doc=document)
            df.columns = [aliases.get(clean_name(str(column)), str(column)) for column in df.columns]
            if not all(column in df.columns for column in required):
                # A continuation may omit headings but keep the preceding columns.
                if previous_columns and min(pages) == previous_page + 1 and len(df.columns) == len(previous_columns):
                    first_row = df.iloc[0].tolist() if len(df) else []
                    if first_row and regex.fullmatch(r"[A-Za-z]{1,4}-?\d[\w.\-]*", str(first_row[0]).strip()):
                        df.columns = previous_columns
                    else:
                        continue
                else:
                    continue
            if df.columns.duplicated().any():
                raise ValueError(f"Ambiguous table columns on pages {pages}.")
            previous_columns = list(df.columns)
            previous_page = max(pages)
            items = read_item_rows(df[required].to_dict("records"), pages)
            if items:
                tables.append({"pages": pages, "items": items})
    if not tables:
        raise ValueError("No detailed bid tables found.")
    return match_bid_tables(summary, tables, page_text)
