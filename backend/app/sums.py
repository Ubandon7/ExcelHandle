from __future__ import annotations

import math
import re
from decimal import Decimal, InvalidOperation
from typing import Any

from openpyxl.utils import get_column_letter

from .workbook import SheetData, WorkbookData, normalize_header


NUMBER_TOKEN = re.compile(r"[+-]?(?:(?:\d{1,3}(?:,\d{3})+)|\d+)(?:\.\d+)?|[+-]?\.\d+")


def parse_number(value: Any) -> tuple[Decimal | None, str]:
    if value is None or value == "":
        return None, "empty"
    if isinstance(value, bool):
        return None, "non_numeric"
    if isinstance(value, Decimal):
        return value, "native"
    if isinstance(value, int):
        return Decimal(value), "native"
    if isinstance(value, float):
        if not math.isfinite(value):
            return None, "non_numeric"
        return Decimal(str(value)), "native"

    text = str(value).strip()
    if not text:
        return None, "empty"
    if any(character.isalpha() for character in text):
        return None, "contains_text"
    tokens = NUMBER_TOKEN.findall(text)
    if len(tokens) != 1:
        return None, "invalid_numeric_format"
    token = tokens[0].replace(",", "")
    try:
        return Decimal(token), "text_converted"
    except InvalidOperation:
        return None, "invalid_numeric_format"


def _select_column(sheet: SheetData, selector: dict[str, Any]) -> tuple[int | None, str, str | None]:
    mode = selector.get("matchMode")
    if mode == "FIXED_POSITION":
        index = selector.get("columnIndex")
        if not isinstance(index, int) or index < 1:
            return None, "", "固定列位置无效"
        if index > len(sheet.columns):
            return None, f"第 {index} 列", "缺少列位置"
        return index - 1, sheet.columns[index - 1].display_name, None

    target = normalize_header(selector.get("headerName"))
    matches = [column for column in sheet.columns if column.normalized_name == target]
    if not target:
        return None, "", "未选择表头"
    if not matches:
        return None, str(selector.get("headerName") or ""), "缺少表头"
    if len(matches) > 1:
        return None, str(selector.get("headerName") or ""), "表头不唯一"
    return matches[0].index - 1, matches[0].display_name, None


def execute_sum(workbook: WorkbookData, selector: dict[str, Any]) -> dict[str, Any]:
    results: list[dict[str, Any]] = []
    invalid_details: list[dict[str, Any]] = []
    grand_total = Decimal("0")

    for sheet in workbook.sheets:
        column_index, column_name, selection_error = _select_column(sheet, selector)
        if selection_error:
            results.append(
                {
                    "sheetId": sheet.id,
                    "sheetName": sheet.name,
                    "columnName": column_name,
                    "dataRowCount": len(sheet.rows),
                    "numericCount": 0,
                    "textConvertedCount": 0,
                    "emptyCount": 0,
                    "nonNumericCount": 0,
                    "errorValueCount": 0,
                    "sum": None,
                    "status": "skipped",
                    "message": selection_error,
                }
            )
            continue

        total = Decimal("0")
        numeric_count = 0
        text_converted_count = 0
        empty_count = 0
        non_numeric_count = 0
        error_value_count = 0
        for offset, row in enumerate(sheet.rows, start=2):
            value = row[column_index] if column_index is not None and column_index < len(row) else None
            number, category = parse_number(value)
            if number is not None:
                total += number
                numeric_count += 1
                if category == "text_converted":
                    text_converted_count += 1
                continue
            if category == "empty":
                empty_count += 1
                continue
            non_numeric_count += 1
            if isinstance(value, str) and value.startswith("#"):
                error_value_count += 1
            invalid_details.append(
                {
                    "sheetName": sheet.name,
                    "sourceRow": offset,
                    "cell": f"{get_column_letter((column_index or 0) + 1)}{offset}",
                    "columnIndex": (column_index or 0) + 1,
                    "columnName": column_name,
                    "originalValue": value,
                    "reason": category,
                }
            )

        grand_total += total
        results.append(
            {
                "sheetId": sheet.id,
                "sheetName": sheet.name,
                "columnName": column_name,
                "dataRowCount": len(sheet.rows),
                "numericCount": numeric_count,
                "textConvertedCount": text_converted_count,
                "emptyCount": empty_count,
                "nonNumericCount": non_numeric_count,
                "errorValueCount": error_value_count,
                "sum": str(total),
                "status": "succeeded",
                "message": "" if non_numeric_count == 0 else f"{non_numeric_count} 个值未参与求和",
            }
        )

    return {
        "results": results,
        "invalidDetails": invalid_details,
        "grandTotal": str(grand_total),
        "sheetCount": len(workbook.sheets),
        "invalidCount": len(invalid_details),
    }
