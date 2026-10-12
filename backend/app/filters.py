from __future__ import annotations

import re
from dataclasses import dataclass
from typing import Any, Iterable

from .models import ColumnRef, ConditionGroup, FilterCondition
from .workbook import ColumnDef, SheetData, WorkbookData, normalize_header


INVALID_SHEET_NAME = re.compile(r"[\\/*?:\[\]]")
MISSING = object()


def validate_sheet_names(groups: list[dict[str, Any]]) -> None:
    seen: set[str] = set()
    for group in groups:
        name = str(group.get("name") or "").strip()
        if not name:
            raise ValueError("结果 Sheet 名不能为空")
        if len(name) > 31:
            raise ValueError(f"结果 Sheet 名不能超过 31 个字符：{name}")
        if INVALID_SHEET_NAME.search(name):
            raise ValueError(f"结果 Sheet 名包含非法字符：{name}")
        if name.casefold() == "未匹配":
            raise ValueError("“未匹配”是系统保留 Sheet 名")
        normalized = name.casefold()
        if normalized in seen:
            raise ValueError(f"结果 Sheet 名重复：{name}")
        seen.add(normalized)


def validate_condition_tree(root: ConditionGroup, max_depth: int = 5, max_leaves: int = 50) -> None:
    leaves = 0

    def walk(node: ConditionGroup | FilterCondition, depth: int) -> None:
        nonlocal leaves
        if depth > max_depth:
            raise ValueError(f"条件分组不能超过 {max_depth} 层")
        if isinstance(node, ConditionGroup):
            if not node.children:
                raise ValueError("条件组至少需要一个条件或子分组")
            for child in node.children:
                walk(child, depth + 1)
            return
        leaves += 1
        if leaves > max_leaves:
            raise ValueError(f"单个结果组不能超过 {max_leaves} 条条件")
        if node.operator in {"starts_with", "ends_with", "contains", "not_contains", "no_prefix", "no_suffix"}:
            if node.value is None or not str(node.value).strip():
                raise ValueError(f"操作符 {node.operator} 需要条件值")
        if node.operator in {"in", "not_in"} and not [value for value in node.values if str(value).strip()]:
            raise ValueError(f"操作符 {node.operator} 至少需要一个集合值")

    walk(root, 1)


def _resolve_column(sheet: SheetData, condition: FilterCondition) -> int | None:
    return _resolve_column_ref(sheet, condition.column_ref)


def _resolve_column_ref(sheet: SheetData, column_ref: ColumnRef) -> int | None:
    target = normalize_header(column_ref.header_name)
    matches = [
        column
        for column in sheet.columns
        if column.normalized_name == target and column.occurrence == column_ref.occurrence
    ]
    if len(matches) == 1:
        return matches[0].index - 1
    fallback = column_ref.fallback_column_index
    if fallback and fallback <= len(sheet.columns):
        candidate = sheet.columns[fallback - 1]
        if candidate.normalized_name == target:
            return fallback - 1
    return None


def _text(value: Any, trim: bool, case_sensitive: bool) -> str:
    result = "" if value is None else str(value)
    if trim:
        result = result.strip()
    return result if case_sensitive else result.casefold()


def evaluate_condition(condition: FilterCondition, sheet: SheetData, row: list[Any]) -> bool:
    column_index = _resolve_column(sheet, condition)
    if column_index is None:
        return False
    value = row[column_index] if column_index < len(row) else None
    options = condition.options
    cell = _text(value, options.trim_cell, options.case_sensitive)
    empty = value is None or cell == ""

    if condition.operator == "empty":
        return empty
    if condition.operator == "not_empty":
        return not empty

    if condition.operator in {"not_contains", "not_in"} and empty:
        return False

    if condition.operator in {"in", "not_in"}:
        values = {
            _text(item, options.trim_input, options.case_sensitive)
            for item in condition.values
            if str(item).strip()
        }
        matched = cell in values
        return matched if condition.operator == "in" else not matched

    target = _text(condition.value, options.trim_input, options.case_sensitive)
    if condition.operator == "starts_with":
        return cell.startswith(target)
    if condition.operator == "ends_with":
        return cell.endswith(target)
    if condition.operator == "contains":
        return target in cell
    if condition.operator == "not_contains":
        return target not in cell
    if condition.operator == "no_prefix":
        position = cell.find(target)
        return position <= 0
    if condition.operator == "no_suffix":
        position = cell.rfind(target)
        return position < 0 or position + len(target) >= len(cell)
    raise ValueError(f"未知筛选操作符：{condition.operator}")


def evaluate_tree(node: ConditionGroup | FilterCondition, sheet: SheetData, row: list[Any]) -> bool:
    if isinstance(node, FilterCondition):
        return evaluate_condition(node, sheet, row)
    results = (evaluate_tree(child, sheet, row) for child in node.children)
    return all(results) if node.relation == "AND" else any(results)


@dataclass(frozen=True)
class UnionColumn:
    key: str
    name: str


def union_columns(sheets: Iterable[SheetData]) -> list[UnionColumn]:
    result: list[UnionColumn] = []
    seen: set[str] = set()
    for sheet in sheets:
        for column in sheet.columns:
            if column.key in seen:
                continue
            seen.add(column.key)
            result.append(UnionColumn(column.key, column.display_name))
    return result


def _map_row(sheet: SheetData, row: list[Any], schema: list[UnionColumn], add_source: bool) -> list[Any]:
    source_by_key = {column.key: column.index - 1 for column in sheet.columns}
    values = [row[source_by_key[column.key]] if column.key in source_by_key else None for column in schema]
    return [sheet.name, *values] if add_source else values


def execute_filter_groups(workbook: WorkbookData, groups: list[dict[str, Any]]) -> dict[str, Any]:
    validate_sheet_names(groups)
    parsed: list[tuple[dict[str, Any], ConditionGroup, set[str]]] = []
    sheet_ids = {sheet.id for sheet in workbook.sheets}
    for group in groups:
        tree = ConditionGroup.model_validate(group["conditionTree"])
        validate_condition_tree(tree)
        scope = group.get("sourceScope") or {"mode": "ALL", "sheetId": None}
        if scope.get("mode") == "SINGLE":
            selected = scope.get("sheetId")
            if selected not in sheet_ids:
                raise ValueError(f"结果组“{group['name']}”选择的源 Sheet 不存在")
            applicable = {selected}
        else:
            applicable = set(sheet_ids)
        parsed.append((group, tree, applicable))

    output: dict[str, Any] = {
        "sheetOrder": [],
        "sheets": {},
        "sourceRowCount": sum(len(sheet.rows) for sheet in workbook.sheets),
    }

    for group, _, applicable in parsed:
        group_sheets = [sheet for sheet in workbook.sheets if sheet.id in applicable]
        add_source = len(group_sheets) != 1 or (group.get("sourceScope") or {}).get("mode") == "ALL"
        schema = union_columns(group_sheets)
        headers = [column.name for column in schema]
        if add_source:
            source_name = "来源 Sheet"
            existing = {normalize_header(header) for header in headers}
            if normalize_header(source_name) in existing:
                source_name = "来源 Sheet（系统）"
            headers = [source_name, *headers]
        key = group["id"]
        output["sheetOrder"].append(key)
        output["sheets"][key] = {
            "key": key,
            "name": group["name"],
            "headers": headers,
            "rows": [],
            "rowCount": 0,
            "isSystem": False,
            "addSourceSheetColumn": add_source,
            "schema": [{"key": column.key, "name": column.name} for column in schema],
        }

    unmatched_schema = union_columns(workbook.sheets)
    unmatched_headers = [column.name for column in unmatched_schema]
    unmatched_source_name = "来源 Sheet"
    if normalize_header(unmatched_source_name) in {normalize_header(value) for value in unmatched_headers}:
        unmatched_source_name = "来源 Sheet（系统）"
    unmatched_key = "__unmatched__"
    output["sheetOrder"].append(unmatched_key)
    output["sheets"][unmatched_key] = {
        "key": unmatched_key,
        "name": "未匹配",
        "headers": [unmatched_source_name, *unmatched_headers],
        "rows": [],
        "rowCount": 0,
        "isSystem": True,
        "addSourceSheetColumn": True,
        "schema": [{"key": column.key, "name": column.name} for column in unmatched_schema],
    }

    for sheet in workbook.sheets:
        for row in sheet.rows:
            assigned = False
            for group, tree, applicable in parsed:
                if sheet.id not in applicable:
                    continue
                if evaluate_tree(tree, sheet, row):
                    dataset = output["sheets"][group["id"]]
                    schema = [UnionColumn(item["key"], item["name"]) for item in dataset["schema"]]
                    dataset["rows"].append(_map_row(sheet, row, schema, dataset["addSourceSheetColumn"]))
                    assigned = True
                    break
            if not assigned:
                dataset = output["sheets"][unmatched_key]
                dataset["rows"].append(_map_row(sheet, row, unmatched_schema, True))

    assigned_count = 0
    for dataset in output["sheets"].values():
        dataset["rowCount"] = len(dataset["rows"])
        assigned_count += dataset["rowCount"]
    output["assignedRowCount"] = assigned_count
    output["rowCountConserved"] = assigned_count == output["sourceRowCount"]
    return output


def _category_identity(value: Any) -> tuple[str, str]:
    if value is None or (isinstance(value, str) and not value.strip()):
        return ("blank", "")
    return (type(value).__name__, str(value))


def _category_label(value: Any) -> str:
    if value is None or (isinstance(value, str) and not value.strip()):
        return "空白"
    return str(value).strip()


def _unique_sheet_name(value: Any, used: set[str]) -> str:
    base = INVALID_SHEET_NAME.sub("_", _category_label(value))
    base = re.sub(r"[\x00-\x1f]", "_", base).strip() or "空白"
    base = base[:31]
    candidate = base
    sequence = 2
    while candidate.casefold() in used:
        suffix = f"_{sequence}"
        candidate = f"{base[:31 - len(suffix)]}{suffix}"
        sequence += 1
    used.add(candidate.casefold())
    return candidate


def execute_column_split(
    workbook: WorkbookData,
    column_ref_value: dict[str, Any],
    source_scope: dict[str, Any],
) -> dict[str, Any]:
    """Split rows into one result sheet per distinct value in the selected column."""
    return _execute_value_split(
        workbook,
        column_ref_value,
        source_scope,
        lambda value: (True, value),
    )


def execute_affix_split(
    workbook: WorkbookData,
    column_ref_value: dict[str, Any],
    source_scope: dict[str, Any],
    kind: str,
    delimiter: str,
) -> dict[str, Any]:
    """Split rows by the prefix or suffix identified with a user-provided delimiter."""
    if kind not in {"PREFIX", "SUFFIX"}:
        raise ValueError("前后缀类型无效")
    if not delimiter or len(delimiter) > 5:
        raise ValueError("连接符长度必须为 1～5 个字符")

    def extract(value: Any) -> tuple[bool, Any]:
        text = "" if value is None else str(value).strip()
        if not text:
            return (False, None)
        if kind == "PREFIX":
            position = text.find(delimiter)
            return (True, text[: position + len(delimiter)]) if position > 0 else (False, None)
        position = text.rfind(delimiter)
        if position >= 0 and position + len(delimiter) < len(text):
            return (True, text[position:])
        return (False, None)

    return _execute_value_split(workbook, column_ref_value, source_scope, extract)


def _execute_value_split(
    workbook: WorkbookData,
    column_ref_value: dict[str, Any],
    source_scope: dict[str, Any],
    classifier: Any,
) -> dict[str, Any]:
    column_ref = ColumnRef.model_validate(column_ref_value)
    if source_scope.get("mode") == "SINGLE":
        selected = source_scope.get("sheetId")
        if not selected:
            raise ValueError("请选择需要分类的源 Sheet")
        try:
            source_sheets = [workbook.sheet_by_id(selected)]
        except KeyError as exc:
            raise ValueError("选择的源 Sheet 不存在") from exc
    else:
        source_sheets = workbook.sheets

    schema = union_columns(source_sheets)
    add_source = len(source_sheets) != 1 or source_scope.get("mode") == "ALL"
    headers = [column.name for column in schema]
    if add_source:
        source_name = "来源 Sheet"
        if normalize_header(source_name) in {normalize_header(header) for header in headers}:
            source_name = "来源 Sheet（系统）"
        headers = [source_name, *headers]

    output: dict[str, Any] = {
        "sheetOrder": [],
        "sheets": {},
        "sourceRowCount": sum(len(sheet.rows) for sheet in source_sheets),
    }
    categories: dict[tuple[str, str], str] = {}
    used_names = {"未匹配".casefold()}
    unmatched_rows: list[list[Any]] = []

    for sheet in source_sheets:
        column_index = _resolve_column_ref(sheet, column_ref)
        for row in sheet.rows:
            if column_index is None:
                unmatched_rows.append(_map_row(sheet, row, schema, add_source))
                continue
            value = row[column_index] if column_index < len(row) else None
            matched, category_value = classifier(value)
            if not matched:
                unmatched_rows.append(_map_row(sheet, row, schema, add_source))
                continue
            identity = _category_identity(category_value)
            key = categories.get(identity)
            if key is None:
                key = f"category-{len(categories) + 1}"
                categories[identity] = key
                output["sheetOrder"].append(key)
                output["sheets"][key] = {
                    "key": key,
                    "name": _unique_sheet_name(category_value, used_names),
                    "headers": headers,
                    "rows": [],
                    "rowCount": 0,
                    "isSystem": False,
                    "addSourceSheetColumn": add_source,
                    "schema": [{"key": column.key, "name": column.name} for column in schema],
                    "categoryValue": category_value,
                }
            output["sheets"][key]["rows"].append(_map_row(sheet, row, schema, add_source))

    if unmatched_rows or not categories:
        key = "__unmatched__"
        output["sheetOrder"].append(key)
        output["sheets"][key] = {
            "key": key,
            "name": "未匹配",
            "headers": headers,
            "rows": unmatched_rows,
            "rowCount": len(unmatched_rows),
            "isSystem": True,
            "addSourceSheetColumn": add_source,
            "schema": [{"key": column.key, "name": column.name} for column in schema],
        }

    assigned_count = 0
    for dataset in output["sheets"].values():
        dataset["rowCount"] = len(dataset["rows"])
        assigned_count += dataset["rowCount"]
    output["assignedRowCount"] = assigned_count
    output["rowCountConserved"] = assigned_count == output["sourceRowCount"]
    return output
