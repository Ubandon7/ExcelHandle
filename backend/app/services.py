from __future__ import annotations

import re
import threading
from pathlib import Path
from typing import Any

from .exports import export_filter_workbook, export_sum_workbook
from .filters import execute_affix_split, execute_column_split, execute_filter_groups, validate_sheet_names
from .storage import Repository
from .sums import execute_sum
from .workbook import SUPPORTED_EXTENSIONS, WorkbookData, WorkbookReader, _build_sheet, normalize_header


class ApplicationService:
    def __init__(self, repository: Repository | None = None, reader: WorkbookReader | None = None) -> None:
        self.repository = repository or Repository()
        self.reader = reader or WorkbookReader()
        self._workbook_cache: dict[str, WorkbookData] = {}
        self._result_cache: dict[str, dict[str, Any]] = {}
        self._cache_lock = threading.RLock()
        self._export_lock = threading.Lock()
        for record in self.repository.list_files():
            for run in self.repository.list_runs("filter", record["id"]):
                if not run.get("saved", False):
                    self.repository.delete_run("filter", run["id"])

    @staticmethod
    def public_file(record: dict[str, Any]) -> dict[str, Any]:
        return {key: value for key, value in record.items() if key != "path"}

    def upload(self, filename: str, content: bytes) -> dict[str, Any]:
        if not filename:
            raise ValueError("请选择文件")
        extension = Path(filename).suffix.casefold()
        if extension not in SUPPORTED_EXTENSIONS:
            supported = ", ".join(sorted(SUPPORTED_EXTENSIONS))
            raise ValueError(f"不支持 {extension or '无扩展名'}；支持：{supported}")
        if not content:
            raise ValueError("上传文件为空")
        if len(content) > 100 * 1024 * 1024:
            raise ValueError("文件不能超过 100 MB")

        staging = self.repository.files_dir / f"staging_{self.repository.new_id('upload')}{extension}"
        staging.write_bytes(content)
        try:
            try:
                workbook = self.reader.read(staging)
                metadata = workbook.metadata()
            except (ValueError, RuntimeError):
                raise
            except Exception as exc:
                raise ValueError("无法解析该文件，请确认文件未损坏、未加密且格式与扩展名一致") from exc
        finally:
            staging.unlink(missing_ok=True)
        record = self.repository.create_file(filename, content, metadata)
        with self._cache_lock:
            self._workbook_cache[record["id"]] = workbook
        return self.public_file(record)

    def get_file(self, file_id: str) -> dict[str, Any] | None:
        record = self.repository.get_file(file_id)
        return self.public_file(record) if record else None

    def list_files(self) -> list[dict[str, Any]]:
        return [self.public_file(record) for record in self.repository.list_files()]

    def load_workbook(self, file_id: str) -> WorkbookData:
        with self._cache_lock:
            cached = self._workbook_cache.get(file_id)
        if cached:
            return cached
        record = self.repository.get_file(file_id)
        if not record:
            raise KeyError("文件不存在或已过期")
        path = Path(record["path"])
        if not path.exists():
            raise KeyError("源文件已经被清理")
        workbook = self.reader.read(path)
        with self._cache_lock:
            self._workbook_cache[file_id] = workbook
        return workbook

    def preview_source(self, file_id: str, sheet_id: str, page: int, page_size: int) -> dict[str, Any]:
        workbook = self.load_workbook(file_id)
        try:
            sheet = workbook.sheet_by_id(sheet_id)
        except KeyError as exc:
            raise KeyError("Sheet 不存在") from exc
        start = (page - 1) * page_size
        end = start + page_size
        return {
            "sheetId": sheet.id,
            "sheetName": sheet.name,
            "headers": sheet.headers,
            "rows": sheet.rows[start:end],
            "page": page,
            "pageSize": page_size,
            "total": len(sheet.rows),
        }

    def affix_suggestions(
        self,
        file_id: str,
        header_name: str,
        occurrence: int,
        kind: str,
        delimiter: str,
        scope_mode: str,
        sheet_id: str | None,
    ) -> dict[str, Any]:
        if kind not in {"PREFIX", "SUFFIX"}:
            raise ValueError("前后缀类型无效")
        if not delimiter or len(delimiter) > 5:
            raise ValueError("标识符长度必须为 1～5 个字符")
        workbook = self.load_workbook(file_id)
        if scope_mode == "SINGLE":
            if not sheet_id:
                raise ValueError("请选择需要识别的 Sheet")
            try:
                sheets = [workbook.sheet_by_id(sheet_id)]
            except KeyError as exc:
                raise ValueError("选择的 Sheet 不存在") from exc
        else:
            sheets = workbook.sheets

        target = normalize_header(header_name)
        counts: dict[str, dict[str, Any]] = {}
        unmarked_count = 0
        scanned_count = 0
        for sheet in sheets:
            columns = [
                column for column in sheet.columns
                if column.normalized_name == target and column.occurrence == occurrence
            ]
            if len(columns) != 1:
                continue
            column_index = columns[0].index - 1
            for row in sheet.rows:
                value = row[column_index] if column_index < len(row) else None
                text = "" if value is None else str(value).strip()
                if not text:
                    continue
                scanned_count += 1
                if kind == "PREFIX":
                    position = text.find(delimiter)
                    affix = text[: position + len(delimiter)] if position > 0 else None
                else:
                    position = text.rfind(delimiter)
                    affix = text[position:] if position >= 0 and position + len(delimiter) < len(text) else None
                if not affix:
                    unmarked_count += 1
                    continue
                key = affix.casefold()
                item = counts.setdefault(key, {"value": affix, "count": 0})
                item["count"] += 1

        items = sorted(counts.values(), key=lambda item: (-item["count"], item["value"].casefold()))
        return {
            "items": items,
            "delimiter": delimiter,
            "kind": kind,
            "scannedCount": scanned_count,
            "unmarkedCount": unmarked_count,
        }

    def create_group(self, file_id: str, payload: dict[str, Any]) -> dict[str, Any]:
        if not self.repository.get_file(file_id):
            raise KeyError("文件不存在或已过期")
        existing = self.repository.list_groups(file_id)
        proposed = dict(payload)
        if not str(proposed.get("name") or "").strip():
            used = {str(group.get("name") or "").casefold() for group in existing}
            sequence = 1
            while f"sheet{sequence}" in used:
                sequence += 1
            proposed["name"] = f"Sheet{sequence}"
        validate_sheet_names([*existing, proposed])
        return self.repository.create_group(file_id, proposed)

    def list_saved_groups(self) -> list[dict[str, Any]]:
        return self.repository.list_saved_groups()

    def create_group_from_saved(self, file_id: str, saved_group_id: str) -> dict[str, Any]:
        target = self.repository.get_file(file_id)
        if not target:
            raise KeyError("文件不存在或已过期")
        saved = self.repository.get_group(saved_group_id)
        if not saved or not saved.get("saved", False):
            raise KeyError("已保存的筛选条件不存在")

        existing = self.repository.list_groups(file_id)
        used_names = {str(group.get("name") or "").casefold() for group in existing}
        base = str(saved.get("name") or "筛选条件").strip()[:31] or "筛选条件"
        name = base
        sequence = 2
        while name.casefold() in used_names or name.casefold() == "未匹配":
            suffix = f"_{sequence}"
            name = f"{base[:31 - len(suffix)]}{suffix}"
            sequence += 1

        scope = dict(saved.get("sourceScope") or {"mode": "ALL", "sheetId": None})
        sheet_ids = {sheet["id"] for sheet in target["metadata"].get("sheets", [])}
        if scope.get("mode") == "SINGLE" and scope.get("sheetId") not in sheet_ids:
            scope = {"mode": "ALL", "sheetId": None}
        payload = {
            "name": name,
            "sourceScope": scope,
            "conditionTree": saved["conditionTree"],
            "saved": False,
        }
        validate_sheet_names([*existing, payload])
        return self.repository.create_group(file_id, payload)

    def update_group(self, group_id: str, payload: dict[str, Any]) -> dict[str, Any]:
        current = self.repository.get_group(group_id)
        if not current:
            raise KeyError("筛选结果组不存在")
        proposed = {**current, **payload}
        peers = [group for group in self.repository.list_groups(current["fileId"]) if group["id"] != group_id]
        validate_sheet_names([*peers, proposed])
        record = self.repository.update_group(group_id, payload)
        if not record:
            raise KeyError("筛选结果组不存在")
        return record

    def clear_current_groups(self, file_id: str) -> int:
        if not self.repository.get_file(file_id):
            raise KeyError("文件不存在或已过期")
        cleared = 0
        for group in self.repository.list_groups(file_id):
            if not group.get("active", True):
                continue
            if group.get("saved", False):
                self.repository.update_group(group["id"], {"active": False})
            else:
                self.repository.delete_group(group["id"])
            cleared += 1
        return cleared

    def create_filter_run(self, file_id: str, ordered_group_ids: list[str]) -> dict[str, Any]:
        if not self.repository.get_file(file_id):
            raise KeyError("文件不存在或已过期")
        if len(ordered_group_ids) != len(set(ordered_group_ids)):
            raise ValueError("结果组顺序中存在重复 ID")
        for run in self.repository.list_runs("filter", file_id):
            if not run.get("saved", False) and run["status"] not in {"queued", "running"}:
                with self._cache_lock:
                    self._result_cache.pop(run["id"], None)
                self.repository.delete_run("filter", run["id"])
        return self.repository.create_run(
            "filter",
            file_id,
            {
                "mode": "GROUPS",
                "matchPolicy": "FIRST_MATCH",
                "orderedGroupIds": ordered_group_ids,
                "unmatchedPolicy": "OUTPUT_TO_UNMATCHED_SHEET",
            },
        )

    def create_column_split_run(self, file_id: str, payload: dict[str, Any]) -> dict[str, Any]:
        if not self.repository.get_file(file_id):
            raise KeyError("文件不存在或已过期")
        scope = payload.get("sourceScope") or {"mode": "ALL", "sheetId": None}
        if scope.get("mode") == "SINGLE" and not scope.get("sheetId"):
            raise ValueError("请选择需要分类的源 Sheet")
        for run in self.repository.list_runs("filter", file_id):
            if not run.get("saved", False) and run["status"] not in {"queued", "running"}:
                with self._cache_lock:
                    self._result_cache.pop(run["id"], None)
                self.repository.delete_run("filter", run["id"])
        return self.repository.create_run(
            "filter",
            file_id,
            {
                "mode": "COLUMN_SPLIT",
                "columnRef": payload["columnRef"],
                "sourceScope": scope,
            },
        )

    def create_affix_split_run(self, file_id: str, payload: dict[str, Any]) -> dict[str, Any]:
        if not self.repository.get_file(file_id):
            raise KeyError("文件不存在或已过期")
        scope = payload.get("sourceScope") or {"mode": "ALL", "sheetId": None}
        if scope.get("mode") == "SINGLE" and not scope.get("sheetId"):
            raise ValueError("请选择需要分类的源 Sheet")
        delimiter = str(payload.get("delimiter") or "")
        if not delimiter or len(delimiter) > 5:
            raise ValueError("连接符长度必须为 1～5 个字符")
        for run in self.repository.list_runs("filter", file_id):
            if not run.get("saved", False) and run["status"] not in {"queued", "running"}:
                with self._cache_lock:
                    self._result_cache.pop(run["id"], None)
                self.repository.delete_run("filter", run["id"])
        return self.repository.create_run(
            "filter",
            file_id,
            {
                "mode": "AFFIX_SPLIT",
                "columnRef": payload["columnRef"],
                "sourceScope": scope,
                "kind": payload["kind"],
                "delimiter": delimiter,
            },
        )

    def execute_filter_run(self, run_id: str) -> None:
        try:
            run = self.repository.get_run("filter", run_id)
            if not run:
                return
            self.repository.update_run("filter", run_id, status="running", progress=10)
            workbook = self.load_workbook(run["fileId"])
            payload = run["payload"]
            if payload.get("mode") == "COLUMN_SPLIT":
                self.repository.update_run("filter", run_id, progress=35)
                result = execute_column_split(workbook, payload["columnRef"], payload["sourceScope"])
            elif payload.get("mode") == "AFFIX_SPLIT":
                self.repository.update_run("filter", run_id, progress=35)
                result = execute_affix_split(
                    workbook,
                    payload["columnRef"],
                    payload["sourceScope"],
                    payload["kind"],
                    payload["delimiter"],
                )
            else:
                groups: list[dict[str, Any]] = []
                for priority, group_id in enumerate(payload["orderedGroupIds"], start=1):
                    group = self.repository.get_group(group_id)
                    if not group or group["fileId"] != run["fileId"]:
                        raise ValueError(f"筛选结果组不存在或不属于当前文件：{group_id}")
                    if not group.get("active", True):
                        raise ValueError(f"筛选结果组当前未启用：{group['name']}")
                    group["priority"] = priority
                    groups.append(group)
                self.repository.update_run("filter", run_id, progress=35)
                result = execute_filter_groups(workbook, groups)
            path = self.repository.write_result(run_id, result)
            summary = [
                {
                    "key": key,
                    "name": result["sheets"][key]["name"],
                    "rowCount": result["sheets"][key]["rowCount"],
                    "headers": result["sheets"][key]["headers"],
                    "isSystem": result["sheets"][key]["isSystem"],
                }
                for key in result["sheetOrder"]
            ]
            with self._cache_lock:
                self._result_cache[run_id] = result
            self.repository.update_run(
                "filter",
                run_id,
                status="succeeded",
                progress=100,
                resultPath=str(path),
                summary=summary,
                sourceRowCount=result["sourceRowCount"],
                assignedRowCount=result["assignedRowCount"],
                rowCountConserved=result["rowCountConserved"],
            )
        except Exception as exc:
            self.repository.update_run("filter", run_id, status="failed", progress=100, error=str(exc))

    def _load_result(self, run: dict[str, Any]) -> dict[str, Any]:
        with self._cache_lock:
            cached = self._result_cache.get(run["id"])
        if cached:
            return cached
        if not run.get("resultPath"):
            raise ValueError("任务尚未生成结果")
        result = self.repository.read_result(run["resultPath"])
        with self._cache_lock:
            self._result_cache[run["id"]] = result
        return result

    def get_filter_run(self, run_id: str) -> dict[str, Any] | None:
        run = self.repository.get_run("filter", run_id)
        if not run:
            return None
        return {key: value for key, value in run.items() if key not in {"resultPath"}}

    def list_filter_runs(self, file_id: str) -> list[dict[str, Any]]:
        if not self.repository.get_file(file_id):
            raise KeyError("文件不存在或已过期")
        return [
            {key: value for key, value in run.items() if key != "resultPath"}
            for run in self.repository.list_runs("filter", file_id)
            if run.get("saved", False)
        ]

    def get_temporary_filter_run(self, file_id: str) -> dict[str, Any] | None:
        if not self.repository.get_file(file_id):
            raise KeyError("文件不存在或已过期")
        run = next(
            (item for item in self.repository.list_runs("filter", file_id) if not item.get("saved", False)),
            None,
        )
        if not run:
            return None
        return {key: value for key, value in run.items() if key != "resultPath"}

    def save_filter_run(self, run_id: str) -> dict[str, Any]:
        run = self.repository.get_run("filter", run_id)
        if not run:
            raise KeyError("筛选任务不存在")
        if run["status"] != "succeeded":
            raise ValueError("只有成功完成的筛选结果可以保存")
        saved = self.repository.update_run("filter", run_id, saved=True)
        return {key: value for key, value in saved.items() if key != "resultPath"}

    def delete_filter_run(self, run_id: str) -> bool:
        run = self.repository.get_run("filter", run_id)
        if not run:
            return False
        if run["status"] in {"queued", "running"}:
            raise ValueError("正在执行的筛选记录不能删除")
        with self._cache_lock:
            self._result_cache.pop(run_id, None)
        return self.repository.delete_run("filter", run_id)

    def preview_filter_result(self, run_id: str, sheet_key: str, page: int, page_size: int) -> dict[str, Any]:
        run = self.repository.get_run("filter", run_id)
        if not run:
            raise KeyError("筛选任务不存在")
        if run["status"] != "succeeded":
            raise ValueError("筛选任务尚未完成")
        result = self._load_result(run)
        dataset = result["sheets"].get(sheet_key)
        if not dataset:
            raise KeyError("结果 Sheet 不存在")
        start = (page - 1) * page_size
        end = start + page_size
        return {
            "key": sheet_key,
            "name": dataset["name"],
            "headers": dataset["headers"],
            "rows": dataset["rows"][start:end],
            "page": page,
            "pageSize": page_size,
            "total": dataset["rowCount"],
            "isSystem": dataset["isSystem"],
        }

    @staticmethod
    def _safe_stem(value: str) -> str:
        stem = Path(value).stem
        return re.sub(r"[\\/:*?\"<>|]", "_", stem)[:80] or "Excel"

    def export_filter(self, run_id: str) -> dict[str, Any]:
        run = self.repository.get_run("filter", run_id)
        if not run or run["status"] != "succeeded":
            raise ValueError("筛选任务尚未成功完成")
        source = self.repository.get_file(run["fileId"])
        if not source:
            raise KeyError("源文件不存在或已过期")
        filename = f"{self._safe_stem(source['name'])}_筛选结果.xlsx"
        destination = self.repository.downloads_dir / f"{run_id}.xlsx"
        with self._export_lock:
            cached = self.repository.get_download_for_run(run_id, "filter")
            if cached:
                return cached
            temporary = destination.with_suffix(".tmp.xlsx")
            temporary.unlink(missing_ok=True)
            try:
                export_filter_workbook(self._load_result(run), temporary)
                temporary.replace(destination)
            finally:
                temporary.unlink(missing_ok=True)
            return self.repository.register_download(run_id, "filter", destination, filename)

    def _filter_result_workbook(self, run: dict[str, Any]) -> WorkbookData:
        result = self._load_result(run)
        sheets = []
        for index, key in enumerate(result["sheetOrder"], start=1):
            dataset = result["sheets"][key]
            sheets.append(_build_sheet(
                f"result-sheet-{index}",
                dataset["name"],
                [dataset["headers"], *dataset["rows"]],
            ))
        return WorkbookData(sheets)

    @staticmethod
    def _filter_result_metadata(run: dict[str, Any]) -> dict[str, Any]:
        summary = run.get("summary") or []
        workbook = WorkbookData([
            _build_sheet(f"result-sheet-{index}", item["name"], [item.get("headers") or ["列A"]])
            for index, item in enumerate(summary, start=1)
        ])
        metadata = workbook.metadata()
        total_rows = 0
        for sheet_meta, item in zip(metadata["sheets"], summary):
            row_count = int(item.get("rowCount") or 0)
            sheet_meta["rowCount"] = row_count
            total_rows += row_count
        metadata["totalRows"] = total_rows
        return metadata

    def list_sum_sources(self) -> list[dict[str, Any]]:
        sources: list[dict[str, Any]] = []
        files = self.repository.list_files()
        for record in files:
            sources.append({
                "id": record["id"],
                "type": "FILE",
                "name": record["name"],
                "sourceFileName": record["name"],
                "fileId": record["id"],
                "saved": True,
                "createdAt": record["createdAt"],
                "metadata": record["metadata"],
            })
            for run in self.repository.list_runs("filter", record["id"]):
                if run.get("status") != "succeeded" or not run.get("resultPath"):
                    continue
                sources.append({
                    "id": run["id"],
                    "type": "FILTER_RESULT",
                    "name": f"{Path(record['name']).stem} · 筛选结果",
                    "sourceFileName": record["name"],
                    "fileId": record["id"],
                    "saved": bool(run.get("saved", False)),
                    "createdAt": run["createdAt"],
                    "metadata": self._filter_result_metadata(run),
                })
        return sorted(sources, key=lambda item: item["createdAt"], reverse=True)

    def _validate_sum_payload(self, payload: dict[str, Any]) -> None:
        selector = payload["columnSelector"]
        if selector["matchMode"] == "NORMALIZED_HEADER" and not selector.get("headerName"):
            raise ValueError("按表头名称求和时必须选择表头")
        if selector["matchMode"] == "FIXED_POSITION" and not selector.get("columnIndex"):
            raise ValueError("按固定位置求和时必须选择列序号")

    def create_sum_run_from_source(self, payload: dict[str, Any]) -> dict[str, Any]:
        source_ref = payload.get("sourceRef")
        if not source_ref:
            raise ValueError("请选择需要求和的数据表")
        if source_ref["type"] == "FILE":
            if not self.repository.get_file(source_ref["id"]):
                raise KeyError("选择的原始工作簿不存在")
            file_id = source_ref["id"]
        else:
            source_run = self.repository.get_run("filter", source_ref["id"])
            if not source_run or source_run.get("status") != "succeeded":
                raise KeyError("选择的筛选结果不存在或尚未完成")
            file_id = source_run["fileId"]
        self._validate_sum_payload(payload)
        return self.repository.create_run("sum", file_id, payload)

    def create_sum_run(self, file_id: str, payload: dict[str, Any]) -> dict[str, Any]:
        if not self.repository.get_file(file_id):
            raise KeyError("文件不存在或已过期")
        payload["sourceRef"] = {"type": "FILE", "id": file_id}
        self._validate_sum_payload(payload)
        return self.repository.create_run("sum", file_id, payload)

    def execute_sum_run(self, run_id: str) -> None:
        try:
            run = self.repository.get_run("sum", run_id)
            if not run:
                return
            self.repository.update_run("sum", run_id, status="running", progress=15)
            source_ref = run["payload"].get("sourceRef") or {"type": "FILE", "id": run["fileId"]}
            if source_ref["type"] == "FILTER_RESULT":
                source_run = self.repository.get_run("filter", source_ref["id"])
                if not source_run or source_run.get("status") != "succeeded":
                    raise ValueError("筛选结果已不存在或尚未完成")
                workbook = self._filter_result_workbook(source_run)
            else:
                workbook = self.load_workbook(source_ref["id"])
            result = execute_sum(workbook, run["payload"]["columnSelector"])
            path = self.repository.write_result(run_id, result)
            with self._cache_lock:
                self._result_cache[run_id] = result
            self.repository.update_run(
                "sum",
                run_id,
                status="succeeded",
                progress=100,
                resultPath=str(path),
                results=result["results"],
                invalidCount=result["invalidCount"],
                grandTotal=result["grandTotal"],
            )
        except Exception as exc:
            self.repository.update_run("sum", run_id, status="failed", progress=100, error=str(exc))

    def get_sum_run(self, run_id: str) -> dict[str, Any] | None:
        run = self.repository.get_run("sum", run_id)
        if not run:
            return None
        return {key: value for key, value in run.items() if key != "resultPath"}

    def get_sum_invalid_details(self, run_id: str) -> list[dict[str, Any]]:
        run = self.repository.get_run("sum", run_id)
        if not run or run["status"] != "succeeded":
            raise ValueError("求和任务尚未成功完成")
        return self._load_result(run)["invalidDetails"]

    def export_sum(self, run_id: str) -> dict[str, Any]:
        run = self.repository.get_run("sum", run_id)
        if not run or run["status"] != "succeeded":
            raise ValueError("求和任务尚未成功完成")
        source = self.repository.get_file(run["fileId"])
        if not source:
            raise KeyError("源文件不存在或已过期")
        source_ref = run["payload"].get("sourceRef") or {"type": "FILE", "id": run["fileId"]}
        suffix = "_筛选结果_跨Sheet求和.xlsx" if source_ref["type"] == "FILTER_RESULT" else "_跨Sheet求和.xlsx"
        filename = f"{self._safe_stem(source['name'])}{suffix}"
        destination = self.repository.downloads_dir / f"{run_id}.xlsx"
        with self._export_lock:
            cached = self.repository.get_download_for_run(run_id, "sum")
            if cached:
                return cached
            temporary = destination.with_suffix(".tmp.xlsx")
            temporary.unlink(missing_ok=True)
            try:
                export_sum_workbook(self._load_result(run), temporary)
                temporary.replace(destination)
            finally:
                temporary.unlink(missing_ok=True)
            return self.repository.register_download(run_id, "sum", destination, filename)

    def delete_file(self, file_id: str) -> bool:
        with self._cache_lock:
            self._workbook_cache.pop(file_id, None)
        return self.repository.delete_file(file_id)
