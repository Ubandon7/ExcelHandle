from __future__ import annotations

import json
import os
import shutil
import threading
import uuid
from datetime import datetime, timezone
from pathlib import Path
from typing import Any


def utc_now() -> str:
    return datetime.now(timezone.utc).isoformat()


class Repository:
    def __init__(self, root: Path | None = None) -> None:
        configured = os.getenv("EXCEL_HANDLE_DATA_DIR")
        self.root = (root or (Path(configured) if configured else Path(__file__).resolve().parents[1] / "data")).resolve()
        self.files_dir = self.root / "files"
        self.results_dir = self.root / "results"
        self.downloads_dir = self.root / "downloads"
        for directory in (self.root, self.files_dir, self.results_dir, self.downloads_dir):
            directory.mkdir(parents=True, exist_ok=True)
        self.state_path = self.root / "state.json"
        self._lock = threading.RLock()
        self.state = self._load_state()

    def _load_state(self) -> dict[str, Any]:
        empty = {"files": {}, "groups": {}, "filterRuns": {}, "sumRuns": {}, "downloads": {}}
        if not self.state_path.exists():
            return empty
        try:
            data = json.loads(self.state_path.read_text(encoding="utf-8"))
            for key, value in empty.items():
                data.setdefault(key, value)
            # Before saved presets were introduced, every persisted group represented a
            # user-saved condition. Preserve those groups in the new preset library.
            for group in data["groups"].values():
                if "saved" not in group:
                    group["saved"] = True
                    group["savedAt"] = group.get("updatedAt") or group.get("createdAt")
                group.setdefault("active", True)
                if "sourceFileName" not in group:
                    source = data["files"].get(group.get("fileId"))
                    group["sourceFileName"] = source.get("name") if source else None
            return data
        except (OSError, json.JSONDecodeError):
            return empty

    def _save(self) -> None:
        temporary = self.state_path.with_suffix(".tmp")
        temporary.write_text(json.dumps(self.state, ensure_ascii=False, indent=2), encoding="utf-8")
        temporary.replace(self.state_path)

    @staticmethod
    def new_id(prefix: str) -> str:
        return f"{prefix}_{uuid.uuid4().hex}"

    def create_file(self, filename: str, content: bytes, metadata: dict[str, Any]) -> dict[str, Any]:
        extension = Path(filename).suffix.casefold()
        file_id = self.new_id("file")
        stored_path = self.files_dir / f"{file_id}{extension}"
        stored_path.write_bytes(content)
        now = utc_now()
        record = {
            "id": file_id,
            "name": Path(filename).name,
            "size": len(content),
            "format": extension.lstrip("."),
            "status": "ready",
            "path": str(stored_path),
            "metadata": metadata,
            "createdAt": now,
            "lastAccessedAt": now,
        }
        with self._lock:
            self.state["files"][file_id] = record
            for group in self.state["groups"].values():
                if (
                    group.get("saved")
                    and group.get("fileId") is None
                    and str(group.get("sourceFileName") or "").casefold() == record["name"].casefold()
                ):
                    group["fileId"] = file_id
                    group["active"] = False
                    group["updatedAt"] = now
            self._save()
        return record

    def get_file(self, file_id: str, touch: bool = True) -> dict[str, Any] | None:
        with self._lock:
            record = self.state["files"].get(file_id)
            if record and touch:
                record["lastAccessedAt"] = utc_now()
                self._save()
            return dict(record) if record else None

    def list_files(self) -> list[dict[str, Any]]:
        with self._lock:
            records = [dict(record) for record in self.state["files"].values()]
        return sorted(records, key=lambda record: record["createdAt"], reverse=True)

    def list_groups(self, file_id: str) -> list[dict[str, Any]]:
        with self._lock:
            groups = [dict(group) for group in self.state["groups"].values() if group["fileId"] == file_id]
        return sorted(groups, key=lambda group: (group["priority"], group["createdAt"]))

    def list_saved_groups(self) -> list[dict[str, Any]]:
        with self._lock:
            groups = [dict(group) for group in self.state["groups"].values() if group.get("saved", False)]
        return sorted(groups, key=lambda group: group.get("savedAt") or group["updatedAt"], reverse=True)

    def create_group(self, file_id: str, payload: dict[str, Any]) -> dict[str, Any]:
        groups = self.list_groups(file_id)
        priority = len(groups) + 1
        group_id = self.new_id("group")
        now = utc_now()
        record = {
            "id": group_id,
            "fileId": file_id,
            "sourceFileName": self.state["files"].get(file_id, {}).get("name"),
            "name": (payload.get("name") or f"Sheet{priority}").strip(),
            "priority": priority,
            "sourceScope": payload.get("sourceScope") or {"mode": "ALL", "sheetId": None},
            "conditionTree": payload.get("conditionTree") or {"type": "group", "relation": "AND", "children": []},
            "status": "draft",
            "saved": bool(payload.get("saved", False)),
            "active": True,
            "savedAt": now if payload.get("saved", False) else None,
            "createdAt": now,
            "updatedAt": now,
        }
        with self._lock:
            self.state["groups"][group_id] = record
            self._save()
        return dict(record)

    def get_group(self, group_id: str) -> dict[str, Any] | None:
        with self._lock:
            value = self.state["groups"].get(group_id)
            return dict(value) if value else None

    def update_group(self, group_id: str, changes: dict[str, Any]) -> dict[str, Any] | None:
        with self._lock:
            record = self.state["groups"].get(group_id)
            if not record:
                return None
            for key in ("name", "priority", "sourceScope", "conditionTree", "saved", "active"):
                if key in changes and changes[key] is not None:
                    record[key] = changes[key]
            if changes.get("saved") is True:
                record["savedAt"] = utc_now()
            elif changes.get("saved") is False:
                record["savedAt"] = None
            record["status"] = "draft"
            record["updatedAt"] = utc_now()
            self._save()
            return dict(record)

    def delete_group(self, group_id: str) -> bool:
        with self._lock:
            group = self.state["groups"].pop(group_id, None)
            if not group:
                return False
            siblings = [value for value in self.state["groups"].values() if value["fileId"] == group["fileId"]]
            for priority, item in enumerate(sorted(siblings, key=lambda value: value["priority"]), start=1):
                item["priority"] = priority
            self._save()
            return True

    def create_run(self, kind: str, file_id: str, payload: dict[str, Any]) -> dict[str, Any]:
        collection = "filterRuns" if kind == "filter" else "sumRuns"
        run_id = self.new_id("filter_run" if kind == "filter" else "sum_run")
        now = utc_now()
        record = {
            "id": run_id,
            "fileId": file_id,
            "kind": kind,
            "status": "queued",
            "progress": 0,
            "error": None,
            "saved": kind != "filter",
            "payload": payload,
            "resultPath": None,
            "createdAt": now,
            "updatedAt": now,
        }
        with self._lock:
            self.state[collection][run_id] = record
            self._save()
        return dict(record)

    def get_run(self, kind: str, run_id: str) -> dict[str, Any] | None:
        collection = "filterRuns" if kind == "filter" else "sumRuns"
        with self._lock:
            record = self.state[collection].get(run_id)
            return dict(record) if record else None

    def list_runs(self, kind: str, file_id: str) -> list[dict[str, Any]]:
        collection = "filterRuns" if kind == "filter" else "sumRuns"
        with self._lock:
            records = [dict(record) for record in self.state[collection].values() if record["fileId"] == file_id]
        return sorted(records, key=lambda record: record["createdAt"], reverse=True)

    def update_run(self, kind: str, run_id: str, **changes: Any) -> dict[str, Any]:
        collection = "filterRuns" if kind == "filter" else "sumRuns"
        with self._lock:
            record = self.state[collection][run_id]
            record.update(changes)
            record["updatedAt"] = utc_now()
            self._save()
            return dict(record)

    def result_path(self, run_id: str) -> Path:
        return self.results_dir / f"{run_id}.json"

    def write_result(self, run_id: str, value: dict[str, Any]) -> Path:
        path = self.result_path(run_id)
        path.write_text(json.dumps(value, ensure_ascii=False, separators=(",", ":")), encoding="utf-8")
        return path

    @staticmethod
    def read_result(path: str | Path) -> dict[str, Any]:
        return json.loads(Path(path).read_text(encoding="utf-8"))

    def register_download(self, run_id: str, kind: str, path: Path, filename: str) -> dict[str, Any]:
        download_id = self.new_id("download")
        record = {
            "id": download_id,
            "runId": run_id,
            "kind": kind,
            "path": str(path),
            "filename": filename,
            "createdAt": utc_now(),
        }
        with self._lock:
            self.state["downloads"][download_id] = record
            self._save()
        return dict(record)

    def get_download(self, download_id: str) -> dict[str, Any] | None:
        with self._lock:
            record = self.state["downloads"].get(download_id)
            return dict(record) if record else None

    def get_download_for_run(self, run_id: str, kind: str) -> dict[str, Any] | None:
        with self._lock:
            records = [
                dict(record)
                for record in self.state["downloads"].values()
                if record["runId"] == run_id and record["kind"] == kind
            ]
        for record in sorted(records, key=lambda item: item["createdAt"], reverse=True):
            path = Path(record["path"])
            if path.exists() and path.is_file() and path.stat().st_size > 0:
                return record
        return None

    def delete_run(self, kind: str, run_id: str) -> bool:
        collection = "filterRuns" if kind == "filter" else "sumRuns"
        with self._lock:
            run = self.state[collection].pop(run_id, None)
            if not run:
                return False
            paths: list[str] = []
            if run.get("resultPath"):
                paths.append(run["resultPath"])
            download_ids = [
                key for key, value in self.state["downloads"].items() if value["runId"] == run_id
            ]
            for download_id in download_ids:
                download = self.state["downloads"].pop(download_id)
                paths.append(download["path"])
            self._save()

        for raw_path in paths:
            target = Path(raw_path).resolve()
            if target.is_relative_to(self.root) and target.exists() and target.is_file():
                target.unlink()
        return True

    def delete_file(self, file_id: str) -> bool:
        with self._lock:
            record = self.state["files"].pop(file_id, None)
            if not record:
                return False
            group_ids = [key for key, value in self.state["groups"].items() if value.get("fileId") == file_id]
            filter_run_ids = [key for key, value in self.state["filterRuns"].items() if value["fileId"] == file_id]
            sum_run_ids = [key for key, value in self.state["sumRuns"].items() if value["fileId"] == file_id]
            run_ids = set(filter_run_ids + sum_run_ids)
            download_ids = [key for key, value in self.state["downloads"].items() if value["runId"] in run_ids]
            paths = [record["path"]]
            for key in filter_run_ids:
                run = self.state["filterRuns"].pop(key)
                if run.get("resultPath"):
                    paths.append(run["resultPath"])
            for key in sum_run_ids:
                run = self.state["sumRuns"].pop(key)
                if run.get("resultPath"):
                    paths.append(run["resultPath"])
            for key in group_ids:
                group = self.state["groups"][key]
                if group.get("saved"):
                    group["fileId"] = None
                    group["active"] = False
                    group["updatedAt"] = utc_now()
                else:
                    self.state["groups"].pop(key, None)
            for key in download_ids:
                download = self.state["downloads"].pop(key)
                paths.append(download["path"])
            self._save()

        for raw_path in paths:
            target = Path(raw_path).resolve()
            if target.is_relative_to(self.root) and target.exists() and target.is_file():
                target.unlink()
        return True
