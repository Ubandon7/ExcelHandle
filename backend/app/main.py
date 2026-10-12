from __future__ import annotations

from pathlib import Path
from typing import Any

from fastapi import BackgroundTasks, FastAPI, File, HTTPException, Query, Request, UploadFile
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse
from fastapi.staticfiles import StaticFiles

from .models import AffixSplitRunRequest, ColumnSplitRunRequest, FilterGroupCreate, FilterGroupUpdate, FilterRunRequest, SumRunRequest
from .services import ApplicationService


app = FastAPI(title="ExcelHandle API", version="0.1.0")
app.add_middleware(
    CORSMiddleware,
    allow_origins=["http://localhost:5173", "http://127.0.0.1:5173"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)
app.state.service = ApplicationService()


def service(request: Request) -> ApplicationService:
    return request.app.state.service


def raise_api_error(exc: Exception) -> None:
    if isinstance(exc, KeyError):
        detail = str(exc).strip("'")
        raise HTTPException(status_code=404, detail=detail) from exc
    if isinstance(exc, (ValueError, RuntimeError)):
        raise HTTPException(status_code=400, detail=str(exc)) from exc
    raise exc


@app.get("/api/health")
def health() -> dict[str, str]:
    return {"status": "ok"}


@app.post("/api/files", status_code=201)
async def upload_file(request: Request, file: UploadFile = File(...)) -> dict[str, Any]:
    try:
        content = await file.read()
        return {"file": service(request).upload(file.filename or "", content)}
    except Exception as exc:
        raise_api_error(exc)


@app.get("/api/files")
def list_files(request: Request) -> dict[str, Any]:
    return {"files": service(request).list_files()}


@app.get("/api/files/{file_id}")
def get_file(request: Request, file_id: str) -> dict[str, Any]:
    record = service(request).get_file(file_id)
    if not record:
        raise HTTPException(status_code=404, detail="文件不存在或已过期")
    return {"file": record}


@app.delete("/api/files/{file_id}", status_code=204)
def delete_file(request: Request, file_id: str) -> None:
    if not service(request).delete_file(file_id):
        raise HTTPException(status_code=404, detail="文件不存在或已过期")


@app.get("/api/files/{file_id}/sheets")
def get_sheets(request: Request, file_id: str) -> dict[str, Any]:
    record = service(request).get_file(file_id)
    if not record:
        raise HTTPException(status_code=404, detail="文件不存在或已过期")
    return record["metadata"]


@app.get("/api/files/{file_id}/sheets/{sheet_id}/preview")
def preview_sheet(
    request: Request,
    file_id: str,
    sheet_id: str,
    page: int = Query(default=1, ge=1),
    page_size: int = Query(default=100, alias="pageSize", ge=1, le=200),
) -> dict[str, Any]:
    try:
        return service(request).preview_source(file_id, sheet_id, page, page_size)
    except Exception as exc:
        raise_api_error(exc)


@app.get("/api/files/{file_id}/affix-suggestions")
def get_affix_suggestions(
    request: Request,
    file_id: str,
    header_name: str = Query(alias="headerName"),
    occurrence: int = Query(default=1, ge=1),
    kind: str = Query(pattern="^(PREFIX|SUFFIX)$"),
    delimiter: str = Query(default="-", min_length=1, max_length=5),
    scope_mode: str = Query(default="ALL", alias="scopeMode", pattern="^(ALL|SINGLE)$"),
    sheet_id: str | None = Query(default=None, alias="sheetId"),
) -> dict[str, Any]:
    try:
        return service(request).affix_suggestions(
            file_id, header_name, occurrence, kind, delimiter, scope_mode, sheet_id
        )
    except Exception as exc:
        raise_api_error(exc)


@app.get("/api/files/{file_id}/filter-groups")
def list_filter_groups(request: Request, file_id: str) -> dict[str, Any]:
    if not service(request).get_file(file_id):
        raise HTTPException(status_code=404, detail="文件不存在或已过期")
    return {"groups": service(request).repository.list_groups(file_id)}


@app.get("/api/filter-groups/saved")
def list_saved_filter_groups(request: Request) -> dict[str, Any]:
    return {"groups": service(request).list_saved_groups()}


@app.post("/api/files/{file_id}/filter-groups", status_code=201)
def create_filter_group(request: Request, file_id: str, payload: FilterGroupCreate) -> dict[str, Any]:
    try:
        group = service(request).create_group(file_id, payload.model_dump(by_alias=True))
        return {"group": group}
    except Exception as exc:
        raise_api_error(exc)


@app.post("/api/files/{file_id}/filter-groups/from-saved/{group_id}", status_code=201)
def create_filter_group_from_saved(request: Request, file_id: str, group_id: str) -> dict[str, Any]:
    try:
        return {"group": service(request).create_group_from_saved(file_id, group_id)}
    except Exception as exc:
        raise_api_error(exc)


@app.delete("/api/files/{file_id}/filter-groups/current")
def clear_current_filter_groups(request: Request, file_id: str) -> dict[str, int]:
    try:
        return {"clearedCount": service(request).clear_current_groups(file_id)}
    except Exception as exc:
        raise_api_error(exc)


@app.patch("/api/filter-groups/{group_id}")
def update_filter_group(request: Request, group_id: str, payload: FilterGroupUpdate) -> dict[str, Any]:
    try:
        changes = payload.model_dump(by_alias=True, exclude_unset=True)
        return {"group": service(request).update_group(group_id, changes)}
    except Exception as exc:
        raise_api_error(exc)


@app.delete("/api/filter-groups/{group_id}", status_code=204)
def delete_filter_group(request: Request, group_id: str) -> None:
    if not service(request).repository.delete_group(group_id):
        raise HTTPException(status_code=404, detail="筛选结果组不存在")


@app.post("/api/files/{file_id}/filter-runs", status_code=202)
def create_filter_run(
    request: Request,
    background_tasks: BackgroundTasks,
    file_id: str,
    payload: FilterRunRequest,
) -> dict[str, Any]:
    try:
        current_service = service(request)
        run = current_service.create_filter_run(file_id, payload.ordered_group_ids)
        background_tasks.add_task(current_service.execute_filter_run, run["id"])
        return {"run": run}
    except Exception as exc:
        raise_api_error(exc)


@app.post("/api/files/{file_id}/column-split-runs", status_code=202)
def create_column_split_run(
    request: Request,
    background_tasks: BackgroundTasks,
    file_id: str,
    payload: ColumnSplitRunRequest,
) -> dict[str, Any]:
    try:
        current_service = service(request)
        run = current_service.create_column_split_run(file_id, payload.model_dump(by_alias=True))
        background_tasks.add_task(current_service.execute_filter_run, run["id"])
        return {"run": run}
    except Exception as exc:
        raise_api_error(exc)


@app.post("/api/files/{file_id}/affix-split-runs", status_code=202)
def create_affix_split_run(
    request: Request,
    background_tasks: BackgroundTasks,
    file_id: str,
    payload: AffixSplitRunRequest,
) -> dict[str, Any]:
    try:
        current_service = service(request)
        run = current_service.create_affix_split_run(file_id, payload.model_dump(by_alias=True))
        background_tasks.add_task(current_service.execute_filter_run, run["id"])
        return {"run": run}
    except Exception as exc:
        raise_api_error(exc)


@app.get("/api/files/{file_id}/filter-runs")
def list_filter_runs(request: Request, file_id: str) -> dict[str, Any]:
    try:
        return {"runs": service(request).list_filter_runs(file_id)}
    except Exception as exc:
        raise_api_error(exc)


@app.get("/api/files/{file_id}/temporary-filter-run")
def get_temporary_filter_run(request: Request, file_id: str) -> dict[str, Any]:
    try:
        return {"run": service(request).get_temporary_filter_run(file_id)}
    except Exception as exc:
        raise_api_error(exc)


@app.get("/api/filter-runs/{run_id}")
def get_filter_run(request: Request, run_id: str) -> dict[str, Any]:
    run = service(request).get_filter_run(run_id)
    if not run:
        raise HTTPException(status_code=404, detail="筛选任务不存在")
    return {"run": run}


@app.delete("/api/filter-runs/{run_id}", status_code=204)
def delete_filter_run(request: Request, run_id: str) -> None:
    try:
        if not service(request).delete_filter_run(run_id):
            raise HTTPException(status_code=404, detail="筛选记录不存在")
    except HTTPException:
        raise
    except Exception as exc:
        raise_api_error(exc)


@app.post("/api/filter-runs/{run_id}/save")
def save_filter_run(request: Request, run_id: str) -> dict[str, Any]:
    try:
        return {"run": service(request).save_filter_run(run_id)}
    except Exception as exc:
        raise_api_error(exc)


@app.get("/api/filter-runs/{run_id}/sheets/{sheet_key}/rows")
def preview_filter_result(
    request: Request,
    run_id: str,
    sheet_key: str,
    page: int = Query(default=1, ge=1),
    page_size: int = Query(default=100, alias="pageSize", ge=1, le=200),
) -> dict[str, Any]:
    try:
        return service(request).preview_filter_result(run_id, sheet_key, page, page_size)
    except Exception as exc:
        raise_api_error(exc)


@app.post("/api/filter-runs/{run_id}/exports")
def create_filter_export(request: Request, run_id: str) -> dict[str, Any]:
    try:
        download = service(request).export_filter(run_id)
        return {"downloadId": download["id"], "downloadUrl": f"/api/downloads/{download['id']}"}
    except Exception as exc:
        raise_api_error(exc)


@app.post("/api/files/{file_id}/sum-runs", status_code=202)
def create_sum_run(
    request: Request,
    background_tasks: BackgroundTasks,
    file_id: str,
    payload: SumRunRequest,
) -> dict[str, Any]:
    try:
        current_service = service(request)
        run = current_service.create_sum_run(file_id, payload.model_dump(by_alias=True))
        background_tasks.add_task(current_service.execute_sum_run, run["id"])
        return {"run": run}
    except Exception as exc:
        raise_api_error(exc)


@app.get("/api/sum-sources")
def list_sum_sources(request: Request) -> dict[str, Any]:
    try:
        return {"sources": service(request).list_sum_sources()}
    except Exception as exc:
        raise_api_error(exc)


@app.post("/api/sum-runs", status_code=202)
def create_sum_run_from_source(
    request: Request,
    background_tasks: BackgroundTasks,
    payload: SumRunRequest,
) -> dict[str, Any]:
    try:
        current_service = service(request)
        run = current_service.create_sum_run_from_source(payload.model_dump(by_alias=True))
        background_tasks.add_task(current_service.execute_sum_run, run["id"])
        return {"run": run}
    except Exception as exc:
        raise_api_error(exc)


@app.get("/api/sum-runs/{run_id}")
def get_sum_run(request: Request, run_id: str) -> dict[str, Any]:
    run = service(request).get_sum_run(run_id)
    if not run:
        raise HTTPException(status_code=404, detail="求和任务不存在")
    return {"run": run}


@app.get("/api/sum-runs/{run_id}/invalid-details")
def get_sum_invalid_details(request: Request, run_id: str) -> dict[str, Any]:
    try:
        return {"items": service(request).get_sum_invalid_details(run_id)}
    except Exception as exc:
        raise_api_error(exc)


@app.post("/api/sum-runs/{run_id}/exports")
def create_sum_export(request: Request, run_id: str) -> dict[str, Any]:
    try:
        download = service(request).export_sum(run_id)
        return {"downloadId": download["id"], "downloadUrl": f"/api/downloads/{download['id']}"}
    except Exception as exc:
        raise_api_error(exc)


@app.get("/api/jobs/{job_id}")
def get_job(request: Request, job_id: str) -> dict[str, Any]:
    current_service = service(request)
    run = current_service.get_filter_run(job_id) or current_service.get_sum_run(job_id)
    if not run:
        raise HTTPException(status_code=404, detail="任务不存在")
    return {"job": run}


@app.get("/api/downloads/{download_id}")
def download_file(request: Request, download_id: str) -> FileResponse:
    record = service(request).repository.get_download(download_id)
    if not record:
        raise HTTPException(status_code=404, detail="下载文件不存在或已过期")
    path = Path(record["path"])
    if not path.exists():
        raise HTTPException(status_code=404, detail="下载文件已经被清理")
    return FileResponse(
        path,
        filename=record["filename"],
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


frontend_dist = Path(__file__).resolve().parents[2] / "frontend" / "dist"
if frontend_dist.exists():
    app.mount("/", StaticFiles(directory=frontend_dist, html=True), name="frontend")
