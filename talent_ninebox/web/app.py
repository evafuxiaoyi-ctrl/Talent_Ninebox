from __future__ import annotations

import os
import base64
import json
import logging
import shutil
import string
import time
import uuid
from datetime import datetime, timedelta
from pathlib import Path

from fastapi import BackgroundTasks, FastAPI, File, Form, HTTPException, Request, UploadFile
from fastapi.responses import FileResponse, HTMLResponse, JSONResponse, RedirectResponse
from fastapi.staticfiles import StaticFiles
from fastapi.templating import Jinja2Templates
from starlette.middleware.sessions import SessionMiddleware
from starlette.concurrency import run_in_threadpool

from talent_ninebox.core.models import ProcessOptions
from talent_ninebox.core.processor import process_workbook, process_zip
from talent_ninebox.core.splitter import list_split_fields, split_workbook_by_field

BASE_DIR = Path(__file__).resolve().parent
PROJECT_ROOT = BASE_DIR.parents[1]
IS_SERVERLESS = bool(os.environ.get("VERCEL") or os.environ.get("FC_FUNCTION_NAME") or os.environ.get("RAILWAY_ENVIRONMENT"))
TMP_ROOT = Path("/tmp/talent-ninebox-web") if IS_SERVERLESS else PROJECT_ROOT / "tmp"
MAX_UPLOAD_BYTES = 100 * 1024 * 1024
SESSION_MAX_AGE = 8 * 60 * 60
RESULT_TTL = timedelta(hours=1)

app = FastAPI(title="人才盘点九宫格工具")
app.add_middleware(
    SessionMiddleware,
    secret_key=os.environ.get("APP_SESSION_SECRET", "dev-session-secret-change-me"),
    max_age=SESSION_MAX_AGE,
    same_site="lax",
)
app.mount("/static", StaticFiles(directory=BASE_DIR / "static"), name="static")
templates = Jinja2Templates(directory=BASE_DIR / "templates")
logger = logging.getLogger("talent_ninebox.web")

TASKS: dict[str, dict[str, object]] = {}
HEX_DIGITS = set(string.hexdigits)


def _safe_error_message(exc: Exception) -> str:
    if isinstance(exc, ValueError):
        return str(exc)
    if isinstance(exc, PermissionError):
        return "文件处理失败：服务暂时无法写入临时文件，请稍后重试。"
    return "处理失败：文件结构可能超出当前工具支持范围。请先检查模板是否一致，或查看源文件是否损坏、加密。"


def _error_guidance(message: str, mode: str | None = None) -> list[str]:
    guidance: list[str] = []
    if mode == "split":
        guidance.append("表格拆分请上传一份 .xlsx 总表，并确认第 3 行包含员工姓名、工号、邮箱、岗位或部门字段。")
    if ".zip" in message or "zip" in message or "压缩包" in message:
        guidance.append("初版整合请上传 .zip 压缩包；zip 内放 1-30 个 .xlsx 文件。")
    if ".xlsx" in message or "Excel" in message:
        guidance.append("终版生成请上传已整合后的 .xlsx 文件，不支持 .xls、加密或损坏文件。")
    if "超过 30" in message:
        guidance.append("请把部门文件拆成多个 zip，每个 zip 内最多 30 个 Excel。")
    if "表头" in message or "模板" in message or "字段" in message:
        if mode == "initial":
            guidance.append("初版整合需要表头包含「姓名/员工姓名」和「初始九宫格落位」。")
        elif mode == "final":
            guidance.append("终版生成需要表头包含「姓名/员工姓名」和「校准后九宫格落位」。")
        else:
            guidance.append("请确认表头在前 20 行内，且各文件字段顺序一致。")
    if "损坏" in message or "加密" in message or "无法读取" in message:
        guidance.append("请用 Excel/WPS 打开源文件确认未加密、未损坏，再另存为 .xlsx 后重试。")
    if not guidance:
        guidance.append("请先确认选择的处理类型和上传文件匹配；详细问题可在成功生成后的「异常报告」Sheet 查看。")
    return guidance


def _error_payload(exc: Exception, mode: str | None = None) -> dict[str, object]:
    message = _safe_error_message(exc)
    return {
        "error": message,
        "guidance": _error_guidance(message, mode),
    }


def _password() -> str:
    return os.environ.get("APP_ACCESS_PASSWORD", "change-me")


def _is_logged_in(request: Request) -> bool:
    return bool(request.session.get("logged_in"))


def _require_login(request: Request) -> None:
    if not _is_logged_in(request):
        raise HTTPException(status_code=303, headers={"Location": "/login"})


def _cleanup_old_tasks() -> None:
    now = datetime.now()
    expired = []
    for task_id, meta in TASKS.items():
        created_at = meta.get("created_at")
        if isinstance(created_at, datetime) and now - created_at > RESULT_TTL:
            expired.append(task_id)
    for task_id in expired:
        task_dir = TASKS[task_id].get("task_dir")
        if isinstance(task_dir, Path):
            shutil.rmtree(task_dir, ignore_errors=True)
        TASKS.pop(task_id, None)

    if TMP_ROOT.exists():
        for task_dir in TMP_ROOT.iterdir():
            if not task_dir.is_dir() or task_dir.name in TASKS:
                continue
            try:
                modified_at = datetime.fromtimestamp(task_dir.stat().st_mtime)
            except OSError:
                continue
            if now - modified_at > RESULT_TTL:
                shutil.rmtree(task_dir, ignore_errors=True)


def _safe_task_dir(task_id: str) -> Path | None:
    if len(task_id) != 32 or any(char not in HEX_DIGITS for char in task_id):
        return None
    return TMP_ROOT / task_id


def _find_output_file(task_id: str) -> Path | None:
    task = TASKS.get(task_id)
    if task:
        output_file = task.get("output_file")
        if isinstance(output_file, Path) and output_file.exists():
            return output_file

    task_dir = _safe_task_dir(task_id)
    if task_dir is None:
        return None
    output_dir = task_dir / "output"
    if not output_dir.exists():
        return None
    candidates = sorted(
        [*output_dir.glob("*.xlsx"), *output_dir.glob("*.zip")],
        key=lambda path: path.stat().st_mtime,
        reverse=True,
    )
    return candidates[0] if candidates else None


def _validate_upload(mode: str, filename: str) -> str | None:
    if mode not in {"split", "initial", "final"}:
        return "处理类型无效。"
    if mode == "split" and not filename.lower().endswith(".xlsx"):
        return "表格拆分仅支持上传一份 .xlsx 总表。"
    if mode == "initial" and not filename.lower().endswith(".zip"):
        return "初版整合仅支持上传 .zip 文件。"
    if mode == "final" and not filename.lower().endswith(".xlsx"):
        return "终版生成仅支持上传已整合后的 .xlsx 文件。"
    return None


async def _run_processing(mode: str, file: UploadFile, split_field: str | None = None):
    started_at = time.monotonic()
    task_id = uuid.uuid4().hex
    task_dir = TMP_ROOT / task_id
    input_dir = task_dir / "input"
    output_dir = task_dir / "output"
    input_dir.mkdir(parents=True, exist_ok=True)
    output_dir.mkdir(parents=True, exist_ok=True)
    input_path = input_dir / ("upload.zip" if mode == "initial" else "upload.xlsx")

    size = 0
    try:
        with input_path.open("wb") as dst:
            while chunk := await file.read(1024 * 1024):
                size += len(chunk)
                if size > MAX_UPLOAD_BYTES:
                    raise ValueError("上传文件超过 100MB 限制。")
                dst.write(chunk)

        if mode == "split":
            if not split_field:
                raise ValueError("请选择拆分依据字段。")
            result = split_workbook_by_field(input_path, output_dir, split_field)
            summary = result.summary
        elif mode == "initial":
            result = process_zip(input_path, output_dir, ProcessOptions(placement_mode="initial", stage_label="初版整合"))
            summary = result.summary.as_dict()
        else:
            result = process_workbook(input_path, output_dir, ProcessOptions(placement_mode="final", stage_label="终版九宫格生成", add_source_columns=False))
            summary = result.summary.as_dict()
        TASKS[task_id] = {
            "created_at": datetime.now(),
            "task_dir": task_dir,
            "output_file": result.output_file,
            "summary": summary,
        }
        duration_ms = int((time.monotonic() - started_at) * 1000)
        logger.info(
            "processing_success task_id=%s mode=%s input_ext=%s upload_bytes=%s duration_ms=%s output_bytes=%s",
            task_id,
            mode,
            input_path.suffix.lower(),
            size,
            duration_ms,
            result.output_file.stat().st_size if result.output_file.exists() else 0,
        )
        return task_id, task_dir, result
    except Exception as exc:
        duration_ms = int((time.monotonic() - started_at) * 1000)
        logger.exception(
            "processing_failed task_id=%s mode=%s input_ext=%s upload_bytes=%s duration_ms=%s error_type=%s",
            task_id,
            mode,
            input_path.suffix.lower(),
            size,
            duration_ms,
            type(exc).__name__,
        )
        shutil.rmtree(task_dir, ignore_errors=True)
        raise


async def _save_upload(task_id: str, mode: str, file: UploadFile) -> tuple[Path, Path, Path, int]:
    task_dir = TMP_ROOT / task_id
    input_dir = task_dir / "input"
    output_dir = task_dir / "output"
    input_dir.mkdir(parents=True, exist_ok=True)
    output_dir.mkdir(parents=True, exist_ok=True)
    input_path = input_dir / ("upload.zip" if mode == "initial" else "upload.xlsx")

    size = 0
    with input_path.open("wb") as dst:
        while chunk := await file.read(1024 * 1024):
            size += len(chunk)
            if size > MAX_UPLOAD_BYTES:
                raise ValueError("上传文件超过 100MB 限制。")
            dst.write(chunk)
    return task_dir, input_path, output_dir, size


async def _process_saved_task(task_id: str, mode: str, input_path: Path, output_dir: Path, upload_size: int, split_field: str | None = None) -> None:
    started_at = time.monotonic()
    task = TASKS.get(task_id)
    if task is None:
        return
    task["status"] = "processing"
    task["message"] = "正在处理 Excel"

    try:
        if mode == "split":
            if not split_field:
                raise ValueError("请选择拆分依据字段。")
            result = await run_in_threadpool(split_workbook_by_field, input_path, output_dir, split_field)
            summary = result.summary
        elif mode == "initial":
            result = await run_in_threadpool(
                process_zip,
                input_path,
                output_dir,
                ProcessOptions(placement_mode="initial", stage_label="初版整合"),
            )
            summary = result.summary.as_dict()
        else:
            result = await run_in_threadpool(
                process_workbook,
                input_path,
                output_dir,
                ProcessOptions(placement_mode="final", stage_label="终版九宫格生成", add_source_columns=False),
            )
            summary = result.summary.as_dict()

        duration_ms = int((time.monotonic() - started_at) * 1000)
        task.update(
            {
                "status": "done",
                "message": "处理完成",
                "completed_at": datetime.now(),
                "output_file": result.output_file,
                "summary": summary,
                "download_name": result.output_file.name,
                "duration_ms": duration_ms,
            }
        )
        logger.info(
            "processing_success task_id=%s mode=%s input_ext=%s upload_bytes=%s duration_ms=%s output_bytes=%s",
            task_id,
            mode,
            input_path.suffix.lower(),
            upload_size,
            duration_ms,
            result.output_file.stat().st_size if result.output_file.exists() else 0,
        )
    except Exception as exc:
        duration_ms = int((time.monotonic() - started_at) * 1000)
        task.update(
            {
                "status": "failed",
                "message": "处理失败",
                "completed_at": datetime.now(),
                "error": _safe_error_message(exc),
                "duration_ms": duration_ms,
            }
        )
        logger.exception(
            "processing_failed task_id=%s mode=%s input_ext=%s upload_bytes=%s duration_ms=%s error_type=%s",
            task_id,
            mode,
            input_path.suffix.lower(),
            upload_size,
            duration_ms,
            type(exc).__name__,
        )


def _task_status_payload(task_id: str, task: dict[str, object]) -> dict[str, object]:
    status = str(task.get("status", "unknown"))
    payload: dict[str, object] = {
        "task_id": task_id,
        "status": status,
        "message": task.get("message", ""),
    }
    if status == "done":
        payload.update(
            {
                "summary": task.get("summary", {}),
                "download_url": f"/download/{task_id}",
                "download_name": task.get("download_name", "人才盘点整合结果.xlsx"),
            }
        )
    if status == "failed":
        error = str(task.get("error", "处理失败，请重新上传处理。"))
        payload["error"] = error
        payload["guidance"] = _error_guidance(error, str(task.get("mode", "")))
    return payload


@app.get("/health")
async def health() -> dict[str, str]:
    return {"status": "ok"}


@app.get("/login", response_class=HTMLResponse)
async def login_page(request: Request) -> HTMLResponse:
    if _is_logged_in(request):
        _cleanup_old_tasks()
        return templates.TemplateResponse(request, "index.html", {"result": None, "error": ""})
    return templates.TemplateResponse(request, "login.html", {"error": ""})


@app.post("/login", response_class=HTMLResponse)
async def login(request: Request, password: str = Form(...)) -> HTMLResponse:
    if password == _password():
        request.session["logged_in"] = True
        _cleanup_old_tasks()
        return templates.TemplateResponse(request, "index.html", {"result": None, "error": ""})
    return templates.TemplateResponse(request, "login.html", {"error": "密码错误"})


@app.post("/logout")
async def logout(request: Request) -> HTMLResponse:
    request.session.clear()
    return templates.TemplateResponse(request, "login.html", {"error": ""})


@app.get("/", response_class=HTMLResponse)
async def index(request: Request) -> HTMLResponse:
    if not _is_logged_in(request):
        return templates.TemplateResponse(request, "login.html", {"error": ""})
    _cleanup_old_tasks()
    return templates.TemplateResponse(request, "index.html", {"result": None, "error": ""})


@app.post("/process", response_class=HTMLResponse)
async def process(request: Request, mode: str = Form("initial"), split_field: str | None = Form(None), file: UploadFile = File(...)) -> HTMLResponse:
    if not _is_logged_in(request):
        return templates.TemplateResponse(request, "login.html", {"error": "请先登录"})
    _cleanup_old_tasks()

    filename = file.filename or ""
    validation_error = _validate_upload(mode, filename)
    if validation_error:
        return templates.TemplateResponse(request, "index.html", {"result": None, "error": validation_error})
    if mode == "split" and not split_field:
        return templates.TemplateResponse(request, "index.html", {"result": None, "error": "请选择拆分依据字段。"})

    try:
        task_id, _, result = await _run_processing(mode, file, split_field)
        summary = result.summary if mode == "split" else result.summary.as_dict()
        view_model = {
            "task_id": task_id,
            "download_name": result.output_file.name,
            "summary": summary,
        }
        return templates.TemplateResponse(request, "index.html", {"result": view_model, "error": ""})
    except Exception as exc:
        return templates.TemplateResponse(request, "index.html", {"result": None, "error": _safe_error_message(exc)})


@app.post("/process-file")
async def process_file(request: Request, mode: str = Form("initial"), split_field: str | None = Form(None), file: UploadFile = File(...)) -> FileResponse:
    if not _is_logged_in(request):
        return JSONResponse({"error": "请先登录。"}, status_code=401)
    _cleanup_old_tasks()

    validation_error = _validate_upload(mode, file.filename or "")
    if validation_error:
        return JSONResponse({"error": validation_error, "guidance": _error_guidance(validation_error, mode)}, status_code=400)
    if mode == "split" and not split_field:
        message = "请选择拆分依据字段。"
        return JSONResponse({"error": message, "guidance": _error_guidance(message, mode)}, status_code=400)

    try:
        _, _, result = await _run_processing(mode, file, split_field)
    except Exception as exc:
        return JSONResponse(_error_payload(exc, mode), status_code=400)

    summary = result.summary if mode == "split" else result.summary.as_dict()
    summary_json = json.dumps(summary, ensure_ascii=False)
    summary_header = base64.b64encode(summary_json.encode("utf-8")).decode("ascii")
    media_type = "application/zip" if result.output_file.suffix.lower() == ".zip" else "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    return FileResponse(
        result.output_file,
        filename=result.output_file.name,
        media_type=media_type,
        headers={"X-Process-Summary": summary_header},
    )


@app.post("/split-fields")
async def split_fields(request: Request, file: UploadFile = File(...)) -> JSONResponse:
    if not _is_logged_in(request):
        return JSONResponse({"error": "请先登录。"}, status_code=401)
    _cleanup_old_tasks()

    validation_error = _validate_upload("split", file.filename or "")
    if validation_error:
        return JSONResponse({"error": validation_error, "guidance": _error_guidance(validation_error, "split")}, status_code=400)

    task_id = uuid.uuid4().hex
    task_dir = TMP_ROOT / task_id
    input_dir = task_dir / "input"
    input_dir.mkdir(parents=True, exist_ok=True)
    input_path = input_dir / "upload.xlsx"
    size = 0
    try:
        with input_path.open("wb") as dst:
            while chunk := await file.read(1024 * 1024):
                size += len(chunk)
                if size > MAX_UPLOAD_BYTES:
                    raise ValueError("上传文件超过 100MB 限制。")
                dst.write(chunk)
        fields = await run_in_threadpool(list_split_fields, input_path)
        return JSONResponse(
            {
                "fields": [
                    {"name": field.name, "column": field.column_letter}
                    for field in fields
                ]
            }
        )
    except Exception as exc:
        return JSONResponse(_error_payload(exc, "split"), status_code=400)
    finally:
        shutil.rmtree(task_dir, ignore_errors=True)


@app.post("/tasks")
async def create_task(
    background_tasks: BackgroundTasks,
    request: Request,
    mode: str = Form("initial"),
    split_field: str | None = Form(None),
    file: UploadFile = File(...),
) -> JSONResponse:
    if not _is_logged_in(request):
        return JSONResponse({"error": "请先登录。"}, status_code=401)
    _cleanup_old_tasks()

    validation_error = _validate_upload(mode, file.filename or "")
    if validation_error:
        return JSONResponse({"error": validation_error, "guidance": _error_guidance(validation_error, mode)}, status_code=400)
    if mode == "split" and not split_field:
        message = "请选择拆分依据字段。"
        return JSONResponse({"error": message, "guidance": _error_guidance(message, mode)}, status_code=400)

    task_id = uuid.uuid4().hex
    try:
        task_dir, input_path, output_dir, upload_size = await _save_upload(task_id, mode, file)
    except Exception as exc:
        task_dir = TMP_ROOT / task_id
        shutil.rmtree(task_dir, ignore_errors=True)
        return JSONResponse(_error_payload(exc, mode), status_code=400)

    TASKS[task_id] = {
        "created_at": datetime.now(),
        "task_dir": task_dir,
        "status": "queued",
        "message": "文件已上传，等待处理",
        "mode": mode,
        "split_field": split_field,
        "upload_size": upload_size,
    }
    background_tasks.add_task(_process_saved_task, task_id, mode, input_path, output_dir, upload_size, split_field)
    return JSONResponse(_task_status_payload(task_id, TASKS[task_id]), status_code=202)


@app.get("/tasks/{task_id}")
async def task_status(request: Request, task_id: str) -> JSONResponse:
    if not _is_logged_in(request):
        return JSONResponse({"error": "请先登录。"}, status_code=401)
    _cleanup_old_tasks()
    if _safe_task_dir(task_id) is None:
        return JSONResponse({"error": "任务不存在或已过期。"}, status_code=404)
    task = TASKS.get(task_id)
    if task is None:
        output_file = _find_output_file(task_id)
        if output_file is None:
            return JSONResponse({"error": "任务不存在或已过期。"}, status_code=404)
        task = {
            "status": "done",
            "message": "处理完成",
            "download_name": output_file.name,
            "summary": {},
        }
    return JSONResponse(_task_status_payload(task_id, task))


@app.get("/download/{task_id}")
async def download(request: Request, task_id: str) -> FileResponse:
    if not _is_logged_in(request):
        raise HTTPException(status_code=401, detail="请先登录。")
    _cleanup_old_tasks()
    output_file = _find_output_file(task_id)
    if output_file is None:
        raise HTTPException(status_code=404, detail="结果文件已过期或不存在，请重新上传处理。")
    media_type = "application/zip" if output_file.suffix.lower() == ".zip" else "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    return FileResponse(output_file, filename=output_file.name, media_type=media_type)
