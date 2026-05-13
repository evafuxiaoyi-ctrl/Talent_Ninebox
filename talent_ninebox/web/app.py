from __future__ import annotations

import os
import base64
import json
import shutil
import string
import uuid
from datetime import datetime, timedelta
from pathlib import Path

from fastapi import FastAPI, File, Form, HTTPException, Request, UploadFile
from fastapi.responses import FileResponse, HTMLResponse, JSONResponse, RedirectResponse
from fastapi.staticfiles import StaticFiles
from fastapi.templating import Jinja2Templates
from starlette.middleware.sessions import SessionMiddleware

from talent_ninebox.core.models import ProcessOptions
from talent_ninebox.core.processor import process_workbook, process_zip

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

TASKS: dict[str, dict[str, object]] = {}
HEX_DIGITS = set(string.hexdigits)


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
    candidates = sorted(output_dir.glob("*.xlsx"), key=lambda path: path.stat().st_mtime, reverse=True)
    return candidates[0] if candidates else None


def _validate_upload(mode: str, filename: str) -> str | None:
    if mode not in {"initial", "final"}:
        return "处理类型无效。"
    if mode == "initial" and not filename.lower().endswith(".zip"):
        return "初版整合仅支持上传 .zip 文件。"
    if mode == "final" and not filename.lower().endswith(".xlsx"):
        return "终版生成仅支持上传已整合后的 .xlsx 文件。"
    return None


async def _run_processing(mode: str, file: UploadFile):
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

        if mode == "initial":
            result = process_zip(input_path, output_dir, ProcessOptions(placement_mode="initial", stage_label="初版整合"))
        else:
            result = process_workbook(input_path, output_dir, ProcessOptions(placement_mode="final", stage_label="终版九宫格生成", add_source_columns=False))
        TASKS[task_id] = {
            "created_at": datetime.now(),
            "task_dir": task_dir,
            "output_file": result.output_file,
            "summary": result.summary.as_dict(),
        }
        return task_id, task_dir, result
    except Exception:
        shutil.rmtree(task_dir, ignore_errors=True)
        raise


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
async def process(request: Request, mode: str = Form("initial"), file: UploadFile = File(...)) -> HTMLResponse:
    if not _is_logged_in(request):
        return templates.TemplateResponse(request, "login.html", {"error": "请先登录"})
    _cleanup_old_tasks()

    filename = file.filename or ""
    validation_error = _validate_upload(mode, filename)
    if validation_error:
        return templates.TemplateResponse(request, "index.html", {"result": None, "error": validation_error})

    try:
        task_id, _, result = await _run_processing(mode, file)
        view_model = {
            "task_id": task_id,
            "download_name": result.output_file.name,
            "summary": result.summary.as_dict(),
        }
        return templates.TemplateResponse(request, "index.html", {"result": view_model, "error": ""})
    except Exception as exc:
        return templates.TemplateResponse(request, "index.html", {"result": None, "error": str(exc)})


@app.post("/process-file")
async def process_file(request: Request, mode: str = Form("initial"), file: UploadFile = File(...)) -> FileResponse:
    if not _is_logged_in(request):
        return JSONResponse({"error": "请先登录。"}, status_code=401)
    _cleanup_old_tasks()

    validation_error = _validate_upload(mode, file.filename or "")
    if validation_error:
        return JSONResponse({"error": validation_error}, status_code=400)

    try:
        _, _, result = await _run_processing(mode, file)
    except Exception as exc:
        return JSONResponse({"error": str(exc)}, status_code=400)

    summary_json = json.dumps(result.summary.as_dict(), ensure_ascii=False)
    summary_header = base64.b64encode(summary_json.encode("utf-8")).decode("ascii")
    return FileResponse(
        result.output_file,
        filename=result.output_file.name,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"X-Process-Summary": summary_header},
    )


@app.get("/download/{task_id}")
async def download(request: Request, task_id: str) -> FileResponse:
    if not _is_logged_in(request):
        raise HTTPException(status_code=401, detail="请先登录。")
    _cleanup_old_tasks()
    output_file = _find_output_file(task_id)
    if output_file is None:
        raise HTTPException(status_code=404, detail="结果文件已过期或不存在，请重新上传处理。")
    return FileResponse(output_file, filename=output_file.name, media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
