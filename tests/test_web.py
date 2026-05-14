from __future__ import annotations

import io
import zipfile

from fastapi.testclient import TestClient
from openpyxl import Workbook

from talent_ninebox.web import app as web_app
from talent_ninebox.web.app import app


def test_login_required() -> None:
    client = TestClient(app)
    response = client.get("/", follow_redirects=False)
    assert response.status_code == 200
    assert "进入" in response.text


def test_health_check() -> None:
    client = TestClient(app)
    response = client.get("/health")
    assert response.status_code == 200
    assert response.json() == {"status": "ok"}


def test_login_success(monkeypatch) -> None:
    monkeypatch.setenv("APP_ACCESS_PASSWORD", "secret")
    client = TestClient(app)
    response = client.post("/login", data={"password": "secret"}, follow_redirects=False)
    assert response.status_code == 200
    assert "初版整合" in response.text


def test_download_falls_back_to_tmp_file(monkeypatch, tmp_path) -> None:
    monkeypatch.setenv("APP_ACCESS_PASSWORD", "secret")
    monkeypatch.setattr(web_app, "TMP_ROOT", tmp_path)
    web_app.TASKS.clear()

    task_id = "a" * 32
    output_dir = tmp_path / task_id / "output"
    output_dir.mkdir(parents=True)
    output_file = output_dir / "人才盘点整合结果_20260512.xlsx"
    workbook = Workbook()
    workbook.active["A1"] = "ok"
    workbook.save(output_file)

    client = TestClient(app)
    client.post("/login", data={"password": "secret"})
    response = client.get(f"/download/{task_id}")

    assert response.status_code == 200
    assert response.headers["content-type"] == "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"


def test_process_file_hides_unexpected_error_details(monkeypatch, tmp_path) -> None:
    monkeypatch.setenv("APP_ACCESS_PASSWORD", "secret")
    monkeypatch.setattr(web_app, "TMP_ROOT", tmp_path)
    web_app.TASKS.clear()

    async def fail_processing(mode, file):
        raise RuntimeError("sensitive workbook path and stack details")

    monkeypatch.setattr(web_app, "_run_processing", fail_processing)

    client = TestClient(app)
    client.post("/login", data={"password": "secret"})
    response = client.post(
        "/process-file",
        data={"mode": "initial"},
        files={"file": ("upload.zip", io.BytesIO(b"not a real zip"), "application/zip")},
    )

    assert response.status_code == 400
    assert "sensitive" not in response.text
    assert response.json()["error"].startswith("处理失败")


def test_process_file_returns_guidance_for_wrong_mode_upload(monkeypatch) -> None:
    monkeypatch.setenv("APP_ACCESS_PASSWORD", "secret")
    client = TestClient(app)
    client.post("/login", data={"password": "secret"})
    response = client.post(
        "/tasks",
        data={"mode": "final"},
        files={"file": ("upload.zip", io.BytesIO(b"fake"), "application/zip")},
    )

    assert response.status_code == 400
    payload = response.json()
    assert payload["error"] == "终版生成仅支持上传已整合后的 .xlsx 文件。"
    assert any(".xlsx" in item for item in payload["guidance"])


def test_task_flow_returns_download_url(monkeypatch, tmp_path) -> None:
    monkeypatch.setenv("APP_ACCESS_PASSWORD", "secret")
    monkeypatch.setattr(web_app, "TMP_ROOT", tmp_path)
    web_app.TASKS.clear()

    workbook_path = tmp_path / "dept.xlsx"
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "盘点"
    sheet.append(["工号", "姓名", "初始九宫格落位"])
    sheet.append(["001", "张三", "高潜高绩"])
    workbook.save(workbook_path)

    zip_bytes = io.BytesIO()
    with zipfile.ZipFile(zip_bytes, "w") as archive:
        archive.write(workbook_path, "dept.xlsx")
    zip_bytes.seek(0)

    client = TestClient(app)
    client.post("/login", data={"password": "secret"})
    response = client.post(
        "/tasks",
        data={"mode": "initial"},
        files={"file": ("upload.zip", zip_bytes, "application/zip")},
    )

    assert response.status_code == 202
    task_id = response.json()["task_id"]
    status_response = client.get(f"/tasks/{task_id}")
    payload = status_response.json()
    assert payload["status"] == "done"
    assert payload["download_url"] == f"/download/{task_id}"
    assert payload["summary"]["总人员数"] == 1
