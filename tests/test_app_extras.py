"""
test_app_extras.py
Covers the newer app.py features: login rate-limiting, upload size cap,
disk cleanup, data preview, and report history.
"""

import os
import time

import pytest


def test_upload_preview_is_populated(app_client, sample_csv_path):
    with open(sample_csv_path, "rb") as f:
        r = app_client.post(
            "/upload",
            data={"file": (f, "sample.csv"), "send_email": "false"},
            content_type="multipart/form-data",
        )
    job_id = r.get_json()["job_id"]

    data = None
    for _ in range(40):
        r = app_client.get(f"/status/{job_id}")
        data = r.get_json()
        if data.get("preview") or data["status"] in ("done", "error"):
            break
        time.sleep(0.2)

    assert "preview" in data
    assert "detected_columns" in data["preview"]
    assert data["preview"]["row_count"] > 0


def test_reports_history_lists_completed_reports(app_client, sample_csv_path):
    with open(sample_csv_path, "rb") as f:
        r = app_client.post(
            "/upload",
            data={"file": (f, "sample.csv"), "send_email": "false"},
            content_type="multipart/form-data",
        )
    job_id = r.get_json()["job_id"]

    for _ in range(40):
        r = app_client.get(f"/status/{job_id}")
        if r.get_json()["status"] in ("done", "error"):
            break
        time.sleep(0.2)

    r = app_client.get("/reports")
    assert r.status_code == 200
    reports = r.get_json()["reports"]
    assert len(reports) >= 1
    assert reports[0]["available"] is True


def test_report_history_download_works(app_client, sample_csv_path):
    with open(sample_csv_path, "rb") as f:
        r = app_client.post(
            "/upload",
            data={"file": (f, "sample.csv"), "send_email": "false"},
            content_type="multipart/form-data",
        )
    job_id = r.get_json()["job_id"]
    for _ in range(40):
        r = app_client.get(f"/status/{job_id}")
        if r.get_json()["status"] in ("done", "error"):
            break
        time.sleep(0.2)

    report_id = app_client.get("/reports").get_json()["reports"][0]["id"]
    r = app_client.get(f"/reports/{report_id}/download")
    assert r.status_code == 200
    assert len(r.data) > 0


def test_report_history_download_404_for_unknown_id(app_client):
    r = app_client.get("/reports/does-not-exist/download")
    assert r.status_code == 404


def test_upload_size_limit_configured():
    """MAX_CONTENT_LENGTH should be a sane positive number, not left unset."""
    import app as app_module
    assert app_module.app.config["MAX_CONTENT_LENGTH"] > 0


def test_cleanup_removes_old_files_but_keeps_recent_ones(monkeypatch, tmp_path):
    monkeypatch.chdir(tmp_path)
    os.makedirs("output/uploads", exist_ok=True)
    os.makedirs("output/charts", exist_ok=True)

    import importlib
    import app as app_module
    importlib.reload(app_module)

    old_file = tmp_path / "output" / "report_old.pptx"
    old_file.write_text("old")
    new_file = tmp_path / "output" / "report_new.pptx"
    new_file.write_text("new")

    # Backdate the "old" file well past the retention window
    old_time = time.time() - (app_module.FILE_RETENTION_HOURS + 1) * 3600
    os.utime(old_file, (old_time, old_time))

    app_module.cleanup_old_files()

    assert not old_file.exists()
    assert new_file.exists()


def test_login_locks_out_after_too_many_failed_attempts(monkeypatch, tmp_path):
    monkeypatch.chdir(tmp_path)
    monkeypatch.setenv("APP_PASSWORD", "correct-password")
    import importlib
    import app as app_module
    importlib.reload(app_module)
    client = app_module.app.test_client()

    for _ in range(app_module.LOGIN_MAX_ATTEMPTS):
        r = client.post("/login", data={"password": "wrong"})
    assert b"Too many attempts" in r.data

    # Even the correct password should now be rejected until the lockout expires
    r = client.post("/login", data={"password": "correct-password"})
    assert b"Too many attempts" in r.data

    monkeypatch.delenv("APP_PASSWORD", raising=False)
    importlib.reload(app_module)


def test_reports_endpoint_requires_auth_when_locked(monkeypatch, tmp_path):
    monkeypatch.chdir(tmp_path)
    monkeypatch.setenv("APP_PASSWORD", "secret")
    import importlib
    import app as app_module
    importlib.reload(app_module)
    client = app_module.app.test_client()

    r = client.get("/reports")
    assert r.status_code == 401

    monkeypatch.delenv("APP_PASSWORD", raising=False)
    importlib.reload(app_module)
