"""
test_app.py
Covers app.py: the full upload -> processing -> download flow, recipient
management, and scheduling endpoints. Run against a real Flask test client
(no live server needed).
"""

import time


def test_index_loads(app_client):
    r = app_client.get("/")
    assert r.status_code == 200
    assert b"Business Report Generator" in r.data


def test_upload_and_generate_report_end_to_end(app_client, sample_csv_path):
    with open(sample_csv_path, "rb") as f:
        r = app_client.post(
            "/upload",
            data={"file": (f, "sample.csv"), "send_email": "false"},
            content_type="multipart/form-data",
        )
    assert r.status_code == 200
    job_id = r.get_json()["job_id"]

    data = None
    for _ in range(40):
        r = app_client.get(f"/status/{job_id}")
        data = r.get_json()
        if data["status"] in ("done", "error"):
            break
        time.sleep(0.25)

    assert data["status"] == "done", f"Report generation failed: {data}"
    assert "download_url" in data

    r = app_client.get(data["download_url"])
    assert r.status_code == 200
    assert len(r.data) > 0


def test_upload_rejects_bad_file_type(app_client, tmp_path):
    bad_file = tmp_path / "notes.txt"
    bad_file.write_text("just some text")
    with open(bad_file, "rb") as f:
        r = app_client.post(
            "/upload",
            data={"file": (f, "notes.txt"), "send_email": "false"},
            content_type="multipart/form-data",
        )
    assert r.status_code == 400


def test_recipients_add_and_remove(app_client):
    r = app_client.post("/recipients", json={"action": "add", "email": "test@example.com"})
    assert r.status_code == 200
    assert "test@example.com" in r.get_json()["recipients"]

    r = app_client.post("/recipients", json={"action": "remove", "email": "test@example.com"})
    assert "test@example.com" not in r.get_json()["recipients"]


def test_recipients_rejects_invalid_email(app_client):
    r = app_client.post("/recipients", json={"action": "add", "email": "not-an-email"})
    assert r.status_code == 400


def test_recipients_bulk_add_parses_multiple_formats(app_client):
    text = "alice@x.com, bob@x.com\ncharlie@x.com not-an-email"
    r = app_client.post("/recipients", json={"action": "bulk_add", "text": text})
    data = r.get_json()
    assert set(data["recipients"]) >= {"alice@x.com", "bob@x.com", "charlie@x.com"}
    assert "not-an-email" in data["skipped"]


def test_schedule_get_default(app_client):
    r = app_client.get("/schedule")
    assert r.status_code == 200
    data = r.get_json()
    assert "enabled" in data
    assert "time" in data


def test_schedule_update_and_validate_time_format(app_client):
    r = app_client.post("/schedule", json={"enabled": True, "time": "14:30"})
    assert r.status_code == 200
    assert r.get_json()["time"] == "14:30"

    r = app_client.post("/schedule", json={"time": "not-a-time"})
    assert r.status_code == 400


def test_login_required_when_password_set(monkeypatch, tmp_path):
    import importlib
    monkeypatch.chdir(tmp_path)
    monkeypatch.setenv("APP_PASSWORD", "secret123")
    import app as app_module
    importlib.reload(app_module)
    client = app_module.app.test_client()

    r = client.get("/", follow_redirects=False)
    assert r.status_code == 302
    assert "/login" in r.headers["Location"]

    r = client.post("/login", data={"password": "wrong"})
    assert b"Incorrect password" in r.data

    r = client.post("/login", data={"password": "secret123"}, follow_redirects=True)
    assert b"Business Report Generator" in r.data

    # No-store cache headers should be present so logout can't be bypassed via browser cache
    r = client.get("/")
    assert "no-store" in r.headers.get("Cache-Control", "")

    r = client.get("/logout", follow_redirects=False)
    assert r.status_code == 302

    r = client.get("/", follow_redirects=False)
    assert r.status_code == 302
    assert "/login" in r.headers["Location"]

    monkeypatch.delenv("APP_PASSWORD", raising=False)
    importlib.reload(app_module)
