"""HTTP contract: upload -> preview -> operations (409) -> undo -> export; legacy routes still mounted."""
from __future__ import annotations

from fastapi.testclient import TestClient

from workspace.app import create_app

MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"


def client():
    return TestClient(create_app())


def test_upload_preview_edit_export(rich_template):
    c = client()
    assert c.get("/api/health").json()["llm_mode"] == "mock"
    r = c.post("/api/workbooks/upload", files={"file": ("rich.xlsx", rich_template.read_bytes(), MIME)})
    assert r.status_code == 200, r.text
    wid = r.json()["id"]
    p = c.get(f"/api/workbooks/{wid}/preview").json()
    assert p["revision"] == 1 and p["sheets"][0]["name"] == "Report"
    op = {"base_revision": 1, "operation_id": "ui-000001",
          "actions": [{"type": "set_value", "sheet": "Report", "cell": "A3", "value": "Alpha"}]}
    r = c.post(f"/api/workbooks/{wid}/operations", json=op)
    assert r.status_code == 200 and r.json()["revision"] == 2 and r.json()["validation"]["pass"]
    r = c.post(f"/api/workbooks/{wid}/operations", json={**op, "operation_id": "ui-000002"})
    assert r.status_code == 409 and r.json()["code"] == "REVISION_CONFLICT" and r.json()["current_revision"] == 2
    bad = {**op, "base_revision": 2, "operation_id": "ui-000003",
           "actions": [{"type": "set_value", "sheet": "Report", "cell": "D3", "value": 1}]}
    r = c.post(f"/api/workbooks/{wid}/operations", json=bad)
    assert r.status_code == 422 and r.json()["code"] == "INVALID_OPERATION"
    r = c.post(f"/api/workbooks/{wid}/operations",
               json={**op, "operation_id": "ui-000004", "actions": [{"type": "insert_row", "sheet": "Report"}]})
    assert r.status_code == 422
    r = c.get(f"/api/workbooks/{wid}/export")
    assert r.status_code == 200 and r.headers["content-type"].startswith(MIME) and r.content[:2] == b"PK"
    r = c.post(f"/api/workbooks/{wid}/undo", json={"base_revision": 2})
    assert r.json()["revision"] == 3
    assert c.get(f"/api/workbooks/{wid}/export").content == rich_template.read_bytes()
    assert c.get(f"/api/workbooks/{wid}/metadata").json()["current_revision"] == 3


def test_upload_rejections():
    c = client()
    r = c.post("/api/workbooks/upload", files={"file": ("m.xlsm", b"PK\x03\x04", MIME)})
    assert r.status_code == 400 and r.json()["code"] == "UNSUPPORTED_FILE"
    r = c.post("/api/workbooks/upload", files={"file": ("x.xlsx", b"not a zip", MIME)})
    assert r.status_code == 400
    assert c.get("/api/workbooks/wb_0000000000000000/metadata").status_code == 404
    assert c.get("/api/workbooks/..%2F..%2Fetc/metadata").status_code == 404


def test_workspace_files():
    c = client()
    ws = c.post("/api/workspaces", json={"mode": "scratch"}).json()
    r = c.post(f"/api/workspaces/{ws['id']}/files", files={"file": ("notes.txt", b"Revenue: 100", "text/plain")},
               data={"role": "source"})
    assert r.status_code == 200, r.text
    fid = r.json()["id"]
    r = c.post(f"/api/workspaces/{ws['id']}/files", files={"file": ("t.xlsx", b"PK", MIME)}, data={"role": "template"})
    assert r.status_code == 422                       # templates only in template mode
    assert c.delete(f"/api/workspaces/{ws['id']}/files/{fid}").json()["files"] == []


def test_workspace_template_preview(rich_template):
    c = client()
    ws = c.post("/api/workspaces", json={"mode": "template"}).json()["id"]
    original = rich_template.read_bytes()
    tpl = c.post(f"/api/workspaces/{ws}/files", files={"file": ("rich.xlsx", original, MIME)},
                 data={"role": "template"}).json()["id"]
    src = c.post(f"/api/workspaces/{ws}/files", files={"file": ("n.txt", b"x", "text/plain")},
                 data={"role": "source"}).json()["id"]
    p = c.get(f"/api/workspaces/{ws}/files/{tpl}/preview")
    assert p.status_code == 200 and p.json()["sheets"][0]["name"] == "Report"
    assert c.get(f"/api/workspaces/{ws}/files/{src}/preview").status_code == 404     # not an xlsx
    other = c.post("/api/workspaces", json={"mode": "template"}).json()["id"]
    assert c.get(f"/api/workspaces/{other}/files/{tpl}/preview").status_code == 404  # not attached there
    assert rich_template.read_bytes() == original


def test_legacy_routes_still_mounted():
    c = client()
    paths = {getattr(r, "path", "") for r in c.app.router.routes}
    assert {"/api/jobs", "/api/jobs/{job_id}", "/api/jobs/{job_id}/download", "/api/jobs/{job_id}/report"} <= paths
    page = c.get("/extractor")
    assert page.status_code == 200 and "/api/jobs" in page.text
    assert c.get("/api/jobs/doesnotexist").status_code == 404
