"""WorkbookService: revisions, 409, idempotency, validation gate, undo, export."""
from __future__ import annotations

import pytest

from workspace import service
from workspace.errors import InvalidOperation, RevisionConflict
from workspace.schemas import OperationBatch


def batch(base, op, *acts, origin="manual"):
    return OperationBatch(base_revision=base, operation_id=op, origin=origin,
                          actions=[{"type": "set_value", "sheet": s, "cell": c, "value": v} for s, c, v in acts])


def test_revision_flow(rich_template):
    meta = service.create_from_path(rich_template, "rich.xlsx", origin="upload")
    wid = meta["id"]
    original = rich_template.read_bytes()
    r = service.apply_operations(wid, batch(1, "op-00001", ("Report", "A3", "Alpha")))
    assert r.revision == 2 and r.validation.passed and r.changed_ranges == ["Report!A3"]
    assert service.revision_path(wid, 1).read_bytes() == original        # rev 1 immutable
    with pytest.raises(RevisionConflict):
        service.apply_operations(wid, batch(1, "op-00002", ("Report", "A4", "Beta")))
    replay = service.apply_operations(wid, batch(1, "op-00001", ("Report", "A3", "Alpha")))
    assert replay.idempotent_replay and replay.revision == 2
    with pytest.raises(InvalidOperation):
        service.apply_operations(wid, batch(2, "op-00001", ("Report", "A3", "Other")))
    service.apply_operations(wid, batch(2, "op-00003", ("Report", "A4", "Beta")))
    u = service.undo(wid, 3)
    assert u.revision == 4
    assert service.revision_path(wid, 4).read_bytes() == service.revision_path(wid, 2).read_bytes()
    u2 = service.undo(wid, 4)                                              # undo again walks back to rev 1
    assert service.revision_path(wid, u2.revision).read_bytes() == original
    assert not service.public_metadata(wid)["can_undo"]
    path, name = service.export_info(wid)
    assert path.read_bytes() == original and name.endswith(".xlsx")
    events = [h["event"] for h in service.history(wid)]
    assert events.count("revision") == 2 and events.count("undo") == 2


def test_invalid_op_publishes_nothing(rich_template):
    wid = service.create_from_path(rich_template, "rich.xlsx", origin="upload")["id"]
    with pytest.raises(InvalidOperation):
        service.apply_operations(wid, batch(1, "op-bad-01", ("Report", "D3", 9)))     # formula cell
    assert service.get_metadata(wid)["current_revision"] == 1


def test_preview_cached_per_revision(simple_sales):
    wid = service.create_from_path(simple_sales, "s.xlsx", origin="upload")["id"]
    p = service.get_preview(wid)
    assert p["revision"] == 1 and p["sheets"][0]["name"] == "Sales"
    cells = {(c["r"], c["c"]): c["v"] for c in p["sheets"][0]["celldata"]}
    assert cells[(5, 3)]["f"] == "=SUM(D3:D5)"
    assert cells[(1, 0)]["bl"] == 1 and cells[(1, 0)]["bg"].upper() == "#1F4E78"
