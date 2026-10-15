"""FortuneSheet adapter: values, formulas, styles, merges, hidden sheets, warnings."""
from __future__ import annotations

from workspace.preview import build_preview


def test_rich_template_preview(rich_template):
    p = build_preview(str(rich_template))
    names = [(s["name"], s.get("hide")) for s in p["sheets"]]
    assert names == [("Report", None), ("Summary", None), ("Lists", 1)]
    rep = p["sheets"][0]
    assert "0_0" in rep["config"]["merge"]
    assert any(c["v"].get("mc", {}).get("r") == 0 and (c["r"], c["c"]) == (0, 1) for c in rep["celldata"])
    assert rep.get("frozen")
    assert any("conditional formatting" in w for w in p["warnings"])
    assert any("comments" in w for w in p["warnings"])


def test_x14_warning(x14_template):
    assert any("x14" in w for w in build_preview(str(x14_template))["warnings"])


def test_number_and_date_display(simple_sales):
    p = build_preview(str(simple_sales))
    cells = {(c["r"], c["c"]): c["v"] for c in p["sheets"][0]["celldata"]}
    assert cells[(2, 2)]["m"] == "15,000,000"
    assert cells[(0, 1)]["ct"]["t"] == "d" and cells[(0, 1)]["m"] == "2026-09-30"


def test_real_sample_preview(wbs):
    p = build_preview(str(wbs))
    assert [s["name"] for s in p["sheets"]][:2] == ["WBS", "High Level Plan"]
    assert sum(len(s["celldata"]) for s in p["sheets"]) > 1000
