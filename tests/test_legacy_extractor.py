"""Regression suite for the existing Template Extractor (xlsx_template). Must stay green unchanged.

No test here calls Groq: the LLM decider is replaced by a fake, and use_cache=False keeps the
user's cache/decisions.json untouched.
"""
from __future__ import annotations

import json
import re
import zipfile
from pathlib import Path

import pytest

from tests.conftest import WBS_PLAN
from xlsx_template import decider as dec
from xlsx_template import graph
from xlsx_template.cache import DecisionCache
from xlsx_template.regions import build
from xlsx_template.validate import validate
from xlsx_template.writer import write_template


# ------------------------------------------------------------------ full pipeline, saved plan
def test_plan_apply_on_real_sample_passes_validation(wbs: Path, tmp_path: Path):
    if not WBS_PLAN.exists():
        pytest.skip("saved plan not present")
    out = tmp_path / "out.xlsx"
    state = graph.run(str(wbs), str(out), plan_in=str(WBS_PLAN), use_cache=False)
    v = state["validation"]
    assert v["ok"], v["errors"]
    assert v["checks"]["cells_cleared"] == 1722
    assert v["checks"]["formulas_preserved"] == 1511
    assert v["checks"]["opens_in_openpyxl"] == 1
    assert Path(state["report_path"]).exists()


def test_x14_and_untouched_parts_survive(wbs: Path, tmp_path: Path):
    if not WBS_PLAN.exists():
        pytest.skip("saved plan not present")
    out = tmp_path / "out.xlsx"
    graph.run(str(wbs), str(out), plan_in=str(WBS_PLAN), use_cache=False)
    with zipfile.ZipFile(wbs) as a, zipfile.ZipFile(out) as b:
        assert a.namelist() == b.namelist()
        allowed = re.compile(r"xl/worksheets/(_rels/)?sheet\d+\.xml(\.rels)?$|xl/sharedStrings\.xml$|xl/workbook\.xml$")
        for n in a.namelist():
            if not allowed.match(n):
                assert a.read(n) == b.read(n), f"{n} changed"
        for n in a.namelist():
            if n.startswith("xl/worksheets/sheet"):
                xo, xn = a.read(n).decode(), b.read(n).decode()
                assert re.findall(r"<x14:dataValidation\b.*?</x14:dataValidation>", xo, re.S) == \
                       re.findall(r"<x14:dataValidation\b.*?</x14:dataValidation>", xn, re.S)
        assert 'fullCalcOnLoad="1"' in b.read("xl/workbook.xml").decode()


def test_run_refuses_to_overwrite_input(wbs: Path):
    with pytest.raises(ValueError):
        graph.run(str(wbs), str(wbs), plan_in=str(WBS_PLAN), use_cache=False)


# ------------------------------------------------------------------ decide stage with a fake LLM
class FakeDecider:
    """Clears every 'data' region, keeps labels; mimics LLMDecider.decide()."""

    def __init__(self, *a, **kw):
        self.model = "fake"

    def decide(self, pw, regions=None, on_progress=None):
        regions = pw.regions if regions is None else regions
        res = dec.DecideResult(decisions={}, model="fake")
        for r in regions:
            res.decisions[r.id] = dec.Decision(r.id, "clear" if r.annotation == "data" and r.vtype == "number" else "keep",
                                               "", "llm:fake")
        res.calls = 1
        return res


def test_dry_run_writes_plan_only(simple_sales: Path, tmp_path: Path, monkeypatch):
    monkeypatch.setattr(dec, "LLMDecider", FakeDecider)
    out = tmp_path / "t.xlsx"
    state = graph.run(str(simple_sales), str(out), dry_run=True, use_cache=False)
    assert not out.exists()
    plan = json.loads(Path(state["plan_path"]).read_text(encoding="utf-8"))
    assert plan["regions"] and {r["action"] for r in plan["regions"]} <= {"keep", "clear", "review"}


def test_fake_decider_end_to_end_keeps_formulas(simple_sales: Path, tmp_path: Path, monkeypatch):
    monkeypatch.setattr(dec, "LLMDecider", FakeDecider)
    out = tmp_path / "t.xlsx"
    state = graph.run(str(simple_sales), str(out), use_cache=False)
    assert state["validation"]["ok"], state["validation"]["errors"]
    assert state["validation"]["checks"]["formulas_preserved"] >= 4


def test_plan_roundtrip_reapplies_same_cells(simple_sales: Path, tmp_path: Path, monkeypatch):
    monkeypatch.setattr(dec, "LLMDecider", FakeDecider)
    s1 = graph.run(str(simple_sales), str(tmp_path / "a.xlsx"), use_cache=False)
    s2 = graph.run(str(simple_sales), str(tmp_path / "b.xlsx"), plan_in=s1["plan_path"], use_cache=False)
    assert s1["clear"] == s2["clear"]
    assert s2["validation"]["ok"]


def test_failed_llm_means_review_which_keeps(simple_sales: Path, monkeypatch):
    """KEEP is the fail-safe: no answer -> review -> cell untouched."""
    monkeypatch.setattr(dec.time, "sleep", lambda s: None)
    d = dec.LLMDecider(api_key="test", fallback_model=None, max_retries=1)
    monkeypatch.setattr(d, "_call", lambda *a, **k: (_ for _ in ()).throw(RuntimeError("boom")))
    pw = build(str(simple_sales))
    res = d.decide(pw)
    assert res.decisions and all(x.action == "review" for x in res.decisions.values())


def test_decider_requires_key(monkeypatch):
    monkeypatch.delenv("GROQ_API_KEY", raising=False)
    with pytest.raises(RuntimeError):
        dec.LLMDecider()


# ------------------------------------------------------------------ cache
def test_cache_store_lookup(simple_sales: Path, tmp_path: Path):
    pw = build(str(simple_sales))
    c = DecisionCache(tmp_path / "c.json")
    hits, pending = c.lookup(pw)
    assert not hits and len(pending) == len(pw.regions)
    n = c.store(pw, {r.id: "keep" for r in pw.regions} | {pw.regions[0].id: "review"})
    assert n == len(pw.regions) - 1          # review is never cached
    hits, pending = c.lookup(build(str(simple_sales)))
    assert len(hits) == n and len(pending) == 1


# ------------------------------------------------------------------ writer / validator directly
def test_writer_clear_drops_hyperlink_and_unused_strings(simple_sales: Path, tmp_path: Path):
    out = tmp_path / "o.xlsx"
    clear = {"Sales": {(3, 1), (4, 1), (5, 1), (8, 1)}}    # products + the mailto cell
    st = write_template(str(simple_sales), str(out), clear)
    assert st.cleared == 4 and st.hyperlinks_removed == 1
    assert st.shared_strings_after < st.shared_strings_before
    rep = validate(str(simple_sales), str(out), clear)
    assert rep.ok, rep.errors
    with zipfile.ZipFile(out) as z:
        ws = z.read("xl/worksheets/sheet1.xml").decode()
        assert "<hyperlink" not in ws
        assert "<f>B3*C3</f>" in ws and "<v>300000000</v>" not in ws   # cache stripped


def test_validator_detects_tampering(simple_sales: Path, tmp_path: Path):
    out = tmp_path / "o.xlsx"
    write_template(str(simple_sales), str(out), {})
    # claim a cell should have been cleared that was not -> must fail
    rep = validate(str(simple_sales), str(out), {"Sales": {(3, 1)}})
    assert not rep.ok


def test_x14_fixture_preserved_by_clear(x14_template: Path, tmp_path: Path):
    out = tmp_path / "o.xlsx"
    clear = {"Lists": {(1, 1)}}
    write_template(str(x14_template), str(out), clear)
    assert validate(str(x14_template), str(out), clear).ok
    with zipfile.ZipFile(out) as z:
        assert "<x14:dataValidation " in z.read("xl/worksheets/sheet1.xml").decode()
