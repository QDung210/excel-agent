"""Shared fixtures. Tests never write into the repo's cache/, output/ or jobs/ folders."""
from __future__ import annotations

import shutil
import sys
from pathlib import Path

import pytest

ROOT = Path(__file__).resolve().parent.parent
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from tests import fixtures as fx  # noqa: E402

WBS = ROOT / "Report_WBS (5) (1) 1.xlsx"
WBS_PLAN = ROOT / "output" / "Report_WBS_TEMPLATE_final.plan.json"


@pytest.fixture
def wbs(tmp_path: Path) -> Path:
    """Copy of the user's real sample workbook (has x14 dropdowns, 5 sheets, ~1500 formulas)."""
    if not WBS.exists():
        pytest.skip("sample workbook not present")
    dst = tmp_path / "wbs.xlsx"
    shutil.copyfile(WBS, dst)
    return dst


@pytest.fixture
def simple_sales(tmp_path: Path) -> Path:
    return fx.simple_sales(tmp_path / "simple_sales.xlsx")


@pytest.fixture
def rich_template(tmp_path: Path) -> Path:
    return fx.rich_template(tmp_path / "rich_template.xlsx")


@pytest.fixture
def x14_template(tmp_path: Path) -> Path:
    return fx.x14_validation(tmp_path / "extended_validation.xlsx")


@pytest.fixture(autouse=True)
def _isolate_data(tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> None:
    """Workspace storage goes to a temp dir; LLM keys are removed so nothing calls Azure by accident."""
    monkeypatch.setenv("WORKSPACE_DATA_DIR", str(tmp_path / "data"))
    for k in ("AZURE_OPENAI_API_KEY", "AZURE_OPENAI_ENDPOINT", "OPENAI_API_KEY", "API_BASE"):
        monkeypatch.delenv(k, raising=False)
    monkeypatch.setenv("WORKSPACE_LLM", "mock")
