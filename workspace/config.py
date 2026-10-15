"""Settings read from the environment at call time (tests change them with monkeypatch)."""
from __future__ import annotations

import os
import re
from pathlib import Path

from dotenv import load_dotenv

ROOT = Path(__file__).resolve().parent.parent
load_dotenv(ROOT / ".env")

# hard limits (spec 08/11)
MAX_SHEETS = 20
MAX_ROWS_PER_SHEET = 5000
MAX_COLUMNS = 60
MAX_FORMULAS = 5000
MAX_STRING = 4000
MAX_ACTIONS_PER_BATCH = 1000
MAX_READ_CELLS = 400
MAX_SOURCE_MB = 5
MAX_PREVIEW_CELLS = 250_000
MAX_UNCOMPRESSED_MB = 300
MAX_COMPRESSION_RATIO = 200


def data_dir() -> Path:
    p = Path(os.environ.get("WORKSPACE_DATA_DIR") or ROOT / "data")
    p.mkdir(parents=True, exist_ok=True)
    return p


def max_upload_bytes() -> int:
    return int(float(os.environ.get("WORKSPACE_MAX_UPLOAD_MB", "30")) * 1024 * 1024)


def llm_mode() -> str:
    """'azure' when an Azure OpenAI key + endpoint are configured and not forced to mock, else 'mock'."""
    if os.environ.get("WORKSPACE_LLM", "").strip().lower() == "mock":
        return "mock"
    return "azure" if azure_configured() else "mock"


def _api_base() -> tuple[str, str]:
    """(endpoint, deployment) from API_BASE='https://<res>.openai.azure.com/openai/deployments/<dep>' if set."""
    m = re.match(r"(https://[^/]+)(?:/openai/deployments/([^/?]+))?", os.environ.get("API_BASE", "").strip())
    return (m[1], m[2] or "") if m else ("", "")


def azure_key() -> str:
    return os.environ.get("AZURE_OPENAI_API_KEY") or os.environ.get("OPENAI_API_KEY") or ""


def azure_endpoint() -> str:
    return os.environ.get("AZURE_OPENAI_ENDPOINT") or _api_base()[0]


def azure_configured() -> bool:
    return bool(azure_key() and azure_endpoint())


def azure_deployment() -> str:
    lm = os.environ.get("LANGUAGE_MODEL", "").strip()
    return (os.environ.get("AZURE_OPENAI_DEPLOYMENT") or _api_base()[1] or lm.split("/")[-1] or "gpt-5")


def active_model() -> str | None:
    return f"azure:{azure_deployment()}" if llm_mode() == "azure" else None


def azure_chat(max_tokens: int | None = None, timeout: int = 60, **extra):
    """The Azure OpenAI chat model (gpt-5 family: reasoning tokens count against max_completion_tokens)."""
    from langchain_openai import AzureChatOpenAI
    dep = azure_deployment()
    kw: dict = dict(extra)
    if max_tokens:
        kw["max_completion_tokens"] = max(max_tokens, 16000)
    if dep.startswith(("gpt-5", "o")):
        kw["reasoning_effort"] = os.environ.get("AZURE_OPENAI_REASONING", "low")
    else:
        kw["temperature"] = 0
    return AzureChatOpenAI(azure_deployment=dep, model=dep, api_key=azure_key(), azure_endpoint=azure_endpoint(),
                           api_version=os.environ.get("AZURE_OPENAI_API_VERSION", "2025-01-01-preview"),
                           timeout=max(timeout, 180), max_retries=1, **kw)


def chat_models(max_tokens: int | None = None, timeout: int = 60) -> list[tuple[str, object]]:
    """Chat models to try in order: only Azure OpenAI (empty when it is not configured)."""
    if not azure_configured():
        return []
    return [(f"azure:{azure_deployment()}", azure_chat(max_tokens, timeout))]
