"""Filesystem storage: opaque ids, atomic writes. User input is never used as a path."""
from __future__ import annotations

import hashlib
import json
import os
import re
import secrets
import tempfile
import time
from pathlib import Path
from typing import Any

from . import config

ID_RE = re.compile(r"^[a-z]{2,4}_[0-9a-f]{16}$")


def new_id(prefix: str) -> str:
    return f"{prefix}_{secrets.token_hex(8)}"


def check_id(value: str, prefix: str) -> str:
    if not ID_RE.match(value or "") or not value.startswith(prefix + "_"):
        raise KeyError(value)
    return value


def now() -> float:
    return round(time.time(), 3)


def sha256(data: bytes) -> str:
    return hashlib.sha256(data).hexdigest()


def atomic_write_bytes(path: Path, data: bytes) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    fd, tmp = tempfile.mkstemp(dir=path.parent, suffix=".tmp")
    try:
        with os.fdopen(fd, "wb") as f:
            f.write(data)
            f.flush()
            os.fsync(f.fileno())
        os.replace(tmp, path)
    except BaseException:
        Path(tmp).unlink(missing_ok=True)
        raise


def atomic_write_json(path: Path, obj: Any) -> None:
    atomic_write_bytes(path, json.dumps(obj, ensure_ascii=False, indent=1, default=str).encode("utf-8"))


def read_json(path: Path, default: Any = None) -> Any:
    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except FileNotFoundError:
        return default


def append_jsonl(path: Path, obj: Any) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    with open(path, "a", encoding="utf-8") as f:
        f.write(json.dumps(obj, ensure_ascii=False, default=str) + "\n")


# ---------------------------------------------------------------- layout
def uploads_dir() -> Path:
    return config.data_dir() / "uploads"


def workbooks_dir() -> Path:
    return config.data_dir() / "workbooks"


def workspaces_dir() -> Path:
    return config.data_dir() / "workspaces"


def templates_dir() -> Path:
    return config.data_dir() / "templates"
