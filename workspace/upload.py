"""Upload checks (spec 11): extension, signature, zip safety, size. Originals stored immutable."""
from __future__ import annotations

import io
import posixpath
import zipfile
from dataclasses import dataclass
from pathlib import Path

from . import config, storage
from .errors import FileTooLarge, NotFound, UnsupportedFile

XLSX_EXT = {".xlsx"}
SOURCE_EXT = {".txt", ".csv"}
REFUSED_EXT = {".xlsm": "Macro-enabled workbooks (.xlsm) are not supported yet.",
               ".xlsb": "Binary workbooks (.xlsb) are not supported. Save as .xlsx.",
               ".xls": "Legacy .xls is not supported. Save as .xlsx."}
OLE_MAGIC = b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1"   # encrypted .xlsx and legacy .xls are OLE files


@dataclass
class StoredFile:
    id: str
    filename: str
    kind: str            # xlsx | txt | csv
    size: int
    sha256: str
    path: Path

    def public(self) -> dict:
        return {"id": self.id, "filename": self.filename, "kind": self.kind, "size": self.size}


def check_xlsx(data: bytes) -> None:
    if data.startswith(OLE_MAGIC):
        raise UnsupportedFile("This workbook is password-protected (encrypted) or a legacy .xls file.")
    if not data.startswith(b"PK\x03\x04"):
        raise UnsupportedFile("Not a valid .xlsx file (missing ZIP signature).")
    try:
        z = zipfile.ZipFile(io.BytesIO(data))
    except zipfile.BadZipFile:
        raise UnsupportedFile("The .xlsx archive is corrupted.") from None
    with z:
        names = [i.filename for i in z.infolist()]
        if len(names) != len(set(names)):
            raise UnsupportedFile("Archive contains duplicate entries.")
        total = 0
        for info in z.infolist():
            n = info.filename
            if n.startswith(("/", "\\")) or "\\" in n or ".." in posixpath.normpath(n).split("/") or ":" in n:
                raise UnsupportedFile(f"Unsafe path inside archive: {n[:80]}")
            if info.external_attr >> 16 & 0o170000 == 0o120000:
                raise UnsupportedFile("Archive contains a symbolic link.")
            if info.flag_bits & 0x1:
                raise UnsupportedFile("Archive entries are encrypted.")
            total += info.file_size
            if info.file_size > 1_000_000 and info.compress_size and info.file_size / info.compress_size > config.MAX_COMPRESSION_RATIO:
                raise UnsupportedFile("Archive looks like a decompression bomb (compression ratio too high).")
        if total > config.MAX_UNCOMPRESSED_MB * 1024 * 1024:
            raise UnsupportedFile("Workbook is too large once uncompressed.")
        lower = {n.lower() for n in names}
        if "[content_types].xml" not in lower or "xl/workbook.xml" not in lower:
            raise UnsupportedFile("Not a spreadsheet: xl/workbook.xml is missing.")
        if any(n.endswith("vbaproject.bin") for n in lower):
            raise UnsupportedFile("Workbook contains macros (VBA); macro files are not supported yet.")


def decode_text(data: bytes) -> str:
    if b"\x00" in data[:4096]:
        raise UnsupportedFile("Binary data is not a text/CSV source.")
    for enc in ("utf-8-sig", "cp1258", "cp1252"):
        try:
            return data.decode(enc)
        except UnicodeDecodeError:
            continue
    raise UnsupportedFile("Could not decode text file (use UTF-8).")


def kind_of(filename: str, allowed: set[str]) -> str:
    ext = Path(filename or "").suffix.lower()
    if ext in REFUSED_EXT:
        raise UnsupportedFile(REFUSED_EXT[ext])
    if ext not in allowed:
        raise UnsupportedFile(f"File type {ext or '(none)'} is not supported. Allowed: {', '.join(sorted(allowed))}.")
    return ext[1:]


def store(data: bytes, filename: str, allowed: set[str] | None = None) -> StoredFile:
    allowed = allowed or (XLSX_EXT | SOURCE_EXT)
    kind = kind_of(filename, allowed)
    limit = config.max_upload_bytes() if kind == "xlsx" else min(config.max_upload_bytes(), config.MAX_SOURCE_MB * 1024 * 1024)
    if len(data) > limit:
        raise FileTooLarge(f"File is larger than {limit // (1024 * 1024)} MB.")
    if not data:
        raise UnsupportedFile("File is empty.")
    if kind == "xlsx":
        check_xlsx(data)
    else:
        decode_text(data)
    fid = storage.new_id("file")
    d = storage.uploads_dir() / fid
    path = d / f"original.{kind}"
    storage.atomic_write_bytes(path, data)
    safe_name = Path(filename).name[:120]
    meta = {"id": fid, "filename": safe_name, "kind": kind, "size": len(data), "sha256": storage.sha256(data),
            "created": storage.now()}
    storage.atomic_write_json(d / "meta.json", meta)
    return StoredFile(fid, safe_name, kind, len(data), meta["sha256"], path)


def get(file_id: str) -> StoredFile:
    try:
        storage.check_id(file_id, "file")
    except KeyError:
        raise NotFound("File not found.") from None
    d = storage.uploads_dir() / file_id
    meta = storage.read_json(d / "meta.json")
    if not meta:
        raise NotFound("File not found.")
    return StoredFile(meta["id"], meta["filename"], meta["kind"], meta["size"], meta["sha256"],
                      d / f"original.{meta['kind']}")


def read_text(file_id: str) -> str:
    f = get(file_id)
    if f.kind == "xlsx":
        raise UnsupportedFile("Not a text source.")
    return decode_text(f.path.read_bytes())
