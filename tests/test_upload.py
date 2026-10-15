"""Upload safety (spec 11): extension, signature, zip safety, size, macro/encrypted rejection."""
from __future__ import annotations

import io
import zipfile

import pytest

from workspace import upload
from workspace.errors import FileTooLarge, NotFound, UnsupportedFile


def _zip(entries: dict[str, bytes]) -> bytes:
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as z:
        for n, d in entries.items():
            z.writestr(n, d)
    return buf.getvalue()


def test_valid_xlsx_stored_immutable_copy(simple_sales):
    data = simple_sales.read_bytes()
    f = upload.store(data, "Sales.xlsx")
    assert f.kind == "xlsx" and f.path.read_bytes() == data
    assert upload.get(f.id).sha256 == f.sha256


@pytest.mark.parametrize("name,msg", [("a.xlsm", "Macro"), ("a.xlsb", "Binary"), ("a.xls", "Legacy"),
                                      ("a.exe", "not supported")])
def test_refused_extensions(name, msg):
    with pytest.raises(UnsupportedFile, match=msg):
        upload.store(b"PK\x03\x04", name)


def test_encrypted_ole_rejected():
    with pytest.raises(UnsupportedFile, match="password"):
        upload.store(upload.OLE_MAGIC + b"\0" * 100, "secret.xlsx")


def test_not_a_zip_rejected():
    with pytest.raises(UnsupportedFile, match="ZIP signature"):
        upload.store(b"hello world", "x.xlsx")


def test_path_traversal_rejected():
    data = _zip({"[Content_Types].xml": b"<x/>", "xl/workbook.xml": b"<x/>", "../evil.txt": b"x"})
    with pytest.raises(UnsupportedFile, match="Unsafe path"):
        upload.check_xlsx(data)


def test_zip_bomb_rejected():
    data = _zip({"[Content_Types].xml": b"<x/>", "xl/workbook.xml": b"<x/>", "xl/big.xml": b"\0" * (5 * 1024 * 1024)})
    with pytest.raises(UnsupportedFile, match="bomb"):
        upload.check_xlsx(data)


def test_vba_rejected_even_with_xlsx_extension():
    data = _zip({"[Content_Types].xml": b"<x/>", "xl/workbook.xml": b"<x/>", "xl/vbaProject.bin": b"x"})
    with pytest.raises(UnsupportedFile, match="macros"):
        upload.store(data, "renamed.xlsx")


def test_missing_workbook_part_rejected():
    with pytest.raises(UnsupportedFile, match="workbook.xml"):
        upload.check_xlsx(_zip({"[Content_Types].xml": b"<x/>"}))


def test_size_limit(monkeypatch):
    monkeypatch.setenv("WORKSPACE_MAX_UPLOAD_MB", "0.001")
    with pytest.raises(FileTooLarge):
        upload.store(b"a" * 5000, "notes.txt")


def test_text_sources_decoded():
    f = upload.store("Doanh thu: 1.200.000 VND\n".encode("utf-8"), "notes.txt")
    assert "Doanh thu" in upload.read_text(f.id)
    with pytest.raises(UnsupportedFile, match="Binary"):
        upload.store(b"\x00\x01\x02", "bin.csv")


def test_unknown_id_not_found():
    with pytest.raises(NotFound):
        upload.get("../../etc")
