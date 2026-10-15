"""Golden fixtures built deterministically at test time (nothing binary is committed).

simple_sales       titles, 3 columns, rows, SUM formulas, number/date formats, hyperlink
rich_template      2 visible sheets + hidden list sheet, merged header, styled blanks, formulas,
                   list dropdown via named range, widths, freeze panes, conditional format, comment
x14_validation     rich_template + an x14 <extLst> data validation injected into the XML
                   (the shape Excel writes for cross-sheet dropdowns; openpyxl drops it on save)
"""
from __future__ import annotations

import datetime as dt
import shutil
import zipfile
from pathlib import Path

import xlsxwriter


def simple_sales(path: Path) -> Path:
    wb = xlsxwriter.Workbook(str(path))
    ws = wb.add_worksheet("Sales")
    bold = wb.add_format({"bold": True, "bg_color": "#1F4E78", "font_color": "#FFFFFF", "border": 1})
    money = wb.add_format({"num_format": "#,##0", "border": 1})
    datef = wb.add_format({"num_format": "yyyy-mm-dd"})
    ws.write("A1", "Sales report")
    ws.write_datetime("B1", dt.datetime(2026, 9, 30), datef)
    for c, h in enumerate(["Product", "Quantity", "Unit price", "Revenue"]):
        ws.write(1, c, h, bold)
    rows = [("Laptop", 20, 15000000), ("Mouse", 35, 200000), ("Keyboard", 12, 450000)]
    for r, (p, q, u) in enumerate(rows, start=2):
        ws.write(r, 0, p)
        ws.write_number(r, 1, q, money)
        ws.write_number(r, 2, u, money)
        ws.write_formula(r, 3, f"=B{r + 1}*C{r + 1}", money, q * u)
    ws.write("A6", "Total", bold)
    ws.write_formula("D6", "=SUM(D3:D5)", money, sum(q * u for _, q, u in rows))
    ws.write_url("A8", "mailto:owner@example.com", string="owner@example.com")
    ws.set_column("A:A", 24)
    ws.set_column("B:D", 16)
    wb.close()
    return path


def rich_template(path: Path) -> Path:
    wb = xlsxwriter.Workbook(str(path))
    ws = wb.add_worksheet("Report")
    sm = wb.add_worksheet("Summary")
    lists = wb.add_worksheet("Lists")
    lists.hide()
    for i, v in enumerate(["Open", "Done", "Blocked"]):
        lists.write(i, 0, v)
    wb.define_name("StatusList", "=Lists!$A$1:$A$3")

    title = wb.add_format({"bold": True, "font_size": 14, "align": "center", "bg_color": "#DDEBF7"})
    head = wb.add_format({"bold": True, "border": 1, "bg_color": "#BDD7EE"})
    blank = wb.add_format({"border": 1, "bg_color": "#FFF2CC"})
    num = wb.add_format({"border": 1, "num_format": "#,##0.00"})
    ws.merge_range("A1:D1", "Monthly status report", title)
    for c, h in enumerate(["Task", "Status", "Hours", "Cost"]):
        ws.write(1, c, h, head)
    for r in range(2, 7):
        ws.write_blank(r, 0, None, blank)
        ws.write_blank(r, 1, None, blank)
        ws.write_blank(r, 2, None, blank)
        ws.write_formula(r, 3, f"=C{r + 1}*25", num, 0)
    ws.write("A8", "Total", head)
    ws.write_formula("C8", "=SUM(C3:C7)", num, 0)
    ws.write_formula("D8", "=SUM(D3:D7)", num, 0)
    ws.data_validation("B3:B7", {"validate": "list", "source": "=StatusList"})
    ws.conditional_format("C3:C7", {"type": "cell", "criteria": ">", "value": 40,
                                    "format": wb.add_format({"font_color": "#9C0006"})})
    ws.write_comment("A2", "Fill one task per row")
    ws.set_column("A:A", 30)
    ws.set_column("B:D", 14)
    ws.freeze_panes(2, 0)

    sm.write("A1", "Total hours", head)
    sm.write_formula("B1", "=Report!C8", num, 0)
    sm.write("A2", "Total cost", head)
    sm.write_formula("B2", "=Report!D8", num, 0)
    sm.write("A3", "Owner")
    sm.write_blank("B3", None, blank)
    wb.close()
    return path


X14_EXT = (
    '<extLst><ext uri="{CCE6A557-97BC-4b89-ADB6-D9C93CAAB3DF}" '
    'xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main">'
    '<x14:dataValidations count="1" xmlns:xm="http://schemas.microsoft.com/office/excel/2006/main">'
    '<x14:dataValidation type="list" allowBlank="1" showInputMessage="1" showErrorMessage="1">'
    "<x14:formula1><xm:f>Lists!$A$1:$A$3</xm:f></x14:formula1><xm:sqref>B3:B7</xm:sqref>"
    "</x14:dataValidation></x14:dataValidations></ext></extLst>"
)


def x14_validation(path: Path) -> Path:
    tmp = path.with_suffix(".base.xlsx")
    rich_template(tmp)
    with zipfile.ZipFile(tmp) as zin, zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as zout:
        for info in zin.infolist():
            data = zin.read(info.filename)
            if info.filename == "xl/worksheets/sheet1.xml":
                xml = data.decode("utf-8")
                # extLst must be the last child of <worksheet>
                xml = xml.replace("</worksheet>", X14_EXT + "</worksheet>")
                data = xml.encode("utf-8")
            zout.writestr(info, data)
    tmp.unlink()
    return path


def copy(src: Path, dst: Path) -> Path:
    shutil.copyfile(src, dst)
    return dst
