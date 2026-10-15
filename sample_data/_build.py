"""Generates the sample templates (.xlsx) and their source data (.csv/.txt) in this folder.

All names, companies and figures are fictional test data. Re-run: .venv-ws/Scripts/python.exe sample_data/_build.py
"""
from __future__ import annotations

import calendar
import datetime as dt
import math
import random
from pathlib import Path

import xlsxwriter

ROOT = Path(__file__).parent


def _fmts(wb: xlsxwriter.Workbook) -> dict:
    border = {"border": 1, "border_color": "#B7C9E2"}
    return {
        "title": wb.add_format({"bold": True, "font_size": 16, "font_color": "#1F4E78"}),
        "sub": wb.add_format({"italic": True, "font_color": "#7F7F7F"}),
        "label": wb.add_format({"bold": True}),
        "input": wb.add_format({"bg_color": "#FFF9E6", **border}),
        "date_in": wb.add_format({"bg_color": "#FFF9E6", "num_format": "dd/mm/yyyy", **border}),
        "num_in": wb.add_format({"bg_color": "#FFF9E6", "num_format": "#,##0", **border}),
        "pct_in": wb.add_format({"bg_color": "#FFF9E6", "num_format": "0.0", **border}),
        "head": wb.add_format({"bold": True, "font_color": "white", "bg_color": "#1F4E78", "align": "center",
                               "valign": "vcenter", "text_wrap": True, **border}),
        "num": wb.add_format({"num_format": "#,##0", **border}),
        "pct": wb.add_format({"num_format": "0.0%", **border}),
        "total": wb.add_format({"bold": True, "bg_color": "#DDEBF7", "top": 2, "num_format": "#,##0", **border}),
        "total_pct": wb.add_format({"bold": True, "bg_color": "#DDEBF7", "top": 2, "num_format": "0.0%",
                                    **border}),
    }


def _kv(ws, f, row: int, items: list[tuple[str, str]]) -> None:
    for i, (lab, fmt) in enumerate(items):
        ws.write(row + i, 0, lab, f["label"])
        ws.write_blank(row + i, 1, None, f[fmt])


def _write(path: Path, text: str) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(text.strip() + "\n", encoding="utf-8")


# ---------------------------------------------------------------- 1. Bảng lương
def payroll(d: Path, rows: int = 15) -> None:
    wb = xlsxwriter.Workbook(str(d / "template_bang_luong.xlsx"))
    f = _fmts(wb)
    ws = wb.add_worksheet("Bảng lương")
    ws.write(0, 0, "BẢNG LƯƠNG NHÂN VIÊN", f["title"])
    _kv(ws, f, 2, [("Công ty:", "input"), ("Tháng:", "input"), ("Người lập:", "input"),
                   ("Ngày công chuẩn:", "num_in")])
    hr = 7
    heads = ["STT", "Họ tên", "Phòng ban", "Lương cơ bản", "Phụ cấp", "Ngày công", "Thực lĩnh"]
    for c, h in enumerate(heads):
        ws.write(hr, c, h, f["head"])
    for i in range(rows):
        r = hr + 1 + i
        x = r + 1
        ws.write_blank(r, 0, None, f["num_in"])
        ws.write_blank(r, 1, None, f["input"])
        ws.write_blank(r, 2, None, f["input"])
        ws.write_blank(r, 3, None, f["num_in"])
        ws.write_blank(r, 4, None, f["num_in"])
        ws.write_blank(r, 5, None, f["num_in"])
        ws.write_formula(r, 6, f'=IF(D{x}="","",ROUND(D{x}*F{x}/IF($B$6="",26,$B$6)+N(E{x}),0))', f["num"], "")
    tr = hr + 1 + rows
    ws.write(tr, 1, "Tổng cộng", f["total"])
    for c in (0, 2, 5):
        ws.write_blank(tr, c, None, f["total"])
    for c, col in ((3, "D"), (4, "E"), (6, "G")):
        ws.write_formula(tr, c, f"=SUM({col}{hr + 2}:{col}{tr})", f["total"], 0)
    ws.set_column(0, 0, 18)
    ws.set_column(1, 1, 26)
    ws.set_column(2, 2, 16)
    ws.set_column(3, 6, 15)
    ws.freeze_panes(hr + 1, 0)
    wb.close()
    _write(d / "nhan_vien_thang_9.csv", """
Họ tên,Phòng ban,Lương cơ bản,Phụ cấp,Ngày công
Nguyễn Văn An,Kinh doanh,15000000,1500000,26
Trần Thị Bình,Kế toán,13500000,1000000,25
Lê Hoàng Cường,Kỹ thuật,18000000,2000000,26
Phạm Thu Dung,Nhân sự,12000000,800000,24
Võ Minh Đức,Kỹ thuật,16500000,2000000,26
Đặng Ngọc Hà,Kinh doanh,14000000,1500000,22
Bùi Quang Huy,Kho vận,10500000,700000,26
Hoàng Lan Anh,Kế toán,13000000,1000000,26
""")
    _write(d / "thong_tin_ky_luong.txt", """
Công ty: Công ty TNHH Thương mại Sao Mai (dữ liệu giả lập)
Tháng: 09/2026
Người lập: Phạm Thu Dung
Ngày công chuẩn: 26
""")
    _write(d / "prompt.txt", "Điền bảng lương tháng 9/2026 từ danh sách nhân viên và thông tin kỳ lương đính kèm.")


# ---------------------------------------------------------------- 2. Báo giá (nhãn khác nhau giữa template và data)
def quotation(d: Path, rows: int = 10) -> None:
    wb = xlsxwriter.Workbook(str(d / "template_bao_gia.xlsx"))
    f = _fmts(wb)
    ws = wb.add_worksheet("Báo giá")
    ws.write(0, 0, "BÁO GIÁ SẢN PHẨM", f["title"])
    ws.write(1, 0, "Ô màu vàng là ô cần nhập", f["sub"])
    _kv(ws, f, 3, [("Khách hàng:", "input"), ("Người liên hệ:", "input"), ("Điện thoại:", "input"),
                   ("Số báo giá:", "input"), ("Ngày báo giá:", "date_in"), ("Hiệu lực (ngày):", "num_in")])
    hr = 10
    for c, h in enumerate(["STT", "Mã hàng", "Tên hàng", "ĐVT", "Số lượng", "Đơn giá", "Thành tiền"]):
        ws.write(hr, c, h, f["head"])
    ws.data_validation(hr + 1, 3, hr + rows, 3, {"validate": "list", "source": ["Cái", "Hộp", "Bộ", "Thùng", "Mét"]})
    for i in range(rows):
        r = hr + 1 + i
        x = r + 1
        ws.write_blank(r, 0, None, f["num_in"])
        for c in (1, 2, 3):
            ws.write_blank(r, c, None, f["input"])
        ws.write_blank(r, 4, None, f["num_in"])
        ws.write_blank(r, 5, None, f["num_in"])
        ws.write_formula(r, 6, f'=IF(E{x}="","",E{x}*F{x})', f["num"], "")
    tr = hr + 1 + rows
    first, last = hr + 2, tr
    ws.write(tr, 2, "Cộng tiền hàng", f["total"])
    ws.write_formula(tr, 6, f"=SUM(G{first}:G{last})", f["total"], 0)
    ws.write(tr + 1, 2, "Chiết khấu (%):", f["label"])
    ws.write_blank(tr + 1, 3, None, f["pct_in"])
    ws.write_formula(tr + 1, 6, f"=-G{tr + 1}*N(D{tr + 2})/100", f["num"], 0)
    ws.write(tr + 2, 2, "VAT (%):", f["label"])
    ws.write_blank(tr + 2, 3, None, f["pct_in"])
    ws.write_formula(tr + 2, 6, f"=(G{tr + 1}+G{tr + 2})*N(D{tr + 3})/100", f["num"], 0)
    ws.write(tr + 3, 2, "Tổng thanh toán", f["total"])
    ws.write_formula(tr + 3, 6, f"=G{tr + 1}+G{tr + 2}+G{tr + 3}", f["total"], 0)
    ws.set_column(0, 0, 18)
    ws.set_column(1, 1, 12)
    ws.set_column(2, 2, 32)
    ws.set_column(3, 6, 14)
    wb.close()
    # nhãn trong data cố ý khác template (Mã SP / Sản phẩm / SL / Giá bán) để thử khả năng map
    _write(d / "danh_sach_hang.csv", """
Mã SP,Sản phẩm,ĐVT,SL,Giá bán
LT-14,Laptop văn phòng 14 inch,Cái,5,"15.900.000"
MN-24,Màn hình 24 inch IPS,Cái,5,"3.250.000"
BP-K1,Bàn phím cơ không dây,Cái,10,"890.000"
CH-HD,Chuột không dây,Hộp,10,"245.000"
CAP-C,Cáp HDMI 2m,Mét,20,"65.000"
""")
    _write(d / "yeu_cau_khach_hang.txt", """
Bên mua: Công ty Cổ phần Giải pháp Hướng Dương (dữ liệu giả lập)
Người liên hệ: Chị Ngô Thanh Mai
Điện thoại: 0901 234 567
Số BG: BG-2026-0915
Ngày lập: 15/09/2026
Hiệu lực (ngày): 30
Chiết khấu (%): 5
Ghi chú: khách yêu cầu giao trong 7 ngày làm việc.
""")
    _write(d / "prompt.txt", "Lập báo giá cho khách theo yêu cầu và danh sách hàng đính kèm. Không có thông tin VAT.")


# ---------------------------------------------------------------- 3. Báo cáo doanh thu chi nhánh (số kiểu VN, 2 sheet)
def branch_report(d: Path, rows: int = 8) -> None:
    wb = xlsxwriter.Workbook(str(d / "template_bao_cao_doanh_thu.xlsx"))
    f = _fmts(wb)
    ws = wb.add_worksheet("Doanh thu")
    ws.write(0, 0, "BÁO CÁO DOANH THU THEO CHI NHÁNH", f["title"])
    _kv(ws, f, 2, [("Kỳ báo cáo:", "input"), ("Đơn vị tính:", "input")])
    hr = 5
    for c, h in enumerate(["Chi nhánh", "Kế hoạch", "Thực hiện", "% hoàn thành", "Ghi chú"]):
        ws.write(hr, c, h, f["head"])
    for i in range(rows):
        r = hr + 1 + i
        x = r + 1
        ws.write_blank(r, 0, None, f["input"])
        ws.write_blank(r, 1, None, f["num_in"])
        ws.write_blank(r, 2, None, f["num_in"])
        ws.write_formula(r, 3, f'=IF(OR(B{x}="",B{x}=0),"",C{x}/B{x})', f["pct"], "")
        ws.write_blank(r, 4, None, f["input"])
    tr = hr + 1 + rows
    ws.write(tr, 0, "Toàn công ty", f["total"])
    ws.write_formula(tr, 1, f"=SUM(B{hr + 2}:B{tr})", f["total"], 0)
    ws.write_formula(tr, 2, f"=SUM(C{hr + 2}:C{tr})", f["total"], 0)
    ws.write_formula(tr, 3, f'=IF(B{tr + 1}=0,"",C{tr + 1}/B{tr + 1})', f["total_pct"], "")
    ws.write_blank(tr, 4, None, f["total"])
    ws.conditional_format(hr + 1, 3, hr + rows, 3, {"type": "cell", "criteria": "<", "value": 1,
                                                     "format": wb.add_format({"font_color": "#C00000"})})
    ws.set_column(0, 0, 22)
    ws.set_column(1, 3, 18)
    ws.set_column(4, 4, 34)
    chart = wb.add_chart({"type": "column"})
    chart.add_series({"name": "Kế hoạch", "categories": ["Doanh thu", hr + 1, 0, hr + rows, 0],
                      "values": ["Doanh thu", hr + 1, 1, hr + rows, 1]})
    chart.add_series({"name": "Thực hiện", "categories": ["Doanh thu", hr + 1, 0, hr + rows, 0],
                      "values": ["Doanh thu", hr + 1, 2, hr + rows, 2]})
    chart.set_title({"name": "Kế hoạch vs thực hiện"})
    ws.insert_chart("G3", chart)
    info = wb.add_worksheet("Hướng dẫn")
    info.write(0, 0, "Sheet này chỉ để đọc: không có ô nhập, công cụ không được điền vào đây.", f["sub"])
    wb.close()
    # số kiểu Việt Nam: dấu chấm ngăn nghìn, file phân tách bằng dấu chấm phẩy
    _write(d / "doanh_thu_q3_2026.csv", """
Chi nhánh;Kế hoạch;Thực hiện;Ghi chú
Hà Nội;12.500.000.000;13.120.000.000;Vượt nhờ dự án trường học
TP. Hồ Chí Minh;18.000.000.000;16.940.000.000;Thiếu hàng tháng 8
Đà Nẵng;6.200.000.000;6.480.000.000;
Cần Thơ;4.100.000.000;3.650.000.000;Mất 1 khách hàng lớn
Hải Phòng;5.300.000.000;5.310.000.000;
""")
    _write(d / "thong_tin_bao_cao.txt", """
Kỳ báo cáo: Quý 3/2026
Đơn vị tính: VNĐ
""")
    _write(d / "prompt.txt", "Điền báo cáo doanh thu quý 3/2026 theo chi nhánh từ file số liệu đính kèm.")


# ---------------------------------------------------------------- 4. Phiếu nhập kho (data dạng bảng markdown trong TXT, dư dòng)
def stock_in(d: Path, rows: int = 6) -> None:
    wb = xlsxwriter.Workbook(str(d / "template_phieu_nhap_kho.xlsx"))
    f = _fmts(wb)
    ws = wb.add_worksheet("Phiếu nhập")
    ws.merge_range(0, 0, 0, 5, "PHIẾU NHẬP KHO", wb.add_format({"bold": True, "font_size": 16, "align": "center",
                                                                 "font_color": "#1F4E78"}))
    _kv(ws, f, 2, [("Số phiếu:", "input"), ("Ngày nhập:", "date_in"), ("Nhà cung cấp:", "input"),
                   ("Kho:", "input"), ("Mã số thuế NCC:", "input")])
    hr = 8
    for c, h in enumerate(["Mã vật tư", "Tên vật tư", "ĐVT", "Số lượng", "Đơn giá", "Thành tiền"]):
        ws.write(hr, c, h, f["head"])
    for i in range(rows):
        r = hr + 1 + i
        x = r + 1
        for c in (0, 1, 2):
            ws.write_blank(r, c, None, f["input"])
        ws.write_blank(r, 3, None, f["num_in"])
        ws.write_blank(r, 4, None, f["num_in"])
        ws.write_formula(r, 5, f'=IF(D{x}="","",D{x}*E{x})', f["num"], "")
    tr = hr + 1 + rows
    ws.write(tr, 1, "Cộng", f["total"])
    ws.write_formula(tr, 5, f"=SUM(F{hr + 2}:F{tr})", f["total"], 0)
    ws.set_column(0, 0, 18)
    ws.set_column(1, 1, 30)
    ws.set_column(2, 5, 14)
    wb.close()
    # 8 dòng hàng nhưng template chỉ có 6 dòng -> phải cảnh báo tràn, không chèn dòng; không có MST -> để trống
    _write(d / "phieu_nhap_ncc.txt", """
Số phiếu: PNK-0926-018
Ngày nhập: 22/09/2026
Nhà cung cấp: Công ty TNHH Vật tư Bình Minh (dữ liệu giả lập)
Kho: Kho Bình Dương

| Mã vật tư | Tên vật tư | ĐVT | Số lượng | Đơn giá |
|---|---|---|---|---|
| VT-001 | Thép hộp 40x80 | Cây | 120 | 285000 |
| VT-002 | Tôn lạnh 0.45mm | Tấm | 80 | 196000 |
| VT-003 | Bulong M12 | Bộ | 500 | 4500 |
| VT-004 | Que hàn 3.2mm | Hộp | 25 | 152000 |
| VT-005 | Sơn chống gỉ | Thùng | 12 | 980000 |
| VT-006 | Ốc vít tự khoan | Hộp | 40 | 68000 |
| VT-007 | Băng keo cách điện | Cuộn | 60 | 9500 |
| VT-008 | Găng tay bảo hộ | Đôi | 100 | 18000 |
""")
    _write(d / "prompt.txt", "Lập phiếu nhập kho theo chứng từ nhà cung cấp đính kèm.")
    # Answer for the clarification question (the supplier tax code is missing from the sources); not a source file.
    _write(d / "tra_loi_cau_hoi.txt", "Mã số thuế NCC: 0100000000")


# ---------------------------------------------------------------- 5. Workbook có sẵn số liệu để test chat agent
def chat_workbook(d: Path) -> None:
    wb = xlsxwriter.Workbook(str(d / "theo_doi_ban_hang.xlsx"))
    f = _fmts(wb)
    ws = wb.add_worksheet("Bán hàng")
    ws.write(0, 0, "THEO DÕI BÁN HÀNG THÁNG 9/2026 (dữ liệu giả lập)", f["title"])
    hr = 2
    for c, h in enumerate(["Sản phẩm", "Số lượng", "Đơn giá", "Doanh thu", "Nhân viên"]):
        ws.write(hr, c, h, f["head"])
    data = [("Laptop văn phòng", 12, 15900000, "An"), ("Màn hình 24 inch", 30, 3250000, "Bình"),
            ("Bàn phím cơ", 45, 890000, "An"), ("Chuột không dây", 80, 245000, "Cường"),
            ("Tai nghe", 25, 650000, "Bình"), ("Webcam HD", 18, 720000, "Cường")]
    for i, (name, q, p, who) in enumerate(data):
        r = hr + 1 + i
        ws.write(r, 0, name)
        ws.write_number(r, 1, q, f["num"])
        ws.write_number(r, 2, p, f["num"])
        ws.write_formula(r, 3, f"=B{r + 1}*C{r + 1}", f["num"], q * p)
        ws.write(r, 4, who)
    tr = hr + 1 + len(data)
    ws.write(tr, 0, "Tổng", f["total"])
    ws.write_formula(tr, 1, f"=SUM(B{hr + 2}:B{tr})", f["total"], sum(x[1] for x in data))
    ws.write_blank(tr, 2, None, f["total"])
    ws.write_formula(tr, 3, f"=SUM(D{hr + 2}:D{tr})", f["total"], sum(x[1] * x[2] for x in data))
    ws.write_blank(tr, 4, None, f["total"])
    ws.set_column(0, 0, 22)
    ws.set_column(1, 3, 16)
    ws.set_column(4, 4, 14)
    t = wb.add_worksheet("Chỉ tiêu")
    t.write_row(0, 0, ["Nhân viên", "Chỉ tiêu", "Đạt"], f["head"])
    for i, (who, goal) in enumerate([("An", 200000000), ("Bình", 120000000), ("Cường", 40000000)]):
        t.write(i + 1, 0, who)
        t.write_number(i + 1, 1, goal, f["num"])
        t.write_formula(i + 1, 2, f"=SUMIF('Bán hàng'!E:E,A{i + 2},'Bán hàng'!D:D)", f["num"])
    t.set_column(0, 2, 16)
    wb.close()


def _xround(v: float) -> int:
    """Excel ROUND(v, 0): half away from zero (Python's round() is banker's rounding)."""
    return int(math.floor(abs(v) + 0.5)) * (1 if v >= 0 else -1)


# ---------------------------------------------------------------- 6. Sổ bán hàng cả năm (300 dòng, nhiều hàm)
PRODUCTS = [("SP-01", "Laptop văn phòng 14 inch", 15900000), ("SP-02", "Màn hình 24 inch IPS", 3250000),
            ("SP-03", "Bàn phím cơ", 890000), ("SP-04", "Chuột không dây", 245000),
            ("SP-05", "Tai nghe chụp tai", 650000), ("SP-06", "Webcam HD", 720000),
            ("SP-07", "Ổ cứng SSD 1TB", 1850000), ("SP-08", "Router Wi-Fi 6", 1290000),
            ("SP-09", "Cáp HDMI 2m", 65000), ("SP-10", "Giá đỡ laptop", 320000)]


def sales_ledger(d: Path, rows: int = 400, lines: int = 300) -> dict:
    wb = xlsxwriter.Workbook(str(d / "template_so_ban_hang.xlsx"))
    f = _fmts(wb)
    ws = wb.add_worksheet("Sổ bán hàng")
    cat = wb.add_worksheet("Danh mục")
    rep = wb.add_worksheet("Tổng hợp")
    date_f = wb.add_format({"num_format": "dd/mm/yyyy", "bg_color": "#FFF9E6", "border": 1, "border_color": "#B7C9E2"})
    pct_f = wb.add_format({"num_format": "0%", "border": 1, "border_color": "#B7C9E2"})

    # --- Sổ bán hàng: 6 cột nhập, 7 cột công thức
    ws.write(0, 0, "SỔ CHI TIẾT BÁN HÀNG", f["title"])
    _kv(ws, f, 2, [("Năm:", "num_in"), ("Đơn vị:", "input"), ("Người lập:", "input")])
    hr, first, last = 6, 8, 7 + rows                       # Excel rows of the input block: 8..407
    heads = ["Ngày", "Số HĐ", "Mã KH", "Mã hàng", "Tên hàng", "Số lượng", "Đơn giá", "Thành tiền",
             "CK (%)", "Tiền CK", "VAT", "Phải thu", "Tháng"]
    for c, h in enumerate(heads):
        ws.write(hr, c, h, f["head"])
    lk = "'Danh mục'!$A$4:$C$13"
    for i in range(rows):
        r = hr + 1 + i
        x = r + 1
        ws.write_blank(r, 0, None, date_f)
        for c in (1, 2, 3):
            ws.write_blank(r, c, None, f["input"])
        ws.write_blank(r, 5, None, f["num_in"])
        ws.write_formula(r, 4, f'=IF(D{x}="","",IFERROR(VLOOKUP(D{x},{lk},2,FALSE),"Không rõ mã"))', None, "")
        ws.write_formula(r, 6, f'=IF(D{x}="","",IFERROR(VLOOKUP(D{x},{lk},3,FALSE),0))', f["num"], "")
        ws.write_formula(r, 7, f'=IF(F{x}="","",F{x}*G{x})', f["num"], "")
        ws.write_formula(r, 8, f'=IF(F{x}="","",IF(F{x}>=50,5%,IF(F{x}>=20,2%,0)))', pct_f, "")
        ws.write_formula(r, 9, f'=IF(H{x}="","",ROUND(H{x}*I{x},0))', f["num"], "")
        ws.write_formula(r, 10, f"=IF(H{x}=\"\",\"\",ROUND((H{x}-J{x})*'Danh mục'!$F$3/100,0))", f["num"], "")
        ws.write_formula(r, 11, f'=IF(H{x}="","",H{x}-J{x}+K{x})', f["num"], "")
        ws.write_formula(r, 12, f'=IF(A{x}="","",MONTH(A{x}))', None, "")
    tr = hr + 1 + rows
    ws.write(tr, 1, "Tổng cộng", f["total"])
    for c in (0, 2, 3, 4, 6, 8, 12):
        ws.write_blank(tr, c, None, f["total"])
    for c in (5, 7, 9, 10, 11):
        col = chr(ord("A") + c)
        ws.write_formula(tr, c, f"=SUBTOTAL(9,{col}{first}:{col}{last})", f["total"], 0)
    ws.autofilter(hr, 0, hr + rows, len(heads) - 1)
    ws.freeze_panes(hr + 1, 0)
    ws.set_column(0, 3, 12)
    ws.set_column(4, 4, 26)
    ws.set_column(5, 12, 13)

    # --- Danh mục: bảng tra cứu tên + đơn giá, tham số VAT
    cat.write(0, 0, "DANH MỤC HÀNG HÓA", f["title"])
    for c, h in enumerate(["Mã hàng", "Tên hàng", "Đơn giá"]):
        cat.write(2, c, h, f["head"])
    for i, (code, name, price) in enumerate(PRODUCTS):
        cat.write(3 + i, 0, code)
        cat.write(3 + i, 1, name)
        cat.write_number(3 + i, 2, price, f["num"])
    cat.write(2, 4, "VAT (%)", f["label"])
    cat.write_number(2, 5, 10)
    cat.set_column(0, 0, 10)
    cat.set_column(1, 1, 28)
    cat.set_column(2, 5, 13)

    # --- Tổng hợp: theo tháng, theo sản phẩm, chỉ số nhanh
    S = lambda col: f"'Sổ bán hàng'!${col}${first}:${col}${last}"  # noqa: E731
    rep.write(0, 0, "TỔNG HỢP BÁN HÀNG", f["title"])
    mh = ["Tháng", "Số dòng", "Số lượng", "Doanh thu", "Chiết khấu", "Phải thu", "TB/dòng", "% cả năm", "Lũy kế"]
    for c, h in enumerate(mh):
        rep.write(2, c, h, f["head"])
    for m in range(1, 13):
        r, x = 2 + m, 3 + m
        rep.write_number(r, 0, m)
        rep.write_formula(r, 1, f"=COUNTIF({S('M')},A{x})", f["num"], 0)
        rep.write_formula(r, 2, f"=SUMIF({S('M')},A{x},{S('F')})", f["num"], 0)
        rep.write_formula(r, 3, f"=SUMIF({S('M')},A{x},{S('H')})", f["num"], 0)
        rep.write_formula(r, 4, f"=SUMIF({S('M')},A{x},{S('J')})", f["num"], 0)
        rep.write_formula(r, 5, f"=SUMIF({S('M')},A{x},{S('L')})", f["num"], 0)
        rep.write_formula(r, 6, f"=IFERROR(AVERAGEIF({S('M')},A{x},{S('L')}),0)", f["num"], 0)
        rep.write_formula(r, 7, f"=IF($F$16=0,0,F{x}/$F$16)", f["pct"], 0)
        rep.write_formula(r, 8, f"=SUM($F$4:F{x})", f["num"], 0)
    rep.write(15, 0, "Cả năm", f["total"])
    for c in range(1, 6):
        col = chr(ord("A") + c)
        rep.write_formula(15, c, f"=SUM({col}4:{col}15)", f["total"], 0)
    rep.write_formula(15, 6, "=IFERROR(F16/B16,0)", f["total"], 0)
    rep.write_formula(15, 7, "=SUM(H4:H15)", f["total_pct"], 0)
    rep.write_blank(15, 8, None, f["total"])
    for c, h in enumerate(["Mã hàng", "Tên hàng", "Số lượng bán", "Doanh thu", "Xếp hạng"]):
        rep.write(18, c, h, f["head"])
    for i, (code, _n, _p) in enumerate(PRODUCTS):
        r, x = 19 + i, 20 + i
        rep.write(r, 0, code)
        rep.write_formula(r, 1, f"=INDEX('Danh mục'!$B$4:$B$13,MATCH(A{x},'Danh mục'!$A$4:$A$13,0))", None, "")
        rep.write_formula(r, 2, f"=SUMIF({S('D')},A{x},{S('F')})", f["num"], 0)
        rep.write_formula(r, 3, f"=SUMIF({S('D')},A{x},{S('H')})", f["num"], 0)
        rep.write_formula(r, 4, f"=RANK(D{x},$D$20:$D$29)", None, 0)
    rep.merge_range(2, 10, 2, 11, "CHỈ SỐ NHANH", f["head"])
    kpis = [("Tổng phải thu cả năm", "=F16"),
            ("Dòng phải thu lớn nhất", f"=MAX({S('L')})"),
            ("Dòng phải thu nhỏ nhất", f"=MIN({S('L')})"),
            ("Trung bình mỗi dòng", f"=IFERROR(ROUND(AVERAGE({S('L')}),0),0)"),
            ("Trung vị mỗi dòng", f"=IFERROR(MEDIAN({S('L')}),0)"),
            ("Số hóa đơn", f'=SUMPRODUCT(({S("B")}<>"")/COUNTIF({S("B")},{S("B")}&""))'),
            ("Số khách hàng", f'=SUMPRODUCT(({S("C")}<>"")/COUNTIF({S("C")},{S("C")}&""))'),
            ("Tháng phải thu cao nhất", "=INDEX(A4:A15,MATCH(MAX(F4:F15),F4:F15,0))"),
            ("Sản phẩm doanh thu cao nhất", "=INDEX(B20:B29,MATCH(MAX(D20:D29),D20:D29,0))"),
            ("Dòng có mã hàng lạ", f'=COUNTIF({S("E")},"Không rõ mã")'),
            ("Dòng được chiết khấu", f'=COUNTIF({S("I")},">0")')]
    for i, (lab, fx) in enumerate(kpis):
        rep.write(3 + i, 10, lab, f["label"])
        rep.write_formula(3 + i, 11, fx, f["num"], 0)
    rep.set_column(0, 0, 10)
    rep.set_column(1, 8, 14)
    rep.set_column(10, 10, 28)
    rep.set_column(11, 11, 18)
    chart = wb.add_chart({"type": "line"})
    chart.add_series({"name": "Phải thu", "categories": ["Tổng hợp", 3, 0, 14, 0],
                      "values": ["Tổng hợp", 3, 5, 14, 5], "marker": {"type": "circle"}})
    chart.set_title({"name": "Phải thu theo tháng"})
    chart.set_legend({"none": True})
    rep.insert_chart("G19", chart)
    wb.close()

    # --- dữ liệu: ~300 dòng, nhiều dòng/hóa đơn, mùa cao điểm cuối năm, 1 mã hàng lạ
    rnd = random.Random(2026)
    weights = [4, 5, 6, 6, 7, 7, 8, 9, 9, 10, 13, 16]             # tháng 11-12 bán nhiều hơn
    out, inv, prices = [], 0, {c: p for c, _n, p in PRODUCTS}
    while len(out) < lines:
        inv += 1
        month = rnd.choices(range(1, 13), weights)[0]
        day = rnd.randint(1, calendar.monthrange(2026, month)[1])
        cust = f"KH{rnd.randint(1, 45):03d}"
        for _ in range(min(rnd.choice([1, 1, 2, 2, 3, 4]), lines - len(out))):
            code = rnd.choice(PRODUCTS)[0]
            qty = rnd.randint(1, 15) if prices[code] >= 1000000 else rnd.randint(5, 120)
            out.append([dt.date(2026, month, day), inv, cust, code, qty])
    out.sort(key=lambda r: (r[0], r[1]))
    out[137][3] = "SP-99"                                           # mã không có trong danh mục
    num = {}
    for r in out:                                                   # đánh số HĐ theo thứ tự ngày
        num.setdefault(r[1], len(num) + 1)
    lines_csv = ["Ngày,Số HĐ,Mã KH,Mã hàng,Số lượng"]
    total = {"thanh_tien": 0, "ck": 0, "vat": 0, "phai_thu": 0}
    for day, i, cust, code, qty in out:
        lines_csv.append(f"{day:%d/%m/%Y},HD-2026-{num[i]:04d},{cust},{code},{qty}")
        h = qty * prices.get(code, 0)
        ck = _xround(h * (0.05 if qty >= 50 else 0.02 if qty >= 20 else 0))
        vat = _xround((h - ck) * 10 / 100)
        for k, v in zip(total, (h, ck, vat, h - ck + vat)):
            total[k] += v
    _write(d / "ban_hang_2026.csv", "\n".join(lines_csv))
    _write(d / "thong_tin_so.txt", """
Năm: 2026
Đơn vị: Công ty TNHH Thiết bị Văn phòng Ánh Dương (dữ liệu giả lập)
Người lập: Lê Thị Hồng
""")
    _write(d / "prompt.txt", "Điền sổ chi tiết bán hàng năm 2026 từ file bán hàng đính kèm.")
    return {"dòng": len(out), "hóa đơn": len(num), **total}


# ---------------------------------------------------------------- 7. Bảng lương 120 người có BHXH + thuế TNCN lũy tiến
DEPTS = {"Kinh doanh": 1500000, "Kế toán": 1000000, "Kỹ thuật": 2000000, "Nhân sự": 800000, "Kho vận": 700000,
         "Chăm sóc KH": 900000}
BRACKETS = [(0, 0.05), (5000000, 0.10), (10000000, 0.15), (18000000, 0.20), (32000000, 0.25), (52000000, 0.30),
            (80000000, 0.35)]
SELF, DEP, INS_RATE, INS_CAP = 11000000, 4400000, 0.105, 46800000


def payroll_tax(d: Path, rows: int = 150, people: int = 120) -> dict:
    wb = xlsxwriter.Workbook(str(d / "template_bang_luong_thue.xlsx"))
    f = _fmts(wb)
    ws = wb.add_worksheet("Bảng lương")
    par = wb.add_worksheet("Tham số")
    rep = wb.add_worksheet("Tổng hợp")
    pct_f = wb.add_format({"num_format": "0.0%"})

    ws.write(0, 0, "BẢNG LƯƠNG, BẢO HIỂM VÀ THUẾ TNCN", f["title"])
    _kv(ws, f, 2, [("Công ty:", "input"), ("Kỳ lương:", "input"), ("Ngày công chuẩn:", "num_in")])
    hr, first, last = 6, 8, 7 + rows
    heads = ["Mã NV", "Họ tên", "Phòng ban", "Lương cơ bản", "Ngày công", "Giờ tăng ca", "Số NPT", "Phụ cấp",
             "Lương theo công", "Tiền tăng ca", "Tổng thu nhập", "BHXH-BHYT-BHTN", "Thu nhập tính thuế",
             "Thuế TNCN", "Thực lĩnh", "Xếp loại công"]
    for c, h in enumerate(heads):
        ws.write(hr, c, h, f["head"])
    std = 'IF(N($B$5)=0,26,$B$5)'
    tp = "'Tham số'!"
    for i in range(rows):
        r = hr + 1 + i
        x = r + 1
        for c in (0, 1, 2):
            ws.write_blank(r, c, None, f["input"])
        for c in (3, 4, 5, 6):
            ws.write_blank(r, c, None, f["num_in"])
        fx = [
            f'=IF(C{x}="","",IFERROR(VLOOKUP(C{x},{tp}$A$20:$B$25,2,FALSE),0))',
            f'=IF(D{x}="","",ROUND(D{x}*E{x}/{std},0))',
            f'=IF(D{x}="","",ROUND(D{x}/{std}/8*1.5*N(F{x}),0))',
            f'=IF(D{x}="","",I{x}+J{x}+N(H{x}))',
            f'=IF(D{x}="","",ROUND(MIN(D{x},{tp}$B$6)*{tp}$B$5,0))',
            f'=IF(D{x}="","",MAX(0,K{x}-L{x}-{tp}$B$3-N(G{x})*{tp}$B$4))',
            f'=IF(D{x}="","",ROUND(SUMPRODUCT(--(M{x}>{tp}$A$10:$A$16),M{x}-{tp}$A$10:$A$16,{tp}$C$10:$C$16),0))',
            f'=IF(D{x}="","",K{x}-L{x}-N{x})',
            f'=IF(E{x}="","",IF(E{x}>={std},"Đủ công",IF(E{x}>={std}-3,"Thiếu ít","Thiếu nhiều")))',
        ]
        for j, formula in enumerate(fx):
            ws.write_formula(r, 7 + j, formula, f["num"] if j < 8 else None, "")
    tr = hr + 1 + rows
    ws.write_formula(tr, 0, f'=COUNTA(B{first}:B{last})&" người"', f["total"], "")
    ws.write(tr, 1, "Tổng cộng", f["total"])
    for c in range(2, 16):
        col = chr(ord("A") + c)
        if c in (3, 5, 7, 8, 9, 10, 11, 13, 14):
            ws.write_formula(tr, c, f"=SUM({col}{first}:{col}{last})", f["total"], 0)
        else:
            ws.write_blank(tr, c, None, f["total"])
    ws.freeze_panes(hr + 1, 2)
    ws.set_column(0, 0, 18)
    ws.set_column(1, 1, 24)
    ws.set_column(2, 2, 14)
    ws.set_column(3, 15, 14)

    par.write(0, 0, "THAM SỐ TÍNH LƯƠNG (minh họa, cần đối chiếu quy định hiện hành trước khi dùng thật)", f["sub"])
    for i, (lab, v, fm) in enumerate([("Giảm trừ bản thân", SELF, f["num"]), ("Giảm trừ mỗi người phụ thuộc", DEP, f["num"]),
                                      ("Tỷ lệ BH người lao động đóng", INS_RATE, pct_f),
                                      ("Trần lương đóng BH", INS_CAP, f["num"])]):
        par.write(2 + i, 0, lab, f["label"])
        par.write_number(2 + i, 1, v, fm)
    par.merge_range(7, 0, 7, 2, "Biểu thuế lũy tiến từng phần (theo tháng)", f["label"])
    for c, h in enumerate(["Thu nhập tính thuế từ", "Thuế suất", "Chênh lệch thuế suất"]):
        par.write(8, c, h, f["head"])
    for i, (lo, rate) in enumerate(BRACKETS):
        par.write_number(9 + i, 0, lo, f["num"])
        par.write_number(9 + i, 1, rate, pct_f)
        par.write_formula(9 + i, 2, "=B10" if i == 0 else f"=B{10 + i}-B{9 + i}", pct_f, rate)
    par.merge_range(17, 0, 17, 1, "Phụ cấp theo phòng ban", f["label"])
    for c, h in enumerate(["Phòng ban", "Phụ cấp"]):
        par.write(18, c, h, f["head"])
    for i, (dep, amt) in enumerate(DEPTS.items()):
        par.write(19 + i, 0, dep)
        par.write_number(19 + i, 1, amt, f["num"])
    par.set_column(0, 0, 30)
    par.set_column(1, 2, 18)

    B = lambda col: f"'Bảng lương'!${col}${first}:${col}${last}"  # noqa: E731
    rep.write(0, 0, "TỔNG HỢP QUỸ LƯƠNG THEO PHÒNG BAN", f["title"])
    for c, h in enumerate(["Phòng ban", "Số NV", "Tổng thu nhập", "Bảo hiểm", "Thuế TNCN", "Thực lĩnh",
                           "TB thực lĩnh", "Tỷ trọng"]):
        rep.write(2, c, h, f["head"])
    deps = [*DEPTS, "Pháp chế"]
    for i, dep in enumerate(deps):
        r, x = 3 + i, 4 + i
        rep.write(r, 0, dep)
        rep.write_formula(r, 1, f"=COUNTIF({B('C')},A{x})", f["num"], 0)
        for c, col in ((2, "K"), (3, "L"), (4, "N"), (5, "O")):
            rep.write_formula(r, c, f"=SUMIF({B('C')},A{x},{B(col)})", f["num"], 0)
        rep.write_formula(r, 6, f"=IFERROR(AVERAGEIF({B('C')},A{x},{B('O')}),0)", f["num"], 0)
        rep.write_formula(r, 7, f"=IF($F${4 + len(deps)}=0,0,F{x}/$F${4 + len(deps)})", f["pct"], 0)
    tr2 = 3 + len(deps)
    rep.write(tr2, 0, "Toàn công ty", f["total"])
    for c in range(1, 6):
        col = chr(ord("A") + c)
        rep.write_formula(tr2, c, f"=SUM({col}4:{col}{tr2})", f["total"], 0)
    rep.write_formula(tr2, 6, f"=IFERROR(F{tr2 + 1}/B{tr2 + 1},0)", f["total"], 0)
    rep.write_formula(tr2, 7, f"=SUM(H4:H{tr2})", f["total_pct"], 0)
    rep.merge_range(2, 9, 2, 10, "CHỈ SỐ NHANH", f["head"])
    kpis = [("Thực lĩnh cao nhất", f"=MAX({B('O')})"), ("Thực lĩnh thấp nhất", f"=MIN({B('O')})"),
            ("Trung vị thực lĩnh", f"=IFERROR(MEDIAN({B('O')}),0)"),
            ("Người lĩnh cao nhất", f"=INDEX({B('B')},MATCH(MAX({B('O')}),{B('O')},0))"),
            ("Số người phải nộp thuế", f'=COUNTIF({B("N")},">0")'),
            ("Số người thiếu nhiều công", f'=COUNTIF({B("P")},"Thiếu nhiều")'),
            ("Tổng giờ tăng ca", f"=SUM({B('F')})"),
            ("Thuế / tổng thu nhập", f"=IFERROR(SUM({B('N')})/SUM({B('K')}),0)")]
    for i, (lab, fx) in enumerate(kpis):
        rep.write(3 + i, 9, lab, f["label"])
        rep.write_formula(3 + i, 10, fx, pct_f if i == 7 else f["num"], 0)
    rep.set_column(0, 0, 16)
    rep.set_column(1, 7, 15)
    rep.set_column(9, 9, 28)
    rep.set_column(10, 10, 18)
    wb.close()

    rnd = random.Random(926)
    ho = ["Nguyễn", "Trần", "Lê", "Phạm", "Hoàng", "Huỳnh", "Phan", "Vũ", "Võ", "Đặng", "Bùi", "Đỗ", "Hồ", "Ngô",
          "Dương", "Lý"]
    dem = ["Văn", "Thị", "Minh", "Ngọc", "Quang", "Thu", "Hoàng", "Thanh", "Đức", "Phương", "Gia", "Bảo"]
    ten = ["An", "Bình", "Châu", "Dũng", "Giang", "Hà", "Hải", "Hạnh", "Hiếu", "Hòa", "Hùng", "Khánh", "Lan",
           "Linh", "Long", "Mai", "Nam", "Nga", "Phúc", "Quân", "Sơn", "Tâm", "Thảo", "Trang", "Tuấn", "Vy", "Yến"]
    base = {"Kinh doanh": (9, 22), "Kế toán": (9, 18), "Kỹ thuật": (12, 35), "Nhân sự": (8, 16),
            "Kho vận": (6, 11), "Chăm sóc KH": (7, 12), "Pháp chế": (15, 25)}
    csv = ["Mã NV,Họ tên,Phòng ban,Lương cơ bản,Ngày công,Giờ tăng ca,Số NPT"]
    tot = {"thu_nhap": 0, "bao_hiem": 0, "thue": 0, "thuc_linh": 0, "nguoi_nop_thue": 0}
    for i in range(people):
        dep = "Pháp chế" if i in (57, 101) else rnd.choice(list(DEPTS))   # phòng không có trong bảng phụ cấp
        lo, hi = base[dep]
        lcb = rnd.randint(lo * 10, hi * 10) * 100000
        days = rnd.choices([26, 25, 24, 22, 20, 18], [60, 15, 10, 7, 5, 3])[0]
        ot = rnd.choice([0, 0, 0, 4, 8, 12, 16, 24, 32])
        npt = rnd.choices([0, 1, 2, 3], [45, 30, 18, 7])[0]
        csv.append(f"NV{i + 1:04d},{rnd.choice(ho)} {rnd.choice(dem)} {rnd.choice(ten)},{dep},{lcb},{days},{ot},{npt}")
        gross = _xround(lcb * days / 26) + _xround(lcb / 26 / 8 * 1.5 * ot) + DEPTS.get(dep, 0)
        ins = _xround(min(lcb, INS_CAP) * INS_RATE)
        taxable = max(0, gross - ins - SELF - npt * DEP)
        tax = _xround(sum(max(0, taxable - lo_) * (rate - (BRACKETS[k - 1][1] if k else 0))
                          for k, (lo_, rate) in enumerate(BRACKETS)))
        for k, v in zip(tot, (gross, ins, tax, gross - ins - tax, int(tax > 0))):
            tot[k] += v
    _write(d / "nhan_vien_thang_10.csv", "\n".join(csv))
    _write(d / "thong_tin_ky_luong.txt", """
Công ty: Công ty Cổ phần Công nghệ Phương Nam (dữ liệu giả lập)
Kỳ lương: 10/2026
Ngày công chuẩn: 26
""")
    _write(d / "prompt.txt", "Điền bảng lương tháng 10/2026 cho toàn bộ nhân viên từ file đính kèm.")
    return {"người": people, **tot}


if __name__ == "__main__":
    for name, fn in [("01_bang_luong", payroll), ("02_bao_gia", quotation), ("03_bao_cao_doanh_thu", branch_report),
                     ("04_phieu_nhap_kho", stock_in), ("05_chat_agent", chat_workbook),
                     ("06_so_ban_hang_nam", sales_ledger), ("07_luong_thue_tncn", payroll_tax)]:
        d = ROOT / name
        d.mkdir(parents=True, exist_ok=True)
        expected = fn(d)
        print("built", name, sorted(p.name for p in d.iterdir()))
        if expected:
            print("   expected totals:", expected)
