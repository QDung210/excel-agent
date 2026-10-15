# Bộ dữ liệu mẫu để test

Toàn bộ là **dữ liệu giả lập** (tên công ty và người đều bịa, có ghi "(dữ liệu giả lập)"). Ô màu vàng trong template là ô cần nhập.

- Tạo lại toàn bộ file: `.venv-ws/Scripts/python.exe sample_data/_build.py`
- Chạy tự động mọi bộ qua API thật, dùng thư mục dữ liệu tạm nên không đụng `data/`:
  - Mock: `PYTHONIOENCODING=utf-8 .venv-ws/Scripts/python.exe sample_data/_verify.py`
  - Groq: `PYTHONIOENCODING=utf-8 .venv-ws/Scripts/python.exe sample_data/_verify.py groq`
  - Chỉ chạy một vài bộ: `... _verify.py mock 06_so_ban_hang_nam 07_luong_thue_tncn`

## Cách dùng trong UI (bộ 01–04)

1. Tạo workspace ở chế độ **Template**.
2. Upload file `template_*.xlsx` với vai trò *template*.
3. Upload các file `.csv`/`.txt` còn lại (trừ `prompt.txt`) với vai trò *source*.
4. Dán nội dung `prompt.txt` vào ô yêu cầu, rồi bấm **Phân tích** và **Tạo**.
5. Xem preview và các cảnh báo, sau đó export.

Có thể upload template vào Gallery trước rồi chọn từ đó.

## Kết quả thực tế đã kiểm (2026-10-13)

| Bộ | Kiểm tra điều gì | Mock | Groq |
|---|---|---|---|
| **01_bang_luong** | Ca chuẩn: header CSV trùng tên với template. 4 ô KV, 8/15 dòng nhân viên. Công thức Thực lĩnh và Tổng cộng giữ nguyên. | Điền đủ, không hỏi lại, không cảnh báo | Như mock |
| **02_bao_gia** | **Header và nhãn khác tên**: CSV dùng `Mã SP / Sản phẩm / SL / Giá bán`, TXT dùng `Bên mua / Số BG / Ngày lập`. Số có dạng `"15.900.000"`. Nguồn không có VAT. | Bảng ghép hiện 2 ô **cần xác nhận** (Số BG→Số báo giá, Ngày lập→Ngày báo giá). Bảng hàng: AI **đoán** `danh_sach_hang.csv` là nguồn và hỏi từng cột (Mã hàng gợi ý `Mã SP`, ĐVT gợi ý `ĐVT`, các cột khác để người dùng chọn, mỗi lựa chọn kèm giá trị ví dụ). Hỏi thêm Khách hàng và VAT vì không tìm thấy. `_verify.py` nhận các gợi ý mặc định nên điền Mã hàng và ĐVT; Tên hàng, SL, Đơn giá để trống kèm cảnh báo. | LLM ghép đủ: Bên mua→Khách hàng, Số BG, Ngày lập, cả 5 cột bảng. **Chỉ hỏi VAT**. Còn cảnh báo "Ghi chú" không có chỗ đặt. |
| **03_bao_cao_doanh_thu** | CSV dùng dấu `;` và số VN `12.500.000.000`. Một số ô ghi chú trống. Template có sheet "Hướng dẫn" chỉ để đọc, conditional formatting và biểu đồ. | 5 chi nhánh được điền. Cột % hoàn thành giữ công thức. Sheet Hướng dẫn và biểu đồ không bị đụng. | Như mock |
| **04_phieu_nhap_kho** | **Tràn dòng**: nguồn có 8 vật tư nhưng template chỉ có 6 dòng. Nguồn không có MST. | **Hỏi lại** MST; `_verify.py` trả lời bằng `tra_loi_cau_hoi.txt` nên B7 = `'0100000000'` (giữ số 0 đầu). Điền 6 dòng. Cảnh báo *2 dòng không vừa (chưa hỗ trợ chèn dòng)*. | Như mock |

| **06_so_ban_hang_nam** | **Dữ liệu dài**: 300 dòng, 127 hóa đơn, 44 khách, ngày dạng `dd/mm/yyyy`. Template có 400 dòng nhập và 3 sheet. Có 1 mã hàng lạ `SP-99`. | Điền đủ 300 dòng (2 lô ghi nên lên revision 3). Công thức ở cả 400 dòng giữ nguyên. | Như mock |
| **07_luong_thue_tncn** | **Dữ liệu dài + công thức nặng**: 120 nhân viên, 2 người thuộc phòng "Pháp chế" không có trong bảng phụ cấp. Template có 150 dòng. | Điền đủ 120 dòng, công thức giữ nguyên | Như mock |

Với bộ 06 và 07 sẽ có thêm cảnh báo *"N table(s) already contain data and were not changed"*. Đây là hành vi đúng: các bảng đó là bảng tra cứu và bảng tổng hợp (sheet Danh mục, Tham số, Tổng hợp), không phải chỗ để nhập.

Khi mở bằng Excel, các công thức (Thành tiền, Tổng, Thực lĩnh, %) sẽ tự tính lại.

### Bộ 06 và 07: các hàm dùng trong template và số đúng để đối chiếu

Các số dưới đây được `_build.py` tính độc lập bằng Python. Sau đó tôi đã mở file đã điền bằng **Excel thật** (tính lại toàn bộ): mọi số khớp và không có ô lỗi `#...`.

**06 – Sổ bán hàng** (sheet *Sổ bán hàng*, *Danh mục*, *Tổng hợp*)
- Mỗi dòng dùng: `VLOOKUP` + `IFERROR` (tra tên hàng và đơn giá; mã lạ hiện "Không rõ mã"), `IF` lồng (chiết khấu 2% khi ≥20 cái, 5% khi ≥50 cái), `ROUND` (tiền CK, VAT lấy từ ô tham số `'Danh mục'!F3`), `MONTH`.
- Tổng hợp dùng: `SUBTOTAL` (dòng tổng, đúng cả khi lọc AutoFilter), `COUNTIF`/`SUMIF`/`AVERAGEIF` theo tháng, lũy kế `SUM($F$4:F4)`, `INDEX`+`MATCH`, `RANK`, `MAX`/`MIN`/`AVERAGE`/`MEDIAN`, `SUMPRODUCT(1/COUNTIF)` để đếm số hóa đơn và số khách không trùng. Có thêm biểu đồ đường.
- Số đúng:

  | Chỉ tiêu | Giá trị |
  |---|---|
  | Thành tiền | 9.276.400.000 |
  | Chiết khấu | 193.825.850 |
  | VAT | 908.257.415 |
  | **Phải thu** | **9.990.831.565** |
  | Số hóa đơn | 127 |
  | Số khách hàng | 44 |
  | Dòng được chiết khấu | 159 |
  | Tháng phải thu cao nhất | 12 |

- "Dòng phải thu nhỏ nhất = 0" là do dòng có mã `SP-99` (đơn giá 0). Đây là chủ ý của dữ liệu.

**07 – Lương + thuế TNCN** (sheet *Bảng lương*, *Tham số*, *Tổng hợp*)
- Mỗi dòng dùng: `VLOOKUP` phụ cấp theo phòng ban (phòng lạ thì `IFERROR` trả về 0), `ROUND` lương theo công và tăng ca ×1,5, `MIN` (trần lương đóng BH), `MAX(0, …)` (thu nhập tính thuế sau giảm trừ bản thân và người phụ thuộc), **`SUMPRODUCT` tính thuế lũy tiến 7 bậc** từ bảng ở sheet *Tham số*, `IF` lồng để xếp loại công.
- Tổng hợp dùng: `COUNTIF`/`SUMIF`/`AVERAGEIF` theo phòng, `INDEX`+`MATCH` (người lĩnh cao nhất), `MEDIAN`, `COUNTIF(">0")`, và chuỗi `COUNTA(...)&" người"`.
- Số đúng:

  | Chỉ tiêu | Giá trị |
  |---|---|
  | Tổng thu nhập | 1.853.337.501 |
  | Bảo hiểm | 174.951.000 |
  | Thuế TNCN | 26.177.995 |
  | **Thực lĩnh** | **1.652.208.506** |
  | Số người phải nộp thuế | 41 / 120 |
  | Số người phòng Pháp chế | 2 |

- Các tham số thuế và bảo hiểm (giảm trừ 11 triệu / 4,4 triệu, 10,5%, trần 46,8 triệu, biểu 7 bậc) chỉ để **minh họa**. Chúng nằm ở sheet *Tham số* và có thể sửa. Không dùng để tính lương thật khi chưa đối chiếu quy định hiện hành.

## Khi nào agent hỏi lại

Agent chỉ hỏi khi **thiếu thông tin thật**, không hỏi cho có. Nếu nguồn đã đủ (bộ 01, 03, 06, 07) thì nó đi thẳng tới bước tạo.

- **Template mode**: ngay ở bước *Phân tích*, agent đọc template và nguồn rồi đưa ra **bảng ghép** (*How the template will be filled*): mỗi ô/cột của mẫu lấy từ đâu trong nguồn, giá trị là gì, trạng thái *Matched / Confirm / No source*. Sau đó chỉ hỏi những chỗ chưa chắc:
  - Ô **khớp đúng nhãn** (hoặc AI chắc chắn): điền luôn, không hỏi.
  - Ô **AI đoán** (nhãn gần giống, ví dụ `Số BG`→`Số báo giá`, hoặc AI đánh dấu chưa chắc): hỏi dạng chọn, mặc định là gợi ý của AI, có lựa chọn *(để trống)*.
  - **Bảng nguồn có tên cột khác hẳn**: AI đoán đó là bảng nguồn và hỏi từng cột lấy từ cột nào (mỗi lựa chọn kèm giá trị ví dụ). Để trống hết thì bảng không được điền.
  - Ô **không có trong nguồn**: hỏi nhập `nhãn: giá trị`. Để trống thì bỏ trống và có cảnh báo.

  AI chỉ được **trỏ** tới dữ liệu (id của fact, cột nguồn, hoặc đoạn trích nguyên văn có thật trong file text). Hệ thống tự kiểm tra lại, và giá trị luôn đọc từ nguồn, AI không tự viết số. Bước *Tạo* dùng lại đúng bảng ghép đã xác nhận, không gọi AI lần hai. Ở chế độ mock (không có Groq) chỉ ghép theo tên nên hỏi nhiều hơn.
- **Không có file nguồn**: hỏi xin dữ liệu, hoặc trả về mẫu trống.
- **Scratch mode**: hỏi đơn vị tiền tệ, kỳ báo cáo, có số liệu hay không, và các cột cần có.
- **Chat**: câu muốn sửa mà không nói ô hoặc giá trị (vd *"điền doanh thu tháng 10"*, *"sửa lại cho đúng"*) thì hỏi lại và không ghi gì. Câu đòi công thức, định dạng hay chèn dòng thì báo là chưa hỗ trợ.

Muốn tự thử: xóa dòng `Điện thoại:` trong `02_bao_gia/yeu_cau_khach_hang.txt` thì câu hỏi sẽ có thêm "Điện thoại".

## Bộ 05_chat_agent — test chat sửa workbook

Mở `theo_doi_ban_hang.xlsx` bằng **Upload workbook**, sau đó chat. `_verify.py` chạy thật cả 8 câu ở cả hai chế độ:

| Câu chat | Mock | Groq |
|---|---|---|
| `tóm tắt workbook` | Liệt kê 2 sheet và vùng dữ liệu | Tóm tắt chi tiết hơn |
| `đọc A3:E9` | In giá trị từng ô | Bảng giá trị |
| `đặt B4 = 15` | Ghi ngay vì ≤ 25 ô, revision +1. Doanh thu D4 và sheet Chỉ tiêu tự tính lại khi mở Excel. | Như mock |
| `xóa A20:F30` | 66 ô > 25, nên hiện kế hoạch chờ bấm **Áp dụng** | Như mock |
| `đồng ý` (sau câu trên) | Áp dụng kế hoạch, revision +1 | Như mock |
| `điền doanh thu tháng 10` | **Hỏi lại**: ghi vào ô nào, giá trị bao nhiêu | **Hỏi lại** ô và số tiền, không ghi gì |
| `thêm công thức tổng cho cột B` | Báo chưa hỗ trợ ghi công thức | Như mock |
| `kiểm tra workbook` | Báo có ô lỗi hay không | Như mock |

Lưu ý khi chạy Groq bằng key free: nếu bị rate limit **sau khi đã ghi**, câu trả lời sẽ là *"Đã ghi thay đổi (workbook giờ ở revision N) nhưng model không phản hồi tiếp"*. Thay đổi vẫn đã được lưu; chờ vài giây rồi chat tiếp.
