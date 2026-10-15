# excel-agent — AI Excel Workspace

Ứng dụng web giúp tạo, điền và chỉnh sửa file Excel (.xlsx) bằng AI, nhưng **không để AI bịa số liệu**.
Mọi giá trị được ghi vào file đều phải lấy từ tài liệu nguồn (.txt/.csv) hoặc từ câu trả lời của người dùng, kèm
vị trí trích dẫn. Nếu không có nguồn thì ô để trống và được liệt kê là còn thiếu.

- **Backend:** FastAPI (`workspace/`), LangGraph / Deep Agents, LLM là **Azure OpenAI** (mặc định `gpt-5`).
- **Frontend:** React + Vite + Tailwind + shadcn/ui (`frontend/`), dùng FortuneSheet để xem trước bảng tính và
  assistant-ui cho khung chat.
- **File Excel:** sửa trực tiếp trên XML (OOXML) của file gốc, không dùng `openpyxl.save()`. Nhờ vậy giữ được style,
  công thức, merge, dropdown, conditional formatting và biểu đồ. Mỗi lần sửa tạo một **revision** mới; file gốc
  không bao giờ bị ghi đè.

> Không cam kết giữ nguyên 100% mọi tính năng của XLSX. Những gì chưa hỗ trợ được liệt kê trong
> [docs/compatibility-matrix.md](docs/compatibility-matrix.md) và hiện thành cảnh báo trong preview.

## Tính năng

| Chức năng | Mô tả |
|---|---|
| **Tạo từ đầu** (scratch) | Từ các file .txt/.csv và một câu yêu cầu, AI đề xuất **bố cục** (sheet, cột, công thức, tổng, bảng gộp nhóm SUMIF). Dữ liệu được chép theo cách tất định từ nguồn, và từng con số đều được đối chiếu lại với nguồn. |
| **Điền template** | Chọn template có sẵn trong Gallery (Hóa đơn bán hàng, Ngân sách chi phí) hoặc upload file .xlsx của bạn. Hệ thống tự nhận diện ô key/value và bảng nhập, AI ghép nhãn với dữ liệu nguồn và hỏi lại những gì còn thiếu (tối đa 5 câu hỏi một lần). Công thức của template được giữ nguyên. |
| **Đổi cấu trúc template theo tài liệu** | Khi nguồn là văn bản tự do (báo cáo, yêu cầu chỉnh sửa), AI đọc tài liệu rồi đề xuất cấu trúc cột mới (giữ / đổi tên / thêm / cột công thức), các ô header và có thể thêm sheet phụ. Kế hoạch được **chạy thử** trên bản sao, bạn xem lại (kèm trích dẫn nguồn của từng ô) và **duyệt**. Backend từ chối generate nếu kế hoạch chưa được duyệt. |
| **Điền dữ liệu giả lập (mock)** | Chỉ chạy khi prompt **yêu cầu rõ** dữ liệu mock/giả lập/mẫu/dummy và không đính kèm file nguồn. AI (Azure) đọc lưới template rồi soạn giá trị cho các ô giá trị. Bạn xem và sửa dạng `Ô: giá trị` trước khi điền. Backend kiểm tra lại từng ô: không ghi vào ô công thức hay ô nằm trong vùng merge. Workbook được gắn cờ `mock_data` kèm cảnh báo rằng đây là dữ liệu hư cấu. |
| **Sửa thủ công** | Sửa ô trực tiếp trên grid. Mỗi lần lưu là một revision, có undo và lịch sử. |
| **Chat agent** | Agent (Deep Agents) đọc và sửa workbook qua một danh sách tool cố định: `inspect_workbook`, `read_range`, `read_source`, `propose_workbook_edits`, `apply_workbook_edits`, `transform_range`, `validate_workbook`, `build_workbook`, `run_template_extractor`. Agent **không** có quyền truy cập shell hay hệ thống file. Thay đổi ≤ 25 ô được áp dụng ngay; thay đổi lớn hơn (tối đa 1000 ô) cần bạn xác nhận. |
| **Biến file đã điền thành template trống** | Dùng extractor cũ (`xlsx_template/`): AI xác định vùng nào là dữ liệu cần xoá, còn công thức và định dạng được giữ lại. |
| **Export** | File .xlsx chuẩn do backend sinh ra, theo revision bạn chọn. Excel tự tính lại công thức khi mở file. |

Khi chưa cấu hình Azure, ứng dụng chạy ở **chế độ mock**: các luồng tất định (ghép theo tên cột, điền, sửa, export)
vẫn hoạt động. Những tính năng bắt buộc cần LLM (extractor, đổi cấu trúc, chat với model thật) sẽ bị tắt kèm
thông báo, không giả lập kết quả.

## Cấu trúc thư mục

```
workspace/            Backend FastAPI (AI Excel Workspace)
  app.py, api.py        App ASGI + router /api/*; phục vụ luôn frontend đã build và /extractor
  config.py             Biến môi trường, giới hạn, khởi tạo Azure OpenAI
  service.py, storage.py, workspaces.py, runs.py   Workbook/revision, workspace, run + SSE
  ooxml_edit.py         Sửa giá trị ô trực tiếp trên XML
  preview.py            Chuyển workbook sang dữ liệu FortuneSheet + cảnh báo tương thích
  sources.py            Đọc .txt/.csv thành "facts" có vị trí (line:N / row:N), parse số kiểu VN/EN
  clarify.py            Phân tích yêu cầu → câu hỏi làm rõ / kế hoạch
  spec_gen.py, builder.py, generate.py   Luồng tạo từ đầu
  templates.py, template_fill.py         Gallery + nhận diện ô cần nhập + điền template
  restructure.py        Đề xuất cấu trúc mới từ tài liệu, đối chiếu trích dẫn, chạy thử, áp dụng
  formula_eval.py       Bộ tính công thức an toàn (allowlist) dùng cho preview
  agent/                Chat agent: runtime, tools, store, mock model
  extractor_adapter.py  Gọi extractor cũ, trỏ sang Azure OpenAI
xlsx_template/        Extractor cũ (LangGraph): file đã điền → template trống. Không sửa.
frontend/             SPA React (Landing, Studio /new/:mode, Workspace /w/:id)
sample_data/          7 bộ dữ liệu giả lập để test + script build/verify
tests/                pytest (backend)
docs/                 api-contract, compatibility-matrix, implementation-notes, phase-status, screenshots
```

## Cài đặt

Cần **Python 3.12** và **Node.js 20+**.

```powershell
# Backend
uv venv .venv-ws --python 3.12          # hoặc: python -m venv .venv-ws
uv pip install -r requirements.txt      # hoặc: .\.venv-ws\Scripts\python.exe -m pip install -r requirements.txt

# Frontend
cd frontend
npm install
npm run build                           # tạo frontend/dist, backend sẽ phục vụ thư mục này
```

> `requirements.txt` cài `excel-parser` từ thư mục local `./excel-parser-main`, nhưng thư mục này **không có trong
> repo**. Hãy đặt mã nguồn excel-parser vào `excel-parser-main/` trước khi cài, hoặc bỏ dòng `-e ./excel-parser-main`.
> Khi thiếu dòng đó thì chỉ có extractor cũ bị ảnh hưởng.

### Cấu hình `.env`

Tạo file `.env` ở thư mục gốc. File này đã nằm trong `.gitignore`, **không commit file này**.

```dotenv
# Azure OpenAI (bắt buộc nếu muốn dùng LLM)
AZURE_OPENAI_API_KEY=<key>
AZURE_OPENAI_ENDPOINT=https://<resource>.openai.azure.com
AZURE_OPENAI_DEPLOYMENT=gpt-5
# AZURE_OPENAI_API_VERSION=2025-01-01-preview
# AZURE_OPENAI_REASONING=low      # mức reasoning cho gpt-5 / o-series: low | medium | high
```

Cũng có thể khai báo theo kiểu OpenAI-compatible:

```dotenv
OPENAI_API_KEY=<key>
API_BASE=https://<resource>.openai.azure.com/openai/deployments/gpt-5
LANGUAGE_MODEL=openai/gpt-5
```

Các biến tuỳ chọn khác:

| Biến | Mặc định | Ý nghĩa |
|---|---|---|
| `WORKSPACE_LLM=mock` | — | Ép chạy chế độ mock kể cả khi đã có key |
| `WORKSPACE_DATA_DIR` | `./data` | Nơi lưu workspace, file upload và các revision |
| `WORKSPACE_MAX_UPLOAD_MB` | `30` | Dung lượng upload tối đa |
| `HOST` / `PORT` | `127.0.0.1` / `8000` | Địa chỉ của server |
| `GROQ_API_KEY` | — | Chỉ dùng cho **CLI cũ** `python -m xlsx_template` (xem bên dưới) |

## Chạy

```powershell
# Backend (đồng thời phục vụ frontend đã build)
.\.venv-ws\Scripts\python.exe -m workspace
# → http://127.0.0.1:8000   (dòng log đầu tiên cho biết LLM mode: azure | mock)

# Phát triển frontend với hot reload (proxy /api sang :8000)
cd frontend; npm run dev
# → http://localhost:5173
```

Cách dùng nhanh:
1. Ở trang chủ, chọn **Template** hoặc **Tạo mới**.
2. Trong Studio: chọn template (hoặc kéo thả file .xlsx), đính kèm các file nguồn .txt/.csv bằng nút kẹp giấy và
   nhập yêu cầu. Ví dụ: dùng `bao_cao_kinh_doanh_thang_09_2026.txt` với template "Hóa đơn bán hàng".
3. Trả lời các câu hỏi làm rõ, hoặc duyệt kế hoạch cấu trúc, rồi bấm **Generate workbook**.
4. Trong Workspace `/w/:id`: xem preview và cảnh báo, sửa tay hoặc chat với agent, undo, rồi **Export**.

Extractor cũ có giao diện riêng tại `http://127.0.0.1:8000/extractor`.

## Thiết kế an toàn dữ liệu

- **Không bịa số:** LLM chỉ đề xuất *bố cục* và cách ghép cột. Mọi giá trị ghi vào ô phải có trích dẫn nguyên văn
  từ một dòng nguồn (cho phép trích dẫn có `...` nếu từng đoạn đều khớp nguyên văn). Chỉ giá trị 0 được phép
  suy ra (và được đánh dấu *assumed*). Giá trị không có nguồn thì để trống và đưa vào danh sách còn thiếu.
- **Công thức được kiểm tra:** chỉ cho phép các hàm trong allowlist, không tham chiếu sang sheet hay ô tuỳ ý,
  không chứa hằng số tự bịa. Công thức theo dòng được bọc điều kiện trống để ô thiếu dữ liệu không thành 0.
- **Revision bất biến:** revision 1 luôn là template gốc chưa sửa. Xung đột revision trả về 409, và
  `operation_id` đảm bảo gửi lại cùng một thao tác không bị áp dụng hai lần.
- **Upload:** kiểm tra phần mở rộng và chữ ký file, chặn zip bomb, từ chối `.xlsm`.
- **Agent:** chỉ dùng được các tool trong allowlist ở trên. Log và sự kiện chỉ ghi địa chỉ ô, không ghi giá trị.
  Nếu agent nói đã sửa mà không có tool nào thực sự sửa, câu trả lời bị gắn cảnh báo.

Chi tiết API: [docs/api-contract.md](docs/api-contract.md). Trạng thái từng giai đoạn và những gì đã kiểm chứng:
[docs/phase-status.md](docs/phase-status.md).

## Test

```powershell
$env:PYTHONIOENCODING="utf-8"
.\.venv-ws\Scripts\python.exe -m pytest -q          # chạy ở chế độ mock, không gọi LLM thật

# Chạy các bộ sample_data qua API thật (dùng thư mục dữ liệu tạm, không đụng tới data/)
.\.venv-ws\Scripts\python.exe sample_data\_verify.py            # mock
.\.venv-ws\Scripts\python.exe sample_data\_verify.py azure      # Azure OpenAI
```

Frontend: `cd frontend; npm run lint; npm run build`. Lệnh build có chạy `tsc`.

## CLI cũ: `xlsx_template` (file đã điền → template trống)

```
input.xlsx → [parse] excel-parser → [regions] gom ô hằng số thành vùng (công thức không bao giờ gửi cho LLM)
           → [decide] LLM gán keep / clear / review cho từng vùng (review hoặc lỗi → giữ nguyên)
           → [apply] chỉ xoá giá trị ô trên XML gốc → [validate] so từng ô / từng part → PASS/FAIL
           → template.xlsx + .plan.json + .report.md
```

Trong web app, extractor này dùng Azure OpenAI thông qua `workspace/extractor_adapter.py`. Còn khi chạy độc lập
bằng dòng lệnh, nó vẫn dùng Groq (`GROQ_API_KEY`):

```powershell
.\.venv-ws\Scripts\python.exe -m xlsx_template input.xlsx -o output\x.xlsx             # chạy đầy đủ
.\.venv-ws\Scripts\python.exe -m xlsx_template input.xlsx -o output\x.xlsx --dry-run   # chỉ xem kế hoạch
.\.venv-ws\Scripts\python.exe -m xlsx_template input.xlsx -o output\x.xlsx --plan output\x.plan.json  # áp plan đã sửa tay
```

Quyết định của LLM được cache theo cấu trúc vùng trong `cache/decisions.json` (`--no-cache` để bỏ qua cache).

## Hạn chế đã biết

- Chưa chèn thêm dòng khi nguồn có nhiều dòng hơn template; các dòng thừa được cảnh báo.
- Đổi cấu trúc không hỗ trợ template có Excel Table (ListObject), shared formula, hay việc sắp xếp lại các cột
  đang có. Cột bị bỏ nhưng vẫn được tham chiếu ở chỗ khác sẽ bị ẩn chứ không xoá.
- Nguồn chỉ hỗ trợ .txt và .csv. Chưa tạo được biểu đồ mới.
- Với gpt-5, bước đề xuất cấu trúc có thể mất khoảng 1–1,5 phút cho mỗi lần gọi.
- Preview chỉ tính các hàm trong allowlist; Excel tự tính lại toàn bộ công thức khi mở file.

Dữ liệu trong `sample_data/` và hai file `.txt` ở thư mục gốc đều là **dữ liệu giả lập**.
