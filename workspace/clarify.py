"""Analyze -> clarify flow (spec 03): decide what is missing before generating, never guess numbers.

State lives in the workspace JSON under `clarification`:
{id, prompt, source_file_ids, template_file_id, template_id, output_filename, status, questions, answers,
 warnings, plan, facts}
"""
from __future__ import annotations

import re
from typing import Any

from . import config, mock_fill, sources, storage, templates, workspaces
from .errors import InvalidOperation
from .schemas import AnalyzeRequest, AnalyzeResponse, ClarifyAnswers, ClarifyQuestion

VI_RE = re.compile(r"[ăâđêôơưàảãáạằẳẵắặầẩẫấậèẻẽéẹềểễếệìỉĩíịòỏõóọồổỗốộờởỡớợùủũúụừửữứựỳỷỹýỵ]", re.I)
REPORT_WORDS = re.compile(r"báo\s*cáo|report|doanh\s*thu|revenue|budget|ngân\s*sách|chi\s*phí|expense|sales|"
                          r"bán\s*hàng|lợi\s*nhuận|profit|forecast|dự\s*báo|kế\s*hoạch|plan|invoice|hóa\s*đơn|"
                          r"thống\s*kê|summary|tổng\s*hợp|quotation|báo\s*giá|lương|payroll", re.I)
PERIOD_WORDS = re.compile(r"tháng|month|quý|quarter|năm|year|weekly|tuần|monthly|kỳ|period", re.I)
CHART_WORDS = re.compile(r"chart|biểu\s*đồ|graph|đồ\s*thị", re.I)
EDIT_WORDS = re.compile(r"(thêm|bỏ|xóa|xoá|bớt|đổi\s*tên|sửa|chỉnh|điều\s*chỉnh)\s*(lại\s*)?(các\s*|những\s*)?"
                        r"(cột|tiêu\s*đề|mẫu|template|công\s*thức|cấu\s*trúc)|"
                        r"(add|remove|delete|drop|rename)\s+(a\s+|the\s+)?columns?|restructure|change\s+the\s+formula",
                        re.I)
CURRENCIES = ["VND", "USD", "EUR"]


def is_vi(text: str) -> bool:
    return bool(VI_RE.search(text or ""))


def _t(vi: bool, a: str, b: str) -> str:
    return a if vi else b


def _questions(prompt: str, b: sources.SourceBundle, mode: str,
               missing: list[str] | None = None) -> tuple[list[ClarifyQuestion], list[str]]:
    vi = is_vi(prompt) or any(is_vi(t) for t in b.texts.values())
    qs: list[ClarifyQuestion] = []
    warnings: list[str] = []
    headers = " ".join(h for t in b.tables for h in t.headers)
    data_numbers = [f for f in b.numeric_facts() if f.status == "observed"]
    money = bool(sources.MONEY_WORDS.search(prompt) or sources.MONEY_WORDS.search(headers)
                 or any(sources.MONEY_WORDS.search(f.field) for f in b.facts))
    if mode == "scratch" and money and not b.currency:
        qs.append(ClarifyQuestion(
            id="currency", type="select", required=True, options=CURRENCIES, default="VND" if vi else "USD",
            question=_t(vi, "Số tiền trong bảng tính dùng đơn vị tiền tệ nào?", "Which currency should amounts use?"),
            explanation=_t(vi, "Không tìm thấy đơn vị tiền tệ rõ ràng trong yêu cầu hoặc file nguồn.",
                           "No unambiguous currency was found in the prompt or sources.")))
    if mode == "scratch" and not data_numbers and REPORT_WORDS.search(prompt):
        qs.append(ClarifyQuestion(
            id="missing_numbers", type="select", required=True,
            options=["blank", "provide"], default="blank",
            question=_t(vi, "Không có số liệu nguồn. Tạo bảng với ô nhập liệu để trống, hay bạn cung cấp số liệu?",
                        "No source figures were found. Leave input cells blank, or will you provide figures?"),
            explanation=_t(vi, "Hệ thống không tự bịa số liệu. 'blank' = để trống ô nhập; 'provide' = nhập ở câu "
                               "hỏi tiếp theo.",
                           "Numbers are never invented. 'blank' leaves input cells empty; 'provide' uses the next "
                           "answer.")))
        qs.append(ClarifyQuestion(
            id="figures", type="text", required=False,
            question=_t(vi, "Số liệu (mỗi dòng 'tên: giá trị'), nếu chọn 'provide'",
                        "Figures (one 'label: value' per line) if you chose 'provide'"),
            explanation=_t(vi, "Ví dụ: Doanh thu tháng 1: 120.000.000", "Example: January revenue: 120,000")))
    if mode == "template" and not b.names and not data_numbers:
        qs.append(ClarifyQuestion(
            id="figures", type="text", required=False,
            question=_t(vi, "Chưa có file nguồn. Nhập dữ liệu để điền mẫu (mỗi dòng 'nhãn: giá trị'), hoặc để trống "
                            "để nhận mẫu trống",
                        "No source file attached. Enter data to fill (one 'label: value' per line), or leave empty "
                        "to get the blank template"),
            explanation=_t(vi, "Nhãn nên trùng với nhãn trong mẫu, vd 'Khách hàng: Công ty A'.",
                           "Labels should match the template, e.g. 'Customer: ACME'.")))
    elif mode == "template" and missing:
        llm = config.llm_mode() == "azure"
        shown = ", ".join(missing[:12]) + (f" (+{len(missing) - 12})" if len(missing) > 12 else "")
        qs.append(ClarifyQuestion(
            id="figures", type="text", required=False,
            question=_t(vi, f"Chưa tìm thấy giá trị trong nguồn cho {len(missing)} ô của mẫu: {shown}. Nhập giá trị "
                            "(mỗi dòng 'nhãn: giá trị'), hoặc để trống",
                        f"No source value found for {len(missing)} template cell(s): {shown}. Enter values (one "
                        "'label: value' per line), or leave empty"),
            explanation=_t(vi, f"Hệ thống không tự đoán. Ví dụ: '{missing[0]}: ...'. " + (
                               "Khi tạo, AI còn thử ghép các nhãn khác tên trong nguồn (vd 'Bên mua' -> 'Khách hàng'); "
                               "ô nào vẫn không có giá trị sẽ để trống và được cảnh báo." if llm else
                               "Ô bạn không nhập sẽ để trống và được cảnh báo."),
                           f"Nothing is guessed. Example: '{missing[0]}: ...'. " + (
                               "When generating, the AI also tries differently-named source labels; cells still "
                               "without a value stay blank and are reported." if llm else
                               "Cells you skip stay blank and are reported."))))
    if mode == "scratch" and not b.period and REPORT_WORDS.search(prompt) and PERIOD_WORDS.search(prompt):
        qs.append(ClarifyQuestion(
            id="period", type="text", required=False,
            question=_t(vi, "Kỳ báo cáo là gì? (vd: Tháng 9/2026)", "Which reporting period? (e.g. Sep 2026)")))
    if mode == "scratch" and not b.tables and len(qs) < 5:
        qs.append(ClarifyQuestion(
            id="columns", type="text", required=False,
            question=_t(vi, "Các cột cần có (phân cách bằng dấu phẩy), để trống để dùng mặc định",
                        "Columns you want (comma separated); leave empty for defaults")))
    if CHART_WORDS.search(prompt):
        warnings.append(_t(vi, "Biểu đồ chưa được hỗ trợ: workbook sẽ không có biểu đồ.",
                           "Charts are not supported yet: the workbook will not contain a chart."))
    return qs[:5], warnings


def _mock_questions(state: dict, prompt: str) -> tuple[list[ClarifyQuestion], list[str]]:
    """The user explicitly asked for mock data and attached no source: the AI drafts values, the user reviews them."""
    from . import template_fill
    vi = is_vi(prompt)
    if config.llm_mode() != "azure":
        return [], [_t(vi, "Tạo dữ liệu giả lập cần Azure OpenAI (đang ở chế độ mock): mẫu sẽ không được điền.",
                       "Generating mock data needs Azure OpenAI (mock mode): the template is not filled.")]
    try:
        src, _ = template_fill._template_source(state)
        text, n, warns = mock_fill.propose(src, prompt)
    except Exception as e:                               # noqa: BLE001 - reported, falls back to the normal question
        return [], [f"Could not generate mock data ({type(e).__name__}); enter values below or retry."]
    if not n:
        return [], warns + ["The AI proposed no mock values for this template."]
    state["mock"] = {"cells": n}
    return [
        ClarifyQuestion(
            id="mock_apply", type="select", required=True, options=["apply", "blank"], default="apply",
            question=_t(vi, f"AI đã soạn dữ liệu GIẢ LẬP cho {n} ô của mẫu. Điền vào file?",
                        f"The AI drafted MOCK data for {n} template cell(s). Fill them in?"),
            explanation=_t(vi, "Đây là dữ liệu giả lập theo yêu cầu của bạn, không phải dữ liệu thật. 'blank' = giữ "
                               "nguyên mẫu. Nhãn và công thức không bị đụng; bản gốc vẫn là revision 1.",
                           "This is fictional data you asked for, not real data. 'blank' keeps the template as is. "
                           "Labels and formulas are untouched; the original stays revision 1.")),
        ClarifyQuestion(
            id="mock_cells", type="text", required=False, default=text,
            question=_t(vi, "Xem và sửa dữ liệu giả lập (mỗi dòng 'Ô: giá trị', xoá dòng để bỏ qua ô đó)",
                        "Review and edit the mock data (one 'CELL: value' per line; delete a line to skip it)"),
            explanation=_t(vi, "\\n trong giá trị là xuống dòng trong ô.", "\\n inside a value is a line break.")),
    ], warns


def _plan(mode: str, b: sources.SourceBundle, ready: dict) -> list[str]:
    steps = [f"Read {len(b.names)} source file(s): {len(b.facts)} fact(s), {len(b.tables)} table(s)"]
    if mode == "template" and ready.get("mock"):
        steps.append(f"Fill up to {ready['mock']['cells']} cell(s) with AI-generated MOCK data that you reviewed "
                     "(labels, formulas and styles untouched); the workbook is marked as mock data")
    elif mode == "template" and ready.get("structure"):
        from . import restructure
        p = restructure.effective(ready["structure"])
        cols = p["columns"]
        n = {k: sum(1 for c in cols if k in c["changes"]) for k in ("add", "rename", "to_input", "formula",
                                                                     "to_formula")}
        if p.get("apply", True):
            steps.append(f"Restructure '{p['sheet']}' as proposed by the AI: {len(cols)} columns ({n['add']} new, "
                         f"{n['rename']} renamed, {len(p['dropped'])} removed/hidden, "
                         f"{n['formula'] + n['to_formula']} formula(s), {n['to_input']} formula->input); "
                         "template styles, filter, widths and other sheets' references follow the moved columns")
        else:
            steps.append("Keep the template's columns unchanged (fill only)")
        steps.append(f"Fill {len(p['data'])} row(s) and {len(p['kv'])} header field(s) with values quoted from the "
                     "sources; anything without a quote stays blank")
    elif mode == "template" and ready.get("proposal"):
        rows = ready.get("mapping") or []
        n = {k: sum(1 for r in rows if r["status"] == k) for k in ("auto", "confirm", "missing")}
        how = "the AI" if str(ready["proposal"]["mode"]).startswith("azure") else "label matching"
        steps.append(f"Field mapping by {how}: {n['auto']} matched, {n['confirm']} to confirm, {n['missing']} without "
                     "a source value (left blank unless you enter one)")
        steps.append("Fill only confirmed fields, with values read from the sources (formulas and styles untouched)")
    elif mode == "template":
        steps.append("Fill the template's input cells from source facts (formulas and styles untouched)")
    else:
        steps.append("Design the workbook layout (sheets, columns, totals)")
        steps.append("Fill data only from sources/answers; missing values stay blank")
        steps.append("Build a styled .xlsx with live formulas")
    steps.append("Validate the workbook and create revision 1")
    return steps


def _response(state: dict) -> AnalyzeResponse:
    return AnalyzeResponse(status=state["status"], clarification_id=state["id"],
                           questions=[ClarifyQuestion(**q) for q in state["questions"]] if state["status"] ==
                           "needs_clarification" else [], reason=state.get("reason"), plan=state["plan"],
                           mapping=state.get("mapping") or [], facts=state["facts"], warnings=state["warnings"],
                           structure=_structure_view(state))


def _structure_view(state: dict) -> dict | None:
    if not state.get("structure"):
        return None
    from . import restructure
    return restructure.view(restructure.effective(state["structure"]))


APPLY_VI, FILL_VI = "Áp dụng cấu trúc AI đề xuất", "Giữ cột như mẫu, chỉ điền dữ liệu"
APPLY_EN, FILL_EN = "Apply the proposed structure", "Keep the template's columns, fill only"
HIDE_VI, KEEP_VI = "Ẩn cột (công thức khác vẫn dùng được)", "Giữ cột hiển thị"
HIDE_EN, KEEP_EN = "Hide the column (other formulas keep working)", "Keep the column visible"


def _template_structure(state: dict, b: sources.SourceBundle, prompt: str) -> dict | None:
    """Template + prose sources (or an explicit change request): the AI proposes the table's column structure."""
    if config.llm_mode() != "azure" or not b.names or (b.tables and not (
            EDIT_WORDS.search(prompt) or any(EDIT_WORDS.search(t) for t in b.texts.values()))):
        return None
    from . import restructure, template_fill
    try:
        src, _ = template_fill._template_source(state)
        slots = templates.detect_slots(src)
        return restructure.propose(src, slots, b, prompt)
    except Exception as e:                               # noqa: BLE001 - fall back to plain field mapping
        if re.search(r"rate limit|request too large|\b(413|429)\b", str(e), re.I):
            state["warnings"].append("The AI model's per-minute limit (Azure OpenAI quota) was reached, so no column "
                                     "structure was proposed. Wait 1-2 minutes and send the request again; until "
                                     "then only the template's existing fields are filled.")
        else:
            state["warnings"].append(f"The AI could not propose a column structure ({type(e).__name__}: "
                                     f"{str(e)[:400]}); falling back to filling the template's existing fields.")
        return None


SHEETS_VI, NOSHEETS_VI = "Thêm các sheet đề xuất", "Không thêm sheet mới"
SHEETS_EN, NOSHEETS_EN = "Add the proposed sheets", "Do not add new sheets"


def _structure_questions(plan: dict, vi: bool) -> list[ClarifyQuestion]:
    """Always ask for approval of the AI's proposal: the preview below the question is what the file will contain."""
    qs: list[ClarifyQuestion] = []
    changed = (any(c["changes"] not in (["keep"], ["default"]) for c in plan["columns"]) or plan["dropped"]
               or plan.get("labels") or plan.get("sheets"))
    n_add = sum(1 for c in plan["columns"] if "add" in c["changes"])
    und = plan.get("understanding") or {}
    prev = plan.get("preview") or {}
    rows = len(plan.get("data") or [])
    if changed:
        q = _t(vi, f"Mẫu '{plan['sheet']}' chưa khớp nội dung tài liệu. AI đề xuất bảng {len(plan['columns'])} cột "
                   f"({n_add} cột mới, {len(plan['dropped'])} cột bỏ, {len(plan.get('labels') or [])} nhãn đổi), "
                   f"{rows} dòng dữ liệu. Xem bản xem trước bên dưới. Duyệt phương án này?",
               f"The template '{plan['sheet']}' does not fully fit the documents. The AI proposes "
               f"{len(plan['columns'])} columns ({n_add} new, {len(plan['dropped'])} removed, "
               f"{len(plan.get('labels') or [])} labels changed) and {rows} data rows. See the preview below. "
               "Approve this plan?")
    else:
        q = _t(vi, f"Mẫu '{plan['sheet']}' đã khớp tài liệu, không cần đổi cấu trúc. Điền {rows} dòng như bản xem "
                   "trước bên dưới?",
               f"The template '{plan['sheet']}' already fits the documents; no structural change is needed. Fill "
               f"{rows} rows as previewed below?")
    notes = []
    if und.get("fit_reason"):
        notes.append(und["fit_reason"])
    miss = [m["what"] for m in plan.get("missing") or []]
    if miss:
        notes.append(_t(vi, "Thiếu trong tài liệu (sẽ để trống, không tự điền): ",
                        "Missing from the documents (left blank, never invented): ") + "; ".join(miss[:8]))
    if prev and not prev.get("ok"):
        notes.append(_t(vi, "Phương án không áp dụng được an toàn: ", "The plan cannot be applied safely: ")
                     + "; ".join(prev.get("errors") or [])[:300])
    notes.append(_t(vi, "Giao diện mẫu (màu, viền, định dạng, bộ lọc) được giữ. Muốn sửa đề xuất, hãy nhắn thêm yêu "
                        "cầu (Yêu cầu chỉnh sửa) để AI phân tích lại.",
                    "The template's look (colors, borders, formats, filter) is kept. To change the proposal, send "
                    "more instructions (Request changes) and the AI analyzes again."))
    apply_ok = not prev or prev.get("ok") or not plan.get("apply", True)
    opts = [_t(vi, APPLY_VI, APPLY_EN), _t(vi, FILL_VI, FILL_EN)] if apply_ok else [_t(vi, FILL_VI, FILL_EN)]
    qs.append(ClarifyQuestion(id="structure_apply", type="select", required=True, options=opts, default=opts[0],
                              question=q[:300], explanation=" ".join(notes)[:400]))
    if plan.get("sheets"):
        names = ", ".join(f"'{s['name']}'" for s in plan["sheets"])
        qs.append(ClarifyQuestion(
            id="structure_sheets", type="select", required=True,
            options=[_t(vi, SHEETS_VI, SHEETS_EN), _t(vi, NOSHEETS_VI, NOSHEETS_EN)],
            default=_t(vi, SHEETS_VI, SHEETS_EN),
            question=_t(vi, f"Tài liệu yêu cầu thêm phần tổng hợp; AI đề xuất thêm sheet {names} (công thức tham chiếu "
                            "bảng chính). Thêm không?",
                        f"The documents ask for summaries; the AI proposes new sheet(s) {names} with formulas over "
                        "the main table. Add them?")[:300]))
    for d in plan["dropped"]:
        if not d["used_by"] or len(qs) >= 6:
            continue
        qs.append(ClarifyQuestion(
            id=f"drop_{d['src']}", type="select", required=True,
            options=[_t(vi, HIDE_VI, HIDE_EN), _t(vi, KEEP_VI, KEEP_EN)], default=_t(vi, HIDE_VI, HIDE_EN),
            question=_t(vi, f"Cột '{d['header']}' được đề xuất bỏ nhưng vẫn được dùng bởi "
                            f"{', '.join(d['used_by'][:4])}. Xóa hẳn sẽ làm hỏng các công thức đó (#REF!). "
                            "Xử lý thế nào?",
                        f"Column '{d['header']}' is proposed for removal but is used by "
                        f"{', '.join(d['used_by'][:4])}. Deleting it would break those formulas (#REF!). "
                        "What should happen?")[:300]))
    return qs


def _apply_structure_answers(plan: dict, answers: dict) -> None:
    plan["apply"] = answers.get("structure_apply") not in (FILL_VI, FILL_EN)
    plan["sheets_on"] = answers.get("structure_sheets") not in (NOSHEETS_VI, NOSHEETS_EN)
    for d in plan["dropped"]:
        a = answers.get(f"drop_{d['src']}")
        if a in (KEEP_VI, KEEP_EN):
            d["mode"] = "keep"
        elif a in (HIDE_VI, HIDE_EN):
            d["mode"] = "hide"


def approve(state: dict, plan: dict) -> None:
    """The user's choices are final: dry-run them again; only a valid preview counts as approved."""
    from . import restructure, template_fill
    try:
        src, _ = template_fill._template_source(state)
        plan["preview"] = restructure.dry_run(src, plan)
    except Exception as e:                               # noqa: BLE001
        plan["preview"] = {"ok": False, "errors": [f"{type(e).__name__}: {str(e)[:300]}"], "code": "error"}
    plan["approved"] = bool(plan["preview"].get("ok"))
    if not plan["approved"]:
        state["status"] = "blocked"
        state["reason"] = ("The approved plan cannot be applied safely: " +
                           "; ".join(plan["preview"].get("errors") or [])[:500] +
                           " Choose 'fill only' or request changes.")


BLANK_VI, BLANK_EN = "(để trống)", "(leave blank)"
MAX_MAP_QUESTIONS = 7


def _template_mapping(state: dict, b: sources.SourceBundle) -> tuple[dict, Any] | None:
    """Template + sources: the AI's (or deterministic) per-field proposal, built once here and reused by generate."""
    from . import template_fill
    try:
        src, _ = template_fill._template_source(state)
        slots = templates.detect_slots(src)
    except Exception:                                    # noqa: BLE001 - generate reports template problems itself
        return None
    return template_fill.propose(slots, b, config.llm_mode() == "azure"), slots


def _opt(text: str, taken: dict) -> str:
    t, k = text[:90], 2
    while t in taken:
        t, k = f"{text[:85]} ({k})", k + 1
    return t


def _map_questions(prop: dict, slots, b: sources.SourceBundle, vi: bool) -> tuple[list[ClarifyQuestion], dict]:
    """select questions for guesses to confirm and for fields that only have weak candidates."""
    from . import template_fill
    byid = {f.id: f for f in template_fill._kv_facts(b)}
    blank = BLANK_VI if vi else BLANK_EN
    qs: list[ClarifyQuestion] = []
    targets: dict[str, dict] = {}
    for i, d in enumerate(prop["kv"]):
        guess = d["target"] and not d["confident"]
        weak = not d["target"] and d["candidates"]
        if not (guess or weak) or len(qs) >= MAX_MAP_QUESTIONS:
            continue
        options: dict[str, str | None] = {}
        default = blank
        for t in d["candidates"]:
            hit = template_fill.resolve(b, t, byid)
            if not hit:
                continue
            name = byid[t].field if t in byid else _t(vi, "trích nguồn", "quote")
            o = _opt(f"{name} = {template_fill._short(hit[0], 50)}", options)
            options[o] = t
            if t == d["target"]:
                default = o
        if not options:
            continue
        options[blank] = None
        qid = f"map_k{i}"
        targets[qid] = {"kind": "kv", "i": i, "options": options}
        where = f"'{d['label']}' ({d['sheet']}!{d['cell']})"
        qs.append(ClarifyQuestion(
            id=qid, type="select", required=False, options=list(options), default=default,
            question=(_t(vi, f"Ô {where}: lấy giá trị nào?", f"Cell {where}: which value?") if weak else
                      _t(vi, f"Ô {where}: AI đề xuất lấy '{default}'. Đúng không?",
                         f"Cell {where}: the AI suggests '{default}'. Correct?"))[:300],
            explanation=(d.get("reason") or _t(vi, "Tên trường trong nguồn khác nhãn của mẫu nên cần bạn xác nhận; "
                                                   "chọn '(để trống)' nếu nguồn không có.",
                                               "The source field name differs from the template label; pick "
                                               "'(leave blank)' if the sources do not have it."))[:400]))
    tabs = template_fill._tabs(slots)
    for x, pt in enumerate(prop["tables"]):
        st = b.tables[pt["source"]] if pt.get("source") is not None else None
        if st is None or x >= len(tabs):
            continue
        taken = {d["hi"] for d in pt["cols"].values() if d.get("hi") is not None}
        for col, d in pt["cols"].items():
            guess = d.get("hi") is not None and not d.get("confident", True)
            free = [hi for hi in range(len(st.headers)) if hi not in taken or hi == d.get("hi")]
            weak = d.get("hi") is None and d.get("via") != "sequence" and free
            if not (guess or weak) or len(qs) >= MAX_MAP_QUESTIONS:
                continue
            options, default = {}, blank
            for hi in free:
                sample = template_fill._short(next((r[hi] for r in st.rows if r[hi] not in (None, "")), ""), 30)
                o = _opt(f"{st.headers[hi]} ({_t(vi, 'vd', 'e.g.')} {sample})", options)
                options[o] = hi
                if hi == d.get("hi"):
                    default = o
            options[blank] = None
            qid = f"map_t{x}c{col}"
            targets[qid] = {"kind": "col", "x": x, "col": col, "options": options}
            qs.append(ClarifyQuestion(
                id=qid, type="select", required=False, options=list(options), default=default,
                question=_t(vi, f"Cột '{d['header']}' ({tabs[x].sheet}): lấy từ cột nào của '{st.name}'?",
                            f"Column '{d['header']}' ({tabs[x].sheet}): which column of '{st.name}'?")[:300],
                explanation=(d.get("reason") or (
                    _t(vi, f"Tên cột của '{st.name}' khác bảng mẫu nên AI chỉ đoán bảng này là nguồn; "
                           "chọn cột đúng, hoặc để trống hết nếu không phải.",
                       f"'{st.name}' has different column names, so the AI only guesses it is the source; "
                       "pick the right columns, or leave all blank if it is not.") if not pt.get("confident", True)
                    else _t(vi, "Không có cột nguồn trùng tên; để trống nếu không cột nào hợp.",
                            "No source column has this name; leave blank if none fits.")))[:400]))
    return qs, targets


def _apply_map_answers(state: dict, answers: dict) -> None:
    prop = state["proposal"]
    for qid, mq in (state.get("map_targets") or {}).items():
        if answers.get(qid) not in mq["options"]:
            continue
        t = mq["options"][answers[qid]]
        if mq["kind"] == "kv":
            d, key = prop["kv"][mq["i"]], "target"
        else:
            d, key = prop["tables"][mq["x"]]["cols"][mq["col"]], "hi"
        changed = t != d.get(key)
        d[key], d["confident"] = t, True
        d["via"] = None if t is None else ("user" if changed else d.get("via"))
        if changed:
            d["reason"] = ""
    for d in prop["kv"]:                                 # guesses that were never asked are not filled
        if d["target"] and not d["confident"]:
            d.update(target=None, via=None, confident=True)
    for x, pt in enumerate(prop["tables"]):
        for d in pt["cols"].values():
            if d.get("hi") is not None and not d.get("confident", True):
                d.update(hi=None, via=None, confident=True)
        if not pt.get("confident", True):                # a guessed source table counts once a column is chosen
            asked = any(m["kind"] == "col" and m["x"] == x and answers.get(q) in m["options"]
                        for q, m in (state.get("map_targets") or {}).items())
            pt["confident"] = asked and any(d.get("hi") is not None for d in pt["cols"].values())


def _mapping_rows(state: dict, b: sources.SourceBundle | None = None, slots=None) -> list[dict]:
    from . import template_fill
    try:
        if slots is None:
            src, _ = template_fill._template_source(state)
            slots = templates.detect_slots(src)
        if b is None:
            b = sources.load(state["source_file_ids"], state["prompt"])
            sources.answer_figures(b, (state.get("answers") or {}).get("figures") or "", keep_text=True)
        return template_fill.describe(slots, b, state["proposal"])
    except Exception:                                    # noqa: BLE001
        return []


def analyze(ws_id: str, req: AnalyzeRequest) -> AnalyzeResponse:
    ws = workspaces.get(ws_id)
    attached = {f["id"]: f for f in ws["files"]}
    src_ids = req.source_file_ids or [f["id"] for f in ws["files"] if f["role"] == "source"]
    for fid in src_ids:
        if attached.get(fid, {}).get("role") != "source":
            raise InvalidOperation(f"Source file {fid} is not attached to this workspace.")
    template_file_id = req.template_file_id or next((f["id"] for f in ws["files"] if f["role"] == "template"), None)
    if template_file_id and attached.get(template_file_id, {}).get("role") != "template":
        raise InvalidOperation("template_file_id is not an attached template.")
    template_id = req.template_id or ws.get("template_id")
    if template_id and not template_file_id:
        templates.get(template_id)                       # 404 for an unknown gallery id
    b = sources.load(src_ids, req.prompt)
    state: dict[str, Any] = {
        "id": storage.new_id("cl"), "prompt": req.prompt, "source_file_ids": src_ids,
        "template_file_id": template_file_id, "template_id": template_id, "output_filename": req.output_filename,
        "answers": {}, "facts": len(b.facts), "warnings": list(b.warnings), "reason": None,
        "currency": b.currency, "period": b.period}
    if ws["mode"] == "template" and not (template_file_id or template_id):
        state.update(status="blocked", questions=[], plan=[],
                     reason="Template mode needs a template: upload an .xlsx or pick one from the gallery.")
    elif not req.prompt.strip() and not src_ids:
        state.update(status="blocked", questions=[], plan=[],
                     reason="Describe the workbook you want or attach at least one source file.")
    else:
        missing, mq = None, []
        structure = _template_structure(state, b, req.prompt) if ws["mode"] == "template" else None
        made = _template_mapping(state, b) if ws["mode"] == "template" and b.names and not structure else None
        if structure:
            vi = is_vi(req.prompt) or any(is_vi(t) for t in b.texts.values())
            state["structure"] = structure
            state["warnings"] += structure["warnings"]
            if req.auto_approve:
                _apply_structure_answers(structure, {})
                approve(state, structure)
            else:
                structure["approved"] = False
                mq = _structure_questions(structure, vi)
            missing = []
        if made:
            prop, slots = made
            vi = is_vi(req.prompt) or any(is_vi(t) for t in b.texts.values())
            mq, state["map_targets"] = _map_questions(prop, slots, b, vi)
            asked = {m["i"] for m in state["map_targets"].values() if m["kind"] == "kv"}
            missing = [d["label"] for i, d in enumerate(prop["kv"]) if not d["target"] and i not in asked]
            state["proposal"], state["warnings"] = prop, state["warnings"] + prop["warnings"]
            state["mapping"] = _mapping_rows(state, b, slots)
        if ws["mode"] == "template" and not b.names and not structure and mock_fill.wants_mock(req.prompt):
            mq, mock_warns = _mock_questions(state, req.prompt)
            state["warnings"] += mock_warns
        qs, warns = _questions(req.prompt, b, ws["mode"], missing)
        if state.get("mock"):
            qs = [q for q in qs if q.id != "figures"]
        qs = mq + qs
        state["warnings"] += warns
        state["questions"] = [q.model_dump() for q in qs]
        if state.get("status") != "blocked":
            state["status"] = "needs_clarification" if qs else "ready"
        state["plan"] = _plan(ws["mode"], b, state)
    ws["prompt"], ws["template_id"] = req.prompt, template_id
    ws["clarification"] = state
    workspaces.save(ws)
    return _response(state)


def clarify(ws_id: str, body: ClarifyAnswers) -> AnalyzeResponse:
    ws = workspaces.get(ws_id)
    state = ws.get("clarification")
    if not state or state["id"] != body.clarification_id:
        raise InvalidOperation("Unknown or outdated clarification_id; run analyze again.")
    qs = {q["id"]: q for q in state["questions"]}
    unknown = set(body.answers) - set(qs)
    if unknown:
        raise InvalidOperation(f"Unknown question id(s): {sorted(unknown)}")
    answers: dict[str, Any] = {}
    for qid, q in qs.items():
        v = body.answers.get(qid, q.get("default"))
        if isinstance(v, str):
            v = v.strip()[:mock_fill.MAX_ANSWER if qid == "mock_cells" else 4000]
        if v in (None, "") and q["required"]:
            raise InvalidOperation(f"Question '{qid}' needs an answer.")
        if q["type"] == "select" and v not in (None, "") and v not in (q["options"] or []):
            raise InvalidOperation(f"'{v}' is not an option for '{qid}'.")
        if q["type"] == "number" and v not in (None, ""):
            try:
                v = float(v)
            except (TypeError, ValueError):
                raise InvalidOperation(f"'{qid}' must be a number.") from None
        if q["type"] == "boolean" and v not in (None, ""):
            v = bool(v)
        answers[qid] = v
    if answers.get("missing_numbers") == "provide" and not answers.get("figures"):
        state["warnings"].append("You chose to provide figures but none were entered; input cells stay blank.")
    state["answers"], state["status"] = answers, "ready"
    if state.get("structure"):
        _apply_structure_answers(state["structure"], answers)
        approve(state, state["structure"])
        state["plan"] = state["plan"][:1] + _plan(ws["mode"], sources.SourceBundle(), state)[1:]
    if state.get("proposal"):
        _apply_map_answers(state, answers)
        state["mapping"] = _mapping_rows(state)
        state["plan"] = state["plan"][:1] + _plan(ws["mode"], sources.SourceBundle(), state)[1:]
    ws["clarification"] = state
    workspaces.save(ws)
    return _response(state)
