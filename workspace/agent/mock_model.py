"""Deterministic stand-in for the chat model when Azure OpenAI is not configured (mode 'mock').

It is NOT an LLM: it recognises a few explicit commands (VI/EN) and turns them into the same tool calls the real
model would make, so the agent loop, tools, plans, revisions and SSE events are exercised end to end. Anything it
does not understand gets a help text listing the supported commands; it never invents values.
"""
from __future__ import annotations

import json
import re
import uuid
from typing import Any

from langchain_core.language_models import BaseChatModel
from langchain_core.messages import AIMessage, HumanMessage, SystemMessage, ToolMessage
from langchain_core.outputs import ChatGeneration, ChatResult

REF = r"(?:(?:'([^']+)'|([A-Za-z_][\w .]*?))!)?\$?([A-Za-z]{1,3})\$?(\d{1,7})"
RANGE = REF + r"(?:\s*:\s*\$?([A-Za-z]{1,3})\$?(\d{1,7}))?"
SET_RE = re.compile(r"(?:đặt|dat|ghi|điền|dien|nhập|nhap|sửa|sua|set|put|write|change)\s+(?:giá trị\s+)?(?:ô|o|cell)?\s*"
                    + REF + r"\s*(?:=|thành|thanh|là|la|bằng|bang|to|:)\s*(.+)", re.I)
CLEAR_RE = re.compile(r"(?:xóa|xoá|xoa|clear|delete|empty)\s+(?:nội dung\s+)?(?:vùng|vung|ô|o|range|cells?)?\s*" + RANGE,
                      re.I)
READ_RE = re.compile(r"(?:đọc|doc|xem|read|show|get)\s+(?:vùng|vung|ô|o|range|cells?)?\s*" + RANGE, re.I)
INSPECT_RE = re.compile(r"tóm tắt|tom tat|mô tả|mo ta|cấu trúc|cau truc|inspect|describe|summari[sz]e|overview", re.I)
VALIDATE_RE = re.compile(r"kiểm tra|kiem tra|validate|check", re.I)
BLANK_RE = re.compile(r"template trống|template trong|mẫu trống|mau trong|blank template", re.I)
CONFIRM_RE = re.compile(r"^\s*(?:đồng ý|dong y|ok(?:ay)?|yes|apply|áp dụng|ap dung|xác nhận|xac nhan|confirm)\b", re.I)
WRITE_INTENT_RE = re.compile(r"^\s*(?:hãy\s+|giúp\s+(?:tôi|mình|em)\s+)?(?:đặt|ghi|điền|dien|nhập|nhap|sửa|sua|cập nhật|"
                             r"thêm|them|set|put|write|fill|update|change|add)\b", re.I)
UNSUPPORTED_RE = re.compile(r"công thức|cong thuc|formula|biểu đồ|bieu do|chart|pivot|macro|chèn\s+(?:dòng|cột)|"
                            r"insert\s+(?:rows?|columns?)|định dạng|format", re.I)
PLAN_RE = re.compile(r"Pending plan: (plan_[0-9a-f]{16})")
VI_RE = re.compile(r"[ăâđêôơưàáảãạèéẻẽẹìíỉĩịòóỏõọùúủũụỳýỷỹỵ]", re.I)

HELP_VI = ("Chế độ mock (chưa cấu hình Azure OpenAI) chỉ hiểu các lệnh rõ ràng: \"đặt B2 = 120\", \"xóa A2:C9\", "
           "\"đọc A1:D10\", \"tóm tắt workbook\", \"kiểm tra workbook\", \"đồng ý\" (áp dụng kế hoạch đang chờ). "
           "Hãy cấu hình AZURE_OPENAI_API_KEY và AZURE_OPENAI_ENDPOINT để chat tự do.")
ASK_VI = ("Bạn muốn ghi vào ô nào và giá trị cụ thể là bao nhiêu? Mình không tự đoán số liệu. "
          "Ví dụ: \"đặt B4 = 120\" hoặc \"ghi 'Bán hàng'!D5 = 3.500.000\". Nếu chưa biết ô, gõ \"tóm tắt workbook\" "
          "để xem cấu trúc trước.")
ASK_EN = ("Which cell should I write, and what exact value? I never guess figures. Example: \"set B4 = 120\". "
          "If you are not sure of the cell, type \"summarize workbook\" first.")
UNSUPPORTED_VI = ("Trình sửa hiện chỉ ghi hoặc xóa giá trị ô. Ghi công thức, định dạng, chèn dòng/cột, biểu đồ, "
                  "pivot và macro chưa được hỗ trợ, nên mình không thực hiện.")
UNSUPPORTED_EN = ("The editor can only set or clear cell values. Formulas, formatting, inserting rows/columns, charts, "
                  "pivots and macros are not supported yet, so nothing was changed.")
HELP_EN = ("Mock mode (Azure OpenAI not configured) understands explicit commands only: \"set B2 = 120\", \"clear A2:C9\", "
           "\"read A1:D10\", \"summarize workbook\", \"check workbook\", \"yes\" (apply the pending plan). "
           "Set AZURE_OPENAI_API_KEY and AZURE_OPENAI_ENDPOINT for free-form chat.")


def _sheet(m: re.Match, i: int = 1) -> str | None:
    return (m.group(i) or m.group(i + 1) or "").strip() or None


def _range(m: re.Match) -> str:
    a = f"{m.group(3)}{m.group(4)}".upper()
    return f"{a}:{m.group(5)}{m.group(6)}".upper() if m.group(5) else a


def _call(name: str, args: dict) -> AIMessage:
    return AIMessage(content="", tool_calls=[{"name": name, "args": args, "id": f"call_{uuid.uuid4().hex[:12]}"}])


def _value(raw: str) -> tuple[Any, str]:
    v = raw.strip().rstrip(".").strip()
    if len(v) >= 2 and v[0] == v[-1] and v[0] in "\"'“”":
        return v[1:-1], "string"
    if re.fullmatch(r"\d{4}-\d{2}-\d{2}", v):
        return v, "date"
    return v, "auto"


class MockChatModel(BaseChatModel):
    model_name: str = "excel-mock"
    bound_tools: list[str] = []

    @property
    def _llm_type(self) -> str:
        return "excel-mock"

    def bind_tools(self, tools, **kw):
        names = []
        for t in tools:
            n = getattr(t, "name", None) or (t.get("name") if isinstance(t, dict) else None)
            n = n or (t.get("function", {}).get("name") if isinstance(t, dict) else None)
            names.append(n)
        type(self).bound_tools = sorted(names)               # inspected by tests: no FS/shell tools exposed
        return self

    def _generate(self, messages, stop=None, run_manager=None, **kw) -> ChatResult:
        return ChatResult(generations=[ChatGeneration(message=self._next(messages))])

    def _next(self, messages: list) -> AIMessage:
        system = next((str(m.content) for m in messages if isinstance(m, SystemMessage)), "")
        user = next((m for m in reversed(messages) if isinstance(m, HumanMessage)), None)
        text = str(user.content) if user else ""
        vi = bool(VI_RE.search(text)) or "Reply in Vietnamese" in system
        last = messages[-1]
        if isinstance(last, ToolMessage):
            return AIMessage(content=self._final(last, vi))
        text = text.strip()
        plan = PLAN_RE.search(system)
        if plan and CONFIRM_RE.search(text):
            return _call("apply_workbook_edits", {"plan_id": plan.group(1)})
        if BLANK_RE.search(text):
            return _call("run_template_extractor", {})
        edits = []
        for part in re.split(r"\s*(?:;|\n|,\s*(?=(?:đặt|ghi|điền|set|put|write)\b)|\bvà\b|\band\b)\s*", text):
            m = SET_RE.search(part)
            if m:
                value, vt = _value(m.group(5))
                edits.append({"op": "set", "sheet": _sheet(m), "cell": f"{m.group(3)}{m.group(4)}".upper(),
                              "value": value, "value_type": vt})
                continue
            m = CLEAR_RE.search(part)
            if m:
                edits.append({"op": "clear", "sheet": _sheet(m), "cell": _range(m)})
        if edits:
            return _call("propose_workbook_edits", {"edits": edits, "reason": text[:200]})
        m = READ_RE.search(text)
        if m:
            return _call("read_range", {"range": _range(m), "sheet": _sheet(m)})
        if INSPECT_RE.search(text):
            return _call("inspect_workbook", {})
        if VALIDATE_RE.search(text):
            return _call("validate_workbook", {})
        if WRITE_INTENT_RE.search(text) and UNSUPPORTED_RE.search(text):
            return AIMessage(content=UNSUPPORTED_VI if vi else UNSUPPORTED_EN)
        if WRITE_INTENT_RE.search(text):                     # wants a change but gave no cell/value: ask, never guess
            return AIMessage(content=ASK_VI if vi else ASK_EN)
        return AIMessage(content=HELP_VI if vi else HELP_EN)

    @staticmethod
    def _final(msg: ToolMessage, vi: bool) -> str:
        try:
            out = json.loads(str(msg.content))
        except (ValueError, TypeError):
            return str(msg.content)[:500]
        if not isinstance(out, dict):
            return str(out)[:500]
        if out.get("error"):
            return ("Không thực hiện được: " if vi else "Could not do it: ") + str(out["error"])
        if out.get("message"):
            return str(out["message"])
        if "cells" in out and isinstance(out["cells"], list):
            shown = ", ".join(f"{c['cell']}={c['value']}" for c in out["cells"][:20])
            more = f" (+{len(out['cells']) - 20})" if len(out["cells"]) > 20 else ""
            return (f"{out['sheet']}!{out['range']}: " + (shown + more if shown else
                                                         ("vùng trống" if vi else "empty range")))
        if "sheets" in out:
            names = ", ".join(f"{s['name']} ({s['dimensions']})" for s in out["sheets"])
            return (f"Workbook có {len(out['sheets'])} sheet: {names}." if vi else
                    f"The workbook has {len(out['sheets'])} sheet(s): {names}.")
        if "cached_error_cells" in out:
            errs = out["cached_error_cells"]
            if vi:
                return f"Revision {out['revision']}: " + (f"{len(errs)} ô lỗi: {', '.join(errs[:10])}" if errs
                                                          else "không thấy ô lỗi.")
            return f"Revision {out['revision']}: " + (f"{len(errs)} error cell(s): {', '.join(errs[:10])}" if errs
                                                      else "no error values found.")
        return json.dumps(out, ensure_ascii=False)[:500]
