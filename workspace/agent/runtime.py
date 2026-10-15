"""Runs one chat turn: a Deep Agents orchestrator over the tool allowlist in tools.py.

The built-in Deep Agents filesystem/shell/todo/sub-agent tools are excluded through a harness profile, so the model
only sees the workbook tools. Azure OpenAI when it is configured (AZURE_OPENAI_*), otherwise the
deterministic MockChatModel. Each turn is a background run streamed over SSE.
"""
from __future__ import annotations

import os
import re
import threading

from langchain_core.messages import AIMessage, AIMessageChunk, HumanMessage

from .. import config, runs, service
from ..errors import RevisionConflict, WorkspaceError
from ..schemas import AgentTurnRequest
from . import store, tools
from .mock_model import VI_RE, MockChatModel

EXCLUDED_TOOLS = frozenset({"ls", "read_file", "write_file", "edit_file", "delete", "glob", "grep", "execute",
                            "write_todos", "task"})
RECURSION_LIMIT = 24
# A reply that claims a change although no tool wrote anything this turn (seen live: "Đã xóa ... (revision 6)").
CLAIM_RE = re.compile(r"\b(?:đã|vừa)\s+(?:được\s+)?(?:ghi|xóa|xoá|cập nhật|điền|sửa|đặt|thay đổi|áp dụng|thêm|nhập)"
                      r"|\b(?:updated|cleared|wrote|written|deleted|changed|applied|filled|set)\b"
                      r"[^.\n]{0,60}\b(?:cells?|range|[A-Z]{1,3}\d+)\b", re.I)
_registered: set[str] = set()
_lock = threading.Lock()

SYSTEM = """You are the assistant inside an Excel workspace. You work on ONE workbook through tools only.

Rules:
- Never invent numbers, dates or names. Every value you write must come from the user's messages, the attached
  source files (read_source) or the workbook itself (read_range). If something is missing, ask the user.
- Ask before writing when the request is vague: if the user did not say which cells/rows to change or did not give
  the values (e.g. "fill October revenue" with no figures), you may read the workbook, then reply with ONE short
  question (which cells? which values?) and call NO write tool. Never reinterpret a request into a different edit
  (e.g. renaming a title or a period label) just to do something.
- Converting or scaling existing numbers (currency conversion such as VND -> USD at a rate the user gave, a % uplift,
  a unit change) is done with transform_range: give the range, the op and the user's rate; the code computes the
  values. Never compute such values yourself and pass them to propose_workbook_edits (they are refused as
  ungrounded). If a tool refuses something, do not retry the same call: explain the refusal to the user.
- "This cell/column/row" ("ô này", "cột này", "dòng này") means the Selection below. For a column, use the
  selected column from the first data row under its header to the last filled row (read_range to find them). Once
  the user has given the target and the values/rate, act; do not ask again. A target you named in your own earlier
  question and the user did not reject counts as given. "ok"/"xác nhận"/"làm đi" answering your own question or
  proposal means: do it now. Do not ask about details with a sensible default (rounding: 2 decimals).
- Inspect before editing: call inspect_workbook and read_range to find the right cells; never guess addresses.
- Write only with propose_workbook_edits. Edits of up to {auto} cells are applied immediately; bigger edits wait for
  the user's approval in the UI. Call apply_workbook_edits only if the user's latest message explicitly approves a
  pending plan.
- The editor can set or clear cell values only. Formulas, styles, inserting rows/columns, charts, pivots and macros
  are not supported: say so instead of attempting them. Do not write text that starts with "=".
- build_workbook creates a separate new workbook; use it only if the user asks for a new workbook.
- Never say that cells were changed unless propose_workbook_edits or apply_workbook_edits returned a new revision
  in THIS turn. To change anything you must call the tool; describing a change does not perform it.
- Keep answers short. After editing, say which cells changed and the new revision.
{lang}

Workbook: {name} (id {wb}), revision {rev}. Sheets: {sheets}.
Active sheet: {active}. Selection: {selection}.
{pending}"""


def _profile(model) -> None:
    from deepagents import HarnessProfile, register_harness_profile
    from deepagents._models import get_model_identifier, get_model_provider
    from deepagents.profiles.harness.harness_profiles import GeneralPurposeSubagentProfile
    key = f"{get_model_provider(model)}:{get_model_identifier(model)}"
    with _lock:
        if key not in _registered:
            register_harness_profile(key, HarnessProfile(
                excluded_tools=EXCLUDED_TOOLS,
                general_purpose_subagent=GeneralPurposeSubagentProfile(enabled=False)))
            _registered.add(key)


def _models() -> list[tuple[str, object]]:
    if config.llm_mode() != "azure":
        return [("mock", MockChatModel())]
    return config.chat_models(timeout=60)


def _system(ctx: tools.Ctx, meta: dict, req: AgentTurnRequest, pending: dict | None) -> str:
    return SYSTEM.format(
        auto=tools.AUTO_APPLY_MAX_CELLS,
        lang="Reply in Vietnamese." if ctx.vi else "Reply in the user's language.",
        name=meta["name"], wb=meta["id"], rev=ctx.revision, sheets=", ".join(meta.get("sheets") or []),
        active=req.active_sheet or "(first sheet)", selection=req.selection or "(none)",
        pending=(f"Pending plan: {pending['id']} ({len(pending['actions'])} actions, made for revision "
                 f"{pending['base_revision']}) awaiting the user's approval." if pending else "No pending plan."))


def _pending(thread: dict, wb_id: str, rev: int) -> dict | None:
    pid = thread.get("pending_plan_id")
    if not pid:
        return None
    try:
        p = store.get_plan(pid)
    except WorkspaceError:
        return None
    return p if p["status"] == "pending" and p["workbook_id"] == wb_id and p["base_revision"] == rev else None


def start_turn(req: AgentTurnRequest) -> dict:
    meta = service.get_metadata(req.workbook_id)
    if req.revision != meta["current_revision"]:
        raise RevisionConflict(f"Workbook is at revision {meta['current_revision']}, not {req.revision}.",
                               current_revision=meta["current_revision"])
    thread = store.get_thread(req.thread_id, req.workbook_id)
    store.save_thread(thread)
    history = store.recent_messages(thread)
    user_text = "\n".join([m["content"] for m in history if m["role"] == "user"] + [req.message])

    def work(run: runs.Run) -> dict:
        ctx = tools.Ctx(run=run, workbook_id=req.workbook_id, revision=req.revision, thread_id=thread["id"],
                        message=req.message, user_text=user_text, active_sheet=req.active_sheet,
                        vi=bool(VI_RE.search(req.message)))
        pending = _pending(thread, req.workbook_id, req.revision)
        msgs = [HumanMessage(m["content"]) if m["role"] == "user" else AIMessage(m["content"]) for m in history]
        msgs.append(HumanMessage(req.message))
        text, used, last_err = "", None, None
        for label, model in _models():
            try:
                text = _invoke(run, model, tools.make_tools(ctx), _system(ctx, meta, req, pending), msgs)
                used = label
                break
            except WorkspaceError:
                raise
            except Exception as e:  # model/transport failure: retry on the fallback only if nothing was written
                last_err = e
                if ctx.mutations or ctx.new_workbook_id or ctx.pending_plan_id:
                    break
                run.emit("model_fallback", {"from": label, "error": type(e).__name__, "detail": str(e)[:200]})
        if used is None:
            if ctx.mutations:                                  # the write succeeded before the model failed
                msg = (f"Đã ghi thay đổi (workbook giờ ở revision {ctx.revision}) nhưng model không phản hồi tiếp; "
                       "hãy kiểm tra lại các ô vừa đổi." if ctx.vi else
                       f"The change was written (workbook is now at revision {ctx.revision}) but the model did not "
                       "respond afterwards; please check the changed cells.")
            elif ctx.pending_plan_id:
                msg = ("Đã tạo kế hoạch sửa đang chờ xác nhận nhưng model không phản hồi tiếp." if ctx.vi else
                       "A plan is waiting for confirmation but the model did not respond afterwards.")
            elif ctx.tool_errors:                               # e.g. the model kept retrying a refused edit
                why = ctx.tool_errors[-1]
                msg = (f"Không thực hiện được yêu cầu, workbook không đổi (vẫn ở revision {ctx.revision}). "
                       f"Lý do từ công cụ: {why}" if ctx.vi else
                       f"The request could not be carried out; the workbook is unchanged (revision {ctx.revision}). "
                       f"Tool said: {why}")
            else:
                msg = ("Model không phản hồi; không có thay đổi nào ngoài những gì đã báo." if ctx.vi else
                       "The model did not respond; nothing changed beyond what was reported.")
            run.emit("text_delta", {"text": msg})
            text = msg
            store.log(thread["id"], {"turn": run.id, "error": type(last_err).__name__})
        elif not (ctx.mutations or ctx.pending_plan_id or ctx.new_workbook_id) and CLAIM_RE.search(text):
            note = ("⚠️ Lưu ý: lượt này KHÔNG có thay đổi nào được áp dụng vào workbook (không có công cụ ghi nào "
                    f"chạy; vẫn ở revision {ctx.revision}). Hãy yêu cầu lại." if ctx.vi else
                    "⚠️ Note: NO change was applied to the workbook in this turn (no write tool ran; still at "
                    f"revision {ctx.revision}). Please ask again.")
            run.emit("text_delta", {"text": "\n\n" + note})
            text = f"{text}\n\n{note}"
            store.log(thread["id"], {"turn": run.id, "unbacked_claim": True})
        thread["messages"] += [{"role": "user", "content": req.message, "t": run.created},
                               {"role": "assistant", "content": text, "turn_id": run.id, "model": used,
                                "revision": ctx.revision}]
        if ctx.pending_plan_id:
            thread["pending_plan_id"] = ctx.pending_plan_id
        elif ctx.mutations:
            thread.pop("pending_plan_id", None)
        store.save_thread(thread)
        out = {"text": text, "workbook_id": req.workbook_id, "revision": ctx.revision, "thread_id": thread["id"],
               "model": used}
        if ctx.new_workbook_id:
            out["new_workbook_id"] = ctx.new_workbook_id
        return out

    run = runs.start("agent", work, prefix="turn")
    return {"turn_id": run.id, "run_id": run.id, "thread_id": thread["id"]}


def _invoke(run: runs.Run, model, tool_list: list, system: str, msgs: list) -> str:
    from deepagents import create_deep_agent
    _profile(model)
    agent = create_deep_agent(model=model, tools=tool_list, system_prompt=system)
    final, streamed = "", False
    for chunk, _meta in agent.stream({"messages": msgs}, config={"recursion_limit": RECURSION_LIMIT},
                                     stream_mode="messages"):
        if not isinstance(chunk, (AIMessage, AIMessageChunk)):
            continue                                         # tool results are reported by tool_completed
        if chunk.tool_calls or getattr(chunk, "tool_call_chunks", None):
            final, streamed = "", False                      # text before a tool call is not the final answer
            continue
        piece = chunk.content if isinstance(chunk.content, str) else "".join(
            p.get("text", "") for p in chunk.content if isinstance(p, dict))
        if piece:
            final = final + piece if isinstance(chunk, AIMessageChunk) else piece
            streamed = True
            run.emit("text_delta", {"text": piece})
    return final.strip() if streamed else ""
