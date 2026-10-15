import type { ChatModelAdapter, ChatModelRunResult, ThreadAssistantMessagePart } from "@assistant-ui/react";
import { ApiError, api, errorMessage } from "@/lib/api";
import { iterateEventStream } from "@/lib/sse";
import { qualify } from "@/lib/a1";
import { useWorkbookStore } from "@/lib/store";
import type {
  ConfirmationRequiredData,
  RunDoneData,
  RunErrorData,
  ToolCompletedData,
  ToolStartedData,
  WorkbookRevisionData,
} from "@/lib/types";

export const CONFIRM_TOOL = "confirm_changes";
export const UNAVAILABLE_TEXT = "Agent backend unavailable";

/** Result stored on a backend tool card. `pending` while the tool is still running. */
export interface ToolCardResult {
  status: "running" | "ok" | "error";
  summary?: string;
}

type MutablePart =
  | { type: "text"; text: string }
  | { type: "tool-call"; toolCallId: string; toolName: string; args: Record<string, unknown>; result: unknown; isError?: boolean };

function userText(content: readonly { type: string; text?: string }[]): string {
  return content
    .filter((p) => p.type === "text")
    .map((p) => p.text ?? "")
    .join("\n")
    .trim();
}

/**
 * Bridges assistant-ui's LocalRuntime to the backend agent:
 *   POST /api/agent/turns  ->  SSE /api/agent/turns/{turn_id}/events
 *
 * Every yield is a full snapshot of the assistant message. Every tool-call part always carries a
 * `result`, and the final yield is `complete/stop`, so LocalRuntime never auto-continues the run.
 */
export function createAgentAdapter(workbookId: string): ChatModelAdapter {
  let threadId: string | null = null;

  return {
    async *run({ messages, abortSignal }) {
      const last = messages[messages.length - 1];
      const message = last?.role === "user" ? userText(last.content as never) : "";
      const st = useWorkbookStore.getState();

      const parts: MutablePart[] = [];
      const toolIndex = new Map<string, number>(); // tool name -> index of its latest running card
      const snapshot = (): ThreadAssistantMessagePart[] =>
        parts.map((p) =>
          p.type === "text"
            ? { type: "text", text: p.text }
            : {
                type: "tool-call",
                toolCallId: p.toolCallId,
                toolName: p.toolName,
                args: p.args as never,
                argsText: JSON.stringify(p.args),
                result: p.result,
                isError: p.isError,
              },
        ) as ThreadAssistantMessagePart[];
      const appendText = (t: string) => {
        const tail = parts[parts.length - 1];
        if (tail?.type === "text") tail.text += t;
        else parts.push({ type: "text", text: t });
      };
      const settleRunningTools = (status: ToolCardResult["status"], summary: string) => {
        for (const p of parts)
          if (p.type === "tool-call" && p.toolName !== CONFIRM_TOOL && (p.result as ToolCardResult)?.status === "running")
            p.result = { status, summary } satisfies ToolCardResult;
      };

      let turnId: string;
      try {
        const res = await api.startTurn(
          {
            workbook_id: workbookId,
            revision: st.revision,
            message,
            thread_id: threadId,
            selection: st.selection && st.activeSheet ? qualify(st.activeSheet, st.selection) : st.selection,
            active_sheet: st.activeSheet,
          },
          { signal: abortSignal },
        );
        threadId = res.thread_id ?? threadId;
        turnId = res.turn_id;
      } catch (e) {
        if (abortSignal.aborted) return;
        const text =
          e instanceof ApiError && e.routeMissing ? UNAVAILABLE_TEXT : `${UNAVAILABLE_TEXT}: ${errorMessage(e)}`;
        yield { content: [{ type: "text", text }], status: { type: "complete", reason: "stop" } };
        return;
      }

      let finalStatus: ChatModelRunResult["status"] = { type: "complete", reason: "stop" };
      try {
        for await (const env of iterateEventStream(api.turnEventsUrl(turnId), abortSignal)) {
          const d = env.data as Record<string, unknown>;
          switch (env.event) {
            case "text_delta": {
              const t = (d.text ?? d.delta ?? "") as string;
              if (t) appendText(t);
              break;
            }
            case "tool_started": {
              const data = d as unknown as ToolStartedData;
              parts.push({
                type: "tool-call",
                toolCallId: `${data.tool}-${env.seq}`,
                toolName: data.tool,
                args: (data.args ?? {}) as Record<string, unknown>,
                result: { status: "running" } satisfies ToolCardResult,
              });
              toolIndex.set(data.tool, parts.length - 1);
              break;
            }
            case "tool_completed": {
              const data = d as unknown as ToolCompletedData;
              const ok = data.status === "ok";
              const result: ToolCardResult = { status: ok ? "ok" : "error", summary: data.summary };
              const idx = toolIndex.get(data.tool);
              const existing = idx !== undefined ? parts[idx] : undefined;
              if (existing?.type === "tool-call") {
                existing.result = result;
                existing.isError = !ok;
                toolIndex.delete(data.tool);
              } else {
                parts.push({
                  type: "tool-call",
                  toolCallId: `${data.tool}-${env.seq}`,
                  toolName: data.tool,
                  args: {},
                  result,
                  isError: !ok,
                });
              }
              break;
            }
            case "confirmation_required": {
              const data = d as unknown as ConfirmationRequiredData;
              parts.push({
                type: "tool-call",
                toolCallId: `confirm-${data.plan_id}`,
                toolName: CONFIRM_TOOL,
                args: data as unknown as Record<string, unknown>,
                // The decision is made through REST (apply/reject), not by resuming this run.
                result: { awaiting: "user" },
              });
              break;
            }
            case "workbook_revision": {
              const data = d as unknown as WorkbookRevisionData;
              if (data.workbook_id === workbookId) useWorkbookStore.getState().setRevision(data.revision);
              break;
            }
            case "done": {
              const data = d as RunDoneData;
              if (data.text && !parts.some((p) => p.type === "text")) appendText(data.text);
              if (typeof data.revision === "number" && data.workbook_id === workbookId)
                useWorkbookStore.getState().setRevision(data.revision);
              break;
            }
            case "error": {
              const data = d as unknown as RunErrorData;
              const msg = data.message || data.code || "The agent run failed.";
              settleRunningTools("error", "Interrupted");
              appendText(`${parts.length ? "\n\n" : ""}**Error:** ${msg}`);
              finalStatus = { type: "incomplete", reason: "error", error: msg };
              break;
            }
            case "cancelled":
              settleRunningTools("error", "Cancelled");
              finalStatus = { type: "incomplete", reason: "cancelled" };
              break;
            default:
              continue;
          }
          yield { content: snapshot(), status: { type: "running" } };
        }
      } catch (e) {
        if (!abortSignal.aborted) {
          settleRunningTools("error", "Connection lost");
          appendText(`${parts.length ? "\n\n" : ""}**Error:** ${errorMessage(e)}`);
          finalStatus = { type: "incomplete", reason: "error", error: errorMessage(e) };
        }
      }

      if (abortSignal.aborted) {
        settleRunningTools("error", "Cancelled");
        finalStatus = { type: "incomplete", reason: "cancelled" };
      }
      settleRunningTools("ok", "");
      if (!parts.length) appendText("(No response)");
      yield { content: snapshot(), status: finalStatus };
    },
  };
}
