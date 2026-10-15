import type { SseEnvelope } from "./types";

/** Every event name the backend may emit on run / agent-turn streams. */
export const SSE_EVENTS = [
  "run_started",
  "analysis_ready",
  "text_delta",
  "tool_started",
  "tool_completed",
  "confirmation_required",
  "clarification_needed",
  "clarification_accepted",
  "workbook_revision",
  "validation_failed",
  "done",
  "error",
  "cancelled",
] as const;

const TERMINAL = new Set(["done", "error", "cancelled"]);
const MAX_SILENT_RECONNECTS = 5;

export interface StreamHandlers {
  onEvent: (env: SseEnvelope) => void;
  /** Stream could not be (re)established; no terminal event was received. */
  onFailure?: (message: string) => void;
}

export interface StreamHandle {
  close: () => void;
}

/**
 * Subscribe to a backend SSE stream. EventSource resends `Last-Event-ID` itself on reconnect
 * (the server emits `id: <seq>` and `retry: 2000`). Closes on done / error / cancelled, and gives up
 * after several reconnects that deliver nothing new.
 */
export function openEventStream(url: string, handlers: StreamHandlers): StreamHandle {
  const es = new EventSource(url);
  let closed = false;
  let silentReconnects = 0;
  let lastSeq = 0;

  const close = () => {
    if (closed) return;
    closed = true;
    es.close();
  };

  const listener = (ev: MessageEvent<string>) => {
    if (closed) return;
    silentReconnects = 0;
    let env: SseEnvelope;
    try {
      env = JSON.parse(ev.data) as SseEnvelope;
    } catch {
      return;
    }
    if (typeof env.seq === "number") {
      if (env.seq <= lastSeq) return; // replayed duplicate
      lastSeq = env.seq;
    }
    handlers.onEvent(env);
    if (TERMINAL.has(env.event)) close();
  };

  for (const name of SSE_EVENTS) es.addEventListener(name, listener as EventListener);
  es.onmessage = listener; // unnamed events, just in case

  es.onerror = () => {
    if (closed) return;
    if (es.readyState === EventSource.CLOSED) {
      // the browser gave up (e.g. 404 / non-event-stream response)
      close();
      handlers.onFailure?.("The event stream is unavailable.");
      return;
    }
    silentReconnects += 1;
    if (silentReconnects > MAX_SILENT_RECONNECTS) {
      close();
      handlers.onFailure?.("Lost connection to the event stream.");
    }
  };

  return { close };
}

/** Async-iterator view of `openEventStream`, for the chat adapter's async generator. */
export async function* iterateEventStream(url: string, signal?: AbortSignal): AsyncGenerator<SseEnvelope> {
  const queue: SseEnvelope[] = [];
  let failure: string | null = null;
  let finished = false;
  let wake: (() => void) | null = null;
  const notify = () => {
    wake?.();
    wake = null;
  };

  const handle = openEventStream(url, {
    onEvent: (env) => {
      queue.push(env);
      if (TERMINAL.has(env.event)) finished = true;
      notify();
    },
    onFailure: (msg) => {
      failure = msg;
      finished = true;
      notify();
    },
  });
  const onAbort = () => {
    handle.close();
    finished = true;
    notify();
  };
  signal?.addEventListener("abort", onAbort);

  try {
    while (true) {
      while (queue.length) yield queue.shift()!;
      if (finished) break;
      await new Promise<void>((r) => (wake = r));
    }
    if (failure && !signal?.aborted) throw new Error(failure);
  } finally {
    signal?.removeEventListener("abort", onAbort);
    handle.close();
  }
}
