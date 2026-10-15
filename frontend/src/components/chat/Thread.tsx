import { memo } from "react";
import {
  AuiIf,
  ComposerPrimitive,
  ErrorPrimitive,
  MessagePrimitive,
  ThreadPrimitive,
} from "@assistant-ui/react";
import { MarkdownTextPrimitive } from "@assistant-ui/react-markdown";
import remarkGfm from "remark-gfm";
import { ArrowDown, ArrowUp, MousePointerSquareDashed, Sparkles, Square } from "lucide-react";
import { Button } from "@/components/ui/button";
import { qualify } from "@/lib/a1";
import { useWorkbookStore } from "@/lib/store";
import { CONFIRM_TOOL } from "./adapter";
import { ConfirmCard } from "./ConfirmCard";
import { ToolCard } from "./ToolCard";

const MarkdownText = memo(function MarkdownText() {
  return (
    <MarkdownTextPrimitive
      remarkPlugins={[remarkGfm]}
      className="chat-markdown text-sm leading-relaxed"
    />
  );
});

const PART_COMPONENTS = {
  Text: MarkdownText,
  tools: { by_name: { [CONFIRM_TOOL]: ConfirmCard }, Fallback: ToolCard },
};

function UserMessage() {
  return (
    <MessagePrimitive.Root className="flex justify-end py-1.5">
      <div className="max-w-[85%] whitespace-pre-wrap rounded-2xl rounded-br-md bg-brand px-3 py-2 text-sm text-brand-foreground">
        <MessagePrimitive.Parts />
      </div>
    </MessagePrimitive.Root>
  );
}

function AssistantMessage() {
  return (
    <MessagePrimitive.Root className="py-1.5">
      <div className="min-w-0 text-foreground">
        <MessagePrimitive.Parts components={PART_COMPONENTS} />
        <MessagePrimitive.Error>
          <ErrorPrimitive.Root className="mt-1 rounded-md border border-destructive/30 bg-destructive/10 px-2 py-1 text-xs text-destructive" role="alert">
            <ErrorPrimitive.Message />
          </ErrorPrimitive.Root>
        </MessagePrimitive.Error>
      </div>
    </MessagePrimitive.Root>
  );
}

const SUGGESTIONS = [
  "Summarize what this workbook contains",
  "Fill the empty cells in the selected range",
  "Check the totals for mistakes",
];

function Welcome() {
  return (
    <div className="flex flex-1 flex-col items-center justify-center gap-3 px-4 py-10 text-center">
      <span className="grid size-10 place-items-center rounded-full bg-brand-soft text-brand">
        <Sparkles className="size-5" aria-hidden />
      </span>
      <div>
        <p className="font-medium">Ask about this workbook</p>
        <p className="text-sm text-muted-foreground">Select cells in the grid to give the assistant context.</p>
      </div>
      <div className="flex w-full flex-col gap-1.5">
        {SUGGESTIONS.map((s) => (
          <ThreadPrimitive.Suggestion
            key={s}
            prompt={s}
            send
            className="rounded-lg border border-border px-3 py-2 text-left text-sm transition-colors hover:bg-muted"
          >
            {s}
          </ThreadPrimitive.Suggestion>
        ))}
      </div>
    </div>
  );
}

function SelectionChip() {
  const sheet = useWorkbookStore((s) => s.activeSheet);
  const range = useWorkbookStore((s) => s.selection);
  if (!range) return null;
  const label = sheet ? qualify(sheet, range) : range;
  return (
    <span
      className="inline-flex max-w-full items-center gap-1 rounded-md bg-brand-soft px-1.5 py-0.5 font-mono text-[11px] text-brand"
      title={`Selection sent with your message: ${label}`}
    >
      <MousePointerSquareDashed className="size-3 shrink-0" aria-hidden />
      <span className="truncate">{label}</span>
    </span>
  );
}

function Composer() {
  return (
    <ComposerPrimitive.Root className="rounded-xl border border-border bg-card shadow-xs focus-within:border-brand/60 focus-within:ring-2 focus-within:ring-brand/20">
      <div className="flex items-center gap-2 px-3 pt-2">
        <SelectionChip />
      </div>
      <ComposerPrimitive.Input
        submitMode="ctrlEnter"
        rows={2}
        autoFocus={false}
        placeholder="Ask the assistant to read or change cells…"
        aria-label="Message the assistant (Ctrl+Enter to send)"
        className="block max-h-40 min-h-12 w-full resize-none bg-transparent px-3 py-2 text-sm outline-none placeholder:text-muted-foreground"
      />
      <div className="flex items-center justify-between px-2 pb-2">
        <span className="pl-1 text-[11px] text-muted-foreground">
          <kbd className="font-sans">Ctrl</kbd>+<kbd className="font-sans">Enter</kbd> to send
        </span>
        <AuiIf condition={(s) => !s.thread.isRunning}>
          <ComposerPrimitive.Send asChild>
            <Button size="icon-sm" aria-label="Send message">
              <ArrowUp aria-hidden />
            </Button>
          </ComposerPrimitive.Send>
        </AuiIf>
        <AuiIf condition={(s) => s.thread.isRunning}>
          <ComposerPrimitive.Cancel asChild>
            <Button size="icon-sm" variant="outline" aria-label="Stop response">
              <Square className="fill-current" aria-hidden />
            </Button>
          </ComposerPrimitive.Cancel>
        </AuiIf>
      </div>
    </ComposerPrimitive.Root>
  );
}

export function Thread() {
  return (
    <ThreadPrimitive.Root className="flex h-full min-h-0 flex-col">
      <ThreadPrimitive.Viewport className="relative flex min-h-0 flex-1 flex-col overflow-y-auto px-3 pt-3" aria-live="polite">
        <AuiIf condition={(s) => s.thread.isEmpty}>
          <Welcome />
        </AuiIf>
        <ThreadPrimitive.Messages>
          {({ message }) => (message.role === "user" ? <UserMessage /> : <AssistantMessage />)}
        </ThreadPrimitive.Messages>
        <AuiIf condition={(s) => s.thread.isRunning}>
          <div className="flex items-center gap-1 py-2" aria-label="Assistant is working">
            {[0, 1, 2].map((i) => (
              <span
                key={i}
                className="size-1.5 animate-pulse rounded-full bg-muted-foreground/60"
                style={{ animationDelay: `${i * 150}ms` }}
              />
            ))}
          </div>
        </AuiIf>
        <ThreadPrimitive.ViewportFooter className="sticky bottom-0 mt-auto bg-background pb-3 pt-2">
          <ThreadPrimitive.ScrollToBottom asChild>
            <Button
              variant="outline"
              size="icon-sm"
              className="absolute -top-10 left-1/2 -translate-x-1/2 rounded-full disabled:invisible"
              aria-label="Scroll to latest message"
            >
              <ArrowDown aria-hidden />
            </Button>
          </ThreadPrimitive.ScrollToBottom>
          <Composer />
        </ThreadPrimitive.ViewportFooter>
      </ThreadPrimitive.Viewport>
    </ThreadPrimitive.Root>
  );
}
