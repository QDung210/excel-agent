import { memo } from "react";
import { MessagesSquare, PanelRightClose } from "lucide-react";
import { Button } from "@/components/ui/button";
import { Thread } from "./Thread";

/** Right-hand assistant panel. The runtime provider lives in the Workspace page so the drawer and the panel share one thread. */
export const ChatPanel = memo(function ChatPanel({ onCollapse }: { onCollapse?: () => void }) {
  return (
    <aside className="flex h-full min-h-0 flex-col bg-background" aria-label="Assistant">
      <div className="flex h-10 shrink-0 items-center gap-2 border-b border-stroke-subtle px-3">
        <MessagesSquare className="size-4 text-brand" aria-hidden />
        <h2 className="flex-1 text-sm font-semibold">Assistant</h2>
        {onCollapse && (
          <Button variant="ghost" size="icon-xs" onClick={onCollapse} aria-label="Collapse assistant panel" title="Collapse assistant panel">
            <PanelRightClose aria-hidden />
          </Button>
        )}
      </div>
      <div className="min-h-0 flex-1">
        <Thread />
      </div>
    </aside>
  );
});
