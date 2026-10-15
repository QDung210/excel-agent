import { Sparkles, FlaskConical, CircleAlert } from "lucide-react";
import { Badge } from "@/components/ui/badge";
import { Skeleton } from "@/components/ui/skeleton";
import { Tooltip, TooltipContent, TooltipTrigger } from "@/components/ui/tooltip";
import { useHealth } from "@/hooks/useHealth";

/** Honest LLM mode badge, straight from /api/health. */
export function LlmBadge() {
  const { health, error } = useHealth();
  if (error) {
    return (
      <Badge variant="destructive" className="gap-1" title={error}>
        <CircleAlert aria-hidden /> API offline
      </Badge>
    );
  }
  if (!health) return <Skeleton className="h-5 w-24 rounded-full" aria-label="Loading model status" />;

  if (health.llm_mode === "mock") {
    return (
      <Tooltip>
        <TooltipTrigger asChild>
          <Badge variant="outline" className="gap-1 border-warning/40 bg-warning-soft text-warning">
            <FlaskConical aria-hidden /> Mock mode
          </Badge>
        </TooltipTrigger>
        <TooltipContent>No LLM key configured. Results come from a deterministic mock.</TooltipContent>
      </Tooltip>
    );
  }
  return (
    <Badge variant="outline" className="gap-1 border-brand/30 bg-brand-soft text-brand">
      <Sparkles aria-hidden /> Azure OpenAI · {(health.model ?? "unknown model").replace(/^azure:/, "")}
    </Badge>
  );
}
