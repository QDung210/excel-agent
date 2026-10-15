import { Link } from "react-router-dom";
import { Table2 } from "lucide-react";

export function AppMark() {
  return (
    <Link to="/" className="flex items-center gap-2 rounded-md font-semibold tracking-tight" aria-label="AI Excel Workspace home">
      <span className="grid size-7 place-items-center rounded-md bg-primary text-primary-foreground">
        <Table2 className="size-4" aria-hidden />
      </span>
      <span className="hidden sm:inline">Excel Workspace</span>
    </Link>
  );
}
