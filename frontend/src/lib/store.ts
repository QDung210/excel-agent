import { create } from "zustand";

/**
 * Shared workbook state between the grid and the chat. The grid subscribes to `revision` and
 * `reloadNonce` only; chat tokens never touch this store, so streaming never re-renders the grid.
 */
interface WorkbookState {
  workbookId: string | null;
  revision: number;
  canUndo: boolean;
  /** Bumped to force a canonical reload of the same revision (e.g. revert a refused edit). */
  reloadNonce: number;
  /** Bumped whenever metadata (name, can_undo, history) should be refetched. */
  metaNonce: number;
  activeSheet: string | null;
  /** A1 range of the current grid selection, without sheet prefix, e.g. "B2:C4". */
  selection: string | null;
  saving: boolean;

  reset: (workbookId: string, revision: number, canUndo: boolean) => void;
  setRevision: (revision: number) => void;
  setCanUndo: (v: boolean) => void;
  reloadCanonical: () => void;
  setSelection: (sheet: string | null, range: string | null) => void;
  setSaving: (v: boolean) => void;
}

export const useWorkbookStore = create<WorkbookState>((set) => ({
  workbookId: null,
  revision: 0,
  canUndo: false,
  reloadNonce: 0,
  metaNonce: 0,
  activeSheet: null,
  selection: null,
  saving: false,

  reset: (workbookId, revision, canUndo) =>
    set({ workbookId, revision, canUndo, activeSheet: null, selection: null, saving: false }),
  setRevision: (revision) =>
    set((s) => (revision === s.revision ? s : { revision, metaNonce: s.metaNonce + 1 })),
  setCanUndo: (canUndo) => set({ canUndo }),
  reloadCanonical: () => set((s) => ({ reloadNonce: s.reloadNonce + 1 })),
  setSelection: (activeSheet, selection) => set({ activeSheet, selection }),
  setSaving: (saving) => set({ saving }),
}));

/** Per-plan confirmation status for agent confirmation cards. */
type PlanStatus = "pending" | "applying" | "applied" | "rejecting" | "rejected" | "error";

interface PlanState {
  plans: Record<string, { status: PlanStatus; message?: string }>;
  setPlan: (planId: string, status: PlanStatus, message?: string) => void;
}

export const usePlanStore = create<PlanState>((set) => ({
  plans: {},
  setPlan: (planId, status, message) => set((s) => ({ plans: { ...s.plans, [planId]: { status, message } } })),
}));
