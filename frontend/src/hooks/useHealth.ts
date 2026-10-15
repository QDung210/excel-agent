import { useEffect, useState } from "react";
import { api, errorMessage } from "@/lib/api";
import type { Health } from "@/lib/types";

let cached: Promise<Health> | null = null;

/** Fetch `/api/health` once per page load and share the result. */
export function useHealth() {
  const [health, setHealth] = useState<Health | null>(null);
  const [error, setError] = useState<string | null>(null);

  useEffect(() => {
    let alive = true;
    cached ??= api.health();
    cached.then(
      (h) => alive && setHealth(h),
      (e) => {
        cached = null; // allow a retry on next mount
        if (alive) setError(errorMessage(e));
      },
    );
    return () => {
      alive = false;
    };
  }, []);

  return { health, error };
}
