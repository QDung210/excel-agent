import { lazy, Suspense } from "react";
import { BrowserRouter, Link, Route, Routes } from "react-router-dom";
import { Toaster } from "@/components/ui/sonner";
import { TooltipProvider } from "@/components/ui/tooltip";
import { Button } from "@/components/ui/button";
import { Skeleton } from "@/components/ui/skeleton";
import Landing from "@/pages/Landing";

// The grid and chat runtime are heavy; keep them off the landing page bundle.
const Studio = lazy(() => import("@/pages/Studio"));
const Workspace = lazy(() => import("@/pages/Workspace"));

function PageFallback() {
  return (
    <div className="flex h-dvh flex-col gap-3 p-6" aria-busy="true" aria-label="Loading">
      <Skeleton className="h-8 w-48" />
      <Skeleton className="flex-1" />
    </div>
  );
}

function NotFound() {
  return (
    <main className="grid h-dvh place-items-center p-6 text-center">
      <div>
        <h1 className="text-xl font-semibold">Page not found</h1>
        <p className="mt-1 text-sm text-muted-foreground">That address does not match any screen.</p>
        <Button className="mt-4" asChild>
          <Link to="/">Back to start</Link>
        </Button>
      </div>
    </main>
  );
}

export default function App() {
  return (
    <BrowserRouter>
      <TooltipProvider delayDuration={300}>
        <Suspense fallback={<PageFallback />}>
          <Routes>
            <Route path="/" element={<Landing />} />
            <Route path="/new/:mode" element={<Studio />} />
            <Route path="/w/:workbookId" element={<Workspace />} />
            <Route path="*" element={<NotFound />} />
          </Routes>
        </Suspense>
        <Toaster position="bottom-right" richColors closeButton />
      </TooltipProvider>
    </BrowserRouter>
  );
}
