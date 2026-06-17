import { createFileRoute } from "@tanstack/react-router";
import { useCallback, useEffect, useMemo, useRef, useState } from "react";
import Papa from "papaparse";
// react-window removed — queue is no longer displayed as a scrolling list.
import { toast } from "sonner";
import {
  Upload, Sun, Moon, Trash2, Send, Mail, Code2, Copy, Check,
  Zap, FileText, Hash, Eye, SkipForward,
} from "lucide-react";
import { Button } from "@/components/ui/button";
import { Input } from "@/components/ui/input";
import { Textarea } from "@/components/ui/textarea";
import { Label } from "@/components/ui/label";
import { Tabs, TabsList, TabsTrigger, TabsContent } from "@/components/ui/tabs";
import {
  Dialog, DialogContent, DialogDescription, DialogFooter,
  DialogHeader, DialogTitle, DialogTrigger,
} from "@/components/ui/dialog";
import { cn } from "@/lib/utils";

export const Route = createFileRoute("/")({
  head: () => ({
    meta: [
      { title: "Midey Enterprises Outreach Lab" },
      { name: "description", content: "High-performance mobile outreach console: parse 50MB+ CSV lead lists, fire personalized mailto handoffs, and craft rich HTML drafts — all 100% client-side." },
      { property: "og:title", content: "Midey Enterprises Outreach Lab" },
      { property: "og:description", content: "Client-side outreach dashboard for rapid, templated email handoffs from massive CSV lead lists." },
    ],
  }),
  component: Index,
});

/* ----------------------------- Types ----------------------------- */

type Row = Record<string, string>;
type RowState = "pending" | "processed" | "skipped";

interface PersistedState {
  headers: string[];
  rows: Row[];
  rowStates: Record<number, RowState>;
  targetEmailHeader: string;
  subjectA: string;
  bodyA: string;
  recipientB: string;
  sampleIdB: number;
  subjectB: string;
  htmlB: string;
}

const STORAGE_KEY = "midey.outreach.v1";
const THEME_KEY = "midey.theme";

const DEFAULT_STATE: PersistedState = {
  headers: [],
  rows: [],
  rowStates: {},
  targetEmailHeader: "",
  subjectA: "Quick question, {first_name}",
  bodyA: "Hi {first_name},\n\nNoticed {company} — wanted to reach out.\n\n— Midey",
  recipientB: "",
  sampleIdB: 0,
  subjectB: "A note for {first_name}",
  htmlB: "<div style=\"font-family:system-ui;line-height:1.55\">\n  <h2 style=\"color:#0ea5e9\">Hi {first_name} 👋</h2>\n  <p>Loved what you're doing at <b>{company}</b>.</p>\n  <p>— Midey Enterprises</p>\n</div>",
};

/* --------------------------- Utilities --------------------------- */

const TOKEN_RE = /\{([^{}]+)\}/g;

function renderTemplate(tpl: string, row: Row | undefined): string {
  if (!row) return tpl;
  return tpl.replace(TOKEN_RE, (_, key: string) => {
    const k = key.trim();
    return row[k] ?? "";
  });
}

function loadState(): PersistedState {
  if (typeof window === "undefined") return DEFAULT_STATE;
  try {
    const raw = localStorage.getItem(STORAGE_KEY);
    if (!raw) return DEFAULT_STATE;
    return { ...DEFAULT_STATE, ...(JSON.parse(raw) as PersistedState) };
  } catch {
    return DEFAULT_STATE;
  }
}

/* --------------------------- Component --------------------------- */

function Index() {
  const [hydrated, setHydrated] = useState(false);
  const [state, setState] = useState<PersistedState>(DEFAULT_STATE);
  const [theme, setTheme] = useState<"dark" | "light">("dark");
  const [parsing, setParsing] = useState(false);
  const [parseProgress, setParseProgress] = useState(0);
  const [copiedToken, setCopiedToken] = useState<string | null>(null);

  // Hydrate from localStorage
  useEffect(() => {
    setState(loadState());
    const t = (localStorage.getItem(THEME_KEY) as "dark" | "light" | null) ?? "dark";
    setTheme(t);
    setHydrated(true);
  }, []);

  // Persist state (debounced microtask)
  useEffect(() => {
    if (!hydrated) return;
    const id = setTimeout(() => {
      try { localStorage.setItem(STORAGE_KEY, JSON.stringify(state)); } catch {}
    }, 150);
    return () => clearTimeout(id);
  }, [state, hydrated]);

  // Apply theme class
  useEffect(() => {
    if (!hydrated) return;
    const root = document.documentElement;
    root.classList.toggle("dark", theme === "dark");
    localStorage.setItem(THEME_KEY, theme);
  }, [theme, hydrated]);

  const patch = useCallback((p: Partial<PersistedState>) => {
    setState((s) => ({ ...s, ...p }));
  }, []);

  /* ---------- CSV Parse (streaming, off main work via Papa worker) ---------- */

  const onFile = useCallback((file: File) => {
    if (!file) return;
    setParsing(true);
    setParseProgress(0);
    toast.info(`Parsing ${(file.size / 1024 / 1024).toFixed(1)} MB …`);

    const collected: Row[] = [];
    let headers: string[] = [];
    const total = file.size || 1;

    Papa.parse<Row>(file, {
      header: true,
      skipEmptyLines: true,
      worker: true,
      chunkSize: 1024 * 512, // 512KB chunks
      chunk: (results, parser) => {
        if (!headers.length && results.meta.fields) headers = results.meta.fields;
        for (const r of results.data) collected.push(r);
        const cursor = (results.meta as { cursor?: number }).cursor ?? 0;
        setParseProgress(Math.min(100, Math.round((cursor / total) * 100)));
        if (collected.length > 500_000) {
          parser.abort();
          toast.error("Row cap (500k) reached — truncated.");
        }
      },
      complete: () => {
        setParsing(false);
        setParseProgress(100);
        const guessEmail =
          headers.find((h) => /e?mail/i.test(h)) ?? headers[0] ?? "";
        setState((s) => ({
          ...s,
          headers,
          rows: collected,
          rowStates: {},
          targetEmailHeader: s.targetEmailHeader || guessEmail,
        }));
        toast.success(`Loaded ${collected.length.toLocaleString()} rows · ${headers.length} columns`);
      },
      error: (err) => {
        setParsing(false);
        toast.error(`Parse failed: ${err.message}`);
      },
    });
  }, []);

  /* ---------- Derived: active queue (pending rows, processed sink to bottom) ---------- */

  const queue = useMemo(() => {
    const pending: number[] = [];
    const processed: number[] = [];
    for (let i = 0; i < state.rows.length; i++) {
      if (state.rowStates[i] === "processed" || state.rowStates[i] === "skipped") {
        processed.push(i);
      } else {
        pending.push(i);
      }
    }
    return [...pending, ...processed];
  }, [state.rows, state.rowStates]);

  const processedCount = useMemo(
    () => Object.values(state.rowStates).filter((v) => v === "processed").length,
    [state.rowStates],
  );

  /* ---------- Row actions ---------- */

  const fireRow = useCallback((rowIndex: number) => {
    const row = state.rows[rowIndex];
    if (!row) return;
    const toAddr = (row[state.targetEmailHeader] || "").trim();
    if (!toAddr) {
      toast.error(`Row ${rowIndex} missing "${state.targetEmailHeader}"`);
      return;
    }
    const subject = renderTemplate(state.subjectA, row);
    const body = renderTemplate(state.bodyA, row);
    const href = `mailto:${encodeURIComponent(toAddr)}?subject=${encodeURIComponent(subject)}&body=${encodeURIComponent(body)}`;
    window.location.href = href;
    setState((s) => ({ ...s, rowStates: { ...s.rowStates, [rowIndex]: "processed" } }));
  }, [state.rows, state.targetEmailHeader, state.subjectA, state.bodyA]);

  const skipRow = useCallback((rowIndex: number) => {
    setState((s) => ({ ...s, rowStates: { ...s.rowStates, [rowIndex]: "skipped" } }));
  }, []);

  const resetRow = useCallback((rowIndex: number) => {
    setState((s) => {
      const next = { ...s.rowStates };
      delete next[rowIndex];
      return { ...s, rowStates: next };
    });
  }, []);

  /* ---------- Token copy ---------- */

  const copyToken = useCallback((tok: string) => {
    const text = `{${tok}}`;
    navigator.clipboard?.writeText(text).then(
      () => {
        setCopiedToken(tok);
        toast.success(`Copied ${text}`);
        setTimeout(() => setCopiedToken((c) => (c === tok ? null : c)), 1200);
      },
      () => toast.error("Clipboard blocked"),
    );
  }, []);

  /* ---------- Clear all ---------- */

  const clearAll = () => {
    localStorage.removeItem(STORAGE_KEY);
    setState(DEFAULT_STATE);
    toast.success("All data cleared.");
  };

  /* ---------- Section B: HTML preview + dual action ---------- */

  const sampleRow = state.rows[state.sampleIdB];
  const renderedHtml = useMemo(() => renderTemplate(state.htmlB, sampleRow), [state.htmlB, sampleRow]);
  const renderedSubjectB = useMemo(() => renderTemplate(state.subjectB, sampleRow), [state.subjectB, sampleRow]);

  const executeHtml = useCallback(async () => {
    if (!state.recipientB.trim()) { toast.error("Recipient required"); return; }
    try {
      const blobHtml = new Blob([renderedHtml], { type: "text/html" });
      const blobText = new Blob([renderedHtml.replace(/<[^>]+>/g, "")], { type: "text/plain" });
      if ("ClipboardItem" in window && navigator.clipboard?.write) {
        await navigator.clipboard.write([
          new ClipboardItem({ "text/html": blobHtml, "text/plain": blobText }),
        ]);
      } else {
        await navigator.clipboard.writeText(renderedHtml);
      }
      toast.success("Rich HTML copied — opening mail in 300ms…");
    } catch (e) {
      toast.error(`Clipboard failed: ${(e as Error).message}`);
      return;
    }
    setTimeout(() => {
      const href = `mailto:${encodeURIComponent(state.recipientB.trim())}?subject=${encodeURIComponent(renderedSubjectB)}`;
      window.location.href = href;
    }, 300);
  }, [state.recipientB, renderedHtml, renderedSubjectB]);

  /* ---------- UI ---------- */

  return (
    <div className="min-h-screen bg-bg-app text-foreground">
      <Header
        theme={theme}
        onToggleTheme={() => setTheme((t) => (t === "dark" ? "light" : "dark"))}
        onClearAll={clearAll}
        totalRows={state.rows.length}
        processedRows={processedCount}
      />

      <main className="mx-auto max-w-5xl px-3 pb-24 pt-4 sm:px-6">
        <IngestPanel
          parsing={parsing}
          progress={parseProgress}
          onFile={onFile}
          headers={state.headers}
          totalRows={state.rows.length}
          processedRows={processedCount}
          targetEmailHeader={state.targetEmailHeader}
          onTargetEmailHeader={(v) => patch({ targetEmailHeader: v })}
          copyToken={copyToken}
          copiedToken={copiedToken}
        />

        <Tabs defaultValue="a" className="mt-6">
          <TabsList className="grid w-full grid-cols-2">
            <TabsTrigger value="a" className="gap-2"><Zap className="size-4" />Section A · Fast Plain</TabsTrigger>
            <TabsTrigger value="b" className="gap-2"><Code2 className="size-4" />Section B · HTML Lab</TabsTrigger>
          </TabsList>

          <TabsContent value="a" className="mt-4 space-y-4">
            <SectionACard
              state={state}
              patch={patch}
              queue={queue}
              processedCount={processedCount}
              fireRow={fireRow}
              skipRow={skipRow}
              resetRow={resetRow}
            />
          </TabsContent>

          <TabsContent value="b" className="mt-4 space-y-4">
            <SectionBCard
              state={state}
              patch={patch}
              renderedHtml={renderedHtml}
              renderedSubject={renderedSubjectB}
              onExecute={executeHtml}
              sampleRow={sampleRow}
            />
          </TabsContent>
        </Tabs>
      </main>
    </div>
  );
}

/* ----------------------------- Header ----------------------------- */

function Header({
  theme, onToggleTheme, onClearAll, totalRows, processedRows,
}: {
  theme: "dark" | "light";
  onToggleTheme: () => void;
  onClearAll: () => void;
  totalRows: number;
  processedRows: number;
}) {
  const [open, setOpen] = useState(false);
  return (
    <header className="sticky top-0 z-30 border-b border-border-strong/60 bg-bg-app/80 backdrop-blur supports-[backdrop-filter]:bg-bg-app/60">
      <div className="mx-auto flex max-w-5xl items-center gap-2 px-3 py-3 sm:px-6">
        <div className="flex min-w-0 flex-1 items-center gap-2">
          <div className="grid size-9 place-items-center rounded-md bg-surface-2 glow-sky">
            <Mail className="size-4 text-sky-glow" />
          </div>
          <div className="min-w-0">
            <h1 className="truncate text-sm font-semibold leading-tight sm:text-base">
              Midey Enterprises <span className="text-sky-glow">Outreach Lab</span>
            </h1>
            <p className="truncate font-mono-data text-[10px] text-muted-foreground sm:text-xs">
              {processedRows.toLocaleString()} / {totalRows.toLocaleString()} processed
            </p>
          </div>
        </div>

        <Button variant="ghost" size="icon" onClick={onToggleTheme} aria-label="Toggle theme">
          {theme === "dark" ? <Sun className="size-4" /> : <Moon className="size-4" />}
        </Button>

        <Dialog open={open} onOpenChange={setOpen}>
          <DialogTrigger asChild>
            <Button variant="ghost" size="icon" aria-label="Clear all data">
              <Trash2 className="size-4 text-destructive" />
            </Button>
          </DialogTrigger>
          <DialogContent>
            <DialogHeader>
              <DialogTitle>Wipe local data?</DialogTitle>
              <DialogDescription>
                This deletes the parsed CSV, templates, and processed-row history from this browser. There is no undo.
              </DialogDescription>
            </DialogHeader>
            <DialogFooter>
              <Button variant="ghost" onClick={() => setOpen(false)}>Cancel</Button>
              <Button variant="destructive" onClick={() => { onClearAll(); setOpen(false); }}>
                <Trash2 /> Clear everything
              </Button>
            </DialogFooter>
          </DialogContent>
        </Dialog>
      </div>
    </header>
  );
}

/* --------------------------- Ingest panel --------------------------- */

function IngestPanel({
  parsing, progress, onFile, headers, totalRows, processedRows,
  targetEmailHeader, onTargetEmailHeader, copyToken, copiedToken,
}: {
  parsing: boolean;
  progress: number;
  onFile: (f: File) => void;
  headers: string[];
  totalRows: number;
  processedRows: number;
  targetEmailHeader: string;
  onTargetEmailHeader: (v: string) => void;
  copyToken: (t: string) => void;
  copiedToken: string | null;
}) {
  const inputRef = useRef<HTMLInputElement>(null);
  return (
    <section className="rounded-xl border border-border-strong/70 bg-surface-1 p-4 shadow-sm">
      <div className="flex flex-wrap items-center gap-3">
        <input
          ref={inputRef}
          type="file"
          accept=".csv,text/csv"
          className="hidden"
          onChange={(e) => {
            const f = e.target.files?.[0];
            if (f) onFile(f);
            e.target.value = "";
          }}
        />
        <Button
          onClick={() => inputRef.current?.click()}
          className="glow-sky"
          disabled={parsing}
        >
          <Upload /> {parsing ? `Parsing… ${progress}%` : totalRows ? "Replace CSV" : "Upload CSV"}
        </Button>

        <div className="flex items-center gap-2 font-mono-data text-xs text-muted-foreground">
          <FileText className="size-3.5" />
          {totalRows ? (
            <>
              <span className="text-foreground">{totalRows.toLocaleString()}</span> rows ·{" "}
              <span className="text-foreground">{headers.length}</span> cols ·{" "}
              <span className="text-sky-glow">{processedRows.toLocaleString()}</span> done
            </>
          ) : (
            <span>No file loaded — streaming parser ready for 50MB+ files.</span>
          )}
        </div>
      </div>

      {parsing && (
        <div className="mt-3 h-1 w-full overflow-hidden rounded-full bg-surface-2">
          <div
            className="h-full bg-[var(--sky)] transition-[width] duration-150"
            style={{ width: `${progress}%` }}
          />
        </div>
      )}

      {headers.length > 0 && (
        <>
          <div className="mt-4 grid gap-2 sm:grid-cols-[200px_1fr] sm:items-center">
            <Label className="text-xs uppercase tracking-wider text-muted-foreground">
              Target email column
            </Label>
            <select
              value={targetEmailHeader}
              onChange={(e) => onTargetEmailHeader(e.target.value)}
              className="h-9 w-full rounded-md border border-border-strong/70 bg-surface-2 px-2 font-mono-data text-sm outline-none focus:glow-sky"
            >
              {headers.map((h) => (
                <option key={h} value={h}>{h}</option>
              ))}
            </select>
          </div>

          <div className="mt-4">
            <Label className="text-xs uppercase tracking-wider text-muted-foreground">
              Tokens · tap to copy
            </Label>
            <div className="mt-2 flex flex-wrap gap-1.5">
              {headers.map((h) => {
                const active = copiedToken === h;
                return (
                  <button
                    key={h}
                    onClick={() => copyToken(h)}
                    className={cn(
                      "inline-flex items-center gap-1 rounded-md border border-border-strong/70 bg-surface-2 px-2 py-1 font-mono-data text-xs transition",
                      "hover:border-[var(--sky)]/60 hover:text-sky-glow active:scale-[0.97]",
                      active && "border-[var(--sky)]/80 text-sky-glow glow-sky",
                    )}
                  >
                    {active ? <Check className="size-3" /> : <Hash className="size-3 opacity-60" />}
                    {`{${h}}`}
                  </button>
                );
              })}
            </div>
          </div>
        </>
      )}
    </section>
  );
}

/* --------------------------- Section A --------------------------- */

function SectionACard({
  state, patch, queue, processedCount, fireRow, skipRow,
}: {
  state: PersistedState;
  patch: (p: Partial<PersistedState>) => void;
  queue: number[];
  processedCount: number;
  fireRow: (i: number) => void;
  skipRow: (i: number) => void;
  resetRow: (i: number) => void;
}) {
  const nextPendingIndex = queue.find(
    (i) => (state.rowStates[i] ?? "pending") === "pending",
  );
  const pendingCount = state.rows.length - processedCount;

  return (
    <div className="space-y-4 rounded-xl border border-border-strong/70 bg-surface-1 p-4">
      <div className="grid gap-3 sm:grid-cols-2">
        <Field label="Subject template">
          <Input
            value={state.subjectA}
            onChange={(e) => patch({ subjectA: e.target.value })}
            className="font-mono-data"
            placeholder="Hi {first_name}…"
          />
        </Field>
        <Field label="Email header column">
          <Input
            value={state.targetEmailHeader}
            onChange={(e) => patch({ targetEmailHeader: e.target.value })}
            className="font-mono-data"
            placeholder="email"
          />
        </Field>
      </div>
      <Field label="Plain-text body template">
        <Textarea
          value={state.bodyA}
          onChange={(e) => patch({ bodyA: e.target.value })}
          rows={6}
          className="font-mono-data text-[13px]"
          placeholder="Hi {first_name}, …"
        />
      </Field>

      <div className="flex items-center justify-between">
        <div className="font-mono-data text-xs text-muted-foreground">
          Queue · <span className="text-foreground">{pendingCount.toLocaleString()}</span> pending ·{" "}
          <span className="text-sky-glow">{processedCount.toLocaleString()}</span> done
        </div>
        <div className="font-mono-data text-[10px] uppercase tracking-wider text-muted-foreground">
          headless
        </div>
      </div>

      <div className="overflow-hidden rounded-lg border border-border-strong/60 bg-bg-app p-4">
        {state.rows.length === 0 ? (
          <div className="grid place-items-center px-6 py-10 text-center">
            <div>
              <Upload className="mx-auto size-6 text-muted-foreground" />
              <p className="mt-2 text-sm text-muted-foreground">
                Upload a CSV to populate the queue.
              </p>
            </div>
          </div>
        ) : nextPendingIndex === undefined ? (
          <p className="py-6 text-center font-mono-data text-xs text-muted-foreground">
            All rows processed.
          </p>
        ) : (
          <div className="flex flex-wrap items-center justify-between gap-3">
            <div className="font-mono-data text-xs text-muted-foreground">
              Next up · row <span className="text-foreground">#{nextPendingIndex}</span>
            </div>
            <div className="flex gap-2">
              <Button size="sm" variant="ghost" onClick={() => skipRow(nextPendingIndex)}>
                <SkipForward className="size-3.5" /> Skip
              </Button>
              <Button
                size="sm"
                className="glow-amber bg-[var(--amber)] text-black hover:bg-[var(--amber)]/90"
                onClick={() => fireRow(nextPendingIndex)}
              >
                <Send className="size-3.5" /> Send next
              </Button>
            </div>
          </div>
        )}
      </div>
    </div>
  );
}

/* --------------------------- Section B --------------------------- */

function SectionBCard({
  state, patch, renderedHtml, renderedSubject, onExecute, sampleRow,
}: {
  state: PersistedState;
  patch: (p: Partial<PersistedState>) => void;
  renderedHtml: string;
  renderedSubject: string;
  onExecute: () => void;
  sampleRow: Row | undefined;
}) {
  return (
    <div className="space-y-4 rounded-xl border border-border-strong/70 bg-surface-1 p-4">
      <div className="grid gap-3 sm:grid-cols-2">
        <Field label="Test recipient">
          <Input
            type="email"
            value={state.recipientB}
            onChange={(e) => patch({ recipientB: e.target.value })}
            placeholder="you@example.com"
            className="font-mono-data"
          />
        </Field>
        <Field label="Sample row ID">
          <Input
            type="number"
            min={0}
            max={Math.max(0, state.rows.length - 1)}
            value={state.sampleIdB}
            onChange={(e) => patch({ sampleIdB: Math.max(0, Number(e.target.value) || 0) })}
            className="font-mono-data"
          />
        </Field>
      </div>

      <Field label="Subject template">
        <Input
          value={state.subjectB}
          onChange={(e) => patch({ subjectB: e.target.value })}
          className="font-mono-data"
        />
      </Field>

      <Field label="HTML body source">
        <Textarea
          value={state.htmlB}
          onChange={(e) => patch({ htmlB: e.target.value })}
          rows={10}
          className="font-mono-data text-[12px] leading-relaxed"
          spellCheck={false}
        />
      </Field>

      <div>
        <div className="mb-2 flex items-center gap-2">
          <Eye className="size-3.5 text-muted-foreground" />
          <span className="font-mono-data text-[10px] uppercase tracking-wider text-muted-foreground">
            Live preview · subject: <span className="text-amber-glow">{renderedSubject || "—"}</span>
          </span>
        </div>
        <div className="overflow-hidden rounded-lg border border-border-strong/60 bg-white">
          <iframe
            title="HTML preview"
            sandbox=""
            srcDoc={`<!doctype html><html><body style="margin:0;padding:12px;font-family:system-ui">${renderedHtml}</body></html>`}
            className="block h-[360px] w-full"
          />
        </div>
        {!sampleRow && state.rows.length > 0 && (
          <p className="mt-2 font-mono-data text-[11px] text-destructive">
            Sample row #{state.sampleIdB} out of range — tokens render blank.
          </p>
        )}
      </div>

      <Button
        onClick={onExecute}
        size="lg"
        className="glow-amber w-full bg-[var(--amber)] text-black hover:bg-[var(--amber)]/90"
      >
        <Copy /> Copy rich HTML & open mail (300ms)
      </Button>
      <p className="text-center font-mono-data text-[11px] text-muted-foreground">
        Body left empty — paste from clipboard inside your mail client to preserve styling.
      </p>
    </div>
  );
}

/* ----------------------------- Field ----------------------------- */

function Field({ label, children }: { label: string; children: React.ReactNode }) {
  return (
    <div className="space-y-1.5">
      <Label className="text-[10px] uppercase tracking-wider text-muted-foreground">
        {label}
      </Label>
      {children}
    </div>
  );
}
