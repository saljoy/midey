import { createFileRoute } from "@tanstack/react-router";
import { useCallback, useEffect, useMemo, useRef, useState } from "react";
import Papa from "papaparse";
import * as XLSX from "xlsx";
// react-window removed — queue is no longer displayed as a scrolling list.
import { toast } from "sonner";
import {
  Upload, Sun, Moon, Trash2, Send, Mail, Code2, Copy,
  Zap, FileText, Eye, SkipForward, Save, AlertTriangle,
  Bold, Italic, Underline, Strikethrough, List, ListOrdered,
  AlignLeft, AlignCenter, AlignRight, AlignJustify, Link2, Minus,
  Palette, Highlighter, CornerDownLeft, RotateCcw,
} from "lucide-react";
import { Button } from "@/components/ui/button";
import { Input } from "@/components/ui/input";
import { Textarea } from "@/components/ui/textarea";
import { Label } from "@/components/ui/label";
import {
  Dialog, DialogContent, DialogDescription, DialogFooter,
  DialogHeader, DialogTitle, DialogTrigger,
} from "@/components/ui/dialog";

export const Route = createFileRoute("/")({
  head: () => ({
    meta: [
      { title: "Wayne Enterprises Outreach Lab" },
      { name: "description", content: "High-performance mobile outreach console: parse 50MB+ CSV lead lists, fire personalized mailto handoffs, and craft rich HTML drafts — all 100% client-side." },
      { property: "og:title", content: "Wayne Enterprises Outreach Lab" },
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
  templateSlotsA: { name: string; subject: string; body: string }[];
  htmlMode: boolean;
}

const STORAGE_KEY = "midey.outreach.v1";
const THEME_KEY = "midey.theme";

const DEFAULT_STATE: PersistedState = {
  headers: [],
  rows: [],
  rowStates: {},
  targetEmailHeader: "",
  subjectA: "Quick question, {first_name}",
  bodyA: "Hi {first_name},\n\nNoticed {company} — wanted to reach out.\n\n— Wayne",
  recipientB: "",
  sampleIdB: 0,
  subjectB: "A note for {first_name}",
  htmlB: "<div style=\"font-family:system-ui;line-height:1.55\">\n  <h2 style=\"color:#0ea5e9\">Hi {first_name} 👋</h2>\n  <p>Loved what you're doing at <b>{company}</b>.</p>\n  <p>— Wayne Enterprises</p>\n</div>",
  templateSlotsA: [],
  htmlMode: false,
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

/**
 * Normalize an email cell that may contain multiple addresses
 * separated by commas or semicolons (with stray whitespace).
 * Returns a comma-joined, space-free string suitable for mailto:.
 */
function cleanEmails(raw: string): string {
  return (raw || "")
    .split(/[,;:\s|]+/)
    .map((e) => e.trim())
    .filter(Boolean)
    .join(",");
}

/**
 * Build a mailto: URL where ALL recipients are placed as a single
 * comma-separated list in the primary "To:" field (no BCC).
 * Subject and Body are independently passed through encodeURIComponent
 * so spaces, line breaks, and reserved symbols survive the trip into
 * mobile Gmail. Recipient commas are percent-encoded (%2C) per RFC 6068
 * to avoid mobile clients truncating the query string.
 */
function buildMailto(rawRecipients: string, params: Record<string, string>): string {
  const to = cleanEmails(rawRecipients).split(",").filter(Boolean).join("%2C");
  const qs: string[] = [];
  for (const [k, v] of Object.entries(params)) {
    if (v === undefined || v === null) continue;
    const encodedKey = encodeURIComponent(k);
    const encodedVal = encodeURIComponent(String(v));
    qs.push(`${encodedKey}=${encodedVal}`);
  }
  return `mailto:${to}${qs.length ? `?${qs.join("&")}` : ""}`;
}

/**
 * Auto-format a raw HTML template so plain newlines in the editor
 * become visible paragraph / line breaks in the rendered output.
 * - Blank-line separated chunks → wrapped in <p>…</p>
 * - Single \n inside a chunk → <br />
 * - Chunks that already start with a block-level tag are left as-is.
 */
const BLOCK_RE =
  /^\s*<(?:p|div|h[1-6]|ul|ol|li|table|thead|tbody|tr|td|th|blockquote|pre|hr|section|article|header|footer|nav|figure|figcaption|img|iframe|br)\b/i;

function autoFormatHtml(src: string): string {
  if (!src) return src;
  const chunks = src.split(/\n{2,}/);
  return chunks
    .map((chunk) => {
      const trimmed = chunk.trim();
      if (!trimmed) return "";
      if (BLOCK_RE.test(trimmed)) {
        // Already starts with a block tag — preserve, but still convert
        // bare single newlines inside to <br /> so layout matches editor.
        return trimmed.replace(/\n/g, "<br />");
      }
      return `<p>${trimmed.replace(/\n/g, "<br />")}</p>`;
    })
    .filter(Boolean)
    .join("\n");
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

    const isXlsx = /\.xlsx$/i.test(file.name);
    if (isXlsx) {
      const reader = new FileReader();
      reader.onerror = () => {
        setParsing(false);
        toast.error("Failed to read Excel file");
      };
      reader.onload = (e) => {
        try {
          const data = new Uint8Array(e.target?.result as ArrayBuffer);
          const wb = XLSX.read(data, { type: "array" });
          const sheetName = wb.SheetNames[0];
          if (!sheetName) throw new Error("Workbook has no sheets");
          const sheet = wb.Sheets[sheetName];
          const aoa = XLSX.utils.sheet_to_json<unknown[]>(sheet, {
            header: 1,
            blankrows: false,
            defval: "",
            raw: false,
          });
          if (!aoa.length) throw new Error("Sheet is empty");
          const seen = new Set<string>();
          const rawHeaders = (aoa[0] as unknown[]).map((h) => String(h ?? "").trim());
          const headers = rawHeaders.filter((h) => {
            if (!h) return false;
            if (h.length > 64) return false;
            if (/[,\n\r"]/.test(h)) return false;
            if (seen.has(h)) return false;
            seen.add(h);
            return true;
          });
          const rows: Row[] = [];
          for (let i = 1; i < aoa.length && rows.length < 500_000; i++) {
            const row = aoa[i] as unknown[];
            const obj: Row = {};
            for (let c = 0; c < rawHeaders.length; c++) {
              const key = rawHeaders[c];
              if (!key || !headers.includes(key)) continue;
              obj[key] = String(row?.[c] ?? "");
            }
            rows.push(obj);
          }
          setParsing(false);
          setParseProgress(100);
          const guessEmail = headers.find((h) => /e?mail/i.test(h)) ?? headers[0] ?? "";
          setState((s) => ({
            ...s,
            headers,
            rows,
            rowStates: {},
            targetEmailHeader: s.targetEmailHeader || guessEmail,
          }));
          toast.success(`Loaded ${rows.length.toLocaleString()} rows · ${headers.length} columns`);
        } catch (err) {
          setParsing(false);
          toast.error(`Parse failed: ${(err as Error).message}`);
        }
      };
      reader.readAsArrayBuffer(file);
      return;
    }

    const collected: Row[] = [];
    let headers: string[] = [];
    const total = file.size || 1;

    Papa.parse<Row>(file, {
      header: true,
      skipEmptyLines: true,
      // NOTE: worker:true + chunk callback is unreliable in some browsers
      // (cursor never updates → progress stuck at 0%, complete never fires).
      // Stream on the main thread in 1MB chunks; React state updates
      // between chunks keep the UI responsive even for 50MB+ files.
      chunkSize: 1024 * 1024,
      chunk: (results, parser) => {
        if (!headers.length && results.meta.fields) {
          // Sanitize: only accept clean column-name tokens from row 1.
          // Reject anything that looks like a data row (commas, newlines,
          // very long strings, or duplicates) so row data can never
          // leak into the token blueprint UI.
          const seen = new Set<string>();
          headers = results.meta.fields
            .map((h) => (h ?? "").trim())
            .filter((h) => {
              if (!h) return false;
              if (h.length > 64) return false;
              if (/[,\n\r"]/.test(h)) return false;
              if (seen.has(h)) return false;
              seen.add(h);
              return true;
            });
        }
        for (const r of results.data) collected.push(r);
        const cursor = (results.meta as { cursor?: number }).cursor ?? 0;
        const pct = cursor > 0
          ? Math.min(99, Math.round((cursor / total) * 100))
          : Math.min(99, Math.round((collected.length / Math.max(1000, collected.length + 1000)) * 100));
        setParseProgress(pct);
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

  /* ---------- Clear all ---------- */

  const clearAll = () => {
    localStorage.removeItem(STORAGE_KEY);
    setState(DEFAULT_STATE);
    toast.success("All data cleared.");
  };

  /* ---------- Section B: HTML preview + dual action ---------- */

  const sampleRow = state.rows[state.sampleIdB];
  const renderedHtml = useMemo(
    () => autoFormatHtml(renderTemplate(state.htmlB, sampleRow)),
    [state.htmlB, sampleRow],
  );
  const renderedSubjectB = useMemo(() => renderTemplate(state.subjectB, sampleRow), [state.subjectB, sampleRow]);

  const executeHtml = useCallback(async () => {
    const recipients = cleanEmails(state.recipientB);
    if (!recipients) { toast.error("Recipient required"); return; }
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
      window.location.href = buildMailto(state.recipientB, { subject: renderedSubjectB });
    }, 300);
  }, [state.recipientB, renderedHtml, renderedSubjectB]);

  /* Plain-text test sandbox: uses sample row + manual recipient, opens mailto with subject+body. */
  const renderedSubjectAPreview = useMemo(
    () => renderTemplate(state.subjectA, sampleRow),
    [state.subjectA, sampleRow],
  );
  const executePlainTest = useCallback(() => {
    const recipients = cleanEmails(state.recipientB);
    if (!recipients) { toast.error("Recipient required"); return; }
    const subject = renderTemplate(state.subjectA, sampleRow);
    const body = renderTemplate(state.bodyA, sampleRow);
    toast.success("Opening test draft…");
    window.location.href = buildMailto(state.recipientB, { subject, body });
  }, [state.recipientB, state.subjectA, state.bodyA, sampleRow]);

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
        />

        <div className="mt-6 space-y-4">
          <SectionACard
            state={state}
            patch={patch}
            queue={queue}
            processedCount={processedCount}
            fireRow={fireRow}
            skipRow={skipRow}
            resetRow={resetRow}
            executeTestHtml={executeHtml}
            renderedTestHtml={renderedHtml}
            renderedTestSubject={renderedSubjectB}
            executeTestPlain={executePlainTest}
            renderedTestSubjectPlain={renderedSubjectAPreview}
            sampleRow={sampleRow}
          />
        </div>
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
              Wayne Enterprises <span className="text-sky-glow">Outreach Lab</span>
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
  targetEmailHeader, onTargetEmailHeader,
}: {
  parsing: boolean;
  progress: number;
  onFile: (f: File) => void;
  headers: string[];
  totalRows: number;
  processedRows: number;
  targetEmailHeader: string;
  onTargetEmailHeader: (v: string) => void;
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
            <Label className="mb-2 block text-[10px] uppercase tracking-wider text-muted-foreground">
              Tokens · tap to copy
            </Label>
            <div className="flex flex-wrap gap-1.5">
              {headers.map((h) => {
                const tok = `{${h}}`;
                return (
                  <button
                    key={h}
                    type="button"
                    onClick={async () => {
                      try {
                        await navigator.clipboard.writeText(tok);
                        toast.success(`Copied ${tok}`);
                      } catch {
                        toast.error("Copy failed");
                      }
                    }}
                    className="rounded-md border border-border-strong/70 bg-surface-2 px-2 py-1 font-mono-data text-[11px] text-foreground hover:glow-sky"
                  >
                    {tok}
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
  executeTestHtml, renderedTestHtml, renderedTestSubject, sampleRow,
  executeTestPlain, renderedTestSubjectPlain,
}: {
  state: PersistedState;
  patch: (p: Partial<PersistedState>) => void;
  queue: number[];
  processedCount: number;
  fireRow: (i: number) => void;
  skipRow: (i: number) => void;
  resetRow: (i: number) => void;
  executeTestHtml: () => void;
  renderedTestHtml: string;
  renderedTestSubject: string;
  executeTestPlain: () => void;
  renderedTestSubjectPlain: string;
  sampleRow: Row | undefined;
}) {
  const firstPendingIndex = queue.find(
    (i) => (state.rowStates[i] ?? "pending") === "pending",
  );
  // Manual override — "Jump to row" input or "Resend" button on a processed row.
  const [activeOverride, setActiveOverride] = useState<number | null>(null);
  const [jumpInput, setJumpInput] = useState<string>("");
  const nextPendingIndex =
    activeOverride !== null && state.rows[activeOverride]
      ? activeOverride
      : firstPendingIndex;
  const pendingCount = state.rows.length - processedCount;
  const [filter, setFilter] = useState<"all" | "active" | "processed">("all");
  const htmlTextareaRef = useRef<HTMLTextAreaElement>(null);
  const [processedSearch, setProcessedSearch] = useState("");
  const [queueSearch, setQueueSearch] = useState("");
  const processedIndices = useMemo(
    () => Object.entries(state.rowStates)
      .filter(([, v]) => v === "processed")
      .map(([k]) => Number(k))
      .sort((a, b) => a - b),
    [state.rowStates],
  );
  const filteredProcessedIndices = useMemo(() => {
    const q = processedSearch.trim().toLowerCase();
    if (!q) return processedIndices;
    return processedIndices.filter((i) => {
      if (String(i).includes(q)) return true;
      const row = state.rows[i];
      if (!row) return false;
      const email = String(row[state.targetEmailHeader] ?? "").toLowerCase();
      if (email.includes(q)) return true;
      return Object.values(row).some((v) => String(v ?? "").toLowerCase().includes(q));
    });
  }, [processedSearch, processedIndices, state.rows, state.targetEmailHeader]);

  /* Active queue email search — matches across all rows respecting the
   * current filter (all / active). Processed mode uses processedSearch. */
  const queueSearchMatches = useMemo(() => {
    const q = queueSearch.trim().toLowerCase();
    if (!q) return [] as number[];
    const matches: number[] = [];
    for (let i = 0; i < state.rows.length; i++) {
      if (filter === "active" && state.rowStates[i] === "processed") continue;
      const email = String(state.rows[i]?.[state.targetEmailHeader] ?? "").toLowerCase();
      if (email.includes(q) || String(i).includes(q)) matches.push(i);
      if (matches.length >= 50) break;
    }
    return matches;
  }, [queueSearch, state.rows, state.rowStates, state.targetEmailHeader, filter]);

  // Char counter for active row's mailto string (plain-text mode only)
  const previewRow = state.rows[nextPendingIndex ?? -1];
  const previewTo = cleanEmails(previewRow?.[state.targetEmailHeader] ?? "");
  const previewSubject = renderTemplate(state.subjectA, previewRow);
  const previewBody = renderTemplate(state.bodyA, previewRow);
  const mailtoLen = previewRow && !state.htmlMode
    ? `mailto:${previewTo}?subject=${encodeURIComponent(previewSubject)}&body=${encodeURIComponent(previewBody)}`.length
    : 0;
  const overLimit = mailtoLen > 2000;

  // Template slot manager
  const [slotName, setSlotName] = useState("");
  const [slotPickerOpen, setSlotPickerOpen] = useState(false);
  const saveSlot = () => {
    const name = slotName.trim();
    if (!name) { toast.error("Slot name required"); return; }
    const next = state.templateSlotsA.filter((s) => s.name !== name);
    next.push({ name, subject: state.subjectA, body: state.bodyA });
    patch({ templateSlotsA: next });
    setSlotName("");
    toast.success(`Saved slot "${name}"`);
  };
  const loadSlot = (name: string) => {
    const slot = state.templateSlotsA.find((s) => s.name === name);
    if (!slot) return;
    patch({ subjectA: slot.subject, bodyA: slot.body });
    setSlotPickerOpen(false);
    toast.success(`Loaded "${name}"`);
  };
  const deleteSlot = (name: string) => {
    patch({ templateSlotsA: state.templateSlotsA.filter((s) => s.name !== name) });
    toast.success(`Deleted "${name}"`);
  };

  return (
    <div className="space-y-4 rounded-xl border border-border-strong/70 bg-surface-1 p-4">
      {/* HTML mode toggle */}
      <label className="flex items-center justify-between gap-3 rounded-lg border border-border-strong/60 bg-surface-2 p-2.5">
        <span className="flex items-center gap-2 font-mono-data text-[11px] uppercase tracking-wider text-muted-foreground">
          <Code2 className="size-3.5" />
          Use HTML template
        </span>
        <input
          type="checkbox"
          checked={state.htmlMode}
          onChange={(e) => patch({ htmlMode: e.target.checked })}
          className="size-4 accent-[var(--sky)]"
        />
      </label>

      {/* Template slot manager */}
      <div className="flex flex-wrap items-center gap-2 rounded-lg border border-border-strong/60 bg-surface-2 p-2">
        <span className="font-mono-data text-[10px] uppercase tracking-wider text-muted-foreground">
          Template slots
        </span>
        <Input
          value={slotName}
          onChange={(e) => setSlotName(e.target.value)}
          placeholder="Name (e.g. Klaviyo Hook)"
          className="h-8 max-w-[180px] font-mono-data text-xs"
        />
        <Button size="sm" variant="ghost" onClick={saveSlot} className="h-8">
          <Save className="size-3.5" /> Save
        </Button>
        <div className="relative">
          <Button
            size="sm"
            variant="ghost"
            className="h-8"
            onClick={() => setSlotPickerOpen((v) => !v)}
            disabled={state.templateSlotsA.length === 0}
          >
            Load ({state.templateSlotsA.length}) ▾
          </Button>
          {slotPickerOpen && state.templateSlotsA.length > 0 && (
            <div className="absolute right-0 z-20 mt-1 w-56 rounded-md border border-border-strong/70 bg-surface-1 p-1 shadow-lg">
              {state.templateSlotsA.map((s) => (
                <div key={s.name} className="flex items-center gap-1">
                  <button
                    type="button"
                    onClick={() => loadSlot(s.name)}
                    className="flex-1 truncate rounded px-2 py-1.5 text-left font-mono-data text-xs hover:bg-surface-2"
                  >
                    {s.name}
                  </button>
                  <button
                    type="button"
                    onClick={() => deleteSlot(s.name)}
                    aria-label={`Delete ${s.name}`}
                    className="rounded px-1.5 py-1 text-muted-foreground hover:text-destructive"
                  >
                    <Trash2 className="size-3" />
                  </button>
                </div>
              ))}
            </div>
          )}
        </div>
      </div>

      <div className="grid gap-3 sm:grid-cols-2">
        <Field label="Subject template">
          <Input
            value={state.htmlMode ? state.subjectB : state.subjectA}
            onChange={(e) =>
              patch(state.htmlMode ? { subjectB: e.target.value } : { subjectA: e.target.value })
            }
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
      {state.htmlMode ? (
        <>
          <Field label="HTML code template">
            <HtmlToolbar
              textareaRef={htmlTextareaRef}
              value={state.htmlB}
              onChange={(v) => patch({ htmlB: v })}
            />
            <Textarea
              ref={htmlTextareaRef}
              value={state.htmlB}
              onChange={(e) => patch({ htmlB: e.target.value })}
              rows={8}
              className="font-mono-data text-[12px] leading-relaxed rounded-t-none border-t-0"
              spellCheck={false}
              placeholder="<div>Hi {first_name}…</div>"
            />
          </Field>
          <div>
            <div className="mb-2 flex items-center gap-2">
              <Eye className="size-3.5 text-muted-foreground" />
              <span className="font-mono-data text-[10px] uppercase tracking-wider text-muted-foreground">
                Live preview
              </span>
            </div>
            <div className="overflow-hidden rounded-lg border border-border-strong/60 bg-white">
              <iframe
                title="HTML preview"
                sandbox=""
                srcDoc={`<!doctype html><html><body style="margin:0;padding:12px;font-family:system-ui">${autoFormatHtml(renderTemplate(state.htmlB, previewRow ?? sampleRow))}</body></html>`}
                className="block h-[300px] w-full"
              />
            </div>
          </div>

          {/* Test sandbox */}
          <div className="space-y-2 rounded-lg border border-amber-glow/40 bg-amber-glow/5 p-3">
            <div className="flex items-center gap-2 font-mono-data text-[10px] uppercase tracking-wider text-amber-glow">
              <Zap className="size-3.5" /> Test sandbox · does not advance queue
            </div>
            <div className="grid gap-2 sm:grid-cols-2">
              <Input
                type="email"
                value={state.recipientB}
                onChange={(e) => patch({ recipientB: e.target.value })}
                placeholder="manual test email"
                className="h-9 font-mono-data text-xs"
              />
              <Input
                type="number"
                min={0}
                max={Math.max(0, state.rows.length - 1)}
                value={state.sampleIdB}
                onChange={(e) => patch({ sampleIdB: Math.max(0, Number(e.target.value) || 0) })}
                placeholder="sample row id"
                className="h-9 font-mono-data text-xs"
              />
            </div>
            <Button
              size="sm"
              onClick={executeTestHtml}
              className="glow-amber w-full bg-[var(--amber)] text-black hover:bg-[var(--amber)]/90"
            >
              <Copy className="size-3.5" /> Send test draft
            </Button>
            <p className="font-mono-data text-[10px] text-muted-foreground">
              Subject preview: <span className="text-amber-glow">{renderedTestSubject || "—"}</span>
            </p>
          </div>
        </>
      ) : (
        <>
          <Field label="Plain-text body template">
            <Textarea
              value={state.bodyA}
              onChange={(e) => patch({ bodyA: e.target.value })}
              rows={6}
              className="font-mono-data text-[13px]"
              placeholder="Hi {first_name}, …"
            />
          </Field>

          {/* Test sandbox · plain text */}
          <div className="space-y-2 rounded-lg border border-amber-glow/40 bg-amber-glow/5 p-3">
            <div className="flex items-center gap-2 font-mono-data text-[10px] uppercase tracking-wider text-amber-glow">
              <Zap className="size-3.5" /> Test sandbox · does not advance queue
            </div>
            <div className="grid gap-2 sm:grid-cols-2">
              <Input
                type="email"
                value={state.recipientB}
                onChange={(e) => patch({ recipientB: e.target.value })}
                placeholder="manual test email"
                className="h-9 font-mono-data text-xs"
              />
              <Input
                type="number"
                min={0}
                max={Math.max(0, state.rows.length - 1)}
                value={state.sampleIdB}
                onChange={(e) => patch({ sampleIdB: Math.max(0, Number(e.target.value) || 0) })}
                placeholder="sample row id"
                className="h-9 font-mono-data text-xs"
              />
            </div>
            <Button
              size="sm"
              onClick={executeTestPlain}
              className="glow-amber w-full bg-[var(--amber)] text-black hover:bg-[var(--amber)]/90"
            >
              <Copy className="size-3.5" /> Send test draft
            </Button>
            <p className="font-mono-data text-[10px] text-muted-foreground">
              Subject preview: <span className="text-amber-glow">{renderedTestSubjectPlain || "—"}</span>
            </p>
          </div>
        </>
      )}
      <div className="flex flex-wrap items-center justify-between gap-2 -mt-2">
        <span className="font-mono-data text-[11px] text-muted-foreground">
          {state.htmlMode
            ? "html mode · body sent via clipboard"
            : <>mailto length: <span className={overLimit ? "text-amber-glow" : "text-foreground"}>{mailtoLen.toLocaleString()}</span> chars</>}
        </span>
        {overLimit && (
          <span className="inline-flex items-center gap-1 rounded-md border border-amber-glow/40 bg-amber-glow/10 px-2 py-0.5 font-mono-data text-[10px] text-amber-glow">
            <AlertTriangle className="size-3" /> Approaching mobile link limit
          </span>
        )}
      </div>

      <div className="flex items-center justify-between">
        <div className="font-mono-data text-xs text-muted-foreground">
          Queue · <span className="text-foreground">{pendingCount.toLocaleString()}</span> pending ·{" "}
          <span className="text-sky-glow">{processedCount.toLocaleString()}</span> done
        </div>
        <div className="font-mono-data text-[10px] uppercase tracking-wider text-muted-foreground">
          headless
        </div>
      </div>

      {/* Quick queue filters */}
      <div className="flex flex-wrap items-center justify-between gap-2">
        <div className="flex gap-1.5">
          {(["all", "active", "processed"] as const).map((k) => (
            <button
              key={k}
              type="button"
              onClick={() => setFilter(k)}
              className={`rounded-md border px-2.5 py-1 font-mono-data text-[11px] capitalize transition ${
                filter === k
                  ? "border-sky-glow/60 bg-sky-glow/10 text-sky-glow glow-sky"
                  : "border-border-strong/60 bg-surface-2 text-muted-foreground hover:text-foreground"
              }`}
            >
              {k === "active" ? "Active only" : k}
            </button>
          ))}
        </div>
        <div className="flex items-center gap-1.5">
          <label className="font-mono-data text-[10px] uppercase tracking-wider text-muted-foreground">
            Jump to row #
          </label>
          <Input
            type="number"
            min={0}
            max={Math.max(0, state.rows.length - 1)}
            value={jumpInput}
            onChange={(e) => {
              const v = e.target.value;
              setJumpInput(v);
              if (v === "") { setActiveOverride(null); return; }
              const n = Number(v);
              if (Number.isFinite(n) && n >= 0 && n < state.rows.length) {
                setActiveOverride(n);
              }
            }}
            className="h-8 w-20 font-mono-data text-xs"
            placeholder="0"
            disabled={state.rows.length === 0}
          />
          {activeOverride !== null && (
            <button
              type="button"
              onClick={() => { setActiveOverride(null); setJumpInput(""); }}
              className="rounded-md border border-border-strong/60 bg-surface-2 px-2 py-1 font-mono-data text-[10px] text-muted-foreground hover:text-foreground"
              title="Clear override · resume normal queue"
            >
              <RotateCcw className="inline size-3" /> reset
            </button>
          )}
        </div>
      </div>

      {/* Active queue email search */}
      {filter !== "processed" && state.rows.length > 0 && (
        <div className="space-y-2">
          <Input
            value={queueSearch}
            onChange={(e) => setQueueSearch(e.target.value)}
            placeholder="Search by Email…"
            className="h-8 font-mono-data text-xs"
          />
          {queueSearch.trim() && (
            <div className="rounded-md border border-border-strong/60 bg-surface-2 p-2">
              {queueSearchMatches.length === 0 ? (
                <p className="py-2 text-center font-mono-data text-[11px] text-muted-foreground">
                  No matches for "{queueSearch}".
                </p>
              ) : (
                <ul className="max-h-56 space-y-1 overflow-auto">
                  {queueSearchMatches.map((i) => {
                    const isProcessed = state.rowStates[i] === "processed";
                    return (
                      <li
                        key={i}
                        className={`flex items-center justify-between gap-2 rounded border border-border-strong/40 bg-bg-app px-2 py-1 font-mono-data text-[11px] ${
                          isProcessed ? "opacity-60" : ""
                        }`}
                      >
                        <span className="min-w-0 truncate">
                          <span className="text-muted-foreground">#{i}</span> ·{" "}
                          <span className="text-foreground">
                            {state.rows[i]?.[state.targetEmailHeader] ?? "—"}
                          </span>
                          {isProcessed && (
                            <span className="ml-2 rounded border border-sky-glow/40 bg-sky-glow/10 px-1 py-0.5 text-[10px] text-sky-glow">
                              done
                            </span>
                          )}
                        </span>
                        <Button
                          size="sm"
                          variant="ghost"
                          className="h-6 px-2 text-sky-glow hover:bg-sky-glow/10"
                          onClick={() => {
                            setActiveOverride(i);
                            setJumpInput(String(i));
                            setQueueSearch("");
                            toast.success(`Loaded row #${i}`);
                          }}
                        >
                          Load
                        </Button>
                      </li>
                    );
                  })}
                </ul>
              )}
            </div>
          )}
        </div>
      )}

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
        ) : filter === "processed" ? (
          processedIndices.length === 0 ? (
            <p className="py-6 text-center font-mono-data text-xs text-muted-foreground">
              No rows processed yet.
            </p>
          ) : (
            <div className="space-y-2">
              <Input
                value={processedSearch}
                onChange={(e) => setProcessedSearch(e.target.value)}
                placeholder="Search processed by email, row #, or any field…"
                className="h-8 font-mono-data text-xs"
              />
              {filteredProcessedIndices.length === 0 ? (
                <p className="py-4 text-center font-mono-data text-xs text-muted-foreground">
                  No matches for "{processedSearch}".
                </p>
              ) : (
                <ul className="max-h-72 space-y-1 overflow-auto">
                  {filteredProcessedIndices.map((i) => (
                <li key={i} className="flex items-center justify-between gap-2 rounded border border-border-strong/40 bg-surface-2 px-2 py-1 font-mono-data text-[11px]">
                  <span className="min-w-0 truncate">
                    <span className="text-muted-foreground">#{i}</span> ·{" "}
                    <span className="text-foreground">{state.rows[i]?.[state.targetEmailHeader] ?? "—"}</span>
                  </span>
                  <Button
                    size="sm"
                    variant="ghost"
                    className="h-6 px-2 text-sky-glow hover:bg-sky-glow/10"
                    onClick={() => {
                      setActiveOverride(i);
                      setJumpInput(String(i));
                      setFilter("all");
                      toast.success(`Loaded row #${i} for resend`);
                    }}
                  >
                    <RotateCcw className="size-3" /> Resend
                  </Button>
                </li>
                  ))}
                </ul>
              )}
            </div>
          )
        ) : nextPendingIndex === undefined ? (
          <p className="py-6 text-center font-mono-data text-xs text-muted-foreground">
            All rows processed.
          </p>
        ) : (
          <NextRowPreview
            rowIndex={nextPendingIndex}
            row={state.rows[nextPendingIndex]}
            targetEmailHeader={state.targetEmailHeader}
            subjectTpl={state.htmlMode ? state.subjectB : state.subjectA}
            bodyTpl={state.bodyA}
            htmlMode={state.htmlMode}
            htmlTpl={state.htmlB}
            onSend={() => {
              fireRow(nextPendingIndex);
              if (activeOverride !== null) {
                setActiveOverride(null);
                setJumpInput("");
              }
            }}
            onSkip={() => {
              skipRow(nextPendingIndex);
              if (activeOverride !== null) {
                setActiveOverride(null);
                setJumpInput("");
              }
            }}
            isResend={state.rowStates[nextPendingIndex] === "processed"}
          />
        )}
      </div>
    </div>
  );
}

function NextRowPreview({
  rowIndex, row, targetEmailHeader, subjectTpl, bodyTpl, htmlMode, htmlTpl, onSend, onSkip, isResend,
}: {
  rowIndex: number;
  row: Row | undefined;
  targetEmailHeader: string;
  subjectTpl: string;
  bodyTpl: string;
  htmlMode: boolean;
  htmlTpl: string;
  onSend: () => void;
  onSkip: () => void;
  isResend?: boolean;
}) {
  const rawRecipients = row?.[targetEmailHeader] ?? "";
  const toAddr = cleanEmails(rawRecipients);
  const subject = renderTemplate(subjectTpl, row);
  const body = renderTemplate(bodyTpl, row);
  const renderedHtml = autoFormatHtml(renderTemplate(htmlTpl, row));
  const plainHref = toAddr
    ? buildMailto(rawRecipients, { subject, body })
    : "";
  const htmlHref = toAddr
    ? buildMailto(rawRecipients, { subject })
    : "";
  const sendHtml = async () => {
    if (!toAddr) return;
    // Snapshot the current row's mailto BEFORE advancing the queue,
    // so re-render from onSend() can't swap in the next row's link.
    const hrefSnapshot = htmlHref;
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
    onSend();
    setTimeout(() => { window.location.href = hrefSnapshot; }, 300);
  };
  return (
    <div className="space-y-3">
      <div className="flex items-center justify-between">
        <div className="font-mono-data text-xs text-muted-foreground">
          {isResend ? "Resend · row " : "Next up · row "}
          <span className="text-foreground">#{rowIndex}</span>
          {isResend && (
            <span className="ml-2 rounded border border-sky-glow/40 bg-sky-glow/10 px-1.5 py-0.5 text-[10px] text-sky-glow">
              processed
            </span>
          )}
        </div>
      </div>
      <div className="space-y-2 rounded-md border border-border-strong/60 bg-surface-2 p-3">
        <div className="font-mono-data text-[11px]">
          <span className="text-muted-foreground">To: </span>
          <span className="text-foreground">{toAddr || <span className="text-destructive">— missing —</span>}</span>
        </div>
        <div className="font-mono-data text-[11px]">
          <span className="text-muted-foreground">Subject: </span>
          <span className="text-amber-glow">{subject || "—"}</span>
        </div>
        {htmlMode ? (
          <div className="overflow-hidden rounded border border-border-strong/40 bg-white">
            <iframe
              title="row html preview"
              sandbox=""
              srcDoc={`<!doctype html><html><body style="margin:0;padding:10px;font-family:system-ui">${renderedHtml}</body></html>`}
              className="block h-40 w-full"
            />
          </div>
        ) : (
        <pre className="max-h-48 overflow-auto whitespace-pre-wrap break-words rounded bg-bg-app p-2 font-mono-data text-[11px] leading-relaxed text-foreground">
{body || "—"}
        </pre>
        )}
      </div>
      <div className="flex justify-end gap-2">
        <Button size="sm" variant="ghost" onClick={onSkip}>
          <SkipForward className="size-3.5" /> Skip
        </Button>
        {!toAddr ? (
          <Button size="sm" disabled>
            <Send className="size-3.5" /> Send current
          </Button>
        ) : htmlMode ? (
          <Button
            size="sm"
            onClick={sendHtml}
            className="glow-amber bg-[var(--amber)] text-black hover:bg-[var(--amber)]/90"
          >
            <Send className="size-3.5" /> Send current
          </Button>
        ) : (
          <Button
            size="sm"
            className="glow-amber bg-[var(--amber)] text-black hover:bg-[var(--amber)]/90"
            onClick={() => {
              // Trigger navigation with the CURRENT row's href first,
              // then mark this row processed so the queue advances after.
              const hrefSnapshot = plainHref;
              window.location.href = hrefSnapshot;
              onSend();
            }}
          >
            <Send className="size-3.5" /> Send current
          </Button>
        )}
      </div>
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

/* --------------------------- HTML Toolbar --------------------------- */

function HtmlToolbar({
  textareaRef, value, onChange,
}: {
  textareaRef: React.RefObject<HTMLTextAreaElement | null>;
  value: string;
  onChange: (v: string) => void;
}) {
  const replaceSelection = useCallback(
    (transform: (sel: string) => string, fallback = "text") => {
      const ta = textareaRef.current;
      if (!ta) return;
      const start = ta.selectionStart ?? value.length;
      const end = ta.selectionEnd ?? value.length;
      const sel = value.slice(start, end) || fallback;
      const out = transform(sel);
      const next = value.slice(0, start) + out + value.slice(end);
      onChange(next);
      requestAnimationFrame(() => {
        ta.focus();
        ta.setSelectionRange(start, start + out.length);
      });
    },
    [textareaRef, value, onChange],
  );

  const wrap = (open: string, close: string) =>
    replaceSelection((s) => `${open}${s}${close}`);

  const makeList = (ordered: boolean) =>
    replaceSelection((s) => {
      const tag = ordered ? "ol" : "ul";
      const items = s.split(/\r?\n/).filter(Boolean).map((l) => `  <li>${l}</li>`).join("\n");
      return `<${tag}>\n${items || "  <li>item</li>"}\n</${tag}>`;
    }, "item 1\nitem 2");

  const align = (a: "left" | "center" | "right" | "justify") =>
    replaceSelection((s) => `<div style="text-align:${a}">${s}</div>`);

  const fontSize = (px: string) =>
    replaceSelection((s) => `<span style="font-size:${px}">${s}</span>`);

  const insertLink = () => {
    const url = window.prompt("Link URL", "https://");
    if (!url) return;
    replaceSelection((s) => `<a href="${url}">${s}</a>`, "link text");
  };

  const insertHr = () => replaceSelection(() => `<hr />`, "");

  const insertBreak = () => {
    const ta = textareaRef.current;
    if (!ta) return;
    const start = ta.selectionStart ?? value.length;
    const end = ta.selectionEnd ?? value.length;
    const ins = `<br />`;
    const next = value.slice(0, start) + ins + value.slice(end);
    onChange(next);
    requestAnimationFrame(() => {
      ta.focus();
      const pos = start + ins.length;
      ta.setSelectionRange(pos, pos);
    });
  };

  const btn =
    "inline-flex h-8 w-8 items-center justify-center rounded border border-border-strong/60 bg-surface-1 text-muted-foreground hover:bg-surface-2 hover:text-foreground transition-colors";

  return (
    <div className="flex flex-wrap items-center gap-1 rounded-t-md border border-b-0 border-border-strong/60 bg-surface-2/40 p-1.5">
      <button type="button" title="Bold" className={btn} onClick={() => wrap("<strong>", "</strong>")}><Bold className="size-3.5" /></button>
      <button type="button" title="Italic" className={btn} onClick={() => wrap("<em>", "</em>")}><Italic className="size-3.5" /></button>
      <button type="button" title="Underline" className={btn} onClick={() => wrap(`<span style="text-decoration: underline;">`, "</span>")}><Underline className="size-3.5" /></button>
      <button type="button" title="Strikethrough" className={btn} onClick={() => wrap("<del>", "</del>")}><Strikethrough className="size-3.5" /></button>

      <span className="mx-1 h-5 w-px bg-border-strong/60" />

      <label title="Text color" className={`${btn} relative cursor-pointer`}>
        <Palette className="size-3.5" />
        <input
          type="color"
          className="absolute inset-0 cursor-pointer opacity-0"
          onChange={(e) => wrap(`<span style="color:${e.target.value}">`, "</span>")}
        />
      </label>
      <label title="Highlight color" className={`${btn} relative cursor-pointer`}>
        <Highlighter className="size-3.5" />
        <input
          type="color"
          className="absolute inset-0 cursor-pointer opacity-0"
          onChange={(e) => wrap(`<span style="background-color:${e.target.value}">`, "</span>")}
        />
      </label>

      <span className="mx-1 h-5 w-px bg-border-strong/60" />

      <select
        title="Font size"
        defaultValue=""
        onChange={(e) => { if (e.target.value) { fontSize(e.target.value); e.target.value = ""; } }}
        className="h-8 rounded border border-border-strong/60 bg-surface-1 px-1.5 font-mono-data text-[11px] text-muted-foreground hover:text-foreground"
      >
        <option value="" disabled>Size</option>
        <option value="12px">Small</option>
        <option value="14px">Normal</option>
        <option value="18px">Large</option>
        <option value="24px">Huge</option>
      </select>

      <span className="mx-1 h-5 w-px bg-border-strong/60" />

      <button type="button" title="Bulleted list" className={btn} onClick={() => makeList(false)}><List className="size-3.5" /></button>
      <button type="button" title="Numbered list" className={btn} onClick={() => makeList(true)}><ListOrdered className="size-3.5" /></button>

      <span className="mx-1 h-5 w-px bg-border-strong/60" />

      <button type="button" title="Align left" className={btn} onClick={() => align("left")}><AlignLeft className="size-3.5" /></button>
      <button type="button" title="Align center" className={btn} onClick={() => align("center")}><AlignCenter className="size-3.5" /></button>
      <button type="button" title="Align right" className={btn} onClick={() => align("right")}><AlignRight className="size-3.5" /></button>
      <button type="button" title="Justify" className={btn} onClick={() => align("justify")}><AlignJustify className="size-3.5" /></button>

      <span className="mx-1 h-5 w-px bg-border-strong/60" />

      <button type="button" title="Insert link" className={btn} onClick={insertLink}><Link2 className="size-3.5" /></button>
      <button type="button" title="Horizontal rule" className={btn} onClick={insertHr}><Minus className="size-3.5" /></button>
      <button type="button" title="Line break (<br />)" className={btn} onClick={insertBreak}><CornerDownLeft className="size-3.5" /></button>
    </div>
  );
}
