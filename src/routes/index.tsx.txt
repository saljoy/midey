import { createFileRoute } from "@tanstack/react-router";
import { createContext, useCallback, useContext, useEffect, useMemo, useRef, useState } from "react";
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
import { Plus, ChevronDown, ChevronUp, Shuffle, Activity, ShieldAlert, History, Layers } from "lucide-react";
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

export interface TemplateItem {
  id: string;
  name: string;
  subject: string;
  body: string;
  html: string;
}

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
  templates: TemplateItem[];
  activeTemplateId: string;
  rotateSubjects: boolean;
  rotateBodies: boolean;
  sendCounter: number;
  dailyGoal: number;
  queueSearchPersisted: string;
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
  templates: [],
  activeTemplateId: "",
  rotateSubjects: false,
  rotateBodies: false,
  sendCounter: 0,
  dailyGoal: 100,
  queueSearchPersisted: "",
};

const AUTOSAVE_KEY = "midey.outreach.autosave.v1";
const SESSION_META_KEY = "midey.outreach.session.v1";
const AI_SETTINGS_KEY = "midey.outreach.ai.v1";

export interface AISettings {
  enabled: boolean;
  provider: "gemini" | "openai";
  apiKey: string;
  prompt: string;
  fallback: string;
  descriptionColumn: string;
}

const DEFAULT_AI: AISettings = {
  enabled: false,
  provider: "gemini",
  apiKey: "",
  prompt:
    "Analyze the following store description and write a single, natural, brief 7-word phrase complimenting a specific product category they specialize in. Do not use corporate jargon or exclamation marks.",
  fallback: "your unique collection",
  descriptionColumn: "store_description",
};

function loadAISettings(): AISettings {
  if (typeof window === "undefined") return DEFAULT_AI;
  try {
    const raw = localStorage.getItem(AI_SETTINGS_KEY);
    if (!raw) return DEFAULT_AI;
    return { ...DEFAULT_AI, ...JSON.parse(raw) };
  } catch { return DEFAULT_AI; }
}

async function generateAIInsight(
  ai: AISettings,
  description: string,
  signal?: AbortSignal,
): Promise<string> {
  const fullPrompt = `${ai.prompt}\n\nStore description:\n"""${description}"""`;
  if (ai.provider === "gemini") {
    const url = `https://generativelanguage.googleapis.com/v1beta/models/gemini-2.0-flash:generateContent?key=${encodeURIComponent(ai.apiKey)}`;
    const res = await fetch(url, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ contents: [{ parts: [{ text: fullPrompt }] }] }),
      signal,
    });
    if (!res.ok) throw new Error(`Gemini ${res.status}`);
    const j = await res.json();
    const txt = j?.candidates?.[0]?.content?.parts?.[0]?.text;
    return String(txt || "").trim().replace(/^["'`]+|["'`]+$/g, "");
  }
  // OpenAI
  const res = await fetch("https://api.openai.com/v1/chat/completions", {
    method: "POST",
    headers: {
      "Content-Type": "application/json",
      Authorization: `Bearer ${ai.apiKey}`,
    },
    body: JSON.stringify({
      model: "gpt-4o-mini",
      messages: [{ role: "user", content: fullPrompt }],
      temperature: 0.7,
    }),
    signal,
  });
  if (!res.ok) throw new Error(`OpenAI ${res.status}`);
  const j = await res.json();
  const txt = j?.choices?.[0]?.message?.content;
  return String(txt || "").trim().replace(/^["'`]+|["'`]+$/g, "");
}

function newId() {
  return `tpl_${Math.random().toString(36).slice(2, 9)}`;
}

/** Spam keywords / patterns commonly tripping aggressive filters. */
const SPAM_TERMS = [
  "free", "guarantee", "guaranteed", "urgent", "act now", "risk-free", "risk free",
  "winner", "cash", "click here", "buy now", "100%", "limited time", "no cost",
  "no obligation", "offer expires", "earn money", "double your", "make money",
  "amazing", "congratulations", "miracle", "lowest price", "best price",
];
function scanSpam(text: string): { hits: string[]; exclaim: number; allCaps: number } {
  const t = (text || "").toLowerCase();
  const hits: string[] = [];
  for (const term of SPAM_TERMS) {
    if (t.includes(term)) hits.push(term);
  }
  const exclaim = (text.match(/!/g) || []).length;
  const words = text.split(/\s+/).filter((w) => w.length >= 4);
  const allCaps = words.filter((w) => /^[A-Z]{4,}$/.test(w)).length;
  return { hits, exclaim, allCaps };
}
/** Validate hyperlinks inside an HTML template. */
function scanLinks(html: string): { ok: number; broken: { tag: string; reason: string }[] } {
  const out: { tag: string; reason: string }[] = [];
  let ok = 0;
  const re = /<a\b[^>]*>/gi;
  const matches = html.match(re) || [];
  for (const tag of matches) {
    const href = /href\s*=\s*("([^"]*)"|'([^']*)')/i.exec(tag);
    const url = href?.[2] ?? href?.[3] ?? "";
    if (!href) { out.push({ tag, reason: "missing href" }); continue; }
    if (!url) { out.push({ tag, reason: "empty href" }); continue; }
    if (/\s/.test(url)) { out.push({ tag, reason: "whitespace in url" }); continue; }
    if (!/^(https?:|mailto:|tel:|#|\/|\{)/i.test(url)) {
      out.push({ tag, reason: "no protocol" }); continue;
    }
    ok++;
  }
  // Unbalanced anchor tags
  const opens = (html.match(/<a\b/gi) || []).length;
  const closes = (html.match(/<\/a>/gi) || []).length;
  if (opens !== closes) out.push({ tag: `${opens} open vs ${closes} close`, reason: "unbalanced <a> tags" });
  return { ok, broken: out };
}

/* --------------------------- Utilities --------------------------- */

const TOKEN_RE = /\{([^{}]+)\}/g;

/**
 * Expand Spintax — {a|b|c} — picking ONE variant per occurrence based on
 * `seed` (the global send counter) instead of random chance, so the Nth
 * email sent always gets the Nth option from every bracket (looping back
 * to the start once a bracket runs out of options). This makes spintax
 * output deterministic and repeatable across a session: send #0 gets
 * option 1 from every bracket, send #1 gets option 2, etc.
 * Iterates innermost-first so nested groups like {Hi|{Hello|Hey}} work.
 */
function expandSpintax(src: string, seed: number): string {
  if (!src || src.indexOf("{") === -1) return src;
  const inner = /\{([^{}]*\|[^{}]*)\}/;
  let out = src;
  let guard = 0;
  while (inner.test(out) && guard++ < 500) {
    out = out.replace(inner, (_, body: string) => {
      const opts = body.split("|");
      const idx = ((seed % opts.length) + opts.length) % opts.length;
      return opts[idx] ?? "";
    });
  }
  return out;
}

function renderTemplate(
  tpl: string,
  row: Row | undefined,
  extras?: Record<string, string>,
  seed = 0,
): string {
  const expanded = expandSpintax(tpl, seed);
  return expanded.replace(TOKEN_RE, (_, key: string) => {
    const k = key.trim();
    if (extras && k in extras) return extras[k] ?? "";
    const raw = row?.[k];
    if (raw === undefined || raw === null || raw === "") return "";
    return String(raw);
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
    const parsed = JSON.parse(raw) as Partial<PersistedState>;
    const merged: PersistedState = { ...DEFAULT_STATE, ...parsed } as PersistedState;
    // Migration: ensure at least one template exists, seeded from legacy fields
    if (!merged.templates || merged.templates.length === 0) {
      const seedId = newId();
      merged.templates = [{
        id: seedId,
        name: "Default",
        subject: merged.subjectA || DEFAULT_STATE.subjectA,
        body: merged.bodyA || DEFAULT_STATE.bodyA,
        html: merged.htmlB || DEFAULT_STATE.htmlB,
      }];
      merged.activeTemplateId = seedId;
    }
    if (!merged.activeTemplateId || !merged.templates.find((t) => t.id === merged.activeTemplateId)) {
      merged.activeTemplateId = merged.templates[0].id;
    }
    return merged;
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

  // AI settings (persisted independently so the API key never leaves localStorage)
  const [ai, setAi] = useState<AISettings>(DEFAULT_AI);
  useEffect(() => {
    if (!hydrated) return;
    try { localStorage.setItem(AI_SETTINGS_KEY, JSON.stringify(ai)); } catch {}
  }, [ai, hydrated]);

  // ---- Draggable "Send current" button state ----
  const [dragUnlocked, setDragUnlocked] = useState(false);
  const [dragPos, setDragPos] = useState<{ x: number; y: number } | null>(() => {
    if (typeof window === "undefined") return null;
    try {
      const raw = localStorage.getItem("midey:sendBtnPos");
      if (!raw) return null;
      const p = JSON.parse(raw);
      if (typeof p?.x === "number" && typeof p?.y === "number") return p;
    } catch {}
    return null;
  });
  useEffect(() => {
    if (!hydrated) return;
    try {
      if (dragPos) localStorage.setItem("midey:sendBtnPos", JSON.stringify(dragPos));
      else localStorage.removeItem("midey:sendBtnPos");
    } catch {}
  }, [dragPos, hydrated]);
  const lastTapRef = useRef(0);
  const tapTimerRef = useRef<number | null>(null);
  const onHeaderTap = useCallback(() => {
    const now = Date.now();
    const elapsed = now - lastTapRef.current;
    if (elapsed < 350 && lastTapRef.current !== 0) {
      // double-tap
      if (tapTimerRef.current) {
        window.clearTimeout(tapTimerRef.current);
        tapTimerRef.current = null;
      }
      lastTapRef.current = 0;
      setDragUnlocked(true);
      toast.info("Send button unlocked — drag it anywhere");
    } else {
      lastTapRef.current = now;
      if (tapTimerRef.current) window.clearTimeout(tapTimerRef.current);
      tapTimerRef.current = window.setTimeout(() => {
        tapTimerRef.current = null;
        lastTapRef.current = 0;
        // single tap
        setDragUnlocked((cur) => {
          if (cur) {
            toast.success("Send button locked in place");
            return false;
          }
          return cur;
        });
      }, 360);
    }
  }, []);

  useEffect(() => {
    setState(loadState());
    const t = (localStorage.getItem(THEME_KEY) as "dark" | "light" | null) ?? "dark";
    setTheme(t);
    setAi(loadAISettings());
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

  /* ---------- Templates / rotation helpers ---------- */
  const activeTemplate: TemplateItem = useMemo(() => {
    const found = state.templates.find((t) => t.id === state.activeTemplateId);
    if (found) return found;
    return state.templates[0] ?? { id: "", name: "Default", subject: state.subjectA, body: state.bodyA, html: state.htmlB };
  }, [state.templates, state.activeTemplateId, state.subjectA, state.bodyA, state.htmlB]);

  const updateTemplate = useCallback((id: string, partial: Partial<TemplateItem>) => {
    setState((s) => ({
      ...s,
      templates: s.templates.map((t) => (t.id === id ? { ...t, ...partial } : t)),
    }));
  }, []);
  const addTemplate = useCallback(() => {
    const id = newId();
    const next: TemplateItem = {
      id,
      name: `Template ${state.templates.length + 1}`,
      subject: activeTemplate.subject,
      body: activeTemplate.body,
      html: activeTemplate.html,
    };
    setState((s) => ({ ...s, templates: [...s.templates, next], activeTemplateId: id }));
    toast.success(`Added "${next.name}"`);
  }, [state.templates.length, activeTemplate]);
  const deleteTemplate = useCallback((id: string) => {
    setState((s) => {
      if (s.templates.length <= 1) {
        toast.error("Keep at least one template");
        return s;
      }
      const remaining = s.templates.filter((t) => t.id !== id);
      return {
        ...s,
        templates: remaining,
        activeTemplateId: s.activeTemplateId === id ? remaining[0].id : s.activeTemplateId,
      };
    });
  }, []);

  /** Resolve the (subject, body, html) actually used for the NEXT send,
   *  honoring rotation toggles. */
  const rotation = useMemo(() => {
    const n = state.templates.length || 1;
    const idx = ((state.sendCounter % n) + n) % n;
    const rotTpl = state.templates[idx] ?? activeTemplate;
    return {
      subject: state.rotateSubjects ? rotTpl.subject : activeTemplate.subject,
      body: state.rotateBodies ? rotTpl.body : activeTemplate.body,
      html: state.rotateBodies ? rotTpl.html : activeTemplate.html,
      rotIndex: idx,
      rotName: rotTpl.name,
    };
  }, [state.templates, state.sendCounter, state.rotateSubjects, state.rotateBodies, activeTemplate]);

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
    setState((s) => ({
      ...s,
      rowStates: { ...s.rowStates, [rowIndex]: "processed" },
      sendCounter: s.sendCounter + 1,
    }));
    sendLogRef.current.push(Date.now());
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
    () => autoFormatHtml(renderTemplate(activeTemplate.html, sampleRow, undefined, state.sendCounter)),
    [activeTemplate.html, sampleRow, state.sendCounter],
  );
  const renderedSubjectB = useMemo(
    () => renderTemplate(activeTemplate.subject, sampleRow, undefined, state.sendCounter),
    [activeTemplate.subject, sampleRow, state.sendCounter],
  );

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
    () => renderTemplate(activeTemplate.subject, sampleRow, undefined, state.sendCounter),
    [activeTemplate.subject, sampleRow, state.sendCounter],
  );
  const executePlainTest = useCallback(() => {
    const recipients = cleanEmails(state.recipientB);
    if (!recipients) { toast.error("Recipient required"); return; }
    const subject = renderTemplate(activeTemplate.subject, sampleRow, undefined, state.sendCounter);
    const body = renderTemplate(activeTemplate.body, sampleRow, undefined, state.sendCounter);
    toast.success("Opening test draft…");
    window.location.href = buildMailto(state.recipientB, { subject, body });
  }, [state.recipientB, activeTemplate, sampleRow]);

  /* ---------- Session stats / autosave / resume ---------- */
  const sessionStartRef = useRef<number>(Date.now());
  const sendLogRef = useRef<number[]>([]);
  const [, forceTick] = useState(0);
  useEffect(() => {
    const id = window.setInterval(() => forceTick((n) => n + 1), 1000);
    return () => window.clearInterval(id);
  }, []);

  // 10s autosave heartbeat
  useEffect(() => {
    if (!hydrated) return;
    const id = window.setInterval(() => {
      try {
        const firstPending = (() => {
          for (let i = 0; i < state.rows.length; i++) {
            if (!state.rowStates[i]) return i;
          }
          return state.rows.length;
        })();
        localStorage.setItem(AUTOSAVE_KEY, JSON.stringify(state));
        localStorage.setItem(
          SESSION_META_KEY,
          JSON.stringify({ ts: Date.now(), lastRow: firstPending, total: state.rows.length }),
        );
      } catch {}
    }, 10_000);
    return () => window.clearInterval(id);
  }, [state, hydrated]);

  const [resume, setResume] = useState<{ lastRow: number; ts: number } | null>(null);
  useEffect(() => {
    if (!hydrated) return;
    try {
      const raw = localStorage.getItem(SESSION_META_KEY);
      if (!raw) return;
      const meta = JSON.parse(raw);
      if (typeof meta?.lastRow === "number" && meta.lastRow > 0 && state.rows.length > 0) {
        setResume({ lastRow: meta.lastRow, ts: meta.ts ?? 0 });
      }
    } catch {}
  }, [hydrated]); // eslint-disable-line react-hooks/exhaustive-deps

  const [resumeTarget, setResumeTarget] = useState<number | null>(null);
  const acceptResume = useCallback(() => {
    if (resume) setResumeTarget(resume.lastRow);
    setResume(null);
  }, [resume]);

  // Velocity (last 30 min)
  const velocity30 = useMemo(() => {
    const cutoff = Date.now() - 30 * 60_000;
    sendLogRef.current = sendLogRef.current.filter((t) => t > cutoff - 1);
    return sendLogRef.current.length;
  }, [state.sendCounter]); // eslint-disable-line react-hooks/exhaustive-deps
  const sessionSeconds = Math.floor((Date.now() - sessionStartRef.current) / 1000);

  /* ---------- UI ---------- */

  return (
    <div className="min-h-screen bg-bg-app text-foreground">
      <Header
        theme={theme}
        onToggleTheme={() => setTheme((t) => (t === "dark" ? "light" : "dark"))}
        onClearAll={clearAll}
        totalRows={state.rows.length}
        processedRows={processedCount}
        onHeaderTap={onHeaderTap}
        dragUnlocked={dragUnlocked}
      />

      <main className="mx-auto max-w-5xl px-3 pb-24 pt-4 sm:px-6">
        <DragContext.Provider value={{ dragUnlocked, dragPos, setDragPos }}>
        <SessionStats
          processedCount={processedCount}
          totalRows={state.rows.length}
          dailyGoal={state.dailyGoal}
          onDailyGoal={(n: number) => patch({ dailyGoal: n })}
          velocity30={velocity30}
          sessionSeconds={sessionSeconds}
        />
        {resume && (
          <ResumeBanner
            lastRow={resume.lastRow}
            ts={resume.ts}
            onRestore={acceptResume}
            onDismiss={() => setResume(null)}
          />
        )}
        <TemplateControlPanel
          templates={state.templates}
          activeTemplateId={state.activeTemplateId}
          onSelect={(id: string) => patch({ activeTemplateId: id })}
          onUpdate={updateTemplate}
          onAdd={addTemplate}
          onDelete={deleteTemplate}
        />
        <AIPersonalizationPanel ai={ai} onChange={setAi} />
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
            activeTemplate={activeTemplate}
            updateTemplate={updateTemplate}
            rotation={rotation}
            resumeTarget={resumeTarget}
            onConsumeResume={() => setResumeTarget(null)}
            ai={ai}
          />
        </div>
        </DragContext.Provider>
      </main>
    </div>
  );
}

/* --------------------- Draggable Send button shell --------------------- */

type DragCtx = {
  dragUnlocked: boolean;
  dragPos: { x: number; y: number } | null;
  setDragPos: (p: { x: number; y: number } | null) => void;
};
const DragContext = createContext<DragCtx>({
  dragUnlocked: false,
  dragPos: null,
  setDragPos: () => {},
});

function DraggableSendShell({ children }: { children: React.ReactNode }) {
  const { dragUnlocked, dragPos, setDragPos } = useContext(DragContext);
  const ref = useRef<HTMLDivElement>(null);
  const dragRef = useRef<{
    active: boolean;
    offX: number;
    offY: number;
    moved: boolean;
  }>({ active: false, offX: 0, offY: 0, moved: false });

  const isFloating = dragUnlocked || dragPos !== null;

  const beginDrag = (clientX: number, clientY: number) => {
    if (!dragUnlocked) return;
    const el = ref.current;
    if (!el) return;
    const r = el.getBoundingClientRect();
    dragRef.current = {
      active: true,
      offX: clientX - r.left,
      offY: clientY - r.top,
      moved: false,
    };
    // Seed a fixed position if we don't have one yet
    if (!dragPos) setDragPos({ x: r.left, y: r.top });
  };

  const moveDrag = (clientX: number, clientY: number) => {
    const d = dragRef.current;
    if (!d.active) return;
    d.moved = true;
    const el = ref.current;
    const w = el?.offsetWidth ?? 0;
    const h = el?.offsetHeight ?? 0;
    const vw = window.innerWidth;
    const vh = window.innerHeight;
    const x = Math.max(4, Math.min(vw - w - 4, clientX - d.offX));
    const y = Math.max(4, Math.min(vh - h - 4, clientY - d.offY));
    setDragPos({ x, y });
  };

  const endDrag = () => {
    dragRef.current.active = false;
  };

  useEffect(() => {
    if (!dragUnlocked) return;
    const onMM = (e: MouseEvent) => moveDrag(e.clientX, e.clientY);
    const onMU = () => endDrag();
    const onTM = (e: TouchEvent) => {
      if (!dragRef.current.active) return;
      const t = e.touches[0];
      if (!t) return;
      e.preventDefault();
      moveDrag(t.clientX, t.clientY);
    };
    const onTE = () => endDrag();
    window.addEventListener("mousemove", onMM);
    window.addEventListener("mouseup", onMU);
    window.addEventListener("touchmove", onTM, { passive: false });
    window.addEventListener("touchend", onTE);
    window.addEventListener("touchcancel", onTE);
    return () => {
      window.removeEventListener("mousemove", onMM);
      window.removeEventListener("mouseup", onMU);
      window.removeEventListener("touchmove", onTM);
      window.removeEventListener("touchend", onTE);
      window.removeEventListener("touchcancel", onTE);
    };
  }, [dragUnlocked]); // eslint-disable-line react-hooks/exhaustive-deps

  const style: React.CSSProperties = isFloating && dragPos
    ? {
        position: "fixed",
        left: dragPos.x,
        top: dragPos.y,
        zIndex: 60,
        touchAction: "none",
      }
    : { touchAction: "auto" };

  return (
    <div
      ref={ref}
      style={style}
      onMouseDown={(e) => {
        if (!dragUnlocked) return;
        beginDrag(e.clientX, e.clientY);
      }}
      onTouchStart={(e) => {
        if (!dragUnlocked) return;
        const t = e.touches[0];
        if (!t) return;
        beginDrag(t.clientX, t.clientY);
      }}
      className={
        isFloating
          ? `rounded-md ${
              dragUnlocked
                ? "animate-pulse ring-2 ring-[var(--amber)] shadow-[0_0_24px_rgba(245,158,11,0.55)]"
                : "ring-1 ring-[var(--amber)]/40"
            } bg-bg-app/90 p-2 backdrop-blur`
          : ""
      }
    >
      {children}
    </div>
  );
}

/* ----------------------------- Header ----------------------------- */

function Header({
  theme, onToggleTheme, onClearAll, totalRows, processedRows, onHeaderTap, dragUnlocked,
}: {
  theme: "dark" | "light";
  onToggleTheme: () => void;
  onClearAll: () => void;
  totalRows: number;
  processedRows: number;
  onHeaderTap: () => void;
  dragUnlocked: boolean;
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
            <h1
              onClick={onHeaderTap}
              title="Double-tap to unlock the Send button · single tap to lock"
              className={`cursor-pointer select-none truncate text-sm font-semibold leading-tight sm:text-base ${
                dragUnlocked ? "text-amber-glow" : ""
              }`}
            >
              Wayne Enterprises <span className="text-sky-glow">Outreach Lab</span>
              {dragUnlocked && (
                <span className="ml-2 align-middle rounded border border-[var(--amber)]/60 bg-[var(--amber)]/15 px-1.5 py-0.5 font-mono-data text-[9px] uppercase tracking-wider text-[var(--amber)]">
                  drag unlocked
                </span>
              )}
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
          accept=".csv,.xlsx,text/csv,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
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
          <Upload /> {parsing ? `Parsing… ${progress}%` : totalRows ? "Replace file" : "Upload CSV / XLSX"}
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
  activeTemplate, updateTemplate, rotation, resumeTarget, onConsumeResume, ai,
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
  activeTemplate: TemplateItem;
  updateTemplate: (id: string, p: Partial<TemplateItem>) => void;
  rotation: { subject: string; body: string; html: string; rotIndex: number; rotName: string };
  resumeTarget: number | null;
  onConsumeResume: () => void;
  ai: AISettings;
}) {
  const firstPendingIndex = queue.find(
    (i) => (state.rowStates[i] ?? "pending") === "pending",
  );
  // Manual override — "Jump to row" input or "Resend" button on a processed row.
  const [activeOverride, setActiveOverride] = useState<number | null>(null);
  const [jumpInput, setJumpInput] = useState<string>("");
  // Apply external "Restore previous session" jump
  useEffect(() => {
    if (resumeTarget !== null) {
      setActiveOverride(resumeTarget);
      setJumpInput(String(resumeTarget));
      onConsumeResume();
    }
  }, [resumeTarget, onConsumeResume]);
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
  const previewSubject = renderTemplate(rotation.subject, previewRow, undefined, state.sendCounter);
  const previewBody = renderTemplate(rotation.body, previewRow, undefined, state.sendCounter);
  const mailtoLen = previewRow && !state.htmlMode
    ? `mailto:${previewTo}?subject=${encodeURIComponent(previewSubject)}&body=${encodeURIComponent(previewBody)}`.length
    : 0;
  const overLimit = mailtoLen > 2000;

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

      {/* Active Template Dropdown + rotation toggles */}
      <div className="space-y-2 rounded-lg border border-border-strong/60 bg-surface-2 p-3">
        <div className="flex flex-wrap items-center gap-2">
          <Label className="font-mono-data text-[10px] uppercase tracking-wider text-muted-foreground">
            Active template
          </Label>
          <select
            value={state.activeTemplateId}
            onChange={(e) => patch({ activeTemplateId: e.target.value })}
            className="h-8 flex-1 min-w-[160px] rounded-md border border-border-strong/70 bg-bg-app px-2 font-mono-data text-xs outline-none focus:glow-sky"
          >
            {state.templates.map((t) => (
              <option key={t.id} value={t.id}>{t.name}</option>
            ))}
          </select>
          <span className="font-mono-data text-[10px] text-muted-foreground">
            {state.templates.length} total
          </span>
        </div>
        <div className="grid gap-2 sm:grid-cols-2">
          <label className="flex items-center justify-between gap-2 rounded border border-border-strong/40 bg-bg-app px-2 py-1.5">
            <span className="flex items-center gap-1.5 font-mono-data text-[11px] text-foreground">
              <Shuffle className="size-3 text-sky-glow" /> Rotate Subjects
            </span>
            <input
              type="checkbox"
              checked={state.rotateSubjects}
              onChange={(e) => patch({ rotateSubjects: e.target.checked })}
              className="size-4 accent-[var(--sky)]"
            />
          </label>
          <label className="flex items-center justify-between gap-2 rounded border border-border-strong/40 bg-bg-app px-2 py-1.5">
            <span className="flex items-center gap-1.5 font-mono-data text-[11px] text-foreground">
              <Shuffle className="size-3 text-amber-glow" /> Rotate Body
            </span>
            <input
              type="checkbox"
              checked={state.rotateBodies}
              onChange={(e) => patch({ rotateBodies: e.target.checked })}
              className="size-4 accent-[var(--amber)]"
            />
          </label>
        </div>
        {(state.rotateSubjects || state.rotateBodies) && state.templates.length > 1 && (
          <p className="font-mono-data text-[10px] text-muted-foreground">
            Next send → <span className="text-sky-glow">{rotation.rotName}</span> (slot #{rotation.rotIndex + 1}/{state.templates.length})
          </p>
        )}
      </div>

      <SpamHealthCheck
        subject={activeTemplate.subject}
        body={state.htmlMode ? activeTemplate.html : activeTemplate.body}
        html={state.htmlMode ? activeTemplate.html : ""}
      />

      <div className="grid gap-3 sm:grid-cols-2">
        <Field label="Subject template">
          <Input
            value={activeTemplate.subject}
            onChange={(e) => updateTemplate(activeTemplate.id, { subject: e.target.value })}
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
              value={activeTemplate.html}
              onChange={(v) => updateTemplate(activeTemplate.id, { html: v })}
            />
            <Textarea
              ref={htmlTextareaRef}
              value={activeTemplate.html}
              onChange={(e) => updateTemplate(activeTemplate.id, { html: e.target.value })}
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
                srcDoc={`<!doctype html><html><body style="margin:0;padding:12px;font-family:system-ui">${autoFormatHtml(renderTemplate(activeTemplate.html, previewRow ?? sampleRow, undefined, state.sendCounter))}</body></html>`}
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
              value={activeTemplate.body}
              onChange={(e) => updateTemplate(activeTemplate.id, { body: e.target.value })}
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
            subjectTpl={rotation.subject}
            bodyTpl={rotation.body}
            htmlMode={state.htmlMode}
            htmlTpl={rotation.html}
            spinSeed={state.sendCounter}
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
            ai={ai}
          />
        )}
      </div>
    </div>
  );
}

function NextRowPreview({
  rowIndex, row, targetEmailHeader, subjectTpl, bodyTpl, htmlMode, htmlTpl, onSend, onSkip, isResend, ai, spinSeed,
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
  ai: AISettings;
  spinSeed: number;
}) {
  // ---- AI {ai_insight} resolver: fetch when active row exposes a description ----
  const description = ((row?.[ai.descriptionColumn] ?? "") as string).trim();
  const usesAi =
    /\{ai_insight\}/.test(subjectTpl) ||
    /\{ai_insight\}/.test(bodyTpl) ||
    /\{ai_insight\}/.test(htmlTpl);
  const [aiInsight, setAiInsight] = useState<string>("");
  const [aiLoading, setAiLoading] = useState(false);
  useEffect(() => {
    if (!usesAi) { setAiInsight(""); setAiLoading(false); return; }
    if (!ai.enabled || !ai.apiKey || !description) {
      setAiInsight(ai.fallback);
      setAiLoading(false);
      return;
    }
    const ctrl = new AbortController();
    setAiLoading(true);
    setAiInsight("");
    generateAIInsight(ai, description, ctrl.signal)
      .then((txt) => setAiInsight(txt || ai.fallback))
      .catch(() => setAiInsight(ai.fallback))
      .finally(() => setAiLoading(false));
    return () => ctrl.abort();
  }, [rowIndex, description, usesAi, ai.enabled, ai.apiKey, ai.provider, ai.prompt, ai.fallback]);

  const extras = useMemo(
    () => ({ ai_insight: aiLoading ? "…" : (aiInsight || ai.fallback) }),
    [aiInsight, aiLoading, ai.fallback],
  );

  const rawRecipients = row?.[targetEmailHeader] ?? "";
  const toAddr = cleanEmails(rawRecipients);
  const subject = renderTemplate(subjectTpl, row, extras, spinSeed);
  const body = renderTemplate(bodyTpl, row, extras, spinSeed);
  const renderedHtml = autoFormatHtml(renderTemplate(htmlTpl, row, extras, spinSeed));
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
        {usesAi && (
          <div className="font-mono-data text-[10px] uppercase tracking-wider">
            {aiLoading ? (
              <span className="inline-flex items-center gap-1.5 text-sky-glow">
                <span className="size-1.5 animate-pulse rounded-full bg-sky-glow" />
                Generating personal insight…
              </span>
            ) : (
              <span className="text-muted-foreground">
                AI insight ·{" "}
                <span className="text-sky-glow normal-case tracking-normal">
                  {aiInsight || ai.fallback}
                </span>
              </span>
            )}
          </div>
        )}
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
      <DraggableSendShell>
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
      </DraggableSendShell>
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

/* ===================== Session Stats Strip ===================== */

function fmtDuration(sec: number) {
  const h = Math.floor(sec / 3600);
  const m = Math.floor((sec % 3600) / 60);
  const s = sec % 60;
  if (h > 0) return `${h}h ${m}m ${s}s`;
  if (m > 0) return `${m}m ${s}s`;
  return `${s}s`;
}

function SessionStats({
  processedCount, totalRows, dailyGoal, onDailyGoal, velocity30, sessionSeconds,
}: {
  processedCount: number;
  totalRows: number;
  dailyGoal: number;
  onDailyGoal: (n: number) => void;
  velocity30: number;
  sessionSeconds: number;
}) {
  const [open, setOpen] = useState(true);
  const goal = Math.max(1, dailyGoal || 1);
  const pct = Math.min(100, Math.round((processedCount / goal) * 100));
  return (
    <section className="mb-3 rounded-xl border border-border-strong/70 bg-surface-1 p-3">
      <button
        type="button"
        onClick={() => setOpen((o) => !o)}
        className="flex w-full items-center justify-between gap-2"
      >
        <span className="flex items-center gap-2 font-mono-data text-[11px] uppercase tracking-wider text-muted-foreground">
          <Activity className="size-3.5 text-sky-glow" /> Session stats
        </span>
        <span className="flex items-center gap-3 font-mono-data text-[11px] text-muted-foreground">
          <span><span className="text-foreground">{processedCount}</span> / {goal}</span>
          <span>{velocity30}/30m</span>
          <span>{fmtDuration(sessionSeconds)}</span>
          {open ? <ChevronUp className="size-3.5" /> : <ChevronDown className="size-3.5" />}
        </span>
      </button>
      {open && (
        <div className="mt-3 space-y-3">
          <div>
            <div className="mb-1 flex items-center justify-between font-mono-data text-[10px] uppercase tracking-wider text-muted-foreground">
              <span>Daily progress</span>
              <span>{pct}%</span>
            </div>
            <div className="h-2 w-full overflow-hidden rounded-full bg-surface-2">
              <div
                className="h-full bg-[var(--sky)] transition-[width] duration-300"
                style={{ width: `${pct}%` }}
              />
            </div>
          </div>
          <div className="grid gap-2 sm:grid-cols-3">
            <div className="rounded border border-border-strong/40 bg-surface-2 p-2">
              <div className="font-mono-data text-[10px] uppercase text-muted-foreground">Targets</div>
              <div className="font-mono-data text-sm text-foreground">
                {processedCount.toLocaleString()} / {goal.toLocaleString()}
              </div>
              <Input
                type="number"
                min={1}
                value={dailyGoal}
                onChange={(e) => onDailyGoal(Math.max(1, Number(e.target.value) || 1))}
                className="mt-1 h-7 font-mono-data text-xs"
              />
            </div>
            <div className="rounded border border-border-strong/40 bg-surface-2 p-2">
              <div className="font-mono-data text-[10px] uppercase text-muted-foreground">Velocity (30m)</div>
              <div className="font-mono-data text-sm text-amber-glow">{velocity30} sent</div>
              <div className="font-mono-data text-[10px] text-muted-foreground">
                {totalRows ? `${totalRows.toLocaleString()} rows loaded` : "no file"}
              </div>
            </div>
            <div className="rounded border border-border-strong/40 bg-surface-2 p-2">
              <div className="font-mono-data text-[10px] uppercase text-muted-foreground">Session timer</div>
              <div className="font-mono-data text-sm text-sky-glow">{fmtDuration(sessionSeconds)}</div>
              <div className="font-mono-data text-[10px] text-muted-foreground">since page load</div>
            </div>
          </div>
        </div>
      )}
    </section>
  );
}

/* ===================== Resume banner ===================== */

function ResumeBanner({
  lastRow, ts, onRestore, onDismiss,
}: { lastRow: number; ts: number; onRestore: () => void; onDismiss: () => void }) {
  const ago = ts ? Math.max(0, Math.floor((Date.now() - ts) / 60000)) : 0;
  return (
    <div className="mb-3 flex flex-wrap items-center justify-between gap-2 rounded-xl border border-sky-glow/40 bg-sky-glow/5 px-3 py-2">
      <div className="flex items-center gap-2 font-mono-data text-xs text-foreground">
        <History className="size-3.5 text-sky-glow" />
        Resume your previous session at Row <span className="text-sky-glow">#{lastRow}</span>?
        {ago > 0 && <span className="text-muted-foreground">· {ago}m ago</span>}
      </div>
      <div className="flex gap-2">
        <Button size="sm" variant="ghost" onClick={onDismiss}>Dismiss</Button>
        <Button size="sm" onClick={onRestore} className="glow-sky bg-[var(--sky)] text-black hover:bg-[var(--sky)]/90">
          Restore
        </Button>
      </div>
    </div>
  );
}

/* ===================== Template Control Panel ===================== */

function TemplateControlPanel({
  templates, activeTemplateId, onSelect, onUpdate, onAdd, onDelete,
}: {
  templates: TemplateItem[];
  activeTemplateId: string;
  onSelect: (id: string) => void;
  onUpdate: (id: string, p: Partial<TemplateItem>) => void;
  onAdd: () => void;
  onDelete: (id: string) => void;
}) {
  const [open, setOpen] = useState(false);
  return (
    <section className="mb-3 rounded-xl border border-border-strong/70 bg-surface-1">
      <button
        type="button"
        onClick={() => setOpen((o) => !o)}
        className="flex w-full items-center justify-between gap-2 p-3"
      >
        <span className="flex items-center gap-2 font-mono-data text-[11px] uppercase tracking-wider text-muted-foreground">
          <Layers className="size-3.5 text-amber-glow" /> Template Control Panel
        </span>
        <span className="flex items-center gap-2 font-mono-data text-[11px] text-muted-foreground">
          {templates.length} template{templates.length === 1 ? "" : "s"}
          {open ? <ChevronUp className="size-3.5" /> : <ChevronDown className="size-3.5" />}
        </span>
      </button>
      {open && (
        <div className="space-y-3 border-t border-border-strong/40 p-3">
          <div className="flex items-center justify-between gap-2">
            <p className="font-mono-data text-[10px] text-muted-foreground">
              Edit every template's Subject &amp; Body. The dropdown below feeds the live preview.
            </p>
            <Button size="sm" onClick={onAdd} className="h-8">
              <Plus className="size-3.5" /> Add Template
            </Button>
          </div>
          <div className="space-y-3">
            {templates.map((t, idx) => {
              const isActive = t.id === activeTemplateId;
              return (
                <div
                  key={t.id}
                  className={`rounded-lg border p-3 ${
                    isActive ? "border-sky-glow/60 bg-sky-glow/5" : "border-border-strong/40 bg-surface-2"
                  }`}
                >
                  <div className="mb-2 flex flex-wrap items-center gap-2">
                    <span className="font-mono-data text-[10px] uppercase text-muted-foreground">#{idx + 1}</span>
                    <Input
                      value={t.name}
                      onChange={(e) => onUpdate(t.id, { name: e.target.value })}
                      className="h-7 max-w-[200px] font-mono-data text-xs"
                      placeholder="Template name"
                    />
                    <Button
                      size="sm"
                      variant="ghost"
                      className="h-7"
                      onClick={() => onSelect(t.id)}
                      disabled={isActive}
                    >
                      {isActive ? "Active" : "Activate"}
                    </Button>
                    <div className="ml-auto">
                      <Button
                        size="sm"
                        variant="ghost"
                        className="h-7 text-destructive hover:bg-destructive/10"
                        onClick={() => onDelete(t.id)}
                        aria-label={`Delete ${t.name}`}
                      >
                        <Trash2 className="size-3.5" />
                      </Button>
                    </div>
                  </div>
                  <div className="grid gap-2">
                    <Input
                      value={t.subject}
                      onChange={(e) => onUpdate(t.id, { subject: e.target.value })}
                      placeholder="Subject"
                      className="h-8 font-mono-data text-xs"
                    />
                    <Textarea
                      value={t.body}
                      onChange={(e) => onUpdate(t.id, { body: e.target.value })}
                      placeholder="Plain-text body"
                      rows={3}
                      className="font-mono-data text-[12px]"
                    />
                    <Textarea
                      value={t.html}
                      onChange={(e) => onUpdate(t.id, { html: e.target.value })}
                      placeholder="HTML body"
                      rows={3}
                      className="font-mono-data text-[12px]"
                      spellCheck={false}
                    />
                  </div>
                </div>
              );
            })}
          </div>
        </div>
      )}
    </section>
  );
}

/* ===================== Spam / Link Health Check ===================== */

function SpamHealthCheck({
  subject, body, html,
}: { subject: string; body: string; html: string }) {
  const subj = useMemo(() => scanSpam(subject), [subject]);
  const bod = useMemo(() => scanSpam(body), [body]);
  const links = useMemo(() => (html ? scanLinks(html) : { ok: 0, broken: [] as { tag: string; reason: string }[] }), [html]);
  const totalHits = subj.hits.length + bod.hits.length;
  const hot = totalHits > 0 || subj.exclaim > 2 || bod.exclaim > 4 || links.broken.length > 0;
  return (
    <div
      className={`rounded-lg border p-3 ${
        hot ? "border-amber-glow/60 bg-amber-glow/10" : "border-border-strong/60 bg-surface-2"
      }`}
    >
      <div className="mb-2 flex items-center gap-2 font-mono-data text-[10px] uppercase tracking-wider text-muted-foreground">
        <ShieldAlert className={`size-3.5 ${hot ? "text-amber-glow" : "text-sky-glow"}`} />
        Health check
        <span className={`ml-auto font-mono-data text-[10px] ${hot ? "text-amber-glow" : "text-sky-glow"}`}>
          {hot ? "review" : "clean"}
        </span>
      </div>
      <div className="grid gap-2 sm:grid-cols-2">
        <div className="space-y-1">
          <div className="font-mono-data text-[10px] uppercase text-muted-foreground">Subject</div>
          {subj.hits.length === 0 && subj.exclaim <= 2 && subj.allCaps === 0 ? (
            <p className="font-mono-data text-[11px] text-muted-foreground">No spam triggers detected.</p>
          ) : (
            <div className="flex flex-wrap gap-1">
              {subj.hits.map((h) => (
                <span key={h} className="rounded border border-amber-glow/40 bg-amber-glow/15 px-1.5 py-0.5 font-mono-data text-[10px] text-amber-glow">{h}</span>
              ))}
              {subj.exclaim > 2 && (
                <span className="rounded border border-amber-glow/40 bg-amber-glow/15 px-1.5 py-0.5 font-mono-data text-[10px] text-amber-glow">{subj.exclaim}×!</span>
              )}
              {subj.allCaps > 0 && (
                <span className="rounded border border-amber-glow/40 bg-amber-glow/15 px-1.5 py-0.5 font-mono-data text-[10px] text-amber-glow">{subj.allCaps} CAPS</span>
              )}
            </div>
          )}
        </div>
        <div className="space-y-1">
          <div className="font-mono-data text-[10px] uppercase text-muted-foreground">Body</div>
          {bod.hits.length === 0 && bod.exclaim <= 4 && bod.allCaps === 0 ? (
            <p className="font-mono-data text-[11px] text-muted-foreground">No spam triggers detected.</p>
          ) : (
            <div className="flex flex-wrap gap-1">
              {bod.hits.map((h) => (
                <span key={h} className="rounded border border-amber-glow/40 bg-amber-glow/15 px-1.5 py-0.5 font-mono-data text-[10px] text-amber-glow">{h}</span>
              ))}
              {bod.exclaim > 4 && (
                <span className="rounded border border-amber-glow/40 bg-amber-glow/15 px-1.5 py-0.5 font-mono-data text-[10px] text-amber-glow">{bod.exclaim}×!</span>
              )}
              {bod.allCaps > 0 && (
                <span className="rounded border border-amber-glow/40 bg-amber-glow/15 px-1.5 py-0.5 font-mono-data text-[10px] text-amber-glow">{bod.allCaps} CAPS</span>
              )}
            </div>
          )}
        </div>
      </div>
      {html && (
        <div className="mt-2 space-y-1">
          <div className="font-mono-data text-[10px] uppercase text-muted-foreground">
            Links · <span className="text-sky-glow">{links.ok} ok</span>
            {links.broken.length > 0 && <> · <span className="text-amber-glow">{links.broken.length} broken</span></>}
          </div>
          {links.broken.length > 0 && (
            <ul className="space-y-0.5">
              {links.broken.slice(0, 4).map((b, i) => (
                <li key={i} className="font-mono-data text-[10px] text-amber-glow">
                  ⚠ {b.reason}: <span className="text-muted-foreground">{b.tag.slice(0, 60)}</span>
                </li>
              ))}
            </ul>
          )}
        </div>
      )}
    </div>
  );
}

/* ===================== AI Personalization Panel ===================== */

function AIPersonalizationPanel({
  ai, onChange,
}: { ai: AISettings; onChange: (next: AISettings) => void }) {
  const [open, setOpen] = useState(false);
  const [reveal, setReveal] = useState(false);
  const set = <K extends keyof AISettings>(k: K, v: AISettings[K]) =>
    onChange({ ...ai, [k]: v });
  return (
    <section className="mb-3 rounded-xl border border-border-strong/70 bg-surface-1">
      <button
        type="button"
        onClick={() => setOpen((o) => !o)}
        className="flex w-full items-center justify-between gap-2 p-3"
      >
        <span className="flex items-center gap-2 font-mono-data text-[11px] uppercase tracking-wider text-muted-foreground">
          <Zap className="size-3.5 text-sky-glow" />
          AI Personalization · {"{ai_insight}"}
          <span
            className={`ml-1 rounded px-1.5 py-0.5 text-[10px] ${
              ai.enabled && ai.apiKey
                ? "border border-sky-glow/40 bg-sky-glow/10 text-sky-glow"
                : "border border-border-strong/50 bg-bg-app text-muted-foreground"
            }`}
          >
            {ai.enabled && ai.apiKey ? "live" : "off"}
          </span>
        </span>
        <span className="flex items-center gap-2 font-mono-data text-[11px] text-muted-foreground">
          {ai.provider}
          {open ? <ChevronUp className="size-3.5" /> : <ChevronDown className="size-3.5" />}
        </span>
      </button>
      {open && (
        <div className="space-y-3 border-t border-border-strong/40 p-3">
          <div className="grid gap-2 sm:grid-cols-[auto_1fr_auto] sm:items-center">
            <label className="flex items-center gap-2 font-mono-data text-[11px] text-foreground">
              <input
                type="checkbox"
                checked={ai.enabled}
                onChange={(e) => set("enabled", e.target.checked)}
                className="size-4 accent-[var(--sky)]"
              />
              Enable
            </label>
            <select
              value={ai.provider}
              onChange={(e) => set("provider", e.target.value as "gemini" | "openai")}
              className="h-8 rounded-md border border-border-strong/70 bg-bg-app px-2 font-mono-data text-xs outline-none focus:glow-sky"
            >
              <option value="gemini">Google Gemini</option>
              <option value="openai">OpenAI</option>
            </select>
            <span className="font-mono-data text-[10px] text-muted-foreground">
              Key saved locally
            </span>
          </div>
          <Field label={`${ai.provider === "gemini" ? "Gemini" : "OpenAI"} API key`}>
            <div className="flex gap-2">
              <Input
                type={reveal ? "text" : "password"}
                value={ai.apiKey}
                onChange={(e) => set("apiKey", e.target.value)}
                placeholder={ai.provider === "gemini" ? "AIza…" : "sk-…"}
                className="font-mono-data text-xs"
                autoComplete="off"
                spellCheck={false}
              />
              <Button
                size="sm"
                variant="ghost"
                type="button"
                onClick={() => setReveal((r) => !r)}
                className="h-9"
              >
                {reveal ? "Hide" : "Show"}
              </Button>
            </div>
          </Field>
          <div className="grid gap-2 sm:grid-cols-2">
            <Field label="Store description column">
              <Input
                value={ai.descriptionColumn}
                onChange={(e) => set("descriptionColumn", e.target.value)}
                placeholder="store_description"
                className="font-mono-data text-xs"
              />
            </Field>
            <Field label="Fallback phrase">
              <Input
                value={ai.fallback}
                onChange={(e) => set("fallback", e.target.value)}
                placeholder="your unique collection"
                className="font-mono-data text-xs"
              />
            </Field>
          </div>
          <Field label="AI Generation Prompt Blueprint">
            <Textarea
              value={ai.prompt}
              onChange={(e) => set("prompt", e.target.value)}
              rows={4}
              className="font-mono-data text-[12px] leading-relaxed"
              placeholder="Write a brief 7-word phrase about a category they specialize in…"
            />
          </Field>
          <p className="font-mono-data text-[10px] text-muted-foreground">
            Insert <span className="text-sky-glow">{"{ai_insight}"}</span> anywhere in a Subject or Body template.
            When the active row exposes <span className="text-sky-glow">{ai.descriptionColumn}</span>, a custom
            sentence is fetched in the background and injected. Missing data → fallback phrase.
          </p>
        </div>
      )}
    </section>
  );
}
