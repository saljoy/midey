import { createFileRoute } from "@tanstack/react-router";
import { createContext, useCallback, useContext, useEffect, useMemo, useRef, useState } from "react";
import Papa from "papaparse";
import * as XLSX from "xlsx";
import { toast } from "sonner";
import {
  Upload, Sun, Moon, Trash2, Send, Mail, Code2, Copy,
  Zap, FileText, Eye, SkipForward, Save, AlertTriangle,
  Bold, Italic, Underline, Strikethrough, List, ListOrdered,
  AlignLeft, AlignCenter, AlignRight, AlignJustify, Link2, Minus,
  Palette, Highlighter, CornerDownLeft, RotateCcw,
  Menu, X, Settings, Key, CheckCircle, XCircle, AlertCircle,
  Globe, Beaker, ChevronRight, Search, BookOpen, PenLine,
  Sparkles, RefreshCw, Clipboard, Download, ExternalLink, EyeOff,
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
      { title: "Midey Enterprises Outreach Lab" },
      { name: "description", content: "High-performance mobile outreach console: parse CSV lead lists, fire personalized mailto handoffs, and craft rich HTML drafts — all 100% client-side." },
    ],
  }),
  component: Index,
});

/* ============================================================
   TYPES
   ============================================================ */

type Row = Record<string, string>;
type RowState = "pending" | "processed" | "skipped";
type HomepageMode = "research" | "queue";

export interface TemplateItem {
  id: string;
  name: string;
  subject: string;
  body: string;
  html: string;
}

export interface ApiKey {
  id: string;
  label: string;
  value: string;
  enabled: boolean;
  status: "unknown" | "active" | "quota" | "invalid" | "testing";
}

interface PersistedState {
  headers: string[];
  rows: Row[];
  rowStates: Record<number, RowState>;
  targetEmailHeader: string;
  domainHeader: string;
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
  homepageMode: HomepageMode;
  // Row indices (within `rows`, keyed by domainHeader) marked as "done" in
  // Research Mode — auto-set once an email is generated for that lead, and
  // manually toggleable. Persists across mode switches and page reloads
  // via the same localStorage-backed state as everything else.
  researchDone: Record<number, boolean>;
  // Leads whose email ends with this domain (e.g. "gmail.com") jump to the
  // front of the send queue, so they go out first. Empty = no reordering.
  priorityDomain: string;
  // Auto-detected column holding each lead's country, plus the currently
  // selected country to prioritize (empty = no country filter).
  countryHeader: string;
  priorityCountry: string;
  // When BOTH priorityDomain and priorityCountry are set: "all" requires a
  // lead to match both to be prioritized, "any" lets either one qualify.
  priorityMatchMode: "all" | "any";
}

/* ============================================================
   STORAGE KEYS
   ============================================================ */

const STORAGE_KEY = "midey.outreach.v1";
const THEME_KEY = "midey.theme";
const AUTOSAVE_KEY = "midey.outreach.autosave.v1";
const SESSION_META_KEY = "midey.outreach.session.v1";
const API_KEYS_KEY = "midey.gemini.keys.v1";
const PROMPTS_KEY = "midey.prompts.v1";

/* ============================================================
   DEFAULT PROMPTS
   ============================================================ */

const DEFAULT_RESEARCH_PROMPT = `I'm going to give you a Shopify store name or link. Research this store and its ownership using web search, and give me a clear, organized brief that helps me understand the business and the person behind it before I reach out to them.

Do the following:

Step 1: Identify the key person
Find the founder, owner, or someone with strong administrative or decision making power in the business (CEO, founder, head of operations, etc). Prioritize whoever appears to be the main decision maker.

Step 2: Find contact details
Search for their business email or any publicly available contact information (this could be on the store's About or Contact page, LinkedIn, press articles, podcast features, interviews, or business directories).

Important: the goal is to reach the actual decision maker, not a generic inbox. Do not settle for or default to generic addresses like info@, support@, hello@, sales@, or a general contact form, even if those are the easiest to find. Only list a generic address as a last resort, and clearly flag it as generic rather than direct.

If you find a direct email tied to the founder or decision maker, list it. If you can only find a likely email format (e.g. firstname@storedomain.com) based on common patterns, mention that it's inferred, not confirmed. Never guess private or personal information that isn't publicly available.

Step 3: Research the person and the store
Search for recent, relevant, and noteworthy information including:
- Their name, role, and background (how they started the business, past experience, education, or other ventures if available)
- Any recent news, launches, milestones, press features, or social media activity from the store or the founder
- Their personality, values, or tone, if it can be picked up from interviews, social posts, or public statements
- Anything unique or specific about the brand, products, mission, or growth stage
- Any pain points or signals that hint at what they might be struggling with (e.g. rapid growth, lots of apps installed, recent funding, complaints, hiring activity, etc)

Step 4: Summarize it clearly
Give a short, well organized brief (not overly long) under these sections:
- Who they are
- What I found about the business
- Recent or noteworthy updates
- Anything useful for understanding their personality or current priorities
- Contact info found (with a note on confidence level: confirmed direct, inferred direct, or generic fallback)

Use your search tool actively throughout this process rather than relying on existing knowledge, since I need current, accurate, and verifiable information.`;

const DEFAULT_EMAIL_PROMPT = `I'm going to give you details about a Shopify store, its founder/owner, and/or recent news about them. Use this to write a short, highly personalized cold outreach message from Ayomide (also known as Midey) at Phoenix Agency, rooted in these core themes (don't use them all as rigid sections, weave them naturally based on what fits the specific lead):

1. Empathy & recognition — Acknowledge the real overwhelm of running a Shopify store (juggling apps, costs, customer queries, day to day operations). Validate that they're likely overworked and stretched thin.

2. Cost reduction, no new expense — Make clear we're not pitching another subscription or tool to add to their stack. Reference the hidden costs of redundant apps, subscriptions, or manual work that could be eliminated.

3. Automation & simplification — Frame the solution as consolidating and automating what's currently manual and repetitive, ideally something built once rather than another recurring cost.

4. Freedom & time recovery — Tie this back to giving them time back for the things that actually matter: product creation, growth, or simply their life outside the business. Introduce the idea of the "Chaos Tax", the hidden cost of disorganization and inefficiency, where relevant.

5. Low pressure, no obligation — Make it clear there's zero commitment. We're not pitching upfront, just sharing something we noticed and offering to show them, nothing more.

6. Personalization & credibility — Use their name and store name, and reference something specific about their store (recent news, product line, design, founder background, etc.) that proves we actually looked at their business before reaching out.

Voice and framing requirement (important):
Write from "we" (Phoenix Agency), never "I". The message should stay centered on the recipient and their store, not on us or our agency. Keep any mention of Phoenix Agency brief and only where it adds credibility, never as the focus of a sentence. When transitioning into explaining the solution, use a short phrase like "what we do at Phoenix is help merchants like you..." or a natural variation of it, to establish agency prowess before going into specifics. Outside of that one transition phrase, the message should read like it's about them and what we noticed about their store, not about what we do or who we are.

Infrastructure paragraph requirement (important):
After the personalized opening, include one detailed paragraph explaining how the solution actually works, written specifically around what was found in the store's details (for example, if they run many sub brands or storefronts, focus on how automation ties those operations together, if they seem to have too many apps, focus on consolidation). This paragraph must clearly communicate, in this order, but in natural flowing language and not as a list:

- We do not hand over a system for them to manage. We build the automation directly into their store.
- It is built using Shopify's own free native tools, such as Shopify Flow and other built in features that match what their specific store needs (mention the relevant native feature if the store details suggest one, e.g. inventory tracking, order tagging, customer notifications).
- Once it is built and set up, it runs entirely on its own and does not require our input or anyone else's to keep working.
- Because it runs on native Shopify tools, most of the external apps and subscriptions currently in use can be removed, since those functions can now be handled for free natively, meaning no more subscription cost tied to them.
- Because everything runs directly from Shopify itself, the store stays fast, stable, and efficient, since there are no extra third party apps slowing it down.

This paragraph should sound like a natural continuation of the message, written in the same human, simple tone as the rest, not like a feature list or technical breakdown.

Tone & language requirements:
- The message must feel highly human, like it was genuinely typed by a real person who cares, not AI generated or templated.
- Do not use hyphens or em dashes anywhere in the message.
- Keep it formal enough to be taken seriously by a business owner, but the English must be simple and easy to understand, the kind a teenager could read and fully grasp. Avoid big or scholarly words, avoid corporate jargon, avoid overly polished phrasing.
- Conversational tone, not salesy, not robotic, no generic compliments.
- Because of the added infrastructure paragraph, the message can run longer than a typical cold email, but should stay under 220 words total and still feel tight and purposeful, not bloated or repetitive.
- End with ONE clear, low friction CTA: asking if they would like us to send over a 60 second video walking through what we found and how this could work for their store. Frame it as a casual reply (yes or no), not a meeting or call.
- Sign off with:

Ayo Midey
Growth Specialist | Phoenix Agency
Helping E-commerce Brands & Authors Scale

(You can create a similar signature in this style if it fits better.)

Angle selection (internal, do not show your reasoning or list options):
Before writing, silently review the lead details against the six themes above. Decide which single angle (or combination) is the strongest fit for this specific lead based on what's most specific and provable from the details given (a recent launch, a visible app stack issue, a founder story, a growth signal, etc). Do not explain this choice or present alternatives. Go straight to writing the final message using whichever angle you determined is strongest.

Also include a subject line that fits the angle and feels personal and curiosity driven, no exclamation marks, no spam sounding words.`;

const DEFAULT_HTML_PROMPT = `Convert the following cold email into clean, professional HTML that renders well in Gmail and other email clients. Use a table-based layout for maximum email client compatibility.

Requirements:
- Use a single centered table, max-width 600px, with a white or very light background
- Use system-safe fonts: Arial, Helvetica, sans-serif
- Keep font size 14-16px for body text, slightly larger for any heading if present
- Maintain all paragraph breaks from the original email
- Make links clickable with a clear color (use #0ea5e9 for link color)
- Add subtle padding around the content (20-30px on sides)
- Keep the signature visually distinct (slightly smaller font, muted color like #6b7280)
- Do NOT add images, logos, or decorative elements unless they were in the original
- Do NOT add an unsubscribe footer
- Return ONLY the complete HTML starting with <!DOCTYPE html>, no explanation, no markdown, no backticks`;

const DEFAULT_PLAIN_PROMPT = `Take the email text below and return it as clean plain text, ready to paste directly into an email client's compose window.

- Preserve every paragraph break and the exact line spacing of the signature
- Do not add any HTML tags, markdown formatting, asterisks, or bullet characters
- Do not add a subject line — that is handled separately
- Do not add any explanation, preamble, or notes of your own
- Return only the raw plain text of the email body, nothing else

Email text:
[EMAIL_TEXT]`;

/* ============================================================
   DEFAULT STATE
   ============================================================ */

const DEFAULT_STATE: PersistedState = {
  headers: [],
  rows: [],
  rowStates: {},
  targetEmailHeader: "",
  domainHeader: "",
  subjectA: "Quick question, {first_name}",
  bodyA: "Hi {first_name},\n\nNoticed {company} — wanted to reach out.\n\n— Midey",
  recipientB: "",
  sampleIdB: 0,
  subjectB: "A note for {first_name}",
  htmlB: "<div style=\"font-family:system-ui;line-height:1.55\">\n  <h2 style=\"color:#0ea5e9\">Hi {first_name} 👋</h2>\n  <p>Loved what you're doing at <b>{company}</b>.</p>\n  <p>— Phoenix Agency</p>\n</div>",
  templateSlotsA: [],
  htmlMode: false,
  templates: [],
  activeTemplateId: "",
  rotateSubjects: false,
  rotateBodies: false,
  sendCounter: 0,
  dailyGoal: 100,
  queueSearchPersisted: "",
  homepageMode: "research",
  researchDone: {},
  priorityDomain: "",
  countryHeader: "",
  priorityCountry: "",
  priorityMatchMode: "all",
};

function newId() {
  return `tpl_${Math.random().toString(36).slice(2, 9)}`;
}

function loadState(): PersistedState {
  if (typeof window === "undefined") return DEFAULT_STATE;
  try {
    const raw = localStorage.getItem(STORAGE_KEY);
    if (!raw) return DEFAULT_STATE;
    const parsed = JSON.parse(raw) as Partial<PersistedState>;
    const merged: PersistedState = { ...DEFAULT_STATE, ...parsed } as PersistedState;
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

function loadApiKeys(): ApiKey[] {
  if (typeof window === "undefined") return [];
  try {
    const raw = localStorage.getItem(API_KEYS_KEY);
    if (!raw) return [];
    return JSON.parse(raw) as ApiKey[];
  } catch { return []; }
}

function saveApiKeys(keys: ApiKey[]) {
  try { localStorage.setItem(API_KEYS_KEY, JSON.stringify(keys)); } catch {}
}

function loadPrompts() {
  const defaults = {
    research: DEFAULT_RESEARCH_PROMPT,
    email: DEFAULT_EMAIL_PROMPT,
    html: DEFAULT_HTML_PROMPT,
    plainFormat: DEFAULT_PLAIN_PROMPT,
  };
  if (typeof window === "undefined") return defaults;
  try {
    const raw = localStorage.getItem(PROMPTS_KEY);
    if (!raw) return defaults;
    const p = JSON.parse(raw);
    // The Email Writing Prompt used to end with a "Workflow" step that told
    // the model to return a JSON array of approach options instead of a
    // finished email — that stage no longer exists. If a saved prompt still
    // carries that instruction, it's a leftover from before the flow
    // simplified, not something you deliberately wrote, so auto-replace it
    // with the current one-shot default rather than silently keeping the
    // stale behavior forever.
    const looksLikeOldApproachPrompt =
      typeof p.email === "string" &&
      (/return\s+only\s+a\s+valid\s+json\s+array/i.test(p.email) || /"number":\s*integer/i.test(p.email));
    return {
      research: p.research || defaults.research,
      email: looksLikeOldApproachPrompt ? defaults.email : (p.email || defaults.email),
      html: p.html || defaults.html,
      plainFormat: p.plainFormat || defaults.plainFormat,
    };
  } catch {
    return defaults;
  }
}

/* ============================================================
   FULL BACKUP / MIGRATION
   ============================================================
   Bundles every piece of the app's data — CSV rows, send progress
   (rowStates/sendCounter), templates, filters, API keys, prompts, and the
   Research Mode per-lead cache — into one JSON file. Used to move
   everything to a new deploy/browser in one shot instead of losing
   progress or re-entering things by hand. Reads/writes localStorage
   directly (raw, unparsed by type) so it works regardless of where in the
   file the relevant interfaces are declared, and so a restore is just a
   page reload away from every component re-hydrating naturally, the same
   way it already does on first load.
   ============================================================ */

const BACKUP_KEYS = {
  state: STORAGE_KEY,
  apiKeys: API_KEYS_KEY,
  prompts: PROMPTS_KEY,
  leadCache: "midey.research.leadCache.v1",
} as const;

function buildFullBackup(): string {
  const bundle: Record<string, unknown> = {
    app: "Midey",
    backupVersion: 1,
    exportedAt: new Date().toISOString(),
  };
  for (const [name, key] of Object.entries(BACKUP_KEYS)) {
    try {
      const raw = localStorage.getItem(key);
      bundle[name] = raw ? JSON.parse(raw) : null;
    } catch {
      bundle[name] = null;
    }
  }
  return JSON.stringify(bundle, null, 2);
}

// Writes every section straight back into its localStorage key. Sections
// missing from an older/partial backup file are simply left untouched
// rather than wiping them out. Caller is responsible for reloading the
// page afterward so every component re-hydrates from the restored data.
function restoreFullBackup(json: string): { ok: boolean; restored: string[]; error?: string } {
  let parsed: Record<string, unknown>;
  try {
    parsed = JSON.parse(json);
  } catch {
    return { ok: false, restored: [], error: "That file isn't valid JSON." };
  }
  if (!parsed || typeof parsed !== "object" || parsed.app !== "Midey") {
    return { ok: false, restored: [], error: "That doesn't look like a Midey backup file." };
  }
  const restored: string[] = [];
  for (const [name, key] of Object.entries(BACKUP_KEYS)) {
    if (parsed[name] === undefined || parsed[name] === null) continue;
    try {
      localStorage.setItem(key, JSON.stringify(parsed[name]));
      restored.push(name);
    } catch {
      // Skip whatever section fails to write rather than aborting the
      // whole restore over one bad section.
    }
  }
  return { ok: restored.length > 0, restored };
}

/* ============================================================
   GEMINI API UTILITIES
   ============================================================ */

// gemini-2.0-flash was deprecated by Google on 2026-03-03 and fully shut
// down on 2026-06-01 — using it now returns errors that look like quota
// exhaustion but are really "model no longer exists". gemini-2.5-flash is
// the current, free-tier-eligible equivalent as of mid-2026.
const GEMINI_BASE = "https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent";

async function callGemini(
  keys: ApiKey[],
  setKeys: (updater: (prev: ApiKey[]) => ApiKey[]) => void,
  activeKeyIdxRef: React.MutableRefObject<number>,
  prompt: string,
  useWebSearch = false,
): Promise<string> {
  const enabledKeys = keys.filter((k) => k.enabled && k.value.trim());
  if (enabledKeys.length === 0) throw new Error("No active API keys. Add and enable at least one Gemini key in Settings.");

  // Start from the current active index within the enabled pool
  let attempts = 0;
  while (attempts < enabledKeys.length) {
    const idx = activeKeyIdxRef.current % enabledKeys.length;
    const key = enabledKeys[idx];

    const body: Record<string, unknown> = {
      contents: [{ parts: [{ text: prompt }] }],
      generationConfig: { temperature: 0.8 },
    };
    if (useWebSearch) {
      body.tools = [{ googleSearch: {} }];
    }

    try {
      const res = await fetch(`${GEMINI_BASE}?key=${encodeURIComponent(key.value)}`, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify(body),
      });

      if (res.status === 429 || res.status === 503) {
        // Quota exhausted — mark this key and rotate to next
        setKeys((prev) => prev.map((k) => k.id === key.id ? { ...k, status: "quota" } : k));
        activeKeyIdxRef.current = (activeKeyIdxRef.current + 1) % enabledKeys.length;
        attempts++;
        continue;
      }

      if (res.status === 400 || res.status === 403) {
        setKeys((prev) => prev.map((k) => k.id === key.id ? { ...k, status: "invalid" } : k));
        activeKeyIdxRef.current = (activeKeyIdxRef.current + 1) % enabledKeys.length;
        attempts++;
        continue;
      }

      if (!res.ok) {
        throw new Error(`Gemini error ${res.status}`);
      }

      const j = await res.json();
      const txt = j?.candidates?.[0]?.content?.parts?.[0]?.text;
      if (!txt) throw new Error("Empty response from Gemini");

      // Mark this key as active/healthy
      setKeys((prev) => prev.map((k) => k.id === key.id ? { ...k, status: "active" } : k));
      return String(txt).trim();

    } catch (err) {
      if ((err as Error).message.includes("Gemini error") || (err as Error).message.includes("Empty response")) {
        throw err;
      }
      attempts++;
    }
  }
  throw new Error("All API keys are exhausted or invalid. Please add more keys or wait for quota reset.");
}

async function testGeminiKey(apiKey: string): Promise<"active" | "quota" | "invalid"> {
  try {
    const res = await fetch(`${GEMINI_BASE}?key=${encodeURIComponent(apiKey)}`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ contents: [{ parts: [{ text: "Reply with the single word: ok" }] }] }),
    });
    if (res.status === 429 || res.status === 503) return "quota";
    if (res.status === 400 || res.status === 403) return "invalid";
    if (!res.ok) return "invalid";
    return "active";
  } catch {
    return "invalid";
  }
}

/* ============================================================
   SPAM / LINK SCAN UTILITIES
   ============================================================ */

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
  const opens = (html.match(/<a\b/gi) || []).length;
  const closes = (html.match(/<\/a>/gi) || []).length;
  if (opens !== closes) out.push({ tag: `${opens} open vs ${closes} close`, reason: "unbalanced <a> tags" });
  return { ok, broken: out };
}

/* ============================================================
   TEMPLATE / RENDER UTILITIES
   ============================================================ */

const TOKEN_RE = /\{([^{}]+)\}/g;

// Matches a chunk that already starts with a block-level HTML tag, so
// autoFormatHtml() leaves it as-is instead of wrapping it in a <p>.
const BLOCK_RE = /^\s*<(p|div|table|thead|tbody|tr|td|th|ul|ol|li|h[1-6]|blockquote|pre|hr|img|a\s)/i;

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

function looksLikeHtmlDocument(src: string): boolean {
  return /<!doctype|<html[\s>]|<head[\s>]|<body[\s>]|<table[\s>]|<thead[\s>]|<tbody[\s>]|<tr[\s>]|<style[\s>]|<!--/i.test(src);
}

function autoFormatHtml(src: string): string {
  if (!src) return src;
  if (looksLikeHtmlDocument(src)) return src;
    const chunks = src.split(/\n{2,}/);
  return chunks
    .map((chunk) => {
      const trimmed = chunk.trim();
      if (!trimmed) return "";
      if (BLOCK_RE.test(trimmed)) {
        return trimmed.replace(/\n/g, "<br />");
      }
      return `<p>${trimmed.replace(/\n/g, "<br />")}</p>`;
    })
    .filter(Boolean)
    .join("\n");
}

function cleanEmails(raw: string): string {
  // Some lead lists separate multiple addresses in one cell with ":" instead
  // of a comma. Gmail (and mailto:) only treats comma/semicolon as "send to
  // both" — a colon-separated string reads as one broken address — so we
  // normalize ":" the same way as "," and ";" here.
  return raw
    .split(/[,;:\n]+/)
    .map((e) => e.trim())
    .filter(Boolean)
    .join(",");
}

function buildMailto(to: string, opts: { subject?: string; body?: string; bcc?: string } = {}): string {
  // Deliberately not using URLSearchParams here: it encodes spaces as "+"
  // (form-encoding convention), but mailto: links need "%20" — mail
  // clients don't turn "+" back into a space the way a web form would,
  // so subjects/bodies would show up with "+" between every word.
  const parts: string[] = [];
  if (opts.subject) parts.push(`subject=${encodeURIComponent(opts.subject)}`);
  if (opts.body) parts.push(`body=${encodeURIComponent(opts.body)}`);
  if (opts.bcc) parts.push(`bcc=${encodeURIComponent(opts.bcc)}`);
  const q = parts.join("&");
  return `mailto:${encodeURIComponent(to)}${q ? `?${q}` : ""}`;
}

// Opens a real, full Google search in a new tab (All / Images / Videos /
// News / AI Mode, exactly as the user's own browser renders it) — this is
// deliberately NOT embedded, since Google blocks its search page from being
// iframed on any other site (X-Frame-Options), and even if it didn't, same
// origin policy would still stop the app from reading anything back out of
// it. A new tab is the only thing that's actually possible here.
function googleSearchUrl(query: string): string {
  return `https://www.google.com/search?q=${encodeURIComponent(query)}`;
}

// Pulls a recipient address out of the research brief's "Contact info"
// section. Prefers a direct (non generic-inbox) address over a fallback
// like info@ / support@ / hello@, matching how the research prompt itself
// is instructed to flag generic addresses as a last resort.
const EMAIL_IN_TEXT_RE = /[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}/g;
const GENERIC_INBOX_RE = /^(info|support|hello|sales|contact|admin|help|team|office|enquiries|inquiries)@/i;
function extractContactEmail(brief: string): string {
  if (!brief) return "";
  const matches = brief.match(EMAIL_IN_TEXT_RE) ?? [];
  if (matches.length === 0) return "";
  const direct = matches.find((e) => !GENERIC_INBOX_RE.test(e));
  return direct ?? matches[0];
}

function fmtDuration(s: number): string {
  if (s < 60) return `${s}s`;
  if (s < 3600) return `${Math.floor(s / 60)}m ${s % 60}s`;
  return `${Math.floor(s / 3600)}h ${Math.floor((s % 3600) / 60)}m`;
}

/* ============================================================
   LEGACY AI SETTINGS TYPE (kept for prop compatibility — 
   AIPersonalizationPanel removed, but SectionACard and 
   NextRowPreview still accept this prop shape; it is always
   passed as a disabled/empty object from Index() now)
   ============================================================ */

export interface AISettings {
  enabled: boolean;
  provider: "gemini" | "openai";
  apiKey: string;
  prompt: string;
  fallback: string;
  descriptionColumn: string;
}

/* ============================================================
   INDEX COMPONENT
   ============================================================ */

function Index() {
  const [hydrated, setHydrated] = useState(false);
  const [state, setState] = useState<PersistedState>(DEFAULT_STATE);
  const [theme, setTheme] = useState<"dark" | "light">("dark");
  const [parsing, setParsing] = useState(false);
  const [parseProgress, setParseProgress] = useState(0);
  const [drawerOpen, setDrawerOpen] = useState(false);

  // API key pool
  const [apiKeys, setApiKeys] = useState<ApiKey[]>([]);
  const activeKeyIdxRef = useRef(0);

  // Prompts
  const [prompts, setPrompts] = useState({
    research: DEFAULT_RESEARCH_PROMPT,
    email: DEFAULT_EMAIL_PROMPT,
    html: DEFAULT_HTML_PROMPT,
    plainFormat: DEFAULT_PLAIN_PROMPT,
  });

  // Draggable send button
  const [dragUnlocked, setDragUnlocked] = useState(false);
  const [buttonsHidden, setButtonsHidden] = useState(false);
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

  const lastTapRef = useRef(0);
  const tapTimerRef = useRef<number | null>(null);
  const onHeaderTap = useCallback(() => {
    const now = Date.now();
    const elapsed = now - lastTapRef.current;
    if (elapsed < 350 && lastTapRef.current !== 0) {
      if (tapTimerRef.current) { window.clearTimeout(tapTimerRef.current); tapTimerRef.current = null; }
      lastTapRef.current = 0;
      setDragUnlocked(true);
      toast.info("Send button unlocked — drag it anywhere");
    } else {
      lastTapRef.current = now;
      if (tapTimerRef.current) window.clearTimeout(tapTimerRef.current);
      tapTimerRef.current = window.setTimeout(() => {
        tapTimerRef.current = null;
        lastTapRef.current = 0;
        setDragUnlocked((cur) => {
          if (cur) { toast.success("Send button locked in place"); setButtonsHidden(false); return false; }
          return cur;
        });
      }, 360);
    }
  }, []);

  // Hydration
  useEffect(() => {
    setState(loadState());
    const t = (localStorage.getItem(THEME_KEY) as "dark" | "light" | null) ?? "dark";
    setTheme(t);
    setApiKeys(loadApiKeys());
    setPrompts(loadPrompts());
    setHydrated(true);
  }, []);

  // Persist state
  useEffect(() => {
    if (!hydrated) return;
    const id = setTimeout(() => {
      try { localStorage.setItem(STORAGE_KEY, JSON.stringify(state)); } catch {}
    }, 150);
    return () => clearTimeout(id);
  }, [state, hydrated]);

  // Persist API keys
  useEffect(() => {
    if (!hydrated) return;
    saveApiKeys(apiKeys);
  }, [apiKeys, hydrated]);

  // Persist prompts
  useEffect(() => {
    if (!hydrated) return;
    try { localStorage.setItem(PROMPTS_KEY, JSON.stringify(prompts)); } catch {}
  }, [prompts, hydrated]);

  // Persist drag pos
  useEffect(() => {
    if (!hydrated) return;
    try {
      if (dragPos) localStorage.setItem("midey:sendBtnPos", JSON.stringify(dragPos));
      else localStorage.removeItem("midey:sendBtnPos");
    } catch {}
  }, [dragPos, hydrated]);

  // Theme
  useEffect(() => {
    if (!hydrated) return;
    document.documentElement.classList.toggle("dark", theme === "dark");
    localStorage.setItem(THEME_KEY, theme);
  }, [theme, hydrated]);

  const patch = useCallback((p: Partial<PersistedState>) => {
    setState((s) => ({ ...s, ...p }));
  }, []);

  // Templates
  const activeTemplate: TemplateItem = useMemo(() => {
    const found = state.templates.find((t) => t.id === state.activeTemplateId);
    if (found) return found;
    return state.templates[0] ?? { id: "", name: "Default", subject: state.subjectA, body: state.bodyA, html: state.htmlB };
  }, [state.templates, state.activeTemplateId, state.subjectA, state.bodyA, state.htmlB]);

  const updateTemplate = useCallback((id: string, partial: Partial<TemplateItem>) => {
    setState((s) => ({ ...s, templates: s.templates.map((t) => (t.id === id ? { ...t, ...partial } : t)) }));
  }, []);
  const addTemplate = useCallback(() => {
    const id = newId();
    const next: TemplateItem = { id, name: `Template ${state.templates.length + 1}`, subject: activeTemplate.subject, body: activeTemplate.body, html: activeTemplate.html };
    setState((s) => ({ ...s, templates: [...s.templates, next], activeTemplateId: id }));
    toast.success(`Added "${next.name}"`);
  }, [state.templates.length, activeTemplate]);
  const deleteTemplate = useCallback((id: string) => {
    setState((s) => {
      if (s.templates.length <= 1) { toast.error("Keep at least one template"); return s; }
      const remaining = s.templates.filter((t) => t.id !== id);
      return { ...s, templates: remaining, activeTemplateId: s.activeTemplateId === id ? remaining[0].id : s.activeTemplateId };
    });
  }, []);

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

  // CSV parsing
  const onFile = useCallback((file: File) => {
    if (!file) return;
    setParsing(true);
    setParseProgress(0);
    toast.info(`Parsing ${(file.size / 1024 / 1024).toFixed(1)} MB …`);
    const isXlsx = /\.xlsx$/i.test(file.name);
    if (isXlsx) {
      const reader = new FileReader();
      reader.onerror = () => { setParsing(false); toast.error("Failed to read Excel file"); };
      reader.onload = (e) => {
        try {
          const data = new Uint8Array(e.target?.result as ArrayBuffer);
          const wb = XLSX.read(data, { type: "array" });
          const sheetName = wb.SheetNames[0];
          if (!sheetName) throw new Error("Workbook has no sheets");
          const sheet = wb.Sheets[sheetName];
          const aoa = XLSX.utils.sheet_to_json<unknown[]>(sheet, { header: 1, blankrows: false, defval: "", raw: false });
          if (!aoa.length) throw new Error("Sheet is empty");
          const seen = new Set<string>();
          const rawHeaders = (aoa[0] as unknown[]).map((h) => String(h ?? "").trim());
          const headers = rawHeaders.filter((h) => {
            if (!h || h.length > 64 || /[,\n\r"]/.test(h) || seen.has(h)) return false;
            seen.add(h); return true;
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
          setParsing(false); setParseProgress(100);
          const guessEmail = headers.find((h) => /e?mail/i.test(h)) ?? headers[0] ?? "";
          const guessDomain = headers.find((h) => /domain/i.test(h)) ?? "";
          const guessCountry =
            headers.find((h) => /^country[\s_-]?code$/i.test(h.trim())) ??
            headers.find((h) => /country/i.test(h)) ??
            "";
          const totalRows = rows.length;
          const keptRows = guessEmail ? rows.filter((r) => String(r[guessEmail] ?? "").trim() !== "") : rows;
          const skipped = totalRows - keptRows.length;
          setState((s) => ({ ...s, headers, rows: keptRows, rowStates: {}, targetEmailHeader: s.targetEmailHeader || guessEmail, domainHeader: s.domainHeader || guessDomain, countryHeader: s.countryHeader || guessCountry }));
          toast.success(`Loaded ${keptRows.length.toLocaleString()} rows · ${headers.length} columns${skipped > 0 ? ` — skipped ${skipped.toLocaleString()} with no email` : ""}`);
        } catch (err) { setParsing(false); toast.error(`Parse failed: ${(err as Error).message}`); }
      };
      reader.readAsArrayBuffer(file); return;
    }
    const collected: Row[] = [];
    let headers: string[] = [];
    const total = file.size || 1;
    Papa.parse<Row>(file, {
      header: true, skipEmptyLines: true, chunkSize: 1024 * 1024,
      chunk: (results, parser) => {
        if (!headers.length && results.meta.fields) {
          const seen = new Set<string>();
          headers = results.meta.fields.map((h) => (h ?? "").trim()).filter((h) => {
            if (!h || h.length > 64 || /[,\n\r"]/.test(h) || seen.has(h)) return false;
            seen.add(h); return true;
          });
        }
        for (const r of results.data) collected.push(r);
        const cursor = (results.meta as { cursor?: number }).cursor ?? 0;
        const pct = cursor > 0 ? Math.min(99, Math.round((cursor / total) * 100)) : Math.min(99, Math.round((collected.length / Math.max(1000, collected.length + 1000)) * 100));
        setParseProgress(pct);
        if (collected.length > 500_000) { parser.abort(); toast.error("Row cap (500k) reached — truncated."); }
      },
      complete: () => {
        setParsing(false); setParseProgress(100);
        const guessEmail = headers.find((h) => /e?mail/i.test(h)) ?? headers[0] ?? "";
        const guessDomain = headers.find((h) => /domain/i.test(h)) ?? "";
        const guessCountry =
            headers.find((h) => /^country[\s_-]?code$/i.test(h.trim())) ??
            headers.find((h) => /country/i.test(h)) ??
            "";
        const totalRows = collected.length;
        const keptRows = guessEmail ? collected.filter((r) => String(r[guessEmail] ?? "").trim() !== "") : collected;
        const skipped = totalRows - keptRows.length;
        setState((s) => ({ ...s, headers, rows: keptRows, rowStates: {}, targetEmailHeader: s.targetEmailHeader || guessEmail, domainHeader: s.domainHeader || guessDomain, countryHeader: s.countryHeader || guessCountry }));
        toast.success(`Loaded ${keptRows.length.toLocaleString()} rows · ${headers.length} columns${skipped > 0 ? ` — skipped ${skipped.toLocaleString()} with no email` : ""}`);
      },
      error: (err) => { setParsing(false); toast.error(`Parse failed: ${err.message}`); },
    });
  }, []);

  // Queue
  const queue = useMemo(() => {
    const domain = state.priorityDomain.trim().toLowerCase().replace(/^@/, "");
    const country = state.priorityCountry.trim().toLowerCase();
    const priority: number[] = [], pending: number[] = [], processed: number[] = [];
    for (let i = 0; i < state.rows.length; i++) {
      if (state.rowStates[i] === "processed" || state.rowStates[i] === "skipped") { processed.push(i); continue; }
      if (domain || country) {
        let domainMatch = true, countryMatch = true;
        if (domain) {
          const email = String(state.rows[i]?.[state.targetEmailHeader] ?? "").toLowerCase();
          // Multi-email cells (colon/comma/semicolon separated) count as a
          // match if ANY address in the cell ends with the target domain.
          domainMatch = email.split(/[,;:\s]+/).some((piece) => piece.endsWith(`@${domain}`));
        }
        if (country) {
          const cell = String(state.rows[i]?.[state.countryHeader] ?? "").trim().toLowerCase();
          countryMatch = cell === country;
        }
        const isMatch = state.priorityMatchMode === "any" && domain && country
          ? domainMatch || countryMatch
          : domainMatch && countryMatch;
        if (isMatch) { priority.push(i); continue; }
      }
      pending.push(i);
    }
    return [...priority, ...pending, ...processed];
  }, [state.rows, state.rowStates, state.priorityDomain, state.priorityCountry, state.priorityMatchMode, state.targetEmailHeader, state.countryHeader]);

  const processedCount = useMemo(
    () => Object.values(state.rowStates).filter((v) => v === "processed").length,
    [state.rowStates],
  );

  const fireRow = useCallback((rowIndex: number) => {
    const row = state.rows[rowIndex];
    if (!row) return;
    const toAddr = (row[state.targetEmailHeader] || "").trim();
    if (!toAddr) { toast.error(`Row ${rowIndex} missing "${state.targetEmailHeader}"`); return; }
    setState((s) => ({ ...s, rowStates: { ...s.rowStates, [rowIndex]: "processed" }, sendCounter: s.sendCounter + 1 }));
    sendLogRef.current.push(Date.now());
  }, [state.rows, state.targetEmailHeader]);

  const skipRow = useCallback((rowIndex: number) => {
    setState((s) => ({ ...s, rowStates: { ...s.rowStates, [rowIndex]: "skipped" } }));
  }, []);
  const resetRow = useCallback((rowIndex: number) => {
    setState((s) => { const next = { ...s.rowStates }; delete next[rowIndex]; return { ...s, rowStates: next }; });
  }, []);

  const clearAll = () => { localStorage.removeItem(STORAGE_KEY); setState(DEFAULT_STATE); toast.success("All data cleared."); };

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
        await navigator.clipboard.write([new ClipboardItem({ "text/html": blobHtml, "text/plain": blobText })]);
      } else { await navigator.clipboard.writeText(renderedHtml); }
      toast.success("Rich HTML copied — opening mail in 300ms…");
    } catch (e) { toast.error(`Clipboard failed: ${(e as Error).message}`); return; }
    setTimeout(() => { window.location.href = buildMailto(recipients, { subject: renderedSubjectB }); }, 300);
  }, [state.recipientB, renderedHtml, renderedSubjectB]);

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
    window.location.href = buildMailto(recipients, { subject, body });
  }, [state.recipientB, activeTemplate, sampleRow]);

  // Session stats
  const sessionStartRef = useRef<number>(Date.now());
  const sendLogRef = useRef<number[]>([]);
  const [, forceTick] = useState(0);
  useEffect(() => { const id = window.setInterval(() => forceTick((n) => n + 1), 1000); return () => window.clearInterval(id); }, []);
  useEffect(() => {
    if (!hydrated) return;
    const id = window.setInterval(() => {
      try {
        const firstPending = (() => { for (let i = 0; i < state.rows.length; i++) { if (!state.rowStates[i]) return i; } return state.rows.length; })();
        localStorage.setItem(AUTOSAVE_KEY, JSON.stringify(state));
        localStorage.setItem(SESSION_META_KEY, JSON.stringify({ ts: Date.now(), lastRow: firstPending, total: state.rows.length }));
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
  const acceptResume = useCallback(() => { if (resume) setResumeTarget(resume.lastRow); setResume(null); }, [resume]);

  const velocity30 = useMemo(() => {
    const cutoff = Date.now() - 30 * 60_000;
    sendLogRef.current = sendLogRef.current.filter((t) => t > cutoff - 1);
    return sendLogRef.current.length;
  }, [state.sendCounter]); // eslint-disable-line react-hooks/exhaustive-deps
  const sessionSeconds = Math.floor((Date.now() - sessionStartRef.current) / 1000);

  // Gemini caller bound to this instance's key pool
  const geminiCall = useCallback((prompt: string, useWebSearch = false) => {
    return callGemini(apiKeys, setApiKeys, activeKeyIdxRef, prompt, useWebSearch);
  }, [apiKeys]);

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
        homepageMode={state.homepageMode}
        onToggleMode={(m) => patch({ homepageMode: m })}
        onOpenDrawer={() => setDrawerOpen(true)}
      />

      {/* Side drawer */}
      <SideDrawer open={drawerOpen} onClose={() => setDrawerOpen(false)}>
        <CollapsibleSection title="Session Stats" icon={<Activity className="size-3.5 text-sky-glow" />} defaultOpen>
          <SessionStats
            processedCount={processedCount}
            totalRows={state.rows.length}
            dailyGoal={state.dailyGoal}
            onDailyGoal={(n: number) => patch({ dailyGoal: n })}
            velocity30={velocity30}
            sessionSeconds={sessionSeconds}
          />
        </CollapsibleSection>
        <CollapsibleSection title="Templates" icon={<Layers className="size-3.5 text-amber-glow" />} defaultOpen={false}>
          <TemplateControlPanel
            templates={state.templates}
            activeTemplateId={state.activeTemplateId}
            onSelect={(id: string) => patch({ activeTemplateId: id })}
            onUpdate={updateTemplate}
            onAdd={addTemplate}
            onDelete={deleteTemplate}
          />
        </CollapsibleSection>
        <CollapsibleSection title="Lead List / CSV" icon={<Upload className="size-3.5 text-sky-glow" />} defaultOpen={state.rows.length === 0} badge={<span className="font-mono-data text-[10px] text-muted-foreground">{state.rows.length.toLocaleString()} rows</span>}>
          <IngestPanel
            parsing={parsing}
            progress={parseProgress}
            onFile={onFile}
            headers={state.headers}
            totalRows={state.rows.length}
            processedRows={processedCount}
            targetEmailHeader={state.targetEmailHeader}
            onTargetEmailHeader={(v) => patch({ targetEmailHeader: v })}
            domainHeader={state.domainHeader}
            onDomainHeader={(v) => patch({ domainHeader: v })}
            countryHeader={state.countryHeader}
            onCountryHeader={(v) => patch({ countryHeader: v })}
          />
        </CollapsibleSection>
        <CollapsibleSection title="Look up by email" icon={<Search className="size-3.5 text-sky-glow" />} defaultOpen={false}>
          <LeadLookupPanel rows={state.rows} headers={state.headers} />
        </CollapsibleSection>
        <CollapsibleSection title="API Keys" icon={<Key className="size-3.5 text-sky-glow" />} defaultOpen={apiKeys.length === 0}>
          <GeminiKeyManager keys={apiKeys} onChange={setApiKeys} />
        </CollapsibleSection>
        <CollapsibleSection title="Prompt Settings" icon={<Settings className="size-3.5 text-amber-glow" />} defaultOpen={false}>
          <PromptSettingsPanel prompts={prompts} onChange={setPrompts} />
        </CollapsibleSection>
        <CollapsibleSection title="Backup & Migrate" icon={<Download className="size-3.5 text-sky-glow" />} defaultOpen={false}>
          <BackupPanel />
        </CollapsibleSection>
      </SideDrawer>

      <main className="mx-auto max-w-5xl px-3 pb-24 pt-4 sm:px-6">
        <DragContext.Provider value={{ dragUnlocked, dragPos, setDragPos, buttonsHidden, setButtonsHidden }}>
          {resume && (
            <ResumeBanner
              lastRow={resume.lastRow}
              ts={resume.ts}
              onRestore={acceptResume}
              onDismiss={() => setResume(null)}
            />
          )}
          {/* Both modes stay mounted at all times — only visibility toggles.
              Conditionally rendering (mount/unmount) here used to wipe out
              all of Research Mode's in-progress state (selected lead,
              brief, generated email) the instant you switched
              to Queue mode and back, since React destroys component state
              on unmount. Hiding with CSS instead preserves it. */}
          <div className={state.homepageMode === "research" ? "" : "hidden"}>
            <ResearchMode
              rows={state.rows}
              headers={state.headers}
              domainHeader={state.domainHeader}
              geminiCall={geminiCall}
              prompts={prompts}
              doneLeads={state.researchDone}
              onToggleDone={(idx) =>
                patch({ researchDone: { ...state.researchDone, [idx]: !state.researchDone[idx] } })
              }
              onMarkDone={(idx) => {
                if (state.researchDone[idx]) return;
                patch({ researchDone: { ...state.researchDone, [idx]: true } });
              }}
            />
          </div>
          <div className={state.homepageMode === "queue" ? "space-y-4" : "hidden"}>
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
              ai={{ enabled: false, provider: "gemini", apiKey: "", prompt: "", fallback: "", descriptionColumn: "" }}
            />
          </div>
        </DragContext.Provider>
      </main>
    </div>
  );
}

/* ============================================================
   SIDE DRAWER
   ============================================================ */

function SideDrawer({ open, onClose, children }: { open: boolean; onClose: () => void; children: React.ReactNode }) {
  // Close on ESC
  useEffect(() => {
    if (!open) return;
    const handler = (e: KeyboardEvent) => { if (e.key === "Escape") onClose(); };
    window.addEventListener("keydown", handler);
    return () => window.removeEventListener("keydown", handler);
  }, [open, onClose]);

  return (
    <>
      {/* Backdrop */}
      {open && (
        <div
          className="fixed inset-0 z-40 bg-black/50 backdrop-blur-sm"
          onClick={onClose}
          aria-hidden
        />
      )}
      {/* Panel */}
      <div
        className={`fixed right-0 top-0 z-50 flex h-full w-full max-w-sm flex-col bg-bg-app shadow-2xl transition-transform duration-200 ${
          open ? "translate-x-0" : "translate-x-full"
        }`}
      >
        <div className="flex items-center justify-between border-b border-border-strong/60 px-4 py-3">
          <span className="font-mono-data text-xs uppercase tracking-wider text-muted-foreground">Settings &amp; Tools</span>
          <button type="button" onClick={onClose} className="rounded-md p-1.5 hover:bg-surface-2" aria-label="Close">
            <X className="size-4" />
          </button>
        </div>
        <div className="flex-1 overflow-y-auto p-3 space-y-2">
          {children}
        </div>
      </div>
    </>
  );
}

/* ============================================================
   GEMINI KEY MANAGER
   ============================================================ */

function GeminiKeyManager({ keys, onChange }: { keys: ApiKey[]; onChange: (updater: (prev: ApiKey[]) => ApiKey[]) => void }) {
  const [newValue, setNewValue] = useState("");
  const [newLabel, setNewLabel] = useState("");
  const [testingAll, setTestingAll] = useState(false);
  const importInputRef = useRef<HTMLInputElement>(null);

  const addKey = () => {
    const v = newValue.trim();
    if (!v) return;
    const key: ApiKey = {
      id: `key_${Math.random().toString(36).slice(2, 9)}`,
      label: newLabel.trim() || `Key ${keys.length + 1}`,
      value: v,
      enabled: true,
      status: "unknown",
    };
    onChange((prev) => [...prev, key]);
    setNewValue("");
    setNewLabel("");
  };

  // One key per line, label in parentheses at the end: "AIza... (Key 1)".
  // Matches how you'd want to read it back on another browser/device.
  const exportKeys = () => {
    if (keys.length === 0) { toast.info("No keys to export yet"); return; }
    const lines = keys.map((k) => `${k.value} (${k.label})`);
    const blob = new Blob([lines.join("\n") + "\n"], { type: "text/plain" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = "gemini-api-keys.txt";
    a.click();
    URL.revokeObjectURL(url);
    toast.success(`Exported ${keys.length} key${keys.length > 1 ? "s" : ""}`);
  };

  const importKeys = (file: File) => {
    const reader = new FileReader();
    reader.onload = () => {
      const text = String(reader.result || "");
      const lines = text.split(/\r?\n/).map((l) => l.trim()).filter(Boolean);
      const existingValues = new Set(keys.map((k) => k.value));
      const imported: ApiKey[] = [];
      for (const line of lines) {
        const match = line.match(/^(.*?)\s*\(([^)]*)\)\s*$/);
        const value = (match ? match[1] : line).trim();
        if (!value || existingValues.has(value)) continue;
        existingValues.add(value);
        const label = (match?.[2] ?? "").trim() || `Key ${keys.length + imported.length + 1}`;
        imported.push({
          id: `key_${Math.random().toString(36).slice(2, 9)}`,
          label,
          value,
          enabled: true,
          status: "unknown",
        });
      }
      if (imported.length === 0) { toast.info("No new keys found in that file"); return; }
      onChange((prev) => [...prev, ...imported]);
      toast.success(`Imported ${imported.length} key${imported.length > 1 ? "s" : ""}`);
    };
    reader.onerror = () => toast.error("Could not read that file");
    reader.readAsText(file);
  };

  const testKey = async (id: string) => {
    const k = keys.find((k) => k.id === id);
    if (!k) return;
    onChange((prev) => prev.map((k) => k.id === id ? { ...k, status: "testing" } : k));
    const result = await testGeminiKey(k.value);
    onChange((prev) => prev.map((k) => k.id === id ? { ...k, status: result } : k));
  };

  const testAll = async () => {
    setTestingAll(true);
    for (const k of keys) {
      onChange((prev) => prev.map((pk) => pk.id === k.id ? { ...pk, status: "testing" } : pk));
      const result = await testGeminiKey(k.value);
      onChange((prev) => prev.map((pk) => pk.id === k.id ? { ...pk, status: result } : pk));
    }
    setTestingAll(false);
    toast.success("All keys tested");
  };

  const statusIcon = (status: ApiKey["status"]) => {
    if (status === "testing") return <RefreshCw className="size-3.5 animate-spin text-muted-foreground" />;
    if (status === "active") return <CheckCircle className="size-3.5 text-green-500" />;
    if (status === "quota") return <AlertCircle className="size-3.5 text-amber-glow" />;
    if (status === "invalid") return <XCircle className="size-3.5 text-destructive" />;
    return <div className="size-3.5 rounded-full border border-border-strong/60" />;
  };

  const statusLabel = (status: ApiKey["status"]) => {
    if (status === "testing") return "Testing…";
    if (status === "active") return "Active";
    if (status === "quota") return "Quota exhausted";
    if (status === "invalid") return "Invalid";
    return "Untested";
  };

  return (
    <div className="space-y-3">
      <p className="font-mono-data text-[10px] leading-relaxed text-muted-foreground">
        Add multiple Gemini API keys from different Google accounts. The app rotates automatically when one is quota-exhausted. Keys are saved only in your browser.
      </p>

      <div className="flex gap-2">
        <button
          type="button"
          onClick={exportKeys}
          className="flex flex-1 items-center justify-center gap-1.5 rounded-md border border-border-strong/60 py-2 font-mono-data text-[11px] text-muted-foreground hover:text-foreground"
        >
          <Download className="size-3.5" /> Export keys (.txt)
        </button>
        <button
          type="button"
          onClick={() => importInputRef.current?.click()}
          className="flex flex-1 items-center justify-center gap-1.5 rounded-md border border-border-strong/60 py-2 font-mono-data text-[11px] text-muted-foreground hover:text-foreground"
        >
          <Upload className="size-3.5" /> Import keys (.txt)
        </button>
        <input
          ref={importInputRef}
          type="file"
          accept=".txt,text/plain"
          className="hidden"
          onChange={(e) => {
            const file = e.target.files?.[0];
            if (file) importKeys(file);
            e.target.value = "";
          }}
        />
      </div>
      <p className="font-mono-data text-[9px] leading-relaxed text-muted-foreground">
        Format: one key per line, label in parentheses — e.g. <span className="text-foreground">AIzaSy... (Key 1)</span>
      </p>

      {/* Key list */}
      {keys.length > 0 && (
        <div className="space-y-2">
          {keys.map((k) => (
            <div key={k.id} className={`rounded-lg border p-2.5 ${k.enabled ? "border-border-strong/60" : "border-border-strong/30 opacity-50"}`}>
              <div className="flex items-center gap-2">
                <input
                  type="checkbox"
                  checked={k.enabled}
                  onChange={(e) => onChange((prev) => prev.map((pk) => pk.id === k.id ? { ...pk, enabled: e.target.checked } : pk))}
                  className="size-4 accent-[var(--sky)] shrink-0"
                  title="Enable / disable this key"
                />
                <span className="min-w-0 flex-1 font-mono-data text-[11px] text-foreground truncate">{k.label}</span>
                <span className="flex items-center gap-1 font-mono-data text-[10px] text-muted-foreground shrink-0">
                  {statusIcon(k.status)}
                  {statusLabel(k.status)}
                </span>
              </div>
              <div className="mt-1.5 flex gap-2">
                <span className="flex-1 font-mono-data text-[10px] text-muted-foreground truncate">
                  {k.value.slice(0, 6)}…{k.value.slice(-4)}
                </span>
                <button
                  type="button"
                  onClick={() => testKey(k.id)}
                  disabled={k.status === "testing"}
                  className="font-mono-data text-[10px] text-sky-glow underline decoration-dotted disabled:opacity-50"
                >
                  Test
                </button>
                <button
                  type="button"
                  onClick={() => onChange((prev) => prev.filter((pk) => pk.id !== k.id))}
                  className="font-mono-data text-[10px] text-destructive underline decoration-dotted"
                >
                  Remove
                </button>
              </div>
            </div>
          ))}
          <button
            type="button"
            onClick={testAll}
            disabled={testingAll}
            className="w-full rounded-md border border-border-strong/60 py-2 font-mono-data text-[11px] text-muted-foreground hover:text-foreground disabled:opacity-50"
          >
            {testingAll ? "Testing all…" : "Test all keys"}
          </button>
        </div>
      )}

      {/* Add new key */}
      <div className="space-y-2 rounded-lg border border-border-strong/40 bg-surface-2 p-2.5">
        <p className="font-mono-data text-[10px] uppercase tracking-wider text-muted-foreground">Add new key</p>
        <Input
          value={newLabel}
          onChange={(e) => setNewLabel(e.target.value)}
          placeholder="Label (e.g. Google Account 1)"
          className="h-8 font-mono-data text-xs"
        />
        <Input
          value={newValue}
          onChange={(e) => setNewValue(e.target.value)}
          placeholder="AIza…"
          type="password"
          className="h-8 font-mono-data text-xs"
          autoComplete="off"
          spellCheck={false}
          onKeyDown={(e) => { if (e.key === "Enter") addKey(); }}
        />
        <Button onClick={addKey} disabled={!newValue.trim()} className="w-full h-8 text-xs">
          <Plus className="size-3.5" /> Add key
        </Button>
      </div>
    </div>
  );
}

/* ============================================================
   FULL BACKUP / MIGRATE PANEL
   ============================================================ */

function BackupPanel() {
  const importInputRef = useRef<HTMLInputElement>(null);
  const [confirming, setConfirming] = useState(false);
  const [pendingFile, setPendingFile] = useState<File | null>(null);

  const exportAll = () => {
    try {
      const json = buildFullBackup();
      const blob = new Blob([json], { type: "application/json" });
      const url = URL.createObjectURL(blob);
      const a = document.createElement("a");
      const stamp = new Date().toISOString().slice(0, 10);
      a.href = url;
      a.download = `midey-backup-${stamp}.json`;
      a.click();
      URL.revokeObjectURL(url);
      toast.success("Backup downloaded");
    } catch (e) {
      toast.error(`Export failed: ${(e as Error).message}`);
    }
  };

  const runImport = (file: File) => {
    const reader = new FileReader();
    reader.onload = () => {
      const text = String(reader.result || "");
      const result = restoreFullBackup(text);
      if (!result.ok) {
        toast.error(result.error ?? "Nothing in that file could be restored");
        return;
      }
      toast.success(`Restored ${result.restored.join(", ")} — reloading…`);
      setTimeout(() => window.location.reload(), 800);
    };
    reader.onerror = () => toast.error("Could not read that file");
    reader.readAsText(file);
  };

  return (
    <div className="space-y-3">
      <p className="font-mono-data text-[10px] leading-relaxed text-muted-foreground">
        Bundles everything into one file: your CSV rows, send progress (what's been sent/skipped and your counters),
        templates, priority filters, API keys, prompts, and everything cached in Research Mode. Use this to move
        to a new deploy or browser without starting over.
      </p>

      <button
        type="button"
        onClick={exportAll}
        className="flex w-full items-center justify-center gap-1.5 rounded-md border border-sky-glow/40 bg-sky-glow/10 py-2.5 font-mono-data text-[11px] text-sky-glow hover:bg-sky-glow/20"
      >
        <Download className="size-3.5" /> Export everything (.json)
      </button>

      <button
        type="button"
        onClick={() => importInputRef.current?.click()}
        className="flex w-full items-center justify-center gap-1.5 rounded-md border border-border-strong/60 py-2.5 font-mono-data text-[11px] text-muted-foreground hover:text-foreground"
      >
        <Upload className="size-3.5" /> Import everything (.json)
      </button>
      <input
        ref={importInputRef}
        type="file"
        accept=".json,application/json"
        className="hidden"
        onChange={(e) => {
          const file = e.target.files?.[0];
          e.target.value = "";
          if (file) { setPendingFile(file); setConfirming(true); }
        }}
      />

      {confirming && pendingFile && (
        <div className="rounded-lg border border-amber-glow/40 bg-amber-glow/5 p-3 space-y-2">
          <p className="font-mono-data text-[11px] text-amber-glow">
            This replaces your current CSV, templates, send progress, API keys, and prompts with what's in{" "}
            <span className="text-foreground">{pendingFile.name}</span>. This can't be undone — export your current
            data first if you're not sure.
          </p>
          <div className="flex gap-2">
            <Button
              size="sm"
              onClick={() => { const f = pendingFile; setConfirming(false); setPendingFile(null); if (f) runImport(f); }}
              className="flex-1 glow-amber bg-[var(--amber)] text-black hover:bg-[var(--amber)]/90"
            >
              Yes, restore and reload
            </Button>
            <Button
              size="sm"
              variant="ghost"
              onClick={() => { setConfirming(false); setPendingFile(null); }}
              className="flex-1"
            >
              Cancel
            </Button>
          </div>
        </div>
      )}
    </div>
  );
}

/* ============================================================
   PROMPT SETTINGS PANEL
   ============================================================ */

const PROMPT_LABELS: Record<"research" | "email" | "html" | "plainFormat", string> = {
  research: "Research Prompt",
  email: "Email Writing Prompt",
  html: "HTML Output Prompt",
  plainFormat: "Plain Text Output Prompt (used when Plain Text mode is selected)",
};

const PROMPT_DEFAULTS = {
  research: DEFAULT_RESEARCH_PROMPT,
  email: DEFAULT_EMAIL_PROMPT,
  html: DEFAULT_HTML_PROMPT,
  plainFormat: DEFAULT_PLAIN_PROMPT,
};

type PromptSet = { research: string; email: string; html: string; plainFormat: string };

function PromptSettingsPanel({
  prompts, onChange,
}: {
  prompts: PromptSet;
  onChange: (p: PromptSet) => void;
}) {
  const [local, setLocal] = useState(prompts);
  const [saved, setSaved] = useState(false);

  useEffect(() => { setLocal(prompts); }, [prompts]);

  const save = () => {
    onChange(local);
    setSaved(true);
    toast.success("Prompts saved");
    setTimeout(() => setSaved(false), 2000);
  };

  const reset = (key: keyof PromptSet) => {
    setLocal((p) => ({ ...p, [key]: PROMPT_DEFAULTS[key] }));
    toast.info(`${key} prompt reset to default`);
  };

  return (
    <div className="space-y-4">
      <p className="font-mono-data text-[10px] leading-relaxed text-muted-foreground">
        Edit the prompts that drive the Research + Email Writer. Changes take effect on the next AI call. Use Reset to restore the original.
      </p>

      {(["research", "email", "plainFormat", "html"] as const).map((key) => (
        <div key={key} className="space-y-1.5">
          <div className="flex items-center justify-between gap-2">
            <Label className="font-mono-data text-[10px] uppercase tracking-wider text-muted-foreground">
              {PROMPT_LABELS[key]}
            </Label>
            <button
              type="button"
              onClick={() => reset(key)}
              className="flex shrink-0 items-center gap-1 font-mono-data text-[10px] text-muted-foreground underline decoration-dotted hover:text-foreground"
            >
              <RotateCcw className="size-3" /> Reset
            </button>
          </div>
          <Textarea
            value={local[key]}
            onChange={(e) => setLocal((p) => ({ ...p, [key]: e.target.value }))}
            rows={key === "html" || key === "plainFormat" ? 6 : 10}
            className="font-mono-data text-[11px] leading-relaxed"
            spellCheck={false}
          />
        </div>
      ))}

      <Button onClick={save} className="w-full glow-sky">
        <Save className="size-3.5" />
        {saved ? "Saved!" : "Save all prompts"}
      </Button>
    </div>
  );
}

/* ============================================================
   RESEARCH MODE
   ============================================================ */

interface GeneratedEmail {
  title: string;
  subject: string;
  plain: string;
  html: string;
  format: "html" | "plain";
}

type ResearchStage = "idle" | "researching" | "brief" | "generatingEmail" | "email";

// Saved per-lead so switching between leads (or reloading the page) never
// forces you to re-spend API calls on a store you've already researched.
interface LeadCacheEntry {
  storeInput: string;
  brief: string;
  extraDetails: string;
  outputFormat: "html" | "plain";
  generatedEmails: GeneratedEmail[];
  savedAt: string;
}

const LEAD_CACHE_KEY = "midey.research.leadCache.v1";

function loadLeadCache(): Record<string, LeadCacheEntry> {
  if (typeof window === "undefined") return {};
  try {
    const raw = localStorage.getItem(LEAD_CACHE_KEY);
    if (!raw) return {};
    return JSON.parse(raw) as Record<string, LeadCacheEntry>;
  } catch { return {}; }
}

function stageForCacheEntry(entry: LeadCacheEntry | undefined): ResearchStage {
  if (!entry) return "idle";
  if (entry.generatedEmails.length > 0) return "email";
  if (entry.brief) return "brief";
  return "idle";
}

function ResearchMode({
  rows, headers, domainHeader, geminiCall, prompts,
  doneLeads, onToggleDone, onMarkDone,
}: {
  rows: Row[];
  headers: string[];
  domainHeader: string;
  geminiCall: (prompt: string, useWebSearch?: boolean) => Promise<string>;
  prompts: PromptSet;
  doneLeads: Record<number, boolean>;
  onToggleDone: (idx: number) => void;
  onMarkDone: (idx: number) => void;
}) {
  const [selectedLeadIdx, setSelectedLeadIdx] = useState<number | null>(null);
  const [manualInput, setManualInput] = useState("");
  const [stage, setStage] = useState<ResearchStage>("idle");
  const [brief, setBrief] = useState("");
  // Free-text box for anything you noticed that the research brief missed —
  // gets folded into the email prompt alongside the brief.
  const [extraDetails, setExtraDetails] = useState("");
  // Whether the generated email is written as HTML (with live preview +
  // HTML conversion step) or as plain text (mailto body prefilled directly).
  const [outputFormat, setOutputFormat] = useState<"html" | "plain">("html");
  const [generatedEmails, setGeneratedEmails] = useState<GeneratedEmail[]>([]);
  const [activeEmailTab, setActiveEmailTab] = useState(0);
  const [editingHtml, setEditingHtml] = useState(false);
  const [searchQuery, setSearchQuery] = useState("");

  // Per-lead cache: research/email already generated for a given store,
  // persisted to localStorage so nothing needs re-fetching just because
  // you switched leads, switched to Queue mode, or reloaded.
  const [leadCache, setLeadCache] = useState<Record<string, LeadCacheEntry>>(() => loadLeadCache());
  useEffect(() => {
    try { localStorage.setItem(LEAD_CACHE_KEY, JSON.stringify(leadCache)); } catch {}
  }, [leadCache]);

  const leadKey = useMemo(() => {
    if (selectedLeadIdx !== null) return `row:${selectedLeadIdx}`;
    const trimmed = manualInput.trim().toLowerCase();
    return trimmed ? `manual:${trimmed}` : null;
  }, [selectedLeadIdx, manualInput]);

  const storeInput = useMemo(() => {
    if (selectedLeadIdx !== null && rows[selectedLeadIdx]) {
      const row = rows[selectedLeadIdx];
      return row[domainHeader] || Object.values(row).join(", ");
    }
    return manualInput.trim();
  }, [selectedLeadIdx, rows, domainHeader, manualInput]);

  const filteredRows = useMemo(() => {
    if (!searchQuery.trim()) return rows.map((_, i) => i);
    const q = searchQuery.toLowerCase();
    return rows.map((_, i) => i).filter((i) => {
      const row = rows[i];
      const domain = row[domainHeader] || "";
      return domain.toLowerCase().includes(q) || Object.values(row).some((v) => String(v).toLowerCase().includes(q));
    });
  }, [rows, domainHeader, searchQuery]);

  // Recipient pulled from the brief's contact-info section — used to
  // prefill the mailto: on "Send this one". Left blank if none was found.
  const contactEmail = useMemo(() => extractContactEmail(brief), [brief]);

  const reset = () => {
    setStage("idle");
    setBrief("");
    setExtraDetails("");
    setOutputFormat("html");
    setGeneratedEmails([]);
    setActiveEmailTab(0);
    setEditingHtml(false);
  };

  // Loads a saved lead's research/email from cache instead of starting
  // from scratch — this is what lets you click back into a lead you've
  // already worked through without spending another API call.
  const hydrateFromCache = (entry: LeadCacheEntry) => {
    setBrief(entry.brief);
    setExtraDetails(entry.extraDetails ?? "");
    setOutputFormat(entry.outputFormat ?? "html");
    setGeneratedEmails(entry.generatedEmails);
    setActiveEmailTab(0);
    setEditingHtml(false);
    setStage(stageForCacheEntry(entry));
  };

  const selectLead = (idx: number) => {
    setSelectedLeadIdx(idx);
    setManualInput("");
    const cached = leadCache[`row:${idx}`];
    if (cached) hydrateFromCache(cached);
    else reset();
  };

  // Merges a partial update into the cache entry for whichever lead is
  // currently active, keyed by row index or by the manually typed store
  // name if no CSV lead is selected.
  const saveToCache = (partial: Partial<LeadCacheEntry>) => {
    if (!leadKey) return;
    setLeadCache((prev) => {
      const existing = prev[leadKey];
      const next: LeadCacheEntry = {
        storeInput,
        brief: existing?.brief ?? "",
        extraDetails: existing?.extraDetails ?? "",
        outputFormat: existing?.outputFormat ?? "html",
        generatedEmails: existing?.generatedEmails ?? [],
        savedAt: new Date().toISOString(),
        ...partial,
      };
      return { ...prev, [leadKey]: next };
    });
  };

  const runResearch = async () => {
    const input = storeInput;
    if (!input) { toast.error("Enter a store name or URL, or select a lead from the list"); return; }
    setStage("researching");
    setBrief("");
    setGeneratedEmails([]);
    try {
      const result = await geminiCall(`${prompts.research}\n\nStore to research: ${input}`, true);
      setBrief(result);
      setStage("brief");
      // Fresh research invalidates any prior generated email for this lead.
      saveToCache({ brief: result, generatedEmails: [] });
    } catch (err) {
      toast.error((err as Error).message);
      setStage("idle");
    }
  };

  // Runs the writing prompt, then either the HTML conversion prompt or the
  // plain-text formatting prompt depending on outputFormat, and returns a
  // finished GeneratedEmail.
  const writeAndFormatEmail = async (writingPrompt: string): Promise<GeneratedEmail> => {
    const plainResult = await geminiCall(writingPrompt);
    const lines = plainResult.split("\n");
    const subjectLine = lines.find((l) => l.toLowerCase().startsWith("subject:")) ?? "";
    const subject = subjectLine.replace(/^subject:\s*/i, "").trim();
    const bodyStart = lines.findIndex((l) => l.toLowerCase().startsWith("subject:"));
    const plainBody = lines.slice(bodyStart + 1).join("\n").trim();

    if (outputFormat === "plain") {
      const hasPlaceholder = prompts.plainFormat.includes("[EMAIL_TEXT]");
      const plainPrompt = hasPlaceholder
        ? prompts.plainFormat.replace("[EMAIL_TEXT]", plainBody)
        : `${prompts.plainFormat}\n\n${plainBody}`;
      const formatted = await geminiCall(plainPrompt);
      return { title: "Email", subject, plain: formatted, html: "", format: "plain" };
    }

    const htmlPrompt = `${prompts.html}\n\nEmail to convert:\n\nSubject: ${subject}\n\n${plainBody}`;
    const htmlResult = await geminiCall(htmlPrompt);
    return { title: "Email", subject, plain: plainBody, html: htmlResult, format: "html" };
  };

  const generateEmail = async () => {
    if (!brief) return;
    setStage("generatingEmail");
    try {
      const detailsBlock = extraDetails.trim()
        ? `\n\nAdditional details to factor in (I noticed these myself, the research brief may have missed them):\n${extraDetails.trim()}`
        : "";
      const emailPrompt = `${prompts.email}\n\nLead details:\n${brief}${detailsBlock}\n\nReturn it as plain text with "Subject: ..." on the first line, then a blank line, then the email body.`;
      const email = await writeAndFormatEmail(emailPrompt);
      setGeneratedEmails([email]);
      setActiveEmailTab(0);
      setStage("email");
      saveToCache({ generatedEmails: [email], extraDetails, outputFormat });
      if (selectedLeadIdx !== null) onMarkDone(selectedLeadIdx);
    } catch (err) {
      toast.error(`Failed to generate email: ${(err as Error).message}`);
      setStage("brief");
    }
  };

  const copyHtml = async (html: string) => {
    try {
      const blobHtml = new Blob([html], { type: "text/html" });
      const blobText = new Blob([html.replace(/<[^>]+>/g, "")], { type: "text/plain" });
      if ("ClipboardItem" in window && navigator.clipboard?.write) {
        await navigator.clipboard.write([new ClipboardItem({ "text/html": blobHtml, "text/plain": blobText })]);
      } else { await navigator.clipboard.writeText(html); }
      toast.success("HTML copied to clipboard — paste into Gmail");
    } catch (e) { toast.error(`Clipboard failed: ${(e as Error).message}`); }
  };

  const copyPlain = async (text: string) => {
    try {
      await navigator.clipboard.writeText(text);
      toast.success("Text copied to clipboard");
    } catch (e) { toast.error(`Clipboard failed: ${(e as Error).message}`); }
  };

  // Copies the HTML to clipboard (same as Copy HTML) then opens a mailto:
  // with the subject and recipient prefilled, so the only remaining step
  // is pasting the already-copied HTML into the compose window body.
  const sendEmail = async (subject: string, html: string) => {
    await copyHtml(html);
    const href = buildMailto(contactEmail, { subject });
    setTimeout(() => { window.location.href = href; }, 300);
  };

  // Plain text can actually go straight into the mailto: body param —
  // no clipboard round trip needed, the compose window opens fully filled in.
  const sendPlainEmail = (subject: string, body: string) => {
    const href = buildMailto(contactEmail, { subject, body });
    window.location.href = href;
  };

  return (
    <div className="space-y-4">
      {/* Lead list (only shown when rows available and no email generated yet) */}
      {rows.length > 0 && stage !== "email" && (
        <div className="rounded-xl border border-border-strong/70 bg-surface-1 overflow-hidden">
          <div className="flex items-center gap-2 border-b border-border-strong/60 px-3 py-2">
            <Search className="size-3.5 text-muted-foreground shrink-0" />
            <input
              value={searchQuery}
              onChange={(e) => setSearchQuery(e.target.value)}
              placeholder={`Search ${rows.length.toLocaleString()} leads…`}
              className="flex-1 bg-transparent font-mono-data text-xs text-foreground outline-none placeholder:text-muted-foreground"
            />
            {searchQuery && (
              <button type="button" onClick={() => setSearchQuery("")} className="text-muted-foreground">
                <X className="size-3.5" />
              </button>
            )}
          </div>
          <div className="max-h-48 overflow-y-auto divide-y divide-border-strong/30">
            {filteredRows.slice(0, 200).map((i) => {
              const row = rows[i];
              const domain = row[domainHeader] || Object.values(row)[0] || `Row ${i}`;
              const isDone = !!doneLeads[i];
              const hasCache = !!leadCache[`row:${i}`];
              return (
                <div
                  key={i}
                  className={`flex w-full items-center gap-2 px-3 py-2 hover:bg-surface-2 ${selectedLeadIdx === i ? "bg-sky-glow/10 border-l-2 border-sky-glow" : ""}`}
                >
                  <button type="button" onClick={() => selectLead(i)} className="flex min-w-0 flex-1 items-center gap-2 text-left">
                    <Globe className="size-3 text-muted-foreground shrink-0" />
                    <span className="min-w-0 flex-1 font-mono-data text-xs text-foreground truncate">{domain}</span>
                    {hasCache && (
                      <span className="shrink-0 font-mono-data text-[9px] uppercase tracking-wider text-sky-glow" title="Saved — opens without using API">
                        saved
                      </span>
                    )}
                  </button>
                  <a
                    href={googleSearchUrl(domain)}
                    target="_blank"
                    rel="noopener noreferrer"
                    onClick={(e) => e.stopPropagation()}
                    title="Search this store on Google (opens a new tab)"
                    className="flex size-5 shrink-0 items-center justify-center rounded-full border border-border-strong/60 text-muted-foreground/50 hover:text-foreground hover:border-foreground/40"
                  >
                    <ExternalLink className="size-3" />
                  </a>
                  <button
                    type="button"
                    onClick={(e) => { e.stopPropagation(); onToggleDone(i); }}
                    title={isDone ? "Marked done — tap to unmark" : "Mark this lead as done"}
                    className={`flex size-5 shrink-0 items-center justify-center rounded-full border transition-colors ${
                      isDone ? "border-emerald-500/60 bg-emerald-500/15 text-emerald-400" : "border-border-strong/60 text-muted-foreground/50 hover:text-muted-foreground"
                    }`}
                  >
                    <CheckCircle className="size-3.5" />
                  </button>
                  <ChevronRight className="size-3 text-muted-foreground shrink-0" onClick={() => selectLead(i)} />
                </div>
              );
            })}
            {filteredRows.length === 0 && (
              <p className="py-4 text-center font-mono-data text-xs text-muted-foreground">No leads match your search</p>
            )}
          </div>
        </div>
      )}

      {/* Manual input */}
      {stage === "idle" || stage === "researching" ? (
        <div className="rounded-xl border border-border-strong/70 bg-surface-1 p-4 space-y-3">
          <p className="font-mono-data text-[10px] uppercase tracking-wider text-muted-foreground">
            {selectedLeadIdx !== null ? "Selected lead" : "Or enter a store manually"}
          </p>
          {selectedLeadIdx !== null && rows[selectedLeadIdx] ? (
            <div className="flex items-center gap-2 rounded-lg border border-sky-glow/40 bg-sky-glow/5 px-3 py-2">
              <Globe className="size-4 text-sky-glow shrink-0" />
              <span className="flex-1 font-mono-data text-sm text-foreground truncate">
                {rows[selectedLeadIdx][domainHeader] || Object.values(rows[selectedLeadIdx])[0]}
              </span>
              <button type="button" onClick={() => { setSelectedLeadIdx(null); reset(); }} className="text-muted-foreground hover:text-foreground">
                <X className="size-4" />
              </button>
            </div>
          ) : (
            <Input
              value={manualInput}
              onChange={(e) => setManualInput(e.target.value)}
              placeholder="e.g. allbirds.com or Allbirds"
              className="font-mono-data text-sm"
              onKeyDown={(e) => { if (e.key === "Enter") runResearch(); }}
            />
          )}
          <div className="flex gap-2">
            <Button
              onClick={runResearch}
              disabled={stage === "researching" || !storeInput}
              className="flex-1 glow-sky"
            >
              {stage === "researching" ? (
                <><RefreshCw className="size-4 animate-spin" /> Researching…</>
              ) : (
                <><Globe className="size-4" /> Research this store</>
              )}
            </Button>
            {storeInput && (
              <a
                href={googleSearchUrl(storeInput)}
                target="_blank"
                rel="noopener noreferrer"
                title="Search this store on Google (opens a new tab — All / Images / News / AI Mode)"
                className="flex shrink-0 items-center justify-center gap-1.5 rounded-md border border-border-strong/60 px-3 font-mono-data text-[11px] text-muted-foreground hover:text-foreground"
              >
                <ExternalLink className="size-3.5" />
              </a>
            )}
          </div>
        </div>
      ) : null}

      {/* Research brief */}
      {(stage === "brief" || stage === "generatingEmail") && brief && (
        <div className="rounded-xl border border-border-strong/70 bg-surface-1 overflow-hidden">
          <div className="flex items-center justify-between border-b border-border-strong/60 px-4 py-2.5">
            <span className="flex items-center gap-2 font-mono-data text-[11px] uppercase tracking-wider text-muted-foreground">
              <BookOpen className="size-3.5 text-sky-glow" /> Research Brief
            </span>
            <div className="flex items-center gap-3">
              <a
                href={googleSearchUrl(storeInput)}
                target="_blank"
                rel="noopener noreferrer"
                className="flex items-center gap-1 font-mono-data text-[10px] text-muted-foreground underline decoration-dotted hover:text-foreground"
                title="Search this store on Google (opens a new tab)"
              >
                <ExternalLink className="size-3" /> Google it
              </a>
              <button type="button" onClick={runResearch} className="font-mono-data text-[10px] text-muted-foreground underline decoration-dotted hover:text-foreground">
                Re-research
              </button>
              <button type="button" onClick={reset} className="font-mono-data text-[10px] text-muted-foreground underline decoration-dotted hover:text-foreground">
                Start over
              </button>
            </div>
          </div>
          <div className="p-4">
            <div className="prose prose-sm prose-invert max-w-none">
              <pre className="whitespace-pre-wrap font-sans text-sm leading-relaxed text-foreground">{brief}</pre>
            </div>
          </div>

          <div className="border-t border-border-strong/60 p-4 space-y-3">
            <div className="space-y-1.5">
              <Label className="font-mono-data text-[10px] uppercase tracking-wider text-muted-foreground">
                Anything to add? (optional)
              </Label>
              <Textarea
                value={extraDetails}
                onChange={(e) => setExtraDetails(e.target.value)}
                placeholder="Anything the research missed — a detail you noticed, a specific angle you want used, something you already know about this lead…"
                rows={3}
                className="font-sans text-sm leading-relaxed"
              />
            </div>

            <div className="rounded-lg border border-border-strong/60 bg-surface-2 p-2.5 space-y-2">
              <p className="font-mono-data text-[10px] uppercase tracking-wider text-muted-foreground">
                Output format
              </p>
              <div className="flex gap-2">
                <button
                  type="button"
                  onClick={() => setOutputFormat("plain")}
                  className={`flex-1 rounded-md border py-2 px-2 text-left font-mono-data text-[11px] ${outputFormat === "plain" ? "border-sky-glow/60 bg-sky-glow/10 text-sky-glow" : "border-border-strong/60 text-muted-foreground"}`}
                >
                  <span className="block font-semibold">Plain text</span>
                  <span className="block text-[10px] opacity-80 normal-case">No HTML — mailto opens with the body already filled in</span>
                </button>
                <button
                  type="button"
                  onClick={() => setOutputFormat("html")}
                  className={`flex-1 rounded-md border py-2 px-2 text-left font-mono-data text-[11px] ${outputFormat === "html" ? "border-sky-glow/60 bg-sky-glow/10 text-sky-glow" : "border-border-strong/60 text-muted-foreground"}`}
                >
                  <span className="block font-semibold">HTML</span>
                  <span className="block text-[10px] opacity-80 normal-case">Styled email, live preview, copy HTML to paste in</span>
                </button>
              </div>
            </div>

            {generatedEmails.length > 0 && (
              <button
                type="button"
                onClick={() => setStage("email")}
                className="w-full font-mono-data text-[10px] text-sky-glow underline decoration-dotted hover:text-sky-glow/80"
              >
                ← Back to the email already generated for this lead
              </button>
            )}

            <Button onClick={generateEmail} disabled={stage === "generatingEmail"} className="w-full glow-sky">
              {stage === "generatingEmail" ? (
                <><RefreshCw className="size-4 animate-spin" /> Writing email…</>
              ) : (
                <><PenLine className="size-4" /> {generatedEmails.length > 0 ? "Regenerate email" : "Generate email"}</>
              )}
            </Button>
          </div>
        </div>
      )}

      {/* Generated emails */}
      {stage === "email" && generatedEmails.length > 0 && (
        <div className="rounded-xl border border-border-strong/70 bg-surface-1 overflow-hidden">
          <div className="flex items-center justify-between border-b border-border-strong/60 px-4 py-2.5">
            <span className="flex items-center gap-2 font-mono-data text-[11px] uppercase tracking-wider text-muted-foreground">
              <PenLine className="size-3.5 text-sky-glow" /> Generated Emails
            </span>
            <button type="button" onClick={() => setStage("brief")} className="font-mono-data text-[10px] text-muted-foreground underline decoration-dotted hover:text-foreground">
              ← Back to brief
            </button>
          </div>

          {/* Tabs if multiple emails */}
          {generatedEmails.length > 1 && (
            <div className="flex border-b border-border-strong/60 overflow-x-auto">
              {generatedEmails.map((e, i) => (
                <button
                  key={i}
                  type="button"
                  onClick={() => setActiveEmailTab(i)}
                  className={`shrink-0 px-4 py-2 font-mono-data text-[11px] whitespace-nowrap ${
                    activeEmailTab === i
                      ? "border-b-2 border-sky-glow text-sky-glow"
                      : "text-muted-foreground hover:text-foreground"
                  }`}
                >
                  #{e.approach} {e.title}
                </button>
              ))}
            </div>
          )}

          {/* Active email */}
          {(() => {
            const email = generatedEmails[activeEmailTab];
            if (!email) return null;
            return (
              <div className="p-4 space-y-3">
                <div className="rounded-lg border border-border-strong/60 bg-surface-2 px-3 py-2">
                  <p className="font-mono-data text-[10px] uppercase tracking-wider text-muted-foreground mb-1">Subject line</p>
                  <input
                    value={email.subject}
                    onChange={(e) => {
                      const newSubject = e.target.value;
                      setGeneratedEmails((prev) => prev.map((em, i) => i === activeEmailTab ? { ...em, subject: newSubject } : em));
                    }}
                    className="w-full bg-transparent font-sans text-sm font-medium text-foreground outline-none"
                    placeholder="Subject line…"
                  />
                </div>

                <div className="rounded-lg border border-border-strong/60 bg-surface-2 px-3 py-2">
                  <p className="font-mono-data text-[10px] uppercase tracking-wider text-muted-foreground mb-1">
                    Recipient {contactEmail ? <span className="text-sky-glow normal-case">(from brief)</span> : <span className="text-amber-glow normal-case">(not found — fill in manually)</span>}
                  </p>
                  <input
                    value={contactEmail}
                    readOnly
                    placeholder="No contact email found in brief"
                    className="w-full bg-transparent font-mono-data text-xs text-foreground outline-none placeholder:text-muted-foreground"
                  />
                </div>

                {email.format === "plain" ? (
                  <>
                    <Textarea
                      value={email.plain}
                      onChange={(e) => {
                        const newPlain = e.target.value;
                        setGeneratedEmails((prev) => prev.map((em, i) => i === activeEmailTab ? { ...em, plain: newPlain } : em));
                      }}
                      rows={12}
                      className="font-sans text-sm leading-relaxed"
                      spellCheck
                    />

                    <Button onClick={() => copyPlain(email.plain)} className="w-full glow-sky">
                      <Clipboard className="size-4" /> Copy text
                    </Button>

                    <Button
                      onClick={() => sendPlainEmail(email.subject, email.plain)}
                      className="w-full glow-amber bg-[var(--amber)] text-black hover:bg-[var(--amber)]/90"
                    >
                      <Send className="size-4" /> Send this one
                    </Button>
                    <p className="font-mono-data text-[10px] leading-relaxed text-muted-foreground">
                      Opens your mail app with the subject, recipient, and body all already filled in — plain text mailto links can carry the body directly.
                    </p>
                  </>
                ) : (
                  <>
                    <div className="flex gap-2">
                      <button
                        type="button"
                        onClick={() => setEditingHtml(false)}
                        className={`flex-1 rounded-md border py-1.5 font-mono-data text-[11px] ${!editingHtml ? "border-sky-glow/60 bg-sky-glow/10 text-sky-glow" : "border-border-strong/60 text-muted-foreground"}`}
                      >
                        Preview
                      </button>
                      <button
                        type="button"
                        onClick={() => setEditingHtml(true)}
                        className={`flex-1 rounded-md border py-1.5 font-mono-data text-[11px] ${editingHtml ? "border-sky-glow/60 bg-sky-glow/10 text-sky-glow" : "border-border-strong/60 text-muted-foreground"}`}
                      >
                        Edit HTML
                      </button>
                    </div>

                    {editingHtml ? (
                      <Textarea
                        value={email.html}
                        onChange={(e) => {
                          const newHtml = e.target.value;
                          setGeneratedEmails((prev) => prev.map((em, i) => i === activeEmailTab ? { ...em, html: newHtml } : em));
                        }}
                        rows={12}
                        className="font-mono-data text-[11px] leading-relaxed"
                        spellCheck={false}
                      />
                    ) : (
                      <div className="overflow-hidden rounded-lg border border-border-strong/60 bg-white">
                        <iframe
                          srcDoc={`<!doctype html><html><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"></head><body style="margin:0;padding:0">${email.html}</body></html>`}
                          className="w-full"
                          style={{ height: "400px", border: "none" }}
                          title="Email preview"
                          sandbox="allow-same-origin"
                        />
                      </div>
                    )}

                    <Button onClick={() => copyHtml(email.html)} className="w-full glow-sky">
                      <Clipboard className="size-4" /> Copy HTML — paste into Gmail
                    </Button>

                    <Button
                      onClick={() => sendEmail(email.subject, email.html)}
                      className="w-full glow-amber bg-[var(--amber)] text-black hover:bg-[var(--amber)]/90"
                    >
                      <Send className="size-4" /> Send this one
                    </Button>
                    <p className="font-mono-data text-[10px] leading-relaxed text-muted-foreground">
                      Copies the HTML to your clipboard, then opens your mail app with the subject and recipient filled in — paste the HTML into the compose body.
                    </p>
                  </>
                )}

                <button
                  type="button"
                  onClick={reset}
                  className="w-full rounded-md border border-border-strong/60 py-2 font-mono-data text-[11px] text-muted-foreground hover:text-foreground"
                >
                  Research another store
                </button>
              </div>
            );
          })()}
        </div>
      )}
    </div>
  );
}


/* --------------------- Draggable Send button shell --------------------- */

type DragCtx = {
  dragUnlocked: boolean;
  dragPos: { x: number; y: number } | null;
  setDragPos: (p: { x: number; y: number } | null) => void;
  // Only ever applies while dragUnlocked (double-tap) mode is active —
  // locked/normal mode always shows Send/Skip regardless of this flag.
  buttonsHidden: boolean;
  setButtonsHidden: (v: boolean) => void;
};
const DragContext = createContext<DragCtx>({
  dragUnlocked: false,
  dragPos: null,
  setDragPos: () => {},
  buttonsHidden: false,
  setButtonsHidden: () => {},
});

function DraggableSendShell({ children }: { children: React.ReactNode }) {
  const { dragUnlocked, dragPos, setDragPos, buttonsHidden, setButtonsHidden } = useContext(DragContext);
  const ref = useRef<HTMLDivElement>(null);
  const dragRef = useRef<{
    active: boolean;
    offX: number;
    offY: number;
    moved: boolean;
  }>({ active: false, offX: 0, offY: 0, moved: false });

  const isFloating = dragUnlocked || dragPos !== null;
  // Hiding only ever takes effect while double-tap (drag) mode is active —
  // in normal locked mode, Send/Skip always render, untouched.
  const collapsed = dragUnlocked && buttonsHidden;

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
      {dragUnlocked && (
        <div className="mb-1 flex justify-end">
          <button
            type="button"
            // Toggling shouldn't also start a drag on the same tap.
            onMouseDown={(e) => e.stopPropagation()}
            onTouchStart={(e) => e.stopPropagation()}
            onClick={() => setButtonsHidden(!buttonsHidden)}
            title={buttonsHidden ? "Show Send/Skip buttons" : "Hide Send/Skip buttons"}
            className="flex items-center gap-1 rounded-full border border-border-strong/60 bg-surface-2 px-2 py-0.5 font-mono-data text-[9px] text-muted-foreground hover:text-foreground"
          >
            {buttonsHidden ? <Eye className="size-3" /> : <EyeOff className="size-3" />}
            {buttonsHidden ? "Show" : "Hide"}
          </button>
        </div>
      )}
      {!collapsed && children}
    </div>
  );
}

/* ----------------------------- Header ----------------------------- */

function Header({
  theme, onToggleTheme, onClearAll, totalRows, processedRows, onHeaderTap, dragUnlocked,
  homepageMode, onToggleMode, onOpenDrawer,
}: {
  theme: "dark" | "light";
  onToggleTheme: () => void;
  onClearAll: () => void;
  totalRows: number;
  processedRows: number;
  onHeaderTap: () => void;
  dragUnlocked: boolean;
  homepageMode: HomepageMode;
  onToggleMode: (m: HomepageMode) => void;
  onOpenDrawer: () => void;
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
              title="Double-tap to unlock the Send button"
              className={`cursor-pointer select-none truncate text-sm font-semibold leading-tight sm:text-base ${
                dragUnlocked ? "text-amber-glow" : ""
              }`}
            >
              <span className="text-sky-glow">Outreach Lab</span>
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

        {/* Mode toggle */}
        <div className="flex rounded-lg border border-border-strong/60 overflow-hidden shrink-0">
          <button
            type="button"
            onClick={() => onToggleMode("research")}
            className={`px-2.5 py-1.5 font-mono-data text-[10px] uppercase tracking-wider transition-colors ${
              homepageMode === "research"
                ? "bg-sky-glow text-black"
                : "bg-transparent text-muted-foreground hover:text-foreground"
            }`}
          >
            Research
          </button>
          <button
            type="button"
            onClick={() => onToggleMode("queue")}
            className={`px-2.5 py-1.5 font-mono-data text-[10px] uppercase tracking-wider transition-colors ${
              homepageMode === "queue"
                ? "bg-sky-glow text-black"
                : "bg-transparent text-muted-foreground hover:text-foreground"
            }`}
          >
            Queue
          </button>
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

        <Button variant="ghost" size="icon" onClick={onOpenDrawer} aria-label="Open settings">
          <Menu className="size-4" />
        </Button>
      </div>
    </header>
  );
}

/* --------------------------- Ingest panel --------------------------- */

function IngestPanel({
  parsing, progress, onFile, headers, totalRows, processedRows,
  targetEmailHeader, onTargetEmailHeader, domainHeader, onDomainHeader, countryHeader, onCountryHeader,
}: {
  parsing: boolean;
  progress: number;
  onFile: (f: File) => void;
  headers: string[];
  totalRows: number;
  processedRows: number;
  targetEmailHeader: string;
  onTargetEmailHeader: (v: string) => void;
  domainHeader: string;
  onDomainHeader: (v: string) => void;
  countryHeader: string;
  onCountryHeader: (v: string) => void;
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
          <div className="mt-2 grid gap-2 sm:grid-cols-[200px_1fr] sm:items-center">
            <Label className="text-xs uppercase tracking-wider text-muted-foreground">
              Domain / store column
            </Label>
            <select
              value={domainHeader}
              onChange={(e) => onDomainHeader(e.target.value)}
              className="h-9 w-full rounded-md border border-border-strong/70 bg-surface-2 px-2 font-mono-data text-sm outline-none focus:glow-sky"
            >
              <option value="">-- Not set --</option>
              {headers.map((h) => (
                <option key={h} value={h}>{h}</option>
              ))}
            </select>
          </div>
          <div className="mt-2 grid gap-2 sm:grid-cols-[200px_1fr] sm:items-center">
            <Label className="text-xs uppercase tracking-wider text-muted-foreground">
              Country column
            </Label>
            <select
              value={countryHeader}
              onChange={(e) => onCountryHeader(e.target.value)}
              className="h-9 w-full rounded-md border border-border-strong/70 bg-surface-2 px-2 font-mono-data text-sm outline-none focus:glow-sky"
              title="Used by the Prioritize country filter and the country badge in the send preview"
            >
              <option value="">-- Not set --</option>
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

/* --------------------------- Email Lookup Panel --------------------------- */

// Search the loaded CSV by email address and show every field for the
// matching row(s) as tap-to-copy chips. Lives only in the drawer so it
// never touches the main Queue/Research interface.
function LeadLookupPanel({ rows, headers }: { rows: Row[]; headers: string[] }) {
  const [query, setQuery] = useState("");
  const [copiedKey, setCopiedKey] = useState<string | null>(null);

  const copyValue = async (key: string, value: string) => {
    if (!value) return;
    try {
      await navigator.clipboard.writeText(value);
      setCopiedKey(key);
      setTimeout(() => setCopiedKey((k) => (k === key ? null : k)), 1200);
    } catch {
      toast.error("Clipboard failed");
    }
  };

  const results = useMemo(() => {
    const q = query.trim().toLowerCase();
    if (!q) return [];
    // Match against every cell, not just a guessed email column — some
    // lists keep a second address in a different field, and colon- or
    // semicolon-separated multi-email cells should still match on either half.
    const matches: number[] = [];
    for (let i = 0; i < rows.length; i++) {
      const row = rows[i];
      const hit = headers.some((h) => {
        const cell = String(row[h] ?? "").toLowerCase();
        if (!cell.includes("@")) return false;
        return cell.split(/[,;:\s]+/).some((piece) => piece.includes(q)) || cell.includes(q);
      });
      if (hit) matches.push(i);
      if (matches.length >= 25) break;
    }
    return matches;
  }, [query, rows, headers]);

  return (
    <div className="space-y-3">
      <p className="font-mono-data text-[10px] leading-relaxed text-muted-foreground">
        Search the loaded CSV by email address to pull up everything known about that lead, without scrolling the queue.
      </p>
      <div className="relative">
        <Search className="pointer-events-none absolute left-2.5 top-1/2 size-3.5 -translate-y-1/2 text-muted-foreground" />
        <Input
          value={query}
          onChange={(e) => setQuery(e.target.value)}
          placeholder="Search by email address…"
          className="pl-8 font-mono-data text-xs"
        />
      </div>

      {query.trim() && results.length === 0 && (
        <p className="font-mono-data text-[11px] text-muted-foreground">No rows match that email.</p>
      )}

      {results.length > 0 && (
        <div className="space-y-3">
          {results.length >= 25 && (
            <p className="font-mono-data text-[10px] text-amber-glow">Showing first 25 matches — narrow your search for more.</p>
          )}
          {results.map((idx) => {
            const row = rows[idx];
            return (
              <div key={idx} className="rounded-lg border border-border-strong/60 bg-surface-2 p-3 space-y-1.5">
                {headers.map((h) => {
                  const value = String(row[h] ?? "");
                  if (!value) return null;
                  const cellKey = `${idx}:${h}`;
                  return (
                    <button
                      key={h}
                      type="button"
                      onClick={() => copyValue(cellKey, value)}
                      className="flex w-full items-start gap-2 rounded-md px-1.5 py-1 text-left hover:bg-surface-1"
                      title="Tap to copy"
                    >
                      <span className="mt-0.5 shrink-0 font-mono-data text-[9px] uppercase tracking-wider text-muted-foreground w-20 truncate">
                        {h}
                      </span>
                      <span className="min-w-0 flex-1 break-all font-mono-data text-[11px] text-foreground">
                        {value}
                      </span>
                      {copiedKey === cellKey ? (
                        <CheckCircle className="size-3.5 shrink-0 text-emerald-400" />
                      ) : (
                        <Copy className="size-3.5 shrink-0 text-muted-foreground/50" />
                      )}
                    </button>
                  );
                })}
              </div>
            );
          })}
        </div>
      )}
    </div>
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
  // This is a ONE-OFF: it points at exactly one row, and gets cleared again
  // after that single send/skip, returning to normal queue order.
  const [activeOverride, setActiveOverride] = useState<number | null>(null);
  const [jumpInput, setJumpInput] = useState<string>("");
  // "Start from row #" — unlike activeOverride, this is STICKY: once set,
  // every subsequent "next pending" lookup is restricted to rows at or
  // after this index, so sending continues forward from here instead of
  // snapping back to the lowest-index pending row. Stays active until the
  // user explicitly clears it (or every row from here on is processed).
  const [queueStartFrom, setQueueStartFrom] = useState<number | null>(null);
  const [startFromInput, setStartFromInput] = useState<string>("");
  // Apply external "Restore previous session" jump
  useEffect(() => {
    if (resumeTarget !== null) {
      setActiveOverride(resumeTarget);
      setJumpInput(String(resumeTarget));
      onConsumeResume();
    }
  }, [resumeTarget, onConsumeResume]);
  const firstPendingFromStart = useMemo(() => {
    if (queueStartFrom === null) return undefined;
    for (let i = queueStartFrom; i < state.rows.length; i++) {
      if ((state.rowStates[i] ?? "pending") === "pending") return i;
    }
    return undefined; // everything from the start point onward is done
  }, [queueStartFrom, state.rows.length, state.rowStates]);
  const nextPendingIndex =
    activeOverride !== null && state.rows[activeOverride]
      ? activeOverride
      : queueStartFrom !== null
        ? firstPendingFromStart
        : firstPendingIndex;
  const pendingCount = state.rows.length - processedCount;
  const priorityMatchCount = useMemo(() => {
    const domain = state.priorityDomain.trim().toLowerCase().replace(/^@/, "");
    const country = state.priorityCountry.trim().toLowerCase();
    if (!domain && !country) return 0;
    let count = 0;
    for (let i = 0; i < state.rows.length; i++) {
      if (state.rowStates[i] === "processed" || state.rowStates[i] === "skipped") continue;
      let domainMatch = true, countryMatch = true;
      if (domain) {
        const email = String(state.rows[i]?.[state.targetEmailHeader] ?? "").toLowerCase();
        domainMatch = email.split(/[,;:\s]+/).some((piece) => piece.endsWith(`@${domain}`));
      }
      if (country) {
        const cell = String(state.rows[i]?.[state.countryHeader] ?? "").trim().toLowerCase();
        countryMatch = cell === country;
      }
      const isMatch = state.priorityMatchMode === "any" && domain && country
        ? domainMatch || countryMatch
        : domainMatch && countryMatch;
      if (isMatch) count++;
    }
    return count;
  }, [state.priorityDomain, state.priorityCountry, state.priorityMatchMode, state.rows, state.rowStates, state.targetEmailHeader, state.countryHeader]);

  // Distinct country values actually present in the loaded CSV, for the
  // dropdown — pulled from the data itself rather than a fixed list, since
  // exact spelling/format (e.g. "US" vs "United States") varies by list.
  const availableCountries = useMemo(() => {
    if (!state.countryHeader) return [];
    const set = new Set<string>();
    for (const row of state.rows) {
      const v = String(row[state.countryHeader] ?? "").trim();
      if (v) set.add(v);
    }
    return Array.from(set).sort((a, b) => a.localeCompare(b));
  }, [state.rows, state.countryHeader]);
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

  // The fully token-substituted HTML — same string that drives the Live
  // Preview iframe, and what "Copy template" copies (never the raw
  // {token} source in the editor above).
  const liveRenderedHtml = useMemo(
    () => autoFormatHtml(renderTemplate(activeTemplate.html, previewRow ?? sampleRow, undefined, state.sendCounter)),
    [activeTemplate.html, previewRow, sampleRow, state.sendCounter]
  );

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
      <CollapsibleSection
        title="Active Template & Rotation"
        icon={<Shuffle className="size-3.5 text-sky-glow" />}
        defaultOpen
        badge={<span className="font-mono-data text-[10px] text-muted-foreground">{state.templates.length} total</span>}
      >
        <div className="space-y-2">
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
      </CollapsibleSection>

      <CollapsibleSection
        title="Spam & Link Health"
        icon={<ShieldAlert className="size-3.5 text-amber-glow" />}
        defaultOpen={false}
      >
        <SpamHealthCheck
          subject={activeTemplate.subject}
          body={state.htmlMode ? activeTemplate.html : activeTemplate.body}
          html={state.htmlMode ? activeTemplate.html : ""}
        />
      </CollapsibleSection>

      <CollapsibleSection
        title="Subject & Sender Settings"
        icon={<Mail className="size-3.5 text-sky-glow" />}
        defaultOpen
      >
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
      </CollapsibleSection>
      {state.htmlMode ? (
        <>
          <CollapsibleSection
            title="HTML Code Template"
            icon={<Code2 className="size-3.5 text-amber-glow" />}
            defaultOpen
          >
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
          </CollapsibleSection>

          {/* Live Preview stays outside the collapsible section above and
              always renders, regardless of whether the HTML editor is
              expanded or collapsed — this is what you're about to send. */}
          <div>
            <div className="mb-2 flex items-center justify-between gap-2">
              <span className="flex items-center gap-2">
                <Eye className="size-3.5 text-muted-foreground" />
                <span className="font-mono-data text-[10px] uppercase tracking-wider text-muted-foreground">
                  Live preview
                </span>
              </span>
              <button
                type="button"
                onClick={async () => {
                  try {
                    await navigator.clipboard.writeText(liveRenderedHtml);
                    toast.success("Rendered HTML copied");
                  } catch {
                    toast.error("Clipboard failed");
                  }
                }}
                className="flex items-center gap-1 rounded-md border border-border-strong/60 bg-surface-2 px-2 py-1 font-mono-data text-[10px] text-muted-foreground hover:text-foreground"
                title="Copy the fully rendered HTML (tokens already filled in), not the raw {token} template"
              >
                <Copy className="size-3" /> Copy template
              </button>
            </div>
            <div className="overflow-hidden rounded-lg border border-border-strong/60 bg-white">
              <iframe
                title="HTML preview"
                sandbox=""
                srcDoc={`<!doctype html><html><body style="margin:0;padding:12px;font-family:system-ui">${liveRenderedHtml}</body></html>`}
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
        <div className="flex items-center gap-1.5">
          <label className="font-mono-data text-[10px] uppercase tracking-wider text-muted-foreground">
            Start from row #
          </label>
          <Input
            type="number"
            min={0}
            max={Math.max(0, state.rows.length - 1)}
            value={startFromInput}
            onChange={(e) => {
              const v = e.target.value;
              setStartFromInput(v);
              if (v === "") { setQueueStartFrom(null); return; }
              const n = Number(v);
              if (Number.isFinite(n) && n >= 0 && n < state.rows.length) {
                setQueueStartFrom(n);
              }
            }}
            className="h-8 w-20 font-mono-data text-xs"
            placeholder="0"
            disabled={state.rows.length === 0}
            title="Unlike Jump to row, this keeps the queue moving forward from here on every send."
          />
          {queueStartFrom !== null && (
            <button
              type="button"
              onClick={() => { setQueueStartFrom(null); setStartFromInput(""); }}
              className="rounded-md border border-sky-glow/40 bg-sky-glow/10 px-2 py-1 font-mono-data text-[10px] text-sky-glow hover:text-foreground"
              title="Clear starting point · resume normal queue"
            >
              <RotateCcw className="inline size-3" /> from #{queueStartFrom}
            </button>
          )}
        </div>
        <div className="flex flex-wrap items-center gap-1.5">
          <label className="font-mono-data text-[10px] uppercase tracking-wider text-muted-foreground">
            Prioritize domain
          </label>
          <Input
            value={state.priorityDomain}
            onChange={(e) => patch({ priorityDomain: e.target.value })}
            placeholder="e.g. gmail.com"
            className="h-8 w-32 font-mono-data text-xs"
            title="Pending leads whose email ends with this domain move to the front of the queue and get sent first"
          />

          <label className="font-mono-data text-[10px] uppercase tracking-wider text-muted-foreground ml-1">
            Country
          </label>
          <select
            value={state.priorityCountry}
            onChange={(e) => patch({ priorityCountry: e.target.value })}
            disabled={availableCountries.length === 0}
            className="h-8 rounded-md border border-border-strong/70 bg-surface-2 px-2 font-mono-data text-xs outline-none focus:glow-sky disabled:opacity-50"
            title={availableCountries.length === 0 ? "No country column detected in this CSV" : "Pending leads from this country move to the front of the queue and get sent first"}
          >
            <option value="">-- Any --</option>
            {availableCountries.map((c) => (
              <option key={c} value={c}>{c}</option>
            ))}
          </select>

          {state.priorityDomain.trim() && state.priorityCountry.trim() && (
            <div className="flex items-center gap-1 rounded-md border border-border-strong/60 bg-surface-2 p-0.5 ml-1">
              <button
                type="button"
                onClick={() => patch({ priorityMatchMode: "all" })}
                title="Only prioritize leads matching BOTH the domain and the country"
                className={`rounded px-1.5 py-0.5 font-mono-data text-[10px] ${state.priorityMatchMode === "all" ? "bg-sky-glow/20 text-sky-glow" : "text-muted-foreground"}`}
              >
                Both
              </button>
              <button
                type="button"
                onClick={() => patch({ priorityMatchMode: "any" })}
                title="Prioritize leads matching EITHER the domain or the country"
                className={`rounded px-1.5 py-0.5 font-mono-data text-[10px] ${state.priorityMatchMode === "any" ? "bg-sky-glow/20 text-sky-glow" : "text-muted-foreground"}`}
              >
                Either
              </button>
            </div>
          )}

          {(state.priorityDomain.trim() || state.priorityCountry.trim()) && (
            <>
              <span className="font-mono-data text-[10px] text-sky-glow">
                {priorityMatchCount} match{priorityMatchCount === 1 ? "" : "es"}
              </span>
              <button
                type="button"
                onClick={() => patch({ priorityDomain: "", priorityCountry: "" })}
                className="rounded-md border border-sky-glow/40 bg-sky-glow/10 px-2 py-1 font-mono-data text-[10px] text-sky-glow hover:text-foreground"
                title="Clear priority · resume normal queue order"
              >
                <RotateCcw className="inline size-3" /> clear
              </button>
            </>
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
            {queueStartFrom !== null
              ? `All rows from #${queueStartFrom} onward are processed.`
              : "All rows processed."}
          </p>
        ) : (
          <NextRowPreview
            rowIndex={nextPendingIndex}
            row={state.rows[nextPendingIndex]}
            targetEmailHeader={state.targetEmailHeader}
            countryHeader={state.countryHeader}
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
  rowIndex, row, targetEmailHeader, countryHeader, subjectTpl, bodyTpl, htmlMode, htmlTpl, onSend, onSkip, isResend, ai, spinSeed,
}: {
  rowIndex: number;
  row: Row | undefined;
  targetEmailHeader: string;
  countryHeader: string;
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
  // ai_insight token removed — tokens are now filled purely from CSV columns
  const extras: Record<string, string> = {};

  const rawRecipients = row?.[targetEmailHeader] ?? "";
  const toAddr = cleanEmails(rawRecipients);
  const subject = renderTemplate(subjectTpl, row, extras, spinSeed);
  const body = renderTemplate(bodyTpl, row, extras, spinSeed);
  // IMPORTANT: this is the preview-iframe value below — it must stay
  // pixel-free, or just scrolling to this row would silently "open" it.
  const renderedHtml = autoFormatHtml(renderTemplate(htmlTpl, row, extras, spinSeed));
  const plainHref = toAddr
    ? buildMailto(toAddr, { subject, body })
    : "";
  const htmlHref = toAddr
    ? buildMailto(toAddr, { subject })
    : "";
  const sendHtml = async () => {
    if (!toAddr) return;
    // Snapshot the current row's mailto BEFORE advancing the queue,
    // so re-render from onSend() can't swap in the next row's link.
    const hrefSnapshot = htmlHref;
    const outgoingHtml = renderedHtml;
    try {
      const blobHtml = new Blob([outgoingHtml], { type: "text/html" });
      const blobText = new Blob([outgoingHtml.replace(/<[^>]+>/g, "")], { type: "text/plain" });
      if ("ClipboardItem" in window && navigator.clipboard?.write) {
        await navigator.clipboard.write([
          new ClipboardItem({ "text/html": blobHtml, "text/plain": blobText }),
        ]);
      } else {
        await navigator.clipboard.writeText(outgoingHtml);
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
          {countryHeader && row?.[countryHeader] && (
            <span className="ml-2 rounded border border-border-strong/60 bg-bg-app px-1.5 py-0.5 text-[10px] text-muted-foreground">
              {row[countryHeader]}
            </span>
          )}
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

/* ------------------------ Collapsible Section ------------------------ */

/**
 * Generic accordion-style wrapper used to keep config/diagnostic panels
 * out of the way until the user actually wants them — tap the header to
 * expand/collapse. `badge` renders inline next to the chevron (e.g. a
 * count) so a collapsed section can still surface a glance-able number.
 */
function CollapsibleSection({
  title, icon, defaultOpen = false, badge, children,
}: {
  title: string;
  icon?: React.ReactNode;
  defaultOpen?: boolean;
  badge?: React.ReactNode;
  children: React.ReactNode;
}) {
  const [open, setOpen] = useState(defaultOpen);
  return (
    <div className="mb-4 overflow-hidden rounded-xl border border-border-strong/70 bg-surface-1">
      <button
        type="button"
        onClick={() => setOpen((o) => !o)}
        aria-expanded={open}
        className="flex w-full items-center justify-between gap-2 px-4 py-3 text-left active:bg-surface-2"
      >
        <span className="flex items-center gap-2 font-mono-data text-xs uppercase tracking-wider text-foreground">
          {icon}
          {title}
        </span>
        <span className="flex items-center gap-2">
          {badge}
          <ChevronDown
            className={`size-4 shrink-0 text-muted-foreground transition-transform duration-150 ${open ? "rotate-180" : ""}`}
          />
        </span>
      </button>
      {open && <div className="space-y-4 border-t border-border-strong/60 p-4">{children}</div>}
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
            {templates.map((t, idx) => (
              <TemplateEditorCard
                key={t.id}
                t={t}
                idx={idx}
                isActive={t.id === activeTemplateId}
                onSelect={onSelect}
                onDelete={onDelete}
                onUpdate={onUpdate}
              />
            ))}
          </div>
        </div>
      )}
    </section>
  );
}

// Holds its own local draft of a template's fields so typing is instant and
// never touches the shared app state (which was re-rendering the whole
// app — including the live preview iframe — on every keystroke). Nothing
// propagates out to the rest of the app until you hit "Sync" or leave the
// field, so typing doesn't fight the preview for CPU.
function TemplateEditorCard({
  t, idx, isActive, onSelect, onDelete, onUpdate,
}: {
  t: TemplateItem;
  idx: number;
  isActive: boolean;
  onSelect: (id: string) => void;
  onDelete: (id: string) => void;
  onUpdate: (id: string, p: Partial<TemplateItem>) => void;
}) {
  const [draft, setDraft] = useState({ name: t.name, subject: t.subject, body: t.body, html: t.html });
  const [dirty, setDirty] = useState(false);

  // Re-sync from props if the template changes from outside (switching
  // templates, reset, import, etc.) — but never clobber an unsynced edit
  // you're still mid-typing.
  useEffect(() => {
    if (!dirty) setDraft({ name: t.name, subject: t.subject, body: t.body, html: t.html });
  }, [t.id, t.name, t.subject, t.body, t.html, dirty]);

  const sync = () => {
    onUpdate(t.id, draft);
    setDirty(false);
  };

  const set = (patch: Partial<typeof draft>) => {
    setDraft((d) => ({ ...d, ...patch }));
    setDirty(true);
  };

  return (
    <div
      className={`rounded-lg border p-3 ${
        isActive ? "border-sky-glow/60 bg-sky-glow/5" : "border-border-strong/40 bg-surface-2"
      }`}
    >
      <div className="mb-2 flex flex-wrap items-center gap-2">
        <span className="font-mono-data text-[10px] uppercase text-muted-foreground">#{idx + 1}</span>
        <Input
          value={draft.name}
          onChange={(e) => set({ name: e.target.value })}
          onBlur={sync}
          className="h-7 max-w-[200px] font-sans text-sm"
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
        {dirty && (
          <Button size="sm" onClick={sync} className="h-7 glow-sky">
            <RefreshCw className="size-3" /> Sync
          </Button>
        )}
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
          value={draft.subject}
          onChange={(e) => set({ subject: e.target.value })}
          onBlur={sync}
          placeholder="Subject"
          className="h-8 font-sans text-sm"
        />
        <Textarea
          value={draft.body}
          onChange={(e) => set({ body: e.target.value })}
          onBlur={sync}
          placeholder="Plain-text body"
          rows={3}
          className="font-sans text-sm leading-relaxed"
        />
        <Textarea
          value={draft.html}
          onChange={(e) => set({ html: e.target.value })}
          onBlur={sync}
          placeholder="HTML body"
          rows={3}
          className="font-sans text-sm leading-relaxed"
          spellCheck={false}
        />
      </div>
      {dirty && (
        <p className="mt-1.5 font-mono-data text-[10px] text-amber-glow">
          Unsynced changes — tap Sync (or tap away) to update the live preview.
        </p>
      )}
    </div>
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

/* ===================== end of file ===================== */


