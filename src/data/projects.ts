// Training template projects: a prompt, the slide plan and the design seed.
// Slides are derived by fixed rules from the project name, layouts and seed (no AI).
import { LAYOUTS, defaultLayouts, type AvatarSpace, type Layout } from "@/data/design";
import type { SlideContent } from "@/data/fill";

export interface Project {
  id: string;
  prompt: string;
  name: string;
  slideCount: number;
  avatar: AvatarSpace;
  seed: number;
  layouts: Layout[] | null;
  /** Text poured in from an uploaded document, one entry per slide; absent for blank templates. */
  content?: SlideContent[] | null;
  /** File name of the document the text came from. */
  source?: string | null;
  updatedAt: number;
}

export const AVATAR_OPTIONS: AvatarSpace[] = ["Right", "Left", "None"];
export const SLIDE_COUNTS = Array.from({ length: 18 }, (_, i) => i + 3);

export function newId(): string {
  return Math.random().toString(36).slice(2, 10) + Date.now().toString(36);
}

// ---------- Prompt parsing ----------

const WORDS: Record<string, number> = {
  three: 3, four: 4, five: 5, six: 6, seven: 7, eight: 8, nine: 9, ten: 10, eleven: 11, twelve: 12,
  thirteen: 13, fourteen: 14, fifteen: 15, sixteen: 16, seventeen: 17, eighteen: 18, nineteen: 19, twenty: 20,
};

export interface ParsedPrompt {
  name: string | null;
  count: number | null;
  avatar: AvatarSpace | null;
}

const titleCase = (s: string) =>
  s
    .split(/\s+/)
    .map((w, i) => {
      if (w.length > 1 && w === w.toUpperCase() && /\p{L}/u.test(w)) return w;
      const l = w.toLowerCase();
      return i > 0 && ["and", "of", "for", "the", "to", "in", "on", "a"].includes(l) ? l : l.charAt(0).toUpperCase() + l.slice(1);
    })
    .join(" ");

export function parsePrompt(prompt: string): ParsedPrompt {
  const text = prompt.replace(/\s+/g, " ").trim();
  if (!text) return { name: null, count: null, avatar: null };

  let count: number | null = null;
  const num = text.match(/(\d{1,2})\s*(?:[a-z-]+\s+){0,2}(?:slides?|pages?|screens?)\b/i);
  const word = text.match(new RegExp(`\\b(${Object.keys(WORDS).join("|")})\\s+(?:[a-z-]+\\s+){0,2}(?:slides?|pages?|screens?)\\b`, "i"));
  if (num) count = parseInt(num[1], 10);
  else if (word) count = WORDS[word[1].toLowerCase()];
  if (count !== null) count = Math.max(3, Math.min(20, count));

  let avatar: AvatarSpace | null = null;
  if (/\b(?:no|without)\s+avatar\b/i.test(text)) avatar = "None";
  else if (/\bavatar\s+(?:on\s+the\s+)?left\b|\bleft\s+avatar\b/i.test(text)) avatar = "Left";
  else if (/\bavatar\s+(?:on\s+the\s+)?right\b|\bright\s+avatar\b/i.test(text)) avatar = "Right";

  let name: string | null = null;
  const before = text.match(/^(.+?)\s*(?:training\s+)?(?:project|course|module|programme|program)\b/i);
  const about = text.match(/\b(?:for|on|about)\s+(?:an?\s+|the\s+)?(.+?)\s*(?:training\s+)?(?:project|course|module|programme|program)?(?=[,.;]|\s+(?:with|create|make|build|design|generate|in|using)\b|$)/i);
  if (before && !/^(?:create|make|build|design|generate)\b/i.test(before[1])) name = before[1];
  else if (about) name = about[1];
  else name = text.split(/[,.;]|\s+(?:create|make|build|design|generate)\b/i)[0];
  name = name.replace(/^(?:create|make|build|design|generate)\s+(?:an?\s+|the\s+)?/i, "").replace(/[^\p{L}\p{N}&' -]/gu, "").trim();
  if (/^\d|slides?$/i.test(name) || !name) name = null;
  return { name: name ? titleCase(name.split(" ").slice(0, 8).join(" ")) : null, count, avatar };
}

// ---------- Seeds ----------

function seed(prompt: string, name: string, slideCount: number, avatar: AvatarSpace, built: boolean, ago: number): Project {
  return { id: newId(), prompt, name, slideCount, avatar, seed: 0, layouts: built ? defaultLayouts(slideCount) : null, updatedAt: Date.now() - ago };
}

export function seedProjects(): Project[] {
  return [
    seed("Maternity leave project, create 8 template slides", "Maternity Leave", 8, "Right", true, 3_600_000),
    seed("Phishing awareness training, 10 slides, avatar on the left", "Phishing Awareness", 10, "Left", true, 86_400_000),
    seed("Quarterly check-ins for people managers, 6 slides, no avatar", "Quarterly Check-ins for People Managers", 6, "None", false, 3 * 86_400_000),
  ];
}

export function blankProject(): Project {
  return { id: newId(), prompt: "", name: "", slideCount: 8, avatar: "Right", seed: 0, layouts: null, updatedAt: Date.now() };
}

// ---------- Browser persistence ----------

const STORE_VERSION = 2;

function isProject(v: unknown): v is Project {
  if (!v || typeof v !== "object") return false;
  const p = v as Project;
  return (
    typeof p.id === "string" &&
    typeof p.prompt === "string" &&
    typeof p.name === "string" &&
    typeof p.slideCount === "number" &&
    AVATAR_OPTIONS.includes(p.avatar) &&
    typeof p.seed === "number" &&
    (p.layouts === null || (Array.isArray(p.layouts) && p.layouts.every((l) => (LAYOUTS as readonly string[]).includes(l)))) &&
    (p.content === undefined || p.content === null || (Array.isArray(p.content) && p.content.every((c) => !!c && typeof c === "object" && typeof c.notes === "string" && !!c.blocks && typeof c.blocks === "object"))) &&
    (p.source === undefined || p.source === null || typeof p.source === "string")
  );
}

/** Version 1 stored video-script projects; keep their name and prompt and start a fresh slide plan. */
function upgradeV1(v: unknown): Project | null {
  if (!v || typeof v !== "object") return null;
  const o = v as { id?: unknown; prompt?: unknown; settings?: { title?: unknown; presenter?: unknown }; scenes?: unknown; updatedAt?: unknown };
  if (typeof o.id !== "string" || typeof o.prompt !== "string" || !o.settings) return null;
  const count = Array.isArray(o.scenes) ? Math.max(3, Math.min(20, o.scenes.length)) : 8;
  return {
    id: o.id,
    prompt: o.prompt,
    name: typeof o.settings.title === "string" ? o.settings.title : "",
    slideCount: count,
    avatar: o.settings.presenter === "Voiceover only" ? "None" : "Right",
    seed: 0,
    layouts: Array.isArray(o.scenes) ? defaultLayouts(count) : null,
    updatedAt: typeof o.updatedAt === "number" ? o.updatedAt : Date.now(),
  };
}

export type LoadResult =
  | { status: "ok"; projects: Project[] }
  | { status: "empty" }
  | { status: "invalid" }
  | { status: "unavailable" };

export function loadProjects(key: string): LoadResult {
  let raw: string | null;
  try {
    raw = window.localStorage.getItem(key);
  } catch {
    return { status: "unavailable" };
  }
  if (raw === null) return { status: "empty" };
  try {
    const data = JSON.parse(raw);
    if (data?.v === STORE_VERSION && Array.isArray(data.projects) && data.projects.every(isProject)) {
      return { status: "ok", projects: data.projects };
    }
    const v1 = Array.isArray(data) ? data : data?.v === 1 ? data.projects : null;
    if (Array.isArray(v1)) {
      const up = v1.map(upgradeV1);
      if (up.every((p): p is Project => p !== null)) return { status: "ok", projects: up };
    }
    return { status: "invalid" };
  } catch {
    return { status: "invalid" };
  }
}

export function saveProjects(key: string, projects: Project[]): boolean {
  try {
    window.localStorage.setItem(key, JSON.stringify({ v: STORE_VERSION, projects }));
    return true;
  } catch {
    return false;
  }
}
