// Template design system: palettes, motifs and slide layouts on a 1600×900 canvas.
// Every project gets a design derived by fixed rules from its name and design seed.

export const W = 1600;
export const H = 900;

/** Solid colour (hex, optional alpha) or a two-stop linear gradient; ang uses PowerPoint degrees (0 = left→right, 90 = top→bottom). */
export type Fill = { c: string; a?: number } | { g: [string, string]; ang: number };
export type ShapeKind = "rect" | "round" | "ellipse";
export type Role = "bg" | "deco" | "text" | "image";

export interface Shape {
  t: ShapeKind;
  x: number;
  y: number;
  w: number;
  h: number;
  fill?: Fill;
  line?: { c: string; w: number; a?: number };
  rot?: number;
  /** Corner rounding for "round" shapes, as a fraction of the shorter side. */
  rad?: number;
  role: Role;
  label?: string;
  /** Font size (pt) and colour used when text is typed into the block. */
  font?: { size: number; color: string; bold?: boolean; align?: "l" | "ctr" };
  /** Filled-in text (from an imported document); blocks stay empty when absent. */
  text?: string;
}

export interface Palette {
  name: string;
  dark: string;
  dark2: string;
  primary: string;
  accent: string;
  light: string;
  ink: string;
  keywords: string[];
}

export const PALETTES: Palette[] = [
  { name: "Blush & Plum", dark: "3B1F3A", dark2: "6B3460", primary: "8E4A7E", accent: "E8A9A0", light: "FBF5F3", ink: "2E1A2D", keywords: ["maternity", "paternity", "parental", "family", "adoption", "childcare", "care"] },
  { name: "Sage & Ivory", dark: "24382F", dark2: "3F5E4F", primary: "5E8670", accent: "D8B98A", light: "F7F6F0", ink: "1E2B25", keywords: ["wellbeing", "wellness", "health", "mental", "benefit", "leave", "medical", "insurance"] },
  { name: "Midnight & Teal", dark: "0D1B2A", dark2: "1B3A4B", primary: "1F6F78", accent: "3FC1C9", light: "F3F7F8", ink: "0D1B2A", keywords: ["security", "cyber", "phishing", "data", "privacy", "it", "technology", "gdpr"] },
  { name: "Navy & Gold", dark: "14213D", dark2: "22335C", primary: "2F4A84", accent: "D4A24C", light: "F8F6F1", ink: "14213D", keywords: ["leadership", "manager", "executive", "strategy", "finance", "performance", "governance"] },
  { name: "Forest & Brass", dark: "1C2E24", dark2: "2D4A3A", primary: "3E6B52", accent: "C49A5A", light: "F5F4EE", ink: "1C2E24", keywords: ["sustainability", "esg", "environment", "green", "climate", "property", "real estate"] },
  { name: "Graphite & Coral", dark: "1F2326", dark2: "353B40", primary: "4A5259", accent: "F07C5F", light: "F7F6F5", ink: "1F2326", keywords: ["sales", "customer", "marketing", "service", "client", "negotiation"] },
  { name: "Indigo & Sky", dark: "1A1A4B", dark2: "2E2E7A", primary: "4747A8", accent: "7FC8F8", light: "F5F6FB", ink: "1A1A4B", keywords: ["onboarding", "induction", "orientation", "new joiner", "welcome", "introduction"] },
  { name: "Charcoal & Amber", dark: "222222", dark2: "3A3A3A", primary: "565656", accent: "F2A541", light: "F8F7F4", ink: "222222", keywords: ["safety", "health and safety", "warehouse", "operations", "risk", "emergency", "fire"] },
  { name: "Aubergine & Rose", dark: "2D1B3D", dark2: "4A2C5E", primary: "6D4590", accent: "F2A7C3", light: "FAF6FB", ink: "2D1B3D", keywords: ["diversity", "inclusion", "dei", "culture", "belonging", "harassment", "posh", "respect"] },
  { name: "Ocean & Sand", dark: "0F2C3F", dark2: "1C4A63", primary: "2A6F8F", accent: "E3C08D", light: "F7F5F0", ink: "0F2C3F", keywords: ["travel", "expense", "policy", "compliance", "code of conduct", "ethics", "procurement"] },
];

export const MOTIFS = ["Arcs", "Diagonal", "Frame", "Ribbons", "Orbs", "Blocks"] as const;
export type Motif = (typeof MOTIFS)[number];
export const HEADERS = ["Accent bar", "Side marker", "Underline"] as const;
export type HeaderStyle = (typeof HEADERS)[number];

export interface Design {
  palette: Palette;
  motif: Motif;
  header: HeaderStyle;
  radius: number;
  heroAlign: "left" | "center";
}

function hash(s: string): number {
  let h = 2166136261;
  for (let i = 0; i < s.length; i++) {
    h ^= s.charCodeAt(i);
    h = Math.imul(h, 16777619);
  }
  return h >>> 0;
}

function rng(seed: number) {
  let s = seed || 1;
  return () => {
    s ^= s << 13;
    s ^= s >>> 17;
    s ^= s << 5;
    return ((s >>> 0) % 10000) / 10000;
  };
}

/** Palette is matched to the project topic by keyword on the first design; later seeds rotate it. */
export function makeDesign(name: string, seed: number): Design {
  const lower = ` ${name.toLowerCase()} `;
  const matched = PALETTES.findIndex((p) => p.keywords.some((k) => lower.includes(k)));
  const h = hash(`${lower}|${seed}`);
  const base = matched >= 0 ? matched : hash(lower) % PALETTES.length;
  const palette = PALETTES[(base + seed) % PALETTES.length];
  return {
    palette,
    motif: MOTIFS[(hash(lower) + seed * 7) % MOTIFS.length],
    header: HEADERS[(h >>> 3) % HEADERS.length],
    radius: [0.04, 0.1, 0.18][(h >>> 7) % 3],
    heroAlign: (h >>> 11) % 2 ? "center" : "left",
  };
}

// ---------- Layouts ----------

export const LAYOUTS = [
  "Title",
  "Agenda",
  "Section divider",
  "Image + text",
  "Three cards",
  "Process steps",
  "Key highlight",
  "Two columns",
  "Timeline",
  "Comparison",
  "Full image",
  "Summary checklist",
  "Knowledge check",
  "Closing",
] as const;
export type Layout = (typeof LAYOUTS)[number];

const MIDDLE: Layout[] = [
  "Agenda",
  "Section divider",
  "Image + text",
  "Three cards",
  "Process steps",
  "Key highlight",
  "Two columns",
  "Timeline",
  "Comparison",
  "Full image",
  "Knowledge check",
  "Summary checklist",
  "Section divider",
  "Image + text",
  "Three cards",
  "Two columns",
  "Process steps",
  "Key highlight",
];

export function defaultLayouts(count: number): Layout[] {
  const n = Math.max(2, Math.min(20, count));
  return ["Title", ...MIDDLE.slice(0, n - 2), "Closing"];
}

export type AvatarSpace = "Right" | "Left" | "None";

const HERO: Layout[] = ["Title", "Section divider", "Key highlight", "Closing"];
export const isHero = (l: Layout) => HERO.includes(l);

export function renderSlide(layout: Layout, d: Design, avatar: AvatarSpace, index: number): Shape[] {
  const p = d.palette;
  const hero = isHero(layout);
  const r = rng(hash(`${p.name}${d.motif}${index}`));
  const out: Shape[] = [];
  const add = (s: Shape) => out.push(s);

  // Content area, leaving room for the presenter avatar.
  const pad = 110;
  const av = avatar === "None" ? 0 : 470;
  const cx = avatar === "Left" ? av + 40 : pad;
  const cw = avatar === "None" ? W - pad * 2 : W - av - 40 - pad;
  const rad = d.radius;
  const onDark = { tint: "FFFFFF", a: 0.1 };
  const onLight = { tint: p.primary, a: 0.07 };
  const surf = hero ? onDark : onLight;
  const textColor = hero ? "FFFFFF" : p.ink;

  const text = (x: number, y: number, w: number, h: number, label: string, size: number, opts: Partial<Shape> = {}) =>
    add({
      t: "round",
      x,
      y,
      w,
      h,
      fill: { c: surf.tint, a: surf.a },
      role: "text",
      label,
      font: { size, color: textColor, bold: size >= 28, align: "l" },
      ...opts,
    });
  const image = (x: number, y: number, w: number, h: number, round = true) =>
    add({ t: round ? "round" : "rect", x, y, w, h, fill: { c: p.primary, a: hero ? 0.35 : 0.16 }, role: "image", label: "Image" });
  const accent = (x: number, y: number, w: number, h: number, a = 1) => add({ t: "rect", x, y, w, h, fill: { c: p.accent, a }, role: "deco" });
  const circle = (x: number, y: number, s: number, c: string, a = 1) => add({ t: "ellipse", x, y, w: s, h: s, fill: { c, a }, role: "deco" });

  // Background
  if (hero) {
    const g: [string, string] = layout === "Section divider" ? [p.primary, p.dark2] : [p.dark, p.dark2];
    add({ t: "rect", x: 0, y: 0, w: W, h: H, fill: { g, ang: 45 }, role: "bg" });
  } else {
    add({ t: "rect", x: 0, y: 0, w: W, h: H, fill: { c: p.light }, role: "bg" });
  }

  // Motif
  const m = d.motif;
  if (m === "Arcs") {
    if (hero) {
      circle(W - 620, -360, 980, p.accent, 0.16);
      circle(W - 360, 420, 640, "FFFFFF", 0.06);
    } else {
      circle(W - 230, -230, 460, p.primary, 0.07);
      circle(-160, H - 140, 300, p.accent, 0.14);
    }
  } else if (m === "Diagonal") {
    if (hero) {
      add({ t: "rect", x: W - 640, y: -300, w: 360, h: 1600, rot: 22, fill: { c: p.accent, a: 0.18 }, role: "deco" });
      add({ t: "rect", x: W - 330, y: -300, w: 120, h: 1600, rot: 22, fill: { c: "FFFFFF", a: 0.07 }, role: "deco" });
    } else {
      add({ t: "rect", x: W - 200, y: -260, w: 140, h: 700, rot: 30, fill: { c: p.accent, a: 0.22 }, role: "deco" });
    }
  } else if (m === "Frame") {
    if (hero) {
      add({ t: "rect", x: 44, y: 44, w: W - 88, h: H - 88, line: { c: p.accent, w: 2, a: 0.7 }, role: "deco" });
      accent(44, 44, 90, 8);
      accent(W - 134, H - 52, 90, 8);
    } else {
      accent(0, 0, W, 8);
      add({ t: "rect", x: 30, y: 30, w: W - 60, h: H - 60, line: { c: p.primary, w: 1, a: 0.18 }, role: "deco" });
    }
  } else if (m === "Ribbons") {
    if (hero) {
      accent(0, 0, 34, H);
      add({ t: "rect", x: 48, y: 0, w: 14, h: H, fill: { c: "FFFFFF", a: 0.16 }, role: "deco" });
      add({ t: "rect", x: 74, y: 0, w: 5, h: H, fill: { c: p.accent, a: 0.6 }, role: "deco" });
    } else {
      add({ t: "rect", x: 0, y: 0, w: 18, h: H, fill: { c: p.primary }, role: "deco" });
      accent(24, 0, 6, H);
    }
  } else if (m === "Orbs") {
    const n = hero ? 5 : 2;
    for (let i = 0; i < n; i++) {
      const s = hero ? 120 + r() * 340 : 90 + r() * 120;
      circle(W * (0.55 + r() * 0.45) - s / 2, H * (hero ? r() : 0.75 + r() * 0.25) - s / 2, s, i % 2 ? p.accent : hero ? "FFFFFF" : p.primary, hero ? 0.08 + r() * 0.12 : 0.08);
    }
  } else {
    const cols = hero ? 4 : 2;
    const size = hero ? 64 : 36;
    const gap = hero ? 18 : 12;
    const ox = W - cols * (size + gap) - (hero ? 70 : 50);
    const oy = hero ? H - cols * (size + gap) - 60 : 50;
    for (let i = 0; i < cols; i++)
      for (let j = 0; j < cols; j++)
        add({ t: "rect", x: ox + i * (size + gap), y: oy + j * (size + gap), w: size, h: size, fill: { c: (i + j) % 3 ? p.primary : p.accent, a: hero ? 0.12 + ((i * j) % 4) * 0.08 : 0.15 }, role: "deco" });
  }

  // Header block for content slides
  const header = (y = 90, w = Math.min(cw * 0.7, 820)) => {
    if (d.header === "Accent bar") accent(cx, y - 26, 96, 8);
    if (d.header === "Side marker") accent(cx - 28, y, 8, 86);
    text(cx, y, w, 86, "Slide title", 36);
    if (d.header === "Underline") add({ t: "rect", x: cx, y: y + 100, w: cw, h: 2, fill: { c: p.primary, a: 0.25 }, role: "deco" });
    return y + 130;
  };
  const footer = () => {
    if (hero) return;
    add({ t: "rect", x: cx, y: H - 62, w: cw, h: 1.5, fill: { c: p.primary, a: 0.2 }, role: "deco" });
    text(cx + cw - 70, H - 52, 70, 32, "No.", 12, { fill: undefined });
  };
  const k = (n: number) => Math.round(n);

  switch (layout) {
    case "Title": {
      const center = d.heroAlign === "center";
      const tw = Math.min(cw * 0.85, 960);
      const tx = center ? cx + (cw - tw) / 2 : cx;
      text(cx, 70, 200, 56, "Logo", 14);
      accent(center ? cx + cw / 2 - 60 : tx, 300, 120, 10);
      text(tx, 340, tw, 160, "Course title", 54, { font: { size: 54, color: "FFFFFF", bold: true, align: center ? "ctr" : "l" } });
      text(center ? cx + (cw - tw * 0.7) / 2 : tx, 530, tw * 0.7, 70, "Subtitle", 22, { font: { size: 22, color: "FFFFFF", align: center ? "ctr" : "l" } });
      text(center ? cx + cw / 2 - 160 : tx, H - 130, 320, 46, "Presenter / date", 14, { font: { size: 14, color: "FFFFFF", align: center ? "ctr" : "l" } });
      break;
    }
    case "Agenda": {
      const y0 = header();
      const rows = 4;
      const rh = k((H - y0 - 110) / rows);
      for (let i = 0; i < rows; i++) {
        const y = y0 + i * rh + 10;
        circle(cx, y, rh - 34, i === 0 ? p.accent : p.primary, i === 0 ? 1 : 0.85);
        text(cx + rh, y, cw * 0.75 - rh, rh - 34, `Agenda item ${i + 1}`, 22);
      }
      break;
    }
    case "Section divider": {
      add({ t: "ellipse", x: cx, y: 250, w: 170, h: 170, line: { c: p.accent, w: 4 }, role: "deco" });
      text(cx + 30, 300, 110, 70, "No.", 40, { fill: undefined, font: { size: 40, color: p.accent, bold: true, align: "ctr" } });
      text(cx, 470, Math.min(cw * 0.8, 900), 120, "Section title", 48);
      text(cx, 610, Math.min(cw * 0.6, 640), 60, "Section summary", 20);
      break;
    }
    case "Image + text": {
      const y0 = 90;
      const iw = k(cw * 0.46);
      image(cx, y0, iw, H - y0 - 110);
      const tx = cx + iw + 50;
      const tw = cw - iw - 50;
      if (d.header === "Accent bar") accent(tx, y0 - 26 + 40, 96, 8);
      text(tx, y0 + 40, tw, 86, "Slide title", 36);
      for (let i = 0; i < 3; i++) text(tx, y0 + 170 + i * 150, tw, 120, `Point ${i + 1}`, 18);
      break;
    }
    case "Three cards": {
      const y0 = header();
      const gap = 36;
      const w = k((cw - gap * 2) / 3);
      const h = H - y0 - 110;
      for (let i = 0; i < 3; i++) {
        const x = cx + i * (w + gap);
        add({ t: "round", x, y: y0, w, h, fill: { c: "FFFFFF" }, line: { c: p.primary, w: 1, a: 0.15 }, role: "deco" });
        accent(x, y0, w, 8, i === 1 ? 1 : 0.6);
        circle(x + 30, y0 + 40, 76, p.primary, 0.12);
        text(x + 30, y0 + 140, w - 60, 56, "Card heading", 22);
        text(x + 30, y0 + 210, w - 60, h - 240, "Card text", 16);
      }
      break;
    }
    case "Process steps": {
      const y0 = header();
      const n = 4;
      const step = k(cw / n);
      const cy = y0 + 60;
      add({ t: "rect", x: cx + 45, y: cy + 43, w: step * (n - 1), h: 4, fill: { c: p.primary, a: 0.3 }, role: "deco" });
      for (let i = 0; i < n; i++) {
        const x = cx + i * step;
        circle(x, cy, 90, i === n - 1 ? p.accent : p.primary);
        text(x + 18, cy + 18, 54, 54, "No.", 24, { fill: undefined, font: { size: 24, color: "FFFFFF", bold: true, align: "ctr" } });
        text(x, cy + 130, step - 30, 56, "Step heading", 20);
        text(x, cy + 200, step - 30, H - cy - 320, "Step detail", 15);
      }
      break;
    }
    case "Key highlight": {
      const w = Math.min(cw * 0.8, 900);
      accent(cx, 230, 14, 300);
      text(cx + 50, 230, w, 190, "Key number or quote", 60);
      text(cx + 50, 440, w * 0.75, 90, "Supporting line", 22);
      text(cx + 50, 560, 300, 44, "Source", 14);
      break;
    }
    case "Two columns": {
      const y0 = header();
      const gap = 50;
      const w = k((cw - gap) / 2);
      for (let i = 0; i < 2; i++) {
        const x = cx + i * (w + gap);
        accent(x, y0, 60, 6, i ? 0.6 : 1);
        text(x, y0 + 26, w, 60, "Column heading", 24);
        text(x, y0 + 106, w, H - y0 - 220, "Column text", 17);
      }
      break;
    }
    case "Timeline": {
      const y0 = header();
      const n = 5;
      const ly = y0 + (H - y0 - 110) / 2;
      add({ t: "rect", x: cx, y: ly - 2, w: cw, h: 4, fill: { c: p.primary, a: 0.35 }, role: "deco" });
      const step = cw / n;
      for (let i = 0; i < n; i++) {
        const x = k(cx + step * i + step / 2);
        circle(x - 16, ly - 16, 32, i % 2 ? p.primary : p.accent);
        const up = i % 2 === 0;
        text(k(x - step / 2 + 10), up ? ly - 180 : ly + 40, k(step - 20), 46, "Date", 16);
        text(k(x - step / 2 + 10), up ? ly - 124 : ly + 96, k(step - 20), 90, "Milestone", 15);
      }
      break;
    }
    case "Comparison": {
      const y0 = header();
      const gap = 40;
      const w = k((cw - gap) / 2);
      for (let i = 0; i < 2; i++) {
        const x = cx + i * (w + gap);
        add({ t: "round", x, y: y0, w, h: 74, fill: { c: i ? p.accent : p.primary }, role: "deco" });
        text(x + 24, y0 + 10, w - 48, 54, i ? "Option B" : "Option A", 22, { fill: undefined, font: { size: 22, color: i ? p.ink : "FFFFFF", bold: true, align: "l" } });
        for (let j = 0; j < 3; j++) text(x, y0 + 100 + j * 120, w, 100, "Comparison point", 16);
      }
      break;
    }
    case "Full image": {
      add({ t: "rect", x: 0, y: 0, w: W, h: H, fill: { c: p.primary, a: 0.22 }, role: "image", label: "Full-bleed image" });
      add({ t: "rect", x: 0, y: H - 300, w: W, h: 300, fill: { c: p.dark, a: 0.72 }, role: "deco" });
      text(cx, H - 210, Math.min(cw * 0.75, 900), 80, "Caption title", 34, { fill: { c: "FFFFFF", a: 0.12 }, font: { size: 34, color: "FFFFFF", bold: true } });
      text(cx, H - 115, Math.min(cw * 0.6, 720), 50, "Caption", 18, { fill: { c: "FFFFFF", a: 0.12 }, font: { size: 18, color: "FFFFFF" } });
      return out.map((s) => (s.t === "round" ? { ...s, rad } : s));
    }
    case "Summary checklist": {
      const y0 = header();
      const lw = k(cw * 0.58);
      for (let i = 0; i < 4; i++) {
        const y = y0 + i * 125;
        circle(cx, y + 18, 56, p.accent);
        add({ t: "rect", x: cx + 15, y: y + 44, w: 26, h: 6, rot: 0, fill: { c: "FFFFFF" }, role: "deco" });
        text(cx + 80, y, lw - 80, 92, `Takeaway ${i + 1}`, 18);
      }
      add({ t: "round", x: cx + lw + 40, y: y0, w: cw - lw - 40, h: H - y0 - 110, fill: { c: p.primary }, role: "deco" });
      text(cx + lw + 70, y0 + 40, cw - lw - 100, 70, "Key message", 24, { fill: { c: "FFFFFF", a: 0.12 }, font: { size: 24, color: "FFFFFF", bold: true } });
      text(cx + lw + 70, y0 + 130, cw - lw - 100, H - y0 - 280, "Key message detail", 16, { fill: { c: "FFFFFF", a: 0.12 }, font: { size: 16, color: "FFFFFF" } });
      break;
    }
    case "Knowledge check": {
      const y0 = header();
      text(cx, y0, cw, 100, "Question", 24);
      const gap = 30;
      const w = k((cw - gap) / 2);
      for (let i = 0; i < 4; i++) {
        const x = cx + (i % 2) * (w + gap);
        const y = y0 + 140 + Math.floor(i / 2) * 150;
        add({ t: "round", x, y, w, h: 120, fill: { c: "FFFFFF" }, line: { c: p.primary, w: 1.5, a: 0.3 }, role: "deco" });
        circle(x + 24, y + 32, 56, i === 0 ? p.accent : p.primary, i === 0 ? 1 : 0.2);
        text(x + 100, y + 22, w - 124, 76, `Option ${"ABCD"[i]}`, 18, { fill: undefined });
      }
      break;
    }
    case "Closing": {
      const center = d.heroAlign === "center";
      const tw = Math.min(cw * 0.8, 900);
      const tx = center ? cx + (cw - tw) / 2 : cx;
      accent(center ? cx + cw / 2 - 60 : tx, 300, 120, 10);
      text(tx, 340, tw, 140, "Closing message", 52, { font: { size: 52, color: "FFFFFF", bold: true, align: center ? "ctr" : "l" } });
      text(center ? cx + (cw - tw * 0.7) / 2 : tx, 510, tw * 0.7, 66, "Next step / call to action", 22, { font: { size: 22, color: "FFFFFF", align: center ? "ctr" : "l" } });
      text(center ? cx + cw / 2 - 180 : tx, 640, 360, 50, "Contact", 16, { font: { size: 16, color: "FFFFFF", align: center ? "ctr" : "l" } });
      break;
    }
  }
  footer();
  return out.map((s) => (s.t === "round" ? { ...s, rad } : s));
}

export function avatarZone(avatar: AvatarSpace): { x: number; y: number; w: number; h: number } | null {
  if (avatar === "None") return null;
  return { x: avatar === "Right" ? W - 470 : 0, y: 180, w: 470, h: H - 180 };
}

// ---------- Filled text ----------

/** Puts document text into the labelled blocks, in order of appearance. */
export function applyContent(shapes: Shape[], blocks: Record<string, string[]> | undefined): Shape[] {
  if (!blocks) return shapes;
  const used: Record<string, number> = {};
  return shapes.map((s) => {
    if (s.role !== "text" || !s.label) return s;
    const list = blocks[s.label];
    const i = used[s.label] ?? 0;
    used[s.label] = i + 1;
    const t = list?.[i]?.trim();
    return t ? { ...s, text: t } : s;
  });
}

const PX_PER_PT = W / 960; // the slide is 960pt wide

/** Shrinks the font until the wrapped text fits the block; returns the size (pt) and the wrapped lines. */
export function fitText(text: string, w: number, h: number, size: number): { size: number; lines: string[] } {
  const innerW = w - 48;
  const innerH = h - 20;
  const min = Math.max(9, Math.round(size * 0.45));
  let best: { size: number; lines: string[] } = { size: min, lines: [] };
  for (let sz = size; sz >= min; sz -= sz > 24 ? 2 : 1) {
    const charW = sz * PX_PER_PT * 0.52;
    const per = Math.max(4, Math.floor(innerW / charW));
    const lines: string[] = [];
    for (const para of text.split("\n")) {
      let line = "";
      for (const word of para.split(" ")) {
        if (!line) line = word;
        else if ((line + " " + word).length <= per) line += " " + word;
        else {
          lines.push(line);
          line = word;
        }
        while (line.length > per) {
          lines.push(line.slice(0, per));
          line = line.slice(per);
        }
      }
      lines.push(line);
    }
    best = { size: sz, lines };
    if (lines.length * sz * PX_PER_PT * 1.2 <= innerH) return best;
  }
  const max = Math.max(1, Math.floor(innerH / (best.size * PX_PER_PT * 1.2)));
  if (best.lines.length > max) best.lines = [...best.lines.slice(0, max - 1), best.lines[max - 1].replace(/\s*\S*$/, "") + "…"];
  return best;
}
