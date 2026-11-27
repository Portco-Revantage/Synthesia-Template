// Pours document sections into slide layouts by fixed rules (no AI).
import type { Layout } from "@/data/design";
import type { ImportedDoc, Section } from "@/lib/docimport";

/** Text for each labelled block on a slide (in order of appearance) plus speaker notes for the avatar script. */
export interface SlideContent {
  blocks: Record<string, string[]>;
  notes: string;
}

export type FillMode = "fit" | "keep";

const MAX = 20;

function split(point: string): [string, string] {
  const m = point.match(/^(.{2,70}?)(?::\s+|\s+[–—-]\s+)(.+)$/);
  if (m) return [m[1].trim(), m[2].trim()];
  const words = point.split(" ");
  if (words.length <= 7) return [point, ""];
  const sentence = point.match(/^(.{10,90}?[.!?])\s+(.+)$/);
  if (sentence) return [sentence[1].replace(/[.]$/, ""), sentence[2]];
  return [words.slice(0, 5).join(" ") + "…", point];
}

const has = (s: string, re: RegExp) => re.test(s.toLowerCase());
const DATE = /\b(?:\d{1,2}(?:st|nd|rd|th)?\s+(?:jan|feb|mar|apr|may|jun|jul|aug|sep|oct|nov|dec)|(?:jan|feb|mar|apr|may|jun|jul|aug|sep|oct|nov|dec)[a-z]*\s+\d|week\s*\d|day\s*\d|month\s*\d|q[1-4]\b|\d{4}|\d+\s+(?:days?|weeks?|months?)|before|after)/i;
const notesFor = (s: Section) => [s.title, ...s.points].filter(Boolean).join("\n");
const bullets = (pts: string[]) => pts.map((p) => `• ${p}`).join("\n");

function choose(s: Section): Layout {
  const t = s.title;
  const n = s.points.length;
  const structured = n > 0 && s.points.filter((p) => split(p)[1]).length >= Math.ceil(n * 0.6);
  if (has(t, /\?|quiz|knowledge check|test your|question/)) return "Knowledge check";
  if (has(t, /summary|takeaway|recap|key points|remember|in short/)) return "Summary checklist";
  if (has(t, /\bvs\b|versus|compar|do'?s and don'?ts|before and after|difference/)) return "Comparison";
  if (n >= 3 && n <= 6 && s.points.filter((p) => DATE.exec(p)?.index === 0).length >= Math.ceil(n * 0.6)) return "Timeline";
  if (n >= 3 && n <= 4 && has(t, /process|steps?|how to|procedure|apply|journey|workflow|stages?/)) return "Process steps";
  if (n === 0) return "Section divider";
  if (n === 1) return s.points[0].length <= 160 ? "Key highlight" : "Image + text";
  if (n === 3 && structured) return "Three cards";
  if (n === 4 && structured) return "Process steps";
  if (n <= 3) return "Image + text";
  return "Two columns";
}

function contentFor(layout: Layout, s: Section, index: number, doc: ImportedDoc, sectionNo: number): SlideContent {
  const P = s.points;
  const b: Record<string, string[]> = { "Slide title": [s.title] };
  switch (layout) {
    case "Title":
      return { blocks: { "Course title": [doc.title], Subtitle: [P[0] ?? ""], "Presenter / date": [""], Logo: [""] }, notes: notesFor({ title: doc.title, points: P }) };
    case "Agenda":
      P.slice(0, 4).forEach((p, i) => (b[`Agenda item ${i + 1}`] = [p]));
      break;
    case "Section divider":
      return { blocks: { "No.": [String(sectionNo).padStart(2, "0")], "Section title": [s.title], "Section summary": [P[0] ?? ""] }, notes: notesFor(s) };
    case "Image + text":
      if (P.length <= 3) P.forEach((p, i) => (b[`Point ${i + 1}`] = [p]));
      else [P.slice(0, 2), P.slice(2, 4), P.slice(4)].forEach((g, i) => (b[`Point ${i + 1}`] = [bullets(g)]));
      break;
    case "Three cards": {
      const parts = P.slice(0, 3).map(split);
      b["Card heading"] = parts.map((x) => x[0]);
      b["Card text"] = parts.map((x) => x[1]);
      break;
    }
    case "Process steps": {
      const parts = P.slice(0, 4).map(split);
      b["Step heading"] = parts.map((x) => x[0]);
      b["Step detail"] = parts.map((x) => x[1]);
      b["No."] = parts.map((_, i) => String(i + 1));
      break;
    }
    case "Key highlight":
      return { blocks: { "Key number or quote": [P[0] ?? s.title], "Supporting line": [P[0] ? s.title : ""], Source: [P[1] ?? ""] }, notes: notesFor(s) };
    case "Two columns": {
      const half = Math.ceil(P.length / 2);
      const cols = [P.slice(0, half), P.slice(half)];
      b["Column heading"] = ["", ""];
      b["Column text"] = cols.map(bullets);
      break;
    }
    case "Timeline": {
      const parts = P.slice(0, 5).map((p) => {
        const m = p.match(DATE);
        const [h, d] = split(p);
        return d && h.length <= 30 ? [h, d] : [m ? m[0].charAt(0).toUpperCase() + m[0].slice(1) : "", p];
      });
      b.Date = parts.map((x) => x[0]);
      b.Milestone = parts.map((x) => x[1]);
      break;
    }
    case "Comparison": {
      const half = Math.ceil(P.length / 2);
      const names = s.title.split(/\s+(?:vs\.?|versus|and)\s+/i);
      b["Option A"] = [names.length === 2 ? names[0] : "Option A"];
      b["Option B"] = [names.length === 2 ? names[1] : "Option B"];
      b["Comparison point"] = [...pad(P.slice(0, Math.min(half, 3)), 3), ...P.slice(half, half + 3)];
      break;
    }
    case "Full image":
      return { blocks: { "Caption title": [s.title], Caption: [P[0] ?? ""] }, notes: notesFor(s) };
    case "Summary checklist":
      P.slice(0, 4).forEach((p, i) => (b[`Takeaway ${i + 1}`] = [p]));
      b["Key message"] = [P[4] ? "Remember" : ""];
      b["Key message detail"] = [P.slice(4).join(" ")];
      break;
    case "Knowledge check":
    {
      const qInPoints = !s.title.includes("?") && /\?$/.test(P[0] ?? "");
      b.Question = [qInPoints ? P[0] : s.title.replace(/^(?:quick check|knowledge check|quiz|question)\s*[:\-–]\s*/i, "").replace(/^./, (c) => c.toUpperCase())];
      P.slice(qInPoints ? 1 : 0).slice(0, 4).forEach((p, i) => (b[`Option ${"ABCD"[i]}`] = [p.replace(/^[A-D][.)]\s*/, "")]));
      b["Slide title"] = ["Knowledge check"];
      break;
    }
    case "Closing":
      return { blocks: { "Closing message": [s.title || "Thank you"], "Next step / call to action": [P[0] ?? ""], Contact: [P[1] ?? ""] }, notes: notesFor(s) };
  }
  b["No."] = [...(b["No."] ?? []), String(index + 1)];
  return { blocks: b, notes: notesFor(s) };
}

const pad = (a: string[], n: number) => [...a, ...Array(Math.max(0, n - a.length)).fill("")];

/** Long sections are split into follow-on slides so blocks stay readable. */
function chunk(sections: Section[]): Section[] {
  const out: Section[] = [];
  for (const s of sections) {
    if (s.points.length <= 6) out.push(s);
    else for (let i = 0; i < s.points.length; i += 6) out.push({ title: i ? `${s.title} (continued)` : s.title, points: s.points.slice(i, i + 6) });
  }
  return out;
}

export interface FillResult {
  layouts: Layout[];
  content: SlideContent[];
  skipped: number;
}

export function fillFromDocument(doc: ImportedDoc, mode: FillMode, current: Layout[] | null): FillResult {
  const closingIdx = doc.sections.findIndex((s, i) => i === doc.sections.length - 1 && has(s.title, /thank|conclusion|next steps?|questions|contact|wrap/));
  const closing = closingIdx >= 0 ? doc.sections[closingIdx] : { title: "Thank you", points: [] };
  let rest = doc.sections.filter((_, i) => i !== closingIdx);
  const intro = { title: doc.title, points: [] as string[] };
  // A short untitled opening becomes the title slide's subtitle.
  if (rest[0] && /^introduction$/i.test(rest[0].title) && rest[0].points.length === 1 && rest[0].points[0].length <= 200) {
    intro.points = rest[0].points;
    rest = rest.slice(1);
  }
  const body = chunk(rest);

  if (mode === "keep" && current?.length) {
    const queue = [...body];
    let sectionNo = 0;
    const content = current.map((l, i) => {
      if (l === "Title") return contentFor(l, intro, i, doc, 0);
      if (l === "Closing") return contentFor(l, closing, i, doc, 0);
      if (l === "Agenda") return contentFor(l, { title: "Agenda", points: body.map((s) => s.title).slice(0, 4) }, i, doc, 0);
      const s = queue.shift();
      if (!s) return { blocks: {}, notes: "" };
      if (l === "Section divider") sectionNo++;
      return contentFor(l, s, i, doc, sectionNo);
    });
    return { layouts: current, content, skipped: queue.length };
  }

  const layouts: Layout[] = ["Title"];
  const content: SlideContent[] = [contentFor("Title", intro, 0, doc, 0)];
  const agenda = body.length >= 4;
  if (agenda) {
    layouts.push("Agenda");
    content.push(contentFor("Agenda", { title: "Agenda", points: body.map((s) => s.title).slice(0, 4) }, 1, doc, 0));
  }
  const room = MAX - layouts.length - 1;
  const used = body.slice(0, room);
  let sectionNo = 0;
  used.forEach((s) => {
    const l = choose(s);
    if (l === "Section divider") sectionNo++;
    content.push(contentFor(l, s, layouts.length, doc, sectionNo));
    layouts.push(l);
  });
  content.push(contentFor("Closing", closing, layouts.length, doc, 0));
  layouts.push("Closing");
  return { layouts, content, skipped: body.length - used.length };
}
