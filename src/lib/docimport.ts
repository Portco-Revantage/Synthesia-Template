// Reads an uploaded Word document, PowerPoint deck, PDF or text file into titled sections.
// Everything runs in the browser; the file never leaves the device.

export interface Section {
  title: string;
  points: string[];
}

export interface ImportedDoc {
  title: string;
  sections: Section[];
  kind: "Word" | "PowerPoint" | "PDF" | "Text";
}

export const ACCEPTED = ".docx,.pptx,.pdf,.txt,.md";

const decode = (s: string) =>
  s
    .replace(/&lt;/g, "<")
    .replace(/&gt;/g, ">")
    .replace(/&quot;/g, '"')
    .replace(/&apos;/g, "'")
    .replace(/&#(\d+);/g, (_, n) => String.fromCodePoint(Number(n)))
    .replace(/&#x([0-9a-f]+);/gi, (_, n) => String.fromCodePoint(parseInt(n, 16)))
    .replace(/&amp;/g, "&");

const clean = (s: string) =>
  Array.from(s, (ch) => (ch.charCodeAt(0) < 32 ? " " : ch))
    .join("")
    .replace(/\s+/g, " ")
    .trim();

// ---------- ZIP reading (stored and deflated entries) ----------

async function inflate(data: Uint8Array): Promise<Uint8Array> {
  const stream = new Blob([data.slice().buffer as ArrayBuffer]).stream().pipeThrough(new DecompressionStream("deflate-raw"));
  return new Uint8Array(await new Response(stream).arrayBuffer());
}

export async function unzip(buf: ArrayBuffer, want: (name: string) => boolean): Promise<Map<string, string>> {
  const b = new Uint8Array(buf);
  const v = new DataView(buf);
  let eocd = -1;
  for (let i = b.length - 22; i >= Math.max(0, b.length - 66000); i--) {
    if (v.getUint32(i, true) === 0x06054b50) {
      eocd = i;
      break;
    }
  }
  if (eocd < 0) throw new Error("not-zip");
  const count = v.getUint16(eocd + 10, true);
  let p = v.getUint32(eocd + 16, true);
  const out = new Map<string, string>();
  const td = new TextDecoder();
  for (let n = 0; n < count; n++) {
    if (v.getUint32(p, true) !== 0x02014b50) break;
    const method = v.getUint16(p + 10, true);
    const csize = v.getUint32(p + 20, true);
    const nlen = v.getUint16(p + 28, true);
    const elen = v.getUint16(p + 30, true);
    const clen = v.getUint16(p + 32, true);
    const local = v.getUint32(p + 42, true);
    const name = td.decode(b.subarray(p + 46, p + 46 + nlen));
    p += 46 + nlen + elen + clen;
    if (!want(name)) continue;
    const start = local + 30 + v.getUint16(local + 26, true) + v.getUint16(local + 28, true);
    const raw = b.subarray(start, start + csize);
    const data = method === 0 ? raw : method === 8 ? await inflate(raw) : null;
    if (data) out.set(name, td.decode(data));
  }
  return out;
}

// ---------- PowerPoint ----------

function pptxSlide(xml: string): Section | null {
  let title = "";
  const points: string[] = [];
  for (const sp of xml.match(/<p:sp>[\s\S]*?<\/p:sp>|<p:graphicFrame>[\s\S]*?<\/p:graphicFrame>/g) ?? []) {
    const isTitle = /<p:ph[^>]*type="(?:title|ctrTitle)"/.test(sp);
    const isMeta = /<p:ph[^>]*type="(?:sldNum|dt|ftr)"/.test(sp);
    if (isMeta) continue;
    const paras = (sp.match(/<a:p>[\s\S]*?<\/a:p>|<a:p\s[\s\S]*?<\/a:p>/g) ?? [])
      .map((para) => clean(decode((para.match(/<a:t>([\s\S]*?)<\/a:t>/g) ?? []).map((t) => t.replace(/<\/?a:t>/g, "")).join(""))))
      .filter((t) => t && !/^\d{1,3}$/.test(t));
    if (!paras.length) continue;
    if (isTitle && !title) title = paras.join(" ");
    else points.push(...paras);
  }
  if (!title && points.length) title = points.shift()!;
  return title || points.length ? { title, points } : null;
}

async function readPptx(buf: ArrayBuffer): Promise<Section[]> {
  const files = await unzip(buf, (n) => /^ppt\/slides\/slide\d+\.xml$/.test(n));
  return [...files.entries()]
    .sort((a, b) => Number(a[0].match(/(\d+)\.xml$/)![1]) - Number(b[0].match(/(\d+)\.xml$/)![1]))
    .map(([, xml]) => pptxSlide(xml))
    .filter((s): s is Section => s !== null);
}

// ---------- Word ----------

interface Para {
  text: string;
  level: number; // 0 = document title, 1–3 = heading level, 9 = body
  list: boolean;
  bold: boolean;
}

async function readDocx(buf: ArrayBuffer): Promise<Para[]> {
  const files = await unzip(buf, (n) => n === "word/document.xml" || n === "word/styles.xml");
  const xml = files.get("word/document.xml");
  if (!xml) throw new Error("not-docx");
  // Map style ids to their names so localised ids ("berschrift1") still resolve.
  const names = new Map<string, string>();
  for (const m of (files.get("word/styles.xml") ?? "").matchAll(/<w:style\b[^>]*w:styleId="([^"]+)"[^>]*>[\s\S]*?<w:name w:val="([^"]+)"/g)) names.set(m[1], m[2]);
  const body = xml.replace(/<w:del\b[\s\S]*?<\/w:del>/g, "");
  const out: Para[] = [];
  for (const p of body.match(/<w:p\b[^>]*>[\s\S]*?<\/w:p>/g) ?? []) {
    const text = clean(decode((p.match(/<w:t(?:\s[^>]*)?>([\s\S]*?)<\/w:t>|<w:tab\/>/g) ?? []).map((t) => (t === "<w:tab/>" ? " " : t.replace(/<\/?w:t[^>]*>/g, ""))).join("")));
    if (!text) continue;
    const id = p.match(/<w:pStyle w:val="([^"]+)"/)?.[1] ?? "";
    const style = (names.get(id) ?? id).toLowerCase().replace(/\s+/g, "");
    const outline = p.match(/<w:outlineLvl w:val="(\d)"/)?.[1];
    let level = 9;
    if (style === "title") level = 0;
    else if (/heading([1-3])$|^h([1-3])$/.test(style)) level = Number(style.match(/([1-3])$/)![1]);
    else if (outline !== undefined && Number(outline) < 3) level = Number(outline) + 1;
    const runs = p.match(/<w:r\b[\s\S]*?<\/w:r>/g) ?? [];
    const bold = runs.length > 0 && runs.every((r) => !/<w:t/.test(r) || /<w:b\/>|<w:b w:val="(?:1|true)"\/>/.test(r));
    out.push({ text, level, list: /<w:numPr>/.test(p) || /list/.test(style), bold });
  }
  return out;
}

// ---------- Text, Markdown, PDF ----------

function readText(src: string): Para[] {
  const out: Para[] = [];
  for (const raw of src.split(/\r?\n/)) {
    const line = raw.trim();
    if (!line) continue;
    const h = line.match(/^(#{1,3})\s+(.*)$/);
    if (h) out.push({ text: clean(h[2]), level: h[1].length === 1 && !out.length ? 0 : h[1].length, list: false, bold: true });
    else out.push({ text: clean(line.replace(/^(?:[-*•▪◦]|\d+[.)])\s+/, "")), level: 9, list: /^(?:[-*•▪◦]|\d+[.)])\s+/.test(line), bold: false });
  }
  return out;
}

async function readPdf(buf: ArrayBuffer): Promise<Section[]> {
  const { extractText, getDocumentProxy } = await import("unpdf");
  const pdf = await getDocumentProxy(new Uint8Array(buf));
  const { text } = await extractText(pdf, { mergePages: false });
  const pages = Array.isArray(text) ? text : [text];
  return pages
    .map((page) => {
      const lines = page.split(/\n+/).map(clean).filter((l) => l && !/^\d{1,3}$/.test(l));
      if (!lines.length) return null;
      const [title, ...rest] = lines;
      return { title, points: mergeLines(rest) };
    })
    .filter((s): s is Section => s !== null);
}

/** PDF text arrives line by line; join wrapped lines back into sentences or bullets. */
function mergeLines(lines: string[]): string[] {
  const out: string[] = [];
  for (const l of lines) {
    const bullet = /^[-*•▪◦]\s*/.test(l);
    const t = l.replace(/^[-*•▪◦]\s*/, "");
    const prev = out[out.length - 1];
    if (prev && !bullet && !/[.!?:;]$/.test(prev) && /^[a-z(]/.test(t)) out[out.length - 1] = `${prev} ${t}`;
    else out.push(t);
  }
  return out;
}

// ---------- Paragraphs → sections ----------

function toSections(paras: Para[]): { title: string; sections: Section[] } {
  let title = "";
  const sections: Section[] = [];
  const hasHeadings = paras.some((p) => p.level >= 1 && p.level <= 3);
  const top = hasHeadings ? Math.min(...paras.filter((p) => p.level >= 1 && p.level <= 3).map((p) => p.level)) : 0;
  let cur: Section | null = null;
  for (const p of paras) {
    if (p.level === 0 && !title) {
      title = p.text;
      continue;
    }
    const isHead = hasHeadings ? p.level >= 1 && p.level <= Math.min(3, top + 1) : p.bold && p.text.length <= 80 && !/[.]$/.test(p.text);
    if (isHead) {
      cur = { title: p.text, points: [] };
      sections.push(cur);
    } else {
      if (!cur) {
        cur = { title: "", points: [] };
        sections.push(cur);
      }
      cur.points.push(p.text);
    }
  }
  // An untitled opening block becomes the introduction.
  if (sections[0] && !sections[0].title) sections[0].title = title ? "Introduction" : sections[0].points.shift() ?? "Introduction";
  return { title, sections: sections.filter((s) => s.title || s.points.length) };
}

export async function readDocument(file: File): Promise<ImportedDoc> {
  const ext = file.name.toLowerCase().split(".").pop() ?? "";
  const buf = await file.arrayBuffer();
  const base = file.name.replace(/\.[^.]+$/, "").replace(/[_-]+/g, " ").trim();
  let doc: ImportedDoc;
  if (ext === "pptx") {
    const sections = await readPptx(buf);
    doc = { title: sections[0]?.title ?? base, sections: sections.slice(1), kind: "PowerPoint" };
    if (!doc.sections.length && sections[0]) doc.sections = [{ title: sections[0].title, points: sections[0].points }];
  } else if (ext === "docx") {
    const r = toSections(await readDocx(buf));
    doc = { title: r.title || base, sections: r.sections, kind: "Word" };
  } else if (ext === "pdf") {
    const sections = await readPdf(buf);
    doc = { title: sections[0]?.title ?? base, sections: sections.length > 1 ? sections.slice(1) : sections, kind: "PDF" };
  } else if (ext === "txt" || ext === "md") {
    const r = toSections(readText(new TextDecoder().decode(buf)));
    doc = { title: r.title || base, sections: r.sections, kind: "Text" };
  } else {
    throw new Error("unsupported");
  }
  doc.sections = doc.sections.map((s) => ({ title: s.title.slice(0, 160), points: s.points.map((p) => p.slice(0, 600)) }));
  return doc;
}
