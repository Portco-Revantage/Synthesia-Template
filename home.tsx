import { useCallback, useEffect, useId, useMemo, useRef, useState } from "react";
import { toast } from "sonner";
import { ArrowRight, Download, Eraser, FilePlus2, FileText, Info, Plus, Shuffle, Trash2, Upload, Wand2, X } from "lucide-react";
import { RadioGroup, RadioGroupItem } from "@/components/ui/radio-group";
import { Spinner } from "@/components/ui/spinner";
import { ACCEPTED, readDocument } from "@/lib/docimport";
import { fillFromDocument, type FillMode, type SlideContent } from "@/data/fill";
import { Button } from "@/components/ui/button";
import { Input } from "@/components/ui/input";
import { Label } from "@/components/ui/label";
import { Textarea } from "@/components/ui/textarea";
import { Switch } from "@/components/ui/switch";
import { Tabs, TabsContent, TabsList, TabsTrigger } from "@/components/ui/tabs";
import { Select, SelectContent, SelectItem, SelectTrigger, SelectValue } from "@/components/ui/select";
import { Popover, PopoverContent, PopoverTrigger } from "@/components/ui/popover";
import { useConfirm } from "@/hooks/use-confirm";
import { cn } from "@/lib/utils";
import { buildPptx } from "@/lib/pptx";
import {
  H,
  LAYOUTS,
  W,
  applyContent,
  avatarZone,
  fitText,
  defaultLayouts,
  makeDesign,
  renderSlide,
  type AvatarSpace,
  type Design,
  type Fill,
  type Layout,
  type Shape,
} from "@/data/design";
import {
  AVATAR_OPTIONS,
  SLIDE_COUNTS,
  blankProject,
  loadProjects,
  parsePrompt,
  saveProjects,
  seedProjects,
  type Project,
} from "@/data/projects";

const STORAGE_KEY = "synthesia-template-builder:projects";
type SaveState = "saved" | "failed" | "unreadable" | "unavailable";

function useProjects() {
  const [projects, setProjects] = useState<Project[]>([]);
  const [save, setSave] = useState<SaveState>("saved");

  useEffect(() => {
    const r = loadProjects(STORAGE_KEY);
    if (r.status === "ok") {
      setProjects(r.projects);
      if (!saveProjects(STORAGE_KEY, r.projects)) setSave("failed");
    } else if (r.status === "empty") {
      const seeded = seedProjects();
      setProjects(seeded);
      if (!saveProjects(STORAGE_KEY, seeded)) setSave("failed");
    } else {
      setProjects(seedProjects());
      setSave(r.status === "invalid" ? "unreadable" : "unavailable");
    }
    const onStorage = (e: StorageEvent) => {
      if (e.key !== STORAGE_KEY) return;
      const fresh = loadProjects(STORAGE_KEY);
      if (fresh.status === "ok") setProjects(fresh.projects);
    };
    window.addEventListener("storage", onStorage);
    return () => window.removeEventListener("storage", onStorage);
  }, []);

  const mutate = useCallback((fn: (list: Project[]) => Project[]) => {
    setProjects((current) => {
      const fresh = loadProjects(STORAGE_KEY);
      const base = fresh.status === "ok" ? fresh.projects : fresh.status === "empty" ? [] : current;
      const next = fn(base);
      if (fresh.status === "invalid") setSave("unreadable");
      else if (saveProjects(STORAGE_KEY, next)) setSave("saved");
      else setSave(fresh.status === "unavailable" ? "unavailable" : "failed");
      return next;
    });
  }, []);

  return { projects, save, mutate };
}

function StorageNote({ save }: { save: SaveState }) {
  const warn = "rounded-md border border-amber-300 bg-amber-50 p-2 text-xs text-amber-950 dark:border-amber-800 dark:bg-amber-950 dark:text-amber-100";
  return (
    <div className="space-y-2">
      {save === "failed" && (
        <p role="alert" className="rounded-md border border-red-300 bg-red-50 p-2 text-xs text-red-900 dark:border-red-800 dark:bg-red-950 dark:text-red-100">
          Your latest changes couldn't be saved in this browser. Keep this page open to avoid losing them.
        </p>
      )}
      {save === "unreadable" && <p role="alert" className={warn}>Saved projects couldn't be loaded. Changes are kept only while this page is open.</p>}
      {save === "unavailable" && <p role="alert" className={warn}>Browser storage is unavailable. Changes are kept only while this page is open.</p>}
      <Popover>
        <PopoverTrigger asChild>
          <button type="button" className="inline-flex items-center gap-1 rounded text-xs text-foreground underline-offset-2 hover:underline focus-visible:outline-2 focus-visible:outline-ring">
            <Info className="size-3.5" aria-hidden="true" />
            Stored in this browser
          </button>
        </PopoverTrigger>
        <PopoverContent className="w-72 text-sm">
          Projects are saved only in this browser. They aren't shared with other people or synced to other browsers or devices, anyone using this
          browser profile can see them, and they're removed if browser data is cleared.
        </PopoverContent>
      </Popover>
    </div>
  );
}

// ---------- Slide rendering (SVG mirror of the PowerPoint shapes) ----------

function fillAttrs(f: Fill | undefined, gradId: string) {
  if (!f) return { fill: "none" };
  if ("g" in f) return { fill: `url(#${gradId})` };
  return { fill: `#${f.c}`, fillOpacity: f.a ?? 1 };
}

function SlideSvg({ shapes, avatar, guides, title, className }: { shapes: Shape[]; avatar: AvatarSpace; guides: boolean; title: string; className?: string }) {
  const uid = useId().replace(/:/g, "");
  const zone = avatarZone(avatar);
  return (
    <svg viewBox={`0 0 ${W} ${H}`} role="img" aria-labelledby={`${uid}-t`} className={cn("block h-auto w-full", className)}>
      <title id={`${uid}-t`}>{title}</title>
      <defs>
        {shapes.map((s, i) =>
          s.fill && "g" in s.fill ? (
            <linearGradient
              key={i}
              id={`${uid}-g${i}`}
              x1={0.5 - Math.cos((s.fill.ang * Math.PI) / 180) / 2}
              y1={0.5 - Math.sin((s.fill.ang * Math.PI) / 180) / 2}
              x2={0.5 + Math.cos((s.fill.ang * Math.PI) / 180) / 2}
              y2={0.5 + Math.sin((s.fill.ang * Math.PI) / 180) / 2}
            >
              <stop offset="0" stopColor={`#${s.fill.g[0]}`} />
              <stop offset="1" stopColor={`#${s.fill.g[1]}`} />
            </linearGradient>
          ) : null,
        )}
      </defs>
      {shapes.map((s, i) => {
        const common = {
          ...fillAttrs(s.fill, `${uid}-g${i}`),
          stroke: s.line ? `#${s.line.c}` : undefined,
          strokeOpacity: s.line?.a,
          strokeWidth: s.line?.w,
          transform: s.rot ? `rotate(${s.rot} ${s.x + s.w / 2} ${s.y + s.h / 2})` : undefined,
        };
        const r = s.t === "round" ? Math.min(s.w, s.h) * (s.rad ?? 0.1) : 0;
        const guide = guides && (s.role === "text" || s.role === "image");
        const fitted = s.text ? fitText(s.text, s.w, s.h, s.font?.size ?? 18) : null;
        const pxSize = fitted ? (fitted.size * W) / 960 : 0;
        const bottom = (s.font?.size ?? 18) >= 28;
        const center = s.font?.align === "ctr";
        const textTop = fitted ? (bottom ? s.y + s.h - 12 - fitted.lines.length * pxSize * 1.2 : s.y + 12) : 0;
        return (
          <g key={i}>
            {s.t === "ellipse" ? (
              <ellipse cx={s.x + s.w / 2} cy={s.y + s.h / 2} rx={s.w / 2} ry={s.h / 2} {...common} />
            ) : (
              <rect x={s.x} y={s.y} width={s.w} height={s.h} rx={r} {...common} />
            )}
            {fitted && (
              <text
                fontSize={pxSize}
                fill={`#${s.font?.color ?? "000000"}`}
                fontWeight={s.font?.bold ? 600 : 400}
                fontFamily="'Segoe UI', 'Geist Variable', sans-serif"
                textAnchor={center ? "middle" : "start"}
              >
                {fitted.lines.map((l, li) => (
                  <tspan key={li} x={center ? s.x + s.w / 2 : s.x + 24} y={textTop + pxSize * (li + 1) * 1.2 - pxSize * 0.25}>
                    {l}
                  </tspan>
                ))}
              </text>
            )}
            {guide && (
              <>
                <rect x={s.x} y={s.y} width={s.w} height={s.h} rx={r} fill="none" stroke={s.font?.color === "FFFFFF" || s.role === "image" ? "#ffffff" : "#000000"} strokeOpacity={0.45} strokeDasharray="10 8" strokeWidth={2} />
                {!fitted && s.h >= 30 && s.w >= 60 && (
                  <text
                    x={s.x + 14}
                    y={s.y + Math.min(s.h / 2 + 9, 38)}
                    fontSize={Math.min(26, Math.max(14, s.h * 0.32))}
                    fill={s.font?.color === "FFFFFF" || s.role === "image" ? "#ffffff" : "#000000"}
                    fillOpacity={0.6}
                    fontFamily="Geist Variable, sans-serif"
                  >
                    {s.label}
                  </text>
                )}
              </>
            )}
          </g>
        );
      })}
      {guides && zone && (
        <g>
          <rect x={zone.x + 30} y={zone.y} width={zone.w - 60} height={zone.h - 20} rx={24} fill="none" stroke="#888888" strokeDasharray="12 10" strokeWidth={3} />
          <text x={zone.x + zone.w / 2} y={zone.y + 50} textAnchor="middle" fontSize={26} fill="#888888" fontFamily="Geist Variable, sans-serif">
            Avatar space
          </text>
        </g>
      )}
    </svg>
  );
}

function Swatches({ d }: { d: Design }) {
  const p = d.palette;
  const sw = [
    ["Dark", p.dark],
    ["Deep", p.dark2],
    ["Primary", p.primary],
    ["Accent", p.accent],
    ["Light", p.light],
  ];
  return (
    <ul className="flex gap-2" aria-label="Palette colours">
      {sw.map(([n, c]) => (
        <li key={n} className="flex flex-col items-center gap-1">
          <span className="size-9 rounded-lg shadow-md ring-1 ring-white/25" style={{ background: `#${c}` }} aria-hidden="true" />
          <span className="text-[11px]">{n}</span>
          <span className="sr-only">#{c}</span>
        </li>
      ))}
    </ul>
  );
}

function Field({ label, htmlFor, children }: { label: string; htmlFor: string; children: React.ReactNode }) {
  return (
    <div className="grid gap-1.5">
      <Label htmlFor={htmlFor}>{label}</Label>
      {children}
    </div>
  );
}

// ---------- Page ----------

export function HomePage() {
  const { projects, save, mutate } = useProjects();
  const confirm = useConfirm();
  const [activeId, setActiveId] = useState<string | null>(null);
  const [tab, setTab] = useState<"prompt" | "template">("prompt");
  const [selected, setSelected] = useState(0);
  const [guides, setGuides] = useState(true);

  const sorted = useMemo(() => [...projects].sort((a, b) => b.updatedAt - a.updatedAt), [projects]);
  const active = projects.find((p) => p.id === activeId) ?? sorted[0] ?? null;
  const parsed = useMemo(() => parsePrompt(active?.prompt ?? ""), [active?.prompt]);
  const design = useMemo(() => (active ? makeDesign(active.name || parsed.name || "Untitled", active.seed) : null), [active, parsed.name]);
  const slides = useMemo(
    () => (active?.layouts && design ? active.layouts.map((l, i) => ({ layout: l, shapes: applyContent(renderSlide(l, design, active.avatar, i), active.content?.[i]?.blocks), notes: active.content?.[i]?.notes ?? "" })) : []),
    [active, design],
  );
  const previewTitle = useMemo(() => (design && active ? renderSlide("Title", design, active.avatar, 0) : []), [design, active]);
  const previewContent = useMemo(() => (design && active ? renderSlide("Three cards", design, active.avatar, 4) : []), [design, active]);

  useEffect(() => {
    if (activeId && !projects.some((p) => p.id === activeId)) setActiveId(null);
  }, [activeId, projects]);
  useEffect(() => {
    if (selected >= slides.length) setSelected(Math.max(0, slides.length - 1));
  }, [selected, slides.length]);

  const update = (patch: Partial<Project>) => {
    if (!active) return;
    mutate((list) => list.map((p) => (p.id === active.id ? { ...p, ...patch, updatedAt: Date.now() } : p)));
  };

  const setPrompt = (prompt: string) => {
    if (!active) return;
    const d = parsePrompt(prompt);
    update({
      prompt,
      ...(d.count ? { slideCount: d.count } : {}),
      ...(d.avatar ? { avatar: d.avatar } : {}),
    });
  };

  const build = async () => {
    if (!active) return;
    const name = active.name.trim() || parsed.name || "";
    if (!name) {
      toast.error("Add a project name or mention it in the prompt.");
      return;
    }
    if (active.layouts) {
      const ok = await confirm({
        title: "Rebuild the slides?",
        description: active.content
          ? "Layout changes and the text imported from your document will be replaced with blank slides. The design stays the same."
          : "Layout changes you made to the current slides will be replaced. The design stays the same.",
        confirmLabel: "Rebuild",
        destructive: true,
      });
      if (!ok) return;
    }
    update({ name, layouts: defaultLayouts(active.slideCount), content: null, source: null });
    setSelected(0);
    setTab("template");
    toast.success(`${active.slideCount} template slides built`);
  };

  const shuffle = () => active && update({ seed: (active.seed + 1) % 600 });

  const setLayout = (i: number, l: Layout) => {
    if (!active?.layouts) return;
    const layouts = active.layouts.map((x, j) => (j === i ? l : x));
    update({ layouts, slideCount: layouts.length });
  };
  const addSlide = () => {
    if (!active?.layouts || active.layouts.length >= 20) return;
    const at = Math.min(selected + 1, active.layouts.length);
    const layouts = [...active.layouts.slice(0, at), "Two columns" as Layout, ...active.layouts.slice(at)];
    const content = active.content ? [...active.content.slice(0, at), { blocks: {}, notes: "" }, ...active.content.slice(at)] : active.content;
    update({ layouts, slideCount: layouts.length, content });
    setSelected(at);
  };
  const removeSlide = (i: number) => {
    if (!active?.layouts || active.layouts.length <= 2) return;
    const layouts = active.layouts.filter((_, j) => j !== i);
    const content = active.content ? active.content.filter((_, j) => j !== i) : active.content;
    update({ layouts, slideCount: layouts.length, content });
    setSelected(Math.max(0, i - 1));
  };

  const addProject = () => {
    const p = blankProject();
    mutate((list) => [p, ...list]);
    setActiveId(p.id);
    setTab("prompt");
  };

  const removeProject = async () => {
    if (!active) return;
    const ok = await confirm({
      title: `Delete "${active.name || "Untitled project"}"?`,
      description: "The prompt and its slide template will be removed from this browser.",
      confirmLabel: "Delete",
      destructive: true,
    });
    if (!ok) return;
    mutate((list) => list.filter((p) => p.id !== active.id));
    setActiveId(null);
  };

  const download = () => {
    if (!active || !design || !slides.length) return;
    try {
      const bytes = buildPptx(active.name, slides.map((s) => s.shapes), design.palette, slides.map((s) => s.notes));
      const blob = new Blob([bytes.buffer as ArrayBuffer], { type: "application/vnd.openxmlformats-officedocument.presentationml.presentation" });
      const url = URL.createObjectURL(blob);
      const a = document.createElement("a");
      a.href = url;
      a.download = `${(active.name || "template").replace(/[^\p{L}\p{N}]+/gu, "-").replace(/^-|-$/g, "")}-${active.content ? "slides" : "template"}.pptx`;
      document.body.appendChild(a);
      a.click();
      a.remove();
      setTimeout(() => URL.revokeObjectURL(url), 5000);
      toast.success("PowerPoint downloaded");
    } catch {
      toast.error("Couldn't create the PowerPoint file. Try again.");
    }
  };

  const fileRef = useRef<HTMLInputElement>(null);
  const [file, setFile] = useState<File | null>(null);
  const [mode, setMode] = useState<FillMode>("fit");
  const [importing, setImporting] = useState(false);
  const [importError, setImportError] = useState<string | null>(null);

  const importDoc = async () => {
    if (!active || !file) return;
    setImportError(null);
    if (file.size > 40 * 1024 * 1024) {
      setImportError("That file is larger than 40 MB. Try a smaller copy.");
      return;
    }
    const useMode: FillMode = mode === "keep" && active.layouts ? "keep" : "fit";
    if (active.content || (useMode === "fit" && active.layouts)) {
      const ok = await confirm({
        title: "Replace the current slides?",
        description:
          useMode === "keep"
            ? "The text on your current slides will be replaced with text from this document. Layouts and design stay the same."
            : "Your current slides will be replaced with slides laid out to fit this document. The design stays the same.",
        confirmLabel: "Replace",
        destructive: true,
      });
      if (!ok) return;
    }
    setImporting(true);
    try {
      const doc = await readDocument(file);
      if (!doc.sections.length && !doc.title) {
        setImportError("No readable text was found in this file. Scanned PDFs and image-only slides can't be read.");
        return;
      }
      const r = fillFromDocument(doc, useMode, active.layouts);
      const name = active.name.trim() || parsed.name || doc.title;
      update({ name, layouts: r.layouts, slideCount: r.layouts.length, content: r.content as SlideContent[], source: file.name });
      setSelected(0);
      setTab("template");
      toast.success(`${r.layouts.length} slides filled from ${file.name}`);
      if (r.skipped > 0)
        toast.warning(
          useMode === "keep"
            ? `${r.skipped} more section${r.skipped === 1 ? "" : "s"} didn't fit your ${r.layouts.length} slides. Choose "Fit slides to the document" to include everything.`
            : `${r.skipped} section${r.skipped === 1 ? "" : "s"} didn't fit within 20 slides.`,
        );
    } catch (e) {
      const msg = e instanceof Error ? e.message : "";
      setImportError(
        msg === "unsupported"
          ? "This file type isn't supported. Use a Word (.docx), PowerPoint (.pptx), PDF or text file."
          : "This file couldn't be read. If it's an older .doc or .ppt file, or password-protected, save it as .docx or .pptx and try again.",
      );
    } finally {
      setImporting(false);
    }
  };

  const clearText = async () => {
    if (!active?.content) return;
    const ok = await confirm({
      title: "Remove the imported text?",
      description: "The slides go back to empty text blocks. Layouts and design stay the same.",
      confirmLabel: "Remove text",
      destructive: true,
    });
    if (ok) update({ content: null, source: null });
  };

  const current = slides[selected];

  return (
    <div className="grid flex-1 gap-4 p-4 md:grid-cols-[17rem_1fr]">
      <nav aria-label="Training projects" className="glass flex flex-col gap-3 rounded-2xl p-4 md:sticky md:top-24 md:max-h-[calc(100vh-7rem)] md:self-start md:overflow-auto">
        <div className="flex items-center justify-between">
          <h2 className="text-sm font-semibold">Projects</h2>
          <Button size="sm" variant="outline" onClick={addProject}>
            <FilePlus2 aria-hidden="true" />
            New
          </Button>
        </div>
        <ul className="flex flex-col gap-1">
          {sorted.map((p) => {
            const d = makeDesign(p.name || "Untitled", p.seed);
            const isCurrent = p.id === active?.id;
            return (
              <li key={p.id}>
                <button
                  type="button"
                  aria-current={isCurrent ? "true" : undefined}
                  onClick={() => {
                    setActiveId(p.id);
                    setSelected(0);
                  }}
                  className={cn(
                    "flex w-full items-center gap-2.5 rounded-md border border-transparent px-2.5 py-2 text-left hover:bg-accent focus-visible:outline-2 focus-visible:outline-ring",
                    isCurrent && "border-white/25 bg-white/15 shadow-lg shadow-violet-950/40",
                  )}
                >
                  <span className="flex h-8 w-3 shrink-0 flex-col overflow-hidden rounded-sm" aria-hidden="true">
                    <span className="flex-1" style={{ background: `#${d.palette.dark}` }} />
                    <span className="flex-1" style={{ background: `#${d.palette.accent}` }} />
                  </span>
                  <span className="min-w-0">
                    <span className="block truncate text-sm font-medium">{p.name || "Untitled project"}</span>
                    <span className="block text-xs">{p.layouts ? `${p.layouts.length} slides` : "Prompt only"}</span>
                  </span>
                </button>
              </li>
            );
          })}
        </ul>
        <div className="mt-auto pt-2">
          <StorageNote save={save} />
        </div>
      </nav>

      <div id="work" tabIndex={-1} className="glass min-w-0 rounded-2xl p-4 outline-none md:p-6">
        {!active || !design ? (
          <div className="flex flex-col items-start gap-3">
            <h1 className="text-xl font-semibold">No projects yet</h1>
            <Button onClick={addProject}>
              <FilePlus2 aria-hidden="true" />
              New project
            </Button>
          </div>
        ) : (
          <Tabs value={tab} onValueChange={(v) => setTab(v as "prompt" | "template")} className="gap-5">
            <div className="flex flex-wrap items-center justify-between gap-3">
              <h1 className="glow-text min-w-0 truncate text-2xl font-semibold tracking-tight">{active.name || parsed.name || "Untitled project"}</h1>
              <div className="flex items-center gap-2">
                <TabsList>
                  <TabsTrigger value="prompt">1 · Prompt</TabsTrigger>
                  <TabsTrigger value="template">2 · Template{active.layouts ? ` (${active.layouts.length})` : ""}</TabsTrigger>
                </TabsList>
                <Button variant="ghost" size="icon" aria-label="Delete this project" onClick={removeProject}>
                  <Trash2 aria-hidden="true" />
                </Button>
              </div>
            </div>

            <TabsContent value="prompt" className="grid gap-6 xl:grid-cols-[1fr_1.1fr]">
              <div className="grid content-start gap-5">
                <div className="grid gap-1.5">
                  <Label htmlFor="prompt">Prompt</Label>
                  <Textarea
                    id="prompt"
                    value={active.prompt}
                    onChange={(e) => setPrompt(e.target.value)}
                    placeholder="Maternity leave project, create 8 template slides"
                    className="min-h-36 text-base leading-relaxed"
                  />
                </div>
                <div className="grid gap-4 sm:grid-cols-3">
                  <Field label="Project name" htmlFor="pname">
                    <Input id="pname" value={active.name} placeholder={parsed.name ?? "e.g. Maternity Leave"} onChange={(e) => update({ name: e.target.value })} />
                  </Field>
                  <Field label="Slides" htmlFor="pcount">
                    <Select value={String(active.slideCount)} onValueChange={(v) => update({ slideCount: Number(v) })}>
                      <SelectTrigger id="pcount" className="w-full">
                        <SelectValue />
                      </SelectTrigger>
                      <SelectContent>
                        {SLIDE_COUNTS.map((n) => (
                          <SelectItem key={n} value={String(n)}>
                            {n} slides
                          </SelectItem>
                        ))}
                      </SelectContent>
                    </Select>
                  </Field>
                  <Field label="Avatar space" htmlFor="pavatar">
                    <Select value={active.avatar} onValueChange={(v) => update({ avatar: v as AvatarSpace })}>
                      <SelectTrigger id="pavatar" className="w-full">
                        <SelectValue />
                      </SelectTrigger>
                      <SelectContent>
                        {AVATAR_OPTIONS.map((o) => (
                          <SelectItem key={o} value={o}>
                            {o === "None" ? "No avatar space" : `${o} side`}
                          </SelectItem>
                        ))}
                      </SelectContent>
                    </Select>
                  </Field>
                </div>
                <div className="flex flex-wrap gap-2">
                  <Button onClick={build}>
                    <Wand2 aria-hidden="true" />
                    {active.layouts ? "Rebuild slides" : "Build template"}
                  </Button>
                  {active.layouts && (
                    <Button variant="outline" onClick={() => setTab("template")}>
                      Open template
                      <ArrowRight aria-hidden="true" />
                    </Button>
                  )}
                </div>
                <ol className="flex flex-wrap gap-1.5" aria-label="Slide plan">
                  {defaultLayouts(active.slideCount).map((l, i) => (
                    <li key={i} className="rounded-full border bg-white/8 px-2.5 py-1 text-xs backdrop-blur">
                      {i + 1}. {l}
                    </li>
                  ))}
                </ol>

                <section aria-labelledby="import-h" className="grid gap-3 rounded-2xl border bg-card p-4 shadow-xl shadow-black/20">
                  <div>
                    <h2 id="import-h" className="text-sm font-semibold">
                      Fill slides from a document
                    </h2>
                    <p className="mt-0.5 text-xs">
                      Upload the project's Word document, deck, PDF or text file. Headings become slide titles, the text goes into the blocks, and
                      each section is added to the speaker notes for the avatar script. The file is read in this browser only.
                    </p>
                  </div>
                  <div className="flex flex-wrap items-center gap-2">
                    <input
                      ref={fileRef}
                      id="docfile"
                      type="file"
                      accept={ACCEPTED}
                      className="sr-only"
                      onChange={(e) => {
                        setFile(e.target.files?.[0] ?? null);
                        setImportError(null);
                      }}
                    />
                    <Button variant="outline" onClick={() => fileRef.current?.click()}>
                      <Upload aria-hidden="true" />
                      {file ? "Choose another file" : "Choose file"}
                    </Button>
                    <span className="min-w-0 truncate text-sm" aria-live="polite">
                      {file ? file.name : "No file chosen"}
                    </span>
                  </div>
                  <RadioGroup value={active.layouts ? mode : "fit"} onValueChange={(v) => setMode(v as FillMode)} aria-label="How to use the document" className="gap-2">
                    <div className="flex items-start gap-2">
                      <RadioGroupItem id="mode-fit" value="fit" className="mt-0.5" />
                      <Label htmlFor="mode-fit" className="grid gap-0.5 font-normal">
                        <span className="font-medium">Fit slides to the document</span>
                        <span className="text-xs">Picks a layout for each section and sets the slide count to match.</span>
                      </Label>
                    </div>
                    <div className="flex items-start gap-2">
                      <RadioGroupItem id="mode-keep" value="keep" disabled={!active.layouts} className="mt-0.5" />
                      <Label htmlFor="mode-keep" className="grid gap-0.5 font-normal">
                        <span className="font-medium">Keep my current slides</span>
                        <span className="text-xs">{active.layouts ? `Fills your ${active.layouts.length} slides in order.` : "Build the template first to use this."}</span>
                      </Label>
                    </div>
                  </RadioGroup>
                  {importError && (
                    <p role="alert" className="rounded-md border border-red-300 bg-red-50 p-2 text-xs text-red-900 dark:border-red-800 dark:bg-red-950 dark:text-red-100">
                      {importError}
                    </p>
                  )}
                  <div>
                    <Button onClick={importDoc} disabled={!file || importing}>
                      {importing ? <Spinner aria-hidden="true" /> : <FileText aria-hidden="true" />}
                      {importing ? "Reading document…" : "Create slides from document"}
                    </Button>
                  </div>
                </section>
              </div>

              <section aria-labelledby="design-h" className="grid content-start gap-4 rounded-2xl border bg-card p-4 shadow-xl shadow-black/20">
                <div className="flex flex-wrap items-center justify-between gap-2">
                  <h2 id="design-h" className="text-sm font-semibold">
                    Design: {design.palette.name}
                  </h2>
                  <Button size="sm" variant="outline" onClick={shuffle}>
                    <Shuffle aria-hidden="true" />
                    Shuffle design
                  </Button>
                </div>
                <div className="grid grid-cols-2 gap-3">
                  <div className="overflow-hidden rounded-xl shadow-xl shadow-black/40 ring-1 ring-white/15">
                    <SlideSvg shapes={previewTitle} avatar={active.avatar} guides={false} title={`Title slide in ${design.palette.name}`} />
                  </div>
                  <div className="overflow-hidden rounded-xl shadow-xl shadow-black/40 ring-1 ring-white/15">
                    <SlideSvg shapes={previewContent} avatar={active.avatar} guides={false} title={`Content slide in ${design.palette.name}`} />
                  </div>
                </div>
                <div className="flex flex-wrap items-end justify-between gap-3">
                  <Swatches d={design} />
                  <dl className="grid grid-cols-[auto_auto] gap-x-3 gap-y-0.5 text-xs">
                    <dt className="font-medium">Motif</dt>
                    <dd>{design.motif}</dd>
                    <dt className="font-medium">Headings</dt>
                    <dd>{design.header}</dd>
                    <dt className="font-medium">Corners</dt>
                    <dd>{design.radius < 0.06 ? "Sharp" : design.radius < 0.15 ? "Soft" : "Rounded"}</dd>
                  </dl>
                </div>
              </section>
            </TabsContent>

            <TabsContent value="template" className="grid gap-4">
              {!active.layouts || !current ? (
                <div className="flex flex-col items-start gap-3 rounded-2xl border border-dashed border-white/25 bg-white/5 p-6">
                  <h2 className="text-base font-semibold">No template yet</h2>
                  <Button onClick={build}>
                    <Wand2 aria-hidden="true" />
                    Build template
                  </Button>
                </div>
              ) : (
                <>
                  <div className="flex flex-wrap items-center gap-2">
                    <Button onClick={download}>
                      <Download aria-hidden="true" />
                      Download PowerPoint
                    </Button>
                    <Button variant="outline" onClick={shuffle}>
                      <Shuffle aria-hidden="true" />
                      Shuffle design
                    </Button>
                    {active.content && (
                      <Button variant="outline" onClick={clearText}>
                        <Eraser aria-hidden="true" />
                        Remove text
                      </Button>
                    )}
                    <div className="ml-auto flex items-center gap-2">
                      <Switch id="guides" checked={guides} onCheckedChange={setGuides} />
                      <Label htmlFor="guides">Show block labels</Label>
                    </div>
                  </div>

                  <div className="grid gap-4 2xl:grid-cols-[1fr_16rem]">
                    <figure className="overflow-hidden rounded-xl shadow-2xl shadow-black/50 ring-1 ring-white/20">
                      <SlideSvg
                        shapes={current.shapes}
                        avatar={active.avatar}
                        guides={guides}
                        title={`Slide ${selected + 1} of ${slides.length}: ${current.layout}`}
                      />
                    </figure>
                    <div className="grid content-start gap-3">
                      <p className="text-sm font-semibold">
                        Slide {selected + 1} of {slides.length}
                      </p>
                      <Field label="Layout" htmlFor="slayout">
                        <Select value={current.layout} onValueChange={(v) => setLayout(selected, v as Layout)}>
                          <SelectTrigger id="slayout" className="w-full">
                            <SelectValue />
                          </SelectTrigger>
                          <SelectContent>
                            {LAYOUTS.map((l) => (
                              <SelectItem key={l} value={l}>
                                {l}
                              </SelectItem>
                            ))}
                          </SelectContent>
                        </Select>
                      </Field>
                      <div className="flex flex-wrap gap-2">
                        <Button size="sm" variant="outline" onClick={addSlide} disabled={slides.length >= 20}>
                          <Plus aria-hidden="true" />
                          Add slide after
                        </Button>
                        <Button size="sm" variant="outline" onClick={() => removeSlide(selected)} disabled={slides.length <= 2}>
                          <X aria-hidden="true" />
                          Remove slide
                        </Button>
                      </div>
                      {slides.length >= 20 && <p className="text-xs">20 slides is the maximum.</p>}
                      {active.source && (
                        <p className="flex items-center gap-1.5 text-xs">
                          <FileText className="size-3.5 shrink-0" aria-hidden="true" />
                          <span className="min-w-0 truncate">Text from {active.source}</span>
                        </p>
                      )}
                      {active.content && (
                        <div className="grid gap-1">
                          <p className="text-xs font-semibold">Speaker notes (avatar script)</p>
                          <p className="max-h-48 overflow-auto whitespace-pre-line rounded-lg border bg-white/6 p-2 text-xs">{current.notes || "No notes for this slide."}</p>
                        </div>
                      )}
                    </div>
                  </div>

                  <ol className="grid grid-cols-2 gap-3 sm:grid-cols-3 lg:grid-cols-4 xl:grid-cols-5" aria-label="All slides">
                    {slides.map((s, i) => (
                      <li key={i}>
                        <button
                          type="button"
                          aria-pressed={i === selected}
                          onClick={() => setSelected(i)}
                          className={cn(
                            "block w-full overflow-hidden rounded-lg text-left shadow-lg shadow-black/30 ring-1 ring-white/15 transition hover:-translate-y-0.5 hover:ring-white/40 focus-visible:outline-2 focus-visible:outline-offset-2 focus-visible:outline-ring",
                            i === selected && "ring-2 ring-cyan-300",
                          )}
                        >
                          <SlideSvg shapes={s.shapes} avatar={active.avatar} guides={false} title={`Slide ${i + 1}: ${s.layout}`} />
                        </button>
                        <p className={cn("mt-1 text-xs", i === selected && "font-semibold")}>
                          {i + 1}. {s.layout}
                        </p>
                      </li>
                    ))}
                  </ol>
                </>
              )}
            </TabsContent>
          </Tabs>
        )}
      </div>
    </div>
  );
}
