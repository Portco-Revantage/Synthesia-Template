// Minimal PowerPoint (.pptx) writer: native shapes and empty text boxes, no dependencies.
import { W, fitText, type Fill, type Shape } from "@/data/design";

const EMU = 12192000 / W; // 16:9 widescreen, 13.333in × 7.5in
const emu = (n: number) => Math.round(n * EMU);
const esc = (s: string) => s.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");

function clr(c: string, a?: number) {
  return a !== undefined && a < 1 ? `<a:srgbClr val="${c}"><a:alpha val="${Math.round(a * 100000)}"/></a:srgbClr>` : `<a:srgbClr val="${c}"/>`;
}

function fillXml(f?: Fill) {
  if (!f) return "<a:noFill/>";
  if ("g" in f) {
    const ang = Math.round((((f.ang % 360) + 360) % 360) * 60000);
    return `<a:gradFill rotWithShape="1"><a:gsLst><a:gs pos="0">${clr(f.g[0])}</a:gs><a:gs pos="100000">${clr(f.g[1])}</a:gs></a:gsLst><a:lin ang="${ang}" scaled="0"/></a:gradFill>`;
  }
  return `<a:solidFill>${clr(f.c, f.a)}</a:solidFill>`;
}

function shapeXml(s: Shape, id: number) {
  const geom = s.t === "ellipse" ? "ellipse" : s.t === "round" ? "roundRect" : "rect";
  const av = s.t === "round" ? `<a:avLst><a:gd name="adj" fmla="val ${Math.round((s.rad ?? 0.1) * 50000)}"/></a:avLst>` : "<a:avLst/>";
  const ln = s.line ? `<a:ln w="${Math.round(s.line.w * 12700)}"><a:solidFill>${clr(s.line.c, s.line.a)}</a:solidFill></a:ln>` : "<a:ln><a:noFill/></a:ln>";
  const rot = s.rot ? ` rot="${Math.round(s.rot * 60000)}"` : "";
  const isText = s.role === "text";
  const name = isText ? `Text ${s.label ?? ""}` : s.role === "image" ? `Image ${s.label ?? ""}` : s.role === "bg" ? "Background" : `Shape ${id}`;
  const f = s.font;
  const size = s.text ? fitText(s.text, s.w, s.h, f?.size ?? 18).size : f?.size ?? 18;
  const rPr = (tag: "a:rPr" | "a:endParaRPr") =>
    `<${tag} lang="en-US" sz="${size * 100}"${f?.bold ? ' b="1"' : ""} dirty="0"><a:solidFill>${clr(f?.color ?? "000000")}</a:solidFill><a:latin typeface="${f?.bold ? "+mj-lt" : "+mn-lt"}"/></${tag}>`;
  const paras = (s.text ?? "").split("\n");
  const body = isText
    ? `<p:txBody><a:bodyPr wrap="square" lIns="182880" tIns="91440" rIns="182880" bIns="91440" anchor="${f && f.size >= 28 ? "b" : "t"}">${s.text ? "" : "<a:normAutofit/>"}</a:bodyPr><a:lstStyle/>${paras
        .map((t) => `<a:p><a:pPr algn="${f?.align ?? "l"}"/>${t ? `<a:r>${rPr("a:rPr")}<a:t>${esc(t)}</a:t></a:r>` : ""}${rPr("a:endParaRPr")}</a:p>`)
        .join("")}</p:txBody>`
    : `<p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:endParaRPr lang="en-US"/></a:p></p:txBody>`;
  return `<p:sp><p:nvSpPr><p:cNvPr id="${id}" name="${esc(name)}"/><p:cNvSpPr${isText ? ' txBox="1"' : ""}/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm${rot}><a:off x="${emu(s.x)}" y="${emu(s.y)}"/><a:ext cx="${emu(s.w)}" cy="${emu(s.h)}"/></a:xfrm><a:prstGeom prst="${geom}">${av}</a:prstGeom>${fillXml(s.fill)}${ln}</p:spPr>${body}</p:sp>`;
}

const NS = 'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"';
const XML = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n';
const REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const emptyTree = `<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>`;

function slideXml(shapes: Shape[]) {
  return `${XML}<p:sld ${NS}><p:cSld><p:spTree>${emptyTree}${shapes.map((s, i) => shapeXml(s, i + 2)).join("")}</p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sld>`;
}

function themeXml(name: string, c: { dark: string; light: string; primary: string; accent: string; ink: string }) {
  const fs = `<a:solidFill><a:schemeClr val="phClr"/></a:solidFill>`;
  return `${XML}<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="${esc(name)}"><a:themeElements><a:clrScheme name="${esc(name)}"><a:dk1><a:srgbClr val="${c.ink}"/></a:dk1><a:lt1><a:srgbClr val="FFFFFF"/></a:lt1><a:dk2><a:srgbClr val="${c.dark}"/></a:dk2><a:lt2><a:srgbClr val="${c.light}"/></a:lt2><a:accent1><a:srgbClr val="${c.primary}"/></a:accent1><a:accent2><a:srgbClr val="${c.accent}"/></a:accent2><a:accent3><a:srgbClr val="${c.dark}"/></a:accent3><a:accent4><a:srgbClr val="${c.primary}"/></a:accent4><a:accent5><a:srgbClr val="${c.accent}"/></a:accent5><a:accent6><a:srgbClr val="${c.ink}"/></a:accent6><a:hlink><a:srgbClr val="${c.primary}"/></a:hlink><a:folHlink><a:srgbClr val="${c.dark}"/></a:folHlink></a:clrScheme><a:fontScheme name="${esc(name)}"><a:majorFont><a:latin typeface="Segoe UI Semibold"/><a:ea typeface=""/><a:cs typeface=""/></a:majorFont><a:minorFont><a:latin typeface="Segoe UI"/><a:ea typeface=""/><a:cs typeface=""/></a:minorFont></a:fontScheme><a:fmtScheme name="${esc(name)}"><a:fillStyleLst>${fs}${fs}${fs}</a:fillStyleLst><a:lnStyleLst><a:ln w="6350"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln><a:ln w="12700"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln><a:ln w="19050"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln></a:lnStyleLst><a:effectStyleLst><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle></a:effectStyleLst><a:bgFillStyleLst>${fs}${fs}${fs}</a:bgFillStyleLst></a:fmtScheme></a:themeElements><a:objectDefaults/><a:extraClrSchemeLst/></a:theme>`;
}

export function buildPptx(
  title: string,
  slides: Shape[][],
  colors: { dark: string; light: string; primary: string; accent: string; ink: string },
  notes: string[] = [],
): Uint8Array {
  const n = slides.length;
  const hasNotes = notes.some((x) => x.trim());
  const files: [string, string][] = [];
  files.push([
    "[Content_Types].xml",
    `${XML}<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/><Override PartName="/ppt/slideMasters/slideMaster1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml"/><Override PartName="/ppt/slideLayouts/slideLayout1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"/><Override PartName="/ppt/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>${
      hasNotes
        ? `<Override PartName="/ppt/theme/theme2.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/><Override PartName="/ppt/notesMasters/notesMaster1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.notesMaster+xml"/>${slides
            .map((_, i) => `<Override PartName="/ppt/notesSlides/notesSlide${i + 1}.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.notesSlide+xml"/>`)
            .join("")}`
        : ""
    }${slides
      .map((_, i) => `<Override PartName="/ppt/slides/slide${i + 1}.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>`)
      .join("")}<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/><Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/></Types>`,
  ]);
  files.push([
    "_rels/.rels",
    `${XML}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${REL}/officeDocument" Target="ppt/presentation.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/><Relationship Id="rId3" Type="${REL}/extended-properties" Target="docProps/app.xml"/></Relationships>`,
  ]);
  files.push([
    "docProps/core.xml",
    `${XML}<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"><dc:title>${esc(title)}</dc:title><dcterms:created xsi:type="dcterms:W3CDTF">${new Date().toISOString().slice(0, 19)}Z</dcterms:created></cp:coreProperties>`,
  ]);
  files.push([
    "docProps/app.xml",
    `${XML}<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"><Application>Microsoft Office PowerPoint</Application><Slides>${n}</Slides></Properties>`,
  ]);
  files.push([
    "ppt/presentation.xml",
    `${XML}<p:presentation ${NS} saveSubsetFonts="1"><p:sldMasterIdLst><p:sldMasterId id="2147483648" r:id="rId1"/></p:sldMasterIdLst>${hasNotes ? `<p:notesMasterIdLst><p:notesMasterId r:id="rIdNM"/></p:notesMasterIdLst>` : ""}<p:sldIdLst>${slides
      .map((_, i) => `<p:sldId id="${256 + i}" r:id="rId${i + 3}"/>`)
      .join("")}</p:sldIdLst><p:sldSz cx="12192000" cy="6858000"/><p:notesSz cx="6858000" cy="9144000"/></p:presentation>`,
  ]);
  files.push([
    "ppt/_rels/presentation.xml.rels",
    `${XML}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${REL}/slideMaster" Target="slideMasters/slideMaster1.xml"/><Relationship Id="rId2" Type="${REL}/theme" Target="theme/theme1.xml"/>${hasNotes ? `<Relationship Id="rIdNM" Type="${REL}/notesMaster" Target="notesMasters/notesMaster1.xml"/>` : ""}${slides
      .map((_, i) => `<Relationship Id="rId${i + 3}" Type="${REL}/slide" Target="slides/slide${i + 1}.xml"/>`)
      .join("")}</Relationships>`,
  ]);
  files.push([
    "ppt/slideMasters/slideMaster1.xml",
    `${XML}<p:sldMaster ${NS}><p:cSld><p:bg><p:bgPr><a:solidFill><a:srgbClr val="${colors.light}"/></a:solidFill><a:effectLst/></p:bgPr></p:bg><p:spTree>${emptyTree}</p:spTree></p:cSld><p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/><p:sldLayoutIdLst><p:sldLayoutId id="2147483649" r:id="rId1"/></p:sldLayoutIdLst></p:sldMaster>`,
  ]);
  files.push([
    "ppt/slideMasters/_rels/slideMaster1.xml.rels",
    `${XML}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${REL}/slideLayout" Target="../slideLayouts/slideLayout1.xml"/><Relationship Id="rId2" Type="${REL}/theme" Target="../theme/theme1.xml"/></Relationships>`,
  ]);
  files.push([
    "ppt/slideLayouts/slideLayout1.xml",
    `${XML}<p:sldLayout ${NS} type="blank" preserve="1"><p:cSld name="Blank"><p:spTree>${emptyTree}</p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sldLayout>`,
  ]);
  files.push([
    "ppt/slideLayouts/_rels/slideLayout1.xml.rels",
    `${XML}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${REL}/slideMaster" Target="../slideMasters/slideMaster1.xml"/></Relationships>`,
  ]);
  files.push(["ppt/theme/theme1.xml", themeXml(title || "Template", colors)]);
  if (hasNotes) {
    files.push(["ppt/theme/theme2.xml", themeXml("Notes", colors)]);
    files.push(["ppt/notesMasters/notesMaster1.xml", notesMasterXml()]);
    files.push([
      "ppt/notesMasters/_rels/notesMaster1.xml.rels",
      `${XML}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${REL}/theme" Target="../theme/theme2.xml"/></Relationships>`,
    ]);
  }
  slides.forEach((shapes, i) => {
    files.push([`ppt/slides/slide${i + 1}.xml`, slideXml(shapes)]);
    files.push([
      `ppt/slides/_rels/slide${i + 1}.xml.rels`,
      `${XML}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${REL}/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>${
        hasNotes ? `<Relationship Id="rId2" Type="${REL}/notesSlide" Target="../notesSlides/notesSlide${i + 1}.xml"/>` : ""
      }</Relationships>`,
    ]);
    if (hasNotes) {
      files.push([`ppt/notesSlides/notesSlide${i + 1}.xml`, notesSlideXml(notes[i] ?? "")]);
      files.push([
        `ppt/notesSlides/_rels/notesSlide${i + 1}.xml.rels`,
        `${XML}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${REL}/notesMaster" Target="../notesMasters/notesMaster1.xml"/><Relationship Id="rId2" Type="${REL}/slide" Target="../slides/slide${i + 1}.xml"/></Relationships>`,
      ]);
    }
  });
  return zip(files);
}

// ---------- Speaker notes (Synthesia reads them as the avatar script) ----------

const notesBody = (text: string) =>
  `<p:txBody><a:bodyPr/><a:lstStyle/>${(text || "")
    .split("\n")
    .map((t) => `<a:p>${t ? `<a:r><a:rPr lang="en-US" dirty="0"/><a:t>${esc(t)}</a:t></a:r>` : `<a:endParaRPr lang="en-US"/>`}</a:p>`)
    .join("")}</p:txBody>`;

function notesMasterXml() {
  return `${XML}<p:notesMaster ${NS}><p:cSld><p:bg><p:bgRef idx="1001"><a:schemeClr val="bg1"/></p:bgRef></p:bg><p:spTree>${emptyTree}<p:sp><p:nvSpPr><p:cNvPr id="2" name="Slide Image Placeholder 1"/><p:cNvSpPr><a:spLocks noGrp="1" noRot="1" noChangeAspect="1"/></p:cNvSpPr><p:nvPr><p:ph type="sldImg" idx="2"/></p:nvPr></p:nvSpPr><p:spPr><a:xfrm><a:off x="381000" y="685800"/><a:ext cx="6096000" cy="3429000"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/></p:spPr></p:sp><p:sp><p:nvSpPr><p:cNvPr id="3" name="Notes Placeholder 2"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr><p:ph type="body" sz="quarter" idx="3"/></p:nvPr></p:nvSpPr><p:spPr><a:xfrm><a:off x="685800" y="4343400"/><a:ext cx="5486400" cy="4114800"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr>${notesBody("")}</p:sp></p:spTree></p:cSld><p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/><p:notesStyle><a:lvl1pPr marL="0" algn="l"><a:defRPr sz="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/></a:defRPr></a:lvl1pPr></p:notesStyle></p:notesMaster>`;
}

function notesSlideXml(text: string) {
  return `${XML}<p:notes ${NS}><p:cSld><p:spTree>${emptyTree}<p:sp><p:nvSpPr><p:cNvPr id="2" name="Slide Image Placeholder 1"/><p:cNvSpPr><a:spLocks noGrp="1" noRot="1" noChangeAspect="1"/></p:cNvSpPr><p:nvPr><p:ph type="sldImg"/></p:nvPr></p:nvSpPr><p:spPr/></p:sp><p:sp><p:nvSpPr><p:cNvPr id="3" name="Notes Placeholder 2"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr><p:ph type="body" idx="1"/></p:nvPr></p:nvSpPr><p:spPr/>${notesBody(text)}</p:sp></p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:notes>`;
}

// ---------- Stored (uncompressed) ZIP ----------

const CRC_TABLE = (() => {
  const t = new Uint32Array(256);
  for (let i = 0; i < 256; i++) {
    let c = i;
    for (let k = 0; k < 8; k++) c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1;
    t[i] = c >>> 0;
  }
  return t;
})();

function crc32(b: Uint8Array) {
  let c = 0xffffffff;
  for (let i = 0; i < b.length; i++) c = CRC_TABLE[(c ^ b[i]) & 0xff] ^ (c >>> 8);
  return (c ^ 0xffffffff) >>> 0;
}

function zip(entries: [string, string][]): Uint8Array {
  const enc = new TextEncoder();
  const chunks: Uint8Array[] = [];
  const central: Uint8Array[] = [];
  let offset = 0;
  for (const [name, content] of entries) {
    const nb = enc.encode(name);
    const data = enc.encode(content);
    const crc = crc32(data);
    const local = new DataView(new ArrayBuffer(30));
    local.setUint32(0, 0x04034b50, true);
    local.setUint16(4, 20, true);
    local.setUint16(6, 0x0800, true);
    local.setUint16(8, 0, true);
    local.setUint32(14, crc, true);
    local.setUint32(18, data.length, true);
    local.setUint32(22, data.length, true);
    local.setUint16(26, nb.length, true);
    chunks.push(new Uint8Array(local.buffer), nb, data);
    const cd = new DataView(new ArrayBuffer(46));
    cd.setUint32(0, 0x02014b50, true);
    cd.setUint16(4, 20, true);
    cd.setUint16(6, 20, true);
    cd.setUint16(8, 0x0800, true);
    cd.setUint32(16, crc, true);
    cd.setUint32(20, data.length, true);
    cd.setUint32(24, data.length, true);
    cd.setUint16(28, nb.length, true);
    cd.setUint32(42, offset, true);
    central.push(new Uint8Array(cd.buffer), nb);
    offset += 30 + nb.length + data.length;
  }
  const cdSize = central.reduce((a, c) => a + c.length, 0);
  const end = new DataView(new ArrayBuffer(22));
  end.setUint32(0, 0x06054b50, true);
  end.setUint16(8, entries.length, true);
  end.setUint16(10, entries.length, true);
  end.setUint32(12, cdSize, true);
  end.setUint32(16, offset, true);
  const all = [...chunks, ...central, new Uint8Array(end.buffer)];
  const out = new Uint8Array(all.reduce((a, c) => a + c.length, 0));
  let p = 0;
  for (const c of all) {
    out.set(c, p);
    p += c.length;
  }
  return out;
}

