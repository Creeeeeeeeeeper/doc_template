// DOCX engine: parse, preprocess, paragraph IO, render.
// Browser/Node compatible — no Tauri APIs in here.
import PizZip from "pizzip";
import Docxtemplater from "docxtemplater";
import ImageModule from "docxtemplater-image-module-free";

// ----------------------------------------------------------------
// XML helpers
// ----------------------------------------------------------------
export function escapeXmlText(s) {
  return String(s)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;");
}

export function escapeXmlAttr(s) {
  return String(s)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&apos;");
}

export function decodeXmlText(s) {
  return String(s)
    .replace(/&apos;/g, "'")
    .replace(/&quot;/g, '"')
    .replace(/&gt;/g, ">")
    .replace(/&lt;/g, "<")
    .replace(/&amp;/g, "&");
}

// ----------------------------------------------------------------
// Field detection (looks for {@text} / {%image})
// ----------------------------------------------------------------
export function extractFieldsFromXml(xml) {
  const stripped = xml.replace(/<[^>]+>/g, "");
  const re = /\{([@%])(\w+)\}/g;
  const found = new Map();
  let m;
  while ((m = re.exec(stripped)) !== null) {
    const [, sigil, name] = m;
    if (!found.has(name)) {
      found.set(name, sigil === "%" ? "image" : "text");
    }
  }
  return [...found.entries()].map(([name, type]) => ({ name, type }));
}

export function extractFieldsFromZip(zip) {
  const xml = zip.file("word/document.xml").asText();
  return extractFieldsFromXml(xml);
}

// ----------------------------------------------------------------
// Paragraph IO — used by edit mode
// ----------------------------------------------------------------
// Returns array of { index, originalXml, originalText, currentText, dirty,
//                    pStart, pEnd, hasComplex, selfClosing }
// Both <w:p>...</w:p> and self-closing <w:p/> are matched so the count
// matches docx-preview's rendering (which emits a <p> for empty paragraphs).
export function parseParagraphs(zip) {
  const xml = zip.file("word/document.xml").asText();
  const re = /<w:p\b([^>/]*)\/>|<w:p\b[^>]*>[\s\S]*?<\/w:p>/g;
  const out = [];
  let m;
  let i = 0;
  while ((m = re.exec(xml)) !== null) {
    const fullXml = m[0];
    const pStart = m.index;
    const pEnd = pStart + fullXml.length;
    const selfClosing = /<w:p\b[^>]*\/>$/.test(fullXml);

    let text = "";
    let hasComplex = false;
    if (!selfClosing) {
      const texts = [...fullXml.matchAll(/<w:t(?:\s[^>]*)?>([\s\S]*?)<\/w:t>/g)].map(
        (x) => decodeXmlText(x[1]),
      );
      text = texts.join("");
      hasComplex =
        /<w:drawing\b/.test(fullXml) ||
        /<w:hyperlink\b/.test(fullXml) ||
        /<w:fldChar\b/.test(fullXml) ||
        /<w:instrText\b/.test(fullXml);
    }

    out.push({
      index: i++,
      originalXml: fullXml,
      originalText: text,
      currentText: text,
      dirty: false,
      pStart,
      pEnd,
      hasComplex,
      selfClosing,
    });
  }
  return out;
}

// Build new <w:p>...</w:p> XML for an edited paragraph.
// Strategy: keep <w:pPr> if present, keep first <w:rPr> as the default run
// formatting, and replace ALL runs with a single run containing the new text.
// Multi-line content uses <w:br/> between text fragments inside the run.
// For previously self-closing <w:p/>, expand to a full paragraph.
export function buildParagraphXml(paragraph) {
  const original = paragraph.originalXml;
  const newText = paragraph.currentText ?? "";

  if (paragraph.selfClosing) {
    // Promote to a non-self-closing paragraph; preserve attributes from <w:p .../>
    const attrs =
      original.match(/^<w:p\b([^>/]*)\/>/)?.[1] ?? "";
    const lines = String(newText).split(/\r?\n/);
    const inner = lines
      .map(
        (line, idx) =>
          `${idx > 0 ? "<w:br/>" : ""}<w:t xml:space="preserve">${escapeXmlText(line)}</w:t>`,
      )
      .join("");
    if (newText.length === 0) {
      return `<w:p${attrs}/>`;
    }
    return `<w:p${attrs}><w:r>${inner}</w:r></w:p>`;
  }

  const pAttrsMatch = original.match(/^<w:p\b([^>]*)>/);
  const pAttrs = pAttrsMatch ? pAttrsMatch[1] : "";

  const pPrMatch = original.match(/<w:pPr>[\s\S]*?<\/w:pPr>/);
  const pPr = pPrMatch ? pPrMatch[0] : "";

  // First <w:rPr> in the paragraph (excluding any inside <w:pPr>)
  const afterPPr = pPrMatch
    ? original.slice(pPrMatch.index + pPrMatch[0].length)
    : original;
  const rPrMatch = afterPPr.match(/<w:rPr>[\s\S]*?<\/w:rPr>/);
  const rPr = rPrMatch ? rPrMatch[0] : "";

  const lines = String(newText).split(/\r?\n/);
  const inner = lines
    .map(
      (line, idx) =>
        `${idx > 0 ? "<w:br/>" : ""}<w:t xml:space="preserve">${escapeXmlText(line)}</w:t>`,
    )
    .join("");

  return `<w:p${pAttrs}>${pPr}<w:r>${rPr}${inner}</w:r></w:p>`;
}

// Apply edits back to the zip's document.xml.
// `paragraphs` must come from `parseParagraphs(zip)` against the SAME zip.
export function applyParagraphEdits(zip, paragraphs) {
  const dirty = paragraphs.filter((p) => p.dirty);
  if (dirty.length === 0) return;

  let xml = zip.file("word/document.xml").asText();

  // Replace in reverse order so earlier indices stay valid
  const sorted = [...dirty].sort((a, b) => b.pStart - a.pStart);
  for (const p of sorted) {
    const replacement = buildParagraphXml(p);
    xml = xml.slice(0, p.pStart) + replacement + xml.slice(p.pEnd);
  }
  zip.file("word/document.xml", xml);
}

// ----------------------------------------------------------------
// Style preprocessing for fill mode
// ----------------------------------------------------------------
function buildRPr({ font, size, color }) {
  const halfPt = Math.round(Number(size || 12) * 2);
  const hex = String(color || "#000000").replace(/^#/, "").toUpperCase();
  const safeFont = escapeXmlAttr(font || "Calibri");
  return (
    `<w:rPr>` +
    `<w:rFonts w:ascii="${safeFont}" w:hAnsi="${safeFont}" w:cs="${safeFont}" w:eastAsia="${safeFont}"/>` +
    `<w:sz w:val="${halfPt}"/><w:szCs w:val="${halfPt}"/>` +
    `<w:color w:val="${hex}"/>` +
    `</w:rPr>`
  );
}

// Walk every <w:r>; if its <w:t> contains {@field}, split the run so each
// {@field} occurrence sits in its own run with our chosen rPr, and replace
// {@field} with {field} for docxtemplater's standard substitution.
export function preprocessTemplateForFill(zip, styleMap) {
  let xml = zip.file("word/document.xml").asText();
  const occurrenceCount = {};

  xml = xml.replace(/<w:r\b[^>]*>[\s\S]*?<\/w:r>/g, (run) => {
    if (!/\{@\w+\}/.test(run)) return run;

    const rPrMatch = run.match(/<w:rPr>[\s\S]*?<\/w:rPr>/);
    const rPrOriginal = rPrMatch ? rPrMatch[0] : "";

    const tMatch = run.match(/<w:t(\s[^>]*)?>([\s\S]*?)<\/w:t>/);
    if (!tMatch) return run;
    const tContent = tMatch[2];
    if (!/\{@\w+\}/.test(tContent)) return run;

    const parts = [];
    let lastIdx = 0;
    for (const m of tContent.matchAll(/\{@(\w+)\}/g)) {
      if (m.index > lastIdx) {
        parts.push({ type: "text", value: tContent.slice(lastIdx, m.index) });
      }
      parts.push({ type: "placeholder", name: m[1] });
      lastIdx = m.index + m[0].length;
    }
    if (lastIdx < tContent.length) {
      parts.push({ type: "text", value: tContent.slice(lastIdx) });
    }

    let out = "";
    for (const p of parts) {
      if (p.type === "text") {
        if (p.value.length === 0) continue;
        out += `<w:r>${rPrOriginal}<w:t xml:space="preserve">${p.value}</w:t></w:r>`;
      } else {
        const occ = occurrenceCount[p.name] || 0;
        occurrenceCount[p.name] = occ + 1;
        const entry = styleMap[p.name];
        let style;
        if (Array.isArray(entry)) {
          style = entry[occ] || entry[0] || {};
        } else {
          style = entry || {};
        }
        const styledRPr = buildRPr(style);
        out += `<w:r>${styledRPr}<w:t xml:space="preserve">{${p.name}}</w:t></w:r>`;
      }
    }
    return out;
  });

  zip.file("word/document.xml", xml);
}

// ----------------------------------------------------------------
// Fill rendering
// ----------------------------------------------------------------
// Read image dimensions from file bytes (supports PNG/JPEG/GIF).
// Returns { naturalW, naturalH } or null.
function getImageDimensions(bytes) {
  if (!bytes || bytes.length < 16) return null;
  try {
    // PNG: 8-byte signature, then IHDR chunk (width at offset 16, height at 20)
    if (bytes[0] === 0x89 && bytes[1] === 0x50 && bytes[2] === 0x4E && bytes[3] === 0x47) {
      const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
      return { naturalW: view.getUint32(16), naturalH: view.getUint32(20) };
    }
    // GIF: width at offset 6, height at 8
    if (bytes[0] === 0x47 && bytes[1] === 0x49 && bytes[2] === 0x46) {
      const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
      return { naturalW: view.getUint16(6), naturalH: view.getUint16(8) };
    }
    // JPEG: scan for SOF marker
    if (bytes[0] === 0xFF && bytes[1] === 0xD8) {
      let offset = 2;
      while (offset < bytes.length - 1) {
        if (bytes[offset] !== 0xFF) { offset++; continue; }
        const marker = bytes[offset + 1];
        if (marker === 0xC0 || marker === 0xC1 || marker === 0xC2) {
          const view = new DataView(bytes.buffer, bytes.byteOffset + offset + 5, 4);
          return { naturalH: view.getUint16(0), naturalW: view.getUint16(2) };
        }
        if (marker === 0xD9 || marker === 0xDA) break;
        const segLen = new DataView(bytes.buffer, bytes.byteOffset + offset + 2, 2).getUint16(0);
        offset += 2 + segLen;
      }
    }
  } catch { /* ignore parse errors */ }
  return null;
}

// values: { fieldName: { type:'text', text, font, size, color } |
//                       { type:'image', bytes: Uint8Array } }
// imageConfigMap: Map<fieldName, imageConfig>
// Returns Uint8Array of the rendered docx bytes.
export function renderFilled(templateBytes, fields, values, occStyleMap, imageConfigMap) {
  const zip = new PizZip(templateBytes);

  // Strip editor-only metadata; the filled doc has no use for it, and a
  // dangling unknown-content-type part trips Word's "corrupt file" check.
  if (zip.file("template/fields.json")) {
    zip.remove("template/fields.json");
  }

  const styleMap = {};
  for (const f of fields) {
    if (f.type === "text") {
      if (occStyleMap && occStyleMap.has(f.name)) {
        styleMap[f.name] = occStyleMap.get(f.name);
      } else {
        const v = values[f.name];
        styleMap[f.name] = {
          font: v?.font || "Calibri",
          size: v?.size || 12,
          color: v?.color || "#000000",
        };
      }
    }
  }
  preprocessTemplateForFill(zip, styleMap);

  const imageMap = new Map();
  const imageDimCache = new Map(); // tagName -> { naturalW, naturalH }

  const imageModule = new ImageModule({
    centered: false,
    getImage: (_tagValue, tagName) => {
      const bytes = imageMap.get(tagName);
      if (!bytes) return bytes;
      // Pre-read natural dimensions from the image bytes
      if (!imageDimCache.has(tagName)) {
        imageDimCache.set(tagName, getImageDimensions(bytes));
      }
      return bytes;
    },
    getSize: (tagValue, _imageData) => {
      const config = (imageConfigMap && imageConfigMap.get(tagValue)) || {};
      const {
        fitMode = "width",
        maintainRatio = true,
        maxWidth = 300,
        maxHeight = 400,
        minWidth = 50,
        minHeight = 50,
      } = config;

      const dims = imageDimCache.get(tagValue);
      const natW = dims?.naturalW || 0;
      const natH = dims?.naturalH || 0;
      const ratio = natW > 0 && natH > 0 ? natW / natH : 1;

      let w, h;
      if (fitMode === "height") {
        h = maxHeight;
        w = maintainRatio ? Math.round(h * ratio) : maxWidth;
      } else if (fitMode === "contain") {
        // Fit inside maxWidth × maxHeight box
        const boxRatio = maxWidth / maxHeight;
        if (ratio > boxRatio) {
          w = maxWidth;
          h = maintainRatio ? Math.round(w / ratio) : maxHeight;
        } else {
          h = maxHeight;
          w = maintainRatio ? Math.round(h * ratio) : maxWidth;
        }
      } else {
        // "width" (default)
        w = maxWidth;
        h = maintainRatio ? Math.round(w / ratio) : maxHeight;
      }

      w = Math.max(minWidth, Math.min(maxWidth, w));
      h = Math.max(minHeight, Math.min(maxHeight, h));
      return [w, h];
    },
  });

  const doc = new Docxtemplater(zip, {
    paragraphLoop: true,
    linebreaks: true,
    modules: [imageModule],
  });

  const data = {};
  for (const f of fields) {
    const v = values[f.name];
    if (f.type === "text") {
      data[f.name] = v?.text ?? "";
    } else {
      if (!v?.bytes) {
        throw new Error(`图片字段 "${f.name}" 尚未选择文件`);
      }
      imageMap.set(f.name, v.bytes);
      data[f.name] = f.name; // truthy string -> normal flow in image module
    }
  }

  doc.render(data);
  return doc.getZip().generate({ type: "uint8array" });
}

// ----------------------------------------------------------------
// Detect run formatting at a given character offset within a paragraph.
// Used by the editor to seed the "default format" of a new placeholder
// from the surrounding text's font/size/color.
// ----------------------------------------------------------------
// Returns { font, size, color, sizeLabel } where any field may be null.
//   font   — string (family name); prefers eastAsia, falls back to ascii.
//   size   — number, in points (e.g. 12).
//   color  — string "#RRGGBB" (uppercase) or null.
// charOffset is in the SAME unit as the concatenated <w:t> content
// (the same unit that parseParagraphs returns as `originalText`).
export function getRunStyleAt(paragraphXml, charOffset) {
  if (!paragraphXml) return null;
  const runRegex = /<w:r\b[^>]*>[\s\S]*?<\/w:r>/g;
  let cumLen = 0;
  let lastMatch = null;
  let m;
  while ((m = runRegex.exec(paragraphXml)) !== null) {
    lastMatch = m;
    const runXml = m[0];
    const texts = [...runXml.matchAll(/<w:t(?:\s[^>]*)?>([\s\S]*?)<\/w:t>/g)];
    const runLen = texts.reduce((sum, t) => sum + t[1].length, 0);
    // Strict greater so a click at the very boundary picks the next run
    if (cumLen + runLen > charOffset && runLen > 0) {
      return parseRunRPr(runXml);
    }
    cumLen += runLen;
  }
  // Past end — fall back to last run's style
  return lastMatch ? parseRunRPr(lastMatch[0]) : null;
}

function parseRunRPr(runXml) {
  const rPrMatch = runXml.match(/<w:rPr>([\s\S]*?)<\/w:rPr>/);
  if (!rPrMatch) return null;
  const rPr = rPrMatch[1];

  let font = null;
  const fontsTag = rPr.match(/<w:rFonts\b[^/>]*\/?>/);
  if (fontsTag) {
    const ea = fontsTag[0].match(/w:eastAsia="([^"]+)"/);
    const ascii = fontsTag[0].match(/w:ascii="([^"]+)"/);
    const hAnsi = fontsTag[0].match(/w:hAnsi="([^"]+)"/);
    font = ea?.[1] || ascii?.[1] || hAnsi?.[1] || null;
  }

  // <w:sz w:val="N"/> — N is in half-points
  const szMatch = rPr.match(/<w:sz\s+w:val="(\d+(?:\.\d+)?)"/);
  const size = szMatch ? parseFloat(szMatch[1]) / 2 : null;

  // <w:color w:val="HEX"/>
  const colorMatch = rPr.match(/<w:color\s+w:val="([0-9A-Fa-f]{6})"/);
  const color = colorMatch ? "#" + colorMatch[1].toUpperCase() : null;

  return { font, size, color };
}

// ----------------------------------------------------------------
// [Content_Types].xml maintenance
// ----------------------------------------------------------------
// OOXML requires every part in the package to have a content type — either
// via a Default extension match or an Override path match. Before this
// helper, we wrote `template/fields.json` without declaring `.json`, and
// Word/WPS rejected the resulting docx as "corrupt".
function ensureJsonContentType(zip) {
  const path = "[Content_Types].xml";
  const file = zip.file(path);
  if (!file) return;
  const ct = file.asText();
  if (/Extension="json"/i.test(ct)) return;
  const inject =
    '<Default Extension="json" ContentType="application/json"/>';
  // Insert just before </Types>; fall back to appending if no closing tag.
  const updated = /<\/Types>/.test(ct)
    ? ct.replace(/<\/Types>/, `${inject}</Types>`)
    : ct + inject;
  zip.file(path, updated);
}

// ----------------------------------------------------------------
// Edit mode save: just apply paragraph edits and return bytes
// ----------------------------------------------------------------
// Update the buildTemplate fieldMeta serialization to include default format.
// (See `default*` fields in the JSON below.)
export function buildTemplate(templateBytes, paragraphs, fieldMeta, occurrenceStyles) {
  const zip = new PizZip(templateBytes);
  applyParagraphEdits(zip, paragraphs);

  if (fieldMeta && fieldMeta.size > 0) {
    const occEntries = [];
    if (occurrenceStyles) {
      for (const [pIdx, styles] of occurrenceStyles) {
        const pText = paragraphs[pIdx]?.currentText ?? "";
        const re = /\{([@%])(\w+)\}/g;
        let occ = 0;
        let m;
        while ((m = re.exec(pText)) !== null) {
          const sigil = m[1];
          const name = m[2];
          const meta = fieldMeta.get(name);
          if (!meta) continue;
          const expectedType = sigil === "%" ? "image" : "text";
          if (meta.type !== expectedType) continue;
          if (styles[occ]) {
            occEntries.push({
              pIdx,
              occ,
              name,
              sigil,
              font: styles[occ].font || null,
              size: styles[occ].size ?? null,
              sizeLabel: styles[occ].sizeLabel || null,
              color: styles[occ].color || null,
            });
          }
          occ++;
        }
      }
    }
    const data = {
      version: 4,
      fields: [...fieldMeta.entries()].map(([name, m]) => {
        const entry = {
          name,
          type: m.type || "text",
          description: m.description || "",
        };
        if (m.type === "image" && m.imageConfig) {
          entry.imageConfig = { ...m.imageConfig };
        }
        return entry;
      }),
      occStyles: occEntries,
    };
    zip.file("template/fields.json", JSON.stringify(data, null, 2));
    ensureJsonContentType(zip);
  } else if (zip.file("template/fields.json")) {
    zip.remove("template/fields.json");
  }

  return zip.generate({ type: "uint8array" });
}

const DEFAULT_IMAGE_CONFIG = {
  fitMode: "width",
  maintainRatio: true,
  maxWidth: 300,
  maxHeight: 400,
  minWidth: 50,
  minHeight: 50,
};

export function readFieldMeta(zip) {
  const file = zip.file("template/fields.json");
  if (!file) return { fieldMeta: new Map(), occStyles: new Map() };
  try {
    const data = JSON.parse(file.asText());
    const map = new Map();
    for (const f of data.fields || []) {
      if (!f.name) continue;
      const entry = {
        type: f.type === "image" ? "image" : "text",
        description: f.description || "",
      };
      if (entry.type === "image") {
        const ic = f.imageConfig || {};
        entry.imageConfig = {
          fitMode: ic.fitMode || DEFAULT_IMAGE_CONFIG.fitMode,
          maintainRatio: ic.maintainRatio !== false,
          maxWidth: Number(ic.maxWidth) || DEFAULT_IMAGE_CONFIG.maxWidth,
          maxHeight: Number(ic.maxHeight) || DEFAULT_IMAGE_CONFIG.maxHeight,
          minWidth: Number(ic.minWidth) || DEFAULT_IMAGE_CONFIG.minWidth,
          minHeight: Number(ic.minHeight) || DEFAULT_IMAGE_CONFIG.minHeight,
        };
      }
      map.set(f.name, entry);
    }
    const occStyles = new Map();
    for (const entry of data.occStyles || []) {
      if (entry.pIdx == null || entry.occ == null) continue;
      if (!occStyles.has(entry.pIdx)) occStyles.set(entry.pIdx, []);
      const arr = occStyles.get(entry.pIdx);
      arr[entry.occ] = {
        name: entry.name || null,
        sigil: entry.sigil || null,
        font: entry.font || null,
        size: entry.size ?? null,
        sizeLabel: entry.sizeLabel || null,
        color: entry.color || null,
      };
    }
    return { fieldMeta: map, occStyles };
  } catch (e) {
    console.warn("template/fields.json parse failed:", e);
    return { fieldMeta: new Map(), occStyles: new Map() };
  }
}

// Re-export PizZip for callers that need it
export { PizZip };
