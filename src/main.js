import { open, save } from "@tauri-apps/plugin-dialog";
import { invoke } from "@tauri-apps/api/core";
import { open as shellOpen } from "@tauri-apps/plugin-shell";
import { writeTextFile, mkdir, exists } from "@tauri-apps/plugin-fs";
import { renderAsync } from "docx-preview";
import * as XLSX from "xlsx";
import {
  PizZip,
  extractFieldsFromZip,
  parseParagraphs,
  renderFilled,
  buildTemplate,
  readFieldMeta,
  getRunStyleAt,
} from "./docx.js";

// ============================================================
// State
// ============================================================
const state = {
  mode: "edit", // 'edit' | 'fill'
  templateBytes: null,
  filename: "",
  isTemplateInput: false,

  // edit / modify mode (shared)
  paragraphs: [],

  // fill mode
  fields: [],
  values: {},

  // Field metadata shared between edit and fill modes.
  // Map<name, { type: 'text'|'image', description: string }>
  // Persisted into the docx as `template/fields.json`.
  fieldMeta: new Map(),

  // Snapshot of placeholders that existed in the source text before editing.
  // Used to avoid deleting literal text like {@name} that was already in doc.
  // Map<paragraphIndex, Map<"@name"|"%name", count>>
  originalPlaceholderBaseline: new Map(),

  // Per-occurrence styling for placeholder tokens.
  // Map<paragraphIndex, Array<{ font, size, sizeLabel, color }>>
  // Each array entry corresponds to the Nth managed placeholder in that paragraph,
  // in order of appearance in the text.
  occurrenceStyles: new Map(),
};

const DEFAULT_IMAGE_CONFIG = {
  fitMode: "width",
  maintainRatio: true,
  maxWidth: 8.0,
  maxHeight: 10.58,
  minWidth: 1.32,
  minHeight: 1.32,
};

function roundCm(value) {
  return Math.round(Number(value) * 100) / 100;
}

function formatCm(value, fallback) {
  const n = Number(value);
  const v = Number.isFinite(n) && n > 0 ? n : fallback;
  return roundCm(v).toFixed(2);
}

const FALLBACK_FONTS = [
  // Common Windows / Office Chinese fonts
  "宋体",
  "新宋体",
  "黑体",
  "微软雅黑",
  "微软雅黑 Light",
  "楷体",
  "楷体_GB2312",
  "仿宋",
  "仿宋_GB2312",
  "隶书",
  "幼圆",
  "等线",
  "等线 Light",
  "华文中宋",
  "华文宋体",
  "华文黑体",
  "华文楷体",
  "华文细黑",
  "华文新魏",
  "华文行楷",
  "华文琥珀",
  "华文隶书",
  "华文彩云",
  "方正书宋_GBK",
  "方正姚体",
  "方正舒体",
  "Arial",
  "Arial Black",
  "Arial Narrow",
  "Calibri",
  "Calibri Light",
  "Cambria",
  "Cambria Math",
  "Candara",
  "Comic Sans MS",
  "Consolas",
  "Constantia",
  "Corbel",
  "Courier New",
  "Georgia",
  "Helvetica",
  "Impact",
  "Lucida Console",
  "Lucida Sans Unicode",
  "Microsoft Sans Serif",
  "Palatino Linotype",
  "Segoe UI",
  "Segoe UI Light",
  "Tahoma",
  "Times New Roman",
  "Trebuchet MS",
  "Verdana",
];

let cachedFonts = null;
async function getFonts() {
  if (cachedFonts) return cachedFonts;

  // 1) Preferred: ask Rust to enumerate fonts. This walks the system font
  //    dirs and parses each TTF/OTF/TTC `name` table directly, so:
  //      - no permission prompt is needed (queryLocalFonts requires one)
  //      - we get the same localized family names Word shows (e.g.
  //        "方正小标宋简体" rather than the PostScript / English name)
  //      - per-user fonts (%LOCALAPPDATA%\Microsoft\Windows\Fonts) are
  //        included, which queryLocalFonts can miss.
  try {
    const fonts = await invoke("list_fonts");
    if (Array.isArray(fonts) && fonts.length > 0) {
      fonts.sort((a, b) => {
        const aZh = /[\u4e00-\u9fa5]/.test(a);
        const bZh = /[\u4e00-\u9fa5]/.test(b);
        if (aZh !== bZh) return aZh ? -1 : 1;
        return a.localeCompare(b, "zh-Hans-CN");
      });
      cachedFonts = fonts;
      console.info(`Loaded ${fonts.length} font families from Rust`);
      return cachedFonts;
    }
  } catch (e) {
    console.warn("list_fonts failed; falling back:", e);
  }

  // 2) Fallback: queryLocalFonts (Chromium Local Font Access API). May show
  //    a permission prompt and may miss some fonts.
  if (
    typeof window !== "undefined" &&
    typeof window.queryLocalFonts === "function"
  ) {
    try {
      const fonts = await window.queryLocalFonts();
      const families = [...new Set(fonts.map((f) => f.family))];
      families.sort((a, b) => {
        const aZh = /[\u4e00-\u9fa5]/.test(a);
        const bZh = /[\u4e00-\u9fa5]/.test(b);
        if (aZh !== bZh) return aZh ? -1 : 1;
        return a.localeCompare(b, "zh-Hans-CN");
      });
      cachedFonts = families;
      console.info(`Loaded ${families.length} font families from queryLocalFonts`);
      return cachedFonts;
    } catch (e) {
      console.warn("queryLocalFonts failed; falling back to preset list:", e);
    }
  }

  // 3) Last resort: a baked-in list of common fonts.
  cachedFonts = FALLBACK_FONTS.slice();
  return cachedFonts;
}

// Word-style font sizes. `value` is pt (used in OOXML). `label` is what we
// show in the dropdown — both Chinese name sizes (初号..八号) and numeric
// pt values, mirroring what Word's font-size combobox shows.
const SIZE_PRESETS_ZH = [
  { value: 42, label: "初号" },
  { value: 36, label: "小初" },
  { value: 26, label: "一号" },
  { value: 24, label: "小一" },
  { value: 22, label: "二号" },
  { value: 18, label: "小二" },
  { value: 16, label: "三号" },
  { value: 15, label: "小三" },
  { value: 14, label: "四号" },
  { value: 12, label: "小四" },
  { value: 10.5, label: "五号" },
  { value: 9, label: "小五" },
  { value: 7.5, label: "六号" },
  { value: 6.5, label: "小六" },
  { value: 5.5, label: "七号" },
  { value: 5, label: "八号" },
];
const SIZE_PRESETS_NUM = [
  { value: 5, label: "5" },
  { value: 5.5, label: "5.5" },
  { value: 6.5, label: "6.5" },
  { value: 7.5, label: "7.5" },
  { value: 8, label: "8" },
  { value: 9, label: "9" },
  { value: 10, label: "10" },
  { value: 10.5, label: "10.5" },
  { value: 11, label: "11" },
  { value: 12, label: "12" },
  { value: 14, label: "14" },
  { value: 16, label: "16" },
  { value: 18, label: "18" },
  { value: 20, label: "20" },
  { value: 22, label: "22" },
  { value: 24, label: "24" },
  { value: 26, label: "26" },
  { value: 28, label: "28" },
  { value: 36, label: "36" },
  { value: 48, label: "48" },
  { value: 72, label: "72" },
];
const ALL_SIZE_PRESETS = [...SIZE_PRESETS_ZH, ...SIZE_PRESETS_NUM];

// ============================================================
// DOM
// ============================================================
const els = {
  tabs: document.querySelectorAll(".mode-tabs .tab"),
  modeEdit: document.getElementById("mode-edit"),
  modeFill: document.getElementById("mode-fill"),

  btnEditLoad: document.getElementById("btn-edit-load"),
  editFilename: document.getElementById("edit-filename"),
  previewContainer: document.getElementById("preview-container"),
  btnRefreshPreview: document.getElementById("btn-refresh-preview"),
  btnInsertText: document.getElementById("btn-insert-text"),
  btnInsertImage: document.getElementById("btn-insert-image"),
  selectionHint: document.getElementById("selection-hint"),
  paragraphList: document.getElementById("paragraph-list"),
  fieldSummaryInline: document.getElementById("field-summary-inline"),
  btnEditSave: document.getElementById("btn-edit-save"),
  previewZoomControls: document.getElementById("preview-zoom-controls"),
  btnZoomOut: document.getElementById("btn-zoom-out"),
  btnZoomReset: document.getElementById("btn-zoom-reset"),
  btnZoomIn: document.getElementById("btn-zoom-in"),
  zoomValue: document.getElementById("zoom-value"),

  btnFillLoad: document.getElementById("btn-fill-load"),
  fillFilename: document.getElementById("fill-filename"),
  formSection: document.getElementById("form-section"),
  btnFillExport: document.getElementById("btn-fill-export"),
  btnBatchExportTemplate: document.getElementById("btn-batch-export-template"),
  btnBatchImport: document.getElementById("btn-batch-import"),

  status: document.getElementById("status"),
};

const PREVIEW_ZOOM_MIN = 0.5;
const PREVIEW_ZOOM_MAX = 2.5;
const PREVIEW_ZOOM_STEP = 0.1;
let previewScale = 1;
let previewFitWidth = true;
let previewRefreshTimer = null;
let previewRefreshInFlight = false;

function clamp(n, min, max) {
  return Math.max(min, Math.min(max, n));
}

function getPreviewDocxElements() {
  return [...els.previewContainer.querySelectorAll(".docx")];
}

function getPreviewWrapperElement() {
  return els.previewContainer.querySelector(".docx-wrapper");
}

function updateZoomUI() {
  if (!els.zoomValue) return;
  if (previewFitWidth) {
    els.zoomValue.textContent = "适应";
  } else {
    els.zoomValue.textContent = `${Math.round(previewScale * 100)}%`;
  }
}

function computeFitScale() {
  const docs = getPreviewDocxElements();
  const docx = docs[0];
  if (!docx) return 1;
  const containerWidth = els.previewContainer.clientWidth - 36;
  const naturalWidth = docx.offsetWidth || 1;
  if (containerWidth <= 0 || naturalWidth <= 0) return 1;
  return clamp(containerWidth / naturalWidth, PREVIEW_ZOOM_MIN, PREVIEW_ZOOM_MAX);
}

function applyPreviewZoom() {
  const docs = getPreviewDocxElements();
  if (docs.length === 0) return;
  const wrapper = getPreviewWrapperElement();
  const scale = previewFitWidth ? computeFitScale() : previewScale;
  // Use CSS zoom instead of transform: zoom affects layout width/height,
  // so fit-width won't leave phantom horizontal scrollbars.
  if (wrapper) {
    wrapper.style.zoom = String(scale);
  }
  for (const docx of docs) {
    docx.style.transform = "none";
    docx.style.marginBottom = "14px";
  }
  updateZoomUI();
}

function setPreviewScale(scale) {
  previewFitWidth = false;
  previewScale = clamp(scale, PREVIEW_ZOOM_MIN, PREVIEW_ZOOM_MAX);
  applyPreviewZoom();
}

function setPreviewFitWidth() {
  previewFitWidth = true;
  applyPreviewZoom();
}

function buildOriginalPlaceholderBaseline() {
  const baseline = new Map();
  for (const p of state.paragraphs) {
    const counts = new Map();
    for (const m of p.originalText.matchAll(/\{([@%])(\w+)\}/g)) {
      const key = `${m[1]}${m[2]}`;
      counts.set(key, (counts.get(key) || 0) + 1);
    }
    baseline.set(p.index, counts);
  }
  state.originalPlaceholderBaseline = baseline;
  syncOccurrenceStyles();
}

function getTemplateFieldType(name) {
  const meta = state.fieldMeta.get(name);
  return meta?.type || null;
}

function syncOccurrenceStyles() {
  for (const p of state.paragraphs) {
    const ranges = getPlaceholderRanges(p.currentText, p.index).filter((r) => r.managed);
    const existing = state.occurrenceStyles.get(p.index) || [];
    const next = ranges.map((r, i) => {
      if (existing[i]) {
        return { ...existing[i], name: r.name, sigil: r.sigil };
      }
      // For new entries, use fieldMeta description as default
      const meta = state.fieldMeta.get(r.name);
      return { name: r.name, sigil: r.sigil, font: null, size: null, sizeLabel: null, color: null, description: meta?.description || null };
    });
    state.occurrenceStyles.set(p.index, next);
  }
}

function isPlaceholderTokenManaged(sigil, name, occurrenceInParagraph, paragraphIndex) {
  const expectedType = sigil === "%" ? "image" : "text";
  const metaType = getTemplateFieldType(name);
  if (metaType !== expectedType) return false;
  if (state.isTemplateInput) return true;
  const baseline = state.originalPlaceholderBaseline.get(paragraphIndex) || new Map();
  const protectedCount = baseline.get(`${sigil}${name}`) || 0;
  return occurrenceInParagraph >= protectedCount;
}

function removeManagedPlaceholderFromText(text, sigil, name, protectedCount) {
  const re = new RegExp(`\\{${escapeRegex(sigil)}${escapeRegex(name)}\\}`, "g");
  let idx = 0;
  return text.replace(re, (full) => {
    const keep = idx < protectedCount;
    idx += 1;
    return keep ? full : "";
  });
}

function getPlaceholderRanges(text, paragraphIndex) {
  const seen = new Map();
  const ranges = [];
  const re = /\{([@%])(\w+)\}/g;
  let m;
  while ((m = re.exec(text)) !== null) {
    const sigil = m[1];
    const name = m[2];
    const key = `${sigil}${name}`;
    const occ = seen.get(key) || 0;
    seen.set(key, occ + 1);
    ranges.push({
      start: m.index,
      end: m.index + m[0].length,
      text: m[0],
      managed: isPlaceholderTokenManaged(sigil, name, occ, paragraphIndex),
      sigil,
      name,
    });
  }
  return ranges;
}

function escapeHtml(s) {
  return String(s)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#39;");
}

function textToEditorHtml(text, paragraphIndex) {
  const ranges = getPlaceholderRanges(text, paragraphIndex);
  let out = "";
  let cursor = 0;
  for (const r of ranges) {
    if (r.start > cursor) {
      out += escapeHtml(text.slice(cursor, r.start));
    }
    if (r.managed) {
      out += `<span class="placeholder-token ${r.sigil === "%" ? "image" : "text"}" contenteditable="false" data-token="1" data-sigil="${r.sigil}" data-name="${escapeHtml(r.name)}">${escapeHtml(r.text)}</span>`;
    } else {
      out += escapeHtml(r.text);
    }
    cursor = r.end;
  }
  if (cursor < text.length) out += escapeHtml(text.slice(cursor));
  return out.replace(/\n/g, "<br>");
}

function editorTextContent(editor) {
  let out = "";
  const walk = (node) => {
    if (node.nodeType === Node.TEXT_NODE) {
      out += node.textContent || "";
      return;
    }
    if (node.nodeType !== Node.ELEMENT_NODE) return;
    if (node.tagName === "BR") {
      out += "\n";
      return;
    }
    if ((node.tagName === "DIV" || node.tagName === "P") && node !== editor) {
      if (out.length > 0 && !out.endsWith("\n")) out += "\n";
    }
    for (const child of node.childNodes) walk(child);
  };
  for (const child of editor.childNodes) walk(child);
  return out;
}

function closeTokenContextMenu() {
  const menu = document.getElementById("token-context-menu");
  if (menu) menu.remove();
}

function getTokenOccurrenceIndex(editor, tokenNode) {
  const tokens = editor.querySelectorAll("[data-token='1']");
  let idx = 0;
  for (const t of tokens) {
    if (t === tokenNode) return idx;
    idx++;
  }
  return -1;
}

let tokenTooltipTimer = null;
let tokenTooltipEl = null;

function removeTokenTooltip() {
  if (tokenTooltipTimer) { clearTimeout(tokenTooltipTimer); tokenTooltipTimer = null; }
  if (tokenTooltipEl) { tokenTooltipEl.remove(); tokenTooltipEl = null; }
}

function showTokenTooltip(tokenNode, paragraphIndex) {
  removeTokenTooltip();
  tokenTooltipTimer = setTimeout(() => {
    const sigil = tokenNode.dataset.sigil;
    const name = tokenNode.dataset.name;
    const typeLabel = sigil === "%" ? "图片" : "文字";
    const typeClass = sigil === "%" ? "image" : "text";
    const occIdx = getTokenOccurrenceIndex(tokenNode.closest(".paragraph-editor"), tokenNode);
    const styles = state.occurrenceStyles.get(paragraphIndex) || [];
    const s = styles[occIdx] || {};

    const el = document.createElement("div");
    el.className = "token-tooltip";
    let html = `<span class="tt-name">{${sigil}${name}}</span><span class="tt-type ${typeClass}">${typeLabel}</span>`;
    if (typeClass === "text") {
      const parts = [];
      if (s.font) parts.push(`字体: ${s.font}`);
      if (s.size) parts.push(`字号: ${s.size}pt`);
      if (s.color) parts.push(`颜色: ${s.color}`);
      if (parts.length > 0) html += `<div class="tt-style">${parts.join(" · ")}</div>`;
    }
    html += `<div class="tt-style">双击编辑 · 右键删除</div>`;
    el.innerHTML = html;
    document.body.appendChild(el);
    tokenTooltipEl = el;

    const rect = tokenNode.getBoundingClientRect();
    el.style.left = `${rect.left}px`;
    el.style.top = `${rect.bottom + 6}px`;
  }, 500);
}

function openTokenContextMenu(x, y, onDelete) {
  closeTokenContextMenu();
  const menu = document.createElement("div");
  menu.id = "token-context-menu";
  menu.className = "token-context-menu";
  const btn = document.createElement("button");
  btn.type = "button";
  btn.textContent = "删除占位符";
  btn.addEventListener("click", () => {
    onDelete();
    closeTokenContextMenu();
  });
  menu.appendChild(btn);
  menu.style.left = `${x}px`;
  menu.style.top = `${y}px`;
  document.body.appendChild(menu);
  setTimeout(() => {
    const onDocClick = (ev) => {
      if (!menu.contains(ev.target)) closeTokenContextMenu();
      document.removeEventListener("mousedown", onDocClick, true);
    };
    document.addEventListener("mousedown", onDocClick, true);
  }, 0);
}

function removeManagedTokenAtSelection(text, paragraphIndex, selStart, selEnd, key) {
  const ranges = getPlaceholderRanges(text, paragraphIndex).filter((r) => r.managed);
  if (ranges.length === 0) return null;

  let targets = [];
  if (selStart !== selEnd) {
    targets = ranges.filter((r) => !(selEnd <= r.start || selStart >= r.end));
  } else if (key === "Backspace") {
    const pos = selStart;
    targets = ranges.filter((r) => pos > r.start && pos <= r.end);
  } else if (key === "Delete") {
    const pos = selStart;
    targets = ranges.filter((r) => pos >= r.start && pos < r.end);
  }
  if (targets.length === 0) return null;

  targets.sort((a, b) => b.start - a.start);
  let out = text;
  for (const t of targets) {
    out = out.slice(0, t.start) + out.slice(t.end);
  }
  const caret = Math.min(...targets.map((t) => t.start));
  return { text: out, caret };
}

function schedulePreviewRefresh() {
  if (!state.templateBytes || state.paragraphs.length === 0) return;
  if (previewRefreshInFlight) return;
  if (previewRefreshTimer) clearTimeout(previewRefreshTimer);
  previewRefreshTimer = setTimeout(async () => {
    previewRefreshTimer = null;
    previewRefreshInFlight = true;
    try {
      syncFieldMetaFromText();
      const updated = buildTemplate(
        state.templateBytes,
        state.paragraphs,
        state.fieldMeta,
        state.occurrenceStyles,
      );
      await renderPreview(updated);
    } catch (e) {
      console.error("auto preview refresh failed", e);
    } finally {
      previewRefreshInFlight = false;
    }
  }, 180);
}

// ============================================================
// Utility
// ============================================================
function setStatus(msg, kind = "") {
  els.status.textContent = msg;
  els.status.className = "status" + (kind ? " " + kind : "");
}

function fileToUint8Array(file) {
  return new Promise((resolve, reject) => {
    const r = new FileReader();
    r.onload = () => resolve(new Uint8Array(r.result));
    r.onerror = () => reject(r.error);
    r.readAsArrayBuffer(file);
  });
}

async function pickDocx(title) {
  const sel = await open({
    multiple: false,
    title,
    filters: [{ name: "Word Document", extensions: ["docx"] }],
  });
  if (!sel) return null;
  return typeof sel === "string" ? sel : sel.path;
}

async function readBytesFromPath(path) {
  const arr = await invoke("read_file_bytes", { path });
  return new Uint8Array(arr);
}

async function saveBytesViaDialog(suggestedName, bytes) {
  const target = await save({
    defaultPath: suggestedName,
    filters: [{ name: "Word Document", extensions: ["docx"] }],
  });
  if (!target) return null;
  await invoke("save_bytes", { path: target, bytes: Array.from(bytes) });
  return target;
}

function basename(path) {
  return path.split(/[\\/]/).pop();
}

// ============================================================
// Mode switching
// ============================================================
// "edit" is for both plain docs and existing templates; we auto-detect which
// one is loaded and adjust save behavior accordingly.

function switchMode(mode) {
  state.mode = mode;
  els.tabs.forEach((t) =>
    t.classList.toggle("active", t.dataset.mode === mode),
  );
  const isEditor = mode === "edit";
  els.modeEdit.classList.toggle("hidden", !isEditor);
  els.modeFill.classList.toggle("hidden", mode !== "fill");
  els.btnEditSave.classList.toggle("hidden", !isEditor);
  els.btnBatchExportTemplate.classList.toggle("hidden", !isEditor);
  els.btnFillExport.classList.toggle("hidden", mode !== "fill");
  els.btnBatchImport.classList.toggle("hidden", mode !== "fill");
  if (els.previewZoomControls) {
    els.previewZoomControls.classList.toggle("hidden", !isEditor);
  }

  setStatus("");
}

els.tabs.forEach((t) =>
  t.addEventListener("click", () => {
    switchMode(t.dataset.mode);
    // Warm the font cache on the user gesture of switching to fill mode.
    if (t.dataset.mode === "fill") getFonts();
  }),
);

// ============================================================
// Edit mode — preview selection capture and field insertion
// ============================================================
// Last captured selection inside the preview, kept even after the user clicks
// the toolbar button (which would otherwise blur the preview's selection).
let previewSelection = null; // { pIdx, start, end, fullText, selectedText }

function findParagraphElement(node) {
  let cur = node;
  while (cur) {
    if (cur.nodeType === 1 && cur.dataset?.pIdx != null) return cur;
    cur = cur.parentNode;
  }
  return null;
}

// Compute the character offset of (container, offset) within rootElement, by
// walking text nodes in document order.
function getCharOffsetWithin(rootElement, container, offset) {
  if (container === rootElement) {
    let count = 0;
    for (let i = 0; i < offset && i < rootElement.childNodes.length; i++) {
      count += rootElement.childNodes[i].textContent?.length ?? 0;
    }
    return count;
  }
  let count = 0;
  const walker = document.createTreeWalker(rootElement, NodeFilter.SHOW_TEXT);
  let n;
  while ((n = walker.nextNode())) {
    if (n === container) return count + offset;
    count += n.textContent.length;
  }
  // Container might be an element node; fall back to range-style offset
  if (container.nodeType === 1) {
    let c = 0;
    for (let i = 0; i < offset && i < container.childNodes.length; i++) {
      c += container.childNodes[i].textContent?.length ?? 0;
    }
    // Need to find the start position of `container` within rootElement
    let pre = 0;
    const w2 = document.createTreeWalker(rootElement, NodeFilter.SHOW_TEXT);
    let node2;
    while ((node2 = w2.nextNode())) {
      if (container.contains(node2)) break;
      pre += node2.textContent.length;
    }
    return pre + c;
  }
  return null;
}

function captureSelection() {
  const sel = window.getSelection();
  if (!sel || sel.rangeCount === 0) return; // keep last
  const range = sel.getRangeAt(0);
  if (!els.previewContainer.contains(range.commonAncestorContainer)) {
    return; // selection moved elsewhere — keep last preview selection
  }

  const startP = findParagraphElement(range.startContainer);
  if (!startP) {
    previewSelection = null;
    updateSelectionUI();
    return;
  }
  const pIdx = Number(startP.dataset.pIdx);
  if (pIdx >= state.paragraphs.length) {
    previewSelection = null;
    updateSelectionUI();
    return;
  }

  const startOffset = getCharOffsetWithin(
    startP,
    range.startContainer,
    range.startOffset,
  );
  if (startOffset == null) {
    previewSelection = null;
    updateSelectionUI();
    return;
  }
  const endP = findParagraphElement(range.endContainer);
  let endOffset;
  if (endP === startP) {
    endOffset = getCharOffsetWithin(startP, range.endContainer, range.endOffset);
    if (endOffset == null) endOffset = startOffset;
  } else {
    // Selection crosses paragraphs — clip to start paragraph end
    endOffset = startP.textContent.length;
  }

  const s = Math.min(startOffset, endOffset);
  const e = Math.max(startOffset, endOffset);
  const fullText = startP.textContent;
  const selectedText = fullText.slice(s, e);

  previewSelection = { pIdx, start: s, end: e, fullText, selectedText };
  updateSelectionUI();
}

function updateSelectionUI() {
  const has = !!previewSelection;
  els.btnInsertText.disabled = !has;
  els.btnInsertImage.disabled = !has;
  if (has) {
    els.selectionHint.classList.add("has-selection");
    const t = previewSelection.selectedText;
    if (t) {
      const trim = t.length > 20 ? t.slice(0, 20) + "…" : t;
      els.selectionHint.textContent = `选中 "${trim}" (${t.length} 字)`;
    } else {
      els.selectionHint.textContent = `光标在第 ${previewSelection.pIdx + 1} 段`;
    }
  } else {
    els.selectionHint.classList.remove("has-selection");
    els.selectionHint.textContent = "在下方文档里点光标或选中文字";
  }
}

function clearPreviewSelection() {
  previewSelection = null;
  updateSelectionUI();
}

async function insertAtSelection(type) {
  if (!previewSelection) return;
  const { pIdx, start, end, fullText } = previewSelection;

  // Auto-detect formatting at the cursor position so the dialog can offer
  // it as the field's default format (font/size/color).
  const p = state.paragraphs[pIdx];
  const detected = getRunStyleAt(p.originalXml, start);

  const result = await openFieldDialog({
    type,
    defaultFont: detected?.font,
    defaultSize: detected?.size,
    defaultColor: detected?.color,
    detectedFromXml: !!detected,
  });
  if (!result) return;
  const {
    name,
    type: fieldType,
    description,
    defaultFont,
    defaultSize,
    defaultSizeLabel,
    defaultColor,
  } = result;

  state.fieldMeta.set(name, {
    type: fieldType,
    description,
    ...(fieldType === "image" && result.imageConfig ? { imageConfig: result.imageConfig } : {}),
  });

  const placeholder = fieldType === "image" ? `{%${name}}` : `{@${name}}`;
  const newText = fullText.slice(0, start) + placeholder + fullText.slice(end);

  p.currentText = newText;
  p.dirty = newText !== p.originalText;

  // Sync occurrence styles and store per-occurrence description
  syncOccurrenceStyles();
  const occStyles = state.occurrenceStyles.get(pIdx) || [];
  const ranges = getPlaceholderRanges(newText, pIdx).filter((r) => r.managed);
  const newOccIdx = ranges.findIndex((r) => r.name === name);
  if (newOccIdx >= 0 && occStyles[newOccIdx]) {
    occStyles[newOccIdx].description = description || null;
  }

  renderParagraphList();
  updateFieldSummary();
  clearPreviewSelection();
  await refreshPreview();
  focusEditCard(pIdx);
}

// listen globally — selectionchange fires when selection moves
document.addEventListener("selectionchange", captureSelection);

// Prevent buttons from stealing focus / clearing the document selection
[els.btnInsertText, els.btnInsertImage].forEach((b) => {
  b.addEventListener("mousedown", (e) => e.preventDefault());
});
els.btnInsertText.addEventListener("click", () => insertAtSelection("text"));
els.btnInsertImage.addEventListener("click", () => insertAtSelection("image"));

// ============================================================
// Edit mode — preview rendering
// ============================================================
function fixPreviewImages(container) {
  const allImgs = container.querySelectorAll("img");
  for (const img of allImgs) {
    let el = img.parentElement;
    // Walk up a few levels to find the absolute-positioned wrapper
    // that docx-preview creates for anchored images
    for (let i = 0; i < 4 && el; i++) {
      const style = el.getAttribute("style") || "";
      if (/position\s*:\s*absolute/i.test(style)) {
        el.style.position = "relative";
        el.style.removeProperty("left");
        el.style.removeProperty("top");
        el.style.removeProperty("right");
        el.style.removeProperty("bottom");
        break;
      }
      el = el.parentElement;
    }
  }
}

async function renderPreview(bytes) {
  const container = els.previewContainer;
  container.innerHTML = "";
  const blob = new Blob([bytes], {
    type: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
  });
  try {
    await renderAsync(blob, container, undefined, {
      inWrapper: true,
      breakPages: true,
      ignoreLastRenderedPageBreak: true,
      renderHeaders: false,
      renderFooters: false,
      renderFootnotes: false,
      renderEndnotes: false,
      experimental: false,
      // Keep the document's own page size/margins so preview matches Word
      // (A4/Letter proportions and layout are driven by DOCX section settings).
      ignoreWidth: false,
      ignoreHeight: false,
    });
  } catch (e) {
    console.error("renderAsync failed", e);
    container.innerHTML =
      '<p class="empty-state">预览渲染失败：' +
      (e?.message || e) +
      "</p>";
    return;
  }

  // Post-process images to fix layout issues
  fixPreviewImages(container);

  // Annotate paragraphs in document order so they map to state.paragraphs[i]
  const ps = container.querySelectorAll("p");
  const limit = Math.min(ps.length, state.paragraphs.length);
  for (let i = 0; i < limit; i++) {
    const p = ps[i];
    p.dataset.pIdx = String(i);
    p.classList.add("preview-paragraph");
    p.addEventListener("click", () => focusEditCard(i));
  }
  if (ps.length !== state.paragraphs.length) {
    console.warn(
      `paragraph count mismatch: preview=${ps.length}, parsed=${state.paragraphs.length}. ` +
        "First " +
        limit +
        " will be linked.",
    );
  }

  // First render uses fit-width so users see complete page without horizontal scroll.
  if (previewFitWidth) {
    setPreviewFitWidth();
  } else {
    applyPreviewZoom();
  }
}

function clearActivePreview() {
  els.previewContainer
    .querySelectorAll(".preview-paragraph.active")
    .forEach((p) => p.classList.remove("active"));
}

function autoResizeEditor(editor) {
  if (!editor) return;
  editor.style.height = "auto";
  editor.style.height = `${editor.scrollHeight}px`;
}

function highlightPreview(idx) {
  clearActivePreview();
  const p = els.previewContainer.querySelector(
    `.preview-paragraph[data-p-idx="${idx}"]`,
  );
  if (p) p.classList.add("active");
}

function focusEditCard(idx) {
  const card = els.paragraphList.querySelector(
    `.paragraph-card[data-p-idx="${idx}"]`,
  );
  if (!card) return;
  card.scrollIntoView({ behavior: "smooth", block: "center" });
  card.classList.add("highlighted");
  setTimeout(() => card.classList.remove("highlighted"), 1200);
  const ed = card.querySelector(".paragraph-editor");
  if (ed) ed.focus();
  highlightPreview(idx);
}

// ============================================================
// Edit mode — paragraph cards
// ============================================================
function detectFieldsInText(text) {
  const re = /\{([@%])(\w+)\}/g;
  const out = [];
  let m;
  while ((m = re.exec(text)) !== null) {
    out.push({ name: m[2], type: m[1] === "%" ? "image" : "text" });
  }
  return out;
}

function updateFieldSummary() {
  // Combine fields detected in text + persisted metadata
  const map = new Map();
  for (const p of state.paragraphs) {
    for (const f of detectFieldsInText(p.currentText)) {
      if (!map.has(f.name)) map.set(f.name, { type: f.type, used: true });
    }
  }
  for (const [name, m] of state.fieldMeta) {
    if (!map.has(name)) map.set(name, { type: m.type, used: false });
  }

  els.fieldSummaryInline.innerHTML = "";
  if (map.size === 0) {
    const span = document.createElement("span");
    span.className = "field-summary-inline-text";
    span.style.opacity = "0.7";
    span.textContent = "（尚无字段，先在预览里选位置插入）";
    els.fieldSummaryInline.appendChild(span);
    return;
  }

  const label = document.createElement("span");
  label.style.opacity = "0.7";
  label.style.marginRight = "4px";
  label.textContent = `${map.size} 个字段：`;
  els.fieldSummaryInline.appendChild(label);

  for (const [name, info] of map) {
    const meta = state.fieldMeta.get(name);
    const occ = info.used ? countOccurrences(name) : 0;

    const wrap = document.createElement("span");
    wrap.className = `field-tag-wrap ${info.type}`;
    if (!info.used) wrap.classList.add("untyped");

    // Name button: opens edit dialog (where user can also rename)
    const tag = document.createElement("button");
    tag.type = "button";
    tag.className = "field-tag-name";
    tag.innerHTML = `<span class="field-tag-text">${name}</span>${
      occ > 0 ? `<span class="field-tag-count">×${occ}</span>` : ""
    }`;
    tag.title =
      (meta?.description ? `描述：${meta.description}\n` : "（点击编辑/重命名）\n") +
      (info.used
        ? `在文档中出现 ${occ} 次`
        : "尚未在文档里出现（仅元数据）");
    tag.addEventListener("click", async () => {
      const updated = await openFieldDialog({
        title: `编辑字段 "${name}"`,
        type: info.type,
        name,
        description: meta?.description || "",
        imageConfig: meta?.imageConfig || null,
        // Name is editable — submit handler validates & runs renameField.
        lockName: false,
        // Type can't change once placeholders exist in the doc, since
        // {@x} and {%x} are not interchangeable mid-document.
        lockType: info.used,
        existingName: name,
        hideDescription: false,
        hideFormat: true,
        hideImageConfig: false,
      });
      if (!updated) return;

      // Name was changed → propagate rename across paragraphs + metadata
      if (updated.name !== name) {
        renameField(name, updated.name);
      }
      const finalMeta = {
        type: updated.type,
        description: updated.description,
      };
      if (updated.imageConfig) {
        finalMeta.imageConfig = updated.imageConfig;
      }
      state.fieldMeta.set(updated.name, finalMeta);

      renderParagraphList();
      updateFieldSummary();
      await refreshPreview();
    });

    // Delete button: strip placeholders + metadata
    const del = document.createElement("button");
    del.type = "button";
    del.className = "field-tag-del";
    del.textContent = "×";
    del.title = "删除此字段";
    del.addEventListener("click", async (e) => {
      e.stopPropagation();
      const msg =
        occ > 0
          ? `删除字段 "${name}"？\n这会从文档里移除 ${occ} 处占位符并删除元数据，不可撤销。`
          : `删除字段元数据 "${name}"？`;
      if (!await showConfirm("删除字段", msg)) return;
      deleteField(name);
      renderParagraphList();
      updateFieldSummary();
      await refreshPreview();
    });

    wrap.append(tag, del);
    els.fieldSummaryInline.appendChild(wrap);
  }
}

function placeCaretAtTextOffset(root, offset) {
  const sel = window.getSelection();
  if (!sel) return;
  let remain = Math.max(0, offset);
  const walker = document.createTreeWalker(root, NodeFilter.SHOW_TEXT);
  let node = null;
  while ((node = walker.nextNode())) {
    const len = node.textContent?.length || 0;
    if (remain <= len) {
      const r = document.createRange();
      r.setStart(node, remain);
      r.collapse(true);
      sel.removeAllRanges();
      sel.addRange(r);
      return;
    }
    remain -= len;
  }
  const r = document.createRange();
  r.selectNodeContents(root);
  r.collapse(false);
  sel.removeAllRanges();
  sel.addRange(r);
}

function getEditorSelectionOffsets(editor) {
  const sel = window.getSelection();
  if (!sel || sel.rangeCount === 0) {
    const t = editorTextContent(editor);
    return { start: t.length, end: t.length };
  }
  const range = sel.getRangeAt(0);
  if (!editor.contains(range.startContainer) || !editor.contains(range.endContainer)) {
    const t = editorTextContent(editor);
    return { start: t.length, end: t.length };
  }
  const a = getCharOffsetWithin(editor, range.startContainer, range.startOffset) ?? 0;
  const b = getCharOffsetWithin(editor, range.endContainer, range.endOffset) ?? a;
  return a <= b ? { start: a, end: b } : { start: b, end: a };
}

function insertAtCursor(editor, snippet, paragraphIndex, overrideStart, overrideEnd) {
  const offsets = (overrideStart != null && overrideEnd != null)
    ? { start: overrideStart, end: overrideEnd }
    : getEditorSelectionOffsets(editor);
  const { start, end } = offsets;
  const oldText = editorTextContent(editor);
  const next = oldText.slice(0, start) + snippet + oldText.slice(end);
  editor.innerHTML = textToEditorHtml(next, paragraphIndex);
  placeCaretAtTextOffset(editor, start + snippet.length);
  editor.focus();
  editor.dispatchEvent(new Event("input", { bubbles: true }));
}

// ============================================================
// Field metadata helpers + add/edit modal
// ============================================================
function syncFieldMetaFromText() {
  // Auto-register any {@xxx}/{%xxx} the user typed directly into a card
  for (const p of state.paragraphs) {
    for (const m of p.currentText.matchAll(/\{([@%])(\w+)\}/g)) {
      const name = m[2];
      const type = m[1] === "%" ? "image" : "text";
      if (!state.fieldMeta.has(name)) {
        state.fieldMeta.set(name, { type, description: "" });
      }
    }
  }
}

function countOccurrences(name) {
  const re = new RegExp(`\\{[@%]${name}\\}`, "g");
  let total = 0;
  for (const p of state.paragraphs) {
    total += (p.currentText.match(re) || []).length;
  }
  return total;
}

function escapeRegex(s) {
  return s.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
}

// Rename a field across the entire document AND its metadata entry.
// Returns true if anything changed. Caller is responsible for refreshing UI.
function renameField(oldName, newName) {
  if (oldName === newName) return false;
  const re = new RegExp(`\\{([@%])${escapeRegex(oldName)}\\}`, "g");
  let touched = false;
  for (const p of state.paragraphs) {
    const next = p.currentText.replace(re, `{$1${newName}}`);
    if (next !== p.currentText) {
      p.currentText = next;
      p.dirty = next !== p.originalText;
      touched = true;
    }
  }
  // Update name in occurrence styles
  for (const [, styles] of state.occurrenceStyles) {
    for (const s of styles) {
      if (s && s.name === oldName) s.name = newName;
    }
  }
  // Transfer metadata key
  const meta = state.fieldMeta.get(oldName);
  state.fieldMeta.delete(oldName);
  if (meta) state.fieldMeta.set(newName, meta);
  return touched || !!meta;
}

// Delete a field: strip {@name} / {%name} from every paragraph and drop
// the metadata entry. Returns true if anything changed.
function deleteField(name) {
  let touched = false;
  for (const p of state.paragraphs) {
    const baseline = state.originalPlaceholderBaseline.get(p.index) || new Map();
    const keepTextCount = baseline.get(`@${name}`) || 0;
    const keepImageCount = baseline.get(`%${name}`) || 0;
    let next = p.currentText;
    next = removeManagedPlaceholderFromText(next, "@", name, keepTextCount);
    next = removeManagedPlaceholderFromText(next, "%", name, keepImageCount);
    if (next !== p.currentText) {
      p.currentText = next;
      p.dirty = next !== p.originalText;
      touched = true;
    }
  }
  if (touched) syncOccurrenceStyles();
  if (state.fieldMeta.delete(name)) touched = true;
  return touched;
}

const fieldDialogEls = {
  dlg: document.getElementById("field-dialog"),
  form: document.getElementById("field-dialog-form"),
  title: document.getElementById("field-dialog-title"),
  type: document.getElementById("field-dialog-type"),
  name: document.getElementById("field-dialog-name"),
  desc: document.getElementById("field-dialog-desc"),
  hint: document.getElementById("field-dialog-hint"),
  cancel: document.getElementById("field-dialog-cancel"),
  format: document.getElementById("field-dialog-format"),
  formatDetected: document.getElementById("field-dialog-format-detected"),
  fontSlot: document.getElementById("field-dialog-font-slot"),
  size: document.getElementById("field-dialog-size"),
  color: document.getElementById("field-dialog-color"),
  colorSwatch: document.getElementById("field-dialog-color-swatch"),
  colorHex: document.getElementById("field-dialog-color-hex"),
  imageConfig: document.getElementById("field-dialog-image-config"),
  fitMode: document.getElementById("field-dialog-fit-mode"),
  maintainRatio: document.getElementById("field-dialog-maintain-ratio"),
  maxW: document.getElementById("field-dialog-max-w"),
  maxH: document.getElementById("field-dialog-max-h"),
  minW: document.getElementById("field-dialog-min-w"),
  minH: document.getElementById("field-dialog-min-h"),
};

let dialogFontPicker = null;

function ensureDialogSizeOptions() {
  if (fieldDialogEls.size.options.length > 0) return;
  const zh = document.createElement("optgroup");
  zh.label = "字号";
  for (const s of SIZE_PRESETS_ZH) {
    const opt = document.createElement("option");
    opt.value = s.label;
    opt.textContent = s.label;
    opt.dataset.pt = String(s.value);
    zh.appendChild(opt);
  }
  const num = document.createElement("optgroup");
  num.label = "磅值";
  for (const s of SIZE_PRESETS_NUM) {
    const opt = document.createElement("option");
    opt.value = s.label;
    opt.textContent = s.label;
    opt.dataset.pt = String(s.value);
    num.appendChild(opt);
  }
  fieldDialogEls.size.append(zh, num);
}

// Open the add/edit field modal. Returns Promise<FieldMeta|null> where
// FieldMeta = { name, type, description, defaultFont, defaultSize,
//               defaultSizeLabel, defaultColor, imageConfig }.
async function openFieldDialog(defaults = {}) {
  const {
    type = "text",
    name = "",
    description = "",
    defaultFont = null,
    defaultSize = null,
    defaultSizeLabel = null,
    defaultColor = null,
    imageConfig = null,
    detectedFromXml = false, // when true, show "(已检测)" hint
    title = "添加字段",
    lockType = false,
    lockName = false,
    hideDescription = false,
    hideFormat = false,
    hideImageConfig = false,
    // When set, this dialog is editing an existing field; the submit
    // handler must allow `name` to equal `existingName` even though that
    // name is "already taken" (it's just unchanged), and warn before
    // colliding with a DIFFERENT existing field.
    existingName = null,
  } = defaults;

  // Lazy-init font picker inside dialog (needs fonts list)
  if (!dialogFontPicker) {
    const fonts = await getFonts();
    dialogFontPicker = createFontPicker(fonts, "宋体", () => {});
    fieldDialogEls.fontSlot.appendChild(dialogFontPicker.wrap);
  }
  ensureDialogSizeOptions();

  return new Promise((resolve) => {
    fieldDialogEls.title.textContent = title;
    fieldDialogEls.type.value = type;
    fieldDialogEls.type.disabled = lockType;
    fieldDialogEls.name.value = name;
    fieldDialogEls.name.disabled = lockName;
    fieldDialogEls.desc.value = description;
    const descRow = fieldDialogEls.desc.closest(".dialog-row");
    if (descRow) descRow.classList.toggle("hidden", hideDescription);
    fieldDialogEls.hint.textContent = "";
    fieldDialogEls.hint.classList.remove("exists");

    // Default format defaults
    dialogFontPicker.input.value = defaultFont || "宋体";
    const sizeLabel =
      defaultSizeLabel ||
      (defaultSize != null ? findSizeByValue(defaultSize)?.label : null) ||
      "小四";
    fieldDialogEls.size.value = sizeLabel;
    if (fieldDialogEls.size.selectedIndex < 0) {
      // size label wasn't in options; fall back to 小四
      fieldDialogEls.size.value = "小四";
    }
    fieldDialogEls.color.value = defaultColor || "#000000";
    if (fieldDialogEls.colorSwatch) {
      fieldDialogEls.colorSwatch.style.background = defaultColor || "#000000";
    }
    if (fieldDialogEls.colorHex) {
      fieldDialogEls.colorHex.textContent = (defaultColor || "#000000").toUpperCase();
    }

    fieldDialogEls.formatDetected.textContent = detectedFromXml
      ? "（已从光标位置自动读取）"
      : "";
    fieldDialogEls.format.classList.toggle("hidden", type === "image" || hideFormat);

    // Color hex display sync
    function onColorInput() {
      if (fieldDialogEls.colorSwatch) {
        fieldDialogEls.colorSwatch.style.background = fieldDialogEls.color.value;
      }
      if (fieldDialogEls.colorHex) {
        fieldDialogEls.colorHex.textContent = fieldDialogEls.color.value.toUpperCase();
      }
    }
    fieldDialogEls.color.addEventListener("input", onColorInput);

    // Image config defaults
    const ic = imageConfig || {};
    fieldDialogEls.fitMode.value = ic.fitMode || DEFAULT_IMAGE_CONFIG.fitMode;
    fieldDialogEls.maintainRatio.checked = ic.maintainRatio !== false;
    fieldDialogEls.maxW.value = formatCm(ic.maxWidth, DEFAULT_IMAGE_CONFIG.maxWidth);
    fieldDialogEls.maxH.value = formatCm(ic.maxHeight, DEFAULT_IMAGE_CONFIG.maxHeight);
    fieldDialogEls.minW.value = formatCm(ic.minWidth, DEFAULT_IMAGE_CONFIG.minWidth);
    fieldDialogEls.minH.value = formatCm(ic.minHeight, DEFAULT_IMAGE_CONFIG.minHeight);
    fieldDialogEls.imageConfig.classList.toggle("hidden", type !== "image" || hideImageConfig);

    let descTouched = !!description;

    function onTypeChange() {
      const isImage = fieldDialogEls.type.value === "image";
      fieldDialogEls.format.classList.toggle("hidden", isImage || hideFormat);
      fieldDialogEls.imageConfig.classList.toggle("hidden", !isImage || hideImageConfig);
    }
    function onNameInput() {
      const n = fieldDialogEls.name.value.trim();
      if (!n) {
        fieldDialogEls.hint.textContent = "";
        fieldDialogEls.hint.classList.remove("exists");
        return;
      }
      // When editing an existing field and the name hasn't changed, this
      // is just the field itself — don't show a duplicate warning.
      if (existingName && n === existingName) {
        fieldDialogEls.hint.textContent = "";
        fieldDialogEls.hint.classList.remove("exists");
        return;
      }
      const existing = state.fieldMeta.get(n);
      if (existing) {
        const cnType = existing.type === "image" ? "图片" : "文字";
        const occ = countOccurrences(n);
        if (existingName) {
          // Editing-rename mode: this is a collision, not "reuse"
          fieldDialogEls.hint.textContent =
            `名称 "${n}" 已被另一个字段占用 (${cnType}${occ > 0 ? `，已用 ${occ} 处` : ""})`;
          fieldDialogEls.hint.classList.remove("exists");
        } else {
          fieldDialogEls.hint.textContent =
            occ > 0
              ? `已存在 "${n}" (${cnType})，文档里已用了 ${occ} 处`
              : `已记录字段 "${n}" (${cnType})`;
          fieldDialogEls.hint.classList.add("exists");
          if (!lockType) fieldDialogEls.type.value = existing.type;
          onTypeChange();
        }
      } else {
        fieldDialogEls.hint.textContent = "";
        fieldDialogEls.hint.classList.remove("exists");
      }
    }
    function onDescInput() {
      descTouched = true;
    }
    function cleanup() {
      fieldDialogEls.type.removeEventListener("change", onTypeChange);
      fieldDialogEls.name.removeEventListener("input", onNameInput);
      fieldDialogEls.desc.removeEventListener("input", onDescInput);
      fieldDialogEls.color.removeEventListener("input", onColorInput);
      fieldDialogEls.cancel.removeEventListener("click", onCancel);
      fieldDialogEls.form.removeEventListener("submit", onSubmit);
      fieldDialogEls.dlg.removeEventListener("cancel", onCancel);
    }
    function onCancel(e) {
      e?.preventDefault?.();
      cleanup();
      fieldDialogEls.dlg.close();
      resolve(null);
    }
    function onSubmit(e) {
      e.preventDefault();
      const n = fieldDialogEls.name.value.trim();
      if (!/^\w+$/.test(n)) {
        fieldDialogEls.hint.textContent =
          "名称只能包含字母、数字、下划线（不要有空格或中文）";
        fieldDialogEls.hint.classList.remove("exists");
        fieldDialogEls.name.focus();
        return;
      }
      // Rename-collision check: if we're editing an existing field and
      // the user typed a NEW name that already belongs to a DIFFERENT
      // field, refuse — otherwise we'd merge two fields' placeholders.
      if (existingName && n !== existingName && state.fieldMeta.has(n)) {
        fieldDialogEls.hint.textContent =
          `已存在字段 "${n}"，重命名会与之冲突。请换一个名字。`;
        fieldDialogEls.hint.classList.remove("exists");
        fieldDialogEls.name.focus();
        return;
      }
      const finalType = fieldDialogEls.type.value;
      const result = {
        name: n,
        type: finalType,
        description: fieldDialogEls.desc.value.trim(),
      };
      if (finalType === "text") {
        const opt = fieldDialogEls.size.options[fieldDialogEls.size.selectedIndex];
        result.defaultFont = dialogFontPicker.input.value.trim() || null;
        result.defaultSizeLabel = opt?.value || null;
        result.defaultSize = opt ? Number(opt.dataset.pt) : null;
        result.defaultColor = fieldDialogEls.color.value || null;
      } else {
        result.imageConfig = {
          fitMode: fieldDialogEls.fitMode.value,
          maintainRatio: fieldDialogEls.maintainRatio.checked,
          maxWidth: roundCm(Number(fieldDialogEls.maxW.value) || DEFAULT_IMAGE_CONFIG.maxWidth),
          maxHeight: roundCm(Number(fieldDialogEls.maxH.value) || DEFAULT_IMAGE_CONFIG.maxHeight),
          minWidth: roundCm(Number(fieldDialogEls.minW.value) || DEFAULT_IMAGE_CONFIG.minWidth),
          minHeight: roundCm(Number(fieldDialogEls.minH.value) || DEFAULT_IMAGE_CONFIG.minHeight),
        };
      }
      cleanup();
      fieldDialogEls.dlg.close();
      resolve(result);
    }

    fieldDialogEls.type.addEventListener("change", onTypeChange);
    fieldDialogEls.name.addEventListener("input", onNameInput);
    fieldDialogEls.desc.addEventListener("input", onDescInput);
    fieldDialogEls.cancel.addEventListener("click", onCancel);
    fieldDialogEls.form.addEventListener("submit", onSubmit);
    fieldDialogEls.dlg.addEventListener("cancel", onCancel);

    fieldDialogEls.dlg.showModal();
    onNameInput();
    setTimeout(() => fieldDialogEls.name.focus(), 0);
  });
}

// Searchable font picker — replaces <datalist> which has tiny popups
// in Chromium-based webviews.
function createFontPicker(fonts, initialValue, onChange) {
  const wrap = document.createElement("div");
  wrap.className = "font-picker";

  const input = document.createElement("input");
  input.type = "text";
  input.className = "input font-input";
  input.value = initialValue || "";
  input.placeholder = "字体（输入搜索）";
  input.autocomplete = "off";
  input.spellcheck = false;

  const list = document.createElement("ul");
  list.className = "font-picker-list";

  let activeIdx = -1;
  let visible = [];

  function paint(filter = "") {
    const f = filter.toLowerCase();
    visible = f ? fonts.filter((x) => x.toLowerCase().includes(f)) : fonts.slice();
    list.innerHTML = "";
    if (visible.length === 0) {
      const empty = document.createElement("div");
      empty.className = "font-picker-empty";
      empty.textContent = "没有匹配的字体（保留输入即可）";
      list.appendChild(empty);
      return;
    }
    visible.forEach((font, i) => {
      const li = document.createElement("li");
      li.className = "font-picker-item";
      li.textContent = font;
      try {
        li.style.fontFamily = `"${font.replace(/"/g, "")}", sans-serif`;
      } catch {}
      if (i === activeIdx) li.classList.add("active");
      li.addEventListener("mousedown", (e) => {
        e.preventDefault();
        select(font);
      });
      list.appendChild(li);
    });
  }
  function select(font) {
    input.value = font;
    onChange(font);
    closeList();
  }
  function openList() {
    list.classList.add("open");
    // Always show the full list on open (don't auto-filter by current value —
    // the family name we have stored may not exactly match the system font's
    // family name, in which case filtering would show "no matches"). User can
    // type to filter.
    paint("");
    // Highlight + scroll to the currently-selected font, if any
    const idx = visible.indexOf(input.value);
    if (idx >= 0) {
      activeIdx = idx;
      paint("");
      scrollToActive();
    }
  }
  function closeList() {
    list.classList.remove("open");
    activeIdx = -1;
  }
  function scrollToActive() {
    const el = list.querySelector(".font-picker-item.active");
    if (el) el.scrollIntoView({ block: "nearest" });
  }

  input.addEventListener("click", openList);
  input.addEventListener("blur", () => setTimeout(closeList, 150));
  input.addEventListener("input", () => {
    onChange(input.value);
    activeIdx = -1;
    // Filter only when user actively types something
    paint(input.value);
    if (!list.classList.contains("open")) list.classList.add("open");
  });
  input.addEventListener("keydown", (e) => {
    if (!list.classList.contains("open")) {
      if (e.key === "ArrowDown" || e.key === "ArrowUp" || e.key === "Enter") {
        openList();
      }
    }
    if (e.key === "ArrowDown") {
      e.preventDefault();
      activeIdx = Math.min(activeIdx + 1, visible.length - 1);
      paint(input.value);
      scrollToActive();
    } else if (e.key === "ArrowUp") {
      e.preventDefault();
      activeIdx = Math.max(activeIdx - 1, 0);
      paint(input.value);
      scrollToActive();
    } else if (e.key === "Enter") {
      if (activeIdx >= 0 && visible[activeIdx]) {
        e.preventDefault();
        select(visible[activeIdx]);
      }
    } else if (e.key === "Escape") {
      closeList();
    }
  });

  wrap.append(input, list);
  return { wrap, input };
}

function renderParagraphList() {
  els.paragraphList.innerHTML = "";
  if (state.paragraphs.length === 0) {
    els.paragraphList.innerHTML =
      '<p class="empty-state">文档里没有可编辑的段落。</p>';
    return;
  }

  for (const p of state.paragraphs) {
    const card = document.createElement("div");
    card.className = "paragraph-card";
    card.dataset.pIdx = String(p.index);

    const num = document.createElement("div");
    num.className = "paragraph-num";
    num.textContent = `${p.index + 1}.`;
    card.appendChild(num);

    const body = document.createElement("div");
    body.className = "paragraph-body";

    const ed = document.createElement("div");
    ed.className = "paragraph-editor";
    ed.contentEditable = "true";
    ed.spellcheck = false;
    ed.dataset.pIdx = String(p.index);
    ed.innerHTML = textToEditorHtml(p.currentText, p.index);
    let savedSel = { start: 0, end: 0 };

    function saveEditorSelection() {
      const sel = window.getSelection();
      if (sel && sel.rangeCount > 0 && ed.contains(sel.getRangeAt(0).startContainer)) {
        savedSel = getEditorSelectionOffsets(ed);
      }
    }

    ed.addEventListener("input", () => {
      p.currentText = editorTextContent(ed);
      p.dirty = p.currentText !== p.originalText;
      card.classList.toggle("dirty", p.dirty);
      card.classList.toggle(
        "has-placeholder",
        getPlaceholderRanges(p.currentText, p.index).some((x) => x.managed),
      );
      const nextHtml = textToEditorHtml(p.currentText, p.index);
      if (ed.innerHTML !== nextHtml) {
        const { start } = getEditorSelectionOffsets(ed);
        ed.innerHTML = nextHtml;
        placeCaretAtTextOffset(ed, start);
      }
      autoResizeEditor(ed);
      syncFieldMetaFromText();
      syncOccurrenceStyles();
      updateFieldSummary();
      schedulePreviewRefresh();
    });

    ed.addEventListener("focus", () => highlightPreview(p.index));
    ed.addEventListener("blur", () => {
      saveEditorSelection();
      clearActivePreview();
    });
    ed.addEventListener("keyup", saveEditorSelection);
    ed.addEventListener("mouseup", saveEditorSelection);

    ed.addEventListener("mouseover", (e) => {
      const token = e.target.closest?.("[data-token='1']");
      if (token && ed.contains(token)) {
        showTokenTooltip(token, p.index);
      }
    });
    ed.addEventListener("mouseout", (e) => {
      const token = e.target.closest?.("[data-token='1']");
      if (token && ed.contains(token)) {
        removeTokenTooltip();
      }
    });

    ed.addEventListener("keydown", (e) => {
      if (e.key === "Enter") {
        e.preventDefault();
        document.execCommand("insertLineBreak");
        return;
      }
      if (e.key !== "Backspace" && e.key !== "Delete") return;
      const { start, end } = getEditorSelectionOffsets(ed);
      const res = removeManagedTokenAtSelection(p.currentText, p.index, start, end, e.key);
      if (!res) return;
      e.preventDefault();
      p.currentText = res.text;
      ed.innerHTML = textToEditorHtml(res.text, p.index);
      placeCaretAtTextOffset(ed, res.caret);
      ed.dispatchEvent(new Event("input", { bubbles: true }));
    });

    ed.addEventListener("contextmenu", (e) => {
      const token = e.target.closest?.("[data-token='1']");
      if (!token || !ed.contains(token)) return;
      e.preventDefault();
      const occurrenceIdx = getTokenOccurrenceIndex(ed, token);
      openTokenContextMenu(e.clientX, e.clientY, () => {
        const ranges = getPlaceholderRanges(p.currentText, p.index).filter((r) => r.managed);
        if (occurrenceIdx < 0 || occurrenceIdx >= ranges.length) return;
        const target = ranges[occurrenceIdx];
        p.currentText = p.currentText.slice(0, target.start) + p.currentText.slice(target.end);
        syncOccurrenceStyles();
        ed.innerHTML = textToEditorHtml(p.currentText, p.index);
        placeCaretAtTextOffset(ed, target.start);
        ed.dispatchEvent(new Event("input", { bubbles: true }));
      });
    });

    ed.addEventListener("dblclick", async (e) => {
      const token = e.target.closest?.("[data-token='1']");
      if (!token || !ed.contains(token)) return;
      const sigil = token.dataset.sigil;
      const name = token.dataset.name;
      const occurrenceIdx = getTokenOccurrenceIndex(ed, token);
      const styles = state.occurrenceStyles.get(p.index) || [];
      const curStyle = styles[occurrenceIdx] || {};
      const detected = getRunStyleAt(p.originalXml, savedSel.start);
      const existingMeta = state.fieldMeta.get(name);
      const result = await openFieldDialog({
        title: `编辑占位符 {${sigil}${name}}`,
        type: sigil === "%" ? "image" : "text",
        name,
        description: (existingMeta?.description) || "",
        defaultFont: curStyle.font || detected?.font || null,
        defaultSize: curStyle.size ?? detected?.size ?? null,
        defaultSizeLabel: curStyle.sizeLabel || (curStyle.size != null ? findSizeByValue(curStyle.size)?.label : null) || null,
        defaultColor: curStyle.color || detected?.color || null,
        imageConfig: existingMeta?.imageConfig || null,
        detectedFromXml: !curStyle.font && !!detected,
        lockType: true,
        lockName: false,
        existingName: name,
        hideDescription: false,
      });
      if (!result) return;
      // If name changed, do a full rename across all paragraphs
      if (result.name !== name) {
        renameField(name, result.name);
        syncOccurrenceStyles();
        ed.innerHTML = textToEditorHtml(p.currentText, p.index);
      }
      // Update imageConfig for image fields
      if (result.imageConfig) {
        const meta = state.fieldMeta.get(result.name);
        if (meta) {
          meta.imageConfig = result.imageConfig;
          state.fieldMeta.set(result.name, meta);
        }
      }
      if (occurrenceIdx >= 0 && occurrenceIdx < styles.length) {
        styles[occurrenceIdx] = {
          ...styles[occurrenceIdx],
          name: result.name,
          font: result.defaultFont || null,
          size: result.defaultSize ?? null,
          sizeLabel: result.defaultSizeLabel || null,
          color: result.defaultColor || null,
        };
        state.occurrenceStyles.set(p.index, styles);
      }
      updateFieldSummary();
      schedulePreviewRefresh();
    });

    body.appendChild(ed);
    requestAnimationFrame(() => autoResizeEditor(ed));

    const actions = document.createElement("div");
    actions.className = "paragraph-actions";

    const btnText = document.createElement("button");
    btnText.type = "button";
    btnText.className = "btn-mini";
    btnText.textContent = "+ 文字字段";
    btnText.addEventListener("click", async () => {
      const sel = savedSel;
      const detected = getRunStyleAt(p.originalXml, sel.start);
      const result = await openFieldDialog({
        type: "text",
        defaultFont: detected?.font,
        defaultSize: detected?.size,
        defaultColor: detected?.color,
        detectedFromXml: !!detected,
      });
      if (!result) return;
      state.fieldMeta.set(result.name, {
        type: result.type,
        description: result.description,
      });
      ed.focus();
      insertAtCursor(
        ed,
        result.type === "image" ? `{%${result.name}}` : `{@${result.name}}`,
        p.index,
        sel.start,
        sel.end,
      );
      if (result.type === "text") {
        const ranges = getPlaceholderRanges(p.currentText, p.index).filter((r) => r.managed);
        const newToken = `{@${result.name}}`;
        const idx = ranges.findIndex((r, i) => r.text === newToken && !(state.occurrenceStyles.get(p.index)?.[i]?.font));
        if (idx >= 0) {
          const styles = state.occurrenceStyles.get(p.index) || [];
          styles[idx] = {
            name: result.name,
            sigil: "@",
            font: result.defaultFont || null,
            size: result.defaultSize ?? null,
            sizeLabel: result.defaultSizeLabel || null,
            color: result.defaultColor || null,
            description: result.description || null,
          };
          state.occurrenceStyles.set(p.index, styles);
        }
      }
      updateFieldSummary();
    });

    const btnImg = document.createElement("button");
    btnImg.type = "button";
    btnImg.className = "btn-mini image";
    btnImg.textContent = "+ 图片字段";
    btnImg.addEventListener("click", async () => {
      const sel = savedSel;
      const result = await openFieldDialog({ type: "image" });
      if (!result) return;
      state.fieldMeta.set(result.name, {
        type: result.type,
        description: result.description,
        ...(result.imageConfig ? { imageConfig: result.imageConfig } : {}),
      });
      ed.focus();
      insertAtCursor(
        ed,
        result.type === "image" ? `{%${result.name}}` : `{@${result.name}}`,
        p.index,
        sel.start,
        sel.end,
      );
      if (result.type === "image") {
        const ranges = getPlaceholderRanges(p.currentText, p.index).filter((r) => r.managed);
        const newToken = `{%${result.name}}`;
        const idx = ranges.findIndex((r, i) => r.text === newToken && !(state.occurrenceStyles.get(p.index)?.[i]?.imageConfig));
        if (idx >= 0) {
          const styles = state.occurrenceStyles.get(p.index) || [];
          styles[idx] = {
            name: result.name,
            sigil: "%",
            font: null,
            size: null,
            sizeLabel: null,
            color: null,
            description: result.description || null,
          };
          state.occurrenceStyles.set(p.index, styles);
        }
      }
      updateFieldSummary();
    });

    const btnReset = document.createElement("button");
    btnReset.type = "button";
    btnReset.className = "btn-mini reset";
    btnReset.textContent = "撤销修改";
    btnReset.addEventListener("click", () => {
      p.currentText = p.originalText;
      p.dirty = false;
      ed.innerHTML = textToEditorHtml(p.originalText, p.index);
      card.classList.remove("dirty");
      card.classList.toggle(
        "has-placeholder",
        getPlaceholderRanges(p.currentText, p.index).some((x) => x.managed),
      );
      state.occurrenceStyles.delete(p.index);
      syncOccurrenceStyles();
      autoResizeEditor(ed);
      updateFieldSummary();
      schedulePreviewRefresh();
    });

    actions.append(btnText, btnImg, btnReset);
    if (p.hasComplex) {
      const warn = document.createElement("span");
      warn.style.cssText =
        "font-size:11px;color:var(--warning);margin-left:auto;";
      warn.title = "此段含图片/超链接等复杂内容，编辑后可能丢失";
      warn.textContent = "⚠ 复杂段落";
      actions.appendChild(warn);
    }
    body.appendChild(actions);

    card.appendChild(body);
    if (getPlaceholderRanges(p.currentText, p.index).some((r) => r.managed)) {
      card.classList.add("has-placeholder");
    }
    els.paragraphList.appendChild(card);
  }
}

async function editLoad() {
  const path = await pickDocx("选择文档或模板");
  if (!path) return;
  try {
    const bytes = await readBytesFromPath(path);
    const zip = new PizZip(bytes);
    state.templateBytes = bytes;
    state.filename = basename(path);
    state.paragraphs = parseParagraphs(zip);
    buildOriginalPlaceholderBaseline();
    const { fieldMeta, occStyles } = readFieldMeta(zip);
    state.fieldMeta = fieldMeta;
    state.occurrenceStyles = occStyles;
    state.isTemplateInput = !!zip.file("template/fields.json");
    syncFieldMetaFromText();
    els.editFilename.textContent = state.filename;
    els.btnEditSave.textContent = state.isTemplateInput ? "保存修改" : "保存为模板";
    setStatus(`已加载，共 ${state.paragraphs.length} 段，正在渲染预览…`);

    renderParagraphList();
    updateFieldSummary();
    clearPreviewSelection();
    els.btnEditSave.disabled = false;
    els.btnRefreshPreview.disabled = false;
    updateBatchExportButton();

    await renderPreview(bytes);
    setStatus(`已加载，共 ${state.paragraphs.length} 段`);
  } catch (e) {
    console.error(e);
    setStatus("加载失败：" + (e?.message || e), "error");
  }
}

async function editSave() {
  if (!state.templateBytes) return;
  setStatus("生成中…");
  try {
    syncFieldMetaFromText();
    const out = buildTemplate(
      state.templateBytes,
      state.paragraphs,
      state.fieldMeta,
      state.occurrenceStyles,
    );
    // If the input is already a template, default to overwriting it
    // (Tauri save dialog still prompts on collision). Otherwise append
    // -template so the original docx isn't replaced.
    const stem = state.filename.replace(/\.docx$/i, "");
    const suggested = state.isTemplateInput ? state.filename : stem + "-template.docx";
    const target = await saveBytesViaDialog(suggested, out);
    if (!target) {
      setStatus("已取消保存");
      return;
    }
    setStatus(state.isTemplateInput ? "已保存修改：" + target : "已保存模板：" + target, "success");
  } catch (e) {
    console.error(e);
    setStatus("保存失败：" + (e?.message || e), "error");
  }
}

async function refreshPreview() {
  if (!state.templateBytes || state.paragraphs.length === 0) return;
  setStatus("刷新预览…");
  try {
    syncFieldMetaFromText();
    const updated = buildTemplate(
      state.templateBytes,
      state.paragraphs,
      state.fieldMeta,
      state.occurrenceStyles,
    );
    await renderPreview(updated);
    const dirty = state.paragraphs.filter((p) => p.dirty).length;
    setStatus(
      dirty > 0 ? `预览已刷新（已修改 ${dirty} 段）` : "预览已刷新",
      "success",
    );
  } catch (e) {
    console.error(e);
    setStatus("刷新失败：" + (e?.message || e), "error");
  }
}

els.btnEditLoad.addEventListener("click", editLoad);
els.btnEditSave.addEventListener("click", editSave);
els.btnRefreshPreview.addEventListener("click", refreshPreview);

els.btnZoomOut?.addEventListener("click", () => {
  const base = previewFitWidth ? computeFitScale() : previewScale;
  setPreviewScale(base - PREVIEW_ZOOM_STEP);
});
els.btnZoomIn?.addEventListener("click", () => {
  const base = previewFitWidth ? computeFitScale() : previewScale;
  setPreviewScale(base + PREVIEW_ZOOM_STEP);
});
els.btnZoomReset?.addEventListener("click", () => {
  setPreviewFitWidth();
});

els.previewContainer.addEventListener(
  "wheel",
  (e) => {
    if (!e.ctrlKey) return;
    e.preventDefault();
    const base = previewFitWidth ? computeFitScale() : previewScale;
    const delta = e.deltaY < 0 ? PREVIEW_ZOOM_STEP : -PREVIEW_ZOOM_STEP;
    setPreviewScale(base + delta);
  },
  { passive: false },
);

window.addEventListener("resize", () => {
  if (previewFitWidth) applyPreviewZoom();
});

// ============================================================
// Fill mode
// ============================================================
async function fillLoad() {
  // Warm the font cache early. Rust enumeration is fast (~50ms) but parses
  // every system font, so kick it off in parallel with the file dialog
  // instead of making the user wait when they first see the form.
  getFonts();

  const path = await pickDocx("选择模板");
  if (!path) return;
  try {
    const bytes = await readBytesFromPath(path);
    const zip = new PizZip(bytes);
    state.templateBytes = bytes;
    state.filename = basename(path);
    state.fields = extractFieldsFromZip(zip);
    const { fieldMeta, occStyles } = readFieldMeta(zip);
    state.fieldMeta = fieldMeta;
    state.occurrenceStyles = occStyles;
    state.values = {};
    fillOccOverrides.clear();
    els.fillFilename.textContent = state.filename;
    setStatus(`已加载模板，发现 ${state.fields.length} 个字段`);
    await renderForm();
    els.btnBatchImport.disabled = false;
  } catch (e) {
    console.error(e);
    setStatus("加载失败：" + (e?.message || e), "error");
  }
}

function labeled(label, child) {
  const wrap = document.createElement("label");
  wrap.className = "labeled";
  const span = document.createElement("span");
  span.textContent = label;
  wrap.append(span, child);
  return wrap;
}

function findSizeByLabel(label) {
  return ALL_SIZE_PRESETS.find((s) => s.label === String(label));
}
function findSizeByValue(value) {
  // Prefer the Chinese name when both exist for the same pt
  return ALL_SIZE_PRESETS.find((s) => Number(s.value) === Number(value));
}

function ensureImageConfig(name) {
  const meta = state.fieldMeta.get(name);
  if (!meta) return null;
  if (!meta.imageConfig) {
    meta.imageConfig = {
      ...DEFAULT_IMAGE_CONFIG,
    };
  }
  return meta.imageConfig;
}

async function renderForm() {
  els.formSection.innerHTML = "";
  if (state.fields.length === 0) {
    els.formSection.innerHTML =
      '<p class="empty-state">没有在模板中找到 {@field} 或 {%field} 占位符。可以切到「制作模板」加占位符。</p>';
    els.btnFillExport.disabled = true;
    return;
  }

  const fonts = await getFonts();

  // Group occurrences by field name for fill mode
  const fieldOccMap = new Map(); // name -> [{font, size, sizeLabel, color, description}, ...]
  for (const [, styles] of state.occurrenceStyles) {
    for (const entry of styles) {
      if (!entry || !entry.name) continue;
      if (!fieldOccMap.has(entry.name)) fieldOccMap.set(entry.name, []);
      fieldOccMap.get(entry.name).push({
        font: entry.font || null,
        size: entry.size ?? null,
        sizeLabel: entry.sizeLabel || null,
        color: entry.color || null,
        description: entry.description || null,
      });
    }
  }

  // Deduplicate field list by name
  const seen = new Set();
  const uniqueFields = [];
  for (const f of state.fields) {
    if (seen.has(f.name)) continue;
    seen.add(f.name);
    uniqueFields.push(f);
  }

  for (const field of uniqueFields) {
    const meta = state.fieldMeta.get(field.name);

    if (field.type === "text") {
      const v = state.values[field.name] || { text: "" };
      state.values[field.name] = v;
      const occs = fieldOccMap.get(field.name) || [];

      // If only one or no occurrences, render a single card
      if (occs.length <= 1) {
        const card = createFillFieldCard(field.name, field.type, meta?.description, null);
        const ta = createFillTextarea(v);
        card.appendChild(ta);
        if (occs.length === 1) {
          card.appendChild(createFillStyleControls(fonts, occs, 0));
        }
        els.formSection.appendChild(card);
      } else {
        // Multiple occurrences: each gets its own card with synced text
        const allTas = [];
        for (let i = 0; i < occs.length; i++) {
          const occDesc = occs[i].description || meta?.description || "";
          const card = createFillFieldCard(field.name, field.type, occDesc, `${i + 1}/${occs.length}`);
          const ta = createFillTextarea(v);
          ta.addEventListener("input", () => {
            // Sync all other textareas
            for (const other of allTas) {
              if (other !== ta) other.value = ta.value;
            }
          });
          allTas.push(ta);
          card.appendChild(ta);
          card.appendChild(createFillStyleControls(fonts, occs, i));
          els.formSection.appendChild(card);
        }
      }
    } else {
      const card = createFillFieldCard(field.name, field.type, meta?.description, null);
      const v = state.values[field.name] || { bytes: null, mime: "", filename: "", previewUrl: "" };
      state.values[field.name] = v;

      const row = document.createElement("div");
      row.className = "image-row";
      const fileInp = document.createElement("input");
      fileInp.type = "file";
      fileInp.accept = "image/png,image/jpeg,image/gif";
      const preview = document.createElement("span");
      preview.className = "image-preview-name";
      preview.textContent = v.filename || "尚未选择图片";
      const imagePreview = document.createElement("div");
      imagePreview.className = "image-preview-box";
      const imagePreviewImg = document.createElement("img");
      imagePreviewImg.className = "image-preview-thumb";
      imagePreviewImg.alt = `${field.name} 预览`;
      const imagePreviewEmpty = document.createElement("div");
      imagePreviewEmpty.className = "image-preview-empty";
      imagePreviewEmpty.textContent = "未选择图片";

      function renderImagePreview() {
        const hasPreview = !!v.previewUrl;
        imagePreview.innerHTML = "";
        if (hasPreview) {
          imagePreviewImg.src = v.previewUrl;
          imagePreview.appendChild(imagePreviewImg);
        } else {
          imagePreview.appendChild(imagePreviewEmpty);
        }
      }

      fileInp.addEventListener("change", async () => {
        const f = fileInp.files?.[0];
        if (!f) return;
        if (v.previewUrl) {
          URL.revokeObjectURL(v.previewUrl);
        }
        v.bytes = await fileToUint8Array(f);
        v.mime = f.type;
        v.filename = f.name;
        v.previewUrl = URL.createObjectURL(f);
        preview.textContent = f.name;
        renderImagePreview();
      });
      row.append(fileInp, preview);
      card.appendChild(row);
      renderImagePreview();
      card.appendChild(imagePreview);

      // Image config controls
      const imgCfg = meta?.imageConfig || {};
      const cfgSection = document.createElement("div");
      cfgSection.className = "image-config-section";
      const cfgTitle = document.createElement("div");
      cfgTitle.className = "image-config-title";
      cfgTitle.textContent = "尺寸配置";
      cfgSection.appendChild(cfgTitle);

      const cfgRow1 = document.createElement("div");
      cfgRow1.className = "image-config-row";
      const fitSel = labeled("自适应模式", (() => {
        const sel = document.createElement("select");
        sel.className = "input";
        sel.innerHTML = '<option value="width">宽度自适应</option><option value="height">高度自适应</option><option value="contain">等比缩放（约束内）</option>';
        sel.value = imgCfg.fitMode || DEFAULT_IMAGE_CONFIG.fitMode;
        const syncFitMode = () => {
          ensureImageConfig(field.name).fitMode = sel.value;
        };
        sel.addEventListener("change", syncFitMode);
        sel.addEventListener("input", syncFitMode);
        return sel;
      })());
      const ratioLabel = document.createElement("label");
      ratioLabel.className = "checkbox-label";
      const ratioCb = document.createElement("input");
      ratioCb.type = "checkbox";
      ratioCb.checked = imgCfg.maintainRatio !== false;
      const syncMaintainRatio = () => {
        ensureImageConfig(field.name).maintainRatio = ratioCb.checked;
      };
      ratioCb.addEventListener("change", syncMaintainRatio);
      ratioCb.addEventListener("input", syncMaintainRatio);
      ratioLabel.append(ratioCb, document.createTextNode(" 保持宽高比"));
      const ratioWrap = labeled("\xa0", ratioLabel);
      ratioWrap.style.flex = "0 0 auto";
      cfgRow1.append(fitSel, ratioWrap);

      const cfgRow2 = document.createElement("div");
      cfgRow2.className = "image-config-row image-config-dims";
      const dims = [
        ["最大宽度", "maxWidth", imgCfg.maxWidth ?? DEFAULT_IMAGE_CONFIG.maxWidth],
        ["最大高度", "maxHeight", imgCfg.maxHeight ?? DEFAULT_IMAGE_CONFIG.maxHeight],
        ["最小宽度", "minWidth", imgCfg.minWidth ?? DEFAULT_IMAGE_CONFIG.minWidth],
        ["最小高度", "minHeight", imgCfg.minHeight ?? DEFAULT_IMAGE_CONFIG.minHeight],
      ];
      for (const [lbl, key, val] of dims) {
        const inp = document.createElement("input");
        inp.type = "number";
        inp.className = "input";
        inp.value = formatCm(val, DEFAULT_IMAGE_CONFIG[key]);
        inp.min = 0.01;
        inp.max = 999.99;
        inp.step = 0.01;
        const syncDim = () => {
          const next = roundCm(Number(inp.value) || val);
          ensureImageConfig(field.name)[key] = next;
          inp.value = formatCm(next, val);
        };
        inp.addEventListener("change", syncDim);
        inp.addEventListener("input", syncDim);
        cfgRow2.appendChild(labeled(`${lbl} (cm)`, inp));
      }

      cfgSection.append(cfgRow1, cfgRow2);
      card.appendChild(cfgSection);
      els.formSection.appendChild(card);
    }
  }
  els.btnFillExport.disabled = false;
}

function createFillFieldCard(name, type, description, occLabel) {
  const card = document.createElement("div");
  card.className = "field-card";
  const head = document.createElement("div");
  head.className = "field-head";
  head.innerHTML = `<span class="field-name">${name}</span><span class="field-type ${type}">${type}</span>` +
    (occLabel ? `<span class="field-occ-label">样式 ${occLabel}</span>` : "");
  card.appendChild(head);
  if (description) {
    const desc = document.createElement("div");
    desc.className = "field-description";
    desc.textContent = description;
    card.appendChild(desc);
  }
  return card;
}

function createFillTextarea(v) {
  const ta = document.createElement("textarea");
  ta.className = "input text-input";
  ta.placeholder = "在此输入文字（支持换行）";
  ta.rows = 2;
  ta.value = v.text;
  ta.addEventListener("input", () => (v.text = ta.value));
  return ta;
}

function createFillStyleControls(fonts, occs, idx) {
  const wrap = document.createElement("div");
  wrap.className = "controls";
  const o = occs[idx];

  const fontPicker = createFontPicker(fonts, o.font || "宋体", (val) => {
    o.font = val.trim() || "Calibri";
  });

  const sizeSel = document.createElement("select");
  sizeSel.className = "input size-input";
  const zhGroup = document.createElement("optgroup");
  zhGroup.label = "字号";
  for (const s of SIZE_PRESETS_ZH) {
    const opt = document.createElement("option");
    opt.value = s.label;
    opt.textContent = s.label;
    opt.dataset.pt = String(s.value);
    zhGroup.appendChild(opt);
  }
  const numGroup = document.createElement("optgroup");
  numGroup.label = "磅值";
  for (const s of SIZE_PRESETS_NUM) {
    const opt = document.createElement("option");
    opt.value = s.label;
    opt.textContent = s.label;
    opt.dataset.pt = String(s.value);
    numGroup.appendChild(opt);
  }
  sizeSel.append(zhGroup, numGroup);
  const sl = o.sizeLabel || (o.size != null ? findSizeByValue(o.size)?.label : null) || "小四";
  sizeSel.value = sl;
  if (sizeSel.selectedIndex < 0) sizeSel.value = "小四";
  sizeSel.addEventListener("change", () => {
    const opt = sizeSel.options[sizeSel.selectedIndex];
    o.sizeLabel = opt.value;
    o.size = Number(opt.dataset.pt);
  });

  const colorInp = document.createElement("input");
  colorInp.type = "color";
  colorInp.className = "color-input";
  colorInp.value = o.color || "#000000";
  const colorSwatch = document.createElement("div");
  colorSwatch.className = "color-swatch";
  colorSwatch.style.background = o.color || "#000000";
  const colorHex = document.createElement("span");
  colorHex.className = "color-hex";
  colorHex.textContent = (o.color || "#000000").toUpperCase();
  const colorWrap = document.createElement("div");
  colorWrap.className = "color-input-wrap";
  colorWrap.append(colorSwatch, colorInp, colorHex);
  colorInp.addEventListener("input", () => {
    o.color = colorInp.value;
    colorSwatch.style.background = colorInp.value;
    colorHex.textContent = colorInp.value.toUpperCase();
  });

  wrap.append(
    labeled("字体", fontPicker.wrap),
    labeled("字号", sizeSel),
    labeled("颜色", colorWrap),
  );
  return wrap;
}

async function fillExport() {
  if (!state.templateBytes) return;
  setStatus("生成中…");

  // Flush any focused image config inputs before reading state.
  // Some browser/Tauri input controls can still hold the latest typed value
  // until blur/change has propagated.
  if (document.activeElement && typeof document.activeElement.blur === "function") {
    document.activeElement.blur();
  }

  syncImageConfigsFromForm();

  // Sync per-occurrence descriptions from fill form back to occurrenceStyles
  for (const [key, override] of fillOccOverrides) {
    const [name, occIdxStr] = key.split(":");
    const occIdx = Number(occIdxStr);
    for (const [, styles] of state.occurrenceStyles) {
      let count = 0;
      for (const s of styles) {
        if (s && s.name === name) {
          if (count === occIdx) {
            s.description = override.description;
            break;
          }
          count++;
        }
      }
    }
  }

  try {
    const values = {};
    const occStyleMap = new Map(); // name -> [{font, size, color}, ...]
    const imageConfigMap = new Map(); // name -> imageConfig

    // Build occStyleMap from occurrenceStyles (entries include name)
    for (const [, styles] of state.occurrenceStyles) {
      for (const entry of styles) {
        if (!entry || !entry.name) continue;
        if (!occStyleMap.has(entry.name)) occStyleMap.set(entry.name, []);
        occStyleMap.get(entry.name).push({
          font: entry.font || "宋体",
          size: entry.size ?? 12,
          color: entry.color || "#000000",
        });
      }
    }

    // Build imageConfigMap from fieldMeta
    for (const [name, meta] of state.fieldMeta) {
      if (meta.type === "image" && meta.imageConfig) {
        imageConfigMap.set(name, meta.imageConfig);
      }
    }

    for (const f of state.fields) {
      const v = state.values[f.name];
      if (f.type === "text") {
        values[f.name] = {
          text: v?.text ?? "",
        };
      } else {
        values[f.name] = { bytes: v?.bytes };
      }
    }
    const out = renderFilled(state.templateBytes, state.fields, values, occStyleMap, imageConfigMap);
    const suggested = state.filename.replace(/\.docx$/i, "") + "-filled.docx";
    const target = await saveBytesViaDialog(suggested, out);
    if (!target) {
      setStatus("已取消保存");
      return;
    }
    setStatus("已保存：" + target, "success");
  } catch (e) {
    console.error(e);
    const msg = e?.properties?.errors
      ? e.properties.errors
          .map((x) => x.properties?.explanation || x.message)
          .join("; ")
      : e?.message || String(e);
    setStatus("生成失败：" + msg, "error");
  }
}

els.btnFillLoad.addEventListener("click", fillLoad);
els.btnFillExport.addEventListener("click", fillExport);

// ============================================================
// Batch: export template + Excel (edit mode)
// ============================================================
// Per-occurrence data edited in fill form (keyed by "name:occIdx")
const fillOccOverrides = new Map();

function showConfirm(title, message) {
  const dlg = document.getElementById("confirm-dialog");
  const titleEl = document.getElementById("confirm-dialog-title");
  const msgEl = document.getElementById("confirm-dialog-message");
  const cancelBtn = document.getElementById("confirm-dialog-cancel");
  titleEl.textContent = title;
  msgEl.textContent = message;
  return new Promise((resolve) => {
    function cleanup() {
      cancelBtn.removeEventListener("click", onCancel);
      dlg.removeEventListener("cancel", onCancel);
      dlg.querySelector("form").removeEventListener("submit", onSubmit);
    }
    function onCancel(e) {
      e?.preventDefault?.();
      cleanup();
      dlg.close();
      resolve(false);
    }
    function onSubmit(e) {
      e.preventDefault();
      cleanup();
      dlg.close();
      resolve(true);
    }
    cancelBtn.addEventListener("click", onCancel);
    dlg.addEventListener("cancel", onCancel);
    dlg.querySelector("form").addEventListener("submit", onSubmit);
    dlg.showModal();
  });
}

function buildSampleDataForPreview() {
  const data = [{}];
  for (const f of state.fields) {
    if (f.type === "text") data[0][f.name] = "";
  }
  return data;
}

function formatFilename(pattern, row, n) {
  let out = pattern.replace(/\{n\}/g, String(n));
  out = out.replace(/\{(\w+)\}/g, (full, name) => {
    const val = row[name];
    if (val == null || val === "") return "";
    return String(val).replace(/[\\/:*?"<>|]/g, "_").substring(0, 50);
  });
  out = out.replace(/[_\s]+/g, "_").replace(/^_|_$/g, "");
  return out || `file_${n}`;
}

function renderFilenamePreview(sampleData) {
  const el = document.getElementById("filename-preview");
  const pattern = document.getElementById("filename-pattern").value || "{n}";
  const lines = [];
  for (let i = 0; i < Math.min(3, sampleData.length); i++) {
    lines.push(formatFilename(pattern, sampleData[i], i + 1) + ".docx");
  }
  if (sampleData.length > 3) lines.push(`… (共 ${sampleData.length} 个文件)`);
  el.innerHTML = lines.map((l) => `<div class="preview-item">${escapeHtml(l)}</div>`).join("");
}

function insertAtCursorSimple(input, text) {
  const start = input.selectionStart;
  const end = input.selectionEnd;
  const old = input.value;
  input.value = old.slice(0, start) + text + old.slice(end);
  input.selectionStart = input.selectionEnd = start + text.length;
  input.focus();
}

function openFilenameDialog(sampleData) {
  const dlg = document.getElementById("filename-dialog");
  const patternInput = document.getElementById("filename-pattern");
  const tagsEl = document.getElementById("filename-field-tags");
  const cancelBtn = document.getElementById("filename-dialog-cancel");
  const hintEl = document.getElementById("filename-dialog-hint");

  // Clear error hint on input
  const clearHint = () => { if (hintEl) hintEl.textContent = ""; };
  patternInput.addEventListener("input", clearHint);

  tagsEl.innerHTML = "";
  const nBtn = document.createElement("button");
  nBtn.type = "button";
  nBtn.className = "filename-field-tag";
  nBtn.textContent = "{n} 自动编号";
  nBtn.addEventListener("click", () => {
    insertAtCursorSimple(patternInput, "{n}");
    renderFilenamePreview(sampleData);
  });
  tagsEl.appendChild(nBtn);
  for (const f of state.fields) {
    if (f.type === "image") continue;
    const b = document.createElement("button");
    b.type = "button";
    b.className = "filename-field-tag";
    b.textContent = `{${f.name}}`;
    b.addEventListener("click", () => {
      insertAtCursorSimple(patternInput, `{${f.name}}`);
      renderFilenamePreview(sampleData);
    });
    tagsEl.appendChild(b);
  }

  patternInput.value = state.filename.replace(/\.docx$/i, "") + "_{n}";
  renderFilenamePreview(sampleData);

  const onInput = () => renderFilenamePreview(sampleData);
  patternInput.addEventListener("input", onInput);

  return new Promise((resolve) => {
    function cleanup() {
      patternInput.removeEventListener("input", onInput);
      patternInput.removeEventListener("input", clearHint);
      cancelBtn.removeEventListener("click", onCancel);
      dlg.removeEventListener("cancel", onCancel);
      dlg.querySelector("form").removeEventListener("submit", onSubmit);
    }
    function onCancel(e) {
      e?.preventDefault?.();
      cleanup();
      dlg.close();
      resolve(null);
    }
    function onSubmit(e) {
      e.preventDefault();
      const val = patternInput.value.trim();
      // Validate: check for Windows-forbidden characters
      const forbidden = val.match(/[\\/:*?"<>|]/g);
      if (forbidden) {
        const chars = [...new Set(forbidden)].map((c) => `"${c}"`).join(" ");
        if (hintEl) hintEl.textContent = `文件名不能包含 ${chars}，请修改`;
        patternInput.focus();
        return;
      }
      if (!val) {
        if (hintEl) hintEl.textContent = "文件名模板不能为空";
        patternInput.focus();
        return;
      }
      cleanup();
      dlg.close();
      resolve(val);
    }
    cancelBtn.addEventListener("click", onCancel);
    dlg.addEventListener("cancel", onCancel);
    dlg.querySelector("form").addEventListener("submit", onSubmit);
    dlg.showModal();
    setTimeout(() => patternInput.focus(), 0);
  });
}

function getOccStyleMap() {
  const occStyleMap = new Map();
  for (const [, styles] of state.occurrenceStyles) {
    for (const entry of styles) {
      if (!entry || !entry.name) continue;
      if (!occStyleMap.has(entry.name)) occStyleMap.set(entry.name, []);
      occStyleMap.get(entry.name).push({
        font: entry.font || "宋体",
        size: entry.size ?? 12,
        color: entry.color || "#000000",
      });
    }
  }
  return occStyleMap;
}

function getImageConfigMap() {
  const imageConfigMap = new Map();
  for (const [name, meta] of state.fieldMeta) {
    if (meta.type === "image" && meta.imageConfig) {
      imageConfigMap.set(name, meta.imageConfig);
    }
  }
  return imageConfigMap;
}

function syncImageConfigsFromForm() {
  const cards = els.formSection.querySelectorAll(".field-card");
  for (const card of cards) {
    const nameEl = card.querySelector(".field-name");
    const typeEl = card.querySelector(".field-type");
    if (!nameEl || !typeEl) continue;
    if ((typeEl.textContent || "").trim() !== "image") continue;
    const name = (nameEl.textContent || "").trim();
    if (!name) continue;
    const meta = state.fieldMeta.get(name);
    if (!meta) continue;
    const fitSelect = card.querySelector("select.input");
    const ratioCheckbox = card.querySelector("input[type='checkbox']");
    const numberInputs = [...card.querySelectorAll("input.input[type='number']")];
    if (numberInputs.length < 4) continue;
    meta.imageConfig = {
      fitMode: fitSelect?.value || meta.imageConfig?.fitMode || DEFAULT_IMAGE_CONFIG.fitMode,
      maintainRatio: ratioCheckbox?.checked ?? (meta.imageConfig?.maintainRatio !== false),
      maxWidth: roundCm(Number(numberInputs[0].value) || meta.imageConfig?.maxWidth || DEFAULT_IMAGE_CONFIG.maxWidth),
      maxHeight: roundCm(Number(numberInputs[1].value) || meta.imageConfig?.maxHeight || DEFAULT_IMAGE_CONFIG.maxHeight),
      minWidth: roundCm(Number(numberInputs[2].value) || meta.imageConfig?.minWidth || DEFAULT_IMAGE_CONFIG.minWidth),
      minHeight: roundCm(Number(numberInputs[3].value) || meta.imageConfig?.minHeight || DEFAULT_IMAGE_CONFIG.minHeight),
    };
    state.fieldMeta.set(name, meta);
  }
}

async function batchExportTemplate() {
  if (!state.templateBytes) return;

  // Sync fieldMeta from paragraph text first
  syncFieldMetaFromText();

  // Check for image fields and warn
  const imageFieldNames = [];
  for (const [name, meta] of state.fieldMeta) {
    if (meta.type === "image") imageFieldNames.push(name);
  }
  if (imageFieldNames.length > 0) {
    const ok = await showConfirm(
      "图片字段提醒",
      `模板中包含图片字段（${imageFieldNames.join("、")}），Excel 无法批量填写图片。\n图片字段需要在「填写模板」中手动设置。\n\n是否继续导出？`
    );
    if (!ok) return;
  }

  // Select output folder
  const folder = await open({ directory: true, title: "选择导出文件夹" });
  if (!folder) {
    setStatus("已取消");
    return;
  }
  const folderPath = typeof folder === "string" ? folder : folder.path;

  setStatus("导出中…");
  try {
    // 1. Save the Word template (same as editSave but to chosen folder)
    syncFieldMetaFromText();
    const templateBytes = buildTemplate(
      state.templateBytes,
      state.paragraphs,
      state.fieldMeta,
      state.occurrenceStyles,
    );
    const templateName = state.isTemplateInput
      ? state.filename
      : state.filename.replace(/\.docx$/i, "") + "-template.docx";
    await invoke("save_bytes", {
      path: `${folderPath}/${templateName}`,
      bytes: Array.from(templateBytes),
    });

    // 2. Build Excel template
    const textFields = [];
    const imageFields = [];
    for (const [name, meta] of state.fieldMeta) {
      if (meta.type === "image") imageFields.push({ name, ...meta });
      else textFields.push({ name, ...meta });
    }
    const headers = textFields.map((f) => f.name);
    const rows = [headers];
    // 10 empty rows for the user to fill
    for (let i = 0; i < 10; i++) {
      rows.push(headers.map(() => ""));
    }
    const ws = XLSX.utils.aoa_to_sheet(rows);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "填写数据");
    if (imageFields.length > 0) {
      const noteRows = [
        ["以下图片字段需要单独收集，Excel 无法嵌入图片："],
        ...imageFields.map((f) => [`  - ${f.name}: ${f.description || ""}`]),
      ];
      XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(noteRows), "图片字段说明");
    }
    const excelBuf = XLSX.write(wb, { bookType: "xlsx", type: "array" });
    const excelName = state.filename.replace(/\.docx$/i, "") + "_批量填写表.xlsx";
    await invoke("save_bytes", {
      path: `${folderPath}/${excelName}`,
      bytes: Array.from(new Uint8Array(excelBuf)),
    });

    setStatus(`已导出：${templateName} + ${excelName} → ${folderPath}`, "success");
  } catch (e) {
    console.error(e);
    setStatus("导出失败：" + (e?.message || e), "error");
  }
}

// Enable batch export button when template is loaded in edit mode
function updateBatchExportButton() {
  if (els.btnBatchExportTemplate) {
    els.btnBatchExportTemplate.disabled = !state.templateBytes || state.paragraphs.length === 0;
  }
}

// ============================================================
// Batch: import Excel + generate Word files (fill mode)
// ============================================================
async function batchImport() {
  if (!state.templateBytes) return;

  // 1. Pick Excel file
  const excelPath = await open({
    multiple: false,
    title: "选择填写好的 Excel 文件",
    filters: [{ name: "Excel", extensions: ["xlsx", "xls"] }],
  });
  if (!excelPath) return;
  const excelFilePath = typeof excelPath === "string" ? excelPath : excelPath.path;

  // 2. Read Excel
  let rows;
  try {
    const excelBytes = await readBytesFromPath(excelFilePath);
    const wb = XLSX.read(excelBytes, { type: "array" });
    const ws = wb.Sheets[wb.SheetNames[0]];
    const data = XLSX.utils.sheet_to_json(ws, { header: 1 });
    if (data.length < 2) throw new Error("Excel 至少需要 2 行（表头 + 1 行数据）");
    const headers = data[0].map((h) => String(h || "").trim());
    rows = [];
    for (let i = 1; i < data.length; i++) {
      const row = {};
      let hasData = false;
      for (let j = 0; j < headers.length; j++) {
        const val = data[i]?.[j];
        row[headers[j]] = val != null ? String(val) : "";
        if (val != null && String(val).trim() !== "") hasData = true;
      }
      if (hasData) rows.push(row);
    }
    if (rows.length === 0) throw new Error("Excel 中没有有效数据行");
  } catch (e) {
    setStatus("读取 Excel 失败：" + (e?.message || e), "error");
    return;
  }

  // 3. Open filename pattern dialog
  const pattern = await openFilenameDialog(rows);
  if (!pattern) return;

  // 4. Select output folder
  const folder = await open({ directory: true, title: "选择输出文件夹" });
  if (!folder) {
    setStatus("已取消");
    return;
  }
  const folderPath = typeof folder === "string" ? folder : folder.path;

  // 5. Generate Word files
  setStatus(`批量生成中（共 ${rows.length} 个）…`);
  let success = 0;
  let fail = 0;
  const occStyleMap = getOccStyleMap();
  const imageConfigMap = getImageConfigMap();

  for (let i = 0; i < rows.length; i++) {
    try {
      const values = {};
      for (const f of state.fields) {
        if (f.type === "text") {
          values[f.name] = { text: rows[i][f.name] || "" };
        } else {
          values[f.name] = { bytes: state.values[f.name]?.bytes };
        }
      }
      const out = renderFilled(state.templateBytes, state.fields, values, occStyleMap, imageConfigMap);
      const fileName = formatFilename(pattern, rows[i], i + 1) + ".docx";
      await invoke("save_bytes", { path: `${folderPath}/${fileName}`, bytes: Array.from(out) });
      success++;
    } catch (e) {
      console.error(`Row ${i + 1} failed:`, e);
      fail++;
    }
  }

  setStatus(`批量生成完成：成功 ${success} 个${fail > 0 ? `，失败 ${fail} 个` : ""}。输出：${folderPath}`, success > 0 ? "success" : "error");
}

els.btnBatchExportTemplate.addEventListener("click", batchExportTemplate);
els.btnBatchImport.addEventListener("click", batchImport);

// Initial state
switchMode("edit");
updateFieldSummary();
updateSelectionUI();

// GitHub link
document.getElementById("github-link")?.addEventListener("click", (e) => {
  e.preventDefault();
  shellOpen("https://github.com/Creeeeeeeeeeper/doc_template");
});
