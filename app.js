import JSZip from "https://esm.sh/jszip@3.10.1";
import { PDFDocument } from "https://esm.sh/pdf-lib@1.17.1";
import * as pdfjsLib from "https://cdn.jsdelivr.net/npm/pdfjs-dist@4.5.136/build/pdf.min.mjs";

pdfjsLib.GlobalWorkerOptions.workerSrc =
  "https://cdn.jsdelivr.net/npm/pdfjs-dist@4.5.136/build/pdf.worker.min.mjs";

const REF_PPTX_WIDTH = 1376;
const REF_PPTX_HEIGHT = 768;
const DEFAULT_PDF_RECT_PTS = [1268, 738, 113, 40];
const DEFAULT_PPTX_RECT_REF = [1265, 740, 103, 18];

const els = {
  fileInput: document.getElementById("fileInput"),
  processBtn: document.getElementById("processBtn"),
  status: document.getElementById("status"),
  log: document.getElementById("log"),
  removeDefault: document.getElementById("removeDefault"),
  browserFileSelect: document.getElementById("browserFileSelect"),
  thumbs: document.getElementById("thumbs"),
  previewCanvas: document.getElementById("previewCanvas"),
  previewWrap: document.getElementById("previewWrap"),
  areasSummary: document.getElementById("areasSummary"),
  pageIndicator: document.getElementById("pageIndicator"),
  prevBtn: document.getElementById("prevBtn"),
  nextBtn: document.getElementById("nextBtn"),
  removeLastBtn: document.getElementById("removeLastBtn"),
  clearPageBtn: document.getElementById("clearPageBtn")
};

const state = {
  contexts: [],
  customAreas: {},
  currentFileIdx: -1,
  currentPageIdx: -1,
  render: {
    displayScale: 1,
    coordScaleX: 1,
    coordScaleY: 1,
    pageWidth: 1,
    pageHeight: 1
  },
  draw: {
    active: false,
    startX: 0,
    startY: 0,
    tempX: 0,
    tempY: 0
  }
};

function log(message) {
  els.log.textContent += `${message}\n`;
  els.log.scrollTop = els.log.scrollHeight;
}

function setStatus(message) {
  if (els.status) {
    els.status.textContent = message;
  }
}

function getConfig() {
  return {
    removeDefault: els.removeDefault.checked,
    pdfRectPts: [...DEFAULT_PDF_RECT_PTS],
    pptxRectRef: [...DEFAULT_PPTX_RECT_REF]
  };
}

function waitForOpenCV() {
  return new Promise((resolve) => {
    if (window.cv && window.cv.Mat) {
      resolve();
      return;
    }

    const timer = setInterval(() => {
      if (window.cv && window.cv.Mat) {
        clearInterval(timer);
        resolve();
      }
    }, 80);
  });
}

function normalizeError(error) {
  if (typeof error === "number" && window.cv?.exceptionFromPtr) {
    try {
      const cvErr = cv.exceptionFromPtr(error);
      return cvErr?.msg || cvErr?.toString?.() || `OpenCV error ptr ${error}`;
    } catch {
      return `OpenCV numeric error: ${error}`;
    }
  }
  if (error?.message) return error.message;
  return String(error);
}

function clampRect([x, y, w, h], width, height) {
  const cx = Math.max(0, Math.min(Math.round(x), Math.max(0, width - 1)));
  const cy = Math.max(0, Math.min(Math.round(y), Math.max(0, height - 1)));
  const cw = Math.max(0, Math.min(Math.round(w), Math.max(0, width - cx)));
  const ch = Math.max(0, Math.min(Math.round(h), Math.max(0, height - cy)));
  return [cx, cy, cw, ch];
}

function drawRectsMask(mask, rects) {
  rects.forEach(([x, y, w, h]) => {
    if (w <= 0 || h <= 0) return;
    const p1 = new cv.Point(x, y);
    const p2 = new cv.Point(x + w, y + h);
    cv.rectangle(mask, p1, p2, new cv.Scalar(255, 255, 255, 255), -1);
  });
}

async function blobToCanvas(blob) {
  const bitmap = await createImageBitmap(blob);
  const canvas = document.createElement("canvas");
  canvas.width = bitmap.width;
  canvas.height = bitmap.height;
  const ctx = canvas.getContext("2d", { willReadFrequently: true });
  ctx.drawImage(bitmap, 0, 0);
  bitmap.close();
  return canvas;
}

async function inpaintCanvas(canvas, rects, radius = 5) {
  if (!rects.length) return canvas;

  const srcRgba = cv.imread(canvas);
  const src = new cv.Mat();
  cv.cvtColor(srcRgba, src, cv.COLOR_RGBA2RGB);
  const mask = new cv.Mat.zeros(src.rows, src.cols, cv.CV_8UC1);

  try {
    drawRectsMask(mask, rects);
    const dst = new cv.Mat();
    cv.inpaint(src, mask, dst, radius, cv.INPAINT_NS);
    cv.imshow(canvas, dst);
    dst.delete();
    return canvas;
  } finally {
    srcRgba.delete();
    src.delete();
    mask.delete();
  }
}

function downloadBytes(bytes, fileName, mimeType) {
  const blob = new Blob([bytes], { type: mimeType });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = fileName;
  document.body.appendChild(a);
  a.click();
  a.remove();
  URL.revokeObjectURL(url);
}

function withCleanSuffix(name, newExt) {
  const dot = name.lastIndexOf(".");
  const stem = dot > 0 ? name.slice(0, dot) : name;
  return `${stem}_clean.${newExt}`;
}

function fileExt(name) {
  return name.split(".").pop()?.toLowerCase() || "";
}

function byNumericTail(a, b, prefix, suffix) {
  const getNum = (value) => {
    const m = value.match(new RegExp(`${prefix}(\\d+)${suffix}$`));
    return m ? Number(m[1]) : Number.MAX_SAFE_INTEGER;
  };
  return getNum(a) - getNum(b);
}

function parseXml(xmlText) {
  return new DOMParser().parseFromString(xmlText, "application/xml");
}

function normalizeZipPath(path) {
  const parts = path.split("/");
  const out = [];
  for (const p of parts) {
    if (!p || p === ".") continue;
    if (p === "..") {
      out.pop();
      continue;
    }
    out.push(p);
  }
  return out.join("/");
}

function resolveTargetPath(basePath, target) {
  const baseParts = basePath.split("/");
  baseParts.pop();
  return normalizeZipPath(`${baseParts.join("/")}/${target}`);
}

function getContextAreas(contextId) {
  if (!state.customAreas[contextId]) {
    state.customAreas[contextId] = {};
  }
  return state.customAreas[contextId];
}

function getPageAreas(contextId, pageNum) {
  const fileAreas = getContextAreas(contextId);
  if (!fileAreas[pageNum]) {
    fileAreas[pageNum] = [];
  }
  return fileAreas[pageNum];
}

function getCurrentContext() {
  if (state.currentFileIdx < 0 || state.currentFileIdx >= state.contexts.length) return null;
  return state.contexts[state.currentFileIdx];
}

function getCurrentPage() {
  const ctx = getCurrentContext();
  if (!ctx) return null;
  if (state.currentPageIdx < 0 || state.currentPageIdx >= ctx.pages.length) return null;
  return ctx.pages[state.currentPageIdx];
}

function countAllAreas() {
  let total = 0;
  for (const areasByPage of Object.values(state.customAreas)) {
    for (const rects of Object.values(areasByPage)) {
      total += rects.length;
    }
  }
  return total;
}

function updateSummary() {
  els.areasSummary.textContent = `Доп. области: ${countAllAreas()}`;
}

function updatePageIndicator() {
  const ctx = getCurrentContext();
  if (!ctx || !ctx.pages.length || state.currentPageIdx < 0) {
    els.pageIndicator.textContent = "— / —";
    return;
  }
  els.pageIndicator.textContent = `${state.currentPageIdx + 1} / ${ctx.pages.length}`;
}

function ensurePageImage(page) {
  if (page.imagePromise) return page.imagePromise;
  page.imagePromise = new Promise((resolve, reject) => {
    const img = new Image();
    img.onload = () => {
      page.image = img;
      resolve(img);
    };
    img.onerror = reject;
    img.src = page.previewUrl;
  });
  return page.imagePromise;
}

function drawCurrentPage() {
  const page = getCurrentPage();
  const ctx = getCurrentContext();
  const canvas = els.previewCanvas;
  const g = canvas.getContext("2d");

  if (!page || !ctx) {
    canvas.width = 1;
    canvas.height = 1;
    g.clearRect(0, 0, 1, 1);
    updatePageIndicator();
    return;
  }

  ensurePageImage(page)
    .then((img) => {
      const maxWidth = Math.max(200, els.previewWrap.clientWidth - 20);
      const maxHeight = Math.max(260, els.previewWrap.clientHeight - 20);
      const displayScale = Math.min(maxWidth / page.width, maxHeight / page.height, 1);
      const cw = Math.max(1, Math.round(page.width * displayScale));
      const ch = Math.max(1, Math.round(page.height * displayScale));

      canvas.width = cw;
      canvas.height = ch;
      g.clearRect(0, 0, cw, ch);
      g.drawImage(img, 0, 0, cw, ch);

      state.render = {
        displayScale,
        coordScaleX: page.coordScaleX,
        coordScaleY: page.coordScaleY,
        pageWidth: page.width,
        pageHeight: page.height
      };

      const rects = getPageAreas(ctx.id, page.pageNum);
      g.lineWidth = 2;
      g.strokeStyle = "#ff6b35";
      g.fillStyle = "rgba(255, 107, 53, 0.25)";
      g.font = "12px Avenir Next, Segoe UI, sans-serif";

      rects.forEach((r, idx) => {
        const x = r[0] * page.coordScaleX * displayScale;
        const y = r[1] * page.coordScaleY * displayScale;
        const w = r[2] * page.coordScaleX * displayScale;
        const h = r[3] * page.coordScaleY * displayScale;
        g.fillRect(x, y, w, h);
        g.strokeRect(x, y, w, h);
        g.fillStyle = "#ff6b35";
        g.fillRect(x, y, 20, 16);
        g.fillStyle = "#ffffff";
        g.fillText(String(idx + 1), x + 6, y + 12);
        g.fillStyle = "rgba(255, 107, 53, 0.25)";
      });

      if (state.draw.active) {
        const x = Math.min(state.draw.startX, state.draw.tempX);
        const y = Math.min(state.draw.startY, state.draw.tempY);
        const w = Math.abs(state.draw.tempX - state.draw.startX);
        const h = Math.abs(state.draw.tempY - state.draw.startY);
        g.strokeStyle = "#f59e0b";
        g.setLineDash([6, 4]);
        g.strokeRect(x, y, w, h);
        g.setLineDash([]);
      }

      updatePageIndicator();
    })
    .catch((error) => {
      log(`Ошибка отрисовки превью: ${normalizeError(error)}`);
    });
}

function renderThumbs() {
  const ctx = getCurrentContext();
  els.thumbs.innerHTML = "";
  if (!ctx) return;

  ctx.pages.forEach((page, index) => {
    const item = document.createElement("button");
    item.type = "button";
    item.className = `thumb ${index === state.currentPageIdx ? "active" : ""}`;
    item.addEventListener("click", () => {
      state.currentPageIdx = index;
      renderThumbs();
      drawCurrentPage();
    });

    const img = document.createElement("img");
    img.src = page.previewUrl;
    img.alt = page.label;

    const meta = document.createElement("div");
    meta.className = "meta";

    const title = document.createElement("div");
    title.className = "title";
    title.textContent = page.label;

    const count = document.createElement("div");
    count.className = "count";
    const areaCount = getPageAreas(ctx.id, page.pageNum).length;
    count.textContent = areaCount > 0 ? `Зон: ${areaCount}` : "Зон нет";

    meta.appendChild(title);
    meta.appendChild(count);
    item.appendChild(img);
    item.appendChild(meta);
    els.thumbs.appendChild(item);
  });
}

function renderFileOptions() {
  els.browserFileSelect.innerHTML = "";
  state.contexts.forEach((ctx, index) => {
    const option = document.createElement("option");
    option.value = String(index);
    option.textContent = `${ctx.file.name} (${ctx.type.toUpperCase()}, ${ctx.pages.length})`;
    els.browserFileSelect.appendChild(option);
  });

  if (state.currentFileIdx >= 0) {
    els.browserFileSelect.value = String(state.currentFileIdx);
  }
}

function selectFile(index) {
  if (index < 0 || index >= state.contexts.length) {
    state.currentFileIdx = -1;
    state.currentPageIdx = -1;
    renderThumbs();
    drawCurrentPage();
    return;
  }

  state.currentFileIdx = index;
  const ctx = state.contexts[index];
  state.currentPageIdx = ctx.pages.length ? 0 : -1;
  renderFileOptions();
  renderThumbs();
  drawCurrentPage();
  updateSummary();
}

function onCanvasPointerDown(event) {
  const page = getCurrentPage();
  if (!page) return;

  const rect = els.previewCanvas.getBoundingClientRect();
  const x = event.clientX - rect.left;
  const y = event.clientY - rect.top;
  state.draw.active = true;
  state.draw.startX = x;
  state.draw.startY = y;
  state.draw.tempX = x;
  state.draw.tempY = y;
  drawCurrentPage();
}

function onCanvasPointerMove(event) {
  if (!state.draw.active) return;
  const rect = els.previewCanvas.getBoundingClientRect();
  state.draw.tempX = event.clientX - rect.left;
  state.draw.tempY = event.clientY - rect.top;
  drawCurrentPage();
}

function onCanvasPointerUp(event) {
  if (!state.draw.active) return;
  state.draw.active = false;

  const ctx = getCurrentContext();
  const page = getCurrentPage();
  if (!ctx || !page) {
    drawCurrentPage();
    return;
  }

  const rect = els.previewCanvas.getBoundingClientRect();
  const endX = event.clientX - rect.left;
  const endY = event.clientY - rect.top;

  const x1 = Math.max(0, Math.min(state.draw.startX, endX));
  const y1 = Math.max(0, Math.min(state.draw.startY, endY));
  const x2 = Math.min(els.previewCanvas.width, Math.max(state.draw.startX, endX));
  const y2 = Math.min(els.previewCanvas.height, Math.max(state.draw.startY, endY));

  const w = x2 - x1;
  const h = y2 - y1;
  if (w < 5 || h < 5) {
    drawCurrentPage();
    return;
  }

  const imageX = x1 / state.render.displayScale;
  const imageY = y1 / state.render.displayScale;
  const imageW = w / state.render.displayScale;
  const imageH = h / state.render.displayScale;

  const coordX = Math.round(imageX / page.coordScaleX);
  const coordY = Math.round(imageY / page.coordScaleY);
  const coordW = Math.round(imageW / page.coordScaleX);
  const coordH = Math.round(imageH / page.coordScaleY);

  if (coordW < 3 || coordH < 3) {
    drawCurrentPage();
    return;
  }

  const areas = getPageAreas(ctx.id, page.pageNum);
  areas.push([coordX, coordY, coordW, coordH]);
  renderThumbs();
  drawCurrentPage();
  updateSummary();
}

function removeLastArea() {
  const ctx = getCurrentContext();
  const page = getCurrentPage();
  if (!ctx || !page) return;
  const areas = getPageAreas(ctx.id, page.pageNum);
  if (!areas.length) return;
  areas.pop();
  renderThumbs();
  drawCurrentPage();
  updateSummary();
}

function clearCurrentPageAreas() {
  const ctx = getCurrentContext();
  const page = getCurrentPage();
  if (!ctx || !page) return;
  getContextAreas(ctx.id)[page.pageNum] = [];
  renderThumbs();
  drawCurrentPage();
  updateSummary();
}

function normalizeRects(rects) {
  const uniq = new Map();
  rects.forEach((r) => {
    const key = r.join(",");
    if (!uniq.has(key)) uniq.set(key, r);
  });
  return [...uniq.values()];
}

function uniqueStrings(values) {
  return [...new Set(values)];
}

function getImageExt(path) {
  const match = path.toLowerCase().match(/\.([a-z0-9]+)$/);
  return match ? match[1] : "";
}

function isProcessableImage(path) {
  const ext = getImageExt(path);
  return ["png", "jpg", "jpeg", "bmp", "webp"].includes(ext);
}

async function buildPdfContext(file, id) {
  const bytes = await file.arrayBuffer();
  const src = await pdfjsLib.getDocument({ data: bytes }).promise;
  const pages = [];

  for (let i = 1; i <= src.numPages; i += 1) {
    const page = await src.getPage(i);
    const previewScale = 0.8;
    const viewport = page.getViewport({ scale: previewScale });
    const canvas = document.createElement("canvas");
    canvas.width = Math.max(1, Math.round(viewport.width));
    canvas.height = Math.max(1, Math.round(viewport.height));
    const g = canvas.getContext("2d", { willReadFrequently: true });
    await page.render({ canvasContext: g, viewport }).promise;

    pages.push({
      pageNum: i,
      label: `Страница ${i}`,
      previewUrl: canvas.toDataURL("image/jpeg", 0.85),
      width: canvas.width,
      height: canvas.height,
      coordScaleX: previewScale,
      coordScaleY: previewScale,
      mediaTargets: []
    });
  }

  return {
    id,
    file,
    type: "pdf",
    pages
  };
}

function extractImageRelationships(relsXmlText) {
  const relDoc = parseXml(relsXmlText);
  const rels = relDoc.getElementsByTagName("Relationship");
  const out = {};
  for (const rel of rels) {
    const type = rel.getAttribute("Type") || "";
    if (!type.includes("/image")) continue;
    const id = rel.getAttribute("Id");
    const target = rel.getAttribute("Target");
    if (id && target) out[id] = target;
  }
  return out;
}

function extractBlipEmbeds(slideXmlText) {
  const slideDoc = parseXml(slideXmlText);
  const blips = slideDoc.getElementsByTagNameNS("*", "blip");
  const out = [];

  for (const blip of blips) {
    const raw = blip.getAttribute("r:embed") || blip.getAttribute("embed");
    if (raw) out.push(raw);
  }
  return out;
}

function buildPlaceholderDataUrl(label) {
  const canvas = document.createElement("canvas");
  canvas.width = 960;
  canvas.height = 540;
  const g = canvas.getContext("2d");
  g.fillStyle = "#111827";
  g.fillRect(0, 0, canvas.width, canvas.height);
  g.strokeStyle = "#334155";
  g.lineWidth = 2;
  g.strokeRect(16, 16, canvas.width - 32, canvas.height - 32);
  g.fillStyle = "#94a3b8";
  g.font = "bold 24px Avenir Next, Segoe UI, sans-serif";
  g.fillText(label, 30, 50);
  g.font = "16px Avenir Next, Segoe UI, sans-serif";
  g.fillText("Не найдено растровое изображение для превью", 30, 82);
  return canvas.toDataURL("image/png");
}

async function buildPptxContext(file, id) {
  const zip = await JSZip.loadAsync(await file.arrayBuffer());
  const slidePaths = Object.keys(zip.files)
    .filter((name) => /^ppt\/slides\/slide\d+\.xml$/i.test(name))
    .sort((a, b) => byNumericTail(a, b, "slide", "\\.xml"));

  const pages = [];

  for (let i = 0; i < slidePaths.length; i += 1) {
    const slidePath = slidePaths[i];
    const relsPath = slidePath.replace("ppt/slides/", "ppt/slides/_rels/") + ".rels";

    const slideXml = await zip.file(slidePath)?.async("string");
    const relsXml = await zip.file(relsPath)?.async("string");

    const relMap = relsXml ? extractImageRelationships(relsXml) : {};
    const embeds = slideXml ? extractBlipEmbeds(slideXml) : [];

    const mediaTargets = uniqueStrings(
      embeds
        .map((embedId) => relMap[embedId])
        .filter(Boolean)
        .map((target) => resolveTargetPath(slidePath, target))
    );

    let previewUrl = "";
    let width = REF_PPTX_WIDTH;
    let height = REF_PPTX_HEIGHT;

    const previewMedia = mediaTargets.find((path) => isProcessableImage(path) && zip.file(path));
    if (previewMedia) {
      const imageBytes = await zip.file(previewMedia).async("uint8array");
      const canvas = await blobToCanvas(new Blob([imageBytes]));
      width = canvas.width;
      height = canvas.height;
      previewUrl = canvas.toDataURL("image/jpeg", 0.85);
    } else {
      previewUrl = buildPlaceholderDataUrl(`Слайд ${i + 1}`);
    }

    pages.push({
      pageNum: i + 1,
      label: `Слайд ${i + 1}`,
      previewUrl,
      width,
      height,
      coordScaleX: width / REF_PPTX_WIDTH,
      coordScaleY: height / REF_PPTX_HEIGHT,
      mediaTargets
    });
  }

  return {
    id,
    file,
    type: "pptx",
    pages
  };
}

async function buildContext(file, index) {
  const ext = fileExt(file.name);
  const id = `${file.name}::${file.size}::${file.lastModified}::${index}`;

  if (ext === "pdf") {
    return buildPdfContext(file, id);
  }
  if (ext === "pptx") {
    return buildPptxContext(file, id);
  }
  return null;
}

async function loadInputFiles() {
  const files = [...els.fileInput.files];
  state.contexts = [];
  state.customAreas = {};
  state.currentFileIdx = -1;
  state.currentPageIdx = -1;
  renderFileOptions();
  renderThumbs();
  drawCurrentPage();
  updateSummary();

  if (!files.length) {
    setStatus("Выбери файлы");
    return;
  }

  setStatus("Подготовка превью...");
  els.processBtn.disabled = true;

  try {
    for (let i = 0; i < files.length; i += 1) {
      const file = files[i];
      log(`Подготовка [${i + 1}/${files.length}]: ${file.name}`);
      const ctx = await buildContext(file, i);
      if (!ctx) {
        log(`Пропуск ${file.name}: формат не поддерживается`);
        continue;
      }
      state.contexts.push(ctx);
      getContextAreas(ctx.id);
    }

    if (!state.contexts.length) {
      setStatus("Нет поддерживаемых файлов");
      return;
    }

    selectFile(0);
    setStatus(`Готово (${state.contexts.length} файлов)`);
  } catch (error) {
    setStatus("Ошибка подготовки");
    log(`ОШИБКА подготовки: ${normalizeError(error)}`);
    console.error(error);
  } finally {
    els.processBtn.disabled = false;
  }
}

async function processPdfContext(context, config) {
  const file = context.file;
  log(`PDF: ${file.name} -> старт`);

  const bytes = await file.arrayBuffer();
  const pdfSrc = await pdfjsLib.getDocument({ data: bytes }).promise;
  const outPdf = await PDFDocument.create();

  for (let i = 1; i <= pdfSrc.numPages; i += 1) {
    const page = await pdfSrc.getPage(i);
    const scale = 2;
    const viewport = page.getViewport({ scale });
    const canvas = document.createElement("canvas");
    canvas.width = Math.floor(viewport.width);
    canvas.height = Math.floor(viewport.height);
    const g = canvas.getContext("2d", { willReadFrequently: true });
    await page.render({ canvasContext: g, viewport }).promise;

    const rects = [];
    if (config.removeDefault) {
      const [x, y, w, h] = config.pdfRectPts;
      rects.push(clampRect([x * scale, y * scale, w * scale, h * scale], canvas.width, canvas.height));
    }

    const custom = getPageAreas(context.id, i);
    custom.forEach(([x, y, w, h]) => {
      rects.push(clampRect([x * scale, y * scale, w * scale, h * scale], canvas.width, canvas.height));
    });

    await inpaintCanvas(canvas, rects);

    const jpgBytes = await new Promise((resolve, reject) => {
      canvas.toBlob(
        async (blob) => {
          if (!blob) {
            reject(new Error("canvas.toBlob вернул пустой результат для PDF-страницы"));
            return;
          }
          resolve(new Uint8Array(await blob.arrayBuffer()));
        },
        "image/jpeg",
        0.95
      );
    });

    const jpg = await outPdf.embedJpg(jpgBytes);
    const outPage = outPdf.addPage([canvas.width, canvas.height]);
    outPage.drawImage(jpg, { x: 0, y: 0, width: canvas.width, height: canvas.height });
    log(`PDF: ${file.name} -> страница ${i}/${pdfSrc.numPages} обработана`);
  }

  const out = await outPdf.save();
  const fileName = withCleanSuffix(file.name, "pdf");
  downloadBytes(out, fileName, "application/pdf");
  log(`PDF: ${file.name} -> готово (${fileName})`);
}

async function processPptxContext(context, config) {
  const file = context.file;
  log(`PPTX: ${file.name} -> старт`);

  const zip = await JSZip.loadAsync(await file.arrayBuffer());
  const mediaRectMap = {};

  context.pages.forEach((page) => {
    const refRects = [];
    if (config.removeDefault) {
      refRects.push([...config.pptxRectRef]);
    }
    const custom = getPageAreas(context.id, page.pageNum);
    custom.forEach((r) => refRects.push([...r]));

    if (!refRects.length) return;

    page.mediaTargets.forEach((mediaPath) => {
      if (!isProcessableImage(mediaPath)) return;
      if (!zip.file(mediaPath)) return;
      if (!mediaRectMap[mediaPath]) mediaRectMap[mediaPath] = [];
      mediaRectMap[mediaPath].push(...refRects);
    });
  });

  const entries = Object.entries(mediaRectMap);
  let processed = 0;

  for (const [mediaPath, refRectsRaw] of entries) {
    const refRects = normalizeRects(refRectsRaw);
    const entry = zip.file(mediaPath);
    if (!entry) continue;

    const srcBlob = new Blob([await entry.async("uint8array")]);
    const canvas = await blobToCanvas(srcBlob);
    const sx = canvas.width / REF_PPTX_WIDTH;
    const sy = canvas.height / REF_PPTX_HEIGHT;

    const pixelRects = refRects.map(([x, y, w, h]) =>
      clampRect([x * sx, y * sy, w * sx, h * sy], canvas.width, canvas.height)
    );

    await inpaintCanvas(canvas, pixelRects);

    const ext = getImageExt(mediaPath);
    const mime = ext === "png" ? "image/png" : "image/jpeg";
    const quality = mime === "image/jpeg" ? 0.95 : undefined;

    const outBlob = await new Promise((resolve, reject) =>
      canvas.toBlob((blob) => (blob ? resolve(blob) : reject(new Error(`canvas.toBlob пустой для ${mediaPath}`))), mime, quality)
    );

    const outBytes = new Uint8Array(await outBlob.arrayBuffer());
    zip.file(mediaPath, outBytes);
    processed += 1;
    log(`PPTX: ${file.name} -> ${mediaPath} обновлен`);
  }

  const outPptx = await zip.generateAsync({ type: "uint8array" });
  const fileName = withCleanSuffix(file.name, "pptx");
  downloadBytes(outPptx, fileName, "application/vnd.openxmlformats-officedocument.presentationml.presentation");
  log(`PPTX: ${file.name} -> готово, изображений обработано: ${processed}`);
}

async function run() {
  if (!state.contexts.length) {
    setStatus("Сначала выбери файлы");
    return;
  }

  els.processBtn.disabled = true;
  els.log.textContent = "";
  setStatus("Инициализация OpenCV...");

  try {
    await waitForOpenCV();
    const config = getConfig();
    setStatus(`Обработка (${state.contexts.length} шт.)...`);
    log(`Начало обработки: ${state.contexts.length} файлов`);

    for (let i = 0; i < state.contexts.length; i += 1) {
      const context = state.contexts[i];
      log(`\n[${i + 1}/${state.contexts.length}] ${context.file.name}`);
      if (context.type === "pdf") {
        await processPdfContext(context, config);
      } else if (context.type === "pptx") {
        await processPptxContext(context, config);
      }
    }

    setStatus("Готово");
    log("\nВсе файлы обработаны");
  } catch (error) {
    setStatus("Ошибка");
    log(`ОШИБКА: ${normalizeError(error)}`);
    console.error(error);
  } finally {
    els.processBtn.disabled = false;
  }
}

els.fileInput.addEventListener("change", () => {
  loadInputFiles();
});

els.browserFileSelect.addEventListener("change", () => {
  const idx = Number(els.browserFileSelect.value);
  selectFile(Number.isFinite(idx) ? idx : -1);
});

els.prevBtn.addEventListener("click", () => {
  const ctx = getCurrentContext();
  if (!ctx || !ctx.pages.length) return;
  if (state.currentPageIdx <= 0) return;
  state.currentPageIdx -= 1;
  renderThumbs();
  drawCurrentPage();
});

els.nextBtn.addEventListener("click", () => {
  const ctx = getCurrentContext();
  if (!ctx || !ctx.pages.length) return;
  if (state.currentPageIdx >= ctx.pages.length - 1) return;
  state.currentPageIdx += 1;
  renderThumbs();
  drawCurrentPage();
});

els.removeLastBtn.addEventListener("click", removeLastArea);
els.clearPageBtn.addEventListener("click", clearCurrentPageAreas);
els.processBtn.addEventListener("click", run);

els.previewCanvas.addEventListener("pointerdown", onCanvasPointerDown);
els.previewCanvas.addEventListener("pointermove", onCanvasPointerMove);
window.addEventListener("pointerup", onCanvasPointerUp);

window.addEventListener("resize", () => {
  drawCurrentPage();
});

updateSummary();
