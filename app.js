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
  dropText: document.getElementById("dropText"),
  fileCount: document.getElementById("fileCount"),
  processBtn: document.getElementById("processBtn"),
  status: document.getElementById("status"),
  log: document.getElementById("log"),
  removeDefault: document.getElementById("removeDefault"),

  browserFileSelect: document.getElementById("browserFileSelect"),
  thumbs: document.getElementById("thumbs"),
  previewCanvas: document.getElementById("previewCanvas"),
  previewWrap: document.getElementById("previewWrap"),
  areasSummary: document.getElementById("areasSummary"),
  areaChips: document.getElementById("areaChips"),
  pageIndicator: document.getElementById("pageIndicator"),

  prevBtn: document.getElementById("prevBtn"),
  nextBtn: document.getElementById("nextBtn"),
  removeLastBtn: document.getElementById("removeLastBtn"),
  clearPageBtn: document.getElementById("clearPageBtn"),

  btnRect: document.getElementById("btnRect"),
  btnBrush: document.getElementById("btnBrush"),
  btnEraser: document.getElementById("btnEraser"),
  brushSize: document.getElementById("brushSize"),
  brushSizeVal: document.getElementById("brushSizeVal"),
  brushSizeLabel: document.getElementById("brushSizeLabel"),
  btnUndo: document.getElementById("btnUndo"),
  btnRedo: document.getElementById("btnRedo")
};

const state = {
  contexts: [],
  customAreas: {},
  brushMasks: {},
  historyUndo: [],
  historyRedo: [],

  currentFileIdx: -1,
  currentPageIdx: -1,
  currentTool: "rect",
  brushSizePx: 16,

  render: {
    displayScale: 1,
    coordScaleX: 1,
    coordScaleY: 1,
    pageWidth: 1,
    pageHeight: 1
  },

  rectDraw: {
    active: false,
    startX: 0,
    startY: 0,
    tempX: 0,
    tempY: 0
  },

  strokeDraw: {
    active: false,
    lastCoordX: 0,
    lastCoordY: 0,
    beforeSnapshot: null
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
    cv.rectangle(mask, new cv.Point(x, y), new cv.Point(x + w, y + h), new cv.Scalar(255), -1);
  });
}

async function blobToCanvas(blob) {
  const bitmap = await createImageBitmap(blob);
  const canvas = document.createElement("canvas");
  canvas.width = bitmap.width;
  canvas.height = bitmap.height;
  canvas.getContext("2d", { willReadFrequently: true }).drawImage(bitmap, 0, 0);
  bitmap.close();
  return canvas;
}

function canvasHasInk(canvas) {
  if (!canvas) return false;
  const ctx = canvas.getContext("2d", { willReadFrequently: true });
  const data = ctx.getImageData(0, 0, canvas.width, canvas.height).data;
  for (let i = 3; i < data.length; i += 4) {
    if (data[i] > 2) return true;
  }
  return false;
}

async function inpaintCanvas(canvas, rects, radius = 5, brushMaskCanvas = null) {
  if (!rects.length && !brushMaskCanvas) return canvas;

  const srcRgba = cv.imread(canvas);
  const src = new cv.Mat();
  cv.cvtColor(srcRgba, src, cv.COLOR_RGBA2RGB);

  const mask = new cv.Mat.zeros(src.rows, src.cols, cv.CV_8UC1);

  let brushRgba = null;
  let brushGray = null;
  let brushBin = null;

  try {
    drawRectsMask(mask, rects);

    if (brushMaskCanvas && canvasHasInk(brushMaskCanvas)) {
      brushRgba = cv.imread(brushMaskCanvas);
      brushGray = new cv.Mat();
      brushBin = new cv.Mat();
      cv.cvtColor(brushRgba, brushGray, cv.COLOR_RGBA2GRAY);
      cv.threshold(brushGray, brushBin, 1, 255, cv.THRESH_BINARY);
      cv.bitwise_or(mask, brushBin, mask);
    }

    const dst = new cv.Mat();
    cv.inpaint(src, mask, dst, radius, cv.INPAINT_NS);
    cv.imshow(canvas, dst);
    dst.delete();
    return canvas;
  } finally {
    if (brushRgba) brushRgba.delete();
    if (brushGray) brushGray.delete();
    if (brushBin) brushBin.delete();
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
    if (p === "..") out.pop();
    else out.push(p);
  }
  return out.join("/");
}

function resolveTargetPath(basePath, target) {
  const baseParts = basePath.split("/");
  baseParts.pop();
  return normalizeZipPath(`${baseParts.join("/")}/${target}`);
}

function uniqueStrings(values) {
  return [...new Set(values)];
}

function normalizeRects(rects) {
  const uniq = new Map();
  rects.forEach((r) => {
    const k = r.join(",");
    if (!uniq.has(k)) uniq.set(k, r);
  });
  return [...uniq.values()];
}

function getContextAreas(contextId) {
  if (!state.customAreas[contextId]) state.customAreas[contextId] = {};
  return state.customAreas[contextId];
}

function getPageAreas(contextId, pageNum) {
  const fileAreas = getContextAreas(contextId);
  if (!fileAreas[pageNum]) fileAreas[pageNum] = [];
  return fileAreas[pageNum];
}

function getContextBrushes(contextId) {
  if (!state.brushMasks[contextId]) state.brushMasks[contextId] = {};
  return state.brushMasks[contextId];
}

function getPageBrushCanvas(contextId, page, create = false) {
  const brushes = getContextBrushes(contextId);
  if (brushes[page.pageNum]) return brushes[page.pageNum];
  if (!create) return null;

  const c = document.createElement("canvas");
  c.width = page.coordWidth;
  c.height = page.coordHeight;
  brushes[page.pageNum] = c;
  return c;
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
    for (const rects of Object.values(areasByPage)) total += rects.length;
  }
  for (const brushByPage of Object.values(state.brushMasks)) {
    for (const c of Object.values(brushByPage)) if (canvasHasInk(c)) total += 1;
  }

  return total;
}

function updateSummary() {
  const total = countAllAreas();
  els.areasSummary.textContent = `Доп. области: ${total}`;
  els.areaChips.innerHTML = "";
  if (!total) {
    els.areaChips.innerHTML = '<span class="no-areas">Доп. область не задана</span>';
    return;
  }

  const ctx = getCurrentContext();
  if (!ctx) return;

  ctx.pages.forEach((page) => {
    const rectCount = getPageAreas(ctx.id, page.pageNum).length;
    const brush = getPageBrushCanvas(ctx.id, page, false);
    const hasBrush = canvasHasInk(brush);
    if (!rectCount && !hasBrush) return;

    if (rectCount) {
      const chip = document.createElement("span");
      chip.className = "area-chip";
      chip.textContent = `⬜ ${page.label}: ${rectCount}`;
      els.areaChips.appendChild(chip);
    }
    if (hasBrush) {
      const chip = document.createElement("span");
      chip.className = "area-chip";
      chip.textContent = `🖌 ${page.label}`;
      els.areaChips.appendChild(chip);
    }
  });
}

function updatePageIndicator() {
  const ctx = getCurrentContext();
  if (!ctx || state.currentPageIdx < 0) {
    els.pageIndicator.textContent = "— / —";
    return;
  }
  els.pageIndicator.textContent = `${state.currentPageIdx + 1} / ${ctx.pages.length}`;
}

function setTool(tool) {
  state.currentTool = tool;

  els.btnRect.classList.remove("active-rect");
  els.btnBrush.classList.remove("active-brush");
  els.btnEraser.classList.remove("active-eraser");

  if (tool === "rect") {
    els.btnRect.classList.add("active-rect");
    els.brushSizeLabel.textContent = "Размер кисти";
    els.previewCanvas.style.cursor = "crosshair";
  } else if (tool === "brush") {
    els.btnBrush.classList.add("active-brush");
    els.brushSizeLabel.textContent = "Размер кисти";
    els.previewCanvas.style.cursor = "none";
  } else {
    els.btnEraser.classList.add("active-eraser");
    els.brushSizeLabel.textContent = "Размер ластика";
    els.previewCanvas.style.cursor = "none";
  }
}

function getPageSnapshot(contextId, page) {
  const rects = getPageAreas(contextId, page.pageNum).map((r) => [...r]);
  const brush = getPageBrushCanvas(contextId, page, false);
  const brushData = brush && canvasHasInk(brush) ? brush.toDataURL("image/png") : null;
  return { contextId, pageNum: page.pageNum, rects, brushData };
}

function applyPageSnapshot(snapshot) {
  const context = state.contexts.find((c) => c.id === snapshot.contextId);
  if (!context) return;
  const page = context.pages.find((p) => p.pageNum === snapshot.pageNum);
  if (!page) return;

  getContextAreas(context.id)[page.pageNum] = snapshot.rects.map((r) => [...r]);

  if (snapshot.brushData) {
    const c = getPageBrushCanvas(context.id, page, true);
    const g = c.getContext("2d");
    g.clearRect(0, 0, c.width, c.height);
    const img = new Image();
    img.onload = () => {
      g.drawImage(img, 0, 0, c.width, c.height);
      if (getCurrentContext()?.id === context.id && getCurrentPage()?.pageNum === page.pageNum) {
        renderThumbs();
        drawCurrentPage();
        updateSummary();
      }
    };
    img.src = snapshot.brushData;
  } else {
    const c = getPageBrushCanvas(context.id, page, false);
    if (c) c.getContext("2d").clearRect(0, 0, c.width, c.height);
  }
}

function pushHistory(before, after) {
  state.historyUndo.push({ before, after });
  state.historyRedo = [];
}

function undo() {
  const action = state.historyUndo.pop();
  if (!action) return;
  state.historyRedo.push(action);
  applyPageSnapshot(action.before);
  renderThumbs();
  drawCurrentPage();
  updateSummary();
}

function redo() {
  const action = state.historyRedo.pop();
  if (!action) return;
  state.historyUndo.push(action);
  applyPageSnapshot(action.after);
  renderThumbs();
  drawCurrentPage();
  updateSummary();
}

function drawCurrentPage() {
  const page = getCurrentPage();
  const context = getCurrentContext();
  const canvas = els.previewCanvas;
  const g = canvas.getContext("2d");

  if (!page || !context) {
    canvas.width = 1;
    canvas.height = 1;
    g.clearRect(0, 0, 1, 1);
    updatePageIndicator();
    return;
  }

  Promise.resolve(page.previewCanvas || null)
    .then(async (sourceCanvas) => {
      if (sourceCanvas) return sourceCanvas;
      if (page.image) return page.image;

      const img = new Image();
      await new Promise((resolve, reject) => {
        img.onload = resolve;
        img.onerror = reject;
        img.src = page.previewUrl;
      });
      page.image = img;
      return img;
    })
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

      const brush = getPageBrushCanvas(context.id, page, false);
      if (brush && canvasHasInk(brush)) {
        g.save();
        g.globalAlpha = 0.65;
        g.drawImage(brush, 0, 0, cw, ch);
        g.restore();
      }

      const rects = getPageAreas(context.id, page.pageNum);
      g.lineWidth = 2;
      g.strokeStyle = "#ff6b35";
      g.fillStyle = "rgba(255, 107, 53, 0.25)";
      g.font = "12px Nunito, sans-serif";

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

      if (state.rectDraw.active && state.currentTool === "rect") {
        const x = Math.min(state.rectDraw.startX, state.rectDraw.tempX);
        const y = Math.min(state.rectDraw.startY, state.rectDraw.tempY);
        const w = Math.abs(state.rectDraw.tempX - state.rectDraw.startX);
        const h = Math.abs(state.rectDraw.tempY - state.rectDraw.startY);
        g.strokeStyle = "#f59e0b";
        g.setLineDash([6, 4]);
        g.strokeRect(x, y, w, h);
        g.setLineDash([]);
      }

      updatePageIndicator();
    })
    .catch((error) => log(`Ошибка отрисовки превью: ${normalizeError(error)}`));
}

function renderThumbs() {
  const context = getCurrentContext();
  els.thumbs.innerHTML = "";
  if (!context) return;

  context.pages.forEach((page, index) => {
    const item = document.createElement("button");
    item.type = "button";
    item.className = `thumb ${index === state.currentPageIdx ? "active" : ""}`;
    item.addEventListener("click", () => {
      state.currentPageIdx = index;
      renderThumbs();
      drawCurrentPage();
      updateSummary();
    });

    const img = document.createElement("img");
    img.src = page.thumbUrl || page.previewUrl;
    img.alt = page.label;

    const meta = document.createElement("div");
    meta.className = "meta";

    const title = document.createElement("div");
    title.className = "title";
    title.textContent = page.label;

    const count = document.createElement("div");
    count.className = "count";
    const rectCount = getPageAreas(context.id, page.pageNum).length;
    const hasBrush = canvasHasInk(getPageBrushCanvas(context.id, page, false));
    if (!rectCount && !hasBrush) count.textContent = "Зон нет";
    else if (rectCount && hasBrush) count.textContent = `Зон: ${rectCount} + кисть`;
    else if (rectCount) count.textContent = `Зон: ${rectCount}`;
    else count.textContent = "Кисть";

    meta.appendChild(title);
    meta.appendChild(count);
    item.appendChild(img);
    item.appendChild(meta);
    els.thumbs.appendChild(item);
  });
}

function renderFileOptions() {
  els.browserFileSelect.innerHTML = "";
  state.contexts.forEach((context, index) => {
    const option = document.createElement("option");
    option.value = String(index);
    option.textContent = `${context.file.name} (${context.type.toUpperCase()}, ${context.pages.length})`;
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
    updateSummary();
    return;
  }

  state.currentFileIdx = index;
  state.currentPageIdx = state.contexts[index].pages.length ? 0 : -1;
  renderFileOptions();
  renderThumbs();
  drawCurrentPage();
  updateSummary();
}

function eventToCanvasPos(event) {
  const rect = els.previewCanvas.getBoundingClientRect();
  return {
    x: Math.max(0, Math.min(els.previewCanvas.width, event.clientX - rect.left)),
    y: Math.max(0, Math.min(els.previewCanvas.height, event.clientY - rect.top))
  };
}

function canvasPosToCoord(pos, page) {
  const imageX = pos.x / state.render.displayScale;
  const imageY = pos.y / state.render.displayScale;
  return {
    x: Math.max(0, Math.min(page.coordWidth, imageX / page.coordScaleX)),
    y: Math.max(0, Math.min(page.coordHeight, imageY / page.coordScaleY))
  };
}

function drawBrushStroke(maskCanvas, fromCoord, toCoord, tool, page) {
  const g = maskCanvas.getContext("2d");
  const avgScale = state.render.displayScale * ((page.coordScaleX + page.coordScaleY) / 2);
  const lineW = Math.max(1, state.brushSizePx / Math.max(avgScale, 0.001));

  g.save();
  g.lineCap = "round";
  g.lineJoin = "round";
  g.lineWidth = lineW;

  if (tool === "brush") {
    g.globalCompositeOperation = "source-over";
    g.strokeStyle = "rgba(240, 123, 63, 0.9)";
  } else {
    g.globalCompositeOperation = "destination-out";
    g.strokeStyle = "rgba(0,0,0,1)";
  }

  g.beginPath();
  g.moveTo(fromCoord.x, fromCoord.y);
  g.lineTo(toCoord.x, toCoord.y);
  g.stroke();
  g.restore();
}

function onCanvasPointerDown(event) {
  const context = getCurrentContext();
  const page = getCurrentPage();
  if (!context || !page) return;

  const pos = eventToCanvasPos(event);

  if (state.currentTool === "rect") {
    state.rectDraw.active = true;
    state.rectDraw.startX = pos.x;
    state.rectDraw.startY = pos.y;
    state.rectDraw.tempX = pos.x;
    state.rectDraw.tempY = pos.y;
    drawCurrentPage();
    return;
  }

  const before = getPageSnapshot(context.id, page);
  state.strokeDraw.beforeSnapshot = before;
  state.strokeDraw.active = true;

  const coord = canvasPosToCoord(pos, page);
  state.strokeDraw.lastCoordX = coord.x;
  state.strokeDraw.lastCoordY = coord.y;

  const maskCanvas = getPageBrushCanvas(context.id, page, true);
  drawBrushStroke(maskCanvas, coord, coord, state.currentTool, page);
  drawCurrentPage();
}

function onCanvasPointerMove(event) {
  const context = getCurrentContext();
  const page = getCurrentPage();
  if (!context || !page) return;

  if (state.currentTool === "rect") {
    if (!state.rectDraw.active) return;
    const pos = eventToCanvasPos(event);
    state.rectDraw.tempX = pos.x;
    state.rectDraw.tempY = pos.y;
    drawCurrentPage();
    return;
  }

  if (!state.strokeDraw.active) return;
  const pos = eventToCanvasPos(event);
  const coord = canvasPosToCoord(pos, page);
  const maskCanvas = getPageBrushCanvas(context.id, page, true);

  drawBrushStroke(maskCanvas, { x: state.strokeDraw.lastCoordX, y: state.strokeDraw.lastCoordY }, coord, state.currentTool, page);

  state.strokeDraw.lastCoordX = coord.x;
  state.strokeDraw.lastCoordY = coord.y;
  drawCurrentPage();
}

function onCanvasPointerUp(event) {
  const context = getCurrentContext();
  const page = getCurrentPage();
  if (!context || !page) return;

  if (state.currentTool === "rect") {
    if (!state.rectDraw.active) return;
    state.rectDraw.active = false;

    const pos = eventToCanvasPos(event);
    const x1 = Math.max(0, Math.min(state.rectDraw.startX, pos.x));
    const y1 = Math.max(0, Math.min(state.rectDraw.startY, pos.y));
    const x2 = Math.min(els.previewCanvas.width, Math.max(state.rectDraw.startX, pos.x));
    const y2 = Math.min(els.previewCanvas.height, Math.max(state.rectDraw.startY, pos.y));

    const w = x2 - x1;
    const h = y2 - y1;
    if (w < 5 || h < 5) {
      drawCurrentPage();
      return;
    }

    const before = getPageSnapshot(context.id, page);

    const c1 = canvasPosToCoord({ x: x1, y: y1 }, page);
    const c2 = canvasPosToCoord({ x: x2, y: y2 }, page);

    const coordRect = [Math.round(c1.x), Math.round(c1.y), Math.max(1, Math.round(c2.x - c1.x)), Math.max(1, Math.round(c2.y - c1.y))];
    getPageAreas(context.id, page.pageNum).push(coordRect);

    const after = getPageSnapshot(context.id, page);
    pushHistory(before, after);

    renderThumbs();
    drawCurrentPage();
    updateSummary();
    return;
  }

  if (!state.strokeDraw.active) return;
  state.strokeDraw.active = false;

  const before = state.strokeDraw.beforeSnapshot;
  const after = getPageSnapshot(context.id, page);
  if (before && JSON.stringify(before) !== JSON.stringify(after)) {
    pushHistory(before, after);
  }

  state.strokeDraw.beforeSnapshot = null;
  renderThumbs();
  drawCurrentPage();
  updateSummary();
}

function removeLastArea() {
  const context = getCurrentContext();
  const page = getCurrentPage();
  if (!context || !page) return;
  const areas = getPageAreas(context.id, page.pageNum);
  if (!areas.length) return;

  const before = getPageSnapshot(context.id, page);
  areas.pop();
  const after = getPageSnapshot(context.id, page);
  pushHistory(before, after);

  renderThumbs();
  drawCurrentPage();
  updateSummary();
}

function clearCurrentPageAreas() {
  const context = getCurrentContext();
  const page = getCurrentPage();
  if (!context || !page) return;

  const before = getPageSnapshot(context.id, page);
  getContextAreas(context.id)[page.pageNum] = [];
  const brush = getPageBrushCanvas(context.id, page, false);
  if (brush) brush.getContext("2d").clearRect(0, 0, brush.width, brush.height);

  const after = getPageSnapshot(context.id, page);
  pushHistory(before, after);

  renderThumbs();
  drawCurrentPage();
  updateSummary();
}

function getImageExt(path) {
  const match = path.toLowerCase().match(/\.([a-z0-9]+)$/);
  return match ? match[1] : "";
}

function isProcessableImage(path) {
  return ["png", "jpg", "jpeg", "bmp", "webp"].includes(getImageExt(path));
}

async function buildPdfContext(file, id) {
  const bytes = await file.arrayBuffer();
  const src = await pdfjsLib.getDocument({ data: bytes, disableWorker: true }).promise;
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

    const coordWidth = Math.max(1, Math.round(canvas.width / previewScale));
    const coordHeight = Math.max(1, Math.round(canvas.height / previewScale));

    pages.push({
      pageNum: i,
      label: `Страница ${i}`,
      previewCanvas: canvas,
      previewUrl: canvas.toDataURL("image/jpeg", 0.85),
      thumbUrl: canvas.toDataURL("image/jpeg", 0.6),
      width: canvas.width,
      height: canvas.height,
      coordScaleX: previewScale,
      coordScaleY: previewScale,
      coordWidth,
      coordHeight,
      mediaTargets: []
    });
  }

  return { id, file, type: "pdf", pages };
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

function buildPlaceholderCanvas(label) {
  const canvas = document.createElement("canvas");
  canvas.width = 960;
  canvas.height = 540;
  const g = canvas.getContext("2d");
  g.fillStyle = "#f0f3f8";
  g.fillRect(0, 0, canvas.width, canvas.height);
  g.fillStyle = "#9aa8ba";
  g.font = "bold 24px Nunito, sans-serif";
  g.fillText(label, 32, 48);
  g.font = "16px Nunito, sans-serif";
  g.fillText("Для этого слайда превью недоступно", 32, 80);
  return canvas;
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

    let previewCanvas = null;
    let previewUrl = "";
    let thumbUrl = "";
    let width = REF_PPTX_WIDTH;
    let height = REF_PPTX_HEIGHT;

    const previewMedia = mediaTargets.find((path) => isProcessableImage(path) && zip.file(path));
    if (previewMedia) {
      const imageBytes = await zip.file(previewMedia).async("uint8array");
      previewCanvas = await blobToCanvas(new Blob([imageBytes]));
      width = previewCanvas.width;
      height = previewCanvas.height;
      previewUrl = previewCanvas.toDataURL("image/jpeg", 0.85);
      thumbUrl = previewCanvas.toDataURL("image/jpeg", 0.6);
    } else {
      previewCanvas = buildPlaceholderCanvas(`Слайд ${i + 1}`);
      width = previewCanvas.width;
      height = previewCanvas.height;
      previewUrl = previewCanvas.toDataURL("image/jpeg", 0.85);
      thumbUrl = previewCanvas.toDataURL("image/jpeg", 0.6);
    }

    pages.push({
      pageNum: i + 1,
      label: `Слайд ${i + 1}`,
      previewCanvas,
      previewUrl,
      thumbUrl,
      width,
      height,
      coordScaleX: width / REF_PPTX_WIDTH,
      coordScaleY: height / REF_PPTX_HEIGHT,
      coordWidth: REF_PPTX_WIDTH,
      coordHeight: REF_PPTX_HEIGHT,
      mediaTargets
    });
  }

  return { id, file, type: "pptx", pages };
}

async function buildContext(file, index) {
  const ext = fileExt(file.name);
  const id = `${file.name}::${file.size}::${file.lastModified}::${index}`;

  if (ext === "pdf") return buildPdfContext(file, id);
  if (ext === "pptx") return buildPptxContext(file, id);
  return null;
}

async function loadInputFiles() {
  const files = [...els.fileInput.files];
  state.contexts = [];
  state.customAreas = {};
  state.brushMasks = {};
  state.historyUndo = [];
  state.historyRedo = [];
  state.currentFileIdx = -1;
  state.currentPageIdx = -1;

  renderFileOptions();
  renderThumbs();
  drawCurrentPage();
  updateSummary();

  els.fileCount.textContent = String(files.length);
  els.dropText.textContent = files.length ? (files.length === 1 ? files[0].name : `${files.length} файлов выбрано`) : "Перетащите файлы сюда";

  if (!files.length) return;

  els.processBtn.disabled = true;
  setStatus("Подготовка превью...");

  try {
    for (let i = 0; i < files.length; i += 1) {
      const file = files[i];
      log(`Подготовка [${i + 1}/${files.length}]: ${file.name}`);
      const context = await buildContext(file, i);
      if (!context) {
        log(`Пропуск ${file.name}: формат не поддерживается`);
        continue;
      }
      state.contexts.push(context);
      getContextAreas(context.id);
      getContextBrushes(context.id);
    }

    if (state.contexts.length) {
      selectFile(0);
      setStatus(`Готово (${state.contexts.length})`);
    }
  } catch (error) {
    setStatus("Ошибка подготовки");
    log(`ОШИБКА подготовки: ${normalizeError(error)}`);
    console.error(error);
  } finally {
    els.processBtn.disabled = false;
  }
}

function buildBrushMaskForTarget(context, page, targetWidth, targetHeight) {
  const source = getPageBrushCanvas(context.id, page, false);
  if (!source || !canvasHasInk(source)) return null;

  const target = document.createElement("canvas");
  target.width = targetWidth;
  target.height = targetHeight;
  const g = target.getContext("2d", { willReadFrequently: true });
  g.drawImage(source, 0, 0, targetWidth, targetHeight);
  return target;
}

async function processPdfContext(context, config) {
  const file = context.file;
  log(`PDF: ${file.name} -> старт`);

  const bytes = await file.arrayBuffer();
  const pdfSrc = await pdfjsLib.getDocument({ data: bytes, disableWorker: true }).promise;
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
    custom.forEach(([x, y, w, h]) => rects.push(clampRect([x * scale, y * scale, w * scale, h * scale], canvas.width, canvas.height)));

    const pageMeta = context.pages.find((p) => p.pageNum === i);
    const brushMask = pageMeta ? buildBrushMaskForTarget(context, pageMeta, canvas.width, canvas.height) : null;

    await inpaintCanvas(canvas, rects, 5, brushMask);

    const jpgBytes = await new Promise((resolve, reject) => {
      canvas.toBlob(async (blob) => {
        if (!blob) {
          reject(new Error("canvas.toBlob вернул пустой результат для PDF-страницы"));
          return;
        }
        resolve(new Uint8Array(await blob.arrayBuffer()));
      }, "image/jpeg", 0.95);
    });

    const jpg = await outPdf.embedJpg(jpgBytes);
    const outPage = outPdf.addPage([canvas.width, canvas.height]);
    outPage.drawImage(jpg, { x: 0, y: 0, width: canvas.width, height: canvas.height });
    log(`PDF: ${file.name} -> страница ${i}/${pdfSrc.numPages} обработана`);
  }

  const out = await outPdf.save();
  downloadBytes(out, withCleanSuffix(file.name, "pdf"), "application/pdf");
  log(`PDF: ${file.name} -> готово`);
}

async function processPptxContext(context, config) {
  const file = context.file;
  log(`PPTX: ${file.name} -> старт`);

  const zip = await JSZip.loadAsync(await file.arrayBuffer());
  const mediaMap = {};

  context.pages.forEach((page) => {
    const refRects = [];
    if (config.removeDefault) refRects.push([...config.pptxRectRef]);
    getPageAreas(context.id, page.pageNum).forEach((r) => refRects.push([...r]));

    const hasBrush = canvasHasInk(getPageBrushCanvas(context.id, page, false));

    if (!refRects.length && !hasBrush) return;

    page.mediaTargets.forEach((mediaPath) => {
      if (!isProcessableImage(mediaPath)) return;
      if (!zip.file(mediaPath)) return;
      if (!mediaMap[mediaPath]) mediaMap[mediaPath] = { rects: [], brushPages: [] };
      mediaMap[mediaPath].rects.push(...refRects);
      if (hasBrush) mediaMap[mediaPath].brushPages.push(page.pageNum);
    });
  });

  let processed = 0;

  for (const [mediaPath, payload] of Object.entries(mediaMap)) {
    const entry = zip.file(mediaPath);
    if (!entry) continue;

    const canvas = await blobToCanvas(new Blob([await entry.async("uint8array")]));
    const sx = canvas.width / REF_PPTX_WIDTH;
    const sy = canvas.height / REF_PPTX_HEIGHT;

    const rects = normalizeRects(payload.rects).map(([x, y, w, h]) =>
      clampRect([x * sx, y * sy, w * sx, h * sy], canvas.width, canvas.height)
    );

    let mergedBrushMask = null;
    if (payload.brushPages.length) {
      mergedBrushMask = document.createElement("canvas");
      mergedBrushMask.width = canvas.width;
      mergedBrushMask.height = canvas.height;
      const mg = mergedBrushMask.getContext("2d");

      for (const pn of uniqueStrings(payload.brushPages)) {
        const page = context.pages.find((p) => p.pageNum === pn);
        if (!page) continue;
        const bm = buildBrushMaskForTarget(context, page, canvas.width, canvas.height);
        if (bm) mg.drawImage(bm, 0, 0);
      }
      if (!canvasHasInk(mergedBrushMask)) mergedBrushMask = null;
    }

    await inpaintCanvas(canvas, rects, 5, mergedBrushMask);

    const ext = getImageExt(mediaPath);
    const mime = ext === "png" ? "image/png" : "image/jpeg";
    const quality = mime === "image/jpeg" ? 0.95 : undefined;

    const outBlob = await new Promise((resolve, reject) =>
      canvas.toBlob((blob) => (blob ? resolve(blob) : reject(new Error(`canvas.toBlob пустой для ${mediaPath}`))), mime, quality)
    );

    zip.file(mediaPath, new Uint8Array(await outBlob.arrayBuffer()));
    processed += 1;
    log(`PPTX: ${file.name} -> ${mediaPath} обновлен`);
  }

  const outPptx = await zip.generateAsync({ type: "uint8array" });
  downloadBytes(outPptx, withCleanSuffix(file.name, "pptx"), "application/vnd.openxmlformats-officedocument.presentationml.presentation");
  log(`PPTX: ${file.name} -> готово, изображений обработано: ${processed}`);
}

async function run() {
  if (!state.contexts.length) return;

  els.processBtn.disabled = true;
  els.log.textContent = "";
  setStatus("Инициализация OpenCV...");

  try {
    await waitForOpenCV();
    const config = getConfig();
    setStatus(`Обработка (${state.contexts.length})...`);
    log(`Начало обработки: ${state.contexts.length} файлов`);

    for (let i = 0; i < state.contexts.length; i += 1) {
      const context = state.contexts[i];
      log(`\n[${i + 1}/${state.contexts.length}] ${context.file.name}`);
      if (context.type === "pdf") await processPdfContext(context, config);
      else if (context.type === "pptx") await processPptxContext(context, config);
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

els.fileInput.addEventListener("change", loadInputFiles);

els.browserFileSelect.addEventListener("change", () => {
  const idx = Number(els.browserFileSelect.value);
  selectFile(Number.isFinite(idx) ? idx : -1);
});

els.prevBtn.addEventListener("click", () => {
  const context = getCurrentContext();
  if (!context || state.currentPageIdx <= 0) return;
  state.currentPageIdx -= 1;
  renderThumbs();
  drawCurrentPage();
  updateSummary();
});

els.nextBtn.addEventListener("click", () => {
  const context = getCurrentContext();
  if (!context || state.currentPageIdx >= context.pages.length - 1) return;
  state.currentPageIdx += 1;
  renderThumbs();
  drawCurrentPage();
  updateSummary();
});

els.removeLastBtn.addEventListener("click", removeLastArea);
els.clearPageBtn.addEventListener("click", clearCurrentPageAreas);
els.processBtn.addEventListener("click", run);

els.btnRect.addEventListener("click", () => setTool("rect"));
els.btnBrush.addEventListener("click", () => setTool("brush"));
els.btnEraser.addEventListener("click", () => setTool("eraser"));

els.brushSize.addEventListener("input", () => {
  state.brushSizePx = Number(els.brushSize.value) || 16;
  els.brushSizeVal.textContent = `${state.brushSizePx}px`;
});

els.btnUndo.addEventListener("click", undo);
els.btnRedo.addEventListener("click", redo);

els.previewCanvas.addEventListener("pointerdown", onCanvasPointerDown);
els.previewCanvas.addEventListener("pointermove", onCanvasPointerMove);
window.addEventListener("pointerup", onCanvasPointerUp);

window.addEventListener("resize", drawCurrentPage);

setTool("rect");
updateSummary();
els.log.textContent = "Watermark Remover готов. Выберите файлы для обработки.\n";
