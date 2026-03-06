import JSZip from "https://esm.sh/jszip@3.10.1";
import { PDFDocument } from "https://esm.sh/pdf-lib@1.17.1";
import * as pdfjsLib from "https://cdn.jsdelivr.net/npm/pdfjs-dist@4.5.136/build/pdf.min.mjs";

pdfjsLib.GlobalWorkerOptions.workerSrc =
  "https://cdn.jsdelivr.net/npm/pdfjs-dist@4.5.136/build/pdf.worker.min.mjs";

const REF_PPTX_WIDTH = 1376;
const REF_PPTX_HEIGHT = 768;

const els = {
  fileInput: document.getElementById("fileInput"),
  processBtn: document.getElementById("processBtn"),
  status: document.getElementById("status"),
  log: document.getElementById("log"),
  removeDefault: document.getElementById("removeDefault"),
  useCustom: document.getElementById("useCustom"),
  customBox: document.getElementById("customBox"),
  pdfX: document.getElementById("pdfX"),
  pdfY: document.getElementById("pdfY"),
  pdfW: document.getElementById("pdfW"),
  pdfH: document.getElementById("pdfH"),
  pptxX: document.getElementById("pptxX"),
  pptxY: document.getElementById("pptxY"),
  pptxW: document.getElementById("pptxW"),
  pptxH: document.getElementById("pptxH"),
  customX: document.getElementById("customX"),
  customY: document.getElementById("customY"),
  customW: document.getElementById("customW"),
  customH: document.getElementById("customH")
};

function log(message) {
  els.log.textContent += `${message}\n`;
  els.log.scrollTop = els.log.scrollHeight;
}

function setStatus(message) {
  els.status.textContent = message;
}

function getNumeric(el, fallback = 0) {
  const value = Number(el.value);
  return Number.isFinite(value) ? value : fallback;
}

function getConfig() {
  const config = {
    removeDefault: els.removeDefault.checked,
    pdfRectPts: [
      getNumeric(els.pdfX),
      getNumeric(els.pdfY),
      getNumeric(els.pdfW),
      getNumeric(els.pdfH)
    ],
    pptxRectRef: [
      getNumeric(els.pptxX),
      getNumeric(els.pptxY),
      getNumeric(els.pptxW),
      getNumeric(els.pptxH)
    ],
    useCustom: els.useCustom.checked,
    customRectPx: [
      getNumeric(els.customX),
      getNumeric(els.customY),
      getNumeric(els.customW),
      getNumeric(els.customH)
    ]
  };
  return config;
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

function clampRect([x, y, w, h], width, height) {
  const cx = Math.max(0, Math.min(Math.round(x), width - 1));
  const cy = Math.max(0, Math.min(Math.round(y), height - 1));
  const cw = Math.max(0, Math.min(Math.round(w), width - cx));
  const ch = Math.max(0, Math.min(Math.round(h), height - cy));
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

function normalizeError(error) {
  if (typeof error === "number" && window.cv?.exceptionFromPtr) {
    try {
      const cvErr = cv.exceptionFromPtr(error);
      const details = cvErr?.msg || cvErr?.toString?.() || `OpenCV error ptr ${error}`;
      return details;
    } catch {
      return `OpenCV numeric error: ${error}`;
    }
  }
  if (error?.message) return error.message;
  return String(error);
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

async function processPdf(file, config) {
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
    const ctx = canvas.getContext("2d", { willReadFrequently: true });

    await page.render({ canvasContext: ctx, viewport }).promise;

    const rects = [];
    if (config.removeDefault) {
      const [x, y, w, h] = config.pdfRectPts;
      rects.push(clampRect([x * scale, y * scale, w * scale, h * scale], canvas.width, canvas.height));
    }
    if (config.useCustom) {
      rects.push(clampRect(config.customRectPx, canvas.width, canvas.height));
    }

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

function getImageExt(path) {
  const match = path.toLowerCase().match(/\.([a-z0-9]+)$/);
  return match ? match[1] : "";
}

function isProcessableImage(path) {
  const ext = getImageExt(path);
  return ["png", "jpg", "jpeg", "bmp", "webp"].includes(ext);
}

async function processPptx(file, config) {
  log(`PPTX: ${file.name} -> старт`);
  const zip = await JSZip.loadAsync(await file.arrayBuffer());
  const mediaEntries = Object.keys(zip.files).filter((name) => name.startsWith("ppt/media/"));

  let processed = 0;
  for (const mediaPath of mediaEntries) {
    if (!isProcessableImage(mediaPath)) continue;
    const entry = zip.file(mediaPath);
    if (!entry) continue;

    const srcBlob = new Blob([await entry.async("uint8array")]);
    const canvas = await blobToCanvas(srcBlob);
    const rects = [];

    if (config.removeDefault) {
      const [x, y, w, h] = config.pptxRectRef;
      const sx = canvas.width / REF_PPTX_WIDTH;
      const sy = canvas.height / REF_PPTX_HEIGHT;
      rects.push(clampRect([x * sx, y * sy, w * sx, h * sy], canvas.width, canvas.height));
    }
    if (config.useCustom) {
      rects.push(clampRect(config.customRectPx, canvas.width, canvas.height));
    }

    await inpaintCanvas(canvas, rects);
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

async function processFile(file, config) {
  const ext = file.name.split(".").pop()?.toLowerCase() ?? "";
  if (ext === "pdf") {
    await processPdf(file, config);
    return;
  }
  if (ext === "pptx") {
    await processPptx(file, config);
    return;
  }
  log(`Пропуск ${file.name}: формат не поддерживается`);
}

async function run() {
  const files = [...els.fileInput.files];
  if (!files.length) {
    setStatus("Сначала выберите файлы");
    return;
  }

  els.processBtn.disabled = true;
  els.log.textContent = "";
  setStatus("Инициализация OpenCV...");

  try {
    await waitForOpenCV();
    const config = getConfig();
    setStatus(`Обработка (${files.length} шт.)...`);
    log(`Начало обработки: ${files.length} файлов`);

    for (let i = 0; i < files.length; i += 1) {
      const file = files[i];
      log(`\n[${i + 1}/${files.length}] ${file.name}`);
      await processFile(file, config);
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

els.useCustom.addEventListener("change", () => {
  els.customBox.classList.toggle("hidden", !els.useCustom.checked);
});

els.processBtn.addEventListener("click", run);
