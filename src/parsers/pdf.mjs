import { createRequire } from "node:module";
import { MAX_PDF_OCR_PAGES, MAX_PDF_VISUAL_PAGES, OCR_CACHE_DIR, OCR_LANGS } from "../config.mjs";
import { normalizeExtractedText } from "./common.mjs";
import { addImageBlock } from "./image.mjs";

const require = createRequire(import.meta.url);

export const _pdfParseRaw = require("pdf-parse");

export let parsePdf;

if (typeof _pdfParseRaw === "function") {
  parsePdf = _pdfParseRaw;
} else if (_pdfParseRaw && typeof _pdfParseRaw.PDFParse === "function") {
  parsePdf = async (buffer) => {
    const instance = new _pdfParseRaw.PDFParse();
    return instance.parse(buffer);
  };
} else {
  parsePdf = async () => {
    throw new Error(
      "pdf-parse module loaded but no usable export was found. Please pin pdf-parse to ^1.1.1."
    );
  };
}

export let pdfOcrDepsPromise = null;

export async function loadPdfOcrDeps() {
  if (!pdfOcrDepsPromise) {
    pdfOcrDepsPromise = Promise.all([
      import("pdf-to-img"),
      import("tesseract.js"),
    ]).then(([pdfToImgModule, tesseractModule]) => {
      const pdf =
        pdfToImgModule.pdf ||
        pdfToImgModule.default?.pdf ||
        pdfToImgModule.default;
      const createWorker =
        tesseractModule.createWorker ||
        tesseractModule.default?.createWorker;

      if (typeof pdf !== "function" || typeof createWorker !== "function") {
        throw new Error("OCR dependencies loaded but did not expose the expected API.");
      }

      return { pdf, createWorker };
    });
  }

  return pdfOcrDepsPromise;
}

export async function runPdfOcr(raw, context) {
  const { pdf, createWorker } = await loadPdfOcrDeps();
  // PDFs come from arbitrary email senders: never let pdf.js compile code
  // from document content (defense in depth against font/JS-eval exploits).
  const document = await pdf(new Uint8Array(raw), {
    scale: 2,
    docInitParams: {
      isEvalSupported: false,
      enableXfa: false,
      disableFontFace: true,
    },
  });

  const worker = await createWorker(OCR_LANGS, undefined, {
    cachePath: OCR_CACHE_DIR,
  }).catch(async (err) => {
    await document.destroy?.().catch(() => { });
    throw err;
  });
  const pageTexts = [];
  let pageCount = 0;
  let attachedPages = 0;

  try {
    for await (const pageImage of document) {
      pageCount += 1;
      if (pageCount > MAX_PDF_OCR_PAGES) break;

      const imageBuffer = Buffer.from(pageImage);

      if (attachedPages < MAX_PDF_VISUAL_PAGES) {
        const attachmentResult = await addImageBlock(context, imageBuffer, "image/png");
        if (attachmentResult.attached) {
          attachedPages += 1;
        }
      }

      const result = await worker.recognize(imageBuffer);
      const pageText = normalizeExtractedText(result?.data?.text || "");
      if (pageText) {
        pageTexts.push(`### OCR Page ${pageCount}\n${pageText}`);
      }
    }
  } finally {
    await worker.terminate().catch(() => { });
    await document.destroy?.().catch(() => { });
  }

  return {
    text: normalizeExtractedText(pageTexts.join("\n\n")),
    pageCount: Math.min(pageCount, MAX_PDF_OCR_PAGES),
    attachedPages,
  };
}
