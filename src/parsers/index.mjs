import ExcelJS from "exceljs";
import mammoth from "mammoth";
import AdmZip from "adm-zip";
import { MAX_DOC_IMAGES } from "../config.mjs";
import { log } from "../logger.mjs";
import { parse7z, parseRar, parseZip } from "./archive.mjs";
import { extractPrintableStrings, getExtension, isProbablyTextLike, normalizeExtractedText } from "./common.mjs";
import { addImageBlock, collectZipImages, inferImageMimeType, isImageAttachment } from "./image.mjs";
import { parseMsg } from "./msg.mjs";
import { extractLegacyPptText, isPowerPointZipExt, parsePptx, worksheetToCsv } from "./office.mjs";
import { parsePdf, runPdfOcr } from "./pdf.mjs";

export async function parseAttachmentPayload(payload, context, depth = 0) {
  const { raw, filename, contentType = "", kind = "attachment" } = payload;
  const ext = getExtension(filename);

  if (isImageAttachment(ext, contentType)) {
    const mimeType = inferImageMimeType(ext, contentType);
    const attached = await addImageBlock(context, raw, mimeType);

    return {
      parser: "image_passthrough",
      text: normalizeExtractedText(
        [
          `[Image attachment] ${filename}`,
          attached.attached
            ? attached.transformed
              ? `The local backend downscaled this image from ${Math.round((attached.originalBytes || 0) / 1024)} KB to ${Math.round((attached.finalBytes || 0) / 1024)} KB and sent it as a visual block so the model can inspect it directly.`
              : "The local backend sent this image as a visual block so the model can inspect it directly."
            : attached.reason === "single_image_too_large"
              ? "The backend detected an image, but even after downscaling it remained too large for Claude Desktop's MCP tool-result payload budget."
              : attached.reason === "total_image_budget_reached"
                ? "The backend detected an image, but the total inline image budget for this tool result was already reached."
                : "The backend detected an image, but it could not be attached inline.",
        ].join("\n\n")
      ),
      details: {
        attachedImage: attached.attached,
        imageResult: attached,
      },
    };
  }

  if (ext === "pdf") {
    let extractedText = normalizeExtractedText((await parsePdf(raw))?.text || "");
    const notes = [];
    let ocrUsed = false;
    let ocrPageCount = 0;

    if (!extractedText) {
      notes.push(
        "No embedded PDF text layer was detected. The backend treated this as a scanned/image PDF and attempted OCR."
      );

      try {
        const ocr = await runPdfOcr(raw, context);
        ocrPageCount = ocr.pageCount;
        if (ocr.text) {
          extractedText = ocr.text;
          ocrUsed = true;
          notes.push(`OCR extracted readable text from ${ocr.pageCount} page(s).`);
        } else {
          notes.push("OCR completed but did not recover readable text from the rendered pages.");
        }
      } catch (ocrErr) {
        log("PDF OCR error", ocrErr);
        notes.push(`OCR could not run successfully: ${ocrErr?.message || String(ocrErr)}`);
      }
    }

    if (!extractedText) {
      extractedText =
        "(This PDF appears to be image-based or scanned. No readable text was extracted.)";
    }

    return {
      parser: ocrUsed ? "pdf_ocr" : "pdf_text",
      text: normalizeExtractedText([...notes, extractedText].join("\n\n")),
      details: {
        ocrUsed,
        ocrPageCount,
      },
    };
  }

  if (ext === "docx" || ext === "doc") {
    const result = await mammoth.extractRawText({ buffer: raw });
    let extractedText =
      normalizeExtractedText(result.value) ||
      "(No readable plain text detected in this document.)";

    if (ext === "docx") {
      try {
        const zip = new AdmZip(raw);
        const imageCount = await collectZipImages(zip, "word/media/", context, MAX_DOC_IMAGES);
        if (imageCount > 0) {
          extractedText += `\n\n[System: MCP extracted ${imageCount} embedded document image(s) and sent them as visual blocks.]`;
        }
      } catch (zipErr) {
        log("DOCX image extraction error", zipErr);
      }
    }

    return {
      parser: ext === "docx" ? "docx" : "doc",
      text: extractedText,
      details: {},
    };
  }

  if (isPowerPointZipExt(ext)) {
    return {
      parser: "pptx",
      text: await parsePptx(raw, context),
      details: {},
    };
  }

  if (ext === "ppt") {
    const extractedText = extractLegacyPptText(raw);
    return {
      parser: "ppt_legacy",
      text: normalizeExtractedText(
        [
          "Legacy .ppt parsing uses best-effort binary text extraction.",
          extractedText || "(No readable text was extracted from this legacy PowerPoint file.)",
        ].join("\n\n")
      ),
      details: {},
    };
  }

  if (ext === "csv") {
    return {
      parser: "csv",
      text: raw.toString("utf8"),
      details: {},
    };
  }

  if (ext === "xlsx" || ext === "xlsm") {
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.load(raw);
    const sheetTexts = [];

    for (const worksheet of workbook.worksheets) {
      const csv = worksheetToCsv(worksheet);
      sheetTexts.push(`### Sheet: ${worksheet.name}\n\`\`\`csv\n${csv}\n\`\`\``);
    }

    return {
      parser: ext,
      text: normalizeExtractedText(sheetTexts.join("\n\n")),
      details: {
        sheetCount: workbook.worksheets.length,
      },
    };
  }

  if (ext === "xls") {
    const extractedText = extractPrintableStrings(raw);

    return {
      parser: "xls_legacy",
      text: normalizeExtractedText(
        [
          "Legacy .xls parsing uses best-effort binary text extraction.",
          extractedText || "(No readable text was extracted from this legacy Excel file.)",
        ].join("\n\n")
      ),
      details: {},
    };
  }

  if (ext === "msg" || String(contentType).toLowerCase() === "application/vnd.ms-outlook") {
    return parseMsg(raw, filename, context, depth);
  }

  if (ext === "zip") {
    return parseZip(raw, context, depth);
  }

  if (ext === "rar") {
    return parseRar(raw, context, depth);
  }

  if (ext === "7z") {
    return parse7z(raw, filename, context, depth);
  }

  if (isProbablyTextLike(ext, contentType, kind)) {
    return {
      parser: "text",
      text: raw.toString("utf8"),
      details: {},
    };
  }

  return {
    parser: "unsupported",
    text: `(Note: This file format is not supported for direct parsing.)\nFilename: ${filename}\nType: ${contentType || "application/octet-stream"}`,
    details: {},
  };
}
