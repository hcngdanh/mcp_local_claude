import AdmZip from "adm-zip";
import { createRequire } from "node:module";
import { MAX_DOC_IMAGES } from "../config.mjs";
import { log } from "../logger.mjs";
import { appendSection, extractPrintableStrings, extractTextFromXml, normalizeExtractedText, toNumberAwareSort } from "./common.mjs";
import { collectZipImages } from "./image.mjs";

const require = createRequire(import.meta.url);

export const CFB = require("cfb");

export function isPowerPointZipExt(ext) {
  return ["pptx", "pptm", "ppsx", "potx"].includes(ext);
}

export function csvEscapeCell(value) {
  const normalized =
    value === null || value === undefined
      ? ""
      : value instanceof Date
        ? value.toISOString()
        : String(value);

  if (/[",\n\r]/.test(normalized)) {
    return `"${normalized.replace(/"/g, "\"\"")}"`;
  }

  return normalized;
}

export function worksheetToCsv(worksheet) {
  let maxColumn = 0;

  worksheet.eachRow({ includeEmpty: true }, (row) => {
    const values = Array.isArray(row.values) ? row.values.slice(1) : [];
    maxColumn = Math.max(maxColumn, values.length);
  });

  const lines = [];
  worksheet.eachRow({ includeEmpty: true }, (row) => {
    const values = Array.isArray(row.values) ? row.values.slice(1) : [];
    const cells = [];

    for (let index = 0; index < maxColumn; index += 1) {
      const cell = row.getCell(index + 1);
      let value = cell?.text;

      if (!value && value !== "0") {
        value = cell?.result ?? cell?.value ?? "";
      }

      cells.push(csvEscapeCell(value));
    }

    lines.push(cells.join(","));
  });

  return lines.join("\n");
}

export function extractLegacyPptText(buffer) {
  try {
    const cfb = CFB.parse(buffer, { type: "buffer" });
    const chunks = [];

    for (const entry of cfb.FileIndex || []) {
      if (entry.type !== 2 || !entry.content) continue;
      const entryName = String(entry.name || "");
      if (
        /PowerPoint Document|Current User|SummaryInformation|DocumentSummaryInformation/i.test(
          entryName
        )
      ) {
        chunks.push(extractPrintableStrings(Buffer.from(entry.content)));
      }
    }

    const combined = normalizeExtractedText(chunks.join("\n"));
    if (combined) return combined;
  } catch (err) {
    log("Legacy PPT parse fallback", err?.message || String(err));
  }

  return extractPrintableStrings(buffer);
}

export async function parsePptx(raw, context) {
  const zip = new AdmZip(raw);
  const slideEntries = toNumberAwareSort(
    zip.getEntries().filter((entry) => /^ppt\/slides\/slide\d+\.xml$/i.test(entry.entryName)),
    /slide(\d+)\.xml/i
  );

  const noteEntries = new Map(
    toNumberAwareSort(
      zip.getEntries().filter((entry) => /^ppt\/notesSlides\/notesSlide\d+\.xml$/i.test(entry.entryName)),
      /notesSlide(\d+)\.xml/i
    ).map((entry) => [
      Number(entry.entryName.match(/notesSlide(\d+)\.xml/i)?.[1] || 0),
      extractTextFromXml(entry.getData().toString("utf8")),
    ])
  );

  const parts = [];

  for (const entry of slideEntries) {
    const slideNumber = Number(entry.entryName.match(/slide(\d+)\.xml/i)?.[1] || 0);
    const slideText = extractTextFromXml(entry.getData().toString("utf8"));
    appendSection(parts, `Slide ${slideNumber || parts.length + 1}`, slideText);

    const noteText = noteEntries.get(slideNumber);
    if (noteText) {
      appendSection(parts, `Slide ${slideNumber || parts.length + 1} Notes`, noteText);
    }
  }

  const addedImages = await collectZipImages(zip, "ppt/media/", context, MAX_DOC_IMAGES);
  let extractedText = normalizeExtractedText(parts.join("\n\n"));

  if (!extractedText) {
    extractedText = "(No readable slide text was detected in this presentation.)";
  }

  if (addedImages > 0) {
    extractedText += `\n\n[System: MCP extracted ${addedImages} embedded presentation image(s) and sent them as visual blocks.]`;
  }

  return extractedText;
}
