import path from "node:path";
import { MAX_IMAGE_BLOCKS, MAX_RETURN_CHARS, MAX_TOTAL_IMAGE_BYTES } from "../config.mjs";

export function createParseContext() {
  return {
    imageBlocks: [],
    maxImageBlocks: MAX_IMAGE_BLOCKS,
    imageBytesUsed: 0,
    maxTotalImageBytes: MAX_TOTAL_IMAGE_BYTES,
    imageNotes: [],
  };
}

export function bufferToArrayBuffer(buffer) {
  return buffer.buffer.slice(
    buffer.byteOffset,
    buffer.byteOffset + buffer.byteLength
  );
}

export function getExtension(filename = "") {
  const ext = path.extname(filename || "").toLowerCase();
  return ext.startsWith(".") ? ext.slice(1) : ext;
}

export function normalizeExtractedText(text) {
  return String(text || "")
    .replace(/\u0000/g, "")
    .replace(/\r\n/g, "\n")
    .replace(/\r/g, "\n")
    .replace(/[ \t]+\n/g, "\n")
    .replace(/\n{3,}/g, "\n\n")
    .trim();
}

export function decodeXmlEntities(text) {
  return String(text || "")
    .replace(/&#x([0-9a-f]+);/gi, (_, hex) =>
      String.fromCodePoint(parseInt(hex, 16))
    )
    .replace(/&#(\d+);/g, (_, dec) =>
      String.fromCodePoint(parseInt(dec, 10))
    )
    .replace(/&lt;/g, "<")
    .replace(/&gt;/g, ">")
    .replace(/&quot;/g, '"')
    .replace(/&apos;/g, "'")
    .replace(/&amp;/g, "&");
}

export function extractTextFromXml(xml) {
  const withBreaks = String(xml || "")
    .replace(/<a:br\s*\/>/gi, "\n")
    .replace(/<w:br\s*\/>/gi, "\n")
    .replace(/<a:tab\s*\/>/gi, "\t")
    .replace(/<w:tab\s*\/>/gi, "\t")
    .replace(/<\/(a:p|w:p|p:txBody|text:p|p:notesTextViewPr)>/gi, "\n");

  const matches = [...withBreaks.matchAll(/>([^<>]+)</g)]
    .map((match) => decodeXmlEntities(match[1]))
    .map((value) => value.replace(/[ \t]+/g, " ").trim())
    .filter(Boolean);

  return normalizeExtractedText(matches.join("\n"));
}

export function truncateExtractedText(text, maxLength = MAX_RETURN_CHARS) {
  const totalLength = String(text || "").length;

  if (totalLength <= maxLength) {
    return {
      text,
      totalLength,
      returnedLength: totalLength,
      omittedLength: 0,
      truncated: false,
    };
  }

  return {
    text: `${String(text || "").slice(0, maxLength)}\n\n... (parsed output truncated)`,
    totalLength,
    returnedLength: maxLength,
    omittedLength: totalLength - maxLength,
    truncated: true,
  };
}

export function toNumberAwareSort(entries, pattern) {
  return [...entries].sort((a, b) => {
    const aNum = Number(a.entryName.match(pattern)?.[1] || 0);
    const bNum = Number(b.entryName.match(pattern)?.[1] || 0);
    return aNum - bNum;
  });
}

export function addContextNote(context, note) {
  if (!note) return;
  if (!context.imageNotes.includes(note)) {
    context.imageNotes.push(note);
  }
}

export function isProbablyTextLike(ext, contentType = "", kind = "") {
  return (
    String(contentType || "").toLowerCase().startsWith("text/") ||
    ["txt", "md", "json", "xml", "html", "csv"].includes(ext) ||
    kind === "itemAttachment"
  );
}

export function formatMetadataBlock(lines) {
  const visible = lines.filter(Boolean);
  return visible.length ? visible.join("\n") : "";
}

export function appendSection(parts, title, body) {
  const normalized = normalizeExtractedText(body);
  if (!normalized) return;
  parts.push(`### ${title}\n${normalized}`);
}

export function sanitizeSnippet(text) {
  return normalizeExtractedText(String(text || "").replace(/[^\S\n]+/g, " "));
}

export function isMeaningfulSnippet(text) {
  const normalized = sanitizeSnippet(text);
  if (normalized.length < 4) return false;

  const signalChars = (normalized.match(/[A-Za-z0-9\u00C0-\uFFFF]/g) || []).length;
  return signalChars >= Math.max(4, Math.floor(normalized.length * 0.3));
}

export function extractPrintableStrings(buffer) {
  const results = [];
  const seen = new Set();

  const utf16Matches =
    buffer
      .toString("utf16le")
      .match(/[\p{L}\p{N}][\p{L}\p{N}\p{P}\p{Zs}\t\n\r]{3,}/gu) || [];
  const asciiMatches =
    buffer.toString("latin1").match(/[A-Za-z0-9][ -~\t\n\r]{3,}/g) || [];

  for (const candidate of [...utf16Matches, ...asciiMatches]) {
    const snippet = sanitizeSnippet(candidate);
    if (!isMeaningfulSnippet(snippet) || seen.has(snippet)) continue;
    seen.add(snippet);
    results.push(snippet);
  }

  return normalizeExtractedText(results.join("\n"));
}

export function htmlToPlainText(html) {
  return normalizeExtractedText(
    decodeXmlEntities(
      String(html || "")
        .replace(/<(script|style)[^>]*>[\s\S]*?<\/\1>/gi, "")
        .replace(/<\/(p|div|li|tr|h[1-6]|blockquote)>/gi, "\n")
        .replace(/<br\s*\/?>/gi, "\n")
        .replace(/<li[^>]*>/gi, "- ")
        .replace(/<[^>]+>/g, "")
    )
  );
}

// ---------------------------------------------------------------------------
// Outbound mail helpers (only used when SEND_ENABLED)
// ---------------------------------------------------------------------------
