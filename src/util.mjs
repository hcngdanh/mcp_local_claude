import { normalizeExtractedText } from "./parsers/common.mjs";

export function normalizeNeedle(value) {
  return normalizeExtractedText(value).toLowerCase();
}

export function matchesNeedle(value, needle) {
  if (!needle) return true;
  return String(value || "").toLowerCase().includes(needle);
}

export function formatSender(fromName, fromAddress) {
  if (fromName && fromAddress) {
    return `${fromName} <${fromAddress}>`;
  }

  return fromName || fromAddress || "(unknown)";
}

export function safeFilename(name) {
  return (name || "attachment.bin").replace(/[<>:"/\\|?*\x00-\x1F]/g, "_");
}
