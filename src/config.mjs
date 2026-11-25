import dotenv from "dotenv";
import fs from "node:fs";
import os from "node:os";
import path from "node:path";
import { fileURLToPath } from "node:url";

// Load .env before any configuration is read.
dotenv.config({ quiet: true });

export const APP_NAME = "m365-attachment-reader-mcp-local";

export const APP_DATA_DIR = resolveAppDataDir();

export function resolveAppDataDir() {
  // Project root (this file lives in <root>/src/).
  const moduleDir = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
  const candidates = [
    process.env.M365_MCP_DATA_DIR,
    process.platform === "win32"
      ? path.join(process.env.LOCALAPPDATA || os.homedir(), APP_NAME)
      : path.join(os.homedir(), `.${APP_NAME}`),
    path.join(moduleDir, `.${APP_NAME}`),
  ].filter(Boolean);

  for (const candidate of candidates) {
    try {
      fs.mkdirSync(candidate, { recursive: true });
      return candidate;
    } catch {
      // Try the next writable location.
    }
  }

  throw new Error(
    `Unable to create an application data directory. Tried: ${candidates.join(", ")}`
  );
}

export const DEBUG_LOG = path.join(APP_DATA_DIR, "debug.log");

export const CLIENT_ID = process.env.M365_CLIENT_ID;

export const TENANT_ID = process.env.M365_TENANT_ID || "common";

export const AUTHORITY = `https://login.microsoftonline.com/${TENANT_ID}`;

// Treat blank values and unsubstituted MCPB placeholders ("${user_config.x}")
// as "not set".
export function readEnv(name) {
  const value = String(process.env[name] ?? "").trim();
  return !value || value.includes("${") ? "" : value;
}

export function envFlag(name, defaultValue = false) {
  const value = readEnv(name).toLowerCase();
  if (!value) return defaultValue;
  return ["1", "true", "yes", "on"].includes(value);
}

export function envList(name, separator) {
  return readEnv(name)
    .split(separator)
    .map((item) => item.trim())
    .filter((item) => item && !item.includes("${"));
}

// ---------------------------------------------------------------------------
// Outbound mail (opt-in). Both switches default OFF and are exposed as toggles
// in the Claude Desktop extension settings (see manifest.json user_config).
//   M365_ENABLE_SEND=true     -> registers send_outlook_email (drafts only)
//   M365_ALLOW_SEND_NOW=true  -> additionally allows send_now=true
// When composing is disabled the server stays strictly read-only and only
// requests Mail.Read. Changing these requires re-running begin_auth so the
// new delegated permissions can be consented.
// ---------------------------------------------------------------------------
export const SEND_ENABLED = envFlag("M365_ENABLE_SEND");

export const SEND_NOW_ALLOWED = SEND_ENABLED && envFlag("M365_ALLOW_SEND_NOW");

export const SEND_ALLOWED_DOMAINS = envList("M365_SEND_ALLOWED_DOMAINS", /[,;\s]+/).map((domain) =>
  domain.toLowerCase().replace(/^@/, "")
);

// Local attachment roots. Empty => local file attachments are DISABLED.
export const SEND_ALLOWED_DIRS = envList("M365_SEND_ALLOWED_DIRS", path.delimiter)
  .map((dir) => {
    try {
      return fs.realpathSync(path.resolve(dir));
    } catch {
      return null;
    }
  })
  .filter(Boolean);

export const SCOPES = [
  "User.Read",
  // Creating drafts / adding attachments needs Mail.ReadWrite (which includes read).
  SEND_ENABLED ? "Mail.ReadWrite" : "Mail.Read",
  ...(SEND_NOW_ALLOWED ? ["Mail.Send"] : []),
];

export const GRAPH_BASE = "https://graph.microsoft.com/v1.0";

export const AUTO_OPEN_BROWSER =
  String(process.env.M365_AUTO_OPEN_BROWSER || "true").toLowerCase() !== "false";

export const MAX_RETURN_CHARS = 35000;

export const MAX_IMAGE_BLOCKS = 20;

export const MAX_IMAGE_BLOCK_BYTES = 450 * 1024;

export const MAX_TOTAL_IMAGE_BYTES = 700 * 1024;

export const IMAGE_RESIZE_STEPS = [1600, 1280, 1024, 800, 640, 480];

export const MAX_DOC_IMAGES = 15;

export const MAX_PDF_OCR_PAGES = 5;

export const MAX_PDF_VISUAL_PAGES = 3;

// OCR languages for scanned PDFs (tesseract codes joined by "+"). Defaults to
// Vietnamese + English. Language data is downloaded once on first use and
// cached under the app data dir (not the process working directory).
export const OCR_LANGS = (() => {
  const langs = String(process.env.M365_OCR_LANGS || "")
    .split(/[+,\s]+/)
    .map((lang) => lang.trim())
    .filter((lang) => /^[a-z_]{3,10}$/i.test(lang));
  return langs.length > 0 ? langs : ["vie", "eng"];
})();

export const OCR_CACHE_DIR = path.join(APP_DATA_DIR, "tessdata");

try {
  fs.mkdirSync(OCR_CACHE_DIR, { recursive: true });
} catch {
  // Falls back to tesseract's default behavior if this fails at runtime.
}

export const MAX_ARCHIVE_ITEMS = 12;

export const MAX_ARCHIVE_ENTRY_BYTES = 20 * 1024 * 1024;

// Upper bound on bytes decompressed from a single archive. Checked against the
// archive's declared entry sizes *before* decompressing, so zip/7z bombs are
// rejected without allocating memory or filling the disk.
export const MAX_ARCHIVE_TOTAL_BYTES = 60 * 1024 * 1024;

export const MAX_ARCHIVE_LISTED_ENTRIES = 5000;

export const MAX_NESTED_DEPTH = 2;

export const MAX_MSG_ATTACHMENTS = 5;

export const IMAGE_MIME_BY_EXT = {
  jpg: "image/jpeg",
  jpeg: "image/jpeg",
  png: "image/png",
  gif: "image/gif",
  webp: "image/webp",
  bmp: "image/bmp",
  tif: "image/tiff",
  tiff: "image/tiff",
};
