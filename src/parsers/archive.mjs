import fs from "node:fs";
import os from "node:os";
import path from "node:path";
import AdmZip from "adm-zip";
import { MAX_ARCHIVE_ENTRY_BYTES, MAX_ARCHIVE_ITEMS, MAX_ARCHIVE_LISTED_ENTRIES, MAX_ARCHIVE_TOTAL_BYTES, MAX_NESTED_DEPTH } from "../config.mjs";
import { appendSection, bufferToArrayBuffer, formatMetadataBlock, normalizeExtractedText } from "./common.mjs";
import { parseAttachmentPayload } from "./index.mjs";
import { safeFilename } from "../util.mjs";

export let rarDepsPromise = null;

export let sevenZipDepsPromise = null;

export function isIgnoredArchiveEntry(name) {
  const normalized = String(name || "").replace(/\\/g, "/");
  const base = path.posix.basename(normalized);

  return (
    !base ||
    base.startsWith(".") ||
    normalized.startsWith("__MACOSX/") ||
    normalized.endsWith("/.DS_Store")
  );
}

export async function loadRarDeps() {
  if (!rarDepsPromise) {
    rarDepsPromise = import("node-unrar-js").then((module) => {
      const createExtractorFromData =
        module.createExtractorFromData ||
        module.default?.createExtractorFromData;

      if (typeof createExtractorFromData !== "function") {
        throw new Error("RAR dependency loaded but createExtractorFromData export is missing.");
      }

      return { createExtractorFromData };
    });
  }

  return rarDepsPromise;
}

export async function loadSevenZipDeps() {
  if (!sevenZipDepsPromise) {
    sevenZipDepsPromise = Promise.all([
      import("node-7z"),
      import("7zip-bin"),
    ]).then(([sevenModule, binModule]) => {
      const extractFull =
        sevenModule.extractFull || sevenModule.default?.extractFull;
      const list = sevenModule.list || sevenModule.default?.list;
      const path7za = binModule.path7za || binModule.default?.path7za;

      if (typeof extractFull !== "function" || typeof list !== "function" || !path7za) {
        throw new Error("7z dependencies loaded but did not expose the expected API.");
      }

      return { extractFull, list, path7za };
    });
  }

  return sevenZipDepsPromise;
}

export async function parseArchiveEntries(entries, archiveKind, context, depth) {
  const visibleEntries = entries.filter(
    (entry) => entry && entry.name && !isIgnoredArchiveEntry(entry.name)
  );
  const selectedEntries = visibleEntries.slice(0, MAX_ARCHIVE_ITEMS);
  const parts = [];

  appendSection(
    parts,
    `${archiveKind.toUpperCase()} Archive Summary`,
    formatMetadataBlock([
      `Files in archive: ${visibleEntries.length}`,
      visibleEntries.length > MAX_ARCHIVE_ITEMS
        ? `Parsed entries: ${MAX_ARCHIVE_ITEMS} (remaining entries were summarized only)`
        : `Parsed entries: ${selectedEntries.length}`,
    ])
  );

  for (const entry of selectedEntries) {
    const entrySize = Math.max(entry.buffer?.length ?? 0, entry.size ?? 0);

    if (entrySize > MAX_ARCHIVE_ENTRY_BYTES) {
      appendSection(
        parts,
        entry.name,
        `Skipped because the extracted entry is too large (${entrySize} bytes).`
      );
      continue;
    }

    if (!entry.buffer) {
      appendSection(
        parts,
        entry.name,
        entry.skipReason || "(Archive entry payload was unavailable.)"
      );
      continue;
    }

    if (depth >= MAX_NESTED_DEPTH) {
      appendSection(
        parts,
        entry.name,
        "(Nested archive parsing skipped because the recursion limit was reached.)"
      );
      continue;
    }

    const parsed = await parseAttachmentPayload(
      {
        raw: entry.buffer,
        filename: entry.name,
        contentType: entry.contentType || "",
        kind: "attachment",
      },
      context,
      depth + 1
    );

    appendSection(parts, entry.name, parsed.text);
  }

  if (visibleEntries.length > MAX_ARCHIVE_ITEMS) {
    appendSection(
      parts,
      `${archiveKind.toUpperCase()} Archive Overflow`,
      `Only the first ${MAX_ARCHIVE_ITEMS} archive entries were parsed out of ${visibleEntries.length}.`
    );
  }

  return {
    parser: `${archiveKind}_archive`,
    text:
      normalizeExtractedText(parts.join("\n\n")) ||
      `(${archiveKind.toUpperCase()} archive had no readable entries.)`,
    details: {
      archiveEntryCount: visibleEntries.length,
      parsedArchiveEntryCount: selectedEntries.length,
    },
  };
}

// Decide, from declared (uncompressed) sizes only, which archive entries are
// safe to decompress. Entries are matched to parseArchiveEntries' selection
// (same ignore filter, first MAX_ARCHIVE_ITEMS), and each must fit both the
// per-entry and the per-archive budget. Returns a Set of selected names.
export function selectArchiveEntriesWithinBudget(listing) {
  const selected = new Set();
  let budget = MAX_ARCHIVE_TOTAL_BYTES;
  let considered = 0;

  for (const item of listing) {
    if (!item.name || isIgnoredArchiveEntry(item.name)) continue;
    if (considered >= MAX_ARCHIVE_ITEMS) break;
    considered += 1;

    const size = Number(item.size) || 0;
    if (size > MAX_ARCHIVE_ENTRY_BYTES || size > budget) continue;

    budget -= size;
    selected.add(item.name);
  }

  return selected;
}

export const ARCHIVE_BUDGET_SKIP_REASON =
  "(Skipped: not decompressed because it is too large or the archive's total extraction budget was exhausted.)";

export async function parseZip(raw, context, depth) {
  const zip = new AdmZip(raw);
  const files = zip
    .getEntries()
    .filter((entry) => !entry.isDirectory)
    .slice(0, MAX_ARCHIVE_LISTED_ENTRIES)
    .map((entry) => ({
      entry,
      name: entry.entryName.replace(/\\/g, "/"),
      size: entry.header?.size || 0,
    }));

  const selected = selectArchiveEntriesWithinBudget(files);
  const entries = files.map(({ entry, name, size }) => {
    if (!selected.has(name)) {
      return { name, buffer: null, size, skipReason: ARCHIVE_BUDGET_SKIP_REASON };
    }
    const buffer = entry.getData();
    // Never trust the declared size alone.
    if (buffer.length > MAX_ARCHIVE_ENTRY_BYTES) {
      return { name, buffer: null, size: buffer.length };
    }
    return { name, buffer, size: buffer.length };
  });

  return parseArchiveEntries(entries, "zip", context, depth);
}

export async function parseRar(raw, context, depth) {
  const { createExtractorFromData } = await loadRarDeps();
  const extractor = await createExtractorFromData({ data: bufferToArrayBuffer(raw) });
  const list = extractor.getFileList();
  const fileHeaders = [...list.fileHeaders]
    .filter((header) => header?.name && !header?.flags?.directory)
    .slice(0, MAX_ARCHIVE_LISTED_ENTRIES);

  const selected = selectArchiveEntriesWithinBudget(
    fileHeaders.map((header) => ({ name: header.name, size: header.unpSize }))
  );

  const extractedByName = new Map();
  if (selected.size > 0) {
    const extracted = extractor.extract({ files: [...selected] });
    for (const file of extracted.files) {
      const name = file?.fileHeader?.name;
      if (name && file?.extraction) {
        extractedByName.set(name, Buffer.from(file.extraction));
      }
    }
  }

  const entries = fileHeaders.map((header) => {
    const buffer = extractedByName.get(header.name) || null;
    return {
      name: header.name,
      buffer,
      size: buffer?.length ?? header.unpSize ?? 0,
      skipReason: selected.has(header.name) ? undefined : ARCHIVE_BUDGET_SKIP_REASON,
    };
  });

  return parseArchiveEntries(entries, "rar", context, depth);
}

export function list7zEntries(archivePath, path7za, listFn) {
  return new Promise((resolve, reject) => {
    const items = [];
    const stream = listFn(archivePath, { $bin: path7za });
    stream.on("data", (item) => {
      if (items.length >= MAX_ARCHIVE_LISTED_ENTRIES) return;
      const isDir = String(item?.attributes || "").startsWith("D");
      if (item?.file && !isDir) {
        items.push({ name: String(item.file).replace(/\\/g, "/"), size: item.size || 0 });
      }
    });
    stream.on("error", reject);
    stream.on("end", () => resolve(items));
  });
}

export async function parse7z(raw, filename, context, depth) {
  const { extractFull, list, path7za } = await loadSevenZipDeps();
  const tempDir = fs.mkdtempSync(path.join(os.tmpdir(), "mcp-local-7z-"));
  const archivePath = path.join(tempDir, safeFilename(filename || "archive.7z"));
  const extractDir = path.join(tempDir, "out");

  fs.mkdirSync(extractDir, { recursive: true });
  fs.writeFileSync(archivePath, raw);

  try {
    // List first and only extract entries whose declared sizes fit the budget,
    // so a 7z bomb can't fill the disk.
    const listing = await list7zEntries(archivePath, path7za, list);
    const selected = selectArchiveEntriesWithinBudget(listing);

    if (selected.size > 0) {
      await new Promise((resolve, reject) => {
        const stream = extractFull(archivePath, extractDir, {
          $bin: path7za,
          $progress: false,
          $cherryPick: [...selected],
          // -spd: treat cherry-picked names literally (no wildcard expansion).
          $raw: ["-spd"],
          overwrite: "a",
        });

        stream.on("error", reject);
        stream.on("end", resolve);
      });
    }

    const resolvedExtractDir = path.resolve(extractDir);
    let remaining = MAX_ARCHIVE_TOTAL_BYTES;
    const entries = listing.map((item) => {
      if (!selected.has(item.name)) {
        return { ...item, buffer: null, skipReason: ARCHIVE_BUDGET_SKIP_REASON };
      }
      const fullPath = path.resolve(extractDir, item.name);
      if (!fullPath.startsWith(resolvedExtractDir + path.sep)) {
        return { ...item, buffer: null, skipReason: "(Skipped: unsafe entry path.)" };
      }
      let stat;
      try {
        stat = fs.lstatSync(fullPath);
      } catch {
        return { ...item, buffer: null };
      }
      if (!stat.isFile() || stat.size > MAX_ARCHIVE_ENTRY_BYTES || stat.size > remaining) {
        return { ...item, buffer: null, size: stat.size, skipReason: ARCHIVE_BUDGET_SKIP_REASON };
      }
      remaining -= stat.size;
      return { name: item.name, buffer: fs.readFileSync(fullPath), size: stat.size };
    });

    return parseArchiveEntries(entries, "7z", context, depth);
  } finally {
    fs.rmSync(tempDir, { recursive: true, force: true });
  }
}
