import fs from "node:fs";
import path from "node:path";
import { GRAPH_BASE, IMAGE_MIME_BY_EXT, SEND_ALLOWED_DIRS, SEND_ALLOWED_DOMAINS, SEND_NOW_ALLOWED } from "../config.mjs";
import { graphGetBytes, graphGetJson, graphRequest } from "../graph.mjs";
import { log } from "../logger.mjs";
import { getExtension } from "../parsers/common.mjs";

export const MAX_SEND_RECIPIENTS = 20;

export const MAX_SEND_ATTACHMENTS = 10;

// Graph's per-request limit is ~4 MB and base64 inflates ~33%, so attachments
// are inlined only while the inline total stays under this budget; anything
// else goes through an attachment upload session.
export const MAX_INLINE_ATTACHMENT_BYTES = 2.5 * 1024 * 1024;

export const MAX_SEND_TOTAL_BYTES = 35 * 1024 * 1024;

// Upload-session chunks must be multiples of 320 KiB and under 4 MB.
export const UPLOAD_CHUNK_BYTES = 320 * 1024 * 12;

export const MAX_LIST_SENDABLE_FILES = 200;

export const MAX_LIST_SENDABLE_DEPTH = 3;

export const EMAIL_RE = /^[^\s@<>(),;:"\[\]]+@[^\s@<>(),;:"\[\]]+\.[^\s@<>(),;:"\[\]]+$/;

export const MIME_BY_EXT = {
  pptx: "application/vnd.openxmlformats-officedocument.presentationml.presentation",
  ppt: "application/vnd.ms-powerpoint",
  docx: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
  doc: "application/msword",
  xlsx: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  xlsm: "application/vnd.ms-excel.sheet.macroEnabled.12",
  xls: "application/vnd.ms-excel",
  pdf: "application/pdf",
  csv: "text/csv",
  txt: "text/plain",
  md: "text/markdown",
  json: "application/json",
  xml: "application/xml",
  html: "text/html",
  zip: "application/zip",
  "7z": "application/x-7z-compressed",
  rar: "application/vnd.rar",
  ...IMAGE_MIME_BY_EXT,
};

export function formatBytes(bytes) {
  if (bytes < 1024) return `${bytes} B`;
  if (bytes < 1024 * 1024) return `${(bytes / 1024).toFixed(1)} KB`;
  return `${(bytes / (1024 * 1024)).toFixed(2)} MB`;
}

export function normalizeRecipients(value) {
  if (value === null || value === undefined) return [];
  const list = Array.isArray(value) ? value : String(value).split(/[;,]/);
  return list.map((item) => String(item).trim()).filter(Boolean);
}

// Throw if an address is malformed or (when an allowlist is set) outside it.
export function assertAddressAllowed(address) {
  if (!EMAIL_RE.test(address)) {
    throw new Error(`Invalid email address: ${address}`);
  }
  if (SEND_ALLOWED_DOMAINS.length > 0) {
    const domain = address.split("@").pop().toLowerCase();
    const allowed = SEND_ALLOWED_DOMAINS.some(
      (allowedDomain) => domain === allowedDomain || domain.endsWith(`.${allowedDomain}`)
    );
    if (!allowed) {
      throw new Error(
        `Recipient ${address} is not in the allowed domains (${SEND_ALLOWED_DOMAINS.join(", ")}).`
      );
    }
  }
}

export function validateRecipients({ to, cc, bcc }) {
  const groups = {
    to: normalizeRecipients(to),
    cc: normalizeRecipients(cc),
    bcc: normalizeRecipients(bcc),
  };
  const all = [...groups.to, ...groups.cc, ...groups.bcc];

  if (groups.to.length === 0) {
    throw new Error("At least one recipient is required in 'to'.");
  }
  if (new Set(all.map((a) => a.toLowerCase())).size > MAX_SEND_RECIPIENTS) {
    throw new Error(`Too many recipients (max ${MAX_SEND_RECIPIENTS} across to/cc/bcc).`);
  }
  for (const address of all) {
    assertAddressAllowed(address);
  }

  const toGraph = (list) => list.map((address) => ({ emailAddress: { address } }));
  return {
    groups,
    toRecipients: toGraph(groups.to),
    ccRecipients: toGraph(groups.cc),
    bccRecipients: toGraph(groups.bcc),
  };
}

export function isInsideDir(child, dir) {
  const norm = (p) => (process.platform === "win32" ? p.toLowerCase() : p);
  const c = norm(child);
  const d = norm(dir);
  return c.startsWith(d.endsWith(path.sep) ? d : d + path.sep);
}

// Resolve a local path to a real file inside one of SEND_ALLOWED_DIRS.
// Symlinks/junctions are resolved first, so they cannot escape the allowlist.
export function resolveSendableLocalFile(filePath) {
  if (SEND_ALLOWED_DIRS.length === 0) {
    throw new Error(
      "Local file attachments are disabled. Choose an attachments folder in the extension settings (M365_SEND_ALLOWED_DIRS) to enable them."
    );
  }
  if (typeof filePath !== "string" || !filePath.trim() || filePath.includes("\u0000")) {
    throw new Error("Each local attachment must be a non-empty file path.");
  }

  const trimmed = filePath.trim();
  const candidates = path.isAbsolute(trimmed)
    ? [trimmed]
    : SEND_ALLOWED_DIRS.map((dir) => path.join(dir, trimmed));

  for (const candidate of candidates) {
    let real;
    try {
      real = fs.realpathSync(candidate);
    } catch {
      continue;
    }
    if (!SEND_ALLOWED_DIRS.some((dir) => isInsideDir(real, dir))) {
      throw new Error(`File is outside the allowed attachments folder: ${trimmed}`);
    }
    const stat = fs.statSync(real);
    if (!stat.isFile()) {
      throw new Error(`Not a regular file: ${trimmed}`);
    }
    return { path: real, name: path.basename(real), size: stat.size };
  }

  throw new Error(`File not found in the allowed attachments folder: ${trimmed}`);
}

export function listSendableFiles(nameContains = "") {
  const needle = String(nameContains || "").toLowerCase();
  const out = [];

  const walk = (root, dir, depth) => {
    if (out.length >= MAX_LIST_SENDABLE_FILES || depth > MAX_LIST_SENDABLE_DEPTH) return;
    let entries;
    try {
      entries = fs.readdirSync(dir, { withFileTypes: true });
    } catch {
      return;
    }
    for (const entry of entries) {
      if (out.length >= MAX_LIST_SENDABLE_FILES) return;
      if (entry.name.startsWith(".") || entry.isSymbolicLink()) continue;
      const full = path.join(dir, entry.name);
      if (entry.isDirectory()) {
        walk(root, full, depth + 1);
      } else if (entry.isFile() && (!needle || entry.name.toLowerCase().includes(needle))) {
        try {
          const stat = fs.statSync(full);
          out.push({
            name: entry.name,
            relativePath: path.relative(root, full).replace(/\\/g, "/"),
            path: full,
            size: stat.size,
            modified: stat.mtime.toISOString(),
          });
        } catch {
          // Skip unreadable entries.
        }
      }
    }
  };

  for (const root of SEND_ALLOWED_DIRS) walk(root, root, 0);
  out.sort((a, b) => b.modified.localeCompare(a.modified));
  return out;
}

// Load every requested attachment into memory, enforcing count/size limits.
export async function loadOutboundAttachments(localFiles = [], emailAttachments = []) {
  const total = localFiles.length + emailAttachments.length;
  if (total > MAX_SEND_ATTACHMENTS) {
    throw new Error(`Too many attachments (max ${MAX_SEND_ATTACHMENTS}).`);
  }

  const items = [];
  let totalBytes = 0;
  const addBytes = (size, name) => {
    totalBytes += size;
    if (totalBytes > MAX_SEND_TOTAL_BYTES) {
      throw new Error(
        `Attachments exceed the ${formatBytes(MAX_SEND_TOTAL_BYTES)} total limit (at "${name}").`
      );
    }
  };

  for (const filePath of localFiles) {
    const file = resolveSendableLocalFile(filePath);
    addBytes(file.size, file.name);
    items.push({
      name: file.name,
      source: `local: ${file.path}`,
      contentType: MIME_BY_EXT[getExtension(file.name)] || "application/octet-stream",
      bytes: fs.readFileSync(file.path),
    });
  }

  for (const ref of emailAttachments) {
    const base = `${GRAPH_BASE}/me/messages/${encodeURIComponent(ref.messageId)}/attachments/${encodeURIComponent(ref.attachmentId)}`;
    const meta = await graphGetJson(base);
    const kind = (meta["@odata.type"] || "").split(".").pop();
    if (kind !== "fileAttachment") {
      throw new Error(
        `Attachment "${meta.name || ref.attachmentId}" is a ${kind || "non-file attachment"} and cannot be re-attached.`
      );
    }
    addBytes(meta.size || 0, meta.name);
    const bytes = await graphGetBytes(`${base}/$value`);
    items.push({
      name: meta.name || "attachment",
      source: `email attachment from message ${ref.messageId}`,
      contentType: meta.contentType || "application/octet-stream",
      bytes,
    });
  }

  return { items, totalBytes };
}

// Add already-loaded attachments to an existing draft: small ones inline via a
// single POST, large ones through an upload session. Used by reply/forward,
// where the draft already exists before attachments are added.
export async function attachToDraft(draftId, items) {
  for (const item of items) {
    if (item.bytes.length <= MAX_INLINE_ATTACHMENT_BYTES) {
      await graphRequest("POST", `${GRAPH_BASE}/me/messages/${encodeURIComponent(draftId)}/attachments`, {
        "@odata.type": "#microsoft.graph.fileAttachment",
        name: item.name,
        contentType: item.contentType,
        contentBytes: item.bytes.toString("base64"),
      });
    } else {
      await uploadLargeAttachment(draftId, item);
    }
  }
}

export async function uploadLargeAttachment(messageId, item) {
  const session = await graphRequest(
    "POST",
    `${GRAPH_BASE}/me/messages/${encodeURIComponent(messageId)}/attachments/createUploadSession`,
    {
      AttachmentItem: {
        attachmentType: "file",
        name: item.name,
        size: item.bytes.length,
        contentType: item.contentType,
      },
    }
  );

  const uploadUrl = session?.uploadUrl;
  if (!uploadUrl || !/^https:\/\//i.test(uploadUrl)) {
    throw new Error(`Graph did not return a valid upload URL for "${item.name}".`);
  }

  // The upload URL is pre-authorized: do NOT send the bearer token to it.
  for (let start = 0; start < item.bytes.length; start += UPLOAD_CHUNK_BYTES) {
    const end = Math.min(start + UPLOAD_CHUNK_BYTES, item.bytes.length);
    const chunk = item.bytes.subarray(start, end);
    const res = await fetch(uploadUrl, {
      method: "PUT",
      headers: {
        "Content-Type": "application/octet-stream",
        "Content-Length": String(chunk.length),
        "Content-Range": `bytes ${start}-${end - 1}/${item.bytes.length}`,
      },
      body: chunk,
    });
    if (!res.ok) {
      const text = await res.text();
      log("upload chunk failed", { name: item.name, status: res.status, text });
      throw new Error(`Uploading "${item.name}" failed (HTTP ${res.status}).`);
    }
  }
}

// Create a draft (always), attach files, then optionally send it.
// Hidden MAPI named property stamped on every draft this tool creates. It lets
// the server prove a draft is its own before deleting it when a revision
// replaces it, so it can never be used to delete arbitrary mail.
export const DRAFT_MARKER_PROPERTY_ID =
  "String {6f1c2b9e-3d4a-4f7b-9c55-2e8a1d0b7c41} Name M365McpCreatedDraft";
export const DRAFT_MARKER_VALUE = "1";

// Delete a previous draft that a revision replaces. Only deletes when the
// message still exists, is still a draft, and carries this tool's marker.
// Deleted drafts go to Deleted Items (recoverable), not hard-deleted.
export async function deleteReplacedDraft(draftId) {
  const filter = encodeURIComponent(`id eq '${DRAFT_MARKER_PROPERTY_ID}'`);
  let message;
  try {
    message = await graphGetJson(
      `${GRAPH_BASE}/me/messages/${encodeURIComponent(draftId)}?$select=id,isDraft,subject&$expand=singleValueExtendedProperties($filter=${filter})`
    );
  } catch (err) {
    if (/Graph API 404/.test(String(err?.message))) {
      return { deleted: false, reason: "the previous draft no longer exists (already sent or deleted)" };
    }
    throw err;
  }

  if (!message?.isDraft) {
    return { deleted: false, reason: "the previous message is not a draft anymore (it may have been sent), left untouched" };
  }

  const marked = (message.singleValueExtendedProperties || []).some(
    (prop) =>
      String(prop?.id || "").toLowerCase() === DRAFT_MARKER_PROPERTY_ID.toLowerCase() &&
      prop?.value === DRAFT_MARKER_VALUE
  );
  if (!marked) {
    return { deleted: false, reason: "the previous draft was not created by this tool, left untouched" };
  }

  await graphRequest("DELETE", `${GRAPH_BASE}/me/messages/${encodeURIComponent(draftId)}`);
  return { deleted: true, subject: message.subject || "" };
}

export async function composeOutlookMessage({
  recipients,
  subject,
  body,
  bodyType,
  attachments,
  sendNow,
}) {
  const inline = [];
  const large = [];
  let inlineBytes = 0;
  for (const item of attachments) {
    if (inlineBytes + item.bytes.length <= MAX_INLINE_ATTACHMENT_BYTES) {
      inlineBytes += item.bytes.length;
      inline.push(item);
    } else {
      large.push(item);
    }
  }

  const draft = await graphRequest("POST", `${GRAPH_BASE}/me/messages`, {
    subject: subject || "",
    body: { contentType: bodyType === "HTML" ? "HTML" : "Text", content: body || "" },
    toRecipients: recipients.toRecipients,
    ccRecipients: recipients.ccRecipients,
    bccRecipients: recipients.bccRecipients,
    singleValueExtendedProperties: [{ id: DRAFT_MARKER_PROPERTY_ID, value: DRAFT_MARKER_VALUE }],
    attachments: inline.map((item) => ({
      "@odata.type": "#microsoft.graph.fileAttachment",
      name: item.name,
      contentType: item.contentType,
      contentBytes: item.bytes.toString("base64"),
    })),
  });

  if (!draft?.id) {
    throw new Error("Graph did not return the created draft.");
  }

  try {
    for (const item of large) {
      await uploadLargeAttachment(draft.id, item);
    }
  } catch (err) {
    // Don't leave a half-built draft (missing attachments) behind.
    await graphRequest("DELETE", `${GRAPH_BASE}/me/messages/${encodeURIComponent(draft.id)}`).catch(
      (deleteErr) => log("draft cleanup failed", deleteErr?.message || String(deleteErr))
    );
    throw err;
  }

  if (sendNow) {
    try {
      await graphRequest("POST", `${GRAPH_BASE}/me/messages/${encodeURIComponent(draft.id)}/send`);
    } catch (err) {
      // Keep the draft: the send may or may not have gone out, so let the user check.
      throw new Error(
        `The draft was created but sending failed: ${err?.message || err}. The draft is kept in your Drafts folder${draft.webLink ? ` (${draft.webLink})` : ""}; check Sent Items before retrying.`
      );
    }
  }

  return { draftId: draft.id, webLink: draft.webLink || null };
}

// Forward an existing message via Graph's createForward (keeps the original
// body, inline images and attachments intact). We always create a forward
// DRAFT first (marked as ours), optionally patch recipients/comment, then send.
export async function forwardOutlookMessage({ messageId, recipients, comment, sendNow }) {
  // 1. Build the forward draft from the original message. Graph copies the
  //    original body + attachments; recipients/comment can be set here.
  const draft = await graphRequest(
    "POST",
    `${GRAPH_BASE}/me/messages/${encodeURIComponent(messageId)}/createForward`,
    {
      toRecipients: recipients.toRecipients,
      ccRecipients: recipients.ccRecipients,
      bccRecipients: recipients.bccRecipients,
      ...(comment ? { comment } : {}),
    }
  );

  if (!draft?.id) {
    throw new Error("Graph did not return the forward draft.");
  }

  // 2. Mark it as created by this tool so our own draft-replace logic can
  //    recognise/clean it up later, matching composeOutlookMessage.
  try {
    await graphRequest("PATCH", `${GRAPH_BASE}/me/messages/${encodeURIComponent(draft.id)}`, {
      singleValueExtendedProperties: [{ id: DRAFT_MARKER_PROPERTY_ID, value: DRAFT_MARKER_VALUE }],
    });
  } catch (err) {
    log("forward draft marking failed", err?.message || String(err));
  }

  if (sendNow) {
    try {
      await graphRequest("POST", `${GRAPH_BASE}/me/messages/${encodeURIComponent(draft.id)}/send`);
    } catch (err) {
      throw new Error(
        `The forward draft was created but sending failed: ${err?.message || err}. The draft is kept in your Drafts folder${draft.webLink ? ` (${draft.webLink})` : ""}; check Sent Items before retrying.`
      );
    }
  }

  return { draftId: draft.id, webLink: draft.webLink || null };
}

// Reply (or reply-all) to an existing message via Graph's createReply(All),
// which keeps the quoted original and pre-fills recipients. We create a reply
// DRAFT first, mark it as ours, verify the pre-filled recipients against the
// allowlist, then optionally send.
export async function replyOutlookMessage({ messageId, comment, replyAll, attachments = [], sendNow }) {
  const action = replyAll ? "createReplyAll" : "createReply";
  const draft = await graphRequest(
    "POST",
    `${GRAPH_BASE}/me/messages/${encodeURIComponent(messageId)}/${action}`,
    comment ? { comment } : undefined
  );

  if (!draft?.id) {
    throw new Error("Graph did not return the reply draft.");
  }

  const cleanup = async (reason) => {
    await graphRequest("DELETE", `${GRAPH_BASE}/me/messages/${encodeURIComponent(draft.id)}`).catch(
      (err) => log("reply draft cleanup failed", err?.message || String(err))
    );
    throw new Error(reason);
  };

  // The recipients were filled in by Graph from the original message. Verify
  // them against the allowlist before anything can be sent, and surface who
  // the reply will go to (reply-all can include unexpected addresses).
  let recipientsTo = [];
  let recipientsCc = [];
  try {
    const filled = await graphGetJson(
      `${GRAPH_BASE}/me/messages/${encodeURIComponent(draft.id)}?$select=toRecipients,ccRecipients`
    );
    recipientsTo = (filled.toRecipients || []).map((r) => r.emailAddress?.address).filter(Boolean);
    recipientsCc = (filled.ccRecipients || []).map((r) => r.emailAddress?.address).filter(Boolean);
  } catch (err) {
    await cleanup(`Could not read the reply recipients: ${err?.message || err}`);
  }

  const everyone = [...recipientsTo, ...recipientsCc];
  if (everyone.length > MAX_SEND_RECIPIENTS) {
    await cleanup(`The reply would go to ${everyone.length} recipients (max ${MAX_SEND_RECIPIENTS}).`);
  }
  try {
    for (const address of everyone) assertAddressAllowed(address);
  } catch (err) {
    await cleanup(String(err?.message || err));
  }

  if (attachments.length) {
    try {
      await attachToDraft(draft.id, attachments);
    } catch (err) {
      await cleanup(`Failed to attach files to the reply: ${err?.message || err}`);
    }
  }

  if (sendNow) {
    try {
      await graphRequest("POST", `${GRAPH_BASE}/me/messages/${encodeURIComponent(draft.id)}/send`);
    } catch (err) {
      throw new Error(
        `The reply draft was created but sending failed: ${err?.message || err}. The draft is kept in your Drafts folder${draft.webLink ? ` (${draft.webLink})` : ""}; check Sent Items before retrying.`
      );
    }
  } else {
    // Mark drafts so the draft-replace logic can recognise them later.
    await graphRequest("PATCH", `${GRAPH_BASE}/me/messages/${encodeURIComponent(draft.id)}`, {
      singleValueExtendedProperties: [{ id: DRAFT_MARKER_PROPERTY_ID, value: DRAFT_MARKER_VALUE }],
    }).catch((err) => log("reply draft marking failed", err?.message || String(err)));
  }

  return { draftId: draft.id, webLink: draft.webLink || null, to: recipientsTo, cc: recipientsCc };
}

export function sendConfigSummary() {
  return [
    `Mode: ${SEND_NOW_ALLOWED ? "drafts or send immediately (send_now=true)" : "drafts only (sending immediately is disabled)"}`,
    `Recipient domains: ${SEND_ALLOWED_DOMAINS.length ? SEND_ALLOWED_DOMAINS.join(", ") : "any"}`,
    `Local attachments folder: ${SEND_ALLOWED_DIRS.length ? SEND_ALLOWED_DIRS.join(", ") : "(disabled)"}`,
    `Limits: ${MAX_SEND_RECIPIENTS} recipients, ${MAX_SEND_ATTACHMENTS} attachments, ${formatBytes(MAX_SEND_TOTAL_BYTES)} total`,
  ].join("\n");
}
