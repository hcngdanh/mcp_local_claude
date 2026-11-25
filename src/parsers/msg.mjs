import { MAX_MSG_ATTACHMENTS, MAX_NESTED_DEPTH } from "../config.mjs";
import { appendSection, formatMetadataBlock, normalizeExtractedText } from "./common.mjs";
import { parseAttachmentPayload } from "./index.mjs";

export let msgReaderPromise = null;

export async function loadMsgReader() {
  if (!msgReaderPromise) {
    msgReaderPromise = import("@kenjiuno/msgreader").then((module) => {
      const MsgReader =
        module.default?.default ||
        module.default ||
        module.MsgReader ||
        module;

      if (typeof MsgReader !== "function") {
        throw new Error(".msg parser dependency loaded but MsgReader export is missing.");
      }

      return MsgReader;
    });
  }

  return msgReaderPromise;
}

export async function parseMsg(raw, filename, context, depth) {
  const MsgReader = await loadMsgReader();
  const msg = new MsgReader(raw);
  const info = msg.getFileData();
  const parts = [];

  appendSection(
    parts,
    "Message Summary",
    formatMetadataBlock([
      info?.subject ? `Subject: ${info.subject}` : "",
      info?.senderName ? `Sender Name: ${info.senderName}` : "",
      info?.senderEmail ? `Sender Email: ${info.senderEmail}` : "",
      info?.messageDeliveryTime ? `Delivered: ${info.messageDeliveryTime}` : "",
      Array.isArray(info?.recipients) && info.recipients.length
        ? `Recipients: ${info.recipients
          .map((item) => item.email || item.name)
          .filter(Boolean)
          .join(", ")}`
        : "",
    ])
  );

  const body = normalizeExtractedText(info?.body || info?.bodyHTML || "");
  appendSection(parts, "Message Body", body || "(No readable message body was detected.)");

  const attachments = Array.isArray(info?.attachments) ? info.attachments : [];
  if (attachments.length) {
    appendSection(
      parts,
      "Embedded Attachments",
      attachments
        .slice(0, MAX_MSG_ATTACHMENTS)
        .map(
          (attachment, index) =>
            `${index + 1}. ${attachment.fileName || attachment.fileNameShort || "unnamed attachment"}`
        )
        .join("\n")
    );

    for (const attachment of attachments.slice(0, MAX_MSG_ATTACHMENTS)) {
      const nested = msg.getAttachment(attachment);
      const nestedName =
        nested?.fileName ||
        attachment.fileName ||
        attachment.fileNameShort ||
        "embedded-attachment.bin";
      const nestedRaw = Buffer.from(nested?.content || []);

      if (!nestedRaw.length) {
        appendSection(parts, `Embedded Attachment: ${nestedName}`, "(Attachment payload was empty.)");
        continue;
      }

      if (depth >= MAX_NESTED_DEPTH) {
        appendSection(
          parts,
          `Embedded Attachment: ${nestedName}`,
          "(Nested attachment parsing skipped because the recursion limit was reached.)"
        );
        continue;
      }

      const nestedParsed = await parseAttachmentPayload(
        {
          raw: nestedRaw,
          filename: nestedName,
          contentType: attachment.mimeType || "",
          kind: "attachment",
        },
        context,
        depth + 1
      );

      appendSection(parts, `Embedded Attachment: ${nestedName}`, nestedParsed.text);
    }

    if (attachments.length > MAX_MSG_ATTACHMENTS) {
      appendSection(
        parts,
        "Embedded Attachments Limit",
        `Only the first ${MAX_MSG_ATTACHMENTS} embedded attachments were parsed out of ${attachments.length}.`
      );
    }
  }

  const text = normalizeExtractedText(parts.join("\n\n"));
  return {
    parser: "outlook_msg",
    text:
      text ||
      `[Outlook .msg attachment]\n\nFilename: ${filename}\n\n(No readable message content was extracted.)`,
    details: {
      embeddedAttachmentCount: attachments.length,
    },
  };
}
