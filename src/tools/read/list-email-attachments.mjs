import { z } from "zod";
import { graphGetJson, mailboxPrefix } from "../../graph.mjs";
import { registerAliases } from "../register.mjs";

export function registerListEmailAttachments(server) {
  const listAttachmentsHandler = async ({ messageId, mailbox }) => {
    try {
      const base = mailboxPrefix(mailbox);
      const url = `${base}/messages/${encodeURIComponent(messageId)}/attachments`;
      const data = await graphGetJson(url);

      const attachments = (data.value || []).map((attachment) => ({
        id: attachment.id,
        name: attachment.name,
        contentType: attachment.contentType,
        size: attachment.size,
        isInline: attachment.isInline,
        kind: (attachment["@odata.type"] || "").split(".").pop() || "attachment",
      }));

      const text =
        attachments.length > 0
          ? [
            "These attachments are fetched directly from Microsoft Graph by attachment ID, so they are not limited by normal chat file-upload size caps. Large inline image previews may be downscaled before being returned to Claude.",
            attachments
              .map(
                (attachment, index) =>
                  `${index + 1}. ${attachment.name || "(unnamed attachment)"}\nAttachment ID: ${attachment.id}\nType: ${attachment.contentType || "application/octet-stream"}\nSize: ${attachment.size || 0} bytes\nInline: ${attachment.isInline ? "yes" : "no"}\nKind: ${attachment.kind}`
              )
              .join("\n\n"),
          ].join("\n\n")
          : "This email has no downloadable attachments.";

      return {
        structuredContent: { attachments },
        content: [{ type: "text", text }],
      };
    } catch (err) {
      return {
        isError: true,
        content: [{ type: "text", text: String(err?.message || err) }],
      };
    }
  };

  registerAliases(
    server,
    "list_email_attachments",
    {
      title: "List Email Attachments",
      description: "List attachments for a specific Outlook email.",
      inputSchema: z.object({
        messageId: z.string(),
        mailbox: z.string().default("me"),
      }),
    },
    listAttachmentsHandler
  );
}
