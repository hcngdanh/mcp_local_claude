import { z } from "zod";
import { graphGetJson, mailboxPrefix } from "../../graph.mjs";
import { htmlToPlainText, normalizeExtractedText, truncateExtractedText } from "../../parsers/common.mjs";
import { registerAliases } from "../register.mjs";
import { formatSender } from "../../util.mjs";

export function registerReadEmail(server) {
  const readEmailHandler = async ({ messageId, mailbox }) => {
    try {
      const base = mailboxPrefix(mailbox);
      const url =
        `${base}/messages/${encodeURIComponent(messageId)}` +
        `?$select=id,subject,from,toRecipients,ccRecipients,receivedDateTime,sentDateTime,hasAttachments,bodyPreview,body,webLink`;
      const m = await graphGetJson(url);

      const fromStr = formatSender(
        m.from?.emailAddress?.name,
        m.from?.emailAddress?.address
      );
      const toStr = (m.toRecipients || [])
        .map((r) => r.emailAddress?.address)
        .filter(Boolean)
        .join(", ");
      const ccStr = (m.ccRecipients || [])
        .map((r) => r.emailAddress?.address)
        .filter(Boolean)
        .join(", ");

      const isHtml = String(m.body?.contentType || "").toLowerCase() === "html";
      const bodyText = isHtml
        ? htmlToPlainText(m.body?.content)
        : normalizeExtractedText(m.body?.content || "");
      const truncated = truncateExtractedText(bodyText);

      const headerLines = [
        `Subject: ${m.subject || "(no subject)"}`,
        `From: ${fromStr}`,
        toStr ? `To: ${toStr}` : "",
        ccStr ? `Cc: ${ccStr}` : "",
        `Received: ${m.receivedDateTime || ""}`,
        `Has attachments: ${m.hasAttachments ? "yes" : "no"}`,
      ].filter(Boolean);

      if (truncated.truncated) {
        headerLines.push(
          `Notice: body truncated to ${truncated.returnedLength.toLocaleString()} of ${truncated.totalLength.toLocaleString()} characters.`
        );
      }

      const textOut = `${headerLines.join("\n")}\n\n${truncated.text || "(empty body)"}`;

      return {
        structuredContent: {
          message: {
            id: m.id,
            subject: m.subject,
            from: fromStr,
            to: toStr,
            cc: ccStr,
            receivedDateTime: m.receivedDateTime,
            sentDateTime: m.sentDateTime,
            hasAttachments: !!m.hasAttachments,
            webLink: m.webLink,
            bodyPreview: m.bodyPreview || "",
            bodyTextLength: bodyText.length,
            truncated: truncated.truncated,
          },
        },
        content: [{ type: "text", text: textOut }],
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
    "read_email",
    {
      title: "Read Outlook Email Body",
      description:
        "Fetch and return the full text of a specific Outlook email: subject, sender, recipients, date, and the message body (HTML is converted to plain text). Use list_recent_messages first to get the messageId. This reads the email body itself, complementing read_email_attachment which reads attachment contents.",
      inputSchema: z.object({
        messageId: z.string(),
        mailbox: z.string().default("me"),
      }),
    },
    readEmailHandler
  );
}
