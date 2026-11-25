import { z } from "zod";
import { graphGetJson, messageCollectionUrl } from "../../graph.mjs";
import { registerAliases } from "../register.mjs";
import { formatSender, matchesNeedle, normalizeNeedle } from "../../util.mjs";

export function registerListRecentMessages(server) {
  const listMessagesHandler = async ({
    mailbox,
    folder,
    top,
    onlyWithAttachments,
    subjectContains,
    fromContains,
  }) => {
    try {
      const requestedTop = top;
      const fetchLimit = Math.min(
        Math.max(
          requestedTop * (onlyWithAttachments || subjectContains || fromContains ? 5 : 2),
          25
        ),
        100
      );
      const collectionUrl = messageCollectionUrl(mailbox, folder);
      const url =
        `${collectionUrl}` +
        `?$select=id,subject,receivedDateTime,hasAttachments,from,bodyPreview` +
        `&$orderby=receivedDateTime desc` +
        `&$top=${fetchLimit}`;

      const data = await graphGetJson(url);
      const subjectNeedle = normalizeNeedle(subjectContains);
      const fromNeedle = normalizeNeedle(fromContains);

      let messages = (data.value || []).map((message) => ({
        id: message.id,
        subject: message.subject,
        receivedDateTime: message.receivedDateTime,
        hasAttachments: message.hasAttachments,
        fromName: message.from?.emailAddress?.name || "",
        fromAddress: message.from?.emailAddress?.address || "",
        bodyPreview: message.bodyPreview || "",
      }));

      if (onlyWithAttachments) {
        messages = messages.filter((message) => message.hasAttachments);
      }

      messages = messages.filter(
        (message) =>
          matchesNeedle(message.subject, subjectNeedle) &&
          (matchesNeedle(message.fromName, fromNeedle) ||
            matchesNeedle(message.fromAddress, fromNeedle))
      );

      messages = messages.slice(0, requestedTop);

      const text =
        messages.length > 0
          ? messages
            .map(
              (message, index) =>
                `${index + 1}. ${message.subject || "(no subject)"}\nFrom: ${formatSender(message.fromName, message.fromAddress)}\nReceived: ${message.receivedDateTime}\nHas Attachments: ${message.hasAttachments}${message.bodyPreview ? `\nPreview: ${message.bodyPreview.slice(0, 200)}` : ""}\nMessage ID: ${message.id}`
            )
            .join("\n\n")
          : [
            "No recent messages matched the current filter.",
            `Folder searched: ${folder || "inbox"}`,
            `Messages scanned: ${data.value?.length || 0}`,
            subjectNeedle ? `Subject filter: ${subjectContains}` : "",
            fromNeedle ? `Sender filter: ${fromContains}` : "",
            onlyWithAttachments ? "Attachment filter: hasAttachments=true" : "",
          ]
            .filter(Boolean)
            .join("\n");

      return {
        structuredContent: {
          mailbox,
          folder: folder || "inbox",
          requestedTop,
          scannedCount: data.value?.length || 0,
          returnedCount: messages.length,
          onlyWithAttachments,
          subjectContains: subjectContains || "",
          fromContains: fromContains || "",
          messages,
        },
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
    "list_recent_messages",
    {
      title: "List Recent Outlook Messages",
      description:
        "List recent Outlook emails from Microsoft 365. By default this searches the Inbox, prefers emails with attachments, and can filter by subject or sender name/address.",
      inputSchema: z.object({
        mailbox: z.string().default("me"),
        folder: z.string().default("inbox"),
        top: z.number().int().min(1).max(50).default(10),
        onlyWithAttachments: z.boolean().default(true),
        subjectContains: z.string().optional().default(""),
        fromContains: z.string().optional().default(""),
      }),
    },
    listMessagesHandler
  );
}
