import { z } from "zod";
import { graphGetBytes, graphGetJson, mailboxPrefix } from "../../graph.mjs";
import { log } from "../../logger.mjs";
import { createParseContext, getExtension, truncateExtractedText } from "../../parsers/common.mjs";
import { parseAttachmentPayload } from "../../parsers/index.mjs";
import { registerAliases } from "../register.mjs";
import { safeFilename } from "../../util.mjs";

export function registerReadEmailAttachment(server) {
  const readAttachmentHandler = async ({ messageId, attachmentId, mailbox }) => {
    try {
      const base = mailboxPrefix(mailbox);
      const metaUrl =
        `${base}/messages/${encodeURIComponent(messageId)}/attachments/${encodeURIComponent(attachmentId)}`;
      const meta = await graphGetJson(metaUrl);
      const kind = (meta["@odata.type"] || "").split(".").pop() || "attachment";

      if (kind === "referenceAttachment") {
        const output = {
          id: meta.id,
          name: meta.name,
          contentType: meta.contentType,
          size: meta.size,
          isInline: meta.isInline,
          kind,
        };

        return {
          structuredContent: { attachment: output },
          content: [
            {
              type: "text",
              text:
                "This is a referenceAttachment. Only metadata is available; content cannot be downloaded directly.\n\n" +
                JSON.stringify(output, null, 2),
            },
          ],
        };
      }

      const rawUrl =
        `${base}/messages/${encodeURIComponent(messageId)}/attachments/${encodeURIComponent(attachmentId)}/$value`;
      const raw = await graphGetBytes(rawUrl);
      const filename = safeFilename(meta.name || attachmentId);
      const ext = getExtension(filename);
      const contentType = (meta.contentType || "application/octet-stream").toLowerCase();
      const parseContext = createParseContext();
      let parsed;

      try {
        parsed = await parseAttachmentPayload(
          {
            raw,
            filename,
            contentType,
            kind,
          },
          parseContext
        );
      } catch (parseError) {
        log("File parsing error", parseError);
        parsed = {
          parser: "parse_error",
          text: `(Warning: An error occurred while parsing this file: ${parseError.message})`,
          details: {},
        };
      }

      const truncated = truncateExtractedText(parsed.text);
      const previewParts = [
        "[Local MCP backend parsed the attachment]",
        `Filename: ${filename}`,
        `Parser: ${parsed.parser}`,
      ];

      if (parseContext.imageBlocks.length > 0) {
        previewParts.push(`Visual blocks attached: ${parseContext.imageBlocks.length}`);
      }

      if (parseContext.imageNotes.length > 0) {
        previewParts.push(
          `Image handling notes:\n${parseContext.imageNotes.map((note) => `- ${note}`).join("\n")}`
        );
      }

      if (truncated.truncated) {
        previewParts.push(
          `Notice: returned ${truncated.returnedLength.toLocaleString()} of ${truncated.totalLength.toLocaleString()} characters; omitted ${truncated.omittedLength.toLocaleString()} characters.`
        );
      }

      previewParts.push(truncated.text);

      return {
        structuredContent: {
          attachment: {
            id: meta.id,
            name: meta.name,
            contentType: meta.contentType,
            size: meta.size,
            isInline: meta.isInline,
            kind,
            ext,
            parser: parsed.parser,
            imageBlockCount: parseContext.imageBlocks.length,
            extractedTextLength: truncated.totalLength,
            returnedTextLength: truncated.returnedLength,
            truncated: truncated.truncated,
            omittedCharacters: truncated.omittedLength,
            ...(parsed.details || {}),
          },
        },
        content: [{ type: "text", text: previewParts.join("\n\n") }, ...parseContext.imageBlocks],
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
    "read_email_attachment",
    {
      title: "Read Email Attachment",
      description:
        "Download an Outlook attachment directly from Microsoft Graph and parse it locally. Supports PDF, OCR-scanned PDF, Word, PowerPoint, Excel, images, archives, MSG, and plain text. Large image previews are automatically downscaled to fit MCP payload limits.",
      inputSchema: z.object({
        messageId: z.string(),
        attachmentId: z.string(),
        mailbox: z.string().default("me"),
      }),
    },
    readAttachmentHandler
  );
}
