import { z } from "zod";
import { SEND_ALLOWED_DIRS } from "../../config.mjs";
import { formatBytes, listSendableFiles } from "../../mail/compose.mjs";
import { errorResult, registerAliases } from "../register.mjs";

export function registerListSendableFiles(server) {
  registerAliases(
    server,
    "list_sendable_files",
    {
      title: "List Files Available to Attach",
      description:
        "List local files inside the configured attachments folder that can be attached with send_outlook_email (newest first). Use this to find the exact file the user means before composing.",
      inputSchema: z.object({
        nameContains: z.string().optional().describe("Case-insensitive filename filter."),
      }),
    },
    async ({ nameContains }) => {
      if (SEND_ALLOWED_DIRS.length === 0) {
        return errorResult(
          "No attachments folder is configured, so local file attachments are disabled. Attachments from existing emails can still be used."
        );
      }
      const files = listSendableFiles(nameContains);
      const text = files.length
        ? files
          .map((f, i) => `${i + 1}. ${f.relativePath} (${formatBytes(f.size)}, modified ${f.modified})\n   Path: ${f.path}`)
          .join("\n")
        : "No matching files.";
      return {
        structuredContent: { folders: SEND_ALLOWED_DIRS, files },
        content: [{ type: "text", text }],
      };
    }
  );
}
