import { z } from "zod";
import { getPublicAuthState } from "../../auth.mjs";
import { APP_NAME, SCOPES, SEND_ALLOWED_DIRS, SEND_ALLOWED_DOMAINS, SEND_ENABLED, SEND_NOW_ALLOWED } from "../../config.mjs";
import { sendConfigSummary } from "../../mail/compose.mjs";
import { registerAliases } from "../register.mjs";

export function registerHealthCheck(server) {
  const healthHandler = async () => ({
    structuredContent: {
      ok: true,
      server: APP_NAME,
      transport: "stdio",
      time: new Date().toISOString(),
      auth: getPublicAuthState(),
      scopes: SCOPES,
      outbound: {
        composeEnabled: SEND_ENABLED,
        sendNowAllowed: SEND_NOW_ALLOWED,
        allowedDomains: SEND_ALLOWED_DOMAINS,
        attachmentFolders: SEND_ALLOWED_DIRS,
      },
    },
    content: [
      {
        type: "text",
        text:
          `${APP_NAME} is running locally over stdio.\nScopes: ${SCOPES.join(", ")}\n` +
          (SEND_ENABLED
            ? `Email composing: ENABLED\n${sendConfigSummary()}`
            : "Email composing: disabled (read-only)."),
      },
    ],
  });

  registerAliases(
    server,
    "health_check",
    {
      title: "Health Check",
      description:
        "Verify that the local Outlook attachment reader MCP server is running and report auth state.",
      inputSchema: z.object({}),
    },
    healthHandler
  );
}
