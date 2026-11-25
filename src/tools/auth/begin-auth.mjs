import { z } from "zod";
import { authPromise, authState, buildBeginAuthText, ensureConfigured, getCachedToken, getPublicAuthState, startDeviceCodeAuth } from "../../auth.mjs";
import { log } from "../../logger.mjs";
import { registerAliases } from "../register.mjs";

export function registerBeginAuth(server) {
  const beginAuthHandler = async () => {
    try {
      ensureConfigured();

      if (getCachedToken()) {
        return {
          structuredContent: { auth: getPublicAuthState() },
          content: [
            {
              type: "text",
              text: "Microsoft 365 is already logged in for this local MCP process.",
            },
          ],
        };
      }

      if (!authPromise && authState.status !== "pending") {
        startDeviceCodeAuth().catch((err) => {
          log("startDeviceCodeAuth background error", err?.message || String(err));
        });
      }

      await new Promise((resolve) => setTimeout(resolve, 1500));

      return {
        structuredContent: { auth: getPublicAuthState() },
        content: [
          {
            type: "text",
            text:
              authState.status === "pending"
                ? buildBeginAuthText()
                : `Current auth status: ${authState.status}${authState.error ? `\nError: ${authState.error}` : ""}`,
          },
        ],
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
    "begin_auth",
    {
      title: "Begin Microsoft 365 Auth",
      description:
        "Start Microsoft 365 device-code login for the local Claude Desktop MCP process.",
      inputSchema: z.object({}),
    },
    beginAuthHandler
  );
}
