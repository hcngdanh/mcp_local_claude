import { z } from "zod";
import { getPublicAuthState } from "../../auth.mjs";
import { registerAliases } from "../register.mjs";

export function registerAuthStatus(server) {
  const authStatusHandler = async () => ({
    structuredContent: { auth: getPublicAuthState() },
    content: [
      {
        type: "text",
        text: JSON.stringify(getPublicAuthState(), null, 2),
      },
    ],
  });

  registerAliases(
    server,
    "auth_status",
    {
      title: "Microsoft 365 Auth Status",
      description:
        "Check whether Microsoft 365 login for this local MCP process has completed.",
      inputSchema: z.object({}),
    },
    authStatusHandler
  );
}
