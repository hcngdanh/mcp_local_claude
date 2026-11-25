import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { APP_NAME, SEND_ENABLED } from "./config.mjs";
import { registerAuthTools } from "./tools/auth/index.mjs";
import { registerReadTools } from "./tools/read/index.mjs";
import { registerSendTools } from "./tools/send/index.mjs";

export function createServer() {
  const server = new McpServer({
    name: APP_NAME,
    version: "0.3.5",
  });

  registerAuthTools(server);
  registerReadTools(server);

  // Outbound mail tools exist only when enabled in the extension settings.
  if (SEND_ENABLED) {
    registerSendTools(server);
  }

  return server;
}
