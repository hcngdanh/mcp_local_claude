#!/usr/bin/env node
import path from "node:path";
import { pathToFileURL } from "node:url";
import { StdioServerTransport } from "@modelcontextprotocol/sdk/server/stdio.js";
import { APP_NAME, CLIENT_ID } from "./src/config.mjs";
import { log } from "./src/logger.mjs";
import { restoreAccountFromCache } from "./src/auth.mjs";
import { createServer } from "./src/server.mjs";

let activeServer = null;

let activeTransport = null;

async function shutdown(signal) {
  log(`shutdown requested: ${signal}`);

  try {
    if (activeTransport) {
      await activeTransport.close();
    }
  } catch (err) {
    log("transport close failed", err);
  }

  try {
    if (activeServer) {
      await activeServer.close();
    }
  } catch (err) {
    log("server close failed", err);
  }
}

async function main() {
  activeServer = createServer();
  activeTransport = new StdioServerTransport();

  // Restore a previously-authenticated account from the persistent token cache
  // so the server is ready to silently refresh tokens after a restart, with no
  // interactive begin_auth needed (key for unattended scheduled runs).
  if (CLIENT_ID) {
    await restoreAccountFromCache().catch((err) =>
      log("startup restoreAccountFromCache failed", err?.message || String(err))
    );
  }

  process.on("SIGINT", () => {
    shutdown("SIGINT")
      .catch((err) => log("SIGINT shutdown failed", err))
      .finally(() => process.exit(0));
  });

  process.on("SIGTERM", () => {
    shutdown("SIGTERM")
      .catch((err) => log("SIGTERM shutdown failed", err))
      .finally(() => process.exit(0));
  });

  await activeServer.connect(activeTransport);
  log(`${APP_NAME} started over stdio`);
}

function isDirectExecution() {
  const entry = process.argv[1];
  if (!entry) return false;
  return import.meta.url === pathToFileURL(path.resolve(entry)).href;
}

export { createServer, main };

if (isDirectExecution()) {
  main().catch((err) => {
    log("main fatal error", err);
    console.error(err);
    process.exit(1);
  });
}
