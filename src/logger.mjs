import fs from "node:fs";
import { DEBUG_LOG } from "./config.mjs";

export function log(...parts) {
  const line =
    `[${new Date().toISOString()}] ` +
    parts
      .map((part) => {
        if (typeof part === "string") return part;
        if (part instanceof Error) return part.stack || part.message;
        try {
          return JSON.stringify(part);
        } catch {
          return String(part);
        }
      })
      .join(" ") +
    "\n";

  fs.appendFileSync(DEBUG_LOG, line, "utf8");
}

process.on("uncaughtException", (err) => {
  log("uncaughtException", err);
  console.error(err);
});

process.on("unhandledRejection", (err) => {
  log("unhandledRejection", err);
  console.error(err);
});
