// Server status and Microsoft 365 sign-in tools. Always registered.
import { registerAuthStatus } from "./auth-status.mjs";
import { registerBeginAuth } from "./begin-auth.mjs";
import { registerHealthCheck } from "./health-check.mjs";

export function registerAuthTools(server) {
  registerHealthCheck(server);
  registerBeginAuth(server);
  registerAuthStatus(server);
}
