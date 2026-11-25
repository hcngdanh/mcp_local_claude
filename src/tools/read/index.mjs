// Read-only mail tools (Mail.Read). Always registered.
import { registerCheckEmailSecurity } from "./check-email-security.mjs";
import { registerListEmailAttachments } from "./list-email-attachments.mjs";
import { registerListRecentMessages } from "./list-recent-messages.mjs";
import { registerReadEmailAttachment } from "./read-email-attachment.mjs";
import { registerReadEmail } from "./read-email.mjs";

export function registerReadTools(server) {
  registerListRecentMessages(server);
  registerListEmailAttachments(server);
  registerReadEmailAttachment(server);
  registerReadEmail(server);
  registerCheckEmailSecurity(server);
}
