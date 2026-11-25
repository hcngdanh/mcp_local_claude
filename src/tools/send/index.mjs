// Outbound mail tools (Mail.ReadWrite / Mail.Send). Registered only when
// "Enable email composing" is turned on (M365_ENABLE_SEND=true).
import { registerForwardEmail } from "./forward-email.mjs";
import { registerListSendableFiles } from "./list-sendable-files.mjs";
import { registerReplyEmail } from "./reply-email.mjs";
import { registerSendOutlookEmail } from "./send-outlook-email.mjs";

export function registerSendTools(server) {
  registerListSendableFiles(server);
  registerSendOutlookEmail(server);
  registerReplyEmail(server);
  registerForwardEmail(server);
}
