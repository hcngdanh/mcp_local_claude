// Pure analysis helpers for check_email_security. No network access here:
// the tool fetches the message from Graph and passes the raw fields in.
//
// The result is a set of *signals* plus a simple, explainable risk level.
// It is meant to support the model's judgment, not to replace Microsoft
// Defender or the user's own verification.

import { decodeXmlEntities, getExtension } from "../parsers/common.mjs";

const MAX_LINKS = 50;

const URL_SHORTENERS = new Set([
  "bit.ly", "tinyurl.com", "t.co", "goo.gl", "ow.ly", "is.gd", "buff.ly", "cutt.ly",
  "rebrand.ly", "shorturl.at", "rb.gy", "tiny.cc", "s.id", "v.gd", "t.ly", "lnkd.in",
]);

// Attachment types commonly used to deliver malware or credential-phishing pages.
const DANGEROUS_EXTENSIONS = new Set([
  "exe", "scr", "com", "pif", "bat", "cmd", "msi", "msix", "appx", "js", "jse", "vbs", "vbe",
  "wsf", "wsh", "ps1", "hta", "cpl", "jar", "lnk", "iso", "img", "vhd", "vhdx", "reg", "chm",
  "html", "htm", "shtml", "svg",
]);
const MACRO_EXTENSIONS = new Set(["docm", "dotm", "xlsm", "xltm", "xlam", "pptm", "potm", "ppsm"]);

// Second-level labels that sit under a 2-letter country TLD (example.com.vn).
const SECOND_LEVEL = new Set(["com", "net", "org", "gov", "edu", "co", "ac", "or", "ne", "go", "info", "biz"]);

export function hostOf(url) {
  try {
    return new URL(url).hostname.toLowerCase().replace(/\.$/, "");
  } catch {
    return "";
  }
}

// Approximate registrable domain ("mail.google.com" -> "google.com",
// "ib.vietcombank.com.vn" -> "vietcombank.com.vn").
export function baseDomain(host) {
  const h = String(host || "").toLowerCase().replace(/\.$/, "");
  if (!h || /^\d{1,3}(\.\d{1,3}){3}$/.test(h) || h.includes(":")) return h;
  const parts = h.split(".");
  if (parts.length >= 3 && parts.at(-1).length === 2 && SECOND_LEVEL.has(parts.at(-2))) {
    return parts.slice(-3).join(".");
  }
  return parts.slice(-2).join(".");
}

export function domainOfAddress(address) {
  const m = String(address || "").trim().toLowerCase().match(/@([^@\s>]+)$/);
  return m ? m[1].replace(/\.$/, "") : "";
}

// Outlook "Safe Links" wrap the real URL; unwrap so we judge the destination.
export function unwrapSafeLink(url) {
  try {
    const u = new URL(url);
    if (u.hostname.toLowerCase().endsWith("safelinks.protection.outlook.com")) {
      const inner = u.searchParams.get("url");
      if (inner) return inner;
    }
  } catch {
    // not a URL
  }
  return url;
}

function cleanText(html) {
  return decodeXmlEntities(String(html || "").replace(/<[^>]+>/g, " "))
    .replace(/\s+/g, " ")
    .trim();
}

// Domain that the visible link text claims to point to, if any.
function claimedDomain(text) {
  const m = String(text || "").toLowerCase().match(/(?:https?:\/\/)?((?:[a-z0-9-]+\.)+[a-z]{2,})(?:[/:?#]|\s|$)/);
  return m ? m[1] : "";
}

export function extractLinks(body, contentType) {
  const links = [];
  const content = String(body || "");

  if (String(contentType || "").toLowerCase() === "html") {
    const re = /<a\b[^>]*?\bhref\s*=\s*(?:"([^"]*)"|'([^']*)'|([^\s>]+))[^>]*>([\s\S]*?)<\/a>/gi;
    let m;
    while ((m = re.exec(content)) && links.length < MAX_LINKS) {
      const href = decodeXmlEntities((m[1] ?? m[2] ?? m[3] ?? "").trim());
      if (!href || href.startsWith("#") || /^mailto:/i.test(href) || /^tel:/i.test(href)) continue;
      links.push({ href, text: cleanText(m[4]).slice(0, 200) });
    }
  } else {
    const re = /\bhttps?:\/\/[^\s<>"')\]]+/gi;
    let m;
    while ((m = re.exec(content)) && links.length < MAX_LINKS) {
      links.push({ href: m[0], text: "" });
    }
  }

  return links.map(analyzeLink);
}

export function analyzeLink({ href, text }) {
  const flags = [];
  const scheme = (String(href).match(/^([a-z][a-z0-9+.-]*):/i)?.[1] || "").toLowerCase();

  if (scheme === "javascript" || scheme === "data" || scheme === "vbscript") {
    return { href, text, host: "", flags: [{ level: "high", reason: `${scheme}: link` }] };
  }

  const target = unwrapSafeLink(href);
  const host = hostOf(target);
  if (!host) return { href, text, host: "", flags };

  if (target !== href) flags.push({ level: "info", reason: "wrapped by Microsoft Safe Links" });
  if (/^\d{1,3}(\.\d{1,3}){3}$/.test(host)) flags.push({ level: "high", reason: "link goes to a raw IP address" });
  if (host.split(".").some((label) => label.startsWith("xn--"))) {
    flags.push({ level: "high", reason: "punycode domain (possible look-alike characters)" });
  }
  if (URL_SHORTENERS.has(host)) flags.push({ level: "medium", reason: "URL shortener hides the real destination" });
  if (/^http:$/i.test(safeProtocol(target))) flags.push({ level: "low", reason: "not HTTPS" });
  if (/^[a-z]+:\/\/[^/]*@/i.test(target)) flags.push({ level: "high", reason: "URL contains '@' (user-info trick)" });

  const claimed = claimedDomain(text);
  if (claimed && baseDomain(claimed) !== baseDomain(host)) {
    flags.push({
      level: "high",
      reason: `link text shows "${claimed}" but it actually goes to "${host}"`,
    });
  }

  return { href: target, text, host, flags };
}

function safeProtocol(url) {
  try {
    return new URL(url).protocol;
  } catch {
    return "";
  }
}

function headerValues(headers, name) {
  const lower = name.toLowerCase();
  return (headers || [])
    .filter((h) => String(h?.name || "").toLowerCase() === lower)
    .map((h) => String(h?.value || ""));
}

export function analyzeHeaders(headers) {
  // The topmost Authentication-Results is the one added by Microsoft's
  // receiving server; lower ones could be forged by the sender.
  const auth = headerValues(headers, "Authentication-Results")[0] || "";
  const pick = (key) => auth.match(new RegExp(`\\b${key}=([a-z]+)`, "i"))?.[1]?.toLowerCase() || null;
  const receivedSpf = headerValues(headers, "Received-SPF")[0] || "";

  const forefront = headerValues(headers, "X-Forefront-Antispam-Report")[0] || "";
  const antispam = headerValues(headers, "X-Microsoft-Antispam")[0] || "";
  const sclHeader = headerValues(headers, "X-MS-Exchange-Organization-SCL")[0];
  const sclMatch = forefront.match(/\bSCL:(-?\d+)/i);
  const scl = sclHeader !== undefined && sclHeader !== "" ? Number(sclHeader) : sclMatch ? Number(sclMatch[1]) : null;
  const bclMatch = antispam.match(/\bBCL:(\d+)/i);

  const returnPath = (headerValues(headers, "Return-Path")[0] || "").replace(/[<>\s]/g, "");

  return {
    spf: pick("spf") || (receivedSpf.split(/\s/)[0] || "").toLowerCase() || null,
    dkim: pick("dkim"),
    dmarc: pick("dmarc"),
    compauth: pick("compauth"),
    scl: Number.isFinite(scl) ? scl : null,
    bcl: bclMatch ? Number(bclMatch[1]) : null,
    spamFilterVerdict: forefront.match(/\bSFV:([A-Z]+)/i)?.[1] || null,
    returnPath: returnPath || null,
    listUnsubscribe: Boolean(headerValues(headers, "List-Unsubscribe")[0]),
    listId: headerValues(headers, "List-Id")[0] || null,
    precedence: (headerValues(headers, "Precedence")[0] || "").toLowerCase() || null,
  };
}

export function analyzeAttachments(attachments) {
  return (attachments || []).map((a) => {
    const name = String(a?.name || "");
    const ext = getExtension(name);
    const flags = [];
    if (DANGEROUS_EXTENSIONS.has(ext)) flags.push({ level: "high", reason: `.${ext} files can run code or host fake login pages` });
    if (MACRO_EXTENSIONS.has(ext)) flags.push({ level: "medium", reason: `.${ext} is a macro-enabled Office file` });
    if (/\.(pdf|docx?|xlsx?|pptx?|jpe?g|png|txt)\.[a-z0-9]{2,5}$/i.test(name)) {
      flags.push({ level: "high", reason: "double extension (disguised file type)" });
    }
    return { name, contentType: a?.contentType || "", size: a?.size || 0, flags };
  });
}

// Combine all signals into an explainable assessment.
export function assessMessage({ message, folderName, attachments }) {
  const headers = analyzeHeaders(message.internetMessageHeaders);
  const links = extractLinks(message.body?.content, message.body?.contentType);
  const atts = analyzeAttachments(attachments);

  const fromAddress = message.from?.emailAddress?.address || "";
  const fromName = message.from?.emailAddress?.name || "";
  const fromDomain = domainOfAddress(fromAddress);
  const replyTo = (message.replyTo || []).map((r) => r?.emailAddress?.address).filter(Boolean);

  const reasons = [];
  const add = (level, reason) => reasons.push({ level, reason });

  // Authentication
  if (headers.dmarc === "fail") add("high", "DMARC failed: the sender domain did not authorize this mail");
  if (headers.compauth === "fail") add("high", "Microsoft composite authentication failed (likely spoofed)");
  if (headers.spf === "fail" && headers.dkim !== "pass") add("high", "SPF failed and DKIM did not pass");
  else if (headers.spf === "softfail" && headers.dkim !== "pass") add("medium", "SPF softfail and DKIM did not pass");
  if (!headers.spf && !headers.dkim && !headers.dmarc) add("info", "no authentication results header (internal mail or not available)");

  // Microsoft's own filtering
  if (headers.scl !== null && headers.scl >= 5) add("high", `Microsoft spam confidence level SCL=${headers.scl}`);
  if (/junk/i.test(folderName || "")) add("high", `message is in the "${folderName}" folder`);

  // Sender identity
  for (const addr of replyTo) {
    const d = domainOfAddress(addr);
    if (d && fromDomain && baseDomain(d) !== baseDomain(fromDomain)) {
      add("medium", `replies go to ${addr}, a different domain than the sender (${fromDomain})`);
    }
  }
  const rpDomain = domainOfAddress(headers.returnPath || "");
  if (rpDomain && fromDomain && baseDomain(rpDomain) !== baseDomain(fromDomain)) {
    add("info", `bounce address domain (${rpDomain}) differs from sender domain; common for mailing services`);
  }
  const nameEmail = fromName.match(/[^\s<>"]+@[^\s<>"]+/)?.[0];
  if (nameEmail && nameEmail.toLowerCase() !== fromAddress.toLowerCase()) {
    add("high", `display name shows "${nameEmail}" but the real address is ${fromAddress}`);
  }
  if (fromDomain.split(".").some((label) => label.startsWith("xn--"))) {
    add("high", "sender domain uses punycode (possible look-alike)");
  }

  for (const link of links) for (const f of link.flags) if (f.level !== "info") add(f.level, `link ${link.host || link.href}: ${f.reason}`);
  for (const att of atts) for (const f of att.flags) add(f.level, `attachment ${att.name}: ${f.reason}`);

  const has = (level) => reasons.some((r) => r.level === level);
  const riskLevel = has("high") ? "high" : has("medium") ? "medium" : "low";

  const marketingSignals = [];
  if (headers.listUnsubscribe) marketingSignals.push("List-Unsubscribe header");
  if (headers.listId) marketingSignals.push("List-Id header");
  if (headers.precedence === "bulk" || headers.precedence === "list") marketingSignals.push(`Precedence: ${headers.precedence}`);
  if (headers.bcl !== null && headers.bcl >= 4) marketingSignals.push(`bulk complaint level BCL=${headers.bcl}`);
  if (message.inferenceClassification === "other") marketingSignals.push('Outlook classified it as "Other" (not Focused)');

  return {
    id: message.id,
    subject: message.subject || "",
    from: fromAddress,
    fromName,
    replyTo,
    received: message.receivedDateTime || null,
    folder: folderName || null,
    focused: message.inferenceClassification || null,
    authentication: { spf: headers.spf, dkim: headers.dkim, dmarc: headers.dmarc, compauth: headers.compauth },
    microsoftFilter: { scl: headers.scl, bcl: headers.bcl, verdict: headers.spamFilterVerdict },
    riskLevel,
    reasons,
    likelyMarketing: marketingSignals.length > 0,
    marketingSignals,
    links: links.map((l) => ({ text: l.text, url: l.href, host: l.host, flags: l.flags.map((f) => `${f.level}: ${f.reason}`) })),
    attachments: atts.map((a) => ({ name: a.name, size: a.size, flags: a.flags.map((f) => `${f.level}: ${f.reason}`) })),
    webLink: message.webLink || null,
  };
}

export function formatAssessment(a) {
  const icon = { high: "🔴", medium: "🟠", low: "🟢" }[a.riskLevel];
  const auth = a.authentication;
  const lines = [
    `${icon} Risk: ${a.riskLevel.toUpperCase()}${a.likelyMarketing ? " · 📢 likely marketing/bulk" : ""}`,
    `Subject: ${a.subject || "(no subject)"}`,
    `From: ${a.fromName ? `${a.fromName} <${a.from}>` : a.from}`,
    a.replyTo.length ? `Reply-To: ${a.replyTo.join(", ")}` : "",
    `Folder: ${a.folder || "?"}${a.focused ? ` · Inbox tab: ${a.focused}` : ""}`,
    `Auth: SPF=${auth.spf ?? "n/a"} DKIM=${auth.dkim ?? "n/a"} DMARC=${auth.dmarc ?? "n/a"}${auth.compauth ? ` compauth=${auth.compauth}` : ""}` +
      (a.microsoftFilter.scl !== null || a.microsoftFilter.bcl !== null
        ? ` · Microsoft SCL=${a.microsoftFilter.scl ?? "n/a"} BCL=${a.microsoftFilter.bcl ?? "n/a"}`
        : ""),
  ];
  const shown = a.reasons.filter((r) => r.level !== "info");
  if (shown.length) lines.push("Warning signs:\n" + shown.map((r) => `- [${r.level}] ${r.reason}`).join("\n"));
  if (a.marketingSignals.length) lines.push("Marketing signals: " + a.marketingSignals.join("; "));
  if (a.links.length) {
    lines.push(
      `Links (${a.links.length}):\n` +
        a.links
          .slice(0, 15)
          .map((l) => `- ${l.text ? `"${l.text.slice(0, 60)}" → ` : ""}${l.host || l.url}${l.flags.some((f) => !f.startsWith("info")) ? " ⚠️" : ""}`)
          .join("\n") +
        (a.links.length > 15 ? `\n- … ${a.links.length - 15} more` : "")
    );
  }
  if (a.attachments.length) {
    lines.push("Attachments: " + a.attachments.map((x) => `${x.name}${x.flags.length ? " ⚠️" : ""}`).join(", "));
  }
  lines.push(`Message ID: ${a.id}`);
  return lines.filter(Boolean).join("\n");
}
