import { getValidAccessToken } from "./auth.mjs";
import { GRAPH_BASE, SCOPES } from "./config.mjs";
import { log } from "./logger.mjs";

export async function graphGetJson(url) {
  const token = await getValidAccessToken();

  const res = await fetch(url, {
    method: "GET",
    headers: {
      Authorization: `Bearer ${token}`,
      Accept: "application/json",
    },
  });

  if (!res.ok) {
    const text = await res.text();
    log("graphGetJson failed", { url, status: res.status, text });
    throw new Error(`Graph API ${res.status}: ${text}`);
  }

  return res.json();
}

export async function graphGetBytes(url) {
  const token = await getValidAccessToken();

  const res = await fetch(url, {
    method: "GET",
    headers: {
      Authorization: `Bearer ${token}`,
    },
  });

  if (!res.ok) {
    const text = await res.text();
    log("graphGetBytes failed", { url, status: res.status, text });
    throw new Error(`Graph API ${res.status}: ${text}`);
  }

  const arrayBuffer = await res.arrayBuffer();
  return Buffer.from(arrayBuffer);
}

export function mailboxPrefix(mailbox = "me") {
  return mailbox === "me"
    ? `${GRAPH_BASE}/me`
    : `${GRAPH_BASE}/users/${encodeURIComponent(mailbox)}`;
}

export function messageCollectionUrl(mailbox = "me", folder = "inbox") {
  const base = mailboxPrefix(mailbox);
  const normalizedFolder = String(folder || "inbox").trim().toLowerCase();

  if (!normalizedFolder || normalizedFolder === "inbox") {
    return `${base}/mailFolders/inbox/messages`;
  }

  if (normalizedFolder === "all" || normalizedFolder === "messages") {
    return `${base}/messages`;
  }

  return `${base}/mailFolders/${encodeURIComponent(folder)}/messages`;
}

export async function graphRequest(method, url, bodyObj) {
  const token = await getValidAccessToken();
  const res = await fetch(url, {
    method,
    headers: {
      Authorization: `Bearer ${token}`,
      "Content-Type": "application/json",
      Accept: "application/json",
    },
    body: bodyObj ? JSON.stringify(bodyObj) : undefined,
  });

  if (!res.ok) {
    const text = await res.text();
    log("graphRequest failed", { method, url, status: res.status, text });
    if (res.status === 403 || /AccessDenied|insufficient privileges/i.test(text)) {
      throw new Error(
        `Graph API ${res.status}: permission denied. Make sure the Entra app has ${SCOPES.join(", ")} and re-run begin_auth to consent. Details: ${text}`
      );
    }
    throw new Error(`Graph API ${res.status}: ${text}`);
  }

  const text = await res.text();
  return text ? JSON.parse(text) : null;
}
