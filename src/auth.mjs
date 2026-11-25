import fs from "node:fs";
import path from "node:path";
import { spawn } from "node:child_process";
import { PublicClientApplication } from "@azure/msal-node";
import { APP_DATA_DIR, AUTHORITY, AUTO_OPEN_BROWSER, CLIENT_ID, SCOPES } from "./config.mjs";
import { log } from "./logger.mjs";

export function getDeviceLoginUrlFromMessage(message) {
  const text = String(message || "");
  const match = text.match(/https?:\/\/[^\s]+/i);
  return match ? match[0] : "https://microsoft.com/devicelogin";
}

export function openUrlInBrowser(url) {
  if (!AUTO_OPEN_BROWSER || !url || typeof url !== "string") {
    return false;
  }

  try {
    if (process.platform === "win32") {
      spawn("cmd", ["/c", "start", "", url], {
        detached: true,
        stdio: "ignore",
      }).unref();
      return true;
    }

    if (process.platform === "darwin") {
      spawn("open", [url], {
        detached: true,
        stdio: "ignore",
      }).unref();
      return true;
    }

    spawn("xdg-open", [url], {
      detached: true,
      stdio: "ignore",
    }).unref();
    return true;
  } catch (err) {
    log("openUrlInBrowser failed", err);
    return false;
  }
}

if (!CLIENT_ID) {
  log("Missing M365_CLIENT_ID in env");
}

export const MSAL_CACHE_PATH = path.join(APP_DATA_DIR, "msal-cache.json");

// Persist MSAL's token cache (including the refresh token) to disk so auth
// survives app/process restarts. With this the user signs in once and the
// server can silently refresh access tokens going forward (refresh tokens roll
// for ~90 days), which is what makes unattended scheduled runs reliable.
// The file holds sensitive tokens, so it is written with 0600 permissions.
export const cachePlugin = {
  async beforeCacheAccess(context) {
    try {
      if (fs.existsSync(MSAL_CACHE_PATH)) {
        context.tokenCache.deserialize(fs.readFileSync(MSAL_CACHE_PATH, "utf8"));
      }
    } catch (err) {
      log("MSAL cache read failed", err?.message || String(err));
    }
  },
  async afterCacheAccess(context) {
    try {
      if (context.cacheHasChanged) {
        fs.writeFileSync(MSAL_CACHE_PATH, context.tokenCache.serialize(), {
          mode: 0o600,
        });
      }
    } catch (err) {
      log("MSAL cache write failed", err?.message || String(err));
    }
  },
};

export const pca = new PublicClientApplication({
  auth: {
    clientId: CLIENT_ID || "missing-client-id",
    authority: AUTHORITY,
  },
  cache: { cachePlugin },
});

export let currentAccount = null;

export let cachedAccessToken = null;

export let cachedTokenExpiresAt = 0;

export let authPromise = null;

export const authState = {
  status: "not_started",
  verificationUri: null,
  userCode: null,
  message: null,
  account: null,
  error: null,
  updatedAt: null,
};

export function setAuthState(patch) {
  Object.assign(authState, patch, {
    updatedAt: new Date().toISOString(),
  });
  log("authState", authState);
}

export function getPublicAuthState() {
  return {
    ...authState,
    configured: !!CLIENT_ID,
    hasCachedToken:
      !!cachedAccessToken && Date.now() < cachedTokenExpiresAt - 60 * 1000,
    tokenExpiresAt: cachedTokenExpiresAt
      ? new Date(cachedTokenExpiresAt).toISOString()
      : null,
  };
}

export function ensureConfigured() {
  if (!CLIENT_ID) {
    throw new Error(
      "Microsoft 365 is not configured. Set M365_CLIENT_ID in the environment before using auth tools."
    );
  }
}

export function getCachedToken() {
  if (
    cachedAccessToken &&
    cachedTokenExpiresAt &&
    Date.now() < cachedTokenExpiresAt - 60 * 1000
  ) {
    return cachedAccessToken;
  }

  return null;
}

export function buildBeginAuthText() {
  const parts = ["Please complete Microsoft 365 login now."];

  if (authState.verificationUri) {
    parts.push(`Open URL: ${authState.verificationUri}`);
  }

  if (authState.userCode) {
    parts.push(`Enter code: ${authState.userCode}`);
  }

  if (authState.message) {
    parts.push(`Full prompt: ${authState.message}`);
  }

  if (!AUTO_OPEN_BROWSER) {
    parts.push("Browser auto-open is disabled by M365_AUTO_OPEN_BROWSER=false.");
  }

  return parts.join("\n");
}

export async function startDeviceCodeAuth() {
  if (authPromise) return authPromise;

  ensureConfigured();

  authPromise = pca
    .acquireTokenByDeviceCode({
      scopes: SCOPES,
      deviceCodeCallback: (response) => {
        log("deviceCodeCallback RAW RESPONSE", JSON.stringify(response));
        const message = response?.message || "";
        const verificationUri = getDeviceLoginUrlFromMessage(message);
        const userCode = response?.userCode || response?.user_code || null;
        const opened = openUrlInBrowser(verificationUri);

        setAuthState({
          status: "pending",
          verificationUri,
          userCode,
          message,
          error: null,
        });

        log("deviceCodeCallback", {
          verificationUri,
          userCode,
          message,
          opened,
        });
      },
    })
    .then((result) => {
      if (!result?.accessToken) {
        throw new Error("No access token returned.");
      }

      currentAccount = result.account ?? null;
      cachedAccessToken = result.accessToken;
      cachedTokenExpiresAt = result.expiresOn
        ? result.expiresOn.getTime()
        : Date.now() + 50 * 60 * 1000;

      setAuthState({
        status: "authenticated",
        account: currentAccount?.username || null,
        error: null,
      });

      return result.accessToken;
    })
    .catch((err) => {
      log("MSAL ERROR FULL DETAIL", {
        name: err?.name,
        message: err?.message,
        errorCode: err?.errorCode,
        errorMessage: err?.errorMessage,
        subError: err?.subError,
        correlationId: err?.correlationId,
        stack: err?.stack,
      });
      setAuthState({
        status: "error",
        error: String(err?.message || err),
      });
      throw err;
    })
    .finally(() => {
      authPromise = null;
    });

  return authPromise;
}

export async function restoreAccountFromCache() {
  if (currentAccount) return currentAccount;
  try {
    const accounts = await pca.getTokenCache().getAllAccounts();
    if (accounts && accounts.length > 0) {
      currentAccount = accounts[0];
      setAuthState({
        status: "authenticated",
        account: currentAccount.username || null,
        error: null,
      });
      log("Restored account from persistent cache", currentAccount.username || "");
    }
  } catch (err) {
    log("restoreAccountFromCache failed", err?.message || String(err));
  }
  return currentAccount;
}

export async function getValidAccessToken() {
  ensureConfigured();

  const cached = getCachedToken();
  if (cached) return cached;

  if (!currentAccount) {
    await restoreAccountFromCache();
  }

  if (currentAccount) {
    try {
      const silent = await pca.acquireTokenSilent({
        account: currentAccount,
        scopes: SCOPES,
      });

      if (silent?.accessToken) {
        cachedAccessToken = silent.accessToken;
        cachedTokenExpiresAt = silent.expiresOn
          ? silent.expiresOn.getTime()
          : Date.now() + 50 * 60 * 1000;

        setAuthState({
          status: "authenticated",
          account: currentAccount?.username || null,
          error: null,
        });

        return silent.accessToken;
      }
    } catch (err) {
      log("acquireTokenSilent failed", err?.message || String(err));
    }
  }

  if (authState.status === "pending") {
    throw new Error(
      "Microsoft 365 login is not yet complete. Finish the browser login first, then call auth_status."
    );
  }

  if (authState.status === "error") {
    throw new Error(
      `Microsoft 365 login is currently in error state: ${authState.error || "unknown error"}. Please call begin_auth again.`
    );
  }

  throw new Error("Microsoft 365 is not logged in yet. Please call begin_auth first.");
}
