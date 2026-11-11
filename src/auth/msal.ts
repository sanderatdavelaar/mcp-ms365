import {
  PublicClientApplication,
  type AccountInfo,
  type DeviceCodeRequest,
} from "@azure/msal-node";
import { join } from "node:path";
import { config } from "../shared/config.js";
import { TokenCacheStore } from "./token-cache.js";
import type { AuthStatus } from "../shared/types.js";

const SCOPES = [
  "Mail.Read", "Mail.ReadWrite", "Mail.Send",
  "Calendars.Read", "Calendars.ReadWrite",
  "Chat.Read", "Chat.ReadWrite",
  "ChannelMessage.Read.All", "ChannelMessage.Send",
  "Team.ReadBasic.All", "Channel.ReadBasic.All",
  "User.Read",
];

let pca: PublicClientApplication;
let currentAccount: AccountInfo | null = null;

let cacheStore: TokenCacheStore;

async function loadCache(): Promise<boolean> {
  const { content, needsMigration } = await cacheStore.load();
  if (content === null) return false;
  try {
    pca.getTokenCache().deserialize(content);
  } catch {
    // Cache file is corrupted — that's fine, user can log in again
    return false;
  }
  return needsMigration;
}

async function persistCache(): Promise<boolean> {
  try {
    await cacheStore.save(pca.getTokenCache().serialize());
    return true;
  } catch (error) {
    console.error("Failed to persist token cache:", error);
    return false;
  }
}

export async function initMsal(): Promise<void> {
  pca = new PublicClientApplication({
    auth: {
      clientId: config.clientId,
      authority: `https://login.microsoftonline.com/${config.tenantId}`,
    },
  });

  cacheStore = new TokenCacheStore(join(config.dataDir, "token-cache.json"), config.cacheKey);
  const needsMigration = await loadCache();
  if (needsMigration) {
    // Onversleutelde cache van een eerdere versie: meteen versleuteld terugschrijven.
    if (await persistCache()) {
      console.log("Tokencache versleuteld (migratie van onversleutelde cache)");
    }
  }

  try {
    const accounts = await pca.getTokenCache().getAllAccounts();
    if (accounts.length > 0) {
      currentAccount = accounts[0];
      // Verify the token is still valid via silent acquisition
      try {
        await pca.acquireTokenSilent({
          account: currentAccount,
          scopes: SCOPES,
        });
        console.log(`MSAL initialized — authenticated as ${currentAccount.username}`);
      } catch {
        // Silent acquisition failed; user will need to re-authenticate
        console.log("MSAL initialized — cached account found but token refresh failed. Use ms365_login to re-authenticate.");
      }
    } else {
      console.log("MSAL initialized — no cached accounts. Use ms365_login to authenticate.");
    }
  } catch {
    console.log("MSAL initialized — no cached accounts. Use ms365_login to authenticate.");
  }
}

export async function getAccessToken(): Promise<string> {
  if (!currentAccount) {
    throw new Error(
      "Not authenticated. Use ms365_login to authenticate via device code flow."
    );
  }

  try {
    const result = await pca.acquireTokenSilent({
      account: currentAccount,
      scopes: SCOPES,
    });
    await persistCache();
    return result.accessToken;
  } catch {
    // Reset account since token is no longer valid
    currentAccount = null;
    throw new Error(
      "Not authenticated. Use ms365_login to authenticate via device code flow."
    );
  }
}

export function startDeviceCodeFlow(): Promise<{
  userCode: string;
  verificationUri: string;
  message: string;
}> {
  return new Promise((resolve, reject) => {
    const tokenPromise = pca.acquireTokenByDeviceCode({
      scopes: SCOPES,
      deviceCodeCallback: (response) => {
        resolve({
          userCode: response.userCode,
          verificationUri: response.verificationUri,
          message: response.message,
        });
      },
    });

    // Handle the token promise in the background
    tokenPromise
      .then(async (result) => {
        if (result) {
          currentAccount = result.account;
          await persistCache();
          console.log(`Authenticated as ${result.account?.username}`);
        }
      })
      .catch((error) => {
        console.error("Device code flow failed:", error);
      });
  });
}

export function getAuthStatus(): AuthStatus {
  if (!currentAccount) {
    return {
      authenticated: false,
      userEmail: null,
      tokenExpires: null,
    };
  }

  // Try to get token expiry from the cached account's ID token claims
  let tokenExpires: string | null = null;
  const claims = currentAccount.idTokenClaims as
    | Record<string, unknown>
    | undefined;
  if (claims?.exp) {
    tokenExpires = new Date((claims.exp as number) * 1000).toISOString();
  }

  return {
    authenticated: true,
    userEmail: currentAccount.username || null,
    tokenExpires,
  };
}

export function isAuthenticated(): boolean {
  return currentAccount !== null;
}
