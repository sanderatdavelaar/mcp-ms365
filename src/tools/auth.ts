import type { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { getAuthStatus, startDeviceCodeFlow, getAccessToken } from "../auth/msal.js";

export function registerAuthTools(server: McpServer): void {
  server.tool(
    "ms365_auth_status",
    "Check Microsoft 365 authentication status. Returns whether you are currently authenticated, the user email, and token expiration.",
    {},
    { readOnlyHint: true },
    async () => {
      try {
        const status = getAuthStatus();
        return {
          content: [{ type: "text" as const, text: JSON.stringify(status, null, 2) }],
        };
      } catch (error) {
        return {
          content: [{ type: "text" as const, text: `Error: ${error}` }],
          isError: true,
        };
      }
    }
  );

  server.tool(
    "ms365_get_access_token",
    "Return the current Microsoft Graph access token (and expiry) for use by trusted internal services such as the agent-v4 Graph subscription manager. Sensitive — do not expose to end users.",
    {},
    { readOnlyHint: true },
    async () => {
      try {
        const token = await getAccessToken();
        const status = getAuthStatus();
        return {
          content: [
            {
              type: "text" as const,
              text: JSON.stringify(
                {
                  access_token: token,
                  user_email: status.userEmail,
                  token_expires: status.tokenExpires,
                },
                null,
                2
              ),
            },
          ],
        };
      } catch (error) {
        return {
          content: [{ type: "text" as const, text: `Error: ${error}` }],
          isError: true,
        };
      }
    }
  );

  server.tool(
    "ms365_login",
    "Authenticate with Microsoft 365 using device code flow. Returns a URL and code — open the URL in a browser and enter the code to sign in. The server will automatically detect when authentication completes.",
    {},
    { readOnlyHint: false },
    async () => {
      try {
        const { userCode, verificationUri, message } = await startDeviceCodeFlow();
        return {
          content: [
            {
              type: "text" as const,
              text: JSON.stringify(
                { user_code: userCode, verification_url: verificationUri, message },
                null,
                2
              ),
            },
          ],
        };
      } catch (error) {
        return {
          content: [{ type: "text" as const, text: `Error: ${error}` }],
          isError: true,
        };
      }
    }
  );
}
