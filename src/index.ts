import express from "express";
import type { Request, Response } from "express";
import { StreamableHTTPServerTransport } from "@modelcontextprotocol/sdk/server/streamableHttp.js";
import { createMs365Server } from "./server.js";
import { createAuthMiddleware } from "./shared/auth-middleware.js";
import { config, validateConfig } from "./shared/config.js";
import { initMsal, getAccessToken, getAuthStatus } from "./auth/msal.js";

async function main() {
  // Validate required env vars
  validateConfig();

  // Initialize MSAL (load token cache, attempt silent auth)
  await initMsal();

  const server = createMs365Server();
  const app = express();
  app.use(express.json());

  // Health check (no auth)
  app.get("/health", (_req: Request, res: Response) => {
    res.json({ status: "ok" });
  });

  const authMiddleware = createAuthMiddleware(config.authToken);

  // Internal token endpoint for trusted services (e.g. agent-v4 Graph
  // subscription manager). Bearer-protected, returns the current Graph
  // access token. Not exposed via MCP to keep tokens out of LLM context.
  app.get("/internal/token", authMiddleware, async (_req: Request, res: Response) => {
    try {
      const token = await getAccessToken();
      const status = getAuthStatus();
      res.json({
        access_token: token,
        user_email: status.userEmail,
        token_expires: status.tokenExpires,
      });
    } catch (error) {
      res.status(401).json({ error: String(error) });
    }
  });

  // MCP endpoint with Streamable HTTP transport
  app.post("/mcp", authMiddleware, async (req: Request, res: Response) => {
    const transport = new StreamableHTTPServerTransport({
      sessionIdGenerator: undefined,
      enableJsonResponse: true,
    });
    res.on("close", () => {
      transport.close();
    });
    await server.connect(transport);
    await transport.handleRequest(req, res, req.body);
  });

  app.listen(config.port, () => {
    console.log(`MCP MS365 server running on port ${config.port}`);
  });
}

main().catch((error) => {
  console.error("Failed to start server:", error);
  process.exit(1);
});
