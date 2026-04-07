import express from "express";
import type { Request, Response } from "express";
import { StreamableHTTPServerTransport } from "@modelcontextprotocol/sdk/server/streamableHttp.js";
import { createMs365Server } from "./server.js";
import { createAuthMiddleware } from "./shared/auth-middleware.js";
import { config, validateConfig } from "./shared/config.js";
import { initMsal } from "./auth/msal.js";

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
