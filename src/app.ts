import express from "express";
import type { Express, Request, Response } from "express";
import type { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { StreamableHTTPServerTransport } from "@modelcontextprotocol/sdk/server/streamableHttp.js";
import { createAuthMiddleware } from "./shared/auth-middleware.js";

export interface AppOptions {
  authToken: string | undefined;
  createServer: () => McpServer;
}

/**
 * Stateless Streamable HTTP: per request een nieuwe McpServer en transport,
 * zodat gelijktijdige requests elkaar niet raken ("Already connected").
 * De auth staat vóór de body-parser: zonder token wordt er niets geparsed.
 */
export function createApp({ authToken, createServer }: AppOptions): Express {
  const authMiddleware = createAuthMiddleware(authToken);
  const app = express();

  app.get("/health", (_req: Request, res: Response) => {
    res.json({ status: "ok" });
  });

  app.post("/mcp", authMiddleware, express.json(), async (req: Request, res: Response) => {
    const server = createServer();
    const transport = new StreamableHTTPServerTransport({
      sessionIdGenerator: undefined,
      enableJsonResponse: true,
    });
    res.on("close", () => {
      void transport.close();
      void server.close();
    });
    try {
      await server.connect(transport);
      await transport.handleRequest(req, res, req.body);
    } catch (error) {
      console.error("Error handling MCP request:", error);
      if (!res.headersSent) {
        res.status(500).json({
          jsonrpc: "2.0",
          error: { code: -32603, message: "Internal server error" },
          id: null,
        });
      }
    }
  });

  return app;
}
