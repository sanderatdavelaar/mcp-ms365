import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { registerAuthTools } from "./tools/auth.js";
import { registerMailTools } from "./tools/mail.js";
import { registerCalendarTools } from "./tools/calendar.js";

export function createMs365Server(): McpServer {
  const server = new McpServer({
    name: "mcp-ms365",
    version: "1.0.0",
  });
  registerAuthTools(server);
  registerMailTools(server);
  registerCalendarTools(server);
  return server;
}
