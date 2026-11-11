import { createApp } from "./app.js";
import { createMs365Server } from "./server.js";
import { config, validateConfig } from "./shared/config.js";
import { initMsal } from "./auth/msal.js";

async function main() {
  // Validate required env vars (zonder MCP_AUTH_TOKEN start de server niet)
  validateConfig();

  // Initialize MSAL (load token cache, attempt silent auth)
  await initMsal();

  const app = createApp({
    authToken: config.authToken,
    createServer: createMs365Server,
  });

  app.listen(config.port, () => {
    console.log(`MCP MS365 server running on port ${config.port}`);
  });
}

main().catch((error) => {
  console.error("Failed to start server:", error);
  process.exit(1);
});
