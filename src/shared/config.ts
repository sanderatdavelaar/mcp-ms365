export const config = {
  authToken: process.env.MCP_AUTH_TOKEN || "",
  port: parseInt(process.env.PORT || "3000", 10),
  clientId: process.env.MS365_CLIENT_ID || "",
  tenantId: process.env.MS365_TENANT_ID || "",
  dataDir: process.env.DATA_DIR || "/data",
};

export function validateConfig(): void {
  if (!config.clientId) {
    throw new Error("MS365_CLIENT_ID environment variable is required");
  }
  if (!config.tenantId) {
    throw new Error("MS365_TENANT_ID environment variable is required");
  }
}
