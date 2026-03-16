# mcp-ms365

MCP server for Microsoft 365 (Outlook Mail) via Microsoft Graph API. Exposes email operations as MCP tools over Streamable HTTP transport.

## Prerequisites

### Azure AD App Registration

1. Go to [Azure Entra ID](https://entra.microsoft.com/) > App registrations > New registration
2. Name: `mcp-ms365` (or whatever you prefer)
3. Supported account types: **Accounts in this organizational directory only** (single tenant)
4. Register, then configure:
   - **Authentication** > Enable **Allow public client flows** = Yes
   - **API permissions** > Add delegated permissions:
     - `Mail.Read`
     - `Mail.ReadWrite`
     - `Mail.Send`
     - `User.Read`
5. Note the **Application (client) ID** and **Directory (tenant) ID**

## Environment Variables

| Variable | Required | Default | Description |
|---|---|---|---|
| `MS365_CLIENT_ID` | Yes | — | Azure AD Application (client) ID |
| `MS365_TENANT_ID` | Yes | — | Azure AD Directory (tenant) ID |
| `MCP_AUTH_TOKEN` | No | — | Bearer token to protect the MCP endpoint |
| `PORT` | No | `3000` | HTTP port |
| `DATA_DIR` | No | `/data` | Directory for persistent token cache |
| `TZ` | No | `Europe/Amsterdam` | Timezone |

## Running Locally

```bash
cp .env.example .env
# Fill in your Azure AD values
npm install
npm run build
npm start
```

## Running with Docker

```bash
docker build -t mcp-ms365 .
docker run -d --name mcp-ms365 \
  --env-file .env \
  -v mcp-ms365-data:/data \
  -p 3000:3000 \
  mcp-ms365
```

## Authentication Flow

This server uses the **device code flow** for user authentication:

1. Call the `ms365_login` tool
2. You receive a URL and a code
3. Open the URL in a browser, enter the code, and sign in with your Microsoft account
4. The server detects authentication automatically
5. The token is cached in `DATA_DIR` and persists across restarts

Check status anytime with `ms365_auth_status`.

## Available Tools

| Tool | Description |
|---|---|
| `ms365_login` | Start device code authentication flow |
| `ms365_auth_status` | Check current authentication status |
| `ms365_list_mail_folders` | List all mail folders with unread/total counts |
| `ms365_search_mail` | Search emails by query, sender, folder, date range, and filters |
| `ms365_get_mail` | Get full content of an email by ID |
| `ms365_send_mail` | Send an email (draft first recommended) |
| `ms365_draft_mail` | Create a draft email for review before sending |
| `ms365_move_mail` | Move an email to another folder |
| `ms365_mark_read` | Mark an email as read or unread |
