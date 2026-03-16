# Claude Code Prompt: MCP Microsoft 365 Server

## Context

Ik bouw een persoonlijke AI-agent met meerdere MCP servers. Nu voegen we Microsoft 365 toe — mijn zakelijke omgeving (Pontifexx). We bouwen dit ZELF op de Microsoft Graph API, geen third-party pakketten. Volledige controle over eigen zakelijke data.

De server volgt hetzelfde patroon als mijn bestaande MCP servers: TypeScript, Streamable HTTP, eigen Docker container op een gedeeld Docker network.

## Einddoel (alle fases)

Een MCP server `mcp-ms365` die toegang biedt tot:
- **Outlook Mail** — lezen, zoeken, versturen, drafts
- **Outlook Calendar** — events lezen, aanmaken, conflicten detecteren
- **Teams** — chats lezen, berichten sturen, kanalen volgen
- **SharePoint/OneDrive** — bestanden zoeken, lezen, downloaden

Alles via de Microsoft Graph API met OAuth 2.0 authenticatie.

## Architectuur

```
┌─────────────────┐         ┌──────────────────────────────────┐
│   agent-core    │  HTTP   │         mcp-ms365                │
│   (Python)      │────────▶│         (TypeScript)             │
│                 │         │                                  │
│                 │         │  Streamable HTTP :3000/mcp       │
│                 │         │  Health check    :3000/health    │
│                 │         │                                  │
│                 │         │  Microsoft Graph API ────────▶ Microsoft 365
│                 │         │  (OAuth 2.0 via MSAL)            │
└─────────────────┘         └──────────────────────────────────┘
```

Agent config:
```python
"ms365": {
    "type": "http",
    "url": "http://mcp-ms365:3000/mcp",
    "headers": {"Authorization": "Bearer ${MCP_AUTH_TOKEN}"}
}
```

## Fases overzicht

```
Fase 1: Fundament + Mail (NU BOUWEN)
  → OAuth, project opzet, mail lezen/zoeken/versturen

Fase 2: Calendar
  → Events lezen, aanmaken, conflicten, agenda-view

Fase 3: Teams
  → Chats lezen, berichten sturen, kanalen

Fase 4: SharePoint & OneDrive
  → Bestanden zoeken, lezen, downloaden
```

---

## Fase 1: Fundament + Mail (BOUW DIT NU)

### Vereisten

**Azure AD App Registration (handmatig, vóór de build):**
De volgende app registration moet bestaan in Azure Entra ID:
- Type: "Accounts in this organizational directory only" (single tenant)
- Redirect URI: `http://localhost:3000/auth/callback` (voor device code flow niet nodig, maar handig voor toekomst)
- API permissions (delegated):
  - `Mail.Read`
  - `Mail.ReadWrite`
  - `Mail.Send`
  - `User.Read`
- Client secret aangemaakt
- De volgende waarden beschikbaar als env vars: `MS365_CLIENT_ID`, `MS365_CLIENT_SECRET`, `MS365_TENANT_ID`

### Technische specificaties

**Taal & runtime:** TypeScript, Node.js 22+, strict mode

**Dependencies:**
- `@modelcontextprotocol/sdk` — MCP SDK
- `@azure/msal-node` — Microsoft Authentication Library (officieel, van Microsoft)
- `express` — HTTP server
- `zod` — Input validatie
- `node-fetch` of ingebouwde fetch — Graph API calls

**Transport:** Streamable HTTP, zelfde patroon als de andere MCP servers

### Projectstructuur

```
mcp-ms365/
├── Dockerfile
├── package.json
├── tsconfig.json
├── README.md
├── .env.example
└── src/
    ├── index.ts              # Entrypoint: Express + MCP setup + auth middleware
    ├── server.ts             # McpServer instantie + tool registraties per fase
    │
    ├── auth/
    │   ├── msal.ts           # MSAL client setup, token cache, device code flow
    │   └── graph-client.ts   # Wrapper voor Graph API calls (GET, POST, PATCH, DELETE)
    │
    ├── tools/
    │   ├── auth.ts           # ms365_auth_status, ms365_login (device code)
    │   └── mail.ts           # Fase 1: mail tools
    │   # Later: calendar.ts, teams.ts, sharepoint.ts
    │
    ├── shared/
    │   ├── auth-middleware.ts # Bearer token check (MCP endpoint)
    │   ├── types.ts          # Gedeelde TypeScript types
    │   └── config.ts         # Config uit env vars
    │
    └── data/                 # Volume mount point
        └── token-cache.json  # MSAL token cache (persistent)
```

### Auth: MSAL + Device Code Flow

**Waarom device code flow:**
De server draait headless op een VPS. Device code flow werkt zonder browser op de server — je krijgt een URL + code, opent die in je browser op je telefoon/laptop, logt in, klaar.

**MSAL setup:**
```typescript
import { PublicClientApplication } from "@azure/msal-node";

const msalConfig = {
    auth: {
        clientId: process.env.MS365_CLIENT_ID,
        authority: `https://login.microsoftonline.com/${process.env.MS365_TENANT_ID}`,
        clientSecret: process.env.MS365_CLIENT_SECRET
    }
};

const pca = new PublicClientApplication(msalConfig);
```

**Token cache:**
- MSAL tokens opslaan in `/data/token-cache.json` (Docker volume, persistent)
- Bij startup: probeer silent token acquisition (cached token)
- Als token verlopen: automatische refresh via MSAL (refresh token)
- Als refresh token ook verlopen: markeer als niet-ingelogd, wacht op `ms365_login`

**Graph API wrapper:**
Een gedeelde functie die alle Graph API calls doet:
```typescript
async function graphRequest(method: string, endpoint: string, body?: any): Promise<any> {
    const token = await getAccessToken();
    const response = await fetch(`https://graph.microsoft.com/v1.0${endpoint}`, {
        method,
        headers: {
            Authorization: `Bearer ${token}`,
            "Content-Type": "application/json"
        },
        body: body ? JSON.stringify(body) : undefined
    });
    // Error handling, rate limiting, pagination
}
```

Met automatische pagination voor lijst-endpoints (Graph API gebruikt `@odata.nextLink`).

### Tools — Fase 1: Auth + Mail

**1. `ms365_auth_status`**
- Check of er een geldige sessie is
- Output: `{authenticated: boolean, user_email: string|null, token_expires: string|null}`
- Annotations: readOnlyHint: true

**2. `ms365_login`**
- Start device code flow
- Output: `{user_code: string, verification_url: string, message: "Open the URL and enter the code to authenticate"}`
- De agent toont dit aan de gebruiker via Telegram
- Na succesvolle authenticatie: token gecached
- Annotations: readOnlyHint: false

**3. `ms365_list_mail_folders`**
- Lijst van mail folders (Inbox, Sent, Drafts, etc.)
- Output: array van `{id, name, unread_count, total_count}`
- Annotations: readOnlyHint: true

**4. `ms365_search_mail`**
- Zoek emails op basis van criteria
- Input:
  - `query` (string, optional): zoekterm (doorzoekt subject, body, from)
  - `from` (string, optional): filter op afzender
  - `folder` (string, optional, default "Inbox"): folder ID of naam
  - `days` (number, optional, default 30): hoe ver terug zoeken
  - `unread_only` (boolean, optional, default false)
  - `limit` (number, optional, default 20)
  - `has_attachments` (boolean, optional): filter op bijlagen
- Graph endpoint: `GET /me/mailFolders/{folder}/messages` met `$filter` en `$search`
- Output: array van `{id, subject, from, to, date, preview, unread, has_attachments, importance}`
- `preview`: eerste ~200 karakters van de body (plain text)
- Sorteer op datum (nieuwste eerst)
- Annotations: readOnlyHint: true

**5. `ms365_get_mail`**
- Haal volledige email op
- Input:
  - `id` (string, required): message ID
- Graph endpoint: `GET /me/messages/{id}`
- Output: `{id, subject, from, to, cc, date, body_text, body_html, attachments: [{name, size, content_type}], importance, categories}`
- Attachments: alleen metadata, niet de inhoud
- Annotations: readOnlyHint: true

**6. `ms365_send_mail`**
- Verstuur een email
- Input:
  - `to` (string, required): ontvanger(s), komma-gescheiden
  - `subject` (string, required)
  - `body` (string, required): plain text of HTML
  - `cc` (string, optional)
  - `bcc` (string, optional)
  - `importance` (string, optional): "low", "normal", "high"
  - `reply_to_id` (string, optional): message ID als dit een reply is
- Graph endpoint: `POST /me/sendMail` of `POST /me/messages/{id}/reply`
- Output: `{message: "Email sent", to, subject}`
- Tool description MOET bevatten: "ALWAYS draft first using ms365_draft_mail and present to user for approval before sending, UNLESS user explicitly says to send directly."
- Annotations: readOnlyHint: false, destructiveHint: false

**7. `ms365_draft_mail`**
- Maak een concept email (sla op in Drafts, verstuur NIET)
- Input: zelfde als ms365_send_mail
- Graph endpoint: `POST /me/messages` (maakt draft aan)
- Output: `{id, subject, to, message: "Draft saved in Drafts folder. Review before sending."}`
- Annotations: readOnlyHint: false

**8. `ms365_move_mail`**
- Verplaats email naar andere folder
- Input:
  - `id` (string, required): message ID
  - `destination_folder` (string, required): folder naam of ID
- Graph endpoint: `POST /me/messages/{id}/move`
- Output: `{id, message: "Moved to {folder}"}`
- Annotations: readOnlyHint: false

**9. `ms365_mark_read`**
- Markeer email als gelezen/ongelezen
- Input:
  - `id` (string, required): message ID
  - `is_read` (boolean, required)
- Graph endpoint: `PATCH /me/messages/{id}`
- Output: `{id, is_read, message: "Marked as read/unread"}`
- Annotations: readOnlyHint: false

### Environment variables

```
# MCP
MCP_AUTH_TOKEN=             # Bearer token voor MCP endpoint
PORT=3000

# Azure AD
MS365_CLIENT_ID=            # Azure AD app registration client ID
MS365_CLIENT_SECRET=        # Azure AD app registration client secret
MS365_TENANT_ID=            # Azure AD tenant ID

# Paden
DATA_DIR=/data              # Persistent volume (token cache)
TZ=Europe/Amsterdam
```

### Dockerfile

- Base image: `node:22-slim`
- Multi-stage build
- Non-root user
- Expose poort 3000
- Volume: `/data` (token cache)
- Health check: `curl -f http://localhost:3000/health || exit 1`

### Error handling

- Niet geauthenticeerd: "Not authenticated. Use ms365_login to authenticate via device code flow."
- Token verlopen + refresh mislukt: "Session expired. Please re-authenticate with ms365_login."
- Graph API 429 (rate limit): retry met backoff, meld als het aanhoudt
- Graph API 403 (insufficient permissions): duidelijke melding welke permission ontbreekt
- Folder niet gevonden: melding met lijst van beschikbare folders

### Test scenario's

1. Server start, `/health` → 200 OK
2. `ms365_auth_status` → not authenticated
3. `ms365_login` → device code + URL
4. Na authenticatie: `ms365_auth_status` → authenticated met email
5. `ms365_list_mail_folders` → Inbox, Sent, Drafts, etc.
6. `ms365_search_mail` → recente emails
7. `ms365_get_mail` met id → volledige email
8. `ms365_draft_mail` → concept in Drafts
9. `ms365_send_mail` → email verstuurd
10. Container restart → token cache behouden, automatisch ingelogd

---

## Fase 2: Calendar (later)

Extra API permissions: `Calendars.Read`, `Calendars.ReadWrite`

Geplande tools:
- `ms365_list_calendars` — alle kalenders
- `ms365_get_events` — events in datumbereik
- `ms365_get_calendar_view` — agenda-weergave (expanded recurring events)
- `ms365_create_event` — nieuw event aanmaken
- `ms365_update_event` — event wijzigen
- `ms365_delete_event` — event verwijderen
- `ms365_check_conflicts` — conflicten detecteren voor een tijdslot

Graph endpoints: `/me/calendars`, `/me/calendarView`, `/me/events`

## Fase 3: Teams (later)

Extra API permissions: `Chat.Read`, `Chat.ReadWrite`, `ChannelMessage.Read.All`, `ChannelMessage.Send`

Geplande tools:
- `ms365_list_chats` — recente Teams chats
- `ms365_get_chat_messages` — berichten uit een chat
- `ms365_send_chat_message` — bericht sturen in chat
- `ms365_list_teams` — teams waar ik lid van ben
- `ms365_list_channels` — kanalen in een team
- `ms365_get_channel_messages` — berichten uit een kanaal
- `ms365_send_channel_message` — bericht in kanaal posten

Graph endpoints: `/me/chats`, `/me/joinedTeams`, `/teams/{id}/channels`

## Fase 4: SharePoint & OneDrive (later)

Extra API permissions: `Files.Read`, `Files.ReadWrite`, `Sites.Read.All`

Geplande tools:
- `ms365_search_files` — bestanden zoeken over OneDrive en SharePoint
- `ms365_list_files` — bestanden in een folder
- `ms365_get_file_content` — bestandsinhoud ophalen (tekst, of download pad)
- `ms365_upload_file` — bestand uploaden
- `ms365_list_sharepoint_sites` — SharePoint sites
- `ms365_search_sharepoint` — zoeken binnen SharePoint

Graph endpoints: `/me/drive`, `/sites`, `/drives`

---

## Wat je NIET moet bouwen

- Geen admin/tenant-level operaties (alleen user-level via delegated permissions)
- Geen attachment download (alleen metadata, te groot voor context)
- Geen real-time webhooks/subscriptions (pull-based, past bij heartbeat)
- Geen OneNote, Planner, To Do (Todoist dekt taken al)
- Geen web UI

## Oplevering Fase 1

- Werkend project met auth + 9 mail tools
- Dockerfile (multi-stage)
- README met: Azure AD setup instructies, env vars, login flow
- .env.example
- `npm run build` compileert foutloos
- Token cache persistent op Docker volume
- Klaar om Fase 2 (calendar) toe te voegen zonder refactoring
