import { z } from "zod";
import type { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { graphGet, graphPost, graphPatch } from "../auth/graph-client.js";

// --- Helper types ---

interface GraphRecipient {
  emailAddress: { address: string; name?: string };
}

interface GraphMailFolder {
  id: string;
  displayName: string;
  unreadItemCount: number;
  totalItemCount: number;
}

interface GraphMessage {
  id: string;
  subject: string;
  from: GraphRecipient;
  toRecipients: GraphRecipient[];
  ccRecipients?: GraphRecipient[];
  receivedDateTime: string;
  bodyPreview?: string;
  body?: { contentType: string; content: string };
  isRead: boolean;
  hasAttachments: boolean;
  importance: string;
  categories?: string[];
  attachments?: Array<{
    id: string;
    name: string;
    size: number;
    contentType: string;
  }>;
}

interface GraphPagedResponse<T> {
  value: T[];
  "@odata.nextLink"?: string;
  "@odata.count"?: number;
}

// --- Helper: folder resolution ---

const WELL_KNOWN_FOLDERS: Record<string, string> = {
  inbox: "Inbox",
  drafts: "Drafts",
  sentitems: "SentItems",
  deleteditems: "DeletedItems",
  junkemail: "JunkEmail",
  archive: "Archive",
};

const folderCache = new Map<string, string>();

async function resolveFolder(folderNameOrId: string): Promise<string> {
  const lower = folderNameOrId.toLowerCase();

  // Check well-known names (case-insensitive)
  if (WELL_KNOWN_FOLDERS[lower]) {
    return WELL_KNOWN_FOLDERS[lower];
  }

  // Check cache
  if (folderCache.has(lower)) {
    return folderCache.get(lower)!;
  }

  // Query Graph API for folder by displayName
  const resp = await graphGet<GraphPagedResponse<GraphMailFolder>>(
    `/me/mailFolders?$filter=displayName eq '${folderNameOrId.replace(/'/g, "''")}'`,
  );

  if (resp.value.length > 0) {
    const folderId = resp.value[0].id;
    folderCache.set(lower, folderId);
    return folderId;
  }

  // Assume it's already a folder ID
  return folderNameOrId;
}

// --- Helper: parse recipients ---

function parseRecipients(str: string): Array<{ emailAddress: { address: string } }> {
  return str
    .split(",")
    .map((s) => s.trim())
    .filter((s) => s.length > 0)
    .map((address) => ({ emailAddress: { address } }));
}

// --- Helper: strip HTML ---

function stripHtml(html: string): string {
  return html
    .replace(/<style[^>]*>[\s\S]*?<\/style>/gi, "")
    .replace(/<script[^>]*>[\s\S]*?<\/script>/gi, "")
    .replace(/<[^>]+>/g, " ")
    .replace(/&nbsp;/g, " ")
    .replace(/&amp;/g, "&")
    .replace(/&lt;/g, "<")
    .replace(/&gt;/g, ">")
    .replace(/&quot;/g, '"')
    .replace(/&#39;/g, "'")
    .replace(/\s+/g, " ")
    .trim();
}

// --- Register all mail tools ---

export function registerMailTools(server: McpServer): void {
  // Tool 1: ms365_list_mail_folders
  server.tool(
    "ms365_list_mail_folders",
    "List all mail folders with unread and total counts.",
    {},
    { readOnlyHint: true },
    async () => {
      try {
        const resp = await graphGet<GraphPagedResponse<GraphMailFolder>>(
          "/me/mailFolders?$top=100",
          true,
        );
        const folders = resp.value.map((f) => ({
          id: f.id,
          name: f.displayName,
          unread_count: f.unreadItemCount,
          total_count: f.totalItemCount,
        }));
        return {
          content: [{ type: "text" as const, text: JSON.stringify(folders, null, 2) }],
        };
      } catch (error) {
        return {
          content: [{ type: "text" as const, text: `Error: ${error}` }],
          isError: true,
        };
      }
    },
  );

  // Tool 2: ms365_search_mail
  server.tool(
    "ms365_search_mail",
    "Search emails by query, sender, folder, date range, and other filters.",
    {
      query: z.string().optional().describe("Free-text search (subject, body, from)"),
      from: z.string().optional().describe("Filter by sender email"),
      folder: z.string().default("Inbox").describe("Folder name or ID"),
      days: z.number().default(30).describe("How far back to search"),
      unread_only: z.boolean().default(false).describe("Only return unread emails"),
      limit: z.number().default(20).describe("Max results (max 50)"),
      has_attachments: z.boolean().optional().describe("Filter on attachments"),
    },
    { readOnlyHint: true },
    async (params) => {
      try {
        const resolvedFolder = await resolveFolder(params.folder);
        const effectiveLimit = Math.min(params.limit, 50);

        const daysAgo = new Date();
        daysAgo.setDate(daysAgo.getDate() - params.days);
        const isoDate = daysAgo.toISOString();

        const queryParams: string[] = [];
        queryParams.push(
          "$select=id,subject,from,toRecipients,receivedDateTime,bodyPreview,isRead,hasAttachments,importance",
        );
        queryParams.push(`$top=${effectiveLimit}`);

        let extraHeaders: Record<string, string> | undefined;

        if (params.query) {
          // $search cannot be combined with $filter on the messages endpoint
          queryParams.push(`$search="${params.query}"`);
          queryParams.push("$count=true");
          extraHeaders = { ConsistencyLevel: "eventual" };
        } else {
          // $filter and $orderby only when NOT using $search
          const filterClauses: string[] = [];
          filterClauses.push(`receivedDateTime ge ${isoDate}`);
          if (params.unread_only) filterClauses.push("isRead eq false");
          if (params.has_attachments !== undefined)
            filterClauses.push(`hasAttachments eq ${params.has_attachments}`);
          if (params.from)
            filterClauses.push(`from/emailAddress/address eq '${params.from.replace(/'/g, "''")}'`);

          if (filterClauses.length > 0) {
            queryParams.push(`$filter=${filterClauses.join(" and ")}`);
          }
          queryParams.push("$orderby=receivedDateTime desc");
        }

        const endpoint = `/me/mailFolders/${resolvedFolder}/messages?${queryParams.join("&")}`;
        const resp = await graphGet<GraphPagedResponse<GraphMessage>>(
          endpoint,
          false,
          extraHeaders,
        );

        const messages = resp.value.map((m) => ({
          id: m.id,
          subject: m.subject,
          from: m.from?.emailAddress?.address,
          to: m.toRecipients?.map((r) => r.emailAddress.address),
          date: m.receivedDateTime,
          preview: m.bodyPreview?.substring(0, 200),
          unread: !m.isRead,
          has_attachments: m.hasAttachments,
          importance: m.importance,
        }));

        return {
          content: [{ type: "text" as const, text: JSON.stringify(messages, null, 2) }],
        };
      } catch (error) {
        return {
          content: [{ type: "text" as const, text: `Error: ${error}` }],
          isError: true,
        };
      }
    },
  );

  // Tool 3: ms365_get_mail
  server.tool(
    "ms365_get_mail",
    "Get the full content of an email by ID.",
    {
      id: z.string().describe("Message ID"),
    },
    { readOnlyHint: true },
    async (params) => {
      try {
        const msg = await graphGet<GraphMessage>(
          `/me/messages/${params.id}?$select=id,subject,from,toRecipients,ccRecipients,receivedDateTime,body,isRead,hasAttachments,importance,categories&$expand=attachments($select=id,name,size,contentType)`,
        );

        const bodyText =
          msg.body?.contentType === "text"
            ? msg.body.content
            : msg.body
              ? stripHtml(msg.body.content)
              : undefined;
        const bodyHtml = msg.body?.contentType === "html" ? msg.body.content : undefined;

        const result = {
          id: msg.id,
          subject: msg.subject,
          from: msg.from?.emailAddress?.address,
          to: msg.toRecipients?.map((r) => r.emailAddress.address),
          cc: msg.ccRecipients?.map((r) => r.emailAddress.address),
          date: msg.receivedDateTime,
          unread: !msg.isRead,
          body_text: bodyText,
          body_html: bodyHtml,
          attachments: msg.attachments?.map((a) => ({
            name: a.name,
            size: a.size,
            content_type: a.contentType,
          })),
          importance: msg.importance,
          categories: msg.categories,
        };

        return {
          content: [{ type: "text" as const, text: JSON.stringify(result, null, 2) }],
        };
      } catch (error) {
        return {
          content: [{ type: "text" as const, text: `Error: ${error}` }],
          isError: true,
        };
      }
    },
  );

  // Tool 4: ms365_send_mail
  server.tool(
    "ms365_send_mail",
    "Send an email. IMPORTANT: ALWAYS draft first using ms365_draft_mail and present to user for approval before sending, UNLESS user explicitly says to send directly.",
    {
      to: z.string().describe("Comma-separated recipient email addresses"),
      subject: z.string().describe("Email subject"),
      body: z.string().describe("Email body (plain text or HTML)"),
      cc: z.string().optional().describe("Comma-separated CC recipients"),
      bcc: z.string().optional().describe("Comma-separated BCC recipients"),
      importance: z
        .enum(["low", "normal", "high"])
        .default("normal")
        .describe("Email importance"),
      reply_to_id: z.string().optional().describe("Message ID to reply to"),
    },
    { readOnlyHint: false, destructiveHint: false },
    async (params) => {
      try {
        if (params.reply_to_id) {
          await graphPost(`/me/messages/${params.reply_to_id}/reply`, {
            comment: params.body,
          });
        } else {
          await graphPost("/me/sendMail", {
            message: {
              subject: params.subject,
              body: { contentType: "Text", content: params.body },
              toRecipients: parseRecipients(params.to),
              ccRecipients: params.cc ? parseRecipients(params.cc) : [],
              bccRecipients: params.bcc ? parseRecipients(params.bcc) : [],
              importance: params.importance,
            },
          });
        }

        const result = {
          message: "Email sent",
          to: params.to,
          subject: params.subject,
        };
        return {
          content: [{ type: "text" as const, text: JSON.stringify(result, null, 2) }],
        };
      } catch (error) {
        return {
          content: [{ type: "text" as const, text: `Error: ${error}` }],
          isError: true,
        };
      }
    },
  );

  // Tool 5: ms365_draft_mail
  server.tool(
    "ms365_draft_mail",
    "Create a draft email in the Drafts folder without sending. Use this to compose emails for review before sending.",
    {
      to: z.string().describe("Comma-separated recipient email addresses"),
      subject: z.string().describe("Email subject"),
      body: z.string().describe("Email body (plain text or HTML)"),
      cc: z.string().optional().describe("Comma-separated CC recipients"),
      bcc: z.string().optional().describe("Comma-separated BCC recipients"),
      importance: z
        .enum(["low", "normal", "high"])
        .default("normal")
        .describe("Email importance"),
      reply_to_id: z.string().optional().describe("Message ID to reply to"),
    },
    { readOnlyHint: false },
    async (params) => {
      try {
        let draft: GraphMessage;

        if (params.reply_to_id) {
          // Create a draft reply, then patch with the desired body/subject
          draft = await graphPost<GraphMessage>(
            `/me/messages/${params.reply_to_id}/createReply`,
            {},
          );
          await graphPatch(`/me/messages/${draft.id}`, {
            subject: params.subject,
            body: { contentType: "Text", content: params.body },
            toRecipients: parseRecipients(params.to),
            ccRecipients: params.cc ? parseRecipients(params.cc) : [],
            bccRecipients: params.bcc ? parseRecipients(params.bcc) : [],
            importance: params.importance,
          });
        } else {
          draft = await graphPost<GraphMessage>("/me/messages", {
            subject: params.subject,
            body: { contentType: "Text", content: params.body },
            toRecipients: parseRecipients(params.to),
            ccRecipients: params.cc ? parseRecipients(params.cc) : [],
            bccRecipients: params.bcc ? parseRecipients(params.bcc) : [],
            importance: params.importance,
          });
        }

        const result = {
          id: draft.id,
          subject: params.subject,
          to: params.to,
          message: "Draft saved in Drafts folder. Review before sending.",
        };
        return {
          content: [{ type: "text" as const, text: JSON.stringify(result, null, 2) }],
        };
      } catch (error) {
        return {
          content: [{ type: "text" as const, text: `Error: ${error}` }],
          isError: true,
        };
      }
    },
  );

  // Tool 6: ms365_move_mail
  server.tool(
    "ms365_move_mail",
    "Move an email to another folder.",
    {
      id: z.string().describe("Message ID"),
      destination_folder: z.string().describe("Destination folder name or ID"),
    },
    { readOnlyHint: false },
    async (params) => {
      try {
        const destinationId = await resolveFolder(params.destination_folder);
        await graphPost(`/me/messages/${params.id}/move`, {
          destinationId,
        });

        const result = {
          id: params.id,
          message: `Moved to ${params.destination_folder}`,
        };
        return {
          content: [{ type: "text" as const, text: JSON.stringify(result, null, 2) }],
        };
      } catch (error) {
        return {
          content: [{ type: "text" as const, text: `Error: ${error}` }],
          isError: true,
        };
      }
    },
  );

  // Tool 7: ms365_mark_read
  server.tool(
    "ms365_mark_read",
    "Mark an email as read or unread.",
    {
      id: z.string().describe("Message ID"),
      is_read: z.boolean().describe("true to mark as read, false to mark as unread"),
    },
    { readOnlyHint: false },
    async (params) => {
      try {
        await graphPatch(`/me/messages/${params.id}`, {
          isRead: params.is_read,
        });

        const result = {
          id: params.id,
          is_read: params.is_read,
          message: `Marked as ${params.is_read ? "read" : "unread"}`,
        };
        return {
          content: [{ type: "text" as const, text: JSON.stringify(result, null, 2) }],
        };
      } catch (error) {
        return {
          content: [{ type: "text" as const, text: `Error: ${error}` }],
          isError: true,
        };
      }
    },
  );
}
