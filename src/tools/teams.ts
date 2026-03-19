import { z } from "zod";
import type { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { graphGet, graphPost } from "../auth/graph-client.js";

// --- Helper types ---

interface GraphChatMessage {
  id: string;
  messageType: string;
  createdDateTime: string;
  deletedDateTime?: string | null;
  from?: {
    user?: { displayName: string; id: string };
    application?: { displayName: string };
  };
  body: { contentType: string; content: string };
  importance: string;
  mentions?: Array<{
    id: number;
    mentionText: string;
    mentioned: { user?: { displayName: string } };
  }>;
  attachments?: Array<{ id: string; contentType: string; name?: string }>;
}

interface GraphChat {
  id: string;
  topic: string | null;
  chatType: string;
  lastUpdatedDateTime: string;
  members?: Array<{ displayName: string; email?: string }>;
}

interface GraphTeam {
  id: string;
  displayName: string;
  description?: string;
}

interface GraphChannel {
  id: string;
  displayName: string;
  description?: string | null;
  membershipType?: string;
}

interface GraphPagedResponse<T> {
  value: T[];
  "@odata.nextLink"?: string;
}

// --- Helpers ---

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

function getSenderName(msg: GraphChatMessage): string {
  if (msg.from?.user?.displayName) return msg.from.user.displayName;
  if (msg.from?.application?.displayName) return msg.from.application.displayName;
  return "Unknown";
}

function formatChatMessage(msg: GraphChatMessage): object {
  const isDeleted = msg.deletedDateTime != null;
  const bodyText = isDeleted
    ? "[This message was deleted]"
    : msg.body.contentType === "html"
      ? stripHtml(msg.body.content)
      : msg.body.content;

  return {
    id: msg.id,
    from: getSenderName(msg),
    date: msg.createdDateTime,
    body_text: bodyText,
    importance: msg.importance,
    message_type: msg.messageType,
    is_deleted: isDeleted,
    mentions:
      msg.mentions?.map((m) => ({
        text: m.mentionText,
        user_name: m.mentioned.user?.displayName,
      })) ?? [],
    has_attachments: (msg.attachments?.length ?? 0) > 0,
  };
}

// --- Register all Teams tools ---

export function registerTeamsTools(server: McpServer): void {
  // Tool 1: ms365_list_chats
  server.tool(
    "ms365_list_chats",
    "List recent 1:1, group, and meeting chats.",
    {
      limit: z.number().default(20).describe("Max results (max 50)"),
    },
    { readOnlyHint: true },
    async (params) => {
      try {
        const effectiveLimit = Math.min(params.limit, 50);
        const resp = await graphGet<GraphPagedResponse<GraphChat>>(
          `/me/chats?$expand=members&$orderby=lastUpdatedDateTime desc&$top=${effectiveLimit}`,
        );

        const chats = resp.value.map((c) => ({
          id: c.id,
          topic: c.topic,
          chat_type: c.chatType,
          last_updated: c.lastUpdatedDateTime,
          members: c.members?.map((m) => ({
            name: m.displayName,
            email: m.email,
          })) ?? [],
        }));

        return {
          content: [{ type: "text" as const, text: JSON.stringify(chats, null, 2) }],
        };
      } catch (error) {
        return {
          content: [{ type: "text" as const, text: `Error: ${error}` }],
          isError: true,
        };
      }
    },
  );

  // Tool 2: ms365_get_chat_messages
  server.tool(
    "ms365_get_chat_messages",
    "Get messages from a 1:1 or group chat.",
    {
      chatId: z.string().describe("Chat ID"),
      limit: z.number().default(20).describe("Max messages (max 50)"),
    },
    { readOnlyHint: true },
    async (params) => {
      try {
        const effectiveLimit = Math.min(params.limit, 50);
        const resp = await graphGet<GraphPagedResponse<GraphChatMessage>>(
          `/me/chats/${params.chatId}/messages?$top=${effectiveLimit}`,
        );

        const messages = resp.value.map(formatChatMessage);
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

  // Tool 3: ms365_send_chat_message
  server.tool(
    "ms365_send_chat_message",
    "Send a message in a 1:1 or group chat. IMPORTANT: ALWAYS draft and present the message to the user for approval before sending, UNLESS user explicitly says to send directly.",
    {
      chatId: z.string().describe("Chat ID"),
      content: z.string().describe("Message content (plain text)"),
      importance: z.enum(["normal", "high", "urgent"]).default("normal").describe("Message importance"),
    },
    { readOnlyHint: false, destructiveHint: false },
    async (params) => {
      try {
        await graphPost(`/me/chats/${params.chatId}/messages`, {
          body: { contentType: "text", content: params.content },
          importance: params.importance,
        });

        return {
          content: [{ type: "text" as const, text: JSON.stringify({ message: "Chat message sent", chatId: params.chatId }, null, 2) }],
        };
      } catch (error) {
        return {
          content: [{ type: "text" as const, text: `Error: ${error}` }],
          isError: true,
        };
      }
    },
  );

  // Tool 4: ms365_list_teams
  server.tool(
    "ms365_list_teams",
    "List all Teams the user has joined.",
    {},
    { readOnlyHint: true },
    async () => {
      try {
        const resp = await graphGet<GraphPagedResponse<GraphTeam>>(
          "/me/joinedTeams",
        );

        const teams = resp.value.map((t) => ({
          id: t.id,
          name: t.displayName,
          description: t.description,
        }));

        return {
          content: [{ type: "text" as const, text: JSON.stringify(teams, null, 2) }],
        };
      } catch (error) {
        return {
          content: [{ type: "text" as const, text: `Error: ${error}` }],
          isError: true,
        };
      }
    },
  );

  // Tool 5: ms365_list_channels
  server.tool(
    "ms365_list_channels",
    "List channels in a Team.",
    {
      teamId: z.string().describe("Team ID"),
    },
    { readOnlyHint: true },
    async (params) => {
      try {
        const resp = await graphGet<GraphPagedResponse<GraphChannel>>(
          `/teams/${params.teamId}/channels`,
        );

        const channels = resp.value.map((c) => ({
          id: c.id,
          name: c.displayName,
          description: c.description,
          membership_type: c.membershipType,
        }));

        return {
          content: [{ type: "text" as const, text: JSON.stringify(channels, null, 2) }],
        };
      } catch (error) {
        return {
          content: [{ type: "text" as const, text: `Error: ${error}` }],
          isError: true,
        };
      }
    },
  );

  // Tool 6: ms365_get_channel_messages
  server.tool(
    "ms365_get_channel_messages",
    "Get messages from a Team channel.",
    {
      teamId: z.string().describe("Team ID"),
      channelId: z.string().describe("Channel ID"),
      limit: z.number().default(20).describe("Max messages (max 50)"),
    },
    { readOnlyHint: true },
    async (params) => {
      try {
        const effectiveLimit = Math.min(params.limit, 50);
        const resp = await graphGet<GraphPagedResponse<GraphChatMessage>>(
          `/teams/${params.teamId}/channels/${params.channelId}/messages?$top=${effectiveLimit}`,
        );

        const messages = resp.value.map(formatChatMessage);
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

  // Tool 7: ms365_send_channel_message
  server.tool(
    "ms365_send_channel_message",
    "Send a message to a Team channel. IMPORTANT: ALWAYS draft and present the message to the user for approval before sending, UNLESS user explicitly says to send directly.",
    {
      teamId: z.string().describe("Team ID"),
      channelId: z.string().describe("Channel ID"),
      content: z.string().describe("Message content (plain text)"),
      importance: z.enum(["normal", "high", "urgent"]).default("normal").describe("Message importance"),
    },
    { readOnlyHint: false, destructiveHint: false },
    async (params) => {
      try {
        await graphPost(`/teams/${params.teamId}/channels/${params.channelId}/messages`, {
          body: { contentType: "text", content: params.content },
          importance: params.importance,
        });

        return {
          content: [{ type: "text" as const, text: JSON.stringify({ message: "Channel message sent", teamId: params.teamId, channelId: params.channelId }, null, 2) }],
        };
      } catch (error) {
        return {
          content: [{ type: "text" as const, text: `Error: ${error}` }],
          isError: true,
        };
      }
    },
  );

  // Tool 8: ms365_reply_to_message
  server.tool(
    "ms365_reply_to_message",
    "Reply to a message in a chat or channel thread. Provide either chatId (for chat replies) OR both teamId+channelId (for channel replies). IMPORTANT: ALWAYS draft and present the reply to the user for approval before sending, UNLESS user explicitly says to send directly.",
    {
      messageId: z.string().describe("Message ID to reply to"),
      content: z.string().describe("Reply content (plain text)"),
      chatId: z.string().optional().describe("Chat ID (for chat replies)"),
      teamId: z.string().optional().describe("Team ID (for channel replies)"),
      channelId: z.string().optional().describe("Channel ID (for channel replies)"),
      importance: z.enum(["normal", "high", "urgent"]).default("normal").describe("Message importance"),
    },
    { readOnlyHint: false, destructiveHint: false },
    async (params) => {
      try {
        const hasChatId = !!params.chatId;
        const hasChannel = !!params.teamId && !!params.channelId;

        if (hasChatId === hasChannel) {
          return {
            content: [{ type: "text" as const, text: "Error: Provide either chatId OR both teamId+channelId, not both or neither." }],
            isError: true,
          };
        }

        const endpoint = hasChatId
          ? `/me/chats/${params.chatId}/messages/${params.messageId}/replies`
          : `/teams/${params.teamId}/channels/${params.channelId}/messages/${params.messageId}/replies`;

        await graphPost(endpoint, {
          body: { contentType: "text", content: params.content },
          importance: params.importance,
        });

        return {
          content: [{ type: "text" as const, text: JSON.stringify({ message: "Reply sent", messageId: params.messageId }, null, 2) }],
        };
      } catch (error) {
        return {
          content: [{ type: "text" as const, text: `Error: ${error}` }],
          isError: true,
        };
      }
    },
  );

  // Tool 9: ms365_get_message_replies
  server.tool(
    "ms365_get_message_replies",
    "Get replies to a message in a chat or channel thread. Provide either chatId (for chat) OR both teamId+channelId (for channel).",
    {
      messageId: z.string().describe("Message ID to get replies for"),
      chatId: z.string().optional().describe("Chat ID (for chat replies)"),
      teamId: z.string().optional().describe("Team ID (for channel replies)"),
      channelId: z.string().optional().describe("Channel ID (for channel replies)"),
      limit: z.number().default(20).describe("Max replies (max 50)"),
    },
    { readOnlyHint: true },
    async (params) => {
      try {
        const hasChatId = !!params.chatId;
        const hasChannel = !!params.teamId && !!params.channelId;

        if (hasChatId === hasChannel) {
          return {
            content: [{ type: "text" as const, text: "Error: Provide either chatId OR both teamId+channelId, not both or neither." }],
            isError: true,
          };
        }

        const effectiveLimit = Math.min(params.limit, 50);
        const endpoint = hasChatId
          ? `/me/chats/${params.chatId}/messages/${params.messageId}/replies?$top=${effectiveLimit}`
          : `/teams/${params.teamId}/channels/${params.channelId}/messages/${params.messageId}/replies?$top=${effectiveLimit}`;

        const resp = await graphGet<GraphPagedResponse<GraphChatMessage>>(endpoint);
        const replies = resp.value.map(formatChatMessage);

        return {
          content: [{ type: "text" as const, text: JSON.stringify(replies, null, 2) }],
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
