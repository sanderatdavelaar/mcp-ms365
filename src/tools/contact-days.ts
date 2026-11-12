import { z } from "zod";
import type { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { graphGet, graphGetPages } from "../auth/graph-client.js";
import { gid } from "../shared/graph-id.js";
import {
  CONTACT_DAYS_DEFAULT,
  CONTACT_DAYS_MAX,
  KNOWN_LOOKBACK_DAYS,
  type Channel,
  type ChatMember,
  type ChatMessageItem,
  type EventItem,
  type MailItem,
  type Window,
  contactWindow,
  countChat,
  countEvents,
  countMail,
  DayCounts,
  formatChannel,
  normalizeAddress,
  otherMember,
} from "../shared/contact-days.js";

const PAGE = 100;
const MAIL_PAGES = 30;
const CHAT_PAGES = 10;
const CHAT_MESSAGE_PAGES = 10;
const EVENT_PAGES = 15;

interface Me {
  id: string;
  mail?: string | null;
  userPrincipalName?: string | null;
  proxyAddresses?: string[] | null;
}

interface ChatSummary {
  id: string;
  chatType: string;
  lastMessagePreview?: { createdDateTime?: string } | null;
}

/** Wie de ander is in een één-op-één-chat verandert niet: één keer ophalen is genoeg. */
const memberCache = new Map<string, { address: string; name: string } | null>();

async function stap<T>(naam: string, fn: () => Promise<T>): Promise<T> {
  try {
    return await fn();
  } catch (error) {
    throw new Error(`${naam}: ${error instanceof Error ? error.message : error}`);
  }
}

async function mailCounts(self: Set<string>, window: Window, known: Window): Promise<{ counts: DayCounts; note: string }> {
  // Elke stap krijgt zijn naam mee in de fout: een 504 zonder tekst zegt anders niet welk verzoek het was.
  const sent = await stap("verzonden", () =>
    graphGetPages<MailItem>(
      `/me/mailFolders/SentItems/messages?$select=sentDateTime,toRecipients,isDraft&$filter=sentDateTime ge ${known.since}&$orderby=sentDateTime desc&$top=${PAGE}`,
      { maxPages: MAIL_PAGES },
    ),
  );
  const received = await stap("ontvangen", () =>
    graphGetPages<MailItem>(
      `/me/messages?$select=receivedDateTime,from,toRecipients,isDraft&$filter=receivedDateTime ge ${window.since}&$orderby=receivedDateTime desc&$top=${PAGE}`,
      { maxPages: MAIL_PAGES },
    ),
  );
  const afgekapt = [sent.truncated ? "verzonden" : "", received.truncated ? "ontvangen" : ""].filter(Boolean);
  return {
    counts: countMail(sent.value, received.value, self, window),
    note: afgekapt.length ? `# let op mail: ${afgekapt.join(" en ")} afgekapt na ${MAIL_PAGES * PAGE} berichten, de oudste dagen zijn niet compleet` : "",
  };
}

async function teamsCounts(me: Me, self: Set<string>, window: Window): Promise<{ counts: DayCounts; note: string }> {
  const counts = new DayCounts(window, self);
  // Gesorteerd op het laatste bericht: zodra een pagina eindigt met een chat van voor het venster, is de rest ouder.
  const chats = await stap("chats", () =>
    graphGetPages<ChatSummary>(`/me/chats?$expand=lastMessagePreview&$orderby=lastMessagePreview/createdDateTime desc&$top=50`, {
      maxPages: CHAT_PAGES,
      more: (page) => {
        const laatste = page.at(-1)?.lastMessagePreview?.createdDateTime;
        return !laatste || laatste >= window.since;
      },
    }),
  );
  let afgekapt = chats.truncated ? 1 : 0;
  for (const chat of chats.value) {
    const laatste = chat.lastMessagePreview?.createdDateTime;
    if (chat.chatType !== "oneOnOne" || !laatste || laatste < window.since) continue;
    if (!memberCache.has(chat.id)) {
      const members = await graphGet<{ value: ChatMember[] }>(`/chats/${gid(chat.id)}/members`);
      memberCache.set(chat.id, otherMember(members.value ?? [], me.id));
    }
    const ander = memberCache.get(chat.id);
    if (!ander) continue;
    const messages = await graphGetPages<ChatMessageItem>(
      `/me/chats/${gid(chat.id)}/messages?$top=50&$orderby=createdDateTime desc`,
      {
        maxPages: CHAT_MESSAGE_PAGES,
        more: (page) => {
          const oudste = page.at(-1)?.createdDateTime;
          return !oudste || oudste >= window.since;
        },
      },
    );
    if (messages.truncated) afgekapt++;
    countChat(counts, ander, messages.value, me.id);
  }
  return { counts, note: afgekapt ? `# let op teams: ${afgekapt} keer afgekapt, de oudste dagen zijn niet compleet` : "" };
}

async function agendaCounts(self: Set<string>, window: Window, now: Date): Promise<{ counts: DayCounts; note: string }> {
  const events = await graphGetPages<EventItem>(
    `/me/calendarView?startDateTime=${window.since}&endDateTime=${now.toISOString()}&$select=start,isAllDay,isCancelled,organizer,attendees,responseStatus&$orderby=start/dateTime&$top=${PAGE}`,
    { maxPages: EVENT_PAGES, headers: { Prefer: 'outlook.timezone="Europe/Amsterdam"' } },
  );
  return {
    counts: countEvents(events.value, self, window),
    note: events.truncated ? `# let op agenda: afgekapt na ${EVENT_PAGES * PAGE} afspraken, de laatste dagen zijn niet compleet` : "",
  };
}

export function registerContactDaysTools(server: McpServer): void {
  server.tool(
    "ms365_contact_days",
    "Per person per day (Europe/Amsterdam): how many Teams 1:1 messages and small meetings (and, only when asked, work mails) went in and out. Counts only, never content. Mail counts recipients on the To line of sent mail, and received mail only from people you wrote to yourself in the last 60 days. Returns plain text: per channel a header line '# contactdagen <mail|teams|agenda> <from>..<to> regels=<n>' followed by one JSON array per line: [channel, email address, name, date, in, out]. A channel that failed gives '# fout <channel>: …' instead of its header. Pass ALL lines on unchanged to `kairos contactdagen -`.",
    {
      days: z.number().default(CONTACT_DAYS_DEFAULT).describe(`Window in days up to and including today (max ${CONTACT_DAYS_MAX})`),
      channels: z
        .array(z.enum(["mail", "teams", "agenda"]))
        .default(["teams", "agenda"])
        .describe("Which channels to count. Mail is off by default: on 12 Nov 2026 its query made the mailbox time out (504) for mail and calendar; only pass it for a supervised test."),
    },
    { readOnlyHint: true },
    async (params) => {
      try {
        const now = new Date();
        const window = contactWindow(now, params.days);
        const known = contactWindow(now, Math.max(Math.trunc(params.days) || 1, KNOWN_LOOKBACK_DAYS), Math.max(CONTACT_DAYS_MAX, KNOWN_LOOKBACK_DAYS));
        // De aliassen horen erbij (mail aan een alias is ook mail aan jezelf); lukt dat niet, dan zonder.
        const me = await graphGet<Me>("/me?$select=id,mail,userPrincipalName,proxyAddresses").catch(() =>
          graphGet<Me>("/me?$select=id,mail,userPrincipalName"),
        );
        const self = new Set(
          [me.mail, me.userPrincipalName, ...(me.proxyAddresses ?? []).map((p) => p.replace(/^smtp:/i, ""))].map(normalizeAddress).filter(Boolean),
        );
        if (!self.size) throw new Error("eigen adres onbekend");

        const alle: Array<[Channel, () => Promise<{ counts: DayCounts; note: string }>]> = [
          ["mail", () => mailCounts(self, window, known)],
          ["teams", () => teamsCounts(me, self, window)],
          ["agenda", () => agendaCounts(self, window, now)],
        ];
        const kanalen = alle.filter(([kanaal]) => params.channels.includes(kanaal));
        // De kanalen staan los van elkaar: een fout in het ene houdt het andere niet tegen. Na elkaar, niet tegelijk:
        // mail en agenda delen één postvak en Graph verdraagt daar weinig verzoeken naast elkaar.
        const parts: string[] = [];
        for (const [kanaal, tel] of kanalen) {
          try {
            const { counts, note } = await tel();
            parts.push(formatChannel(kanaal, window, counts) + (note ? `\n${note}` : ""));
          } catch (error) {
            // Geen kopregel: Kairos houdt dit kanaal dan op de vorige stand.
            parts.push(`# fout ${kanaal}: ${String(error instanceof Error ? error.message : error).replace(/\s+/g, " ").slice(0, 160)}`);
          }
        }
        return { content: [{ type: "text" as const, text: parts.join("\n") }] };
      } catch (error) {
        return { content: [{ type: "text" as const, text: `Error: ${error}` }], isError: true };
      }
    },
  );
}
