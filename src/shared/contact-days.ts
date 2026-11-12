/**
 * Contactdagen: per adres per dag (Europe/Amsterdam) hoeveel er binnenkwam en
 * hoeveel er uitging. Alleen tellingen, nooit inhoud. Dit bestand bevat het
 * rekenwerk zonder netwerk; de tool in tools/contact-days.ts haalt de gegevens
 * bij Graph.
 *
 * Regelformaat (het contract met `kairos contactdagen -`):
 *
 *   # contactdagen <kanaal> <van>..<tot> regels=<n>
 *   ["<kanaal>","<adres>","<naam>","JJJJ-MM-DD",<in>,<uit>]
 */

export const CONTACT_DAYS_DEFAULT = 4;
export const CONTACT_DAYS_MAX = 90;
/** Zo ver terug kijken we naar wie de gebruiker zelf mailde: alleen van hen telt inkomende mail. */
export const KNOWN_LOOKBACK_DAYS = 60;
/** Mail aan meer geadresseerden is een rondzending, geen contact. */
export const MAX_RECIPIENTS = 10;
/** Een afspraak telt met hooguit zoveel anderen; daarboven is het een bijeenkomst. */
export const MAX_OTHER_ATTENDEES = 5;

export type Channel = "mail" | "teams" | "agenda";

const DAY = new Intl.DateTimeFormat("en-CA", {
  timeZone: "Europe/Amsterdam",
  year: "numeric",
  month: "2-digit",
  day: "2-digit",
});

/** De datum (JJJJ-MM-DD) in Europe/Amsterdam van een tijdstip. */
export function amsterdamDay(at: string | Date): string {
  return DAY.format(typeof at === "string" ? new Date(at) : at);
}

function addDays(day: string, n: number): string {
  const d = new Date(`${day}T12:00:00Z`);
  d.setUTCDate(d.getUTCDate() + n);
  return d.toISOString().slice(0, 10);
}

export interface Window {
  /** Eerste en laatste dag van het venster, in Europe/Amsterdam. */
  from: string;
  to: string;
  /** Ruime ondergrens voor filters bij Graph (UTC); de dag zelf bepaalt wat telt. */
  since: string;
}

/** Venster van `days` dagen tot en met vandaag. `days` wordt begrensd op 1..max. */
export function contactWindow(now: Date, days: number, max = CONTACT_DAYS_MAX): Window {
  const n = Math.min(Math.max(Math.trunc(Number(days)) || 1, 1), max);
  const to = amsterdamDay(now);
  const from = addDays(to, -(n - 1));
  // Middernacht in Amsterdam is 22:00 of 23:00 UTC de dag ervoor.
  return { from, to, since: `${addDays(from, -1)}T21:00:00Z` };
}

interface Count {
  name: string;
  in: number;
  out: number;
}

/** Tellingen per (adres, dag) voor één kanaal. */
export class DayCounts {
  private readonly rows = new Map<string, Count>();
  private readonly names = new Map<string, string>();

  constructor(
    private readonly window: Window,
    private readonly self: Set<string>,
  ) {}

  /** Telt één bericht of afspraak. Buiten het venster, zonder adres of van jezelf: niets. */
  add(address: string | undefined | null, name: string | undefined | null, day: string, direction: "in" | "out"): void {
    const adres = normalizeAddress(address);
    if (!adres || this.self.has(adres) || day < this.window.from || day > this.window.to) return;
    const schoon = cleanName(name, adres);
    // De eerste naam die we zien wint; de aanroeper biedt de nieuwste berichten eerst aan.
    if (schoon && !this.names.has(adres)) this.names.set(adres, schoon);
    const key = `${adres}\n${day}`;
    const row = this.rows.get(key) ?? { name: "", in: 0, out: 0 };
    row[direction]++;
    this.rows.set(key, row);
  }

  get size(): number {
    return this.rows.size;
  }

  lines(channel: Channel): string[] {
    return [...this.rows.keys()].sort().map((key) => {
      const [adres, day] = key.split("\n");
      const row = this.rows.get(key)!;
      return JSON.stringify([channel, adres, this.names.get(adres) ?? "", day, row.in, row.out]);
    });
  }
}

/** Kopregel plus regels voor één kanaal. */
export function formatChannel(channel: Channel, window: Window, counts: DayCounts): string {
  return [`# contactdagen ${channel} ${window.from}..${window.to} regels=${counts.size}`, ...counts.lines(channel)].join("\n");
}

export function normalizeAddress(address: string | undefined | null): string {
  const a = String(address ?? "").trim().toLowerCase();
  // Interne Exchange-adressen ("/o=exchangelabs/…") zijn geen adres waar iemand op te vinden is.
  return /^[^\s@/]+@[^\s@/]+$/.test(a) && a.length <= 200 ? a : "";
}

function cleanName(name: string | undefined | null, address: string): string {
  const n = String(name ?? "").replace(/\s+/g, " ").trim().slice(0, 120);
  return n.toLowerCase() === address ? "" : n;
}

/** Afzenders waar geen mens achter zit. */
export function isAutomated(address: string): boolean {
  return /^(no[-_.]?reply|do[-_.]?not[-_.]?reply|mailer-daemon|postmaster|bounces?)([-+._].*)?@/.test(address);
}

// ---------- Mail ----------

export interface MailAddress {
  emailAddress?: { address?: string; name?: string };
}

export interface MailItem {
  sentDateTime?: string;
  receivedDateTime?: string;
  isDraft?: boolean;
  from?: MailAddress;
  toRecipients?: MailAddress[];
}

/**
 * Mail: verzonden telt per geadresseerde (alleen Aan, niet Cc), ontvangen telt
 * alleen als de gebruiker zelf in Aan staat en de afzender iemand is aan wie hij
 * de afgelopen tijd zelf schreef. Zo blijven nieuwsbrieven en meldingen buiten
 * het register. Bied beide lijsten met de nieuwste eerst aan.
 */
export function countMail(sent: MailItem[], received: MailItem[], self: Set<string>, window: Window): DayCounts {
  const counts = new DayCounts(window, self);
  const known = new Set<string>();
  const recipients = (m: MailItem) => (m.toRecipients ?? []).map((r) => ({ adres: normalizeAddress(r.emailAddress?.address), naam: r.emailAddress?.name }));

  for (const m of sent) {
    if (m.isDraft || !m.sentDateTime) continue;
    const aan = recipients(m).filter((r) => r.adres && !self.has(r.adres) && !isAutomated(r.adres));
    if (aan.length > MAX_RECIPIENTS) continue;
    const day = amsterdamDay(m.sentDateTime);
    for (const r of aan) {
      known.add(r.adres);
      counts.add(r.adres, r.naam, day, "out");
    }
  }
  for (const m of received) {
    if (m.isDraft || !m.receivedDateTime) continue;
    const van = normalizeAddress(m.from?.emailAddress?.address);
    if (!van || self.has(van) || !known.has(van)) continue;
    const aan = recipients(m);
    if (aan.length > MAX_RECIPIENTS || !aan.some((r) => self.has(r.adres))) continue;
    counts.add(van, m.from?.emailAddress?.name, amsterdamDay(m.receivedDateTime), "in");
  }
  return counts;
}

// ---------- Teams ----------

export interface ChatMember {
  userId?: string;
  email?: string | null;
  displayName?: string | null;
}

export interface ChatMessageItem {
  messageType?: string;
  createdDateTime?: string;
  deletedDateTime?: string | null;
  from?: { user?: { id?: string } | null } | null;
}

/** De ander in een één-op-één-chat, of null als die geen e-mailadres heeft (bv. een bot). */
export function otherMember(members: ChatMember[], myId: string): { address: string; name: string } | null {
  const ander = members.find((m) => m.userId && m.userId !== myId);
  const address = normalizeAddress(ander?.email);
  return address ? { address, name: ander?.displayName ?? "" } : null;
}

/** Telt de echte berichten van één chat (geen systeemmeldingen, geen verwijderde berichten). */
export function countChat(counts: DayCounts, other: { address: string; name: string }, messages: ChatMessageItem[], myId: string): void {
  for (const m of messages) {
    if (m.messageType !== "message" || m.deletedDateTime || !m.createdDateTime || !m.from?.user?.id) continue;
    counts.add(other.address, other.name, amsterdamDay(m.createdDateTime), m.from.user.id === myId ? "out" : "in");
  }
}

// ---------- Agenda ----------

export interface EventItem {
  isCancelled?: boolean;
  isAllDay?: boolean;
  /** Met de header Prefer: outlook.timezone="Europe/Amsterdam" is dit lokale tijd. */
  start?: { dateTime?: string };
  organizer?: MailAddress;
  responseStatus?: { response?: string };
  attendees?: Array<MailAddress & { type?: string; status?: { response?: string } }>;
}

/**
 * Afspraken met een klein gezelschap: elke andere deelnemer krijgt die dag een
 * telling. `uit` als de gebruiker zelf uitnodigde, anders `in`. Geannuleerd,
 * de hele dag, afgewezen (door de gebruiker of door de ander) telt niet.
 */
export function countEvents(events: EventItem[], self: Set<string>, window: Window): DayCounts {
  const counts = new DayCounts(window, self);
  for (const e of events) {
    const day = e.start?.dateTime?.slice(0, 10) ?? "";
    if (e.isCancelled || e.isAllDay || !/^\d{4}-\d{2}-\d{2}$/.test(day) || e.responseStatus?.response === "declined") continue;
    const anderen = (e.attendees ?? [])
      .filter((a) => a.type !== "resource")
      .map((a) => ({ adres: normalizeAddress(a.emailAddress?.address), naam: a.emailAddress?.name, weg: a.status?.response === "declined" }))
      .filter((a) => a.adres && !self.has(a.adres));
    if (!anderen.length || anderen.length > MAX_OTHER_ATTENDEES) continue;
    const richting = self.has(normalizeAddress(e.organizer?.emailAddress?.address)) ? "out" : "in";
    for (const a of anderen) if (!a.weg) counts.add(a.adres, a.naam, day, richting);
  }
  return counts;
}
