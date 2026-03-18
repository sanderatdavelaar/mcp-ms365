import { z } from "zod";
import type { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { graphGet, graphPost, graphPatch, graphDelete } from "../auth/graph-client.js";

// --- Helper types ---

interface GraphEvent {
  id: string;
  subject: string;
  start: { dateTime: string; timeZone: string };
  end: { dateTime: string; timeZone: string };
  location?: { displayName?: string };
  organizer?: { emailAddress: { address: string; name?: string } };
  attendees?: Array<{
    emailAddress: { address: string; name?: string };
    status: { response: string };
    type: string;
  }>;
  isAllDay: boolean;
  isCancelled: boolean;
  isOnlineMeeting: boolean;
  onlineMeetingUrl?: string;
  bodyPreview?: string;
  body?: { contentType: string; content: string };
  showAs: string;
  importance: string;
  sensitivity: string;
  recurrence?: unknown;
  webLink?: string;
}

interface GraphCalendar {
  id: string;
  name: string;
  color: string;
  isDefaultCalendar: boolean;
  canEdit: boolean;
}

interface GraphPagedResponse<T> {
  value: T[];
  "@odata.nextLink"?: string;
}

// --- Helpers ---

const TIMEZONE = "Europe/Amsterdam";

const CALENDAR_VIEW_SELECT = [
  "id", "subject", "start", "end", "location", "organizer", "attendees",
  "isAllDay", "isCancelled", "isOnlineMeeting", "onlineMeetingUrl",
  "bodyPreview", "showAs", "importance", "recurrence", "webLink",
].join(",");

const EVENT_FULL_SELECT = [
  ...CALENDAR_VIEW_SELECT.split(","), "body", "sensitivity",
].join(",");

function toDateTimeParam(input: string): string {
  // If only a date is provided (no T), append midnight
  return input.includes("T") ? input : `${input}T00:00:00`;
}

function formatEventSummary(e: GraphEvent) {
  return {
    id: e.id,
    subject: e.subject,
    start: e.start,
    end: e.end,
    location: e.location?.displayName || null,
    organizer: e.organizer?.emailAddress?.address || null,
    is_all_day: e.isAllDay,
    is_cancelled: e.isCancelled,
    is_online_meeting: e.isOnlineMeeting,
    online_meeting_url: e.onlineMeetingUrl || null,
    preview: e.bodyPreview?.substring(0, 200) || null,
    show_as: e.showAs,
    importance: e.importance,
    has_recurrence: e.recurrence != null,
    web_link: e.webLink || null,
  };
}

function parseAttendees(str: string): Array<{ emailAddress: { address: string }; type: string }> {
  return str
    .split(",")
    .map((s) => s.trim())
    .filter((s) => s.length > 0)
    .map((address) => ({ emailAddress: { address }, type: "required" }));
}

// --- Register all calendar tools ---

export function registerCalendarTools(server: McpServer): void {
  // Tool 1: ms365_list_calendars
  server.tool(
    "ms365_list_calendars",
    "List all calendars for the current user.",
    {},
    { readOnlyHint: true },
    async () => {
      try {
        const resp = await graphGet<GraphPagedResponse<GraphCalendar>>(
          "/me/calendars?$top=100",
          true,
        );
        const calendars = resp.value.map((c) => ({
          id: c.id,
          name: c.name,
          color: c.color,
          is_default: c.isDefaultCalendar,
          can_edit: c.canEdit,
        }));
        return {
          content: [{ type: "text" as const, text: JSON.stringify(calendars, null, 2) }],
        };
      } catch (error) {
        return {
          content: [{ type: "text" as const, text: `Error: ${error}` }],
          isError: true,
        };
      }
    },
  );

  // Tool 2: ms365_get_events
  server.tool(
    "ms365_get_events",
    "Get calendar events in a date range. Uses calendarView to expand recurring events.",
    {
      start_date: z.string().describe("Start date/datetime in ISO format, e.g. '2026-03-18' or '2026-03-18T09:00:00'"),
      end_date: z.string().describe("End date/datetime in ISO format"),
      calendar_id: z.string().optional().describe("Calendar ID (default: primary calendar)"),
      limit: z.number().default(50).describe("Max number of events to return"),
    },
    { readOnlyHint: true },
    async (params) => {
      try {
        const startDt = toDateTimeParam(params.start_date);
        const endDt = toDateTimeParam(params.end_date);
        const effectiveLimit = Math.min(params.limit, 100);

        const basePath = params.calendar_id
          ? `/me/calendars/${params.calendar_id}/calendarView`
          : "/me/calendarView";

        const queryParams = [
          `startDateTime=${startDt}`,
          `endDateTime=${endDt}`,
          `$select=${CALENDAR_VIEW_SELECT}`,
          `$orderby=start/dateTime`,
          `$top=${effectiveLimit}`,
        ].join("&");

        const resp = await graphGet<GraphPagedResponse<GraphEvent>>(
          `${basePath}?${queryParams}`,
          false,
          { Prefer: `outlook.timezone="${TIMEZONE}"` },
        );

        const events = resp.value.map(formatEventSummary);
        return {
          content: [{ type: "text" as const, text: JSON.stringify(events, null, 2) }],
        };
      } catch (error) {
        return {
          content: [{ type: "text" as const, text: `Error: ${error}` }],
          isError: true,
        };
      }
    },
  );

  // Tool 3: ms365_get_event
  server.tool(
    "ms365_get_event",
    "Get full details of a single calendar event by ID, including body and attendee details.",
    {
      id: z.string().describe("Event ID"),
    },
    { readOnlyHint: true },
    async (params) => {
      try {
        const event = await graphGet<GraphEvent>(
          `/me/events/${params.id}?$select=${EVENT_FULL_SELECT}`,
          false,
          { Prefer: `outlook.timezone="${TIMEZONE}"` },
        );

        const result = {
          id: event.id,
          subject: event.subject,
          start: event.start,
          end: event.end,
          location: event.location?.displayName || null,
          organizer: event.organizer?.emailAddress || null,
          attendees: event.attendees?.map((a) => ({
            email: a.emailAddress.address,
            name: a.emailAddress.name || null,
            response: a.status.response,
            type: a.type,
          })) || [],
          is_all_day: event.isAllDay,
          is_cancelled: event.isCancelled,
          is_online_meeting: event.isOnlineMeeting,
          online_meeting_url: event.onlineMeetingUrl || null,
          body: event.body?.content || null,
          body_type: event.body?.contentType || null,
          show_as: event.showAs,
          importance: event.importance,
          sensitivity: event.sensitivity,
          has_recurrence: event.recurrence != null,
          web_link: event.webLink || null,
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

  // Tool 4: ms365_create_event
  server.tool(
    "ms365_create_event",
    "Create a new calendar event.",
    {
      subject: z.string().describe("Event subject/title"),
      start: z.string().describe("Start datetime in ISO format, e.g. '2026-03-20T14:00:00'"),
      end: z.string().describe("End datetime in ISO format"),
      location: z.string().optional().describe("Event location"),
      body: z.string().optional().describe("Event description/body"),
      attendees: z.string().optional().describe("Comma-separated email addresses of attendees"),
      is_all_day: z.boolean().default(false).describe("All-day event"),
      is_online_meeting: z.boolean().default(false).describe("Create as Teams online meeting"),
      show_as: z.enum(["free", "tentative", "busy", "oof", "workingElsewhere"]).default("busy").describe("Free/busy status"),
      importance: z.enum(["low", "normal", "high"]).default("normal").describe("Event importance"),
      calendar_id: z.string().optional().describe("Calendar ID (default: primary calendar)"),
    },
    { readOnlyHint: false },
    async (params) => {
      try {
        const eventBody: Record<string, unknown> = {
          subject: params.subject,
          start: { dateTime: toDateTimeParam(params.start), timeZone: TIMEZONE },
          end: { dateTime: toDateTimeParam(params.end), timeZone: TIMEZONE },
          isAllDay: params.is_all_day,
          showAs: params.show_as,
          importance: params.importance,
        };

        if (params.location) {
          eventBody.location = { displayName: params.location };
        }

        if (params.body) {
          eventBody.body = { contentType: "Text", content: params.body };
        }

        if (params.attendees) {
          eventBody.attendees = parseAttendees(params.attendees);
        }

        if (params.is_online_meeting) {
          eventBody.isOnlineMeeting = true;
          eventBody.onlineMeetingProvider = "teamsForBusiness";
        }

        const endpoint = params.calendar_id
          ? `/me/calendars/${params.calendar_id}/events`
          : "/me/events";

        const created = await graphPost<GraphEvent>(endpoint, eventBody);

        const result = {
          id: created.id,
          subject: created.subject,
          start: created.start,
          end: created.end,
          location: created.location?.displayName || null,
          web_link: created.webLink || null,
          message: "Event created",
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

  // Tool 5: ms365_update_event
  server.tool(
    "ms365_update_event",
    "Update an existing calendar event.",
    {
      id: z.string().describe("Event ID"),
      subject: z.string().optional().describe("New subject/title"),
      start: z.string().optional().describe("New start datetime in ISO format"),
      end: z.string().optional().describe("New end datetime in ISO format"),
      location: z.string().optional().describe("New location"),
      body: z.string().optional().describe("New description/body"),
      attendees: z.string().optional().describe("Comma-separated email addresses (replaces existing)"),
      is_all_day: z.boolean().optional().describe("All-day event"),
      is_online_meeting: z.boolean().optional().describe("Teams online meeting"),
      show_as: z.enum(["free", "tentative", "busy", "oof", "workingElsewhere"]).optional().describe("Free/busy status"),
      importance: z.enum(["low", "normal", "high"]).optional().describe("Event importance"),
    },
    { readOnlyHint: false },
    async (params) => {
      try {
        const updates: Record<string, unknown> = {};

        if (params.subject !== undefined) updates.subject = params.subject;
        if (params.start !== undefined) updates.start = { dateTime: toDateTimeParam(params.start), timeZone: TIMEZONE };
        if (params.end !== undefined) updates.end = { dateTime: toDateTimeParam(params.end), timeZone: TIMEZONE };
        if (params.location !== undefined) updates.location = { displayName: params.location };
        if (params.body !== undefined) updates.body = { contentType: "Text", content: params.body };
        if (params.attendees !== undefined) updates.attendees = parseAttendees(params.attendees);
        if (params.is_all_day !== undefined) updates.isAllDay = params.is_all_day;
        if (params.show_as !== undefined) updates.showAs = params.show_as;
        if (params.importance !== undefined) updates.importance = params.importance;
        if (params.is_online_meeting !== undefined) {
          updates.isOnlineMeeting = params.is_online_meeting;
          if (params.is_online_meeting) {
            updates.onlineMeetingProvider = "teamsForBusiness";
          }
        }

        const updated = await graphPatch<GraphEvent>(`/me/events/${params.id}`, updates);

        const result = {
          id: updated.id,
          subject: updated.subject,
          start: updated.start,
          end: updated.end,
          message: "Event updated",
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

  // Tool 6: ms365_delete_event
  server.tool(
    "ms365_delete_event",
    "Delete a calendar event.",
    {
      id: z.string().describe("Event ID"),
    },
    { readOnlyHint: false, destructiveHint: true },
    async (params) => {
      try {
        await graphDelete(`/me/events/${params.id}`);

        return {
          content: [{ type: "text" as const, text: JSON.stringify({ message: "Event deleted" }, null, 2) }],
        };
      } catch (error) {
        return {
          content: [{ type: "text" as const, text: `Error: ${error}` }],
          isError: true,
        };
      }
    },
  );

  // Tool 7: ms365_check_conflicts
  server.tool(
    "ms365_check_conflicts",
    "Check if a time slot has conflicts with existing calendar events.",
    {
      start: z.string().describe("Start datetime in ISO format"),
      end: z.string().describe("End datetime in ISO format"),
      calendar_id: z.string().optional().describe("Calendar ID (default: primary calendar)"),
    },
    { readOnlyHint: true },
    async (params) => {
      try {
        const startDt = toDateTimeParam(params.start);
        const endDt = toDateTimeParam(params.end);

        const basePath = params.calendar_id
          ? `/me/calendars/${params.calendar_id}/calendarView`
          : "/me/calendarView";

        const queryParams = [
          `startDateTime=${startDt}`,
          `endDateTime=${endDt}`,
          `$select=id,subject,start,end,showAs,isCancelled`,
          `$orderby=start/dateTime`,
          `$top=50`,
        ].join("&");

        const resp = await graphGet<GraphPagedResponse<GraphEvent>>(
          `${basePath}?${queryParams}`,
          false,
          { Prefer: `outlook.timezone="${TIMEZONE}"` },
        );

        const conflicts = resp.value
          .filter((e) => !e.isCancelled && e.showAs !== "free")
          .map((e) => ({
            id: e.id,
            subject: e.subject,
            start: e.start,
            end: e.end,
            show_as: e.showAs,
          }));

        const result = {
          has_conflicts: conflicts.length > 0,
          conflicts,
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
