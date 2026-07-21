import { MSGraphClientV3 } from '@microsoft/sp-http';
import { cacheGet, cacheSet } from './cache';
import { ICalendarEvent, IFreeSlot } from './types';

/**
 * Delegated Microsoft Graph reads and pure derivations for the intranet
 * homepage web parts.
 *
 * The library never acquires its own client: every call takes an
 * `MSGraphClientV3` that the consuming web part passes in, so all Graph access
 * runs under the signed-in user's delegated permissions. Every network call is
 * cached in sessionStorage and every failure returns a safe empty value.
 *
 * All calendar times are requested in GMT Standard Time and handled as
 * UTC-anchored Dates.
 */

/** Cache lifetime for the shared calendar/mail/task reads (5 minutes). */
const CACHE_TTL_MS: number = 5 * 60 * 1000;

/** sessionStorage keys, shared by every consumer of this library. */
const EVENTS_CACHE_KEY: string = 'hero:events';
const MAIL_CACHE_KEY: string = 'hero:mailcount';
const TASKS_CACHE_KEY: string = 'hero:taskcount';

/** Outlook timezone requested for all calendar reads. */
const OUTLOOK_TIME_ZONE: string = 'GMT Standard Time';

/** Working-day window used by the free-slot derivation. */
const WORKDAY_START_HOUR: number = 9;
const WORKDAY_END_HOUR: number = 18;

/** Smallest gap the free-slot derivation will report. */
const FREE_SLOT_MINIMUM_MS: number = 30 * 60 * 1000;

/** Graph collection envelope, plus the optional `$count` annotation. */
interface IGraphCollection<T> {
  value?: T[];
  '@odata.count'?: number;
}

/** Subset of a Graph calendar event this library reads. */
interface IRawEvent {
  subject?: string;
  isAllDay?: boolean;
  start?: { dateTime: string };
  end?: { dateTime: string };
  onlineMeeting?: { joinUrl?: string } | null;
  location?: { displayName?: string } | null;
}

/** Subset of a Graph To Do list this library reads. */
interface IRawTaskList {
  id: string;
}

/** Subset of a Graph To Do task this library reads. */
interface IRawTask {
  status?: string;
  dueDateTime?: { dateTime?: string } | null;
}

// --- calendar ------------------------------------------------------------

/**
 * The single cached calendarView call. One request powers the next meeting,
 * the "N today" count and the next free slot — every homepage web part shares
 * this call and its cache entry (`hero:events`).
 */
export async function getTodayEvents(client: MSGraphClientV3): Promise<ICalendarEvent[]> {
  const cached: ICalendarEvent[] | undefined = cacheGet<ICalendarEvent[]>(EVENTS_CACHE_KEY, CACHE_TTL_MS);
  if (cached) {
    // Dates are serialised to ISO strings in sessionStorage — rehydrate them.
    // (The cached values are strings at runtime despite the Date type.)
    return cached.map((e: ICalendarEvent): ICalendarEvent => {
      return {
        ...e,
        start: new Date(e.start as unknown as string),
        end: new Date(e.end as unknown as string)
      };
    });
  }

  try {
    const now: Date = new Date();
    const endOfDay: Date = new Date(now.getFullYear(), now.getMonth(), now.getDate(), 23, 59, 59);

    const res: IGraphCollection<IRawEvent> = await client
      .api('/me/calendarView')
      .query({ startDateTime: now.toISOString(), endDateTime: endOfDay.toISOString() })
      .select('subject,start,end,isAllDay,onlineMeeting,location')
      .orderby('start/dateTime')
      .top(15)
      .header('Prefer', `outlook.timezone="${OUTLOOK_TIME_ZONE}"`)
      .get();

    const events: ICalendarEvent[] = (res.value ?? []).map((e: IRawEvent): ICalendarEvent => {
      return {
        subject: e.subject ? e.subject : '(no subject)',
        start: new Date((e.start ? e.start.dateTime : '') + 'Z'),
        end: new Date((e.end ? e.end.dateTime : '') + 'Z'),
        isAllDay: !!e.isAllDay,
        joinUrl: e.onlineMeeting ? e.onlineMeeting.joinUrl : undefined,
        location: e.location ? e.location.displayName : undefined
      };
    });

    cacheSet(EVENTS_CACHE_KEY, events);
    return events;
  } catch {
    return []; // never let a calendar failure break the homepage
  }
}

// --- important mail count ------------------------------------------------

/**
 * Count of high-importance or flagged messages in the inbox. `$count` requires
 * the `ConsistencyLevel: eventual` header.
 */
export async function getImportantMailCount(client: MSGraphClientV3): Promise<number> {
  const cached: number | undefined = cacheGet<number>(MAIL_CACHE_KEY, CACHE_TTL_MS);
  if (cached !== undefined) {
    return cached;
  }

  try {
    const res: IGraphCollection<IRawEvent> = await client
      .api('/me/mailFolders/inbox/messages')
      .filter("importance eq 'high' or flag/flagStatus eq 'flagged'")
      .header('ConsistencyLevel', 'eventual')
      .count(true)
      .top(1)
      .get();

    const count: number = res['@odata.count'] ?? 0;
    cacheSet(MAIL_CACHE_KEY, count);
    return count;
  } catch {
    return 0; // never let a mail-count failure break the homepage
  }
}

// --- tasks due today -----------------------------------------------------

/**
 * Count of open To Do tasks due today, across the user's task lists. This is
 * the most expensive element (one call per list), so consumers typically gate
 * it behind a web-part property to protect first paint.
 */
export async function getTasksDueTodayCount(client: MSGraphClientV3): Promise<number> {
  const cached: number | undefined = cacheGet<number>(TASKS_CACHE_KEY, CACHE_TTL_MS);
  if (cached !== undefined) {
    return cached;
  }

  try {
    const lists: IGraphCollection<IRawTaskList> = await client.api('/me/todo/lists').select('id').get();
    const today: Date = new Date();
    let due: number = 0;

    for (const list of lists.value ?? []) {
      const tasks: IGraphCollection<IRawTask> = await client
        .api(`/me/todo/lists/${list.id}/tasks`)
        .filter("status ne 'completed'")
        .select('dueDateTime,status')
        .top(50)
        .get();

      for (const t of tasks.value ?? []) {
        if (!t.dueDateTime || !t.dueDateTime.dateTime) {
          continue;
        }
        const d: Date = new Date(t.dueDateTime.dateTime + 'Z');
        if (
          d.getFullYear() === today.getFullYear() &&
          d.getMonth() === today.getMonth() &&
          d.getDate() === today.getDate()
        ) {
          due += 1;
        }
      }
    }

    cacheSet(TASKS_CACHE_KEY, due);
    return due;
  } catch {
    return 0; // never let a task-count failure break the homepage
  }
}

// --- derivations (no network) --------------------------------------------

/** The first timed meeting still to finish. */
export function findNextMeeting(events: ICalendarEvent[]): ICalendarEvent | undefined {
  const now: number = Date.now();
  for (const e of events) {
    if (!e.isAllDay && e.end.getTime() > now) {
      return e;
    }
  }
  return undefined;
}

/** Number of timed (non-all-day) meetings today, past ones included. */
export function countMeetingsToday(events: ICalendarEvent[]): number {
  return events.filter((e: ICalendarEvent): boolean => !e.isAllDay).length;
}

/**
 * Walk today's timed events and return the first gap of at least 30 minutes
 * between now (or the start of the working day) and the end of the working day.
 */
export function findNextFreeSlot(events: ICalendarEvent[]): IFreeSlot | undefined {
  const now: Date = new Date();
  const dayEnd: Date = new Date(now.getFullYear(), now.getMonth(), now.getDate(), WORKDAY_END_HOUR, 0, 0);
  const dayStart: Date = new Date(now.getFullYear(), now.getMonth(), now.getDate(), WORKDAY_START_HOUR, 0, 0);
  let cursor: Date = new Date(Math.max(now.getTime(), dayStart.getTime()));

  const timed: ICalendarEvent[] = events
    .filter((e: ICalendarEvent): boolean => !e.isAllDay)
    .sort((a: ICalendarEvent, b: ICalendarEvent): number => a.start.getTime() - b.start.getTime());

  for (const e of timed) {
    if (e.start.getTime() - cursor.getTime() >= FREE_SLOT_MINIMUM_MS) {
      return { start: new Date(cursor.getTime()), end: new Date(e.start.getTime()) };
    }
    if (e.end.getTime() > cursor.getTime()) {
      cursor = new Date(e.end.getTime());
    }
  }

  if (dayEnd.getTime() - cursor.getTime() >= FREE_SLOT_MINIMUM_MS) {
    return { start: cursor, end: dayEnd };
  }
  return undefined;
}
