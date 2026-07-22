import { MSGraphClientV3 } from '@microsoft/sp-http';
import { cacheGet, cacheSet } from './cache';
import { ICalendarEvent, IFreeSlot, IMailMessage } from './types';

/**
 * Delegated Microsoft Graph reads and pure derivations for the intranet
 * homepage web parts.
 *
 * The library never acquires its own client: every call takes an
 * `MSGraphClientV3` that the consuming web part passes in, so all Graph access
 * runs under the signed-in user's delegated permissions. Every network call is
 * cached in sessionStorage and every failure returns a safe empty value.
 *
 * Calendar reads are requested in GMT Standard Time and converted to correct
 * UTC-anchored Dates; all day-boundary and working-day maths is done in
 * Europe/London so behaviour is independent of the browser's timezone.
 */

/** Cache lifetime for the shared calendar/mail/task reads (5 minutes). */
const CACHE_TTL_MS: number = 5 * 60 * 1000;

/** sessionStorage keys, shared by every consumer of this library. */
const EVENTS_CACHE_KEY: string = 'hero:events';
const MAIL_CACHE_KEY: string = 'hero:mailcount';
const MAIL_LIST_CACHE_KEY: string = 'hero:mail';
const TASKS_CACHE_KEY: string = 'hero:taskcount';

/** Largest page of important mail the list reader will request. */
const MAIL_LIST_MAX: number = 25;

/** Outlook timezone requested for all calendar reads. */
const OUTLOOK_TIME_ZONE: string = 'GMT Standard Time';

/** IANA equivalent of OUTLOOK_TIME_ZONE, used for browser-side date maths. */
const LONDON_TIME_ZONE: string = 'Europe/London';

/** Working-day window (Europe/London) used by the free-slot derivation. */
const WORKDAY_START_HOUR: number = 8;
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
  isCancelled?: boolean;
  showAs?: string;
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

/** Subset of a Graph message this library reads for the important-mail list. */
interface IRawMessage {
  id?: string;
  subject?: string;
  from?: { emailAddress?: { name?: string; address?: string } } | null;
  receivedDateTime?: string;
  importance?: string;
  flag?: { flagStatus?: string } | null;
  isRead?: boolean;
  webLink?: string;
  bodyPreview?: string;
}

// --- Europe/London helpers ----------------------------------------------

function pad2(value: number): string {
  return value < 10 ? `0${value}` : String(value);
}

/** Offset of London wall-clock time from UTC, in minutes (0 in GMT, 60 in BST). */
function londonOffsetMinutes(at: Date): number {
  const utcMs: number = new Date(at.toLocaleString('en-US', { timeZone: 'UTC' })).getTime();
  const londonMs: number = new Date(at.toLocaleString('en-US', { timeZone: LONDON_TIME_ZONE })).getTime();
  return Math.round((londonMs - utcMs) / 60000);
}

/** Today's date in London as YYYY-MM-DD (en-CA formats ISO-style). */
function londonDateISO(at: Date): string {
  return at.toLocaleDateString('en-CA', { timeZone: LONDON_TIME_ZONE });
}

/**
 * Convert a London wall-clock stamp ('YYYY-MM-DDTHH:MM:SS', no offset) to a
 * correct UTC-anchored Date, using the given London offset. calendarView with
 * `Prefer: outlook.timezone` returns wall-clock times without an offset, so
 * this is what turns them into real instants (and fixes the naive `+ 'Z'`
 * that is an hour out during British Summer Time).
 */
function londonWallToUtc(wallClock: string, offsetMinutes: number): Date {
  return new Date(new Date(`${wallClock}Z`).getTime() - offsetMinutes * 60000);
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
    const offsetMinutes: number = londonOffsetMinutes(now);
    const date: string = londonDateISO(now);

    // calendarView interprets offset-less boundaries as UTC, so convert the
    // London day boundaries to UTC instants before querying.
    const startUtc: string = londonWallToUtc(`${date}T00:00:00`, offsetMinutes).toISOString();
    const endUtc: string = londonWallToUtc(`${date}T23:59:59`, offsetMinutes).toISOString();

    const res: IGraphCollection<IRawEvent> = await client
      .api('/me/calendarView')
      .query({ startDateTime: startUtc, endDateTime: endUtc })
      .select('subject,start,end,isAllDay,isCancelled,showAs,onlineMeeting,location')
      .orderby('start/dateTime')
      .top(50)
      .header('Prefer', `outlook.timezone="${OUTLOOK_TIME_ZONE}"`)
      .get();

    const events: ICalendarEvent[] = (res.value ?? []).map((e: IRawEvent): ICalendarEvent => {
      const startWall: string = e.start ? e.start.dateTime.substring(0, 19) : '';
      const endWall: string = e.end ? e.end.dateTime.substring(0, 19) : '';
      return {
        subject: e.subject ? e.subject : '(no subject)',
        start: londonWallToUtc(startWall, offsetMinutes),
        end: londonWallToUtc(endWall, offsetMinutes),
        isAllDay: !!e.isAllDay,
        isCancelled: !!e.isCancelled,
        showAs: e.showAs,
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

// --- important mail list -------------------------------------------------

/**
 * The high-importance or flagged inbox messages themselves — the list behind
 * `getImportantMailCount`. Same advanced-query filter (`ConsistencyLevel:
 * eventual` + `$count`), newest first. `top` is clamped to [1, 25]; the result
 * is cached under `hero:mail` (5-min TTL) and every failure returns an empty
 * array so a mail hiccup never breaks the homepage.
 *
 * Ordering is done client-side (Graph rejects combining this `$filter` with a
 * `$orderby` on a different property), so the call stays a single request.
 */
export async function getImportantMail(client: MSGraphClientV3, top: number = 10): Promise<IMailMessage[]> {
  const cached: IMailMessage[] | undefined = cacheGet<IMailMessage[]>(MAIL_LIST_CACHE_KEY, CACHE_TTL_MS);
  if (cached) {
    // `received` is serialised to an ISO string in sessionStorage — rehydrate it.
    return cached.map((m: IMailMessage): IMailMessage => {
      return { ...m, received: new Date(m.received as unknown as string) };
    });
  }

  try {
    const limit: number = Math.min(Math.max(Math.floor(top), 1), MAIL_LIST_MAX);

    const res: IGraphCollection<IRawMessage> = await client
      .api('/me/mailFolders/inbox/messages')
      .filter("importance eq 'high' or flag/flagStatus eq 'flagged'")
      .select('subject,from,receivedDateTime,importance,flag,isRead,webLink,bodyPreview')
      .header('ConsistencyLevel', 'eventual')
      .count(true)
      .top(limit)
      .get();

    const messages: IMailMessage[] = (res.value ?? []).map((m: IRawMessage): IMailMessage => {
      const emailAddress: { name?: string; address?: string } | undefined = m.from ? m.from.emailAddress : undefined;
      const name: string | undefined = emailAddress && emailAddress.name ? emailAddress.name : undefined;
      const address: string | undefined = emailAddress && emailAddress.address ? emailAddress.address : undefined;
      return {
        id: m.id ? m.id : '',
        subject: m.subject ? m.subject : '(no subject)',
        from: name ?? address ?? 'Unknown sender',
        fromAddress: address,
        received: m.receivedDateTime ? new Date(m.receivedDateTime) : new Date(0),
        importance: m.importance ? m.importance : 'normal',
        isFlagged: !!(m.flag && m.flag.flagStatus === 'flagged'),
        isRead: m.isRead !== false,
        webLink: m.webLink,
        preview: m.bodyPreview
      };
    });

    // Newest first, independent of the order Graph happened to return.
    messages.sort((a: IMailMessage, b: IMailMessage): number => b.received.getTime() - a.received.getTime());

    cacheSet(MAIL_LIST_CACHE_KEY, messages);
    return messages;
  } catch {
    return []; // never let a mail failure break the homepage
  }
}

// --- tasks due today -----------------------------------------------------

/**
 * Count of open To Do tasks due today or overdue, across the user's task
 * lists (compared in Europe/London). This is the most expensive element (one
 * call per list), so consumers typically gate it behind a web-part property to
 * protect first paint.
 */
export async function getTasksDueTodayCount(client: MSGraphClientV3): Promise<number> {
  const cached: number | undefined = cacheGet<number>(TASKS_CACHE_KEY, CACHE_TTL_MS);
  if (cached !== undefined) {
    return cached;
  }

  try {
    const today: string = londonDateISO(new Date());
    const lists: IGraphCollection<IRawTaskList> = await client.api('/me/todo/lists').select('id').get();

    const perListCounts: number[] = await Promise.all(
      (lists.value ?? []).map((list: IRawTaskList): Promise<number> => {
        return client
          .api(`/me/todo/lists/${list.id}/tasks`)
          .select('dueDateTime,status')
          .top(100)
          .get()
          .then((tasks: IGraphCollection<IRawTask>): number => {
            return (tasks.value ?? []).filter((task: IRawTask): boolean => {
              return (
                task.status !== 'completed' &&
                !!task.dueDateTime &&
                !!task.dueDateTime.dateTime &&
                task.dueDateTime.dateTime.substring(0, 10) <= today
              );
            }).length;
          })
          .catch((): number => 0);
      })
    );

    let due: number = 0;
    for (const count of perListCounts) {
      due += count;
    }

    cacheSet(TASKS_CACHE_KEY, due);
    return due;
  } catch {
    return 0; // never let a task-count failure break the homepage
  }
}

// --- derivations (no network) --------------------------------------------

/** The first timed, non-cancelled meeting still to finish. */
export function findNextMeeting(events: ICalendarEvent[]): ICalendarEvent | undefined {
  const now: number = Date.now();
  const meetings: ICalendarEvent[] = events
    .filter((e: ICalendarEvent): boolean => !e.isAllDay && !e.isCancelled)
    .sort((a: ICalendarEvent, b: ICalendarEvent): number => a.start.getTime() - b.start.getTime());

  for (const e of meetings) {
    if (e.end.getTime() > now) {
      return e;
    }
  }
  return undefined;
}

/** Number of timed, non-cancelled meetings today, past ones included. */
export function countMeetingsToday(events: ICalendarEvent[]): number {
  return events.filter((e: ICalendarEvent): boolean => !e.isAllDay && !e.isCancelled).length;
}

/**
 * The first gap of at least 30 minutes between now (or the start of the
 * working day) and the end of the working day, in Europe/London. Events that
 * are all-day, cancelled, or shown as 'free' do not count as busy.
 */
export function findNextFreeSlot(events: ICalendarEvent[]): IFreeSlot | undefined {
  const now: Date = new Date();
  const offsetMinutes: number = londonOffsetMinutes(now);
  const date: string = londonDateISO(now);
  const dayStart: Date = londonWallToUtc(`${date}T${pad2(WORKDAY_START_HOUR)}:00:00`, offsetMinutes);
  const dayEnd: Date = londonWallToUtc(`${date}T${pad2(WORKDAY_END_HOUR)}:00:00`, offsetMinutes);

  let cursor: number = Math.max(now.getTime(), dayStart.getTime());
  if (cursor >= dayEnd.getTime()) {
    return undefined;
  }

  const busy: number[][] = [];
  for (const e of events) {
    if (e.isAllDay || e.isCancelled || e.showAs === 'free') {
      continue;
    }
    // Clamp events to today's working-day window.
    const start: number = Math.max(e.start.getTime(), dayStart.getTime());
    const end: number = Math.min(e.end.getTime(), dayEnd.getTime());
    if (end > start) {
      busy.push([start, end]);
    }
  }
  busy.sort((a: number[], b: number[]): number => a[0] - b[0]);

  for (const interval of busy) {
    const gapEnd: number = Math.min(interval[0], dayEnd.getTime());
    if (gapEnd - cursor >= FREE_SLOT_MINIMUM_MS) {
      return { start: new Date(cursor), end: new Date(gapEnd) };
    }
    if (interval[1] > cursor) {
      cursor = interval[1];
    }
    if (cursor >= dayEnd.getTime()) {
      return undefined;
    }
  }

  if (dayEnd.getTime() - cursor >= FREE_SLOT_MINIMUM_MS) {
    return { start: new Date(cursor), end: new Date(dayEnd.getTime()) };
  }
  return undefined;
}
