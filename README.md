# @wbd/hub-core

Shared SPFx **library component** holding the calendar/mail data layer for the WBD
intranet homepage web parts. It is consumed by two separate web-part solutions —
the existing **Welcome Hero** and the future **My Meetings** — and published to the
WBD Azure Artifacts npm feed, with consumers pinning a version.

- **Solution name:** `wbd-hub-core`
- **npm package:** `@wbd/hub-core`
- **Package (.sppkg):** `sharepoint/solution/wbd-hub-core.sppkg`
- **Library componentId (cross-solution contract — never change):**
  `79f58958-95b7-4500-ba34-d006f351fd47`
- **SPFx:** 1.21.1 · **Node:** 22.x

## Scope

This library contains **data access, caching, types and pure derivations only** —
no React and no web-part-specific composition. `getGlanceData` deliberately stays
in the Welcome Hero, not here.

All Microsoft Graph access is **delegated**: every function takes an
`MSGraphClientV3` supplied by the consuming web part. The library never acquires
its own client. Calendar reads use GMT Standard Time; every network failure
returns a safe empty value.

## Delegated Graph permissions

Because SPFx grants API permissions to the shared tenant service principal, the
delegated permission requests are declared **only in this solution**
(`config/package-solution.json`), not in the consumers:

- `Calendars.Read`
- `Mail.Read`
- `Tasks.Read`

## Public API

Everything a consumer imports from `@wbd/hub-core`:

| Export | Kind | Notes |
| --- | --- | --- |
| `ICalendarEvent` | type | Today's calendar event (renamed from the hero's `IHeroEvent`) |
| `IFreeSlot` | type | A free gap in the working day |
| `cacheGet(key, ttlMs)` | function | sessionStorage read, `{ ts, value }` envelope, configurable TTL |
| `cacheSet(key, value)` | function | sessionStorage write, stamps `ts` |
| `getTodayEvents(client)` | async | Single cached `calendarView` call — cache key `hero:events`, 5-min TTL. Shared by both web parts. Excludes cancelled meetings and honours `showAs`; Europe/London day boundaries |
| `getImportantMailCount(client)` | async | High-importance / flagged inbox count |
| `getTasksDueTodayCount(client)` | async | Open To Do tasks due today |
| `findNextMeeting(events)` | function | First timed meeting still to finish |
| `countMeetingsToday(events)` | function | Count of timed meetings today |
| `findNextFreeSlot(events)` | function | First ≥ 30-min gap before end of working day |
| `HubCore` | class | SPFx Library component anchor (carries the componentId) |

## Consuming from a web part

```ts
import {
  getTodayEvents,
  findNextMeeting,
  countMeetingsToday,
  findNextFreeSlot,
  ICalendarEvent
} from '@wbd/hub-core';

const client = await this.context.msGraphClientFactory.getClient('3');
const events: ICalendarEvent[] = await getTodayEvents(client);
const next = findNextMeeting(events);
```

## Build

```bash
npm install
gulp build
gulp bundle --ship
gulp package-solution --ship   # -> sharepoint/solution/wbd-hub-core.sppkg
npm pack                       # -> wbd-hub-core-1.1.0.tgz for the feed
```
