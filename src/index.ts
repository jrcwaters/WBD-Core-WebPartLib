// A file is required to be in the root of the /src directory by the TypeScript compiler.
//
// Public API of @wbd/hub-core — the shared calendar/mail data layer for the WBD
// intranet homepage web parts (Welcome Hero and My Meetings). Everything a
// consumer imports from '@wbd/hub-core' is re-exported here.

// SPFx Library component anchor (carries the stable componentId).
export { HubCore } from './libraries/hubCore/HubCore';

// Types
export { ICalendarEvent, IFreeSlot, IMailMessage } from './services/types';

// sessionStorage cache helpers ({ ts, value } envelope, configurable TTL)
export { cacheGet, cacheSet } from './services/cache';

// Shared display formatting (en-GB London time, duration, relative date)
export { formatLondonTime, formatDuration, formatRelativeDate } from './services/format';

// Delegated Microsoft Graph data access + pure derivations
export {
  getTodayEvents,
  getImportantMailCount,
  getImportantMail,
  getTasksDueTodayCount,
  findNextMeeting,
  countMeetingsToday,
  findNextFreeSlot
} from './services/graphData';
