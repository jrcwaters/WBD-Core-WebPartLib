/**
 * Shared types for the @wbd/hub-core calendar/mail data layer.
 *
 * These are the public data shapes the intranet homepage web parts
 * (Welcome Hero and My Meetings) consume from this library.
 */

/**
 * A single calendar event for today, projected from a Microsoft Graph
 * calendarView response.
 *
 * Renamed from the hero's original `IHeroEvent`; the hero now imports
 * `ICalendarEvent` from this library.
 */
export interface ICalendarEvent {
  /** Event subject, or '(no subject)' when Graph returns none. */
  subject: string;
  /** Start of the event as a UTC-anchored Date. */
  start: Date;
  /** End of the event as a UTC-anchored Date. */
  end: Date;
  /** True for all-day events, which are excluded from meeting derivations. */
  isAllDay: boolean;
  /** Teams/online meeting join URL, when the event has one. */
  joinUrl?: string;
  /** Location display name, when the event has one. */
  location?: string;
}

/** A free gap in the working day, bounded by two instants. */
export interface IFreeSlot {
  /** Start of the free slot. */
  start: Date;
  /** End of the free slot. */
  end: Date;
}
