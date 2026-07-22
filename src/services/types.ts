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
  /** Start of the event as a correct UTC-anchored Date. */
  start: Date;
  /** End of the event as a correct UTC-anchored Date. */
  end: Date;
  /** True for all-day events, which are excluded from meeting derivations. */
  isAllDay: boolean;
  /** True for cancelled meetings, which are excluded from meeting derivations. */
  isCancelled: boolean;
  /**
   * Graph free/busy status: 'free' | 'tentative' | 'busy' | 'oof' |
   * 'workingElsewhere' | 'unknown'. Events shown as 'free' do not block the
   * next-free-slot search.
   */
  showAs?: string;
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

/**
 * A single high-importance or flagged inbox message, projected from a Microsoft
 * Graph messages response. This is the list counterpart of
 * `getImportantMailCount`: the Important Emails web part renders these rows,
 * while the Welcome Hero keeps using the count.
 */
export interface IMailMessage {
  /** Graph message id (stable key for React lists). */
  id: string;
  /** Message subject, or '(no subject)' when Graph returns none. */
  subject: string;
  /** Sender display name, falling back to the address, then 'Unknown sender'. */
  from: string;
  /** Sender email address, when Graph returns one. */
  fromAddress?: string;
  /** Received time as a correct UTC-anchored Date. */
  received: Date;
  /** Graph importance: 'low' | 'normal' | 'high'. */
  importance: string;
  /** True when the message is flagged for follow-up (flag/flagStatus === 'flagged'). */
  isFlagged: boolean;
  /** True when the message has been read. */
  isRead: boolean;
  /** Deep link that opens the message in Outlook on the web. */
  webLink?: string;
  /** Short plain-text preview of the body, when Graph returns one. */
  preview?: string;
}
