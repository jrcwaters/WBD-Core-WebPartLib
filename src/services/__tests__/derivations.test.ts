import { findNextMeeting, countMeetingsToday, findNextFreeSlot } from '../graphData';
import { ICalendarEvent } from '../types';
import { formatLondonTime } from '../format';

/** Build a calendar event, overriding only the fields a test cares about. */
function ev(partial: Partial<ICalendarEvent>): ICalendarEvent {
  return {
    subject: 'Meeting',
    start: new Date('2026-07-15T09:00:00Z'),
    end: new Date('2026-07-15T09:30:00Z'),
    isAllDay: false,
    isCancelled: false,
    ...partial
  };
}

describe('findNextMeeting', () => {
  beforeEach(() => {
    jest.useFakeTimers();
    jest.setSystemTime(new Date('2026-07-15T09:15:00Z'));
  });
  afterEach(() => jest.useRealTimers());

  it('returns the first meeting still to finish, ignoring finished ones and honouring order', () => {
    const past = ev({ subject: 'Past', start: new Date('2026-07-15T08:00:00Z'), end: new Date('2026-07-15T08:30:00Z') });
    const current = ev({ subject: 'Current', start: new Date('2026-07-15T09:00:00Z'), end: new Date('2026-07-15T09:30:00Z') });
    const later = ev({ subject: 'Later', start: new Date('2026-07-15T11:00:00Z'), end: new Date('2026-07-15T11:30:00Z') });
    // Deliberately unsorted input.
    expect(findNextMeeting([later, past, current])?.subject).toBe('Current');
  });

  it('skips all-day and cancelled events', () => {
    const allDay = ev({ subject: 'AllDay', isAllDay: true, start: new Date('2026-07-15T00:00:00Z'), end: new Date('2026-07-16T00:00:00Z') });
    const cancelled = ev({ subject: 'Cancelled', isCancelled: true, start: new Date('2026-07-15T10:00:00Z'), end: new Date('2026-07-15T10:30:00Z') });
    const real = ev({ subject: 'Real', start: new Date('2026-07-15T12:00:00Z'), end: new Date('2026-07-15T12:30:00Z') });
    expect(findNextMeeting([allDay, cancelled, real])?.subject).toBe('Real');
  });

  it('returns undefined when nothing remains today', () => {
    const past = ev({ start: new Date('2026-07-15T08:00:00Z'), end: new Date('2026-07-15T08:30:00Z') });
    expect(findNextMeeting([past])).toBeUndefined();
  });
});

describe('countMeetingsToday', () => {
  it('counts timed, non-cancelled meetings, past ones included', () => {
    const events = [
      ev({}),
      ev({ start: new Date('2026-07-15T08:00:00Z'), end: new Date('2026-07-15T08:30:00Z') }),
      ev({ isAllDay: true }),
      ev({ isCancelled: true })
    ];
    expect(countMeetingsToday(events)).toBe(2);
  });
});

describe('findNextFreeSlot', () => {
  beforeEach(() => {
    jest.useFakeTimers();
    // 09:00Z == 10:00 London (BST). Working day 08:00–18:00 London == 07:00Z–17:00Z.
    jest.setSystemTime(new Date('2026-07-15T09:00:00Z'));
  });
  afterEach(() => jest.useRealTimers());

  it('offers the rest of the working day when the calendar is clear', () => {
    const slot = findNextFreeSlot([]);
    expect(slot).toBeDefined();
    expect(formatLondonTime(slot!.start)).toBe('10:00'); // now
    expect(formatLondonTime(slot!.end)).toBe('18:00'); // end of the working day
  });

  it('returns the gap before the next meeting when it is at least 30 minutes', () => {
    const busy = ev({ start: new Date('2026-07-15T09:30:00Z'), end: new Date('2026-07-15T10:30:00Z') });
    const slot = findNextFreeSlot([busy]);
    expect(formatLondonTime(slot!.start)).toBe('10:00');
    expect(formatLondonTime(slot!.end)).toBe('10:30');
  });

  it('skips a gap shorter than 30 minutes and finds the next one', () => {
    const soon = ev({ start: new Date('2026-07-15T09:20:00Z'), end: new Date('2026-07-15T12:00:00Z') });
    const slot = findNextFreeSlot([soon]);
    // First gap (10:00–10:20 London) is only 20 min; next free is after the meeting.
    expect(formatLondonTime(slot!.start)).toBe('13:00'); // 12:00Z == 13:00 London
    expect(formatLondonTime(slot!.end)).toBe('18:00');
  });

  it('does not treat events shown as "free" as busy', () => {
    const free = ev({ showAs: 'free', start: new Date('2026-07-15T09:30:00Z'), end: new Date('2026-07-15T10:30:00Z') });
    const slot = findNextFreeSlot([free]);
    expect(formatLondonTime(slot!.end)).toBe('18:00');
  });

  it('returns undefined once the working day is over', () => {
    jest.setSystemTime(new Date('2026-07-15T17:30:00Z')); // 18:30 London
    expect(findNextFreeSlot([])).toBeUndefined();
  });
});
