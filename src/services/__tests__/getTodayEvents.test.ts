import { getTodayEvents } from '../graphData';
import type { MSGraphClientV3 } from '@microsoft/sp-http';
import { ICalendarEvent } from '../types';

interface IRawResponse {
  value: unknown[];
}

/** A fluent MSGraphClientV3 stub: every builder call returns the builder; get() resolves the canned response. */
function mockClient(response: IRawResponse): MSGraphClientV3 {
  const builder: Record<string, unknown> = {};
  const chain = (): unknown => builder;
  builder.api = jest.fn(chain);
  builder.query = jest.fn(chain);
  builder.select = jest.fn(chain);
  builder.orderby = jest.fn(chain);
  builder.top = jest.fn(chain);
  builder.header = jest.fn(chain);
  builder.get = jest.fn(() => Promise.resolve(response));
  return builder as unknown as MSGraphClientV3;
}

describe('getTodayEvents', () => {
  beforeEach(() => {
    window.sessionStorage.clear();
    jest.useFakeTimers();
    jest.setSystemTime(new Date('2026-07-15T09:00:00Z')); // summer / BST
  });
  afterEach(() => jest.useRealTimers());

  it('maps a BST wall-clock event to the correct UTC instant (locks the +Z bug)', async () => {
    // calendarView with `Prefer: outlook.timezone` returns London wall-clock with
    // no offset. 10:00 London in July is 09:00Z; the naive `new Date(dt + "Z")`
    // would wrongly yield 10:00Z (an hour early).
    const client = mockClient({
      value: [
        {
          subject: 'Stand-up',
          start: { dateTime: '2026-07-15T10:00:00.0000000' },
          end: { dateTime: '2026-07-15T10:30:00.0000000' },
          isAllDay: false,
          isCancelled: false,
          showAs: 'busy',
          onlineMeeting: { joinUrl: 'https://teams.example/join' },
          location: { displayName: 'Room 1' }
        }
      ]
    });

    const events: ICalendarEvent[] = await getTodayEvents(client);

    expect(events).toHaveLength(1);
    expect(events[0].start.toISOString()).toBe('2026-07-15T09:00:00.000Z');
    expect(events[0].end.toISOString()).toBe('2026-07-15T09:30:00.000Z');
    expect(events[0].subject).toBe('Stand-up');
    expect(events[0].joinUrl).toBe('https://teams.example/join');
    expect(events[0].location).toBe('Room 1');
    expect(events[0].isCancelled).toBe(false);
  });

  it('caches the result and serves rehydrated Dates without a second Graph call', async () => {
    const client = mockClient({
      value: [
        {
          subject: 'Cached',
          start: { dateTime: '2026-07-15T14:00:00.0000000' },
          end: { dateTime: '2026-07-15T15:00:00.0000000' },
          isAllDay: false,
          isCancelled: false,
          showAs: 'busy'
        }
      ]
    });
    await getTodayEvents(client);
    expect(window.sessionStorage.getItem('hero:events')).not.toBeNull();

    const client2 = mockClient({ value: [] }); // would blank the result if it were called
    const events = await getTodayEvents(client2);

    expect(client2.get as jest.Mock).not.toHaveBeenCalled();
    expect(events[0].subject).toBe('Cached');
    expect(events[0].start instanceof Date).toBe(true);
    expect(events[0].start.toISOString()).toBe('2026-07-15T13:00:00.000Z'); // 14:00 London BST
  });

  it('falls back to an empty array when the Graph call fails', async () => {
    const client = mockClient({ value: [] });
    (client.get as jest.Mock).mockImplementationOnce(() => Promise.reject(new Error('boom')));
    const events = await getTodayEvents(client);
    expect(events).toEqual([]);
  });
});
