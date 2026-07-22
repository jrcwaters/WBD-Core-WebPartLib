import { formatLondonTime, formatDuration, formatRelativeDate } from '../format';

describe('formatLondonTime', () => {
  it('formats an instant in GMT (winter)', () => {
    expect(formatLondonTime(new Date('2026-01-15T08:30:00Z'))).toBe('08:30');
  });

  it('formats an instant in BST (summer, +1h) — the season-correct case', () => {
    expect(formatLondonTime(new Date('2026-07-15T08:30:00Z'))).toBe('09:30');
  });

  it('renders midnight as 00:00, not 24:00', () => {
    // 23:00Z in July is 00:00 the next day in London.
    expect(formatLondonTime(new Date('2026-07-15T23:00:00Z'))).toBe('00:00');
  });
});

describe('formatDuration', () => {
  const at = (iso: string): Date => new Date(`2026-07-15T${iso}Z`);

  it('formats a sub-hour duration', () => {
    expect(formatDuration(at('09:00:00'), at('09:30:00'))).toBe('30 mins');
  });

  it('uses the singular for one minute', () => {
    expect(formatDuration(at('09:00:00'), at('09:01:00'))).toBe('1 min');
  });

  it('formats a whole hour', () => {
    expect(formatDuration(at('09:00:00'), at('10:00:00'))).toBe('1 hr');
  });

  it('formats hours and minutes together', () => {
    expect(formatDuration(at('09:00:00'), at('10:30:00'))).toBe('1 hr 30 mins');
  });

  it('never returns a negative duration', () => {
    expect(formatDuration(at('10:00:00'), at('09:00:00'))).toBe('0 mins');
  });
});

describe('formatRelativeDate', () => {
  beforeEach(() => {
    jest.useFakeTimers();
    jest.setSystemTime(new Date('2026-07-15T12:00:00Z'));
  });
  afterEach(() => jest.useRealTimers());

  it('returns an empty string for undefined', () => {
    expect(formatRelativeDate(undefined)).toBe('');
  });

  it('returns an empty string for an unparseable date', () => {
    expect(formatRelativeDate('not-a-date')).toBe('');
  });

  it('returns "Today" for the current day', () => {
    expect(formatRelativeDate('2026-07-15T08:00:00Z')).toBe('Today');
  });

  it('returns "Yesterday" for the previous day', () => {
    expect(formatRelativeDate('2026-07-14T12:00:00Z')).toBe('Yesterday');
  });

  it('returns the weekday name within the past week', () => {
    const threeDaysAgo = '2026-07-12T12:00:00Z';
    const expected = new Date(threeDaysAgo).toLocaleDateString('en-GB', { weekday: 'long' });
    expect(formatRelativeDate(threeDaysAgo)).toBe(expected);
  });

  it('returns a short date beyond a week', () => {
    expect(formatRelativeDate('2026-07-01T12:00:00Z')).toBe('1 Jul');
  });
});
