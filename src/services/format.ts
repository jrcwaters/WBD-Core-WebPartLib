/**
 * Shared display formatting for the WBD Hub homepage web parts.
 *
 * These are pure, presentation-neutral helpers consolidated out of the
 * individual web parts (the hero's `londonTimeHHMM`, My Meetings' `formatUkTime`
 * / `formatDuration`, and News' `formatRelativeDate`) so every web part formats
 * times and dates the same way from one implementation.
 */

/** IANA timezone all Hub time formatting is anchored to. */
const LONDON_TIME_ZONE: string = 'Europe/London';

/**
 * Formatter for en-GB 24-hour wall-clock time in Europe/London. (`hourCycle` is
 * avoided because it is absent from the older DOM lib typings this library
 * compiles against; the `24:xx` that some engines emit for midnight is
 * normalised in `formatLondonTime`.)
 */
const LONDON_HHMM: Intl.DateTimeFormat = new Intl.DateTimeFormat('en-GB', {
  timeZone: LONDON_TIME_ZONE,
  hour: '2-digit',
  minute: '2-digit',
  hour12: false
});

/**
 * Format an instant as `HH:MM` (24-hour) in Europe/London. The input must be a
 * correct UTC-anchored `Date` (as produced by `getTodayEvents`); the London
 * offset — GMT or BST — is applied by the formatter, so it is season-correct.
 */
export function formatLondonTime(date: Date): string {
  const raw: string = LONDON_HHMM.format(date);
  // Some engines render midnight as 24:xx with hour12: false.
  return raw.indexOf('24:') === 0 ? `00:${raw.substring(3)}` : raw;
}

/**
 * Compact British duration between two instants, e.g. `30 mins`, `1 hr`,
 * `1 hr 30 mins`. Never negative.
 */
export function formatDuration(start: Date, end: Date): string {
  const totalMin: number = Math.max(0, Math.round((end.getTime() - start.getTime()) / 60000));
  if (totalMin < 60) {
    return `${totalMin} min${totalMin === 1 ? '' : 's'}`;
  }
  const hours: number = Math.floor(totalMin / 60);
  const mins: number = totalMin % 60;
  const hoursLabel: string = `${hours} hr${hours === 1 ? '' : 's'}`;
  return mins === 0 ? hoursLabel : `${hoursLabel} ${mins} min${mins === 1 ? '' : 's'}`;
}

/**
 * Human-friendly relative date for a byline: `Today`, `Yesterday`, a weekday
 * name for the past week, otherwise a short `12 Jul`-style date. Accepts a
 * `Date`, an ISO string, or `undefined` (which formats to an empty string).
 */
export function formatRelativeDate(input: Date | string | undefined): string {
  if (!input) {
    return '';
  }
  const date: Date = typeof input === 'string' ? new Date(input) : input;
  if (isNaN(date.getTime())) {
    return '';
  }

  const startOfDay = (d: Date): number => new Date(d.getFullYear(), d.getMonth(), d.getDate()).getTime();

  const oneDay: number = 24 * 60 * 60 * 1000;
  const diffDays: number = Math.round((startOfDay(new Date()) - startOfDay(date)) / oneDay);

  if (diffDays <= 0) {
    return 'Today';
  }
  if (diffDays === 1) {
    return 'Yesterday';
  }
  if (diffDays < 7) {
    return date.toLocaleDateString('en-GB', { weekday: 'long' });
  }
  return date.toLocaleDateString('en-GB', { day: 'numeric', month: 'short' });
}
