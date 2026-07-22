import { cacheGet, cacheSet } from '../cache';

describe('cache', () => {
  beforeEach(() => {
    window.sessionStorage.clear();
    jest.useRealTimers();
  });

  it('round-trips a value within its TTL', () => {
    cacheSet('k', { a: 1, b: 'two' });
    expect(cacheGet<{ a: number; b: string }>('k', 60000)).toEqual({ a: 1, b: 'two' });
  });

  it('treats a value older than the TTL as a miss', () => {
    jest.useFakeTimers();
    jest.setSystemTime(new Date('2026-07-15T10:00:00Z'));
    cacheSet('k', 'v');
    jest.setSystemTime(new Date('2026-07-15T10:06:00Z')); // +6 min, TTL is 5
    expect(cacheGet<string>('k', 5 * 60 * 1000)).toBeUndefined();
  });

  it('still returns a value just inside the TTL boundary', () => {
    jest.useFakeTimers();
    jest.setSystemTime(new Date('2026-07-15T10:00:00Z'));
    cacheSet('k', 'v');
    jest.setSystemTime(new Date('2026-07-15T10:04:59Z')); // +4m59s
    expect(cacheGet<string>('k', 5 * 60 * 1000)).toBe('v');
  });

  it('returns undefined for a missing key', () => {
    expect(cacheGet('missing', 60000)).toBeUndefined();
  });

  it('returns undefined for malformed JSON', () => {
    window.sessionStorage.setItem('bad', '{ not valid json');
    expect(cacheGet('bad', 60000)).toBeUndefined();
  });

  it('returns undefined when the envelope has no numeric timestamp', () => {
    window.sessionStorage.setItem('noTs', JSON.stringify({ value: 'v' }));
    expect(cacheGet('noTs', 60000)).toBeUndefined();
  });

  it('degrades silently (no throw) when storage rejects a write', () => {
    const spy = jest
      .spyOn(window.Storage.prototype, 'setItem')
      .mockImplementation(() => {
        throw new Error('QuotaExceeded');
      });
    expect(() => cacheSet('k', 'v')).not.toThrow();
    spy.mockRestore();
  });
});
