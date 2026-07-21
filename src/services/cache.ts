/**
 * Tiny sessionStorage cache shared across the intranet homepage web parts.
 *
 * Values are wrapped in a `{ ts, value }` envelope, where `ts` is the write
 * time in epoch milliseconds. The TTL is supplied by the caller on read, so
 * each cache key can age out on its own schedule (a "configurable TTL").
 *
 * Every failure — storage unavailable, quota exceeded, malformed JSON —
 * degrades silently to a cache miss. A cache must never break the homepage.
 */

/** The `{ ts, value }` envelope every cached value is wrapped in. */
interface ICacheEnvelope<T> {
  ts: number;
  value: T;
}

/**
 * Read a cached value, treating anything written more than `ttlMs`
 * milliseconds ago as a miss.
 */
export function cacheGet<T>(key: string, ttlMs: number): T | undefined {
  try {
    const raw: string | null = window.sessionStorage.getItem(key);
    if (!raw) {
      return undefined;
    }
    const envelope: ICacheEnvelope<T> = JSON.parse(raw);
    if (!envelope || typeof envelope.ts !== 'number' || Date.now() - envelope.ts > ttlMs) {
      return undefined;
    }
    return envelope.value;
  } catch {
    return undefined;
  }
}

/** Write a value under `key`, stamped with the current time. */
export function cacheSet<T>(key: string, value: T): void {
  try {
    const envelope: ICacheEnvelope<T> = { ts: Date.now(), value: value };
    window.sessionStorage.setItem(key, JSON.stringify(envelope));
  } catch {
    /* storage full or unavailable — degrade silently */
  }
}
