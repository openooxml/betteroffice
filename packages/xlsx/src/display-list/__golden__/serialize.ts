/**
 * Deterministic canonicalizer for golden snapshots.
 *
 * Produces a stable, sorted-key, pretty-printed JSON string with every float
 * rounded to a fixed precision and `-0` collapsed to `0`, so measurement
 * micro-noise never causes a spurious diff. Pure and reproducible run-to-run —
 * no Date, no random, no dependence on object insertion order.
 */

/** decimal places kept for every number before comparison. */
export const GOLDEN_PRECISION = 3;

function roundNumber(n: number): number {
  if (!Number.isFinite(n)) return n;
  const factor = 10 ** GOLDEN_PRECISION;
  const rounded = Math.round(n * factor) / factor;
  return Object.is(rounded, -0) ? 0 : rounded;
}

function canonicalize(value: unknown): unknown {
  if (typeof value === 'number') return roundNumber(value);
  if (Array.isArray(value)) return value.map(canonicalize);
  if (value && typeof value === 'object') {
    const out: Record<string, unknown> = {};
    for (const key of Object.keys(value as Record<string, unknown>).sort()) {
      const child = (value as Record<string, unknown>)[key];
      if (child === undefined) continue;
      out[key] = canonicalize(child);
    }
    return out;
  }
  return value;
}

/**
 * Serialize a value to its canonical golden string (trailing newline included).
 */
export function serialize(value: unknown): string {
  return `${JSON.stringify(canonicalize(value), null, 2)}\n`;
}
