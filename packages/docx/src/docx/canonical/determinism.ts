import type { ParseOptions } from '../parser';
import { MAX_HEX_ID_EXCLUSIVE, type HexIdAllocator } from '../../utils/hexId';

/** Parse/save nondeterminism pinned by the canonical harness. */
export interface CanonicalDeterminism {
  seed: string;
  allocateHexId: HexIdAllocator;
  now: () => Date;
}

const contexts = new WeakMap<object, CanonicalDeterminism>();

/**
 * Attach an internal deterministic context without changing the public
 * ParseOptions contract. Rust consumes the seed for golden-locked parse IDs.
 */
export function withCanonicalDeterminism<T extends ParseOptions>(
  options: T,
  context: CanonicalDeterminism
): T {
  contexts.set(options, context);
  return options;
}

export function canonicalDeterminismFor(options: ParseOptions): CanonicalDeterminism | undefined {
  return contexts.get(options);
}

/** Seed IDs and the future serializer clock from a fixture SHA-256. */
export function createCanonicalDeterminism(fixtureSha256: string): CanonicalDeterminism {
  if (!/^[0-9a-f]{64}$/i.test(fixtureSha256)) {
    throw new TypeError('canonical determinism seed must be a SHA-256 hex digest');
  }
  let state = 0;
  for (let offset = 0; offset < fixtureSha256.length; offset += 8) {
    state = (state ^ parseInt(fixtureSha256.slice(offset, offset + 8), 16)) >>> 0;
  }
  if (state === 0) state = 0x6d2b79f5;

  const allocateHexId: HexIdAllocator = () => {
    // xorshift32: deterministic, fast, and sufficient for fixture identities.
    state ^= state << 13;
    state ^= state >>> 17;
    state ^= state << 5;
    const value = (state >>> 0) % MAX_HEX_ID_EXCLUSIVE;
    return value.toString(16).toUpperCase().padStart(8, '0');
  };

  const clockOffset = parseInt(fixtureSha256.slice(0, 12), 16) % (100 * 365 * 24 * 60 * 60 * 1000);
  const timestamp = Date.UTC(2000, 0, 1) + clockOffset;
  return { seed: fixtureSha256.toLowerCase(), allocateHexId, now: () => new Date(timestamp) };
}
