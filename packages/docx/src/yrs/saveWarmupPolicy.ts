import { WASM32_MEMORY_LIMIT_BYTES } from '../wasm/loadWasmAsset';

const SAVE_WARMUP_MIN_HEADROOM_BYTES = 64 * 1024 * 1024;
const SAVE_WARMUP_HEADROOM_FRACTION = 0.25;

export type SaveWarmupPolicyReason = 'ok' | 'engine-native' | 'no-limit' | 'no-stats' | 'low-headroom';

export interface SaveWarmupPolicy {
  warm: boolean;
  primeSeeds: boolean;
  reason: SaveWarmupPolicyReason;
}

export interface SaveWarmupMemory {
  liveBytes?: number;
  maximumBytes?: number;
}

function off(reason: SaveWarmupPolicyReason): SaveWarmupPolicy {
  return { warm: false, primeSeeds: false, reason };
}

function headroom(limit: number, liveBytes: number | undefined): 'ok' | 'no-stats' | 'low-headroom' {
  if (typeof liveBytes !== 'number' || !Number.isSafeInteger(liveBytes) || liveBytes < 0) return 'no-stats';
  return limit - liveBytes >= Math.max(SAVE_WARMUP_MIN_HEADROOM_BYTES, limit * SAVE_WARMUP_HEADROOM_FRACTION)
    ? 'ok'
    : 'low-headroom';
}

/**
 * Decides whether the worker warms the save path and primes save seed states.
 * Both engines measure headroom the same way; native keeps both off.
 */
export function saveWarmupPolicy(
  engine: 'wasm' | 'native',
  heapLimitBytes: number | undefined,
  memory: SaveWarmupMemory | undefined,
): SaveWarmupPolicy {
  const validLimit = heapLimitBytes === undefined || (Number.isSafeInteger(heapLimitBytes) && heapLimitBytes > 0);
  if (engine === 'native') {
    if (heapLimitBytes === undefined || !validLimit) return off('no-limit');
    const reason = headroom(heapLimitBytes, memory?.liveBytes);
    return off(reason === 'ok' ? 'engine-native' : reason);
  }
  if (!validLimit) return off('no-limit');
  const limit = Math.min(heapLimitBytes ?? WASM32_MEMORY_LIMIT_BYTES,
    memory?.maximumBytes ?? WASM32_MEMORY_LIMIT_BYTES, WASM32_MEMORY_LIMIT_BYTES);
  if (!Number.isSafeInteger(limit) || limit <= 0) return off('no-limit');
  const reason = headroom(limit, memory?.liveBytes);
  const ok = reason === 'ok';
  return { warm: ok, primeSeeds: ok, reason };
}
