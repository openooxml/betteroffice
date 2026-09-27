import { MAX_COLLABORATION_FRAME_BYTES } from '../../../shared/collaboration-limits';
import type { Checkpoint, LogMutation, RetainedEntry } from './retention';

const UPDATE_PREFIX = 'update:';
const CHECKPOINT_KEY = 'checkpoint';
const CHUNK_PREFIX = 'checkpoint:';
const CHUNK_BYTES = 64 * 1024;
const BATCH_SIZE = 128;

interface Manifest {
  version: 1;
  seq: number;
  length: number;
  chunks: number;
}

export function updateKey(seq: number): string {
  return UPDATE_PREFIX + String(seq).padStart(16, '0');
}

function chunkKey(index: number): string {
  return CHUNK_PREFIX + String(index).padStart(4, '0');
}

export async function readRoom(storage: DurableObjectStorage) {
  const manifest = await storage.get<Manifest>(CHECKPOINT_KEY);
  let checkpoint: Checkpoint | undefined;
  if (manifest !== undefined) {
    if (manifest.version !== 1 || !Number.isSafeInteger(manifest.seq) || manifest.seq < 0 ||
      !Number.isSafeInteger(manifest.length) || manifest.length <= 0 || manifest.length > MAX_COLLABORATION_FRAME_BYTES ||
      manifest.chunks !== Math.ceil(manifest.length / CHUNK_BYTES)) throw new Error('Invalid checkpoint manifest');
    const bytes = new Uint8Array(manifest.length);
    for (let start = 0; start < manifest.chunks; start += BATCH_SIZE) {
      const keys = Array.from({ length: Math.min(BATCH_SIZE, manifest.chunks - start) }, (_, i) => chunkKey(start + i));
      const chunks = await storage.get<Uint8Array>(keys);
      for (let i = 0; i < keys.length; i++) {
        const chunk = chunks.get(keys[i]);
        const offset = (start + i) * CHUNK_BYTES;
        if (!(chunk instanceof Uint8Array) || chunk.length !== Math.min(CHUNK_BYTES, bytes.length - offset)) throw new Error('Incomplete relay checkpoint');
        bytes.set(chunk, offset);
      }
    }
    checkpoint = { version: 1, seq: manifest.seq, bytes };
  }
  const stored = await storage.list<Uint8Array>({ prefix: UPDATE_PREFIX });
  const entries: RetainedEntry[] = [];
  for (const [key, bytes] of stored) {
    if (!/^update:\d{16}$/.test(key) || !(bytes instanceof Uint8Array)) throw new Error('Invalid retained update');
    entries.push({ seq: Number(key.slice(UPDATE_PREFIX.length)), bytes });
  }
  const legacy = await storage.get<unknown>('updates');
  if (legacy !== undefined) {
    if (!Array.isArray(legacy) || legacy.some(bytes => !(bytes instanceof Uint8Array))) throw new Error('Invalid legacy update log');
    let seq = Math.max(checkpoint?.seq ?? -1, ...entries.map(entry => entry.seq)) + 1;
    for (const bytes of legacy) entries.push({ seq: seq++, bytes });
  }
  return { checkpoint, entries, legacy: legacy !== undefined };
}

export async function persistMutation(storage: DurableObjectStorage, mutation: LogMutation, removeLegacy = false) {
  await storage.transaction(async transaction => {
    const deletes = mutation.deletes.map(updateKey);
    if (mutation.checkpoint) {
      const { bytes, seq } = mutation.checkpoint;
      const previous = await transaction.get<Manifest>(CHECKPOINT_KEY);
      const chunks = Math.ceil(bytes.length / CHUNK_BYTES);
      const entries: Record<string, unknown> = {
        [CHECKPOINT_KEY]: { version: 1, seq, length: bytes.length, chunks } satisfies Manifest,
      };
      for (let i = 0; i < chunks; i++) entries[chunkKey(i)] = bytes.slice(i * CHUNK_BYTES, (i + 1) * CHUNK_BYTES);
      const pairs = Object.entries(entries);
      for (let i = 0; i < pairs.length; i += BATCH_SIZE) await transaction.put(Object.fromEntries(pairs.slice(i, i + BATCH_SIZE)));
      for (let i = chunks; i < (previous?.chunks ?? 0); i++) deletes.push(chunkKey(i));
    }
    for (let i = 0; i < mutation.puts.length; i += BATCH_SIZE) {
      await transaction.put(Object.fromEntries(mutation.puts.slice(i, i + BATCH_SIZE).map(entry => [updateKey(entry.seq), entry.bytes])));
    }
    if (removeLegacy) deletes.push('updates');
    for (let i = 0; i < deletes.length; i += BATCH_SIZE) await transaction.delete(deletes.slice(i, i + BATCH_SIZE));
  });
}
