import { mkdtempSync, readFileSync, rmSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join, resolve } from 'node:path';

const root = resolve(import.meta.dir, '../..');

/** Every page PNG export laid out for each shared fixture on one engine, as display-list JSON. */
export interface DisplayListDump { engine: 'wasm' | 'native'; fonts: 'pinned' | 'default'; pages: Record<string, string[]> }

/**
 * Runs `displayListParity.dump.ts` in a process of its own on one engine, since
 * the engine is chosen per process. `pinned` passes both engines the same
 * complete font provider; `default` leaves each engine its own default.
 */
export function dumpDisplayLists(engine: 'wasm' | 'native', fonts: 'pinned' | 'default'): DisplayListDump {
  const dir = mkdtempSync(join(tmpdir(), 'docx-display-list-parity-'));
  try {
    const out = join(dir, `${engine}.json`);
    const preloads = [
      'bindings/docx-native/displayListParity.capture.ts',
      'scripts/test-preload.ts',
      ...(engine === 'native' ? ['scripts/native-engine-preload.ts'] : []),
    ].flatMap(path => ['--preload', resolve(root, path)]);
    const env = { ...process.env };
    delete env.BUN_OPTIONS;
    const child = Bun.spawnSync([process.execPath, ...preloads, resolve(import.meta.dir, 'displayListParity.dump.ts'), out, fonts], { cwd: root, env, stderr: 'pipe' });
    if (child.exitCode !== 0) throw new Error(`${engine} display-list dump failed (${child.exitCode}): ${child.stderr.toString().slice(-4000)}`);
    const dump = JSON.parse(readFileSync(out, 'utf8')) as DisplayListDump;
    if (dump.engine !== engine) throw new Error(`the ${engine} dump ran on the ${dump.engine} engine`);
    return dump;
  } finally {
    rmSync(dir, { recursive: true, force: true });
  }
}

/** One recorded page: the font chains the engine was handed, the face behind each id, and the page's display list. */
export interface RecordedPage { fontChains?: Record<string, number[]>; faces?: Record<string, string>; page: { primitives?: unknown[] } & Record<string, unknown> }

/** A recorded page after `canonicalFaceIds`: `faces[n]` is the face canonical id `n` stands for. */
export interface CanonicalPage { fontChains?: Record<string, number[]>; faces: string[]; page: RecordedPage['page'] }

/**
 * Renumbers the face ids of one recorded page by first occurrence: chains in
 * sorted key order, then chain position, then every `fontId` in a walk of the
 * page. An id that no chain holds is numbered after all chain ids, so it never
 * shares a number with a chained face. Each canonical id keeps the face the
 * recording named for its original id; an id the recording names no face for
 * throws.
 */
export function canonicalFaceIds(json: string): CanonicalPage {
  const recorded = JSON.parse(json) as RecordedPage;
  const numbers = new Map<number, number>();
  const faces: string[] = [];
  const renumber = (id: number) => {
    if (!numbers.has(id)) {
      const face = recorded.faces?.[id];
      if (face === undefined) throw new Error(`face id ${id} has no recorded face`);
      numbers.set(id, numbers.size);
      faces.push(face);
    }
    return numbers.get(id)!;
  };
  const chains = recorded.fontChains;
  const fontChains = chains && Object.fromEntries(Object.keys(chains).sort().map(key => [key, chains[key]!.map(renumber)]));
  const walk = (value: unknown): unknown => {
    if (Array.isArray(value)) return value.map(walk);
    if (value === null || typeof value !== 'object') return value;
    return Object.fromEntries(Object.entries(value).map(([key, field]) => [key, key === 'fontId' && typeof field === 'number' ? renumber(field) : walk(field)]));
  };
  const page = walk(recorded.page) as RecordedPage['page'];
  return { ...(fontChains ? { fontChains } : {}), faces, page };
}

/** Where two dumps first differ for each fixture: the page, its font chains or faces when they differ, and the first differing primitive or field, all after `canonicalFaceIds`. */
export function firstDifferences(expected: DisplayListDump, actual: DisplayListDump): string[] {
  const found: string[] = [];
  for (const name of Object.keys(expected.pages).sort()) {
    const ours = (expected.pages[name] ?? []).map(canonicalFaceIds);
    const theirs = (actual.pages[name] ?? []).map(canonicalFaceIds);
    if (ours.length !== theirs.length) {
      found.push(`${name}: ${expected.engine} lays out ${ours.length} page(s), ${actual.engine} ${theirs.length}`);
      continue;
    }
    const page = ours.findIndex((recorded, index) => JSON.stringify(recorded) !== JSON.stringify(theirs[index]));
    if (page < 0) continue;
    const a = ours[page]!;
    const b = theirs[page]!;
    const clip = (value: unknown) => JSON.stringify(value)?.slice(0, 600);
    const chainsDiffer = JSON.stringify(a.fontChains) !== JSON.stringify(b.fontChains);
    if (chainsDiffer) {
      found.push(`${name} page ${page}: font chains differ: ${clip(a.fontChains)} vs ${clip(b.fontChains)}`);
    }
    const facesDiffer = JSON.stringify(a.faces) !== JSON.stringify(b.faces);
    if (facesDiffer) {
      found.push(`${name} page ${page}: faces differ: ${clip(a.faces)} vs ${clip(b.faces)}`);
    }
    const left = a.page.primitives ?? [];
    const right = b.page.primitives ?? [];
    const index = left.findIndex((primitive, at) => JSON.stringify(primitive) !== JSON.stringify(right[at]));
    if (index >= 0 || left.length !== right.length) {
      const at = index >= 0 ? index : Math.min(left.length, right.length);
      found.push(`${name} page ${page} primitive ${at} of ${left.length}/${right.length}: ${expected.engine} ${clip(left[at])} vs ${actual.engine} ${clip(right[at])}`);
      continue;
    }
    const field = Object.keys({ ...a.page, ...b.page }).find(key => JSON.stringify(a.page[key]) !== JSON.stringify(b.page[key]));
    if (field !== undefined || (!chainsDiffer && !facesDiffer)) found.push(`${name} page ${page} field ${field}: ${clip(field && a.page[field])} vs ${clip(field && b.page[field])}`);
  }
  return found;
}

/** How many pages and primitives differ between two dumps, per fixture, after `canonicalFaceIds`. */
export function differenceCounts(expected: DisplayListDump, actual: DisplayListDump): Record<string, { pages: string; primitives: string }> {
  const counts: Record<string, { pages: string; primitives: string }> = {};
  for (const name of Object.keys(expected.pages).sort()) {
    const ours = (expected.pages[name] ?? []).map(canonicalFaceIds);
    const theirs = (actual.pages[name] ?? []).map(canonicalFaceIds);
    let pages = 0;
    let primitives = 0;
    let total = 0;
    for (let index = 0; index < Math.max(ours.length, theirs.length); index++) {
      const left = ours[index]?.page.primitives ?? [];
      const right = theirs[index]?.page.primitives ?? [];
      if (JSON.stringify(ours[index]) !== JSON.stringify(theirs[index])) pages++;
      total += Math.max(left.length, right.length);
      for (let at = 0; at < Math.max(left.length, right.length); at++) {
        if (JSON.stringify(left[at]) !== JSON.stringify(right[at])) primitives++;
      }
    }
    counts[name] = { pages: `${pages}/${Math.max(ours.length, theirs.length)}`, primitives: `${primitives}/${total}` };
  }
  return counts;
}
