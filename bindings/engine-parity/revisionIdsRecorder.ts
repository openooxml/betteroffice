import { mock } from 'bun:test';
import { AsyncLocalStorage } from 'node:async_hooks';
import { randomUUID } from 'node:crypto';
import { mkdirSync, writeFileSync } from 'node:fs';
import { join } from 'node:path';

const out = process.env.DOCX_REVISION_IDS_DIR;
const recording = new AsyncLocalStorage<{ document: string; name: string; count: number }>();
const MAX_ID = 2 ** 31;
const COMMENT_MARKERS = new Set(['commentRangeStart', 'commentRangeEnd', 'commentReference']);
const ANNOTATIONS = new Set([
  ...COMMENT_MARKERS,
  'bookmarkStart',
  'bookmarkEnd',
  'moveFromRangeStart',
  'moveFromRangeEnd',
  'moveToRangeStart',
  'moveToRangeEnd',
]);
const XML_ID = /\bw:id\s*=\s*["'](\d+)["']/g;

function pointer(path: string[]): string {
  return path.map((key) => '/' + key.replace(/~/g, '~0').replace(/\//g, '~1')).join('');
}

export async function recordRevisionIdCases(name: string, save: () => Promise<void>): Promise<number> {
  const current = { document: name, name: name.replace(/[^a-zA-Z0-9_-]/g, '-').slice(0, 160), count: 0 };
  await recording.run(current, save);
  return current.count;
}

if (out) {
  const { isSessionRevisionInfo, markRevisionInfo, withSessionRevisionIds } =
    await import('../../packages/docx/src/docx/sessionRevisionIds');

  function inspect(tree: object) {
    const sites: Array<{ path: string[]; placeholder: number }> = [];
    const reservations: Array<{ path: string; id: number }> = [];
    const used = new Set<number>();
    const outOfRange = new Set<string>();
    const path: string[] = [];
    let max = 0;

    function note(id: unknown, raisesBase: boolean, site: string[]): void {
      if (typeof id !== 'number' || !Number.isInteger(id) || id < 0) return;
      if (id >= MAX_ID) {
        outOfRange.add(BigInt(id).toString());
        return;
      }
      used.add(id);
      reservations.push({ path: pointer(site), id });
      if (raisesBase && id > max) max = id;
    }

    function child(value: unknown, key: string): void {
      path.push(key);
      visit(value);
      path.pop();
    }

    function visit(value: unknown): void {
      if (typeof value === 'string') {
        if (value.includes('w:id')) {
          for (const [, id] of value.matchAll(XML_ID)) {
            const digits = id!.replace(/^0+(?=\d)/, '');
            if (digits.length > 10 || (digits.length === 10 && digits >= '2147483648')) {
              outOfRange.add(digits);
            } else {
              note(Number(digits), true, path);
            }
          }
        }
        return;
      }
      if (value === null || typeof value !== 'object' || ArrayBuffer.isView(value)) return;
      if (isSessionRevisionInfo(value)) {
        sites.push({ path: [...path, 'id'], placeholder: (value as { id: number }).id });
        return;
      }
      if (Array.isArray(value)) {
        value.forEach((entry, index) => child(entry, String(index)));
        return;
      }
      const record = value as Record<string, unknown>;
      if (typeof record.author === 'string') {
        const raisesBase = !Array.isArray(record.content);
        note(record.id, raisesBase, [...path, 'id']);
        note(record.revisionId, raisesBase, [...path, 'revisionId']);
      } else if (typeof record.type === 'string' && ANNOTATIONS.has(record.type)) {
        note(record.id, !COMMENT_MARKERS.has(record.type), [...path, 'id']);
      }
      for (const key of Object.keys(record)) child(record[key], key);
    }

    visit(tree);
    return { sites, reservations, used: [...used], max, outOfRange: [...outOfRange] };
  }

  mock.module('../../packages/docx/src/docx/sessionRevisionIds', () => ({
    isSessionRevisionInfo,
    markRevisionInfo,
    withSessionRevisionIds<T>(tree: object, write: () => T): T {
      if (!('determinism' in tree)) return withSessionRevisionIds(tree, write);
      const { sites, reservations, used, max, outOfRange } = inspect(tree);
      if (!sites.length) return withSessionRevisionIds(tree, write);

      const request = JSON.stringify(tree);
      const wire = withSessionRevisionIds(tree, write);
      if (typeof wire !== 'string') throw new Error('S13 revision ID recording requires a JSON wire');
      const assigned: unknown = JSON.parse(wire);
      const assignedSites = sites.map(({ path, placeholder }) => {
        const site = pointer(path);
        let id = assigned;
        for (const key of path) {
          id = id !== null && typeof id === 'object' ? (id as Record<string, unknown>)[key] : undefined;
        }
        if (typeof id !== 'number') throw new Error(`missing assigned revision ID at ${site}`);
        return { path: site, placeholder, id };
      });
      const session = [...new Set(sites.map(({ placeholder }) => placeholder))];
      const current = recording.getStore();
      const token = randomUUID();
      const seed = (tree.determinism as { seed?: string } | null)?.seed;
      const document = current ? `script:${current.document}` : `seed:${seed ?? token}`;
      mkdirSync(out, { recursive: true });
      writeFileSync(
        join(out, `${current?.name ?? 'save'}-${token}.json`),
        JSON.stringify({ document, request, session, wire, sites: assignedSites, reservations, used, max, outOfRange })
      );
      if (current) current.count += 1;
      return wire;
    },
  }));
}
