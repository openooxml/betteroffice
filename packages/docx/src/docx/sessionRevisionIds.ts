/** Word revision ids for the tracked changes a session made, assigned as the save serializes. */

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

const sessionInfos = new WeakSet<object>();
let marked = false;

/** Notes `info` for the save; a `session` info's id is a placeholder until then. @internal */
export function markRevisionInfo<T extends object>(info: T, session: boolean): T {
  if (session) sessionInfos.add(info);
  marked = true;
  return info;
}

interface Site {
  holder: Record<string | number, unknown>;
  key: string | number;
  info: { id: unknown };
}

/**
 * Runs `write` with each session revision in `tree` given one Word id, shared
 * by all its elements, then restores the placeholders. Ids are numbered in
 * document order above every non-comment annotation id the tree holds,
 * skipping any id already in use and wrapping to the lowest free id past 2^31 - 1.
 * @internal
 */
export function withSessionRevisionIds<T>(tree: object, write: () => T): T {
  if (!marked) return write();
  const sites: Site[] = [];
  let max = 0;
  const used = new Set<number>();
  const note = (id: unknown, raisesBase = true): void => {
    if (typeof id !== 'number' || !Number.isInteger(id) || id < 0 || id >= MAX_ID) return;
    used.add(id);
    if (raisesBase && id > max) max = id;
  };
  const visit = (value: unknown, holder: Site['holder'], key: string | number): void => {
    if (typeof value === 'string') {
      if (value.includes('w:id')) for (const [, id] of value.matchAll(XML_ID)) note(Number(id));
      return;
    }
    if (value === null || typeof value !== 'object' || ArrayBuffer.isView(value)) return;
    if (sessionInfos.has(value)) {
      sites.push({ holder, key, info: value as Site['info'] });
      return;
    }
    if (Array.isArray(value)) {
      value.forEach((entry, index) => visit(entry, value as unknown as Site['holder'], index));
      return;
    }
    const record = value as Record<string, unknown>;
    if (typeof record.author === 'string') {
      const comment = Array.isArray(record.content);
      note(record.id, !comment);
      note(record.revisionId, !comment);
    } else if (typeof record.type === 'string' && ANNOTATIONS.has(record.type)) {
      note(record.id, !COMMENT_MARKERS.has(record.type));
    }
    for (const child of Object.keys(record)) visit(record[child], record, child);
  };
  visit(tree, {}, '');
  if (sites.length === 0) return write();
  let next = max;
  const ids = new Map<unknown, number>();
  for (const { info } of sites) {
    if (ids.has(info.id)) continue;
    do {
      next = next + 1 < MAX_ID ? next + 1 : 1;
      if (next === max) throw new RangeError('no Word revision ids left in range');
    } while (used.has(next));
    ids.set(info.id, next);
  }
  for (const site of sites) site.holder[site.key] = { ...site.info, id: ids.get(site.info.id) };
  try {
    return write();
  } finally {
    for (const site of sites) site.holder[site.key] = site.info;
  }
}
