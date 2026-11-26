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
const NON_STORY_FIELDS = new Set(['media', 'numbering', 'styles', 'theme', 'fontTable', 'charts',
  'relationships', 'relationshipEntries']);
const XML_ID = /\bw:id\s*=\s*["'](\d+)["']/g;

const sessionInfos = new WeakMap<object, { key: unknown; id: unknown }>();
let marked = false;

/** Notes `info` for the save; a `session` info's id is a placeholder until then. @internal */
export function markRevisionInfo<T extends object>(info: T, session: boolean): T {
  if (session) {
    const id = (info as { id?: unknown }).id;
    sessionInfos.set(info, { key: id, id });
    marked = true;
  }
  return info;
}

interface Site {
  holder: Record<string | number, unknown>;
  key: string | number;
  info: { id: unknown };
}

function revisionSites(tree: object): { sites: Site[]; ids: Map<unknown, number> } {
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
      if (!value.startsWith('data:') && value.includes('w:id')) for (const [, id] of value.matchAll(XML_ID)) note(Number(id));
      return;
    }
    if (value === null || typeof value !== 'object' || ArrayBuffer.isView(value)) return;
    const session = sessionInfos.get(value);
    if (session && session.id === (value as Site['info']).id) {
      sites.push({ holder, key, info: value as Site['info'] });
      return;
    }
    if (value instanceof Map) {
      for (const child of value.values()) visit(child, {}, '');
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
    for (const child of Object.keys(record)) {
      if (!NON_STORY_FIELDS.has(child)) visit(record[child], record, child);
    }
  };
  visit(tree, {}, '');
  let next = max;
  const ids = new Map<unknown, number>();
  for (const { info } of sites) {
    const key = sessionInfos.get(info)?.key;
    if (ids.has(key)) continue;
    do {
      next = next + 1 < MAX_ID ? next + 1 : 1;
      if (next === max) throw new RangeError('no Word revision ids left in range');
    } while (used.has(next));
    ids.set(key, next);
  }
  return { sites, ids };
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
  const { sites, ids } = revisionSites(tree);
  for (const site of sites) site.holder[site.key] = { ...site.info, id: ids.get(sessionInfos.get(site.info)?.key) };
  try {
    return write();
  } finally {
    for (const site of sites) site.holder[site.key] = site.info;
  }
}

export function materializeSessionRevisionIds<T extends object>(
  tree: T,
  onClone: (source: object, clone: object) => void
): T {
  if (!marked) return tree;
  const { sites, ids } = revisionSites(tree);
  if (!sites.length) return tree;
  const copies = new WeakMap<object, object>();
  for (const { info } of sites) {
    const key = sessionInfos.get(info)?.key;
    const copy = { ...info, id: ids.get(key) };
    sessionInfos.set(copy, { key, id: copy.id });
    copies.set(info, copy);
  }
  const replace = (value: unknown): unknown => {
    if (value === null || typeof value !== 'object' ||
        value instanceof ArrayBuffer || ArrayBuffer.isView(value)) return value;
    const existing = copies.get(value);
    if (existing) return existing;
    let clone = value;
    if (value instanceof Map) {
      for (const [key, child] of value) {
        const next = replace(child);
        if (next === child) continue;
        if (clone === value) clone = new Map(value);
        (clone as Map<unknown, unknown>).set(key, next);
      }
    } else {
      for (const [key, child] of Object.entries(value)) {
        if (NON_STORY_FIELDS.has(key)) continue;
        const next = replace(child);
        if (next === child) continue;
        if (clone === value) clone = Array.isArray(value) ? value.slice() : { ...value };
        (clone as Record<string, unknown>)[key] = next;
      }
    }
    if (clone !== value) onClone(value, clone);
    copies.set(value, clone);
    return clone;
  };
  return replace(tree) as T;
}
