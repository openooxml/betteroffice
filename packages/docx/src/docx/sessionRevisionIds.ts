/** Word revision ids for the tracked changes a session made, assigned as the save serializes. */

const MAX_ID = 0x7fffffff;
const ANNOTATIONS = new Set([
  'bookmarkStart',
  'bookmarkEnd',
  'commentRangeStart',
  'commentRangeEnd',
  'commentReference',
  'moveFromRangeStart',
  'moveFromRangeEnd',
  'moveToRangeStart',
  'moveToRangeEnd',
]);
const XML_ID = /\bw:id="(\d+)"/g;
const ELEMENT = new Map(
  Object.entries({
    insertion: 'ins',
    pPrIns: 'ins',
    tableRowInsertion: 'ins',
    deletion: 'del',
    pPrDel: 'del',
    tableRowDeletion: 'del',
    moveFrom: 'moveFrom',
    moveTo: 'moveTo',
    tableCellInsertion: 'cellIns',
    tableCellDeletion: 'cellDel',
    tableCellMerge: 'cellMerge',
  })
);

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
  info: object;
}

/**
 * Runs `write` with each session revision element in `tree` numbered, in
 * document order, above every annotation id the tree already holds, then
 * restores the placeholders. A source id that reappears on another element
 * kind is numbered the same way; its first element keeps it.
 * @internal
 */
export function withSessionRevisionIds<T>(tree: object, write: () => T): T {
  if (!marked) return write();
  const sites: Site[] = [];
  let max = 0;
  const kinds = new Map<number, string>();
  const note = (id: unknown): void => {
    if (typeof id === 'number' && Number.isInteger(id) && id > max && id < MAX_ID) max = id;
  };
  const visit = (value: unknown, holder: Site['holder'], key: string | number): void => {
    if (typeof value === 'string') {
      if (value.includes('w:id=')) for (const [, id] of value.matchAll(XML_ID)) note(Number(id));
      return;
    }
    if (value === null || typeof value !== 'object' || ArrayBuffer.isView(value)) return;
    if (sessionInfos.has(value)) {
      sites.push({ holder, key, info: value });
      return;
    }
    if (Array.isArray(value)) {
      value.forEach((entry, index) => visit(entry, value as unknown as Site['holder'], index));
      return;
    }
    const record = value as Record<string, unknown>;
    if (typeof record.author === 'string') {
      const kind = ELEMENT.get(key === 'info' ? String(holder.type) : String(key));
      if (kind !== undefined && typeof record.id === 'number') {
        const first = kinds.get(record.id);
        if (first === undefined) kinds.set(record.id, kind);
        else if (first !== kind) {
          sites.push({ holder, key, info: value });
          return;
        }
      }
      note(record.id);
      note(record.revisionId);
    } else if (typeof record.type === 'string' && ANNOTATIONS.has(record.type)) note(record.id);
    for (const child of Object.keys(record)) visit(record[child], record, child);
  };
  visit(tree, {}, '');
  if (sites.length === 0) return write();
  if (max + sites.length >= MAX_ID) throw new RangeError('no Word revision ids left in range');
  sites.forEach((site, index) => {
    site.holder[site.key] = { ...site.info, id: max + 1 + index };
  });
  try {
    return write();
  } finally {
    for (const site of sites) site.holder[site.key] = site.info;
  }
}
