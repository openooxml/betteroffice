/** Word revision ids for the tracked changes a session made, assigned as the save serializes. */

const MAX_ID = 0x7fffffff;
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
    runPropertyChange: 'rPrChange',
    paragraphPropertyChange: 'pPrChange',
    tablePropertyChange: 'tblPrChange',
    tableRowPropertyChange: 'trPrChange',
    tableCellPropertyChange: 'tcPrChange',
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
 * document order, above every non-comment annotation id the tree holds and
 * past any id already in use, then restores the placeholders. A source id
 * that reappears on another element kind is numbered the same way; its first
 * element keeps it.
 * @internal
 */
export function withSessionRevisionIds<T>(tree: object, write: () => T): T {
  if (!marked) return write();
  const sites: Site[] = [];
  let max = 0;
  const used = new Set<number>();
  const kinds = new Map<number, string>();
  const note = (id: unknown, raisesBase = true): void => {
    if (typeof id !== 'number' || !Number.isInteger(id) || id < 0 || id >= MAX_ID) return;
    used.add(id);
    if (raisesBase && id > max) max = id;
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
  const ids = sites.map(() => {
    next += 1;
    while (used.has(next)) next += 1;
    if (next >= MAX_ID) throw new RangeError('no Word revision ids left in range');
    return next;
  });
  sites.forEach((site, index) => {
    site.holder[site.key] = { ...site.info, id: ids[index] };
  });
  try {
    return write();
  } finally {
    for (const site of sites) site.holder[site.key] = site.info;
  }
}
