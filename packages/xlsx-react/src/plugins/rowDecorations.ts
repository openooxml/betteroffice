import type { XlsxPluginRowDecoration, XlsxPluginRowDecorations } from './types';

/** Items one plugin may show at once. */
export const MAX_ROW_DECORATIONS = 10_000;
/** Items a gutter row shows before the rest collapse into a count. */
export const ROW_DECORATIONS_PER_ROW = 2;
/** Unzoomed CSS pixels. */
export const ROW_GUTTER_WIDTH = 24;

const MAX_ROWS = 1_048_576;
const MAX_TEXT = 4;
const KINDS: ReadonlySet<unknown> = new Set(['badge', 'dot']);
const TONES: ReadonlySet<unknown> = new Set(['neutral', 'info', 'success', 'warning', 'danger']);

function optionalString(value: unknown): value is string | undefined {
  return value === undefined || typeof value === 'string';
}

/** The set a hook returned, copied and frozen; throws a TypeError naming what is malformed. */
export function checkedRowDecorations(value: unknown): XlsxPluginRowDecorations | null {
  if (value === null) return null;
  const set = value as { version?: unknown; items?: unknown };
  if (typeof set !== 'object' || typeof set.version !== 'string' || !Array.isArray(set.items)) {
    throw new TypeError('rowDecorations must return null or { version, items }');
  }
  if (set.items.length > MAX_ROW_DECORATIONS) {
    throw new TypeError(`rowDecorations returned more than ${MAX_ROW_DECORATIONS} items`);
  }
  const ids = new Set<string>();
  const items: XlsxPluginRowDecoration[] = [];
  for (const entry of set.items as unknown[]) {
    const item = entry as Partial<Record<keyof XlsxPluginRowDecoration, unknown>> | null;
    if (!item || typeof item !== 'object' || typeof item.id !== 'string' || typeof item.sheetId !== 'string') {
      throw new TypeError('A row decoration needs a string id and sheetId');
    }
    if (ids.has(item.id)) throw new TypeError(`Duplicate row decoration id "${item.id}"`);
    ids.add(item.id);
    const row = item.row;
    if (typeof row !== 'number' || !Number.isInteger(row) || row < 0 || row >= MAX_ROWS) {
      throw new TypeError(`Row decoration "${item.id}" has an invalid row`);
    }
    if (!KINDS.has(item.kind)) throw new TypeError(`Row decoration "${item.id}" has an unknown kind`);
    if (item.tone !== undefined && !TONES.has(item.tone)) {
      throw new TypeError(`Row decoration "${item.id}" has an unknown tone`);
    }
    if (!optionalString(item.text) || !optionalString(item.tooltip)) {
      throw new TypeError(`Row decoration "${item.id}" text and tooltip must be strings`);
    }
    if (item.text !== undefined && Array.from(item.text).length > MAX_TEXT) {
      throw new TypeError(`Row decoration "${item.id}" text is longer than ${MAX_TEXT} characters`);
    }
    items.push(Object.freeze({
      id: item.id,
      sheetId: item.sheetId,
      row,
      kind: item.kind as XlsxPluginRowDecoration['kind'],
      ...(item.text !== undefined && { text: item.text }),
      ...(item.tone !== undefined && { tone: item.tone as XlsxPluginRowDecoration['tone'] }),
      ...(item.tooltip !== undefined && { tooltip: item.tooltip }),
    }));
  }
  return Object.freeze({ version: set.version, items: Object.freeze(items) });
}

/** One plugin's current set, as the gutter shows it. */
export interface RowDecorationSet {
  pluginId: string;
  /** The activation that computed it. */
  key: string;
  items: readonly XlsxPluginRowDecoration[];
}

export interface RowDecorationEntry {
  pluginId: string;
  key: string;
  item: XlsxPluginRowDecoration;
}

/** What a gutter row shows: its first items with their labels and the rest as a count. */
export interface RowGutterRow {
  shown: readonly { entry: RowDecorationEntry; label: string }[];
  /** Items beyond `shown`. */
  more: number;
  /** The first hidden items' labels, then `+n` for any beyond them; empty without hidden items. */
  tooltip: string;
}

export interface RowGutterIndex {
  sheetId: string;
  rows: ReadonlyMap<number, RowGutterRow>;
}

/** Hidden items a row's tooltip names before counting the rest. */
export const ROW_TOOLTIP_ITEMS = 20;

/**
 * The items of `sheetId` by row, in plugin then item order, built once per set change so a
 * paint only looks rows up; null when there are none.
 */
export function rowDecorationIndex(
  sets: readonly RowDecorationSet[],
  sheetId: string | null,
  rowLabel: (row: number) => string
): RowGutterIndex | null {
  if (sheetId === null) return null;
  const rows = new Map<number, { shown: { entry: RowDecorationEntry; label: string }[]; more: number;
    named: string[]; tooltip: string }>();
  const label = (item: XlsxPluginRowDecoration) => item.tooltip ?? item.text ?? rowLabel(item.row);
  for (const set of sets) {
    for (const item of set.items) {
      if (item.sheetId !== sheetId) continue;
      let row = rows.get(item.row);
      if (!row) {
        row = { shown: [], more: 0, named: [], tooltip: '' };
        rows.set(item.row, row);
      }
      if (row.shown.length < ROW_DECORATIONS_PER_ROW) {
        row.shown.push({ entry: { pluginId: set.pluginId, key: set.key, item }, label: label(item) });
        continue;
      }
      row.more += 1;
      if (row.named.length < ROW_TOOLTIP_ITEMS) row.named.push(label(item));
    }
  }
  for (const row of rows.values()) {
    if (!row.more) continue;
    const unnamed = row.more - row.named.length;
    row.tooltip = unnamed > 0 ? [...row.named, `+${unnamed}`].join('\n') : row.named.join('\n');
  }
  return rows.size ? { sheetId, rows } : null;
}
