import type { DisplayList } from '../display-list/types';
import type { Viewport } from './loader';
import type { XlsxPngOptions } from './pngTarget';

/** A target with its sheet resolved to an index and its range to an A1 string or 'used'. */
export type ResolvedPngTarget =
  | { sheet: number; viewport: Viewport }
  | { sheet: number; range: string };

const RANGE_SHAPE = 'range must be an A1 string or a range address';

function isIndex(value: unknown): value is number {
  return typeof value === 'number' && Number.isInteger(value) && value >= 0;
}

function hasExactly(value: object, keys: string[]): boolean {
  const own = Object.keys(value);
  return own.length === keys.length && keys.every((key) => own.includes(key));
}

function columnLetters(col: number): string {
  let letters = '';
  for (let n = col + 1; n > 0; n = Math.floor((n - 1) / 26)) letters = String.fromCharCode(65 + ((n - 1) % 26)) + letters;
  return letters;
}

function position(value: unknown): { row: number; col: number } {
  if (typeof value !== 'object' || value === null || Array.isArray(value) || !hasExactly(value, ['row', 'col'])) throw new TypeError(RANGE_SHAPE);
  const { row, col } = value as { row: unknown; col: unknown };
  if (!isIndex(row) || !isIndex(col)) throw new TypeError(RANGE_SHAPE);
  return { row, col };
}

/** The A1 string a range address names. A1 text is parsed by the engine, as for the string form. */
export function rangeA1(range: unknown): string {
  if (typeof range === 'string') return range;
  if (typeof range !== 'object' || range === null || Array.isArray(range)) throw new TypeError(RANGE_SHAPE);
  const kind = (range as { kind?: unknown }).kind;
  if (kind === 'a1' && hasExactly(range, ['kind', 'a1']) && typeof (range as { a1: unknown }).a1 === 'string') {
    return (range as { a1: string }).a1;
  }
  if (kind === 'rowCol' && hasExactly(range, ['kind', 'start', 'end'])) {
    const start = position((range as { start: unknown }).start);
    const end = position((range as { end: unknown }).end);
    if (start.row > end.row || start.col > end.col) throw new RangeError('range start must be above and left of its end');
    return `${columnLetters(start.col)}${start.row + 1}:${columnLetters(end.col)}${end.row + 1}`;
  }
  throw new TypeError(RANGE_SHAPE);
}

/** Resolves a sheet reference against the workbook's tab names. */
export function sheetIndex(sheet: unknown, names: readonly string[]): number {
  if (isIndex(sheet)) {
    if (sheet >= names.length) throw new RangeError(`sheet index ${sheet} out of range for ${names.length} sheet(s)`);
    return sheet;
  }
  if (typeof sheet === 'string') {
    const index = names.indexOf(sheet);
    if (index < 0) throw new RangeError(`no sheet named ${JSON.stringify(sheet)}`);
    return index;
  }
  throw new TypeError('sheet must be a name (string) or an index (number)');
}

/** Checks a target and options the way every engine does before it paints. */
export function resolvePngTarget(target: unknown, options: XlsxPngOptions, names: readonly string[]): ResolvedPngTarget {
  if (typeof target !== 'object' || target === null) throw new TypeError('target must be { sheet, viewport } or { sheet, range }');
  const scale = options.scale ?? 1;
  if (typeof scale !== 'number' || !Number.isFinite(scale) || scale <= 0) throw new RangeError(`scale must be a positive number, got ${String(scale)}`);
  for (const key of ['maxWidth', 'maxHeight'] as const) {
    const limit = options[key];
    if (limit !== undefined && (!Number.isInteger(limit) || limit <= 0)) throw new RangeError(`${key} must be a positive integer`);
  }
  const { sheet } = target as { sheet?: unknown };
  if ('viewport' in target && !('range' in target)) {
    const viewport = (target as { viewport: unknown }).viewport as Viewport;
    if (typeof viewport !== 'object' || viewport === null || !['x', 'y', 'width', 'height'].every((key) => Number.isFinite((viewport as unknown as Record<string, unknown>)[key]))) {
      throw new TypeError('viewport must be { x, y, width, height } in finite numbers');
    }
    return { sheet: sheetIndex(sheet, names), viewport: { x: viewport.x, y: viewport.y, width: viewport.width, height: viewport.height } };
  }
  if ('range' in target && !('viewport' in target)) {
    const range = (target as { range: unknown }).range;
    return { sheet: sheetIndex(sheet, names), range: range === 'used' ? 'used' : rangeA1(range) };
  }
  throw new TypeError('target must be { sheet, viewport } or { sheet, range }');
}

/** The engine's largest PNG side and area, in pixels. */
export const MAX_PNG_SIDE = 16_384;
export const MAX_PNG_PIXELS = 16_777_216;

const f32 = new Float32Array(1);
const bits = new Uint32Array(f32.buffer);
function nextDownF32(value: number): number {
  f32[0] = value;
  bits[0] -= 1;
  return f32[0];
}

/** An extent cut so it fits `cap` pixels at `scale`, in the engine's single-precision arithmetic. */
export function cappedExtent(extent: number, cap: number | undefined, scale: number): number {
  if (cap === undefined) return extent;
  const s = Math.fround(scale);
  let limit = Math.fround(Math.fround(cap) / s);
  while (limit > 0 && Math.ceil(Math.fround(limit * s)) > cap) limit = nextDownF32(limit);
  return Math.min(Math.fround(extent), limit);
}

/** What a display-list read needs: sheet switching and the engine document's PNG list calls. */
export interface PngListSource {
  onSheet<T>(sheet: number, read: () => T): T;
  doc: { pngDisplayListJson(viewportJson: string): string; rangePngDisplayListJson(argsJson: string): string };
}

/** A viewport target cut to the size caps, as the engine cuts it. */
export function cappedViewport(viewport: Viewport, options: XlsxPngOptions): Viewport {
  const scale = options.scale ?? 1;
  return {
    ...viewport,
    width: cappedExtent(viewport.width, options.maxWidth, scale),
    height: cappedExtent(viewport.height, options.maxHeight, scale),
  };
}

/** The engine's range arguments; `'used'` leaves the range out, which is the engine's used range. */
export function rangeArgs(range: string, options: XlsxPngOptions): string {
  return JSON.stringify({ range: range === 'used' ? undefined : range, scale: options.scale ?? 1, maxWidth: options.maxWidth, maxHeight: options.maxHeight });
}

/** The display list a canvas export of `target` paints, with its device scale and PNG size. */
export function pngDisplayList(source: PngListSource, target: ResolvedPngTarget, options: XlsxPngOptions): {
  list: DisplayList; dpr: number; width: number; height: number;
} {
  const scale = options.scale ?? 1;
  if ('viewport' in target) {
    const viewport = cappedViewport(target.viewport, options);
    const list = JSON.parse(source.onSheet(target.sheet, () => source.doc.pngDisplayListJson(JSON.stringify(viewport)))) as DisplayList;
    return {
      list, dpr: scale,
      width: Math.max(1, Math.ceil(Math.fround(viewport.width * scale))),
      height: Math.max(1, Math.ceil(Math.fround(viewport.height * scale))),
    };
  }
  const list = JSON.parse(source.onSheet(target.sheet, () => source.doc.rangePngDisplayListJson(rangeArgs(target.range, options)))) as DisplayList;
  return { list, dpr: 1, width: Math.max(1, Math.ceil(list.width)), height: Math.max(1, Math.ceil(list.height)) };
}
