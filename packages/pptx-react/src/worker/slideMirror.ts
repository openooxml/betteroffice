import type { ShapeAltText, SlideDisplayList, SlidePrimitive, TextLanguage, TextLink, TextLinkTarget } from '@betteroffice/pptx';
import { sanitizeHref } from '../../../../shared/safeHref';

export interface SlideMirrorSegment {
  text: string;
  lang: string | null;
  link?: TextLinkTarget;
}

export interface SlideMirrorParagraph {
  text: string;
  heading: boolean;
  lang: string | null;
  segments: SlideMirrorSegment[];
}

export interface SlideMirrorCell {
  row: number;
  column: number;
  rowSpan: number;
  columnSpan: number;
  /** Only a cell with text heads its column or row. */
  header: 'column' | 'row' | null;
  paragraphs: SlideMirrorParagraph[];
}

/** A table's cells by row, one entry per grid row; a row whose cells all continue spans from above is empty. */
export interface SlideMirrorTable {
  rows: SlideMirrorCell[][];
  columns: number;
}

/** One shape as a screen reader meets it, in slide pixels. */
export interface SlideMirrorItem {
  key: string;
  x: number;
  y: number;
  w: number;
  h: number;
  label: string | null;
  paragraphs: SlideMirrorParagraph[];
  table?: SlideMirrorTable;
  chart?: SlideMirrorChart;
  /** A chart that caches no points: it reads as its name and that it has no data. */
  chartNoData?: boolean;
}

/** A chart's data: a column per series and a row per category. */
export interface SlideMirrorChart {
  series: (string | null)[];
  rows: { category: string; values: (string | null)[] }[];
}

export const slideLinkOptions = { relative: false, protocols: ['http:', 'https:', 'mailto:', 'tel:'] } as const;

export function sameSlideLink(a: TextLinkTarget | undefined, b: TextLinkTarget | undefined): boolean {
  if (a === b) return true;
  if (!a || !b) return false;
  return a.kind === 'url' && b.kind === 'url' ? a.href === b.href
    : a.kind === 'slide' && b.kind === 'slide' ? a.index === b.index
    : a.kind === 'jump' && b.kind === 'jump' && a.jump === b.jump;
}

const inherited = (shapeId: string) => shapeId.startsWith('master:') || shapeId.startsWith('layout:');
const spoken = (text: string) => text.replace(/[\s\u000b]+/g, ' ');

const midWord = (before: string, after: string) => {
  const unspaced = /[\p{Script=Han}\p{Script=Hiragana}\p{Script=Katakana}\p{Script=Hangul}\p{Script=Thai}]/u;
  const [tail] = /[\p{L}\p{N}\p{M}]$/u.exec(before) ?? [];
  const [head] = /^[\p{L}\p{N}]/u.exec(after) ?? [];
  return tail !== undefined && head !== undefined && !unspaced.test(tail) && !unspaced.test(head);
};
const letters = /\p{L}.*\p{L}/su;
const primaryTag = (tag: string) => tag.split('-')[0]!.toLowerCase();

/** Language tags land on arbitrary runs, so a change only counts between whole words of two letters or more. */
function wordAligned(parts: SlideMirrorSegment[]): SlideMirrorSegment[] {
  const aligned: SlideMirrorSegment[] = [];
  for (const part of parts) {
    let { text } = part;
    const previous = aligned[aligned.length - 1];
    if (previous && !sameSlideLink(previous.link, part.link)) {
      aligned.push({ ...part });
      continue;
    }
    if (previous && previous.lang !== part.lang) {
      if (midWord(previous.text, text)) {
        const [rest] = /^[\p{L}\p{N}\p{M}]+/u.exec(text)!;
        previous.text += rest;
        text = text.slice(rest.length);
      }
      if (!letters.test(text)) {
        previous.text += text;
        continue;
      }
      if (!letters.test(previous.text)) {
        previous.text += text;
        previous.lang = part.lang;
        continue;
      }
    } else if (previous) {
      previous.text += text;
      continue;
    }
    aligned.push({ ...part, text });
  }
  return aligned;
}

function mirrorParagraph(source: SlideMirrorSegment[], heading = false): SlideMirrorParagraph | null {
  const parts = wordAligned(source);
  const counts = new Map<string, number>();
  for (const part of parts) if (part.lang) counts.set(part.lang, (counts.get(part.lang) ?? 0) + part.text.replace(/\P{L}/gu, '').length);
  let lang: string | null = null;
  let most = -1;
  if (parts.every((part) => part.lang)) for (const [tag, count] of counts) if (count > most) { lang = tag; most = count; }
  const segments: SlideMirrorSegment[] = [];
  for (const part of parts) {
    let text = spoken(part.text);
    const previous = segments[segments.length - 1];
    if (!previous || previous.text.endsWith(' ')) text = text.trimStart();
    if (!text) continue;
    const language = part.lang && lang && primaryTag(part.lang) === primaryTag(lang) ? lang : part.lang;
    if (previous && previous.lang === language && sameSlideLink(previous.link, part.link)) previous.text += text;
    else segments.push({ text, lang: language, ...(part.link ? { link: part.link } : {}) });
  }
  const last = segments[segments.length - 1];
  if (last) {
    last.text = last.text.trimEnd();
    if (!last.text) segments.pop();
  }
  const text = segments.map((segment) => segment.text).join('');
  return text ? { text, heading, lang, segments } : null;
}

function languageParts(text: string, languages: TextLanguage[] = []): SlideMirrorSegment[] {
  const parts: SlideMirrorSegment[] = [];
  let offset = 0;
  for (const language of languages) {
    const start = Math.max(offset, Math.min(text.length, language.start));
    const end = Math.max(start, Math.min(text.length, language.end));
    if (start > offset) parts.push({ text: text.slice(offset, start), lang: null });
    const tag = language.lang.trim();
    if (end > start) parts.push({ text: text.slice(start, end), lang: tag && tag.toLowerCase() !== 'x-none' ? tag : null });
    offset = end;
  }
  if (offset < text.length) parts.push({ text: text.slice(offset), lang: null });
  return parts;
}

function linkedParts(text: string, languages: TextLanguage[] = [], links: TextLink[] = []): SlideMirrorSegment[] {
  const parts = wordAligned(languageParts(text, languages));
  const valid = links.filter((link) => Number.isInteger(link.start) && Number.isInteger(link.end) &&
    link.start >= 0 && link.end > link.start && link.end <= text.length &&
    (link.target.kind !== 'url' || sanitizeHref(link.target.href, slideLinkOptions) !== undefined))
    .sort((a, b) => a.start - b.start);
  const result: SlideMirrorSegment[] = [];
  let offset = 0;
  let linkIndex = 0;
  for (const part of parts) {
    const end = offset + part.text.length;
    while (offset < end) {
      while (linkIndex < valid.length && valid[linkIndex]!.end <= offset) linkIndex++;
      const span = valid[linkIndex];
      const stop = span ? Math.min(end, span.start > offset ? span.start : span.end) : end;
      const link = span && span.start <= offset ? span.target : undefined;
      result.push({ text: text.slice(offset, stop), lang: part.lang, ...(link ? { link } : {}) });
      offset = stop;
    }
  }
  return result;
}

function altLabel(entry: ShapeAltText | undefined): string | null {
  return [entry?.title, entry?.description].filter(Boolean).join(': ') || null;
}

function paragraphs(primitive: SlidePrimitive): SlideMirrorParagraph[] {
  if (primitive.kind === 'textBox') return primitive.paragraphs.flatMap((paragraph) => {
    const value = mirrorParagraph(linkedParts(paragraph.runs.map((run) => run.text).join(''), paragraph.languages, paragraph.links));
    return value ? [value] : [];
  });
  return primitive.kind === 'table' ? primitive.primitives.flatMap(paragraphs) : [];
}

const count = (value: unknown, min: number): value is number => Number.isInteger(value) && (value as number) >= min;

/** The table's cells, or null when the engine sent no grid or one that does not fit the table. */
const MAX_TABLE_SLOTS = 1_000_000;
const MAX_CHART_CELLS = 10_000;
const MAX_CHART_TEXT = 1024 * 1024;

const chartText = (value: unknown, budget: { left: number }): value is string => {
  if (typeof value !== 'string') return false;
  budget.left -= value.length;
  return budget.left >= 0;
};

/** The chart's data, or null when the engine sent none or it does not fit the chart's shape or the caps. */
function mirrorChart(primitive: Extract<SlidePrimitive, { kind: 'chart' }>): SlideMirrorChart | null {
  const data = primitive.data;
  if (typeof data !== 'object' || data === null || !Array.isArray(data.categories) || !Array.isArray(data.series)) return null;
  const { categories, series } = data;
  if (!categories.length || !series.length || categories.length * series.length > MAX_CHART_CELLS) return null;
  const budget = { left: MAX_CHART_TEXT };
  if (!categories.every((category) => chartText(category, budget))) return null;
  const names: (string | null)[] = [];
  for (const entry of series) {
    if (typeof entry !== 'object' || entry === null || !Array.isArray(entry.values) || entry.values.length !== categories.length) return null;
    if (entry.name !== undefined && !chartText(entry.name, budget)) return null;
    if (!entry.values.every((value) => value === null || chartText(value, budget))) return null;
    names.push(entry.name?.trim() ? entry.name : null);
  }
  return {
    series: names,
    rows: categories.map((category, index) => ({ category, values: series.map((entry) => entry.values[index] ?? null) })),
  };
}

const plotsNoPoints = ({ data }: Extract<SlidePrimitive, { kind: 'chart' }>) =>
  Array.isArray(data?.categories) && Array.isArray(data?.series) && !data.categories.length && !data.series.length;

function mirrorTable(primitive: Extract<SlidePrimitive, { kind: 'table' }>): SlideMirrorTable | null {
  const grid = primitive.grid;
  if (!grid || !count(grid.rows, 1) || !count(grid.columns, 1) || grid.rows * grid.columns > MAX_TABLE_SLOTS ||
    !Array.isArray(grid.cells) || grid.cells.length === 0) return null;
  const children = primitive.primitives;
  const rows: SlideMirrorCell[][] = Array.from({ length: grid.rows }, () => []);
  const taken = new Set<number>();
  const ranges: [number, number][] = [];
  for (const cell of grid.cells) {
    if (typeof cell !== 'object' || cell === null) return null;
    const [start, end] = Array.isArray(cell.primitives) ? cell.primitives : [];
    if (!count(cell.row, 0) || !count(cell.column, 0) || !count(cell.rowSpan, 1) || !count(cell.columnSpan, 1) ||
      cell.row + cell.rowSpan > grid.rows || cell.column + cell.columnSpan > grid.columns ||
      !count(start, 0) || !count(end, start) || end > children.length) return null;
    for (let row = cell.row; row < cell.row + cell.rowSpan; row++) {
      for (let column = cell.column; column < cell.column + cell.columnSpan; column++) {
        const slot = row * grid.columns + column;
        if (taken.has(slot)) return null;
        taken.add(slot);
      }
    }
    ranges.push([start, end]);
    const text = children.slice(start, end).flatMap(paragraphs);
    rows[cell.row]!.push({
      row: cell.row, column: cell.column, rowSpan: cell.rowSpan, columnSpan: cell.columnSpan,
      header: text.length && (cell.header === 'column' || cell.header === 'row') ? cell.header : null,
      paragraphs: text,
    });
  }
  const covered = new Uint8Array(children.length);
  for (const [start, end] of ranges) {
    for (let index = start; index < end; index++) {
      if (covered[index]) return null;
      covered[index] = 1;
    }
  }
  if (children.some((child, index) => !covered[index] && paragraphs(child).length > 0)) return null;
  for (const row of rows) row.sort((a, b) => a.column - b.column);
  return { rows, columns: grid.columns };
}

/**
 * The slide in reading order: its own shapes back to front, as PowerPoint reads
 * them, then what the layout and master paint on every slide.
 */
export function slideMirrorItems(list: SlideDisplayList, picture: string): SlideMirrorItem[] {
  const alt = list.altText ?? {};
  const items = new Map<string, SlideMirrorItem>();
  const item = (key: string, { x, y, w, h }: SlidePrimitive) => {
    let value = items.get(key);
    if (!value) items.set(key, value = { key, x, y, w, h, label: null, paragraphs: [] });
    return value;
  };
  const describeGroups = (primitive: SlidePrimitive, group: string | undefined) => {
    if (!group || items.has(group)) return;
    describeGroups(primitive, alt[group]?.group);
    if (!alt[group]?.decorative) item(group, primitive).label = altLabel(alt[group]);
  };
  const painted = new Map(Object.entries(alt).filter(([, entry]) => entry.primitive !== undefined)
    .map(([id, entry]) => [entry.primitive, id] as const));
  for (const [index, primitive] of list.primitives.entries()) {
    const id = primitive.shapeId ?? painted.get(index);
    if (!id) continue;
    const entry = alt[id];
    describeGroups(primitive, entry?.group);
    const own = item(id, primitive);
    const table = primitive.kind === 'table' && !own.table && !own.paragraphs.length ? mirrorTable(primitive) : null;
    if (table) own.table = table;
    else own.paragraphs.push(...paragraphs(primitive));
    const chart = primitive.kind === 'chart' && !entry?.decorative && !own.chart ? mirrorChart(primitive) : null;
    if (chart) own.chart = chart;
    if (primitive.kind === 'chart' && !entry?.decorative && plotsNoPoints(primitive)) own.chartNoData = true;
    if (entry?.decorative || primitive.kind === 'textBox') continue;
    const described = altLabel(entry);
    const artwork = primitive.kind === 'image' || painted.get(index) === id;
    const summary = primitive.kind === 'shape' || primitive.kind === 'image' ? null : primitive.label ?? primitive.name;
    const fallback = entry?.group ? null : artwork ? primitive.name || picture : summary;
    own.label ??= described ?? fallback;
  }
  const ordered = [...items.values()].filter((value) => value.label !== null || value.paragraphs.length > 0 || value.table || value.chart);
  const title = list.titleShape ? items.get(list.titleShape) : undefined;
  if (title?.paragraphs.length) {
    const parts = title.paragraphs.flatMap((paragraph, index) => [
      ...(index ? [{ text: ' ', lang: null }] : []), ...paragraph.segments,
    ]);
    title.paragraphs = [mirrorParagraph(parts, true)!];
  }
  return [...ordered.filter((value) => !inherited(value.key)), ...ordered.filter((value) => inherited(value.key))];
}

/** The text of the slide's title placeholder, when it has any. */
export function slideTitle(list: SlideDisplayList): string | null {
  const id = list.titleShape;
  if (!id) return null;
  return list.primitives.filter((primitive) => primitive.shapeId === id).flatMap(paragraphs).map((paragraph) => paragraph.text).join(' ') || null;
}

/** How an announcement names a shape: its alternative text, its name, then its text. */
export function slideObjectName(list: SlideDisplayList, shapeId: string, fallback: string): string {
  const entry = list.altText?.[shapeId];
  const described = entry?.decorative ? null : altLabel(entry);
  if (described) return described;
  const start = entry?.primitive;
  const primitives = list.primitives.filter((primitive, index) => primitive.shapeId === shapeId || index === start);
  for (const primitive of primitives) if (primitive.kind !== 'textBox' && primitive.name) return primitive.name;
  return primitives.flatMap(paragraphs)[0]?.text ?? fallback;
}
