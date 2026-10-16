import type { DocxContentControl } from './contentControls';
import type { YrsContentControlValue, YrsStorySegment } from './index';
import type { ResidentEngineSession } from './residentEngineSession';
import type { RetainedFrame } from '../layout/render/frameDelta';
import { DisplayPositionIndex } from './displayPositionIndex';
import type { YrsPositionProjection } from './yrsPositionProjection';
import { displayPrimitiveRect } from '../layout/render/displayListGeometry';
import type { DisplayListQueries } from '../layout/render/displayListQueries';
import type { DisplayPrimitive } from '../layout/render/displayList';

export interface ResidentControlPosition { story: string; index: number; encoded: number[] }
export interface ResidentControlTarget { control: DocxContentControl; at: ResidentControlPosition }
export interface ResidentControlWidget {
  kind: 'checkbox' | 'dropdown' | 'date' | 'repeatingItem';
  checked?: boolean;
  value?: string;
  items?: { displayText: string; value: string }[];
}
export interface ResidentControlSnapshot {
  version: string;
  frameEpoch: number | null;
  layoutRevision: number;
  controls: {
    handle: string;
    control: DocxContentControl;
    location: { story: string; index: number; rootStory: string; position: number | null };
    widget: ResidentControlWidget | null;
    rects: { pageId: string; pageIndex: number; x: number; y: number; width: number; height: number }[];
  }[];
}
interface ControlHandle { position: ResidentControlPosition }
type ControlRect = ResidentControlSnapshot['controls'][number]['rects'][number];
interface ControlEmbed { kind: string; payload: Record<string, unknown>; paraId: string; offset: number; valid: boolean }
/** Control embeds and block control positions of one document version. */
interface ControlLayout {
  version: string;
  embeds: Map<string, Map<number, ControlEmbed>>;
  blocks: Map<string, Map<string, number>>;
}

const NESTING_LIMIT = 256;
const SIZE_LIMIT = 1_000_000;
const handles = new WeakMap<ResidentEngineSession, Map<string, ControlHandle>>();
const positionIndexes = new WeakMap<ResidentEngineSession, DisplayPositionIndex>();
const layouts = new WeakMap<ResidentEngineSession, ControlLayout>();

function identity(position: ResidentControlPosition): string { return JSON.stringify([position.story, Array.from(position.encoded)]); }

/** Display positions with each inline control spanning its content, as the layout counts them. */
function controlPositionIndex(engine: ResidentEngineSession): DisplayPositionIndex {
  let index = positionIndexes.get(engine);
  if (!index) {
    index = new DisplayPositionIndex({ ...engine.geometryReader, selectionText: engine.selectionText,
      storySegments: story => engine.geometryReader.storySegments(story).map(segment => segment.kind === 'embed' && segment.embedKind === 'sdt'
        ? { kind: 'text' as const, text: ' '.repeat(inlineControlSize(segment.payload).size), attributes: segment.attributes } : segment) });
    positionIndexes.set(engine, index);
  }
  return index;
}

/** The embeds of `segments` by story index, each with its paragraph offset in display units. */
function storyEmbeds(segments: readonly YrsStorySegment[]): Map<number, ControlEmbed> {
  const embeds = new Map<number, ControlEmbed>();
  let open: ControlEmbed[] = [];
  let raw = 0;
  let paragraphStart = 0;
  let expansion = 0;
  for (const segment of segments) {
    if (segment.kind === 'pilcrow') {
      for (const embed of open) embed.paraId = segment.paraId;
      open = [];
      raw += 1;
      paragraphStart = raw;
      expansion = 0;
      continue;
    }
    if (segment.kind === 'embed') {
      const size = segment.embedKind === 'sdt' ? inlineControlSize(segment.payload) : null;
      const embed = { kind: segment.embedKind, payload: segment.payload, paraId: '', offset: raw - paragraphStart + expansion, valid: size?.valid ?? true };
      embeds.set(raw, embed);
      open.push(embed);
      expansion += size ? size.size - 1 : 0;
    }
    raw += segment.kind === 'text' ? segment.text.length : 1;
  }
  return embeds;
}

/** The display position of each block control in `projection`, keyed by its story and content story. */
function blockControlPositions(projection: YrsPositionProjection): Map<string, number> {
  const tables = new Map(projection.allTables().map(table => [table.start, table]));
  const positions = new Map<string, number>();
  const walk = (from: number, to: number): void => {
    for (let position = from; position < to;) {
      const node = projection.nodeAt(position);
      if (!node) { position += 1; continue; }
      if (node.kind === 'blockSdt') {
        const key = JSON.stringify([node.story, node.attrs.story]);
        if (!positions.has(key)) positions.set(key, position);
        walk(position + 1, position + node.nodeSize - 1);
      }
      for (const cell of node.kind === 'table' ? tables.get(position)?.cells ?? [] : []) walk(cell.contentStart, cell.contentStart + cell.nodeSize - 2);
      position += Math.max(1, node.nodeSize);
    }
  };
  walk(0, projection.size);
  return positions;
}

/** Page rects of each `sdt@` display group, by root story. */
function controlGroupRects(engine: ResidentEngineSession, frame: RetainedFrame): Map<string, Map<string, ControlRect[]>> {
  const partStories = new Map<string, string>();
  const partStory = (rId: string) => {
    let story = partStories.get(rId);
    if (story === undefined) { story = engine.partStory(rId); partStories.set(rId, story); }
    return story;
  };
  const groups = new Map<string, Map<string, ControlRect[]>>();
  for (const page of frame.pages) {
    if (page.page.unbuilt) continue;
    const regions: Array<[string, readonly DisplayPrimitive[]]> = [['body', page.page.primitives]];
    for (const part of [page.page.header, page.page.footer]) if (part) regions.push([partStory(part.rId), part.primitives]);
    for (const area of page.page.noteAreas ?? []) for (const primitive of area.primitives ?? []) {
      const note = /^(footnote|endnote)-(.+)$/.exec(primitive.groupId ?? '');
      if (note) regions.push([`${note[1] === 'footnote' ? 'fn' : 'en'}:${note[2]}`, [primitive]]);
    }
    const bounds = new Map<string, Map<string, { left: number; top: number; right: number; bottom: number }>>();
    for (const [root, primitives] of regions) for (const primitive of primitives) {
      const ids = new Set([primitive.inlineSdtWidget?.groupId, primitive.sdt?.groupId, ...(primitive.sdtPath ?? []).map(sdt => sdt.groupId)]
        .filter((id): id is string => id?.startsWith('sdt@') === true));
      if (!ids.size) continue;
      const box = displayPrimitiveRect(primitive);
      if (!Number.isFinite(box.x + box.y + box.w + box.h) || box.w <= 0 || box.h <= 0) continue;
      for (const id of ids) {
        let rootBounds = bounds.get(root);
        if (!rootBounds) { rootBounds = new Map(); bounds.set(root, rootBounds); }
        const current = rootBounds.get(id);
        rootBounds.set(id, current ? { left: Math.min(current.left, box.x), top: Math.min(current.top, box.y),
          right: Math.max(current.right, box.x + box.w), bottom: Math.max(current.bottom, box.y + box.h) }
          : { left: box.x, top: box.y, right: box.x + box.w, bottom: box.y + box.h });
      }
    }
    for (const [root, rootBounds] of bounds) {
      let rootGroups = groups.get(root);
      if (!rootGroups) { rootGroups = new Map(); groups.set(root, rootGroups); }
      for (const [id, box] of rootBounds) {
        const rects = rootGroups.get(id) ?? [];
        rects.push({ pageId: page.pageId.toString(), pageIndex: page.pageIndex, x: box.left, y: box.top, width: box.right - box.left, height: box.bottom - box.top });
        rootGroups.set(id, rects);
      }
    }
  }
  return groups;
}

/** Page rects of the display range `from..to` of `root`, one per page. */
function rangeRects(frame: RetainedFrame, queries: DisplayListQueries, root: string, from: number, to: number): ControlRect[] {
  const [kind, id] = root.split(':');
  const raw = root === 'body' ? queries.rangeRects(from, to)
    : kind === 'hf' ? [...queries.hfRangeRects('header', id!, from, to), ...queries.hfRangeRects('footer', id!, from, to)]
    : kind === 'fn' || kind === 'en' ? queries.noteRangeRects(kind === 'fn' ? 'footnote' : 'endnote', Number(id), from, to) : [];
  const pages = new Map<number, ControlRect>();
  for (const rect of raw) {
    const page = frame.pages.find(entry => entry.pageIndex === rect.pageIndex);
    if (!page || !(rect.width > 0 && rect.height > 0)) continue;
    const current = pages.get(rect.pageIndex);
    const left = Math.min(current?.x ?? rect.x, rect.x);
    const top = Math.min(current?.y ?? rect.y, rect.y);
    const right = Math.max(current ? current.x + current.width : rect.x + rect.width, rect.x + rect.width);
    const bottom = Math.max(current ? current.y + current.height : rect.y + rect.height, rect.y + rect.height);
    pages.set(rect.pageIndex, { pageId: page.pageId.toString(), pageIndex: rect.pageIndex, x: left, y: top, width: right - left, height: bottom - top });
  }
  return [...pages.values()];
}

/**
 * The controls of the current version with their widgets and, given the frame that shows it, their page rects:
 * the layout's `sdt@` group, else the painted content of an inline control the layout groups no widget for.
 */
export function readControlHandles(engine: ResidentEngineSession, frame?: RetainedFrame, layoutRevision = 0, queries?: () => DisplayListQueries): ResidentControlSnapshot {
  let registry = handles.get(engine);
  if (!registry) { registry = new Map(); handles.set(engine, registry); }
  const index = controlPositionIndex(engine);
  const version = engine.geometryReader.version();
  let layout = layouts.get(engine);
  if (layout?.version !== version) { layout = { version, embeds: new Map(), blocks: new Map() }; layouts.set(engine, layout); }
  const { embeds, blocks } = layout;
  let rects: Map<string, Map<string, ControlRect[]>> | undefined;
  const controls = engine.controlTargets().map(({ at, control }) => {
    const position = at;
    const known = [...registry!].find(([, entry]) => identity(entry.position) === identity(position));
    const handle = known?.[0] ?? crypto.randomUUID();
    if (!known) registry!.set(handle, { position });
    let storyEmbedsAt = embeds.get(at.story);
    if (!storyEmbedsAt) { storyEmbedsAt = storyEmbeds(engine.geometryReader.storySegments(at.story)); embeds.set(at.story, storyEmbedsAt); }
    const embed = storyEmbedsAt.get(at.index);
    const rootStory = at.story.startsWith('body') ? 'body' : at.story.split(':').slice(0, 2).join(':');
    let display: number | null = null;
    if (embed?.kind === 'blockSdt') {
      let positions = blocks.get(rootStory);
      if (!positions) {
        const projection = index.projection(rootStory);
        positions = projection ? blockControlPositions(projection) : new Map();
        blocks.set(rootStory, positions);
      }
      display = positions.get(JSON.stringify([at.story, embed.payload.story])) ?? null;
    } else if (embed?.kind === 'sdt' && embed.valid && embed.paraId) {
      display = index.positionOf({ story: at.story, paraId: embed.paraId, offset: embed.offset }, rootStory);
    }
    const location = { story: at.story, index: at.index, rootStory, position: display };
    const widget = embed?.valid ? controlWidget(control, embed.payload) : null;
    const grouped = frame && display !== null && widget ? (rects ??= controlGroupRects(engine, frame)).get(rootStory)?.get(`sdt@${display}`) : undefined;
    const painted = grouped ?? (frame && queries && display !== null && widget && embed?.kind === 'sdt'
      ? rangeRects(frame, queries(), rootStory, display + 1, display + inlineControlSize(embed.payload).size - 1) : undefined);
    return { handle, control, location, widget, rects: painted ? [...painted] : [] };
  });
  return { version: engine.proposalEngine.version(), frameEpoch: frame?.frameEpoch ?? null, layoutRevision, controls };
}

/** Display size of an inline control payload, and whether it stays within the projection limits. */
function inlineControlSize(payload: Record<string, unknown>): { size: number; valid: boolean } {
  let size = 0;
  let depth = 0;
  const pending: Array<{ content: unknown; depth: number }> = [{ content: payload.content, depth: 0 }];
  for (let next = pending.pop(); next; next = pending.pop()) {
    depth = Math.max(depth, next.depth);
    size += 2;
    for (const child of Array.isArray(next.content) ? next.content : []) {
      if (!child || typeof child !== 'object') continue;
      if (child.kind === 'text') size += typeof child.text === 'string' ? child.text.length : 0;
      else if (child.kind !== 'sdt') size += 1;
      else if (child.payload && typeof child.payload === 'object') pending.push({ content: child.payload.content, depth: next.depth + 1 });
      else size += 2;
    }
  }
  return { size, valid: depth <= NESTING_LIMIT && size <= SIZE_LIMIT };
}

/** `value`, parsed when it is JSON text; null when that text is malformed. */
function decoded(value: unknown): { value: unknown } | null {
  if (typeof value !== 'string') return { value };
  try { return { value: JSON.parse(value) }; } catch { return null; }
}

function controlWidget(control: DocxContentControl, payload: Record<string, unknown>): ResidentControlWidget | null {
  const properties = typeof payload.propertiesJson === 'string' ? decoded(payload.propertiesJson) : { value: {} };
  if (!properties) return null;
  const data = { ...(properties.value && typeof properties.value === 'object' ? properties.value as Record<string, unknown> : {}), ...payload };
  const current = data.value && typeof data.value === 'object' ? data.value as Record<string, unknown> : {};
  const authored = data.controlState && typeof data.controlState === 'object' ? data.controlState as Record<string, unknown> : {};
  const date = data.dateState && typeof data.dateState === 'object' ? data.dateState as Record<string, unknown> : {};
  const text = control.value.kind === 'text' ? control.value.text : '';
  if (control.controlType === 'checkbox') return { kind: 'checkbox', checked: current.checked === undefined ? data.checked === true : current.checked === true };
  if (control.controlType === 'date') return { kind: 'date', value: String(current.date ?? date.fullDate ?? authored.value ?? text).slice(0, 10) };
  if (control.controlType === 'repeatingSectionItem') return { kind: 'repeatingItem' };
  if (control.controlType !== 'dropDownList' && control.controlType !== 'comboBox') return null;
  const listed = decoded(data.listItems);
  if (!listed) return null;
  const items = (Array.isArray(listed.value) ? listed.value : []).flatMap(item => {
    if (!item || typeof item !== 'object' || typeof item.value !== 'string') return [];
    return [{ value: item.value, displayText: typeof item.displayText === 'string' ? item.displayText : item.value }];
  });
  return { kind: 'dropdown', items, value: typeof current.value === 'string' ? current.value : typeof authored.value === 'string' ? authored.value :
    typeof data.listLastValue === 'string' ? data.listLastValue : items.find(item => item.displayText === text)?.value ?? text };
}

export function resolveControlHandle(engine: ResidentEngineSession, handle: string): ResidentControlTarget | null {
  const entry = handles.get(engine)?.get(handle);
  if (!entry) return null;
  return engine.controlTargets().find(target => identity(target.at) === identity(entry.position)) ?? null;
}

export function validControlValue(value: unknown): value is YrsContentControlValue {
  if (typeof value === 'string') return !/[\uD800-\uDBFF](?![\uDC00-\uDFFF])|(?<![\uD800-\uDBFF])[\uDC00-\uDFFF]/u.test(value);
  if (!value || typeof value !== 'object' || Array.isArray(value)) return false;
  const data = value as Record<string, unknown>;
  if (data.kind === 'checkbox') return Object.keys(data).every(key => ['kind', 'checked'].includes(key)) && typeof data.checked === 'boolean';
  if (data.kind === 'dropdown') return Object.keys(data).every(key => ['kind', 'value'].includes(key)) && typeof data.value === 'string';
  if (data.kind === 'date') return Object.keys(data).every(key => ['kind', 'date'].includes(key)) && typeof data.date === 'string';
  return false;
}
