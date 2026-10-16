import type { DocxTextRange } from './edits';
import type { DocxParagraphAnchor } from './paragraphIdentity';
import type { DocxOccurrence } from './proposals';
import { paragraphInlineStart, resolveAnchorTarget, type AnchorReader, type RawAnchorRange } from './proposalGeometry';

/** Text a reservation anchors to: a range read at `version`, or matches of `text` in a paragraph. */
export type DocxReservationAnchor =
  | { kind: 'range'; version: string; range: DocxTextRange }
  | { kind: 'search'; paragraph: DocxParagraphAnchor; text: string; occurrence?: DocxOccurrence };

/** A whole paragraph, mark included. */
export interface DocxReservationParagraph {
  kind: 'paragraph';
  paragraph: DocxParagraphAnchor;
}

/**
 * Layout space reserved without changing the document, in CSS pixels at zoom 1. An `inline` box
 * sits at its anchor's start on the baseline, takes no document positions and is never split;
 * `hideAnchor` also hides the anchored text. A `block` sits beside the paragraph its anchor starts
 * in, spans the column unless `width` is set and is clipped to the column. `hide` removes its
 * anchor from render only.
 */
export type DocxReservation =
  | {
      id: string;
      kind: 'inline';
      anchor: DocxReservationAnchor;
      width: number;
      height: number;
      hideAnchor?: boolean;
    }
  | {
      id: string;
      kind: 'block';
      anchor: DocxReservationAnchor | DocxReservationParagraph;
      side: 'before' | 'after';
      height: number;
      width?: number;
    }
  | { id: string; kind: 'hide'; anchor: DocxReservationAnchor | DocxReservationParagraph };

/** @internal Every owner's reservations, computed at `version`. */
export interface DocxReservationsPayload {
  version: string;
  /** Names this payload in the frames laid out with it. */
  key?: string;
  sets: readonly { owner: string; version?: string; items: readonly DocxReservation[] }[];
}

/**
 * @internal One reservation resolved to units counted from its paragraph's first inline unit, the
 * mark at the inline length.
 */
export interface DocxResolvedReservation {
  owner: string;
  id: string;
  story: string;
  paraId: string;
  start: number;
  end: number;
  kind: 'inline' | 'block' | 'hide';
  width?: number;
  height?: number;
  side?: 'before' | 'after';
  hide?: boolean;
}

/** @internal Resolved reservations by key, as `renderEnv.reservations` carries them. */
export type DocxResolvedReservations = Record<string, DocxResolvedReservation>;

/** @internal Retains applied entries for the payload's owners. */
export function retainReservationOwners(
  payload: DocxReservationsPayload,
  previous: DocxResolvedReservations | null
): DocxResolvedReservations {
  const owners = new Set(payload.sets.map(({ owner }) => owner));
  return Object.fromEntries(Object.entries(previous ?? {}).filter(([, entry]) => owners.has(entry.owner)));
}

const MAX_EXTENT = 100_000;

type Fields = Record<string, unknown>;
const record = (value: unknown): value is Fields =>
  typeof value === 'object' && value !== null && !Array.isArray(value);
const text = (value: unknown): value is string => typeof value === 'string';
const count = (value: unknown): boolean => Number.isInteger(value) && (value as number) >= 0;
const extent = (value: unknown): boolean =>
  typeof value === 'number' && Number.isFinite(value) && value >= 0 && value <= MAX_EXTENT;

function validParagraph(value: unknown): boolean {
  if (!record(value)) return false;
  switch (value.kind) {
    case 'session':
      return text(value.sessionId) && text(value.story) && text(value.paraId);
    case 'source':
      return text(value.packageSha256) && text(value.partUri) && count(value.paragraphOrdinal);
    case 'persisted':
      return record(value.story) && text(value.story.partUri) && text(value.story.kind) &&
        text(value.paraId);
    default:
      return false;
  }
}

const validPosition = (value: unknown): boolean =>
  record(value) && text(value.paraId) && count(value.offset);

function validAnchor(value: unknown, paragraph: boolean, single: boolean): boolean {
  if (!record(value)) return false;
  switch (value.kind) {
    case 'paragraph':
      return paragraph && validParagraph(value.paragraph);
    case 'range':
      return text(value.version) && record(value.range) && text(value.range.story) &&
        validPosition(value.range.start) && validPosition(value.range.end) &&
        (value.range.view === 'accepted' || value.range.view === 'original');
    case 'search':
      return validParagraph(value.paragraph) && text(value.text) && value.text.length > 0 &&
        (value.occurrence === undefined || value.occurrence === 'first' ||
          (!single && value.occurrence === 'all') ||
          (Number.isInteger(value.occurrence) && (value.occurrence as number) >= 1));
    default:
      return false;
  }
}

function validItem(value: unknown): value is DocxReservation {
  if (!record(value) || !text(value.id) || !value.id) return false;
  switch (value.kind) {
    case 'inline':
      return validAnchor(value.anchor, false, true) && extent(value.width) &&
        extent(value.height) &&
        (value.hideAnchor === undefined || typeof value.hideAnchor === 'boolean');
    case 'block':
      return validAnchor(value.anchor, true, true) &&
        (value.side === 'before' || value.side === 'after') && extent(value.height) &&
        (value.width === undefined || extent(value.width));
    case 'hide':
      return validAnchor(value.anchor, true, false);
    default:
      return false;
  }
}

/** @internal Whether `value` is a well-formed payload: ids unique per owner, owners unique. */
export function validReservationsPayload(value: unknown): value is DocxReservationsPayload {
  if (!record(value) || !text(value.version) || !Array.isArray(value.sets)) return false;
  if (value.key !== undefined && !text(value.key)) return false;
  const owners = new Set<string>();
  return value.sets.every((set: unknown) => {
    if (!record(set) || !text(set.owner) || !set.owner || owners.has(set.owner)) return false;
    if (set.version !== undefined && (!text(set.version) || !set.version)) return false;
    if (!Array.isArray(set.items)) return false;
    owners.add(set.owner);
    const ids = new Set<string>();
    return set.items.every((item: unknown) => {
      if (!validItem(item) || ids.has(item.id)) return false;
      ids.add(item.id);
      return true;
    });
  });
}

interface Segment {
  story: string;
  paraId: string;
  start: number;
  end: number;
}

/** `range` cut at its paragraphs; `mark` takes the last paragraph's mark too. */
interface StoryIndex {
  spans: ReturnType<AnchorReader['paragraphSpans']>;
  positions: Map<string, number>;
}

function segments(indexFor: (story: string) => StoryIndex, range: RawAnchorRange, mark: boolean): Segment[] {
  const { story } = range.start;
  if (range.start.paraId === range.end.paraId) {
    return [{
      story,
      paraId: range.start.paraId,
      start: range.start.offset,
      end: range.end.offset + (mark ? 1 : 0),
    }];
  }
  const { spans, positions } = indexFor(story);
  const first = positions.get(range.start.paraId) ?? -1;
  const last = positions.get(range.end.paraId) ?? -1;
  if (first < 0 || last < first) return [];
  return spans.slice(first, last + 1).map((span, index, covered) => ({
    story,
    paraId: span.paraId,
    start: index === 0 ? range.start.offset : 0,
    end: index === covered.length - 1 ? range.end.offset + (mark ? 1 : 0) : span.length + 1,
  }));
}

/**
 * @internal Resolves a payload against `reader` at `readerVersion`, which a range anchor at the
 * payload's version is read at. Anchors that do not resolve are left out.
 */
export function resolveReservations(
  reader: AnchorReader,
  payload: DocxReservationsPayload,
  readerVersion: string = payload.version,
  previous: DocxResolvedReservations | null = null
): DocxResolvedReservations {
  const stories = new Map<string, StoryIndex>();
  const indexFor = (story: string): StoryIndex => {
    let index = stories.get(story);
    if (!index) {
      const spans = reader.paragraphSpans(story);
      const positions = new Map<string, number>();
      spans.forEach((span, position) => {
        if (!positions.has(span.paraId)) positions.set(span.paraId, position);
      });
      index = { spans, positions };
      stories.set(story, index);
    }
    return index;
  };
  const storySegments = new Map<string, ReturnType<AnchorReader['storySegments']>>();
  const cachedReader: AnchorReader = {
    resolveParagraphAnchor: (anchor) => reader.resolveParagraphAnchor(anchor),
    hasStory: (story) => reader.hasStory(story),
    findText: (...args) => reader.findText(...args),
    listRevisions: () => reader.listRevisions(),
    paragraphSpans: (story) => indexFor(story).spans,
    storySegments: (story) => {
      let segments = storySegments.get(story);
      if (!segments) {
        segments = reader.storySegments(story);
        storySegments.set(story, segments);
      }
      return segments;
    },
  };
  const resolved: DocxResolvedReservations = {};
  const priorOwners = new Map<string, [string, DocxResolvedReservation][]>();
  for (const [key, entry] of Object.entries(previous ?? {})) {
    let entries = priorOwners.get(entry.owner);
    if (!entries) priorOwners.set(entry.owner, entries = []);
    entries.push([key, entry]);
  }
  for (const { owner, version, items } of payload.sets) {
    if (version !== undefined && version !== payload.version) {
      for (const [key, entry] of priorOwners.get(owner) ?? []) resolved[key] = entry;
      continue;
    }
    for (const item of items) {
      const anchor = item.anchor.kind === 'range'
        ? { ...item.anchor, version: item.anchor.version === payload.version ? readerVersion : '' }
        : item.anchor;
      const result = resolveAnchorTarget(cachedReader, anchor, readerVersion);
      if (!result.ok || !result.ranges.length) continue;
      const mark = item.anchor.kind === 'paragraph';
      const parts = (item.kind === 'inline' ? result.ranges.slice(0, 1) : result.ranges)
        .flatMap((range) => segments(indexFor, range, mark))
        .map((part) => {
          const lead = paragraphInlineStart(cachedReader, readerVersion, part.story, part.paraId);
          return { ...part, start: Math.max(0, part.start - lead), end: part.end - lead };
        });
      const first = parts[0];
      if (!first) continue;
      const key = (index: number) => JSON.stringify(index ? [owner, item.id, index] : [owner, item.id]);
      if (item.kind === 'block') {
        resolved[key(0)] = {
          owner, id: item.id, story: first.story, paraId: first.paraId, start: 0, end: 0,
          kind: 'block', side: item.side, height: item.height,
          ...(item.width === undefined ? {} : { width: item.width }),
        };
        continue;
      }
      if (item.kind === 'inline') {
        resolved[key(0)] = {
          owner, id: item.id, ...first, end: item.hideAnchor ? first.end : first.start,
          kind: 'inline', width: item.width, height: item.height,
          ...(item.hideAnchor ? { hide: true } : {}),
        };
        if (!item.hideAnchor) continue;
      }
      parts.forEach((part, index) => {
        if (item.kind === 'hide' || index > 0) {
          resolved[key(index)] = { owner, id: item.id, ...part, kind: 'hide' };
        }
      });
    }
  }
  return resolved;
}
