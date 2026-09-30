import type { YrsLoc, YrsSession, YrsStorySegment } from '../../yrs';

type SidebarReader = Pick<
  YrsSession,
  'storyIds' | 'storySegments' | 'version' | 'paragraphs' | 'locateParagraph'
>;

/** A yrs location projected into the position space used by the display list. */
export interface YrsSidebarDisplayPoint {
  story: string;
  position: number;
  /** Header/footer relationship id; absent for body and nested body stories. */
  hfRid?: string;
}

/** Read-only converter used by the yrs-backed sidebar data sources. */
export interface YrsSidebarProjection {
  locToDisplayPoint(loc: YrsLoc): YrsSidebarDisplayPoint | null;
  storyOffsetToDisplayPoint(story: string, offset: number): YrsSidebarDisplayPoint | null;
}

/**
 * Where a sidebar projection reads story segments. `segments` must answer for the session's
 * current state on every call, as `YrsSession.storySegments` does; the arrays it returns are shared
 * and never mutated. A projection remembers the last source given for a session and reads
 * through it for the session's lifetime.
 */
export interface YrsStorySegmentSource {
  segments(story: string): YrsStorySegment[];
}

/**
 * Convert a string yrs identity into the numeric id used by the existing
 * layout/sidebar contract. This is the same UTF-8 FNV-1a projection used by
 * the yrs render and Document bridges; numeric OOXML ids pass through.
 */
export function yrsIdToNumericId(value: string): number {
  const parsed = Number(value);
  if (Number.isFinite(parsed)) return parsed;

  let hash = 0xcbf29ce484222325n;
  for (const byte of new TextEncoder().encode(value)) {
    hash = (hash ^ BigInt(byte)) * 0x100000001b3n;
    hash &= 0xffffffffffffffffn;
  }
  return Number(hash & ((1n << 53n) - 1n));
}

interface ParagraphDisplaySpan {
  pmStart: number;
  pmEnd: number;
}

interface StoryGeometryRoot {
  story: string;
  hfRid?: string;
}

type StorySegmentsReader = (story: string) => YrsStorySegment[];

function asRecord(value: unknown): Record<string, unknown> | null {
  return value !== null && typeof value === 'object' && !Array.isArray(value)
    ? (value as Record<string, unknown>)
    : null;
}

function childStories(payload: Record<string, unknown>): string[] {
  const direct = typeof payload.story === 'string' ? [payload.story] : [];
  const rows = Array.isArray(payload.rows) ? payload.rows : [];
  for (const rawRow of rows) {
    const row = asRecord(rawRow);
    const cells = row && Array.isArray(row.cells) ? row.cells : [];
    for (const rawCell of cells) {
      const story = asRecord(rawCell)?.story;
      if (typeof story === 'string') direct.push(story);
    }
  }
  return direct;
}

function geometryRoots(
  session: SidebarReader,
  segmentsByStory: Map<string, YrsStorySegment[]>,
  readSegments: StorySegmentsReader
): Map<string, StoryGeometryRoot> {
  const storyIds = session.storyIds();
  const childrenByStory = new Map<string, string[]>();
  const nestedStories = new Set<string>();
  for (const story of storyIds) {
    const children: string[] = [];
    try {
      const segments = readSegments(story);
      segmentsByStory.set(story, segments);
      for (const segment of segments) {
        if (segment.kind !== 'embed') continue;
        for (const child of childStories(segment.payload)) {
          children.push(child);
          nestedStories.add(child);
        }
      }
    } catch {
      // A malformed/unsupported story stays unmapped rather than reaching the fallback.
    }
    childrenByStory.set(story, children);
  }

  const roots = new Map<string, StoryGeometryRoot>();
  const registerTree = (story: string, root: StoryGeometryRoot, active: Set<string>): void => {
    if (active.has(story) || roots.has(story)) return;
    active.add(story);
    roots.set(story, root);
    for (const child of childrenByStory.get(story) ?? []) registerTree(child, root, active);
    active.delete(story);
  };

  if (storyIds.includes('body')) registerTree('body', { story: 'body' }, new Set());
  for (const story of storyIds) {
    if (!story.startsWith('hf:') || nestedStories.has(story)) continue;
    registerTree(story, { story, hfRid: story.slice('hf:'.length) }, new Set());
  }
  // DisplayListQueries has body and header/footer range APIs, but no region
  // query for independent footnote/endnote stories yet.
  return roots;
}

function tableNodeSize(
  segmentsOf: StorySegmentsReader,
  payload: Record<string, unknown>,
  pmStart: number,
  paragraphs: Map<string, ParagraphDisplaySpan>,
  activeStories: Set<string>
): number {
  const rows = Array.isArray(payload.rows) ? payload.rows : [];
  let rowPmStart = pmStart + 1;
  for (const rawRow of rows) {
    const row = asRecord(rawRow);
    const cells = row && Array.isArray(row.cells) ? row.cells : [];
    let cellPmStart = rowPmStart + 1;
    for (const rawCell of cells) {
      const cell = asRecord(rawCell);
      const childStory = typeof cell?.story === 'string' ? cell.story : null;
      const contentSize = childStory
        ? indexStory(segmentsOf, childStory, cellPmStart + 1, paragraphs, activeStories)
        : 0;
      cellPmStart += contentSize + 2;
    }
    rowPmStart = cellPmStart + 1;
  }
  return rowPmStart + 1 - pmStart;
}

/** Index one yrs story in the same token space used by the renderer. */
function indexStory(
  segmentsOf: StorySegmentsReader,
  story: string,
  pmBase: number,
  paragraphs: Map<string, ParagraphDisplaySpan>,
  activeStories: Set<string>
): number {
  if (activeStories.has(story)) throw new Error(`recursive yrs story: ${story}`);
  activeStories.add(story);
  try {
    let pmCursor = pmBase;
    let paragraphPmStart = pmBase;
    let paragraphPmUnits = 0;
    let atBlockBoundary = true;

    for (const segment of segmentsOf(story)) {
      if (segment.kind === 'text') {
        paragraphPmUnits += segment.text.length;
        atBlockBoundary = false;
        continue;
      }
      if (segment.kind === 'pilcrow') {
        const pmEnd = paragraphPmStart + paragraphPmUnits + 2;
        paragraphs.set(segment.paraId, { pmStart: paragraphPmStart, pmEnd });
        pmCursor = pmEnd;
        paragraphPmStart = pmCursor;
        paragraphPmUnits = 0;
        atBlockBoundary = true;
        continue;
      }

      if (segment.embedKind === 'table' && atBlockBoundary) {
        pmCursor += tableNodeSize(segmentsOf, segment.payload, pmCursor, paragraphs, activeStories);
        paragraphPmStart = pmCursor;
        continue;
      }
      if (segment.embedKind === 'blockSdt' && atBlockBoundary) {
        const childStory = typeof segment.payload.story === 'string' ? segment.payload.story : null;
        const contentSize = childStory
          ? indexStory(segmentsOf, childStory, pmCursor + 1, paragraphs, activeStories)
          : 0;
        pmCursor += contentSize + 2;
        paragraphPmStart = pmCursor;
        continue;
      }
      if (
        atBlockBoundary &&
        (segment.embedKind === 'pageBreak' || segment.embedKind === 'columnBreak')
      ) {
        pmCursor += 1;
        paragraphPmStart = pmCursor;
        continue;
      }

      // Inline atoms occupy one position inside their paragraph.
      paragraphPmUnits += 1;
      atBlockBoundary = false;
    }

    return pmCursor - pmBase;
  } finally {
    activeStories.delete(story);
  }
}

const projections = new WeakMap<
  SidebarReader,
  { version: string; projection: YrsSidebarProjection }
>();
const segmentSources = new WeakMap<SidebarReader, YrsStorySegmentSource>();

/**
 * Build a lazy projection from live yrs stories to display positions.
 * The canonical yrs segment stream supplies paragraph/atom units; table-cell
 * and block-SDT stories are recursively sized so container tokens are included.
 * A session gets the same projection back until its document changes.
 */
export function createYrsSidebarProjection(
  session: SidebarReader,
  source?: YrsStorySegmentSource
): YrsSidebarProjection {
  if (source) segmentSources.set(session, source);
  const version = session.version();
  const cached = projections.get(session);
  if (cached?.version === version) return cached.projection;
  const projection = projectSession(session, segmentSources.get(session));
  projections.set(session, { version, projection });
  return projection;
}

function projectSession(
  session: SidebarReader,
  source?: YrsStorySegmentSource
): YrsSidebarProjection {
  const paragraphMaps = new Map<string, Map<string, ParagraphDisplaySpan> | null>();
  const segmentsByStory = new Map<string, YrsStorySegment[]>();
  const readSegments: StorySegmentsReader = (story) =>
    source ? source.segments(story) : session.storySegments(story);
  const segmentsOf: StorySegmentsReader = (story) => {
    const segments = segmentsByStory.get(story);
    if (segments !== undefined) {
      segmentsByStory.delete(story);
      return segments;
    }
    return readSegments(story);
  };
  const roots = geometryRoots(session, segmentsByStory, readSegments);
  // Only the stories a root indexes are read again.
  for (const story of segmentsByStory.keys()) {
    if (!roots.has(story)) segmentsByStory.delete(story);
  }

  const paragraphsForStory = (story: string): Map<string, ParagraphDisplaySpan> | null => {
    const root = roots.get(story)?.story;
    if (!root) return null;
    if (paragraphMaps.has(root)) return paragraphMaps.get(root) ?? null;

    try {
      const paragraphs = new Map<string, ParagraphDisplaySpan>();
      indexStory(segmentsOf, root, 0, paragraphs, new Set());
      paragraphMaps.set(root, paragraphs);
      return paragraphs;
    } catch {
      // Unsupported embeds take the normal layout path's fallback, but the
      // gated sidebar read must not fall back to it. Leave that story unplaced.
      paragraphMaps.set(root, null);
      return null;
    } finally {
      for (const [story, owner] of roots) {
        if (owner.story === root) segmentsByStory.delete(story);
      }
    }
  };

  const locToDisplayPoint = (loc: YrsLoc): YrsSidebarDisplayPoint | null => {
    const root = roots.get(loc.story);
    if (!root) return null;
    const block = paragraphsForStory(loc.story)?.get(loc.paraId);
    if (!block) return null;
    const paragraphLength = Math.max(0, block.pmEnd - block.pmStart - 2);
    const offset = Math.min(Math.max(0, Math.trunc(loc.offset)), paragraphLength);
    return {
      story: loc.story,
      position: block.pmStart + 1 + offset,
      ...(root.hfRid ? { hfRid: root.hfRid } : {}),
    };
  };

  const storyOffsetToDisplayPoint = (
    story: string,
    offset: number
  ): YrsSidebarDisplayPoint | null => {
    if (!roots.has(story)) return null;
    const target = Math.max(0, Math.trunc(offset));
    try {
      for (const paragraph of session.paragraphs(story)) {
        const span = session.locateParagraph(story, paragraph.paraId);
        if (target < span.start || target > span.end) continue;
        return locToDisplayPoint({
          story,
          paraId: paragraph.paraId,
          offset: Math.min(target - span.start, span.end - span.start),
        });
      }
    } catch {
      return null;
    }
    return null;
  };

  return { locToDisplayPoint, storyOffsetToDisplayPoint };
}
