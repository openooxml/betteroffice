import type { YrsLoc, YrsSession } from '../../yrs';

/** A yrs location projected into the PM-shaped position space used by the display list. */
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

function geometryRoots(session: YrsSession): Map<string, StoryGeometryRoot> {
  const storyIds = session.storyIds();
  const childrenByStory = new Map<string, string[]>();
  const nestedStories = new Set<string>();
  for (const story of storyIds) {
    const children: string[] = [];
    try {
      for (const segment of session.storySegments(story)) {
        if (segment.kind !== 'embed') continue;
        for (const child of childStories(segment.payload)) {
          children.push(child);
          nestedStories.add(child);
        }
      }
    } catch {
      // A malformed/unsupported story stays unmapped rather than reaching PM.
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
  session: YrsSession,
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
        ? indexStory(session, childStory, cellPmStart + 1, paragraphs, activeStories)
        : 0;
      cellPmStart += contentSize + 2;
    }
    rowPmStart = cellPmStart + 1;
  }
  return rowPmStart + 1 - pmStart;
}

/** Index one yrs story in the same PM-shaped token space used by the renderer. */
function indexStory(
  session: YrsSession,
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

    for (const segment of session.storySegments(story)) {
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
        pmCursor += tableNodeSize(session, segment.payload, pmCursor, paragraphs, activeStories);
        paragraphPmStart = pmCursor;
        continue;
      }
      if (segment.embedKind === 'blockSdt' && atBlockBoundary) {
        const childStory = typeof segment.payload.story === 'string' ? segment.payload.story : null;
        const contentSize = childStory
          ? indexStory(session, childStory, pmCursor + 1, paragraphs, activeStories)
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

      // Inline atoms occupy one PM position inside their paragraph.
      paragraphPmUnits += 1;
      atBlockBoundary = false;
    }

    return pmCursor - pmBase;
  } finally {
    activeStories.delete(story);
  }
}

/**
 * Build a lazy PM-free projection from live yrs stories to display positions.
 * The canonical yrs segment stream supplies paragraph/atom units; table-cell
 * and block-SDT stories are recursively sized so container tokens are included.
 */
export function createYrsSidebarProjection(session: YrsSession): YrsSidebarProjection {
  const paragraphMaps = new Map<string, Map<string, ParagraphDisplaySpan> | null>();
  const roots = geometryRoots(session);

  const paragraphsForStory = (story: string): Map<string, ParagraphDisplaySpan> | null => {
    const root = roots.get(story)?.story;
    if (!root) return null;
    if (paragraphMaps.has(root)) return paragraphMaps.get(root) ?? null;

    try {
      const paragraphs = new Map<string, ParagraphDisplaySpan>();
      indexStory(session, root, 0, paragraphs, new Set());
      paragraphMaps.set(root, paragraphs);
      return paragraphs;
    } catch {
      // Unsupported embeds take the normal layout path's PM fallback, but the
      // gated sidebar read must not fall back to PM. Leave that story unplaced.
      paragraphMaps.set(root, null);
      return null;
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
