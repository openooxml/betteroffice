import type { YrsLoc } from './index';

/** One native paragraph in the yrs-backed display-position projection. */
export interface YrsInputParagraphSpan {
  paraId: string;
  /** Number of UTF-16/inline units before the paragraph's pilcrow. */
  length: number;
  /** Block start emitted by the yrs display-position projection. */
  displayStart: number;
}

/**
 * Live converter between display-list positions and public yrs Locs.
 *
 * A native paragraph of length N occupies N+2 display positions (open token,
 * content, close token), so this map can be rebuilt directly from the live yrs
 * paragraph spans after every structural edit.
 */
export interface YrsInputPositionMap {
  story: string;
  paragraphs: readonly YrsInputParagraphSpan[];
}

export interface YrsInputParagraphSeed {
  paraId: string;
  length: number;
}

export function createYrsInputPositionMap(
  story: string,
  seeds: readonly YrsInputParagraphSeed[]
): YrsInputPositionMap {
  let displayStart = 0;
  const paragraphs = seeds.map((seed) => {
    const paragraph = {
      paraId: seed.paraId,
      length: Math.max(0, Math.trunc(seed.length)),
      displayStart,
    };
    displayStart += paragraph.length + 2;
    return paragraph;
  });
  return { story, paragraphs };
}

/** Convert a live yrs location to the position used by display-list geometry. */
export function yrsLocToDisplayPosition(map: YrsInputPositionMap, loc: YrsLoc): number {
  const paragraph = map.paragraphs.find((entry) => entry.paraId === loc.paraId);
  if (!paragraph || loc.story !== map.story) return map.paragraphs[0]?.displayStart ?? 0;
  return paragraph.displayStart + 1 + Math.min(Math.max(0, loc.offset), paragraph.length);
}

/** Convert a display-list hit position to a paragraph-keyed yrs location. */
export function displayPositionToYrsLoc(map: YrsInputPositionMap, position: number): YrsLoc | null {
  if (map.paragraphs.length === 0) return null;
  const pos = Math.max(0, Math.trunc(position));
  let paragraph = map.paragraphs[0];
  for (const candidate of map.paragraphs) {
    if (candidate.displayStart > pos) break;
    paragraph = candidate;
  }
  return {
    story: map.story,
    paraId: paragraph.paraId,
    offset: Math.min(Math.max(0, pos - (paragraph.displayStart + 1)), paragraph.length),
  };
}

/** Compare two Locs in document order using one live map. */
export function compareYrsLocs(map: YrsInputPositionMap, a: YrsLoc, b: YrsLoc): number {
  return yrsLocToDisplayPosition(map, a) - yrsLocToDisplayPosition(map, b);
}
