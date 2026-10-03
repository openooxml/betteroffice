import type { YrsLoc, YrsSession, YrsStickyPosition } from './index';
import { createYrsInputPositionMap, type YrsInputPositionMap } from './inputPositionMap';
import {
  createYrsPositionProjection,
  createYrsLocProjectionFromOutline,
  yrsLocToProjectedDisplayPosition,
  type YrsPositionOutline,
} from './yrsPositionProjection';

/** @internal */
export interface ResidentSearchMatch {
  story: string;
  paraId: string;
  start: number;
  displayFrom: number;
  displayTo: number;
}

/** @internal */
export interface ResidentSearchResult {
  matches: ResidentSearchMatch[];
  carried: number;
}

type SearchReader = Pick<
  YrsSession,
  | 'hasStory'
  | 'storySegments'
  | 'paragraphSpans'
  | 'searchText'
  | 'resolveStickyPosition'
> & { positionOutline?(root: string): YrsPositionOutline | null };

/** @internal */
export function residentBodyPositions(
  reader: Pick<SearchReader, 'hasStory' | 'storySegments' | 'paragraphSpans' | 'positionOutline'>
): (loc: YrsLoc) => number | null {
  const outline = reader.hasStory('body') ? reader.positionOutline?.('body') : null;
  const projection = outline ? createYrsLocProjectionFromOutline(outline) :
    createYrsPositionProjection(reader, 'body');
  const maps = new Map<string, YrsInputPositionMap | null>();
  const inputMap = (story: string): YrsInputPositionMap | null => {
    if (!maps.has(story)) {
      maps.set(story, reader.hasStory(story)
        ? createYrsInputPositionMap(story, reader.paragraphSpans(story))
        : null);
    }
    return maps.get(story)!;
  };
  return (loc: YrsLoc): number | null =>
    loc.story === 'body' || loc.story.startsWith('body:')
      ? yrsLocToProjectedDisplayPosition(reader, () => projection, loc, 'body', inputMap)
      : null;
}

/** @internal */
export function readResidentSearch(
  reader: SearchReader,
  query: string,
  caseSensitive: boolean,
  carry?: YrsStickyPosition | null
): ResidentSearchResult {
  if (!query) return { matches: [], carried: -1 };
  const hits = reader.searchText(query, { caseSensitive }).filter((hit) =>
    hit.story === 'body' || hit.story.startsWith('body:')
  );
  if (hits.length === 0) return { matches: [], carried: -1 };
  const positionFor = residentBodyPositions(reader);
  const matches: ResidentSearchMatch[] = [];
  for (const hit of hits) {
    const loc = { story: hit.story, paraId: hit.paraId, offset: hit.start };
    const displayFrom = positionFor(loc);
    const displayTo = positionFor({ ...loc, offset: hit.end });
    if (displayFrom == null || displayTo == null || displayFrom >= displayTo) continue;
    matches.push({
      story: hit.story, paraId: hit.paraId, start: hit.start, displayFrom, displayTo,
    });
  }
  matches.sort((a, b) => a.displayFrom - b.displayFrom);
  if (matches.length === 0) return { matches, carried: -1 };
  let loc: YrsLoc | null = null;
  try {
    loc = carry ? reader.resolveStickyPosition(carry) : null;
  } catch {}
  if (!loc) return { matches, carried: 0 };
  const exact = matches.findIndex((match) =>
    match.story === loc.story && match.paraId === loc.paraId && match.start === loc.offset
  );
  if (exact >= 0) return { matches, carried: exact };
  const position = positionFor(loc);
  const after = position == null ? -1 : matches.findIndex((match) => match.displayFrom >= position);
  return { matches, carried: after >= 0 ? after : matches.length - 1 };
}
