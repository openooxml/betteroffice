import type { PointPosition } from '../plugin-api';
import type { DocxTextRange } from './edits';
import type { YrsLoc } from './index';
import { projectYrsDisplayPosition } from './yrsPositionProjection';
import { displayPositionToYrsLoc } from './inputPositionMap';
import { DisplayPositionIndex, type DisplayPositionReader } from './displayPositionIndex';

/** @internal */
export interface DocxResolvedPointPosition extends PointPosition {
  version: string;
  target: { kind: 'range' } & DocxTextRange;
}

/** Accepted-view offset of `loc` in its paragraph, or null when the story cannot answer. */
export function acceptedOffsetOf(reader: DisplayPositionReader, loc: YrsLoc): number | null {
  try {
    return reader.selectionText({ story: loc.story, start: loc, end: loc }).before.length;
  } catch {
    return null;
  }
}

/** @internal Resolves a display-list hit of the layout of `expectVersion` to a collapsed text range. */
export function resolveYrsPointPosition(
  index: DisplayPositionIndex,
  hit: PointPosition,
  expectVersion: string
): DocxResolvedPointPosition | null {
  const reader = index.reader;
  if (reader.version() !== expectVersion) return null;
  const target = projectYrsDisplayPosition(hit, (story) => index.projection(story));
  if (!target || !reader.hasStory(target.story)) return null;
  const map = index.inputMap(target.story);
  const loc = map ? displayPositionToYrsLoc(map, target.displayPosition) : null;
  if (!loc) return null;
  const offset = acceptedOffsetOf(reader, loc);
  if (offset === null) return null;
  return {
    ...hit,
    version: expectVersion,
    target: {
      kind: 'range',
      story: loc.story,
      start: { paraId: loc.paraId, offset },
      end: { paraId: loc.paraId, offset },
      view: 'accepted',
    },
  };
}
