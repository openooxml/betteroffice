import type { PointPosition } from '../plugin-api';
import type { DocxTextRange } from './edits';
import type { YrsSession } from './index';
import { createYrsInputPositionMap, displayPositionToYrsLoc } from './inputPositionMap';
import { createYrsPositionProjection, projectYrsDisplayPosition } from './yrsPositionProjection';

/** @internal */
export interface DocxResolvedPointPosition extends PointPosition {
  version: string;
  target: { kind: 'range' } & DocxTextRange;
}

/** @internal */
export function resolveYrsPointPosition(
  reader: Pick<YrsSession, 'version' | 'hasStory' | 'storySegments' | 'paragraphSpans' | 'selectionText'>,
  hit: PointPosition,
  expectVersion: string
): DocxResolvedPointPosition | null {
  if (reader.version() !== expectVersion) return null;
  const target = projectYrsDisplayPosition(hit, (story) => createYrsPositionProjection(reader, story));
  if (!target || !reader.hasStory(target.story)) return null;
  const map = createYrsInputPositionMap(target.story, reader.paragraphSpans(target.story));
  const loc = displayPositionToYrsLoc(map, target.displayPosition);
  if (!loc) return null;
  try {
    const offset = reader.selectionText({ story: loc.story, start: loc, end: loc }).before.length;
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
  } catch {
    return null;
  }
}
