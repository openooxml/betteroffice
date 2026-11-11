// @betteroffice/docx declaration yrs/pointPosition.d.ts
// Regenerate: bun scripts/api-surface.ts --write docx
import type { PointPosition } from '../plugin-api';
import type { DocxTextRange } from './edits';
import type { YrsLoc } from './index';
import { DisplayPositionIndex, type DisplayPositionReader } from './displayPositionIndex';
/** @internal */
export interface DocxResolvedPointPosition extends PointPosition {
    version: string;
    target: {
        kind: 'range';
    } & DocxTextRange;
}
/** Accepted-view offset of `loc` in its paragraph, or null when the story cannot answer. */
export declare function acceptedOffsetOf(reader: DisplayPositionReader, loc: YrsLoc): number | null;
/** @internal Resolves a display-list hit of the layout of `expectVersion` to a collapsed text range. */
export declare function resolveYrsPointPosition(index: DisplayPositionIndex, hit: PointPosition, expectVersion: string): DocxResolvedPointPosition | null;
