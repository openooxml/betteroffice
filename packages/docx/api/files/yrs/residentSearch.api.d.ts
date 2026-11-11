// @betteroffice/docx declaration yrs/residentSearch.d.ts
// Regenerate: bun scripts/api-surface.ts --write docx
import type { YrsLoc, YrsSession, YrsStickyPosition } from './index';
import { type YrsPositionOutline } from './yrsPositionProjection';
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
type SearchReader = Pick<YrsSession, 'hasStory' | 'storySegments' | 'paragraphSpans' | 'searchText' | 'resolveStickyPosition'> & {
    positionOutline?(root: string): YrsPositionOutline | null;
};
/** @internal */
export declare function residentBodyPositions(reader: Pick<SearchReader, 'hasStory' | 'storySegments' | 'paragraphSpans' | 'positionOutline'>): (loc: YrsLoc) => number | null;
/** @internal */
export declare function readResidentSearch(reader: SearchReader, query: string, caseSensitive: boolean, carry?: YrsStickyPosition | null): ResidentSearchResult;
export {};
