// @betteroffice/docx declaration yrs/findMatches.d.ts
// Regenerate: bun scripts/api-surface.ts --write docx
import type { YrsLoc, YrsSession, YrsStoryRange } from './index';
import { type FindOptions } from '../utils/findReplace';
/** @internal */
export interface DocxFindDisplayMatch {
    paragraphIndex: number;
    contentIndex: number;
    startOffset: number;
    endOffset: number;
    text: string;
    displayFrom: number;
    displayTo: number;
    yrsRange: YrsStoryRange;
}
/** @internal */
export declare function findBodyMatches(reader: Pick<YrsSession, 'paragraphs'>, locToDisplayPosition: (loc: YrsLoc) => number | null, searchText: string, options: FindOptions): DocxFindDisplayMatch[];
