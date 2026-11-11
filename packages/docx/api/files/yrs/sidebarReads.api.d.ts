// @betteroffice/docx declaration yrs/sidebarReads.d.ts
// Regenerate: bun scripts/api-surface.ts --write docx
import type { YrsLoc, YrsSession } from './index';
import type { TrackedChangeEntry } from '../utils/comments';
import { type YrsSidebarDisplayPoint } from '../layout/render/yrsSidebarProjection';
export type DocxSidebarReader = Pick<YrsSession, 'version' | 'listRevisions' | 'resolveComment' | 'headings' | 'paragraphs' | 'storyIds' | 'storySegments' | 'locateParagraph'>;
export interface DocxSidebarAnchorPoints {
    start: YrsSidebarDisplayPoint | null;
    end: YrsSidebarDisplayPoint | null;
}
export interface DocxSidebarRead {
    comments: Array<{
        id: string;
        anchors: DocxSidebarAnchorPoints[];
    }>;
    revisions: Array<{
        key: string;
        start: YrsSidebarDisplayPoint | null;
    }>;
    trackedChanges: {
        entries: TrackedChangeEntry[];
        commentToRevision: Array<[number, number]>;
    };
}
export interface DocxOutlineHeading {
    story: string;
    paraId: string;
    text: string;
    level: number;
    position: number;
}
export declare function readSidebar(reader: DocxSidebarReader, commentIds: readonly string[], expectVersion: string): DocxSidebarRead | null;
/** Hybrid review navigation: the first change after the selection, or the last before it, wrapping at the ends. */
export declare function readAdjacentRevision(reader: DocxSidebarReader, selection: {
    anchor: YrsLoc;
    focus: YrsLoc;
} | null, direction: 'next' | 'previous', excluded: ReadonlySet<string>): {
    revisionId: string;
} | null;
export declare function readOutlineHeadings(reader: DocxSidebarReader, expectVersion: string): DocxOutlineHeading[] | null;
