// @betteroffice/docx declaration yrs/residentCommandSelection.d.ts
// Regenerate: bun scripts/api-surface.ts --write docx
import { type RetainedFrame } from '../layout/render/frameDelta';
import { type DisplayListRect } from '../layout/render/displayListQueries';
import type { YrsStoryRange, YrsTableRange } from './index';
import type { DocxCommandApplyContext, DocxCommandIntent, DocxCommandSelection, DocxWorkerCommandEngine, DocxWorkerCommandState } from './residentCommandProtocol';
/** @internal */
export type DocxSelectionCommandIntent = Extract<DocxCommandIntent, {
    kind: 'navigate' | 'selectPoint' | 'extendSelection';
}>;
/** @internal Layout read after the predecessor command has resolved. */
export interface DocxCommandSelectionLayout {
    frame: RetainedFrame;
    layoutRevision: number;
    /** The canvas backing scale page boxes are snapped to. */
    rasterScale?: number;
}
/** @internal */
export interface ResolveDocxSelectionIntentInput {
    engine: DocxWorkerCommandEngine;
    state: DocxWorkerCommandState;
    command: DocxSelectionCommandIntent;
    context: DocxCommandApplyContext;
    readLayout?: () => DocxCommandSelectionLayout;
}
/** @internal Stale pointers always resolve or clamp; they cannot refuse a dependency. */
export type ResolveDocxSelectionIntentResult = {
    status: 'resolved' | 'clamped';
    selection: DocxCommandSelection;
    range: YrsStoryRange;
    goalX?: number;
    cells?: YrsTableRange | null;
};
export declare function releaseResidentSelectionReads(engine: DocxWorkerCommandEngine): void;
export declare function selectResidentPartPage(engine: DocxWorkerCommandEngine, story: string, pageIndex: number): void;
export declare function readDocxSelectionFocus(engine: DocxWorkerCommandEngine, selection: DocxCommandSelection, layout: DocxCommandSelectionLayout): DisplayListRect | null;
/** @internal */
export declare function resolveDocxSelectionIntent(input: ResolveDocxSelectionIntentInput): ResolveDocxSelectionIntentResult;
