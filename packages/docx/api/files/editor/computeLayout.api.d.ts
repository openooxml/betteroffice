// @betteroffice/docx declaration editor/computeLayout.d.ts
// Regenerate: bun scripts/api-surface.ts --write docx
import { type LayoutMetaV1 } from '../yrs/layoutMeta';
import type { ResidentMeasurementConfig } from '../layout/measure';
import type { Layout, LayoutOptions, MeasuredBlock } from '../layout/pagination';
import type { DisplayListHeadersFooters } from '../layout/render/rustDisplayList';
import type { Document, NoteKind, SectionProperties } from '../types/document';
import type { YrsRenderEnv, YrsSession } from '../yrs';
export interface ResidentRegionLayoutRequest {
    bodyStory: 'body';
    /** @internal Use saved page totals while the layout is partial. */
    cachedPageTotals?: boolean;
    options: Pick<LayoutOptions, 'contractVersion' | 'pageGap'>;
    regions: {
        sections: Array<{
            sectionId?: string;
            properties: SectionProperties;
        }>;
        settings?: Document['package']['settings'];
        watermark?: SectionProperties['watermark'];
    };
    notes: {
        contents: Array<{
            id: number;
            noteKind: NoteKind;
            height: 0;
        }>;
    };
    renderEnv: YrsRenderEnv;
    measurement?: ResidentMeasurementConfig;
}
export interface ComputeLayoutInputs {
    document: Document | null;
    pageGap: number;
    session: Pick<YrsSession, 'layoutDocumentWithRegionsRetainedJson' | 'residentWorkerProbe' | 'retainedKernelInputsJson'>;
    renderEnv: YrsRenderEnv;
    measurement: ResidentMeasurementConfig;
    /** @internal See `ResidentRegionLayoutRequest.cachedPageTotals`. */
    cachedPageTotals?: boolean;
}
export interface LayoutComputation {
    layout: Layout;
    notesConverged: boolean;
}
export declare function buildResidentRegionLayoutRequest(document: Document | null, pageGap: number, renderEnv: YrsRenderEnv): ResidentRegionLayoutRequest;
/** What a fresh editor renders a document with, before its view options. */
export declare function documentRenderEnv(document: Document | null | undefined, colorScheme?: object): YrsRenderEnv;
export declare function getLayoutKernelInputs(layout: Layout): {
    measured: MeasuredBlock[];
    options: unknown;
    headersFooters?: DisplayListHeadersFooters;
    layoutRevision?: number;
} | undefined;
export declare function computeLayout(inputs: ComputeLayoutInputs): LayoutComputation;
/**
 * The layout a resident worker ran, from its region layout reply. The
 * measured blocks stay in the worker, so the main-thread display fallback has
 * to lay the document out on this thread first.
 */
export declare function workerLayoutComputation(layoutReply: string | LayoutMetaV1, layoutRevision?: number, headersFooters?: string): LayoutComputation;
