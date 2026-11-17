// @betteroffice/docx declaration yrs/layoutMeta.d.ts
// Regenerate: bun scripts/api-surface.ts --write docx
import type { Layout } from '../layout/pagination';
export declare const LAYOUT_META_VERSION = 1;
export interface LayoutMetaV1 {
    v: 1;
    layoutRevision: number;
    pageCount: number;
    partial: boolean;
    provisional: boolean;
    notesConverged: boolean;
    pageSizes: Float64Array;
    layoutShell: string;
    headersFootersEpoch: number;
    headersFooters?: string;
}
export type RetainedLayoutMeta = Pick<LayoutMetaV1, 'pageCount' | 'partial' | 'provisional' | 'notesConverged' | 'pageSizes' | 'layoutShell'>;
export declare function isLayoutMetaV1(value: unknown): value is LayoutMetaV1;
export declare function readRetainedLayoutMeta(meta: {
    readonly page_count: number;
    readonly partial: boolean;
    readonly provisional: boolean;
    readonly notes_converged: boolean;
    page_sizes(): Float64Array;
    layout_shell_json(): string;
    free(): void;
}): RetainedLayoutMeta;
export declare function layoutMetaSummary(meta: LayoutMetaV1): Layout & {
    summaryOnly: true;
};
