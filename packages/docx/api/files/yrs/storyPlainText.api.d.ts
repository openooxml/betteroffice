// @betteroffice/docx declaration yrs/storyPlainText.d.ts
// Regenerate: bun scripts/api-surface.ts --write docx
import type { YrsLoc, YrsSession } from './index';
/** @internal */
export type PlainTextReader = Pick<YrsSession, 'storySegments'>;
/** @internal */
export declare const MAX_TABLE_COLUMNS = 16384;
interface TablePayloadCell {
    story: string;
    tcPr?: Record<string, unknown>;
}
/** @internal */
export interface TablePayloadRow {
    trPr?: Record<string, unknown>;
    cells: TablePayloadCell[];
}
/** @internal */
export interface TablePayload {
    tblPr?: Record<string, unknown>;
    grid?: unknown[];
    rows: TablePayloadRow[];
}
/** @internal */
export interface TableCellAnchor {
    row: number;
    column: number;
    rowspan: number;
    colspan: number;
    story: string;
}
/** Plain text of one non-table embed; tables use {@link tablePlainText}. @internal */
export declare function embedPlainText(reader: PlainTextReader, kind: string, payload: Record<string, unknown>): string;
/** Tab-separated rows, without grid padding for oversized blocks. @internal */
export declare function tablePlainText(reader: PlainTextReader, payload: TablePayload, range?: {
    top: number;
    bottom: number;
    left: number;
    right: number;
}): string[];
/** Plain text of story units `[from, to)`: paragraphs end in newlines, tables become tab-separated rows. @internal */
export declare function storyPlainText(reader: PlainTextReader, story: string, from?: number, to?: number): string;
export declare function tableRowGrid(trPr?: Record<string, unknown>): {
    before: number;
    after: number;
};
/** @internal */
export declare function tableAnchors(payload: TablePayload): {
    anchors: TableCellAnchor[];
    columns: number;
};
/** @internal Story offset of `loc`: the paragraph's span start plus the loc offset. */
export declare function storyOffsetForLoc(reader: Pick<YrsSession, 'locateParagraph'>, loc: YrsLoc): number;
export {};
