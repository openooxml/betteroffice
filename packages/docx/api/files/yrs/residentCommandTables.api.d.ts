// @betteroffice/docx declaration yrs/residentCommandTables.d.ts
// Regenerate: bun scripts/api-surface.ts --write docx
import type { YrsCellLoc, YrsLoc, YrsTableLoc, YrsTableRange } from './index';
import type { DocxCommandIntent, DocxCommandSelection, DocxWorkerCommandEngine, DocxWorkerCommandState } from './residentCommandProtocol';
export type TableCommandKind = 'insertTableRow' | 'insertTableColumn' | 'deleteTableRow' | 'deleteTableColumn' | 'deleteTable' | 'mergeTableCells' | 'splitTableCell' | 'setTableCellShading' | 'setTableBorders' | 'setTableProperties';
export type TableCommandIntent = Extract<DocxCommandIntent, {
    kind: TableCommandKind;
}>;
export interface DocxTableCommandTarget {
    focused: YrsCellLoc;
    range: YrsTableRange;
    table: YrsTableLoc;
    near: YrsLoc | null;
}
export declare function isTableCommandKind(kind: string): kind is TableCommandKind;
export declare function nearTable(engine: DocxWorkerCommandEngine, table: YrsTableLoc): YrsLoc | null;
/** Live tables holding a story, innermost first; empty outside tables. */
export declare function liveTableAncestors(engine: DocxWorkerCommandEngine, story: string): YrsTableLoc[];
export declare function preflightTableCommand(engine: DocxWorkerCommandEngine, command: TableCommandIntent, selection: DocxCommandSelection): {
    target: DocxTableCommandTarget;
} | {
    reason: string;
};
export declare function syncCellSelectionWithCaret(engine: DocxWorkerCommandEngine, story: string): void;
export interface TableCommandSink {
    caret(at: YrsLoc): void;
    repair(): boolean;
}
/** Runs inside the caller's discrete undo group. */
export declare function applyTableCommand(engine: DocxWorkerCommandEngine, state: DocxWorkerCommandState, command: TableCommandIntent, target: DocxTableCommandTarget, keep: YrsLoc, sink: TableCommandSink): void;
