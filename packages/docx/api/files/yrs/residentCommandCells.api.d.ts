// @betteroffice/docx declaration yrs/residentCommandCells.d.ts
// Regenerate: bun scripts/api-surface.ts --write docx
import type { YrsCellLoc, YrsTableRange } from './index';
import type { DocxCommandPayloads, DocxCommandSelection, DocxWorkerCommandEngine } from './residentCommandProtocol';
export interface DocxCellNavigationPlan {
    selection: DocxCommandSelection;
    cells: YrsTableRange | null;
    insertRow?: YrsCellLoc;
}
export declare function resolveCellNavigation(engine: DocxWorkerCommandEngine, selection: DocxCommandSelection, intent: DocxCommandPayloads['navigateCell'], readOnly: boolean): DocxCellNavigationPlan | 'read-only' | null;
export declare function applyCellNavigation(engine: DocxWorkerCommandEngine, plan: DocxCellNavigationPlan, intent: DocxCommandPayloads['navigateCell']): DocxCommandSelection;
