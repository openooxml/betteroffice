// @betteroffice/docx declaration managers/types.d.ts
// Regenerate: bun scripts/api-surface.ts --write docx
/**
 * Manager Types
 *
 * Framework-agnostic interfaces for the editor's manager classes.
 * @packageDocumentation
 * @public
 */
/** Cell coordinates in a table */
export interface CellCoordinates {
    tableIndex: number;
    rowIndex: number;
    columnIndex: number;
}
/** TableSelectionManager snapshot */
export interface TableSelectionSnapshot {
    /** Currently selected cell, or null if no selection */
    selectedCell: CellCoordinates | null;
}
/** Error severity levels */
export type ErrorSeverity = 'error' | 'warning' | 'info';
/** Error notification */
export interface ErrorNotification {
    id: string;
    message: string;
    severity: ErrorSeverity;
    details?: string;
    timestamp: number;
    dismissed?: boolean;
}
/** ErrorManager snapshot */
export interface ErrorManagerSnapshot {
    notifications: ErrorNotification[];
}
