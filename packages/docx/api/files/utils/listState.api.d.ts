// @betteroffice/docx declaration utils/listState.d.ts
// Regenerate: bun scripts/api-surface.ts --write docx
/**
 * Pure list-state helpers that track whether the selection is in a
 * bullet or numbered list and at what indent level.
 * @packageDocumentation
 * @public
 */
export type ListType = 'bullet' | 'numbered' | 'none';
export interface ListState {
    type: ListType;
    level: number;
    isInList: boolean;
    numId?: number;
}
export declare function createDefaultListState(): ListState;
export declare function createBulletListState(level?: number, numId?: number): ListState;
export declare function createNumberedListState(level?: number, numId?: number): ListState;
export declare function isBulletListState(state: ListState | undefined): boolean;
export declare function isNumberedListState(state: ListState | undefined): boolean;
export declare function isAnyListState(state: ListState | undefined): boolean;
export declare function getNextIndentLevel(currentLevel: number): number;
export declare function getPreviousIndentLevel(currentLevel: number): number;
export declare function toggleListType(state: ListState | undefined, targetType: ListType): ListState;
