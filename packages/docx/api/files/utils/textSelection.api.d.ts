// @betteroffice/docx declaration utils/textSelection.d.ts
// Regenerate: bun scripts/api-surface.ts --write docx
/**
 * Text Selection Utilities
 *
 * Utilities for word-level and paragraph-level text selection.
 * Used for double-click (word) and triple-click (paragraph) selection.
 * @packageDocumentation
 * @public
 */
/**
 * Check if a character is a word character
 */
export declare function isWordCharacter(char: string): boolean;
/**
 * Check if a character is whitespace
 */
export declare function isWhitespace(char: string): boolean;
/**
 * Find word boundaries around a position in text
 * Returns [startIndex, endIndex] inclusive start, exclusive end
 */
export declare function findWordBoundaries(text: string, position: number): [number, number];
/**
 * Get the word at a position in text
 */
export declare function getWordAt(text: string, position: number): string;
/**
 * Word selection result
 */
export interface WordSelectionResult {
    /** The selected word */
    word: string;
    /** Start index in the text (inclusive) */
    startIndex: number;
    /** End index in the text (exclusive) */
    endIndex: number;
}
/**
 * Find the word at a position and return detailed info
 */
export declare function findWordAt(text: string, position: number): WordSelectionResult;
/**
 * Select a word at the current cursor position using the browser's native APIs.
 * This works reliably across different browsers and handles contentEditable well.
 */
export declare function selectWordAtCursor(): boolean;
/**
 * Select a word in a specific text node at the given offset
 */
export declare function selectWordInTextNode(textNode: Text, offset: number): boolean;
/**
 * Expand the current selection to word boundaries.
 * If there's a collapsed selection (cursor), selects the word at cursor.
 * If there's an existing selection, expands to include complete words.
 */
export declare function expandSelectionToWordBoundaries(): boolean;
/**
 * Select the entire paragraph containing the current selection.
 * Looks for the nearest element with [data-paragraph-index] attribute.
 */
export declare function selectParagraphAtCursor(): boolean;
/**
 * Handle click event for multi-click detection.
 * Call this in your click handler.
 * Returns the click count (1 = single, 2 = double, 3 = triple).
 */
export declare function handleClickForMultiClick(event: MouseEvent): number;
/**
 * Create a double-click handler that selects words.
 * Returns a function that should be called on dblclick events.
 */
export declare function createDoubleClickWordSelector(): (event: MouseEvent) => void;
/**
 * Create a triple-click handler that selects paragraphs.
 * This uses our custom click counting since browsers have inconsistent triple-click.
 */
export declare function createTripleClickParagraphSelector(): (event: MouseEvent) => void;
declare const _default: {
    isWordCharacter: typeof isWordCharacter;
    isWhitespace: typeof isWhitespace;
    findWordBoundaries: typeof findWordBoundaries;
    getWordAt: typeof getWordAt;
    findWordAt: typeof findWordAt;
    selectWordAtCursor: typeof selectWordAtCursor;
    selectWordInTextNode: typeof selectWordInTextNode;
    expandSelectionToWordBoundaries: typeof expandSelectionToWordBoundaries;
    selectParagraphAtCursor: typeof selectParagraphAtCursor;
    handleClickForMultiClick: typeof handleClickForMultiClick;
    createDoubleClickWordSelector: typeof createDoubleClickWordSelector;
    createTripleClickParagraphSelector: typeof createTripleClickParagraphSelector;
};
export default _default;
