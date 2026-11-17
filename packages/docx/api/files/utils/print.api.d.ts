// @betteroffice/docx declaration utils/print.d.ts
// Regenerate: bun scripts/api-surface.ts --write docx
/**
 * Framework-agnostic print helpers shared by the React and Vue
 * adapters. Lifted from packages/docx-react/src/components/ui/PrintPreview.tsx
 * so both adapters use the same parsing / preview-window code path.
 *
 * The thin button component + the print-time CSS injection stay
 * adapter-local (they're framework-specific JSX/SFC bits); the data
 * helpers below are pure functions.
 */
export interface PrintOptions {
    includeHeaders?: boolean;
    includeFooters?: boolean;
    includePageNumbers?: boolean;
    pageRange?: {
        start: number;
        end: number;
    } | null;
    scale?: number;
    printBackground?: boolean;
    margins?: 'default' | 'none' | 'minimum';
}
export declare function getDefaultPrintOptions(): PrintOptions;
/** Trigger browser print dialog for the current document. */
export declare function triggerPrint(): void;
/**
 * Open a new window with print-optimised body content.
 *
 * Built entirely via DOM APIs (no `document.write` of interpolated strings):
 * `title` is assigned as a property so a value like `</title><script>` is
 * treated as text and cannot break out, and `content` is parsed in an inert
 * document and imported rather than concatenated into markup. `content` is the
 * caller's already-rendered print HTML; provide trusted markup.
 * Sibling copy: packages/docx-react/src/components/ui/PrintPreview.tsx.
 */
export declare function openPrintWindow(title: string | undefined, content: string): Window | null;
/** Parse "1", "1-5", etc. into a page range, or null on invalid. */
export declare function parsePageRange(input: string, maxPages: number): {
    start: number;
    end: number;
} | null;
export declare function formatPageRange(range: {
    start: number;
    end: number;
} | null, totalPages: number): string;
export declare function isPrintSupported(): boolean;
