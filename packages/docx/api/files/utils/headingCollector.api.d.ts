// @betteroffice/docx declaration utils/headingCollector.d.ts
// Regenerate: bun scripts/api-surface.ts --write docx
/**
 * @packageDocumentation
 * @public
 */
import type { EditorTreeNode } from '../types/editorTree';
/**
 * Information about a heading found in the document.
 */
export interface HeadingInfo {
    /** The text content of the heading */
    text: string;
    /** Outline level (0 = Heading 1, 1 = Heading 2, etc.) */
    level: number;
    /** Document position of the paragraph node */
    pmPos: number;
}
/**
 * Collect all headings from a document.
 *
 * Detection logic:
 * 1. Check `outlineLevel` attr (set by OOXML parsing or style resolution)
 * 2. Fallback to `styleId` matching /^[Hh]eading(\d)$/
 *
 * @deprecated A legacy tree walker kept for compatibility. `YrsSession.headings(story)` classifies
 * headings with the engine's resolver, which also reads style inheritance and document defaults.
 */
export declare function collectHeadings(doc: EditorTreeNode): HeadingInfo[];
