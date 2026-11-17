// @betteroffice/docx declaration editor/finalSection.d.ts
// Regenerate: bun scripts/api-surface.ts --write docx
import type { DocumentBody, SectionProperties } from '../types/document';
/**
 * The last section with inherited references; `finalSectionProperties`
 * is the body sectPr as authored.
 */
export declare function resolvedFinalSectionProperties(body: DocumentBody | null | undefined): SectionProperties | undefined;
/**
 * Applies an edit to the last section on both views: the resolved entry the
 * engine lays out and the authored sectPr the serializer saves.
 */
export declare function updateFinalSectionProperties<T extends Pick<DocumentBody, 'sections' | 'finalSectionProperties'>>(body: T, update: (properties: SectionProperties) => SectionProperties): T;
