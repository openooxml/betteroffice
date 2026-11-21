// @betteroffice/docx declaration yrs/listCommands.d.ts
// Regenerate: bun scripts/api-surface.ts --write docx
import type { YrsAuthor, YrsParagraphAttrs, YrsSession, YrsStoryRange } from './index';
/** The attributes that take a paragraph off its list. @internal */
export declare const LIST_OFF_ATTRS: Readonly<YrsParagraphAttrs>;
/**
 * The attributes that take a paragraph with `properties` off its list: `LIST_OFF_ATTRS`, and for a
 * paragraph its style numbers the `numId` 0 that keeps it off that numbering in the saved file.
 * `styleNumbers` is the engine's answer for the style the paragraph has now; without it the flag
 * stored on the paragraph when the file was read decides.
 * @internal
 */
export declare function listOffAttrs(properties: Readonly<Record<string, unknown>>, styleNumbers?: boolean): YrsParagraphAttrs;
/**
 * Bullets or Numbering over the paragraphs of `range`: every paragraph off its list when the first
 * is on a list of `kind`, otherwise the first onto the list `ensureListDefinition` gives it and the
 * others onto that list after it.
 * @internal
 */
export declare function applyListToggle(engine: Pick<YrsSession, 'paragraphs' | 'selectionContext' | 'ensureListDefinition' | 'setParagraphAttrs'> & Partial<Pick<YrsSession, 'paragraphStyleNumbers'>>, range: YrsStoryRange, kind: 'bullet' | 'numbered', suggesting?: YrsAuthor): void;
