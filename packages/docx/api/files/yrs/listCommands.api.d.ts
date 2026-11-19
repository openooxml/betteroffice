// @betteroffice/docx declaration yrs/listCommands.d.ts
// Regenerate: bun scripts/api-surface.ts --write docx
import type { YrsAuthor, YrsParagraphAttrs, YrsSession, YrsStoryRange } from './index';
/** The attributes that take a paragraph off its list. @internal */
export declare const LIST_OFF_ATTRS: Readonly<YrsParagraphAttrs>;
/**
 * Bullets or Numbering over the paragraphs of `range`: every paragraph off its list when the first
 * is on a list of `kind`, otherwise each in turn onto the list `ensureListDefinition` gives it.
 * @internal
 */
export declare function applyListToggle(engine: Pick<YrsSession, 'paragraphs' | 'selectionContext' | 'ensureListDefinition' | 'setParagraphAttrs'>, range: YrsStoryRange, kind: 'bullet' | 'numbered', suggesting?: YrsAuthor): void;
