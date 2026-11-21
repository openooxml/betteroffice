import type { YrsParagraphAttrs } from '@betteroffice/docx/yrs';
import type { DocxCommandArgs } from './types';

export function paragraphSpacingAttrs(
  spacing: DocxCommandArgs['paragraphSpacing'] & { lineSpacing?: number }
): YrsParagraphAttrs {
  return {
    ...(spacing.before !== undefined ? { spaceBefore: spacing.before } : {}),
    ...(spacing.after !== undefined ? { spaceAfter: spacing.after } : {}),
    ...(spacing.lineSpacing !== undefined ? { lineSpacing: spacing.lineSpacing, lineSpacingRule: 'auto' as const } : {}),
  };
}
