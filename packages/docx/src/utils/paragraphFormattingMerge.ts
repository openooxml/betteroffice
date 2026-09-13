import type { ParagraphFormatting } from '../types/document';
import { mergeTextFormatting } from './textFormattingMerge';

export function mergeParagraphFormatting(
  target: ParagraphFormatting | undefined,
  source: ParagraphFormatting | undefined
): ParagraphFormatting | undefined {
  if (!source) return target;
  if (!target) return source ? { ...source } : undefined;

  const result = { ...target };

  for (const key of Object.keys(source) as (keyof ParagraphFormatting)[]) {
    const value = source[key];
    if (value !== undefined) {
      if (key === 'runProperties') {
        result.runProperties = mergeTextFormatting(result.runProperties, source.runProperties);
      } else if (key === 'borders' || key === 'numPr' || key === 'frame') {
        const baseValue = result[key] as Record<string, unknown> | undefined;
        const sourceValue = value as Record<string, unknown> | undefined;
        (result as Record<string, unknown>)[key] = {
          ...(baseValue || {}),
          ...(sourceValue || {}),
        };
      } else if (key === 'tabs' && Array.isArray(value)) {
        result.tabs = [...value];
      } else {
        (result as Record<string, unknown>)[key] = value;
      }
    }
  }

  return result;
}
