import type { BundledFontProvider } from '@betteroffice/docx/layout';
import type { FontDefinition } from '@betteroffice/docx/utils';

const hostedFamiliesKey = Symbol.for('@betteroffice/docx/hostedFontFamilies');

export function setHostedFontFamilies(source: object, fonts: ReadonlyArray<FontDefinition> | undefined): void {
  Object.assign(source, { [hostedFamiliesKey]: new Set((fonts ?? []).map(font => font.family.trim().toLowerCase())) });
}

export function hostedFontKey(source: object | undefined): string {
  const families = (source as { [hostedFamiliesKey]?: ReadonlySet<string> } | undefined)?.[hostedFamiliesKey];
  return families ? [...families].sort().join('\n') : '';
}

export function hostedFontProvider(
  provider: BundledFontProvider | undefined,
  fonts: ReadonlyArray<FontDefinition> | undefined
): BundledFontProvider | undefined {
  if (provider || !fonts?.length) return provider;
  const implicit = { resolve: () => undefined };
  setHostedFontFamilies(implicit, fonts);
  return implicit;
}
