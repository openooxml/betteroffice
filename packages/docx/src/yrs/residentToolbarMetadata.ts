import type { StyleDefinitions } from '../types/styles';
import { resolveParagraphStyleOptions, type ResolvedStyleOption } from '../utils/stylePreview';

export interface DocxToolbarMetadata {
  styles: ResolvedStyleOption[];
  fontFamilies: string[];
}

export function readResidentToolbarMetadata(hostJson: string): DocxToolbarMetadata {
  const host = JSON.parse(hostJson) as {
    referencedFonts: string[];
    envelope: { wireVersion: number; document: { package: { styles?: StyleDefinitions } } };
  };
  if (host.envelope?.wireVersion !== 1 || !Array.isArray(host.referencedFonts) || !host.referencedFonts.every(font => typeof font === 'string')) throw new TypeError('Invalid toolbar metadata');
  return { styles: resolveParagraphStyleOptions(host.envelope.document.package.styles?.styles), fontFamilies: [...new Set(host.referencedFonts)].sort() };
}
