import type { Document } from '../types/document';
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

const FONT_KEYS = new Set(['fontFamily', 'listMarkerFontFamily', 'markerFontFamily']);
const FONT_SLOTS = ['ascii', 'hAnsi', 'eastAsia', 'cs'] as const;

/** The metadata a source open reports, read from a parsed document: its font table and every font a value names. */
export function readParsedToolbarMetadata(document: Document): DocxToolbarMetadata {
  const fonts = new Set<string>();
  const add = (name: unknown) => {
    if (typeof name === 'string' && name.trim()) fonts.add(name.trim());
  };
  for (const font of document.package.fontTable?.fonts ?? []) {
    add(font.name);
    add(font.altName);
  }
  const pending: unknown[] = [document];
  while (pending.length > 0) {
    const value = pending.pop();
    if (Array.isArray(value)) {
      for (const entry of value) pending.push(entry);
      continue;
    }
    if (typeof value !== 'object' || value === null || Object.getPrototypeOf(value) !== Object.prototype) continue;
    for (const [key, entry] of Object.entries(value)) {
      if (FONT_KEYS.has(key)) {
        if (typeof entry === 'string') add(entry);
        else if (typeof entry === 'object' && entry !== null) for (const slot of FONT_SLOTS) add((entry as Record<string, unknown>)[slot]);
      }
      pending.push(entry);
    }
  }
  return { styles: resolveParagraphStyleOptions(document.package.styles?.styles), fontFamilies: [...fonts].sort() };
}
