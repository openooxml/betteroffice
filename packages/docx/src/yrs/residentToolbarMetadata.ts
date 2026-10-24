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

function symbolChar(value: unknown): boolean {
  if (typeof value !== 'string' || !/^\+?[0-9a-f]+$/i.test(value)) return false;
  const codepoint = Number.parseInt(value, 16);
  return codepoint <= 0x10ffff && (codepoint < 0xd800 || codepoint > 0xdfff);
}

/** The metadata a source open reports, read from a parsed document: its font table and every font its values name. */
export function readParsedToolbarMetadata(document: Document): DocxToolbarMetadata {
  const fonts = new Set<string>();
  const add = (name: unknown) => {
    if (typeof name === 'string' && name.trim()) fonts.add(name.trim());
  };
  for (const font of document.package.fontTable?.fonts ?? []) {
    add(font.name);
    add(font.altName);
  }
  const seen = new Set<object>();
  const pending: unknown[] = [document];
  while (pending.length > 0) {
    const value = pending.pop();
    if (typeof value !== 'object' || value === null || seen.has(value)) continue;
    seen.add(value);
    if (Array.isArray(value)) {
      for (const entry of value) pending.push(entry);
      continue;
    }
    if (value instanceof Map) {
      for (const entry of value.values()) pending.push(entry);
      continue;
    }
    if (Object.getPrototypeOf(value) !== Object.prototype) continue;
    const record = value as Record<string, unknown>;
    if (record.type === 'symbol' && symbolChar(record.char)) add(record.font);
    for (const key in record) {
      const entry = record[key];
      if (FONT_KEYS.has(key)) {
        if (typeof entry === 'string') add(entry);
        else if (typeof entry === 'object' && entry !== null) for (const slot of FONT_SLOTS) add((entry as Record<string, unknown>)[slot]);
      }
      pending.push(entry);
    }
  }
  return { styles: resolveParagraphStyleOptions(document.package.styles?.styles), fontFamilies: [...fonts].sort() };
}
