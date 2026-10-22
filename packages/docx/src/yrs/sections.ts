import type { SectionProperties } from '../types/document';

export type SectionStartKind = 'nextPage' | 'continuous' | 'evenPage' | 'oddPage';
export interface SectionPropertyPatch {
  orientation?: 'portrait' | 'landscape';
  pageWidth?: number;
  pageHeight?: number;
  margins?: Partial<Record<'top' | 'bottom' | 'left' | 'right' | 'header' | 'footer' | 'gutter', number>>;
}
export interface SectionInfo {
  closingParagraphId: string;
  firstParagraph: number;
  lastParagraph: number;
  authored: SectionProperties;
  resolved: SectionProperties;
}
export class DocxSectionError extends Error {
  constructor(readonly code: string, message: string) {
    super(message);
    this.name = 'DocxSectionError';
  }
}

export function sectionPropertiesEqual(left: unknown, right: unknown): boolean {
  const sorted = (value: unknown): unknown => Array.isArray(value) ? value.map(sorted)
    : value && typeof value === 'object' ? Object.fromEntries(Object.entries(value).filter(([, value]) => value !== undefined).sort(([a], [b]) => a.localeCompare(b)).map(([key, value]) => [key, sorted(value)])) : value;
  return JSON.stringify(sorted(left)) === JSON.stringify(sorted(right));
}
