import type { Document } from '../types/document';
import type { ResidentCommandInitRequest } from './residentCommandProtocol';

export interface ResidentSourcePreviewOptions extends Partial<Pick<ResidentCommandInitRequest,
  'fonts' | 'fontRegistrationSlots' | 'layoutInput' | 'extras' | 'displayWindow'>> {
  pages?: number;
}

export type ResidentSourceOpenRequest = (
  | { document: Document; bytes?: never }
  | { bytes: Uint8Array; document?: never }
) & {
  preview?: boolean | ResidentSourcePreviewOptions;
  /** @internal No preview here: the first command frame shows the font-ready first pages instead. */
  deferPreview?: boolean;
  mediaTokens?: boolean;
  heapLimitBytes?: number;
  generation?: string;
  digest?: string;
};

export function validateResidentSource(source: ResidentSourceOpenRequest): void {
  if (!source || (source.document === undefined) === (source.bytes === undefined)) {
    throw new TypeError('DOCX source requires exactly one of document or bytes');
  }
  if (source.bytes !== undefined && !(source.bytes instanceof Uint8Array)) {
    throw new TypeError('DOCX source bytes must be a Uint8Array');
  }
  if (source.document !== undefined && !Array.isArray(source.document?.package?.document?.content)) {
    throw new TypeError('DOCX source document must contain a document body');
  }
  if (source.preview !== undefined && typeof source.preview !== 'boolean') {
    if (!source.preview || typeof source.preview !== 'object') throw new TypeError('Invalid DOCX preview options');
    const pages = source.preview?.pages ?? 1;
    if (!Number.isSafeInteger(pages) || pages < 1) throw new TypeError('Invalid DOCX preview pages');
  }
}

export function residentSourceMetadata(document: Document): Document {
  const pkg = document.package;
  return {
    ...document, originalBuffer: undefined,
    package: {
      ...pkg,
      document: { ...pkg.document, content: [], sections: pkg.document.sections?.map(section => ({ ...section, content: [] })) },
      headers: pkg.headers && new Map([...pkg.headers].map(([id, part]) => [id, { ...part, content: [] }])),
      footers: pkg.footers && new Map([...pkg.footers].map(([id, part]) => [id, { ...part, content: [] }])),
      footnotes: pkg.footnotes?.map(note => ({ ...note, content: [] })),
      endnotes: pkg.endnotes?.map(note => ({ ...note, content: [] })),
    },
  };
}

export function tokenizeResidentDocument(document: Document, sources: Map<string, string>): Document {
  const copy = structuredClone(document);
  const tokens = new Map<string, string>();
  const visit = (value: unknown): void => {
    if (!value || typeof value !== 'object' || value instanceof ArrayBuffer || ArrayBuffer.isView(value)) return;
    if (value instanceof Map) {
      for (const item of value.values()) visit(item);
      return;
    }
    for (const [key, child] of Object.entries(value)) {
      if (key === 'src' && typeof child === 'string' && child.startsWith('data:')) {
        let token = tokens.get(child);
        if (!token) {
          token = `media:${tokens.size}`;
          tokens.set(child, token);
          sources.set(token, child);
        }
        (value as Record<string, unknown>)[key] = token;
      } else visit(child);
    }
  };
  visit(copy.package);
  return copy;
}
