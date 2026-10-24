import type { Document } from '../types/document';
import type { DocxCommandInitRequest } from './residentCommandProtocol';

export interface DocxSourcePreviewOptions extends Partial<Pick<DocxCommandInitRequest,
  'fonts' | 'fontRegistrationSlots' | 'layoutInput' | 'extras' | 'displayWindow'>> {
  pages?: number;
}

export type DocxSourceOpenRequest = (
  | { document: Document; bytes?: never }
  | { bytes: Uint8Array; document?: never }
) & {
  preview?: boolean | DocxSourcePreviewOptions;
  /** Opens a display-only cut of the first body blocks of `bytes`; a later whole open replaces it. */
  cut?: { blocks: number; paragraphBudget?: number };
  mediaTokens?: boolean;
  heapLimitBytes?: number;
  generation?: string;
  digest?: string;
  /**
   * One replica seeds from bytes; joiners pass `initialUpdate`. Two seeds merge by story last-writer-wins.
   * `clientClock` is the highest clock of `clientId` the caller has held. A seed would write that id's clocks again
   * and a joined state behind it would reuse later ones, so either opens under a fresh client id.
   */
  collaboration?: { clientId: number; initialUpdate?: Uint8Array; clientClock?: number };
  /** Yrs state restored over the unseeded package of `bytes`, such as a command checkpoint. */
  state?: Uint8Array;
};

export function validateResidentSource(source: DocxSourceOpenRequest): void {
  if (!source || (source.document === undefined) === (source.bytes === undefined)) {
    throw new TypeError('DOCX source requires exactly one of document or bytes');
  }
  if (source.bytes !== undefined && !(source.bytes instanceof Uint8Array)) {
    throw new TypeError('DOCX source bytes must be a Uint8Array');
  }
  if (source.document !== undefined && !Array.isArray(source.document?.package?.document?.content)) {
    throw new TypeError('DOCX source document must contain a document body');
  }
  if (source.cut !== undefined && (source.bytes === undefined || !Number.isSafeInteger(source.cut?.blocks) || source.cut.blocks < 1 ||
    (source.cut.paragraphBudget !== undefined && (!Number.isSafeInteger(source.cut.paragraphBudget) || source.cut.paragraphBudget < 1)))) {
    throw new TypeError('Invalid DOCX source cut');
  }
  if (source.state !== undefined && (!(source.state instanceof Uint8Array) || source.bytes === undefined ||
    source.cut !== undefined || source.preview)) {
    throw new TypeError('Invalid DOCX source state');
  }
  if (source.preview !== undefined && typeof source.preview !== 'boolean') {
    if (!source.preview || typeof source.preview !== 'object') throw new TypeError('Invalid DOCX preview options');
    const pages = source.preview?.pages ?? 1;
    if (!Number.isSafeInteger(pages) || pages < 1) throw new TypeError('Invalid DOCX preview pages');
  }
  const collaboration = source.collaboration;
  if (collaboration !== undefined && (!collaboration || !Number.isSafeInteger(collaboration.clientId) || collaboration.clientId < 0 ||
    (collaboration.initialUpdate !== undefined && !(collaboration.initialUpdate instanceof Uint8Array)) ||
    (collaboration.clientClock !== undefined && (!Number.isSafeInteger(collaboration.clientClock) || collaboration.clientClock < 0)))) {
    throw new TypeError('Invalid DOCX collaboration options');
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
