/**
 * Public DOCX parser facade.
 *
 * Rust S9 owns ZIP/XML parsing. TypeScript only normalizes host inputs,
 * rehydrates the public `Document`, reports progress, and registers fonts.
 * @packageDocumentation
 * @public
 */

import type { Document, DocumentBody, StyleDefinitions, Theme } from '../types/document';
import type { DocxInput } from '../utils/docxInput';
import { toArrayBuffer } from '../utils/docxInput';
import { loadEmbeddedFonts } from '../utils/embeddedFonts';
import { loadFontsWithMapping } from '../utils/fontLoader';
import { preloadOpcWasm } from './wasm';
import { preloadParseWasm } from './parseWasm';
import {
  parseDocumentWithRust,
  type RustS9ParseOptions,
  type RustS9Result,
} from './rustParseFacade';

export type ProgressCallback = (stage: string, percent: number) => void;

export interface ParseOptions {
  onProgress?: ProgressCallback;
  preloadFonts?: boolean;
  parseHeadersFooters?: boolean;
  parseNotes?: boolean;
}

export async function parseDocx(input: DocxInput, options: ParseOptions = {}): Promise<Document> {
  // The container + parser wasm are external assets; this async entry is where
  // browsers load them (Node/Bun sync-inits from disk on first use instead).
  await preloadOpcWasm();
  await preloadParseWasm();
  const buffer = input instanceof ArrayBuffer ? input : await toArrayBuffer(input);
  const onProgress = options.onProgress ?? (() => {});
  try {
    onProgress('Extracting DOCX...', 0);
    const parsed = parseDocumentWithRust(buffer, rustOptions(options));
    emitProgress(options, onProgress);
    await loadHostFonts(parsed, options.preloadFonts ?? true, onProgress);
    onProgress('Assembling document...', 95);
    onProgress('Complete', 100);
    return parsed.document;
  } catch (error) {
    const message = error instanceof Error ? error.message : String(error);
    if (message.startsWith('Failed to parse DOCX: ')) throw error;
    const mapped = mapRustParseError(message);
    const publicMessage =
      mapped.code === 'Container' ? message.replace(/^container error: /, '') : message;
    const wrapped = new Error(`Failed to parse DOCX: ${publicMessage}`, { cause: error });
    Object.assign(wrapped, mapped);
    throw wrapped;
  }
}

function rustOptions(options: ParseOptions): RustS9ParseOptions {
  return {
    parseHeadersFooters: options.parseHeadersFooters ?? true,
    parseNotes: options.parseNotes ?? true,
    detectVariables: false,
    includeCanonical: false,
  };
}

function emitProgress(options: ParseOptions, onProgress: ProgressCallback): void {
  onProgress('Extracted DOCX', 10);
  onProgress('Parsing relationships...', 10);
  onProgress('Parsed relationships', 15);
  onProgress('Parsing theme...', 15);
  onProgress('Parsed theme', 20);
  onProgress('Parsing styles...', 20);
  onProgress('Parsed styles', 30);
  onProgress('Parsing numbering...', 30);
  onProgress('Parsed numbering', 35);
  onProgress('Processing media files...', 35);
  onProgress('Processed media', 40);
  onProgress('Parsing document body...', 40);
  onProgress('Parsed document body', 55);
  if (options.parseHeadersFooters ?? true) {
    onProgress('Parsing headers/footers...', 55);
    onProgress('Parsed headers/footers', 65);
  } else {
    onProgress('Skipping headers/footers', 65);
  }
  if (options.parseNotes ?? true) {
    onProgress('Parsing footnotes/endnotes...', 65);
    onProgress('Parsed footnotes/endnotes', 75);
  } else {
    onProgress('Skipping footnotes/endnotes', 75);
  }
  onProgress('Parsing comments...', 75);
  onProgress('Parsed comments', 80);
}

async function loadHostFonts(
  parsed: RustS9Result,
  preloadFonts: boolean,
  onProgress: ProgressCallback
): Promise<void> {
  if (!preloadFonts) {
    onProgress('Skipping font loading', 95);
    return;
  }
  onProgress('Loading fonts...', 80);
  await loadDocumentFonts(
    parsed.document.package.theme ?? null,
    parsed.document.package.styles,
    parsed.document.package.document
  );
  await loadEmbeddedFonts(
    parsed.document.package.fontTable,
    parsed.embeddedFonts,
    parsed.fontTableRelationshipsXml
  );
  onProgress('Loaded fonts', 95);
}

async function loadDocumentFonts(
  theme: Theme | null,
  styles: StyleDefinitions | undefined,
  body: DocumentBody
): Promise<void> {
  const fonts = new Set<string>();
  if (theme?.fontScheme?.majorFont?.latin) fonts.add(theme.fontScheme.majorFont.latin);
  if (theme?.fontScheme?.minorFont?.latin) fonts.add(theme.fontScheme.minorFont.latin);
  if (styles?.docDefaults?.rPr?.fontFamily?.ascii) {
    fonts.add(styles.docDefaults.rPr.fontFamily.ascii);
  }
  for (const style of styles?.styles ?? []) {
    if (style.rPr?.fontFamily?.ascii) fonts.add(style.rPr.fontFamily.ascii);
    if (style.rPr?.fontFamily?.hAnsi) fonts.add(style.rPr.fontFamily.hAnsi);
  }
  for (const block of body.content) {
    if (block.type !== 'paragraph') continue;
    for (const item of block.content) {
      if (item.type !== 'run') continue;
      if (item.formatting?.fontFamily?.ascii) fonts.add(item.formatting.fontFamily.ascii);
      if (item.formatting?.fontFamily?.hAnsi) fonts.add(item.formatting.fontFamily.hAnsi);
    }
  }
  if (fonts.size === 0) return;
  try {
    await loadFontsWithMapping([...fonts]);
  } catch (error) {
    console.warn('Failed to load some fonts:', error);
  }
}

function mapRustParseError(message: string): { code: string; part?: string; offset?: number } {
  const part =
    /(?:unsafe XML|resource limit .* exceeded|malformed XML|invalid relationship) in ([^:]+)(?::| at)/.exec(
      message
    )?.[1];
  const offset = / at byte (\d+):/.exec(message)?.[1];
  const code = message.startsWith('unsafe XML')
    ? 'UnsafeXml'
    : message.startsWith('resource limit')
      ? 'ResourceLimit'
      : message.startsWith('malformed XML')
        ? 'MalformedXml'
        : message.startsWith('invalid relationship')
          ? 'Relationship'
          : message.startsWith('container error')
            ? 'Container'
            : message.startsWith('wire.')
              ? 'Wire'
              : 'Parse';
  return {
    code,
    ...(part ? { part } : {}),
    ...(offset ? { offset: Number(offset) } : {}),
  };
}

export async function quickParseDocx(buffer: ArrayBuffer): Promise<Document> {
  return parseDocx(buffer, {
    preloadFonts: false,
    parseHeadersFooters: false,
    parseNotes: false,
  });
}

export async function fullParseDocx(
  buffer: ArrayBuffer,
  onProgress?: ProgressCallback
): Promise<Document> {
  return parseDocx(buffer, {
    onProgress,
    preloadFonts: true,
    parseHeadersFooters: true,
    parseNotes: true,
  });
}

export async function getDocxSummary(buffer: ArrayBuffer): Promise<{
  hasDocument: boolean;
  hasStyles: boolean;
  hasTheme: boolean;
  hasNumbering: boolean;
  headerCount: number;
  footerCount: number;
  mediaCount: number;
}> {
  const document = await parseDocx(buffer, { preloadFonts: false });
  const pkg = document.package;
  const mediaPaths = new Set([...(pkg.media?.values() ?? [])].map((media) => media.path));
  return {
    hasDocument: true,
    hasStyles: pkg.styles !== undefined,
    hasTheme: pkg.theme !== undefined,
    hasNumbering: pkg.numbering !== undefined,
    headerCount: pkg.headers?.size ?? 0,
    footerCount: pkg.footers?.size ?? 0,
    mediaCount: mediaPaths.size,
  };
}
