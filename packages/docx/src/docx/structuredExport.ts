/**
 * Headless structured export: DOCX bytes in, read-only structured content or Markdown out. The
 * Rust engine does the reading; no session, editor, DOM or font is involved, and anchors address
 * the returned snapshot only.
 */

import type {
  DocxExportFailure,
  DocxExportOptions,
  DocxMarkdownContent,
  DocxMarkdownOptions,
  DocxStructuredContent,
} from '../yrs/structuredExport';

/** An export the engine refused as data: unusable options. */
export class DocxExportError extends Error {
  readonly failure: DocxExportFailure;

  constructor(failure: DocxExportFailure) {
    super(failure.message);
    this.name = 'DocxExportError';
    this.failure = failure;
  }
}

type Outcome<T> = { ok: true; content: T } | { ok: false; failure: DocxExportFailure };

async function edit() {
  const wasm = await import('../wasm/edit');
  await wasm.preloadEditWasm();
  return wasm;
}

function unwrap<T>(json: string): T {
  const outcome = JSON.parse(json) as Outcome<T>;
  if (!outcome.ok) throw new DocxExportError(outcome.failure);
  return outcome.content;
}

/**
 * Exports DOCX bytes as structured content. Throws {@link DocxExportError} for unusable options
 * and an `Error` for bytes that are not a readable DOCX.
 */
export async function exportDocxStructured(
  bytes: Uint8Array,
  options: DocxExportOptions
): Promise<DocxStructuredContent> {
  const wasm = await edit();
  return unwrap(wasm.exportDocxStructuredJson(bytes, JSON.stringify(options)));
}

/** {@link exportDocxStructured} rendered as Markdown with source anchors. */
export async function exportDocxMarkdown(
  bytes: Uint8Array,
  options: DocxExportOptions
): Promise<DocxMarkdownContent> {
  const wasm = await edit();
  return unwrap(wasm.exportDocxMarkdownJson(bytes, JSON.stringify(options)));
}

/** Renders structured content (schema version 1) as Markdown. */
export async function renderDocxMarkdown(
  content: DocxStructuredContent,
  options: DocxMarkdownOptions = {}
): Promise<DocxMarkdownContent> {
  const wasm = await edit();
  return unwrap(wasm.renderDocxMarkdownJson(JSON.stringify(content), JSON.stringify(options)));
}
