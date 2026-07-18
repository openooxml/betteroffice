import type { ParseOptions } from './parser';

/** Internal parser seam; Rust is the sole implementation after S9 retirement. */
export type DocxParseBackend = 'rust';

// Internal-only parameter channel: no ParseOptions/public API expansion.
const optionBackends = new WeakMap<object, DocxParseBackend>();

/** Select a parser backend for an internal call without changing ParseOptions. */
export function withDocxParseBackend<T extends ParseOptions>(
  options: T,
  backend: DocxParseBackend
): T {
  optionBackends.set(options, backend);
  return options;
}

export function parseBackendFor(options: ParseOptions): DocxParseBackend {
  const option = optionBackends.get(options);
  if (option) return option;
  const configured =
    typeof process === 'undefined'
      ? undefined
      : (process.env.OPENOOXML_DOCX_PARSE_BACKEND ?? process.env.DOCX_PARSE_BACKEND);
  if (configured === 'rust') return configured;
  return DEFAULT_DOCX_PARSE_BACKEND;
}

// S9 gate passed and the TypeScript oracle has been retired.
export const DEFAULT_DOCX_PARSE_BACKEND: DocxParseBackend = 'rust';
