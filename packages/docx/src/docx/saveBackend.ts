/**
 * Internal save-backend seam. Rust is the sole implementation after S13
 * retirement; keeping the selector makes a future backend experiment local.
 */
export type DocxSaveBackend = 'rust';

// Internal-only parameter channel: no public save-options expansion.
const optionBackends = new WeakMap<object, DocxSaveBackend>();

/** Select a save backend for one internal call without changing public options. */
export function withDocxSaveBackend<T extends object>(
  options: T,
  backend: DocxSaveBackend
): T {
  optionBackends.set(options, backend);
  return options;
}

export function saveBackendFor(options: object): DocxSaveBackend {
  const option = optionBackends.get(options);
  if (option) return option;
  const configured =
    typeof process === 'undefined'
      ? undefined
      : (process.env.OPENOOXML_DOCX_SAVE_BACKEND ?? process.env.DOCX_SAVE_BACKEND);
  if (configured === 'rust') return configured;
  return DEFAULT_DOCX_SAVE_BACKEND;
}

// S13 gate passed and the TypeScript oracle has been retired.
export const DEFAULT_DOCX_SAVE_BACKEND: DocxSaveBackend = 'rust';
