/**
 * Zip-container READ facade for the DOCX open path.
 *
 * Backed by the inlined docx-container wasm (see ./wasm): every non-directory
 * entry is inflated in Rust — with the decompression-bomb and path-traversal
 * guards enforced there by construction — and exposed as a normalized, eager
 * read view. Consumers pull parts by path.
 *
 * Read path only. The write/rezip builder assembles a parts map and hands it
 * to the same wasm container via `rezipContainer` (see ./rezip).
 */

import { unzipContainer } from './wasm';

/**
 * Eager, normalized read view of an OPC (DOCX) container. All entries are
 * already materialized; lookups are synchronous.
 */
export interface ZipContainerReader {
  paths(): string[];
  file(path: string): Uint8Array | null;
  text(path: string): string | null;
}

// build a reader over an already-materialized part map.
function readerFromEntries(entries: Map<string, Uint8Array>): ZipContainerReader {
  // ignoreBOM keeps a leading BOM in the string, matching a literal utf-8 decode
  // (the default TextDecoder would strip it and diverge on round-trip).
  const decoder = new TextDecoder('utf-8', { ignoreBOM: true });
  return {
    paths: () => [...entries.keys()],
    file: (path) => entries.get(path) ?? null,
    text: (path) => {
      const bytes = entries.get(path);
      return bytes ? decoder.decode(bytes) : null;
    },
  };
}

// reads a DOCX container via the wasm backend. requires a CSP permitting wasm
// compilation (`wasm-unsafe-eval`).
export function readDocxContainer(buffer: ArrayBuffer): ZipContainerReader {
  const parts = unzipContainer(new Uint8Array(buffer));
  return readerFromEntries(new Map(Object.entries(parts)));
}
