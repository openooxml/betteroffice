/**
 * Headless content-control discovery: DOCX bytes in, the document's content controls out. The Rust
 * engine does the reading; no session, editor or DOM is involved, and control ids and anchors
 * address the returned snapshot only.
 */

import type {
  DocxContentControlQuery,
  DocxContentControlReadFailure,
  DocxContentControlsOptions,
  DocxContentControlsSnapshot,
} from '../yrs/contentControls';

/** A read the engine refused as data: unusable options or a limit. */
export class DocxContentControlsError extends Error {
  readonly failure: DocxContentControlReadFailure;

  constructor(failure: DocxContentControlReadFailure) {
    super(failure.message);
    this.name = 'DocxContentControlsError';
    this.failure = failure;
  }
}

type Outcome = { ok: true; content: DocxContentControlsSnapshot } | {
  ok: false;
  failure: DocxContentControlReadFailure;
};

async function edit() {
  const wasm = await import('../wasm/edit');
  await wasm.preloadEditWasm();
  return wasm;
}

function unwrap(json: string): DocxContentControlsSnapshot {
  const outcome = JSON.parse(json) as Outcome;
  if (!outcome.ok) throw new DocxContentControlsError(outcome.failure);
  return outcome.content;
}

/**
 * Lists the content controls of DOCX bytes in document order. Throws
 * {@link DocxContentControlsError} for unusable options or an exceeded limit and an `Error` for
 * bytes that are not a readable DOCX.
 */
export async function listDocxContentControls(
  bytes: Uint8Array,
  options: DocxContentControlsOptions = {}
): Promise<DocxContentControlsSnapshot> {
  const wasm = await edit();
  return unwrap(wasm.listDocxContentControlsJson(bytes, JSON.stringify(options)));
}

/** The content controls of DOCX bytes that match `query` exactly; none or several. */
export async function findDocxContentControls(
  bytes: Uint8Array,
  query: DocxContentControlQuery,
  options: DocxContentControlsOptions = {}
): Promise<DocxContentControlsSnapshot> {
  const wasm = await edit();
  return unwrap(
    wasm.findDocxContentControlsJson(bytes, JSON.stringify(query), JSON.stringify(options))
  );
}
