/** Runs a comparison in a private session; Rust applies, verifies and finalizes the result. */

import type {
  DocxCompareDiagnostic,
  DocxComparedChange,
  DocxCompareOptions,
  DocxCompareResult,
} from '../docx/compare';
import { writeDocumentWithRust } from '../docx/rustSaveFacade';
import { sessionInternals } from './sessionInternals';
import { type RevisionNumbers, yrsBodyToDocumentWithRevisionIds } from './yrsToDocument';

const COMPARE_CLIENT_ID = 1;

interface CompareSave {
  partSha256: string;
  paragraphs: Array<{ path: number[]; block: number }>;
  revisionIds: Record<string, RevisionNumbers>;
  seed: string;
  now: string;
}

type Final =
  | { ok: true; changes: DocxComparedChange[]; diagnostics: DocxCompareDiagnostic[] }
  | { ok: false; diagnostics: DocxCompareDiagnostic[] };

type Prepared = Final & { noop?: boolean; save?: CompareSave };

function exactBuffer(bytes: Uint8Array): ArrayBuffer {
  const buffer = new ArrayBuffer(bytes.byteLength);
  new Uint8Array(buffer).set(bytes);
  return buffer;
}

export async function compareDocxInSession(
  original: Uint8Array,
  revised: Uint8Array,
  options: DocxCompareOptions
): Promise<DocxCompareResult> {
  // A static import would move the facade, and its relative worker URL, out of the yrs entry.
  const { createYrsSession } = await import('./index');
  const session = await createYrsSession({ clientId: COMPARE_CLIENT_ID });
  try {
    const internals = sessionInternals(session);
    const prepared = JSON.parse(
      internals.compareDocx(original, revised, JSON.stringify(options))
    ) as Prepared;
    if (!prepared.ok) return { ok: false, diagnostics: prepared.diagnostics };
    if (!prepared.save) {
      return {
        ok: true,
        docx: original.slice(),
        changes: prepared.changes,
        diagnostics: prepared.diagnostics,
      };
    }
    const save = prepared.save;
    const base = session.materializeDocx();
    if (!base) throw new Error('the compared session holds no document');
    let saved: Uint8Array;
    try {
      const projected = yrsBodyToDocumentWithRevisionIds(
        session,
        base,
        new Map(Object.entries(save.revisionIds))
      );
      const result = await writeDocumentWithRust(
        projected,
        exactBuffer(original),
        { updateModifiedDate: false },
        {
          changedParaIds: [],
          sourceParagraphs: { partSha256: save.partSha256, paragraphs: save.paragraphs },
        },
        { seed: save.seed, now: save.now }
      );
      saved = new Uint8Array(result.buffer);
    } catch (error) {
      const message = error instanceof Error ? error.message : String(error);
      return JSON.parse(internals.failComparedDocx(message)) as Final & { ok: false };
    }
    const finished = JSON.parse(internals.finishComparedDocx(saved)) as Final;
    return finished.ok ? { ...finished, docx: saved } : finished;
  } finally {
    session.destroy();
  }
}
