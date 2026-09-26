import type {
  DocxEditFailure,
  DocxEditRefusal,
  DocxEditRequest,
  DocxEditResult,
  YrsSession,
} from '@betteroffice/docx/yrs';
import type { PagedEditorRef } from './PagedEditor';
import type { EditorMode } from './internals/editing-modes';

export type EditorFlush =
  | { ok: true; editor: PagedEditorRef; session: YrsSession }
  | { ok: false; code: 'editor-unavailable' | 'document-replaced' | 'input-failed'; error: Error };

export type EditorFlushFailure = Extract<EditorFlush, { ok: false }>;

/**
 * Commits pending input. Handles are rebuilt on layout changes, so only the session identifies
 * the document the flush started on.
 */
export async function flushEditorInput(
  pagedEditorRef: React.RefObject<PagedEditorRef | null>
): Promise<EditorFlush> {
  const editor = pagedEditorRef.current;
  const session = editor?.getYrsSession();
  if (!editor || !session) {
    return {
      ok: false,
      code: 'editor-unavailable',
      error: new Error('The editor input is unavailable'),
    };
  }
  try {
    await editor.flushPendingInput();
  } catch (error) {
    return {
      ok: false,
      code:
        pagedEditorRef.current?.getYrsSession() === session ? 'input-failed' : 'document-replaced',
      error: error instanceof Error ? error : new Error(String(error)),
    };
  }
  const flushed = pagedEditorRef.current;
  if (!flushed || flushed.getYrsSession() !== session) {
    return {
      ok: false,
      code: 'document-replaced',
      error: new Error('The document changed while flushing input'),
    };
  }
  return { ok: true, editor: flushed, session };
}

/** Flushes pending input and returns the current handle; throws when that fails. */
export async function flushedSession(
  pagedEditorRef: React.RefObject<PagedEditorRef | null>
): Promise<{ editor: PagedEditorRef; session: YrsSession }> {
  const flushed = await flushEditorInput(pagedEditorRef);
  if (!flushed.ok) throw flushed.error;
  return flushed;
}

function refusal(session: YrsSession, failure: DocxEditFailure): DocxEditRefusal {
  return { ok: false, version: session.version(), failure };
}

/** Refuses writes the editor's mode does not allow; suggesting mode needs `suggest` on every step. */
export function modeRefusal(
  session: YrsSession,
  mode: EditorMode,
  request: DocxEditRequest
): DocxEditRefusal | null {
  if (mode === 'viewing') {
    return refusal(session, { code: 'read-only', message: 'The editor is read-only' });
  }
  const direct = request.steps.findIndex((step) => !step.suggest);
  if (mode === 'suggesting' && direct >= 0) {
    return refusal(session, {
      code: 'invalid-step',
      stepIndex: direct,
      message: 'Suggesting mode records every step as a tracked change; supply suggest metadata',
    });
  }
  return null;
}

/**
 * The editor's batch path: the mode gate, an input flush, `authorize` immediately before the
 * mutation, which runs inside `commit`, then one refresh of every changed story. A refusal never
 * rolls back flushed typing.
 */
export async function applyEditBatch<Refusal = never>(
  pagedEditorRef: React.RefObject<PagedEditorRef | null>,
  mode: () => EditorMode,
  request: DocxEditRequest,
  authorize?: () => Refusal | null,
  commit: <T>(write: () => T) => T = (write) => write()
): Promise<{ flush: EditorFlushFailure } | { result: DocxEditResult | Refusal }> {
  const session = pagedEditorRef.current?.getYrsSession();
  if (!session) {
    return {
      flush: {
        ok: false,
        code: 'editor-unavailable',
        error: new Error('The editor input is unavailable'),
      },
    };
  }
  const early = modeRefusal(session, mode(), request);
  if (early) return { result: early };
  const flushed = await flushEditorInput(pagedEditorRef);
  if (!flushed.ok) return { flush: flushed };
  if (flushed.session !== session || pagedEditorRef.current?.getYrsSession() !== session) {
    return {
      flush: {
        ok: false,
        code: 'document-replaced',
        error: new Error('The document changed while flushing input'),
      },
    };
  }
  const denied = authorize?.() ?? null;
  if (denied) return { result: denied };
  const refused = modeRefusal(session, mode(), request);
  if (refused) return { result: refused };
  const result = commit(() => session.applyEdits(request));
  if (result.ok && result.applied) {
    try {
      flushed.editor.syncYrsInputState(true, result.changedStories);
    } catch (error) {
      console.error('[DocxEditor] refreshing after an applied edit batch failed', error);
    }
  }
  return { result };
}
