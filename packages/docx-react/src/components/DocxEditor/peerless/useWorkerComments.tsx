import { useCallback, useLayoutEffect, useMemo, useState, type Dispatch, type SetStateAction } from 'react';
import type { CommentCallbacks } from '../../../hooks/useCommentSidebarItems';
import { Tooltip } from '../../ui/Tooltip';
import { MaterialSymbol } from '../../ui/Icons';
import type { createPeerlessShellController, WorkerCommentAction, WorkerCommentDraft } from './shellController';

export function useWorkerComments(options: {
  controller: ReturnType<typeof createPeerlessShellController>;
  documentKey: string | undefined;
  readOnly: boolean;
  setOpen: Dispatch<SetStateAction<boolean>>;
  setExpanded: Dispatch<SetStateAction<string | null>>;
  selectionAnchor(): { y: number; right: number } | null;
  focus(): void;
  onError(error: Error): void;
}) {
  const { controller, documentKey, readOnly, setOpen, setExpanded, selectionAnchor, focus, onError } = options;
  const [draft, setDraft] = useState<{ draft: WorkerCommentDraft; y: number } | null>(null);
  useLayoutEffect(() => setDraft(null), [documentKey]);
  const beginAddComment = useCallback(async () => {
    try {
      const next = await controller.beginComment();
      const anchor = selectionAnchor();
      if (!next || !anchor) return;
      setOpen(true);
      setDraft({ draft: next, y: anchor.y });
    } catch (error) { onError(error instanceof Error ? error : new Error(String(error))); }
  }, [controller, onError, selectionAnchor, setOpen]);
  const callbacks = useMemo<CommentCallbacks>(() => {
    const run = async (action: WorkerCommentAction) => {
      try {
        const result = await controller.comment(action);
        if (!result.ok) onError(new Error(result.failure.message));
        return result.ok;
      } catch (error) { onError(error instanceof Error ? error : new Error(String(error))); return false; }
    };
    return {
      onCommentReply: (id, text) => { void run({ kind: 'reply', id, text }); },
      onCommentResolve: (id) => { void run({ kind: 'resolve', id }); setExpanded(item => item === `comment-${id}` ? null : item); },
      onCommentUnresolve: (id) => { void run({ kind: 'reopen', id }); },
      onCommentDelete: (id) => { void run({ kind: 'delete', id }); },
      onAddComment: async (text) => {
        if (!draft || !await run({ kind: 'add', draft: draft.draft, text })) return false;
        setDraft(current => current === draft ? null : current);
        return true;
      },
      onCancelAddComment: () => { setDraft(null); focus(); },
    };
  }, [controller, draft, focus, onError, setExpanded]);
  return { callbacks, beginAddComment, adding: draft !== null, composerHidden: draft !== null && readOnly, addCommentY: draft?.y ?? null };
}

export function WorkerCommentButton({ top, left, onAdd }: { top: number; left: number; onAdd(): void }) {
  return <Tooltip content="Add comment" side="bottom" delayMs={300}>
    <button type="button" aria-label="Add comment" onMouseDown={(event) => { event.preventDefault(); event.stopPropagation(); onAdd(); }}
      style={{ position: 'absolute', top, left, transform: 'translate(-50%, -50%)', zIndex: 50, width: 28, height: 28, borderRadius: 6,
        border: '1px solid var(--doc-focus-ring)', backgroundColor: 'var(--doc-surface)', color: 'var(--doc-primary)', cursor: 'pointer',
        display: 'flex', alignItems: 'center', justifyContent: 'center', boxShadow: '0 1px 3px var(--doc-shadow)' }}>
      <MaterialSymbol name="add_comment" size={16} />
    </button>
  </Tooltip>;
}
