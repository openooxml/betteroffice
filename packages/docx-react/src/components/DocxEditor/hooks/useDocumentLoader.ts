import { useCallback, useEffect, useRef, useState } from 'react';
import type { Document } from '@betteroffice/docx/types/document';
import type { Comment } from '@betteroffice/docx/types/content';
import { parseDocx } from '@betteroffice/docx/docx';
import {
  loadDocumentFonts,
  getRenderableDocumentFonts,
  getEmbeddedFontFamilies,
  type DocxInput,
} from '@betteroffice/docx/utils';
import type { FontOption } from '@betteroffice/docx/utils/fontOptions';
import type { UseHistoryReturn } from '../../../hooks/useHistory';
import type { PagedEditorRef } from '../PagedEditor';
import type { CommentIdAllocator } from '../commentFactories';

/**
 * Document lifecycle: load buffer / pre-parsed doc, react to
 * `documentBuffer` / `document` prop changes, and extract any baked-in
 * comments from the document model on initial load.
 *
 * State reset across the editor on a fresh load is heavy (~10 distinct
 * state setters across multiple hooks), so the parent assembles a
 * single `resetForNewDocument` callback and threads it in.
 */
export function useDocumentLoader({
  documentBuffer,
  initialDocument,
  externalContent,
  history,
  pagedEditorRef,
  setLoadingState,
  setComments,
  setShowCommentsSidebar,
  onError,
  resetForNewDocument,
  commentsLoadedRef,
  commentIdAllocator,
  setDocumentFonts,
}: {
  documentBuffer: DocxInput | null | undefined;
  initialDocument: Document | null | undefined;
  externalContent: boolean | undefined;
  history: UseHistoryReturn<Document | null>;
  pagedEditorRef: React.RefObject<PagedEditorRef | null>;
  // The full EditorState shape lives in the parent; we only need to flip
  // `isLoading` and `parseError`, so the parent exposes a focused callback.
  setLoadingState: (state: { isLoading: boolean; parseError: string | null }) => void;
  setComments: React.Dispatch<React.SetStateAction<Comment[]>>;
  setShowCommentsSidebar: React.Dispatch<React.SetStateAction<boolean>>;
  onError: ((error: Error) => void) | undefined;
  resetForNewDocument: () => void;
  // `resetForNewDocument` (declared earlier in the parent) needs to clear
  // this ref on every load. Lifted out of the hook for that reason.
  commentsLoadedRef: React.RefObject<boolean>;
  // Per-editor-instance ID allocator; seeded above the loaded doc's max ID.
  commentIdAllocator: CommentIdAllocator;
  // Fonts the document references that the browser can actually render
  // (embedded or system-resolved), surfaced in the picker's "Document fonts"
  // group.
  setDocumentFonts: (fonts: FontOption[]) => void;
}) {
  // The live history document changes after every edit, but yrs must only be
  // reseeded when a new source document is loaded. Keep that load boundary
  // separate so PagedEditor can replace its session without treating normal
  // edits as fresh documents.
  const [yrsSeedDocument, setYrsSeedDocument] = useState<Document | null>(
    initialDocument ?? null
  );
  // Monotonically increasing generation counter so a late `parseDocx`
  // result doesn't overwrite a newer load that started while we were
  // parsing.
  const loadGenerationRef = useRef(0);

  const loadParsedDocument = useCallback(
    (doc: Document) => {
      resetForNewDocument();
      setYrsSeedDocument(doc);
      history.reset(doc);
      setLoadingState({ isLoading: false, parseError: null });
      loadDocumentFonts(doc).catch((err) => {
        console.warn('Failed to load document fonts:', err);
      });
      // Offer the document's own renderable fonts (embedded faces are loaded by
      // parseDocx; system fonts are probed) in the picker.
      setDocumentFonts(
        getRenderableDocumentFonts(doc, {
          embeddedFamilies: getEmbeddedFontFamilies(doc.package?.fontTable),
        })
      );
    },
    [resetForNewDocument, history, setLoadingState, setDocumentFonts]
  );

  const loadBuffer = useCallback(
    async (buffer: DocxInput) => {
      const generation = ++loadGenerationRef.current;
      resetForNewDocument();
      setLoadingState({ isLoading: true, parseError: null });
      try {
        const doc = await parseDocx(buffer);
        if (loadGenerationRef.current !== generation) return;
        loadParsedDocument(doc);
      } catch (error) {
        if (loadGenerationRef.current !== generation) return;
        const message = error instanceof Error ? error.message : 'Failed to parse document';
        setLoadingState({ isLoading: false, parseError: message });
        onError?.(error instanceof Error ? error : new Error(message));
      }
    },
    [resetForNewDocument, loadParsedDocument, onError, setLoadingState]
  );

  // React to documentBuffer / document prop changes.
  useEffect(() => {
    // External-content mode: the caller (e.g. ySyncPlugin) populates PM
    // directly — skip the load.
    if (externalContent) return;

    if (!documentBuffer) {
      if (initialDocument) {
        loadParsedDocument(initialDocument);
      }
      return;
    }

    loadBuffer(documentBuffer);
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [documentBuffer, initialDocument, externalContent]);

  // Extract any baked-in comments from the document model on first load.
  // Bumps the shared comment/revision ID counter above all loaded IDs so new
  // comments and tracked changes don't collide with existing ones (they
  // share the OOXML ID space).
  useEffect(() => {
    if (commentsLoadedRef.current) return;
    const doc = history.state;
    if (!doc) return;
    commentsLoadedRef.current = true;
    const bodyComments = doc.package?.document?.comments;
    if (bodyComments && bodyComments.length > 0) {
      setComments(bodyComments);
      setShowCommentsSidebar(true);
    }
    // New Yrs revisions have replica-stable string IDs; the numeric OOXML
    // comment allocator only needs to stay above loaded comment/reply IDs.
    commentIdAllocator.seedAbove(
      (bodyComments ?? []).reduce((max, comment) => Math.max(max, comment.id), 0)
    );
  }, [
    history.state,
    pagedEditorRef,
    setComments,
    setShowCommentsSidebar,
    commentsLoadedRef,
    commentIdAllocator,
  ]);

  return {
    loadParsedDocument,
    loadBuffer,
    yrsSeedDocument,
  };
}
