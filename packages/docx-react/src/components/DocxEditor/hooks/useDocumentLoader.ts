import { useCallback, useEffect, useRef, useState } from 'react';
import type { Document } from '@betteroffice/docx/types/document';
import type { Comment } from '@betteroffice/docx/types/content';
import type { YrsDocxHost, YrsSession } from '@betteroffice/docx/yrs';
import { createDocx, repackDocx } from '@betteroffice/docx/docx';
import {
  extractEmbeddedFontFaces,
  extractFontsFromDocument,
  loadEmbeddedFontFamilies,
  registerDocumentFaces,
  getRenderableDocumentFonts,
  getEmbeddedFontFamilies,
  selectRenderableFonts,
  toArrayBuffer,
  type DocxInput,
  type FontLoadScope,
} from '@betteroffice/docx/utils';
import type { FontOption } from '@betteroffice/docx/utils/fontOptions';
import type { UseHistoryReturn } from '../../../hooks/useHistory';
import type { PagedEditorRef } from '../PagedEditor';
import type { CommentIdAllocator } from '../commentFactories';
import { DocumentLoadGeneration } from './documentLoadGeneration';
import { DocxWorkerError } from '../internals/docxWorkerError';

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
  workerViewer = false,
  workerOpen = workerViewer,
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
  fontScope,
}: {
  documentBuffer: DocxInput | null | undefined;
  initialDocument: Document | null | undefined;
  workerViewer?: boolean;
  workerOpen?: boolean;
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
  /** The editor instance's font loads, see `useFontLoadScope`. */
  fontScope: FontLoadScope;
}) {
  // The live history document changes after every edit, but yrs must only be
  // reseeded when a new source document is loaded. Keep that load boundary
  // separate so PagedEditor can replace its session without treating normal
  // edits as fresh documents.
  const [yrsSeedDocument, setYrsSeedDocument] = useState<Document | null>(
    workerOpen ? null : initialDocument ?? null
  );
  const [yrsSeedBytes, setYrsSeedBytes] = useState<Uint8Array | null>(null);
  const [yrsSeedGeneration, setYrsSeedGeneration] = useState(0);
  const [loadGeneration] = useState(() => new DocumentLoadGeneration());
  const previewDocumentRef = useRef<Document | null>(null);
  // Counts accepted host documents: a preview's font loads end once the full
  // document of its load is accepted.
  const hostDocumentsRef = useRef(0);
  // Embedded families registered under an alias because another live document
  // registered different faces under the same name.
  const [fontAliases, setFontAliases] = useState<ReadonlyMap<string, string>>(NO_FONT_ALIASES);
  const skippedFontsRef = useRef<SkippedFonts | null>(null);

  const failHostDocument = useCallback(
    (error: Error, generation: number, options?: { opened: boolean }) => {
      // A load that fails after its document was accepted has completed.
      if (
        options?.opened
          ? !loadGeneration.isCurrent(generation)
          : !loadGeneration.complete(generation)
      ) {
        return;
      }
      loadGeneration.fail(generation);
      // A preview's first pages, or a document that failed to show, are not
      // the document the load opened.
      if (options?.opened || previewDocumentRef.current) {
        previewDocumentRef.current = null;
        history.reset(null);
      }
      setYrsSeedDocument(null);
      setYrsSeedBytes(null);
      setLoadingState({ isLoading: false, parseError: error.message });
      onError?.(error);
    },
    [loadGeneration, history, onError, setLoadingState]
  );

  const loadParsedDocument = useCallback(
    (doc: Document, seedBytes?: Uint8Array) => {
      const generation = loadGeneration.begin();
      resetForNewDocument();
      if (workerOpen) {
        setYrsSeedDocument(null);
        setYrsSeedBytes(null);
        setYrsSeedGeneration(generation);
        history.reset(null);
        setLoadingState({ isLoading: true, parseError: null });
        setFontAliases(NO_FONT_ALIASES);
        void (async () => {
          try {
            const buffer = await (doc.originalBuffer ? repackDocx(doc) : createDocx(doc));
            if (!loadGeneration.isCurrent(generation)) return;
            setYrsSeedBytes(new Uint8Array(buffer));
          } catch (cause) {
            failHostDocument(new DocxWorkerError('open', cause), generation);
          }
        })();
        return;
      }
      setYrsSeedDocument(doc);
      setYrsSeedBytes(seedBytes?.slice() ?? null);
      setYrsSeedGeneration(generation);
      history.reset(doc);
      setLoadingState({ isLoading: false, parseError: null });
      setFontAliases(NO_FONT_ALIASES);
      // parseDocx registered the embedded faces for the page; this editor
      // holds them again, under the names its neighbours leave free.
      loadDocumentFontsInOrder(
        registerDocumentFaces(doc.package ? extractEmbeddedFontFaces(doc) : [], fontScope),
        () => loadGeneration.isCurrent(generation),
        setFontAliases,
        () => fontScope.loadDocumentFonts(doc)
      );
      // Offer the document's own renderable fonts (embedded faces are loaded by
      // parseDocx; system fonts are probed) in the picker.
      setDocumentFonts(
        getRenderableDocumentFonts(doc, {
          embeddedFamilies: getEmbeddedFontFamilies(doc.package?.fontTable),
        })
      );
    },
    [
      workerOpen,
      loadGeneration,
      resetForNewDocument,
      history,
      setLoadingState,
      setDocumentFonts,
      fontScope,
      failHostDocument,
    ]
  );

  const loadBuffer = useCallback(
    async (buffer: DocxInput) => {
      const generation = loadGeneration.begin();
      resetForNewDocument();
      setLoadingState({ isLoading: true, parseError: null });
      setFontAliases(NO_FONT_ALIASES);
      setYrsSeedDocument(null);
      setYrsSeedBytes(null);
      setYrsSeedGeneration(generation);
      try {
        const source = buffer instanceof ArrayBuffer ? buffer : await toArrayBuffer(buffer);
        if (!loadGeneration.isCurrent(generation)) return;
        setYrsSeedBytes(new Uint8Array(source));
        history.reset(null);
        await loadGeneration.waitForCompletion(generation);
      } catch (error) {
        if (!loadGeneration.complete(generation)) return;
        const message = error instanceof Error ? error.message : 'Failed to parse document';
        setLoadingState({ isLoading: false, parseError: message });
        onError?.(error instanceof Error ? error : new Error(message));
      }
    },
    [loadGeneration, resetForNewDocument, history, onError, setLoadingState]
  );

  const acceptHostDocument = useCallback(
    (
      host: YrsDocxHost,
      generation: number,
      session?: Pick<YrsSession, 'onUpdate'>,
      options?: { preview: boolean }
    ) => {
      // A preview shows the load's first pages; the full document completes it.
      if (
        options?.preview
          ? !loadGeneration.isCurrent(generation)
          : !loadGeneration.complete(generation)
      ) {
        // A session replaced within this load may hold another document.
        if (session && loadGeneration.isCurrent(generation)) skippedFontsRef.current?.changed();
        return;
      }
      const doc = host.document;
      const accepted = ++hostDocumentsRef.current;
      previewDocumentRef.current = options?.preview ? doc : null;
      history.reset(doc);
      setLoadingState({ isLoading: false, parseError: null });
      const embeddedFamilies = getEmbeddedFontFamilies(doc.package.fontTable);
      const documentFonts = [
        ...getRenderableDocumentFonts(doc, { embeddedFamilies }),
        ...selectRenderableFonts(host.referencedFonts, { embeddedFamilies }),
      ];
      setDocumentFonts(
        [...new Map(documentFonts.map((font) => [font.name.toLowerCase(), font])).values()]
      );
      // A preview's font loads stop once the full document is accepted.
      const isCurrent = () =>
        loadGeneration.isCurrent(generation) && hostDocumentsRef.current === accepted;
      // A preview never changes, so what its first pages skip stays skipped.
      const skipped = new Set(
        session || options?.preview ? host.unusedScriptFonts?.map(fontKey) : undefined
      );
      const isSkipped = (family: string) => skipped.has(fontKey(family));
      const skippedFonts =
        session && !options?.preview && skipped.size > 0
          ? skipUntilChanged(session, () => {
              if (!isCurrent()) return;
              fontScope
                .loadFontsWithMapping(
                  [...host.referencedFonts, ...extractFontsFromDocument(doc)].filter(isSkipped)
                )
                .catch((error) => console.warn('Failed to load document fonts:', error));
            })
          : null;
      skippedFontsRef.current = skippedFonts;
      loadDocumentFontsInOrder(
        loadEmbeddedFontFamilies(
          doc.package.fontTable,
          host.embeddedFonts,
          host.fontTableRelationshipsXml,
          fontScope
        ),
        isCurrent,
        setFontAliases,
        () => {
          const used = (family: string) => !isSkipped(family);
          const loaded = Promise.all([
            fontScope.loadFontsWithMapping(host.referencedFonts.filter(used)),
            fontScope.loadFontsWithMapping([...extractFontsFromDocument(doc)].filter(used)),
          ]);
          skippedFonts?.start();
          return loaded;
        }
      );
    },
    [loadGeneration, history, setDocumentFonts, setLoadingState, fontScope]
  );

  const isCurrentLoad = useCallback(
    (generation: number) => loadGeneration.isCurrent(generation),
    [loadGeneration]
  );

  const reportLayoutError = useCallback(
    (error: Error, onCurrentError?: (error: Error) => void) =>
      loadGeneration.reportError(yrsSeedGeneration, error, (current) => {
        onCurrentError?.(current);
        onError?.(current);
      }),
    [loadGeneration, yrsSeedGeneration, onError]
  );

  // React to documentBuffer / document prop changes.
  useEffect(() => {
    // External-content mode: the caller populates the document directly —
    // skip the load.
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
    const bodyComments = doc.package?.document?.comments;
    // A preview's parse generates other IDs than the full document's, whose
    // comments are the ones loaded; its sidebar opens now all the same.
    if (doc === previewDocumentRef.current) {
      if (bodyComments && bodyComments.length > 0) setShowCommentsSidebar(true);
      return;
    }
    commentsLoadedRef.current = true;
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

  useEffect(
    () => () => {
      loadGeneration.invalidate();
    },
    [loadGeneration]
  );

  return {
    loadParsedDocument,
    loadBuffer,
    yrsSeedDocument,
    yrsSeedBytes,
    yrsSeedGeneration,
    isCurrentLoad,
    acceptHostDocument,
    failHostDocument,
    reportLayoutError,
    fontAliases,
  };
}

const NO_FONT_ALIASES: ReadonlyMap<string, string> = new Map();

const fontKey = (family: string): string => family.trim().toLowerCase();

interface SkippedFonts {
  /** The fonts loaded at open have started loading. */
  start(): void;
  /** The document changed outside `session`'s updates. */
  changed(): void;
}

/**
 * Runs `load` once, after `start` and the first change to the document,
 * local or remote: an edit can give a font skipped at open text to draw.
 */
function skipUntilChanged(
  session: Pick<YrsSession, 'onUpdate'>,
  load: () => void
): SkippedFonts {
  let started = false;
  let changed = false;
  let loaded = false;
  const run = () => {
    if (!started || !changed || loaded) return;
    loaded = true;
    load();
  };
  const onChange = () => {
    unsubscribe();
    changed = true;
    run();
  };
  const unsubscribe = session.onUpdate(onChange);
  return {
    start: () => {
      started = true;
      run();
    },
    changed: onChange,
  };
}

/**
 * Takes the aliases of a document's embedded faces once they registered, and
 * only then loads the fonts it references: the previous document's faces are
 * released by then, so none of their loaded state stands in for this one's.
 */
function loadDocumentFontsInOrder(
  embedded: Promise<ReadonlyMap<string, string>>,
  isCurrent: () => boolean,
  setFontAliases: React.Dispatch<React.SetStateAction<ReadonlyMap<string, string>>>,
  loadReferenced: () => Promise<unknown>
): void {
  void embedded
    .then((families) => {
      if (!isCurrent()) return;
      const aliases = [...families].filter(([family, cssFamily]) => family !== cssFamily);
      setFontAliases((current) =>
        aliases.length === 0 && current.size === 0 ? current : new Map(aliases)
      );
    })
    .catch((error) => {
      console.warn('Failed to load embedded document fonts:', error);
    })
    .then(() => (isCurrent() ? loadReferenced() : undefined))
    .catch((error) => {
      console.warn('Failed to load document fonts:', error);
    });
}
