import { useImperativeHandle, useMemo } from 'react';
import type { Comment } from '@betteroffice/docx/types/content';
import type { Document } from '@betteroffice/docx/types/document';
import type {
  DocxEditStep,
  DocxExportResult,
  DocxLayoutMap,
  DocxPageExportOptions,
  DocxPagedStructuredContent,
  DocxProposalResult,
  DocxTextTarget,
  YrsInlineFormatDelta,
  YrsLoc,
  YrsParagraph,
  YrsSession,
  YrsStoryRange,
  WasmModuleMemory,
} from '@betteroffice/docx/yrs';
import { createStyleResolver } from '@betteroffice/docx/styles';
import type { DocxInput, ScrollToParaIdOptions } from '@betteroffice/docx/utils';
import type { DisplayList } from '@betteroffice/docx/layout/render';
import type { DocxEditorRef } from '../../DocxEditor';
import type { DocxCommandStore } from '../../../commands/types';
import type { PagedEditorRef } from '../PagedEditor';
import type { CommentIdAllocator } from '../commentFactories';
import { createComment } from '../commentFactories';
import { applyEditBatch, applyProposalCall, flushedSession, modeRefusal } from '../editorBatches';
import type { EditorMode } from '../internals/editing-modes';
import type { SelectionState } from '../types';
import { readMemoryStats } from '../memoryStats';
import { documentPageCount } from './documentPageCount';
import type { DocxHostSearch } from './useHostSearch';
import {
  awaitWorkerOpenReplica,
  ensureWorkerOpenReplica,
  requestOnDemandWorkerOpenReplica,
  requestWorkerOpenReplica,
  workerOpenReplicaOnDemand,
} from '../internals/workerOpenReplica';
import {
  handedOverRequest,
  workerProposalAuthority,
  type WorkerProposalAuthority,
} from '../internals/workerProposalAuthority';

export const DOCX_REF_REPLICA_ACCESS = {
  commands: 'commands',
  getDocument: 'sync',
  getEditorRef: 'sync',
  flushPendingInput: 'await',
  save: 'await',
  setZoom: 'independent',
  getZoom: 'independent',
  focus: 'sync',
  getCurrentPage: 'independent',
  getTotalPages: 'independent',
  getMemoryStats: 'independent',
  whenLayoutComplete: 'await',
  scrollToPage: 'independent',
  scrollToPosition: 'sync',
  openPrintPreview: 'sync',
  print: 'sync',
  loadDocument: 'independent',
  loadDocumentBuffer: 'independent',
  readParagraphs: 'await',
  getParagraphIdentities: 'await',
  resolveParagraphAnchors: 'await',
  listContentControls: 'await',
  findContentControls: 'await',
  findText: 'await',
  validateEdits: 'await',
  applyEdits: 'await',
  proposeChanges: 'await',
  setProposalStates: 'await',
  withdrawProposals: 'await',
  getProposals: 'await',
  exportStructuredWithPages: 'await',
  getPositionAtPoint: 'sync',
  addComment: 'sync',
  replyToComment: 'independent',
  resolveComment: 'independent',
  proposeChange: 'sync',
  applyFormatting: 'sync',
  setParagraphStyle: 'sync',
  insertBreak: 'sync',
  getPageContent: 'sync',
  scrollToParaId: 'sync',
  scrollToCommentId: 'sync',
  scrollToChangeId: 'sync',
  highlightRange: 'sync',
  findInDocument: 'sync',
  getSelectionInfo: 'sync',
  getComments: 'independent',
  search: 'await',
  searchNext: 'independent',
  searchPrevious: 'independent',
  searchGoTo: 'independent',
  clearSearch: 'independent',
  getSearchState: 'independent',
  onSearchChange: 'independent',
  onContentChange: 'independent',
  onSelectionChange: 'independent',
} as const satisfies Record<keyof DocxEditorRef, 'await' | 'sync' | 'independent' | 'commands'>;

/**
 * Synchronous APIs that an on-demand replica still loading answers without loading it at once:
 * `direct` needs no replica (the display list, print, focus, or a call that waits for the replica
 * itself), `unselected` has nothing selected before the replica, and `request` answers as unready
 * and asks for the replica.
 */
const ON_DEMAND_SYNC_ACCESS: Partial<Record<keyof DocxEditorRef, 'direct' | 'unselected' | 'request'>> = {
  focus: 'direct',
  scrollToPosition: 'direct',
  openPrintPreview: 'direct',
  print: 'direct',
  highlightRange: 'direct',
  getSelectionInfo: 'unselected',
  getPositionAtPoint: 'request',
};

/**
 * Thrown by a synchronous editor ref member that needs the document on the main thread while a
 * read-only `experimentalWorkerOpen` editor still holds it, with host proposals, in its worker. The
 * document starts loading; await `flushPendingInput()` (or the member's async counterpart) and call
 * it again.
 */
export class DocxReplicaNotReadyError extends Error {
  constructor(readonly member: string) {
    super(
      `${member} needs the document on the main thread, which is still loading; ` +
        'await flushPendingInput() and call it again'
    );
    this.name = 'DocxReplicaNotReadyError';
  }
}

function withDeadline(ready: Promise<void>, timeoutMs: number | undefined): Promise<void> {
  if (timeoutMs === undefined) return ready;
  return new Promise<void>((resolve, reject) => {
    const timer = setTimeout(
      () => reject(new Error('The document did not finish rendering')),
      timeoutMs
    );
    ready.then(
      () => {
        clearTimeout(timer);
        resolve();
      },
      (error: unknown) => {
        clearTimeout(timer);
        reject(error);
      }
    );
  });
}

const WORKER_PROPOSAL_ACCESS: ReadonlySet<keyof DocxEditorRef> = new Set([
  'proposeChanges', 'setProposalStates', 'withdrawProposals', 'getProposals',
  'readParagraphs', 'getParagraphIdentities', 'resolveParagraphAnchors',
]);

function gateReplicaAccess(
  api: DocxEditorRef,
  pagedEditorRef: React.RefObject<PagedEditorRef | null>,
  enabled: boolean
): DocxEditorRef {
  if (!enabled) return api;
  const gated = { ...api };
  for (const key of Object.keys(DOCX_REF_REPLICA_ACCESS) as Array<keyof DocxEditorRef>) {
    const access = DOCX_REF_REPLICA_ACCESS[key];
    const call = api[key];
    if ((access !== 'await' && access !== 'sync') || typeof call !== 'function') continue;
    Object.defineProperty(gated, key, {
      value: (...args: unknown[]) => {
        const session = pagedEditorRef.current?.getYrsSession();
        if (session) {
          if (WORKER_PROPOSAL_ACCESS.has(key) && workerProposalAuthority(session)) {
            return Reflect.apply(call, api, args);
          }
          if (access === 'sync') {
            const onDemand = workerOpenReplicaOnDemand(session) ? ON_DEMAND_SYNC_ACCESS[key] : undefined;
            if (onDemand === 'unselected') return null;
            if (onDemand === 'request') {
              requestOnDemandWorkerOpenReplica(session);
              return null;
            }
            // Proposals only the worker holds cannot be rebuilt here: the replica takes them over.
            if (onDemand === undefined && workerProposalAuthority(session)?.holdsWorkerState()) {
              void requestWorkerOpenReplica(session)?.catch(() => {});
              throw new DocxReplicaNotReadyError(key);
            }
            if (onDemand === undefined) ensureWorkerOpenReplica(session);
          } else {
            if (key === 'whenLayoutComplete' && workerOpenReplicaOnDemand(session)) {
              return Reflect.apply(call, api, args);
            }
            const ready = awaitWorkerOpenReplica(session);
            if (ready) {
              const timeoutMs =
                key === 'whenLayoutComplete'
                  ? (args[0] as { timeoutMs?: number } | undefined)?.timeoutMs
                  : undefined;
              const started = Date.now();
              return withDeadline(ready, timeoutMs).then(() => {
                if (pagedEditorRef.current?.getYrsSession() !== session) {
                  throw new Error('The document changed while opening the replica');
                }
                // The layout deadline covers the replica wait.
                const rest =
                  timeoutMs === undefined
                    ? args
                    : [
                        {
                          ...(args[0] as object),
                          timeoutMs: Math.max(0, timeoutMs - (Date.now() - started)),
                        },
                      ];
                return Reflect.apply(call, api, rest);
              });
            }
          }
        }
        return Reflect.apply(call, api, args);
      },
      enumerable: true,
    });
  }
  return gated;
}

const noWorkerMemory = (): null => null;

type LocatedParagraph = {
  story: string;
  paragraph: YrsParagraph;
};

function bodyStoryIds(session: YrsSession): string[] {
  return session.storyIds().filter((story) => story === 'body' || story.startsWith('body:'));
}

function locateParagraph(session: YrsSession, paraId: string): LocatedParagraph | null {
  for (const story of bodyStoryIds(session)) {
    const paragraph = session.paragraphs(story).find((candidate) => candidate.paraId === paraId);
    if (paragraph) return { story, paragraph };
  }
  return null;
}

/** The legacy helpers' `{ paraId, search? }` target, resolved in the accepted view by Rust. */
function helperTarget(story: string, paraId: string, search?: string): DocxTextTarget {
  return search === undefined
    ? { kind: 'paragraph', story, paraId }
    : { kind: 'search', text: search, within: { kind: 'paragraph', story, paraId }, view: 'accepted' };
}

function storyOffset(session: YrsSession, loc: YrsLoc): number {
  return session.locateParagraph(loc.story, loc.paraId).start + loc.offset;
}

/** How long a paged export waits for fonts and a layout of the flushed document. */
const LAYOUT_WAIT_MS = 2_000;
const LAYOUT_POLL_MS = 16;
const LAYOUT_REFUSALS: ReadonlySet<string> = new Set([
  'stale-document',
  'stale-layout',
  'layout-unavailable',
]);

/**
 * Flushes input, then exports with pages when the session's retained layout is of the current
 * version, lowered from its own stories, and computed from the inputs the editor would lay the
 * document out with now. Lays the document out once when it is not, then waits for fonts or a
 * pass the pipeline deferred. The session, document version and current inputs are checked again
 * after every wait and by the export itself, which runs synchronously right before returning;
 * waiting for a painted frame is not needed.
 */
async function exportWithPages(
  pagedEditorRef: React.RefObject<PagedEditorRef | null>,
  options: DocxPageExportOptions,
  experimentalWorkerOpen = false
): Promise<DocxExportResult<DocxPagedStructuredContent<DocxLayoutMap>>> {
  const { session } = await flushedSession(pagedEditorRef, experimentalWorkerOpen);
  const editor = (): PagedEditorRef => {
    const current = pagedEditorRef.current;
    if (!current || current.getYrsSession() !== session) {
      throw new Error('The document changed while it was being laid out');
    }
    return current;
  };
  const attempt = (): DocxExportResult<DocxPagedStructuredContent<DocxLayoutMap>> => {
    const request = editor().getLayoutRequest();
    if (request === null) {
      return {
        ok: false,
        version: session.version(),
        failure: {
          code: 'layout-unavailable',
          target: null,
          message: 'The fonts this document uses are not loaded yet.',
        },
      };
    }
    return session.exportStructuredWithPagesFor(options, request);
  };
  // A pass that changes only the revision preview lays out in the resident worker, so the
  // session's retained layout can still preview decisions the editor no longer shows.
  const showsMarkup = (): boolean => {
    const request = editor().getLayoutRequest();
    if (request === null) return false;
    const preview = (JSON.parse(request) as { renderEnv?: { revisionPreview?: object } })
      .renderEnv?.revisionPreview;
    return !preview || Object.keys(preview).length === 0;
  };
  let result = attempt();
  if (options.expectLayoutVersion !== undefined) return result;
  if (
    !result.ok &&
    (LAYOUT_REFUSALS.has(result.failure.code) ||
      (result.failure.code === 'unsupported-revision-layout' && showsMarkup()))
  ) {
    editor().relayout({ onHost: true });
    result = attempt();
  }
  const deadline = Date.now() + LAYOUT_WAIT_MS;
  while (!result.ok && LAYOUT_REFUSALS.has(result.failure.code) && Date.now() < deadline) {
    await new Promise((resolve) => setTimeout(resolve, LAYOUT_POLL_MS));
    result = attempt();
  }
  return result;
}

function normalizeSelection(session: YrsSession): YrsStoryRange | null {
  const selection = session.selection();
  if (!selection || selection.anchor.story !== selection.head.story) return null;
  const anchorOffset = storyOffset(session, selection.anchor);
  const headOffset = storyOffset(session, selection.head);
  const [start, end] =
    anchorOffset <= headOffset ? [selection.anchor, selection.head] : [selection.head, selection.anchor];
  return {
    story: start.story,
    start: { paraId: start.paraId, offset: start.offset },
    end: { paraId: end.paraId, offset: end.offset },
  };
}

function formattingDelta(marks: Parameters<DocxEditorRef['applyFormatting']>[0]['marks']) {
  const delta: YrsInlineFormatDelta = {};
  if (marks.bold !== undefined) delta.bold = marks.bold;
  if (marks.italic !== undefined) delta.italic = marks.italic;
  if (marks.underline !== undefined) {
    delta.underline = marks.underline
      ? typeof marks.underline === 'object'
        ? { style: marks.underline.style }
        : true
      : null;
  }
  if (marks.strike !== undefined) delta.strike = marks.strike;
  if (marks.color !== undefined) {
    delta.color = marks.color.rgb
      ? { rgb: marks.color.rgb }
      : marks.color.themeColor
        ? { themeColor: marks.color.themeColor }
        : null;
  }
  if (marks.highlight !== undefined) delta.highlight = marks.highlight || null;
  if (marks.fontSize !== undefined) delta.fontSize = marks.fontSize > 0 ? marks.fontSize : null;
  if (marks.fontFamily !== undefined) {
    const ascii = marks.fontFamily.ascii ?? marks.fontFamily.hAnsi;
    delta.fontFamily = ascii
      ? { ascii, hAnsi: marks.fontFamily.hAnsi ?? ascii }
      : null;
  }
  return delta;
}

/** Owns the public imperative surface without exposing an editor-view handle. */
export function useDocxEditorRefApi({
  ref,
  document,
  documentFromYrs,
  historyStateRef,
  pagedEditorRef: hostEditorRef,
  handleSave,
  zoom,
  setZoom,
  scrollPageInfo,
  readCurrentPage,
  loadParsedDocument,
  loadBuffer,
  comments,
  setComments,
  setShowCommentsSidebar,
  contentChangeSubscribersRef,
  onContentSubscribersChange,
  selectionChangeSubscribersRef,
  getCachedStyleResolver,
  commentIdAllocator,
  commands,
  modeRef,
  openingRef,
  allowHostProposalsRef,
  workerMemory = noWorkerMemory,
  settledDisplayList,
  awaitingDocument,
  experimentalWorkerOpen = false,
  hostSearch,
}: {
  ref: React.ForwardedRef<DocxEditorRef>;
  document: Document | null;
  documentFromYrs: () => Document | null;
  historyStateRef: React.RefObject<Document | null>;
  pagedEditorRef: React.RefObject<PagedEditorRef | null>;
  handleSave: () => Promise<ArrayBuffer | null>;
  zoom: number;
  setZoom: (zoom: number) => void;
  scrollPageInfo: { currentPage: number; totalPages: number; visible: boolean };
  /** The page the scroll position shows now, where it can be read. */
  readCurrentPage?: () => number | null;
  loadParsedDocument: (doc: Document) => void;
  loadBuffer: (buffer: DocxInput) => Promise<void>;
  comments: Comment[];
  setComments: React.Dispatch<React.SetStateAction<Comment[]>>;
  setShowCommentsSidebar: React.Dispatch<React.SetStateAction<boolean>>;
  contentChangeSubscribersRef: React.RefObject<Set<(doc: Document) => void>>;
  onContentSubscribersChange?: (count: number) => void;
  selectionChangeSubscribersRef: React.RefObject<Set<(state: SelectionState | null) => void>>;
  getCachedStyleResolver: (
    styles: Parameters<typeof createStyleResolver>[0]
  ) => ReturnType<typeof createStyleResolver>;
  commentIdAllocator: CommentIdAllocator;
  commands: DocxCommandStore;
  /** The editor's current write mode; `viewing` also stands for a read-only editor. */
  modeRef: React.RefObject<EditorMode>;
  /** While the document opens, the API has no editor and no document, as during a load. */
  openingRef?: React.RefObject<boolean>;
  /** Whether proposal methods run while the editor is read-only. */
  allowHostProposalsRef: React.RefObject<boolean>;
  /** The resident worker's wasm memories as of its latest reply. */
  workerMemory?: () => WasmModuleMemory[] | null;
  /** The renderer's display list once it shows the whole current document. */
  settledDisplayList?: (
    relayout: null,
    timeoutMs: number | null,
    scope?: 'document' | 'window'
  ) => Promise<DisplayList>;
  /** Whether a document load has not yet produced its first layout. */
  awaitingDocument?: () => boolean;
  experimentalWorkerOpen?: boolean;
  hostSearch: DocxHostSearch;
}) {
  const opening = (): boolean => openingRef?.current === true;
  const pagedEditorRef = useMemo<React.RefObject<PagedEditorRef | null>>(
    () => ({
      get current() {
        return openingRef?.current === true ? null : hostEditorRef.current;
      },
    }),
    [hostEditorRef, openingRef]
  );
  const hostProposalsAllowed = () =>
    modeRef.current !== 'viewing' || allowHostProposalsRef.current === true;
  const proposalAuthority = () => {
    const session = pagedEditorRef.current?.getYrsSession();
    return experimentalWorkerOpen && session ? workerProposalAuthority(session) : null;
  };
  /** A proposal call on the worker's registry while it holds them, else on the main session. */
  const routedProposalCall = <R extends { expectVersion: string }>(
    request: R,
    onWorker: (
      authority: WorkerProposalAuthority,
      main: (request: R) => Promise<DocxProposalResult>
    ) => Promise<DocxProposalResult>,
    call: (session: YrsSession, request: R) => DocxProposalResult
  ): Promise<DocxProposalResult> => {
    const main = (input: R) =>
      applyProposalCall(
        pagedEditorRef,
        hostProposalsAllowed,
        (session) => call(session, handedOverRequest(session, input)),
        experimentalWorkerOpen
      );
    const authority = proposalAuthority();
    if (!authority) return main(request);
    if (!hostProposalsAllowed()) {
      return Promise.resolve({
        ok: false,
        version: pagedEditorRef.current!.getYrsSession()!.version(),
        failure: { code: 'read-only', message: 'The editor is read-only' },
      });
    }
    return onWorker(authority, main);
  };
  useImperativeHandle(
    ref,
    () => gateReplicaAccess({
      commands,
      getDocument: () =>
        opening() ? null : (pagedEditorRef.current?.getDocument() ?? documentFromYrs() ?? document),
      getEditorRef: () => pagedEditorRef.current,
      flushPendingInput: async () => {
        await flushedSession(pagedEditorRef, experimentalWorkerOpen);
      },
      save: async () => (opening() ? null : handleSave()),
      setZoom,
      getZoom: () => zoom,
      focus: () => pagedEditorRef.current?.focus(),
      getCurrentPage: () => readCurrentPage?.() ?? scrollPageInfo.currentPage,
      // A preview's layouts are partial, so the count is the full document's even
      // before its pages replace the preview's, as `whenLayoutComplete` reports it.
      getTotalPages: () =>
        awaitingDocument?.() ? 0 : documentPageCount(hostEditorRef.current?.getLayout()),
      whenLayoutComplete: async (options) => {
        if (!settledDisplayList) throw new Error('This editor paints no display list');
        return (await settledDisplayList(
          null, options?.timeoutMs ?? null, experimentalWorkerOpen ? 'window' : 'document'
        )).pages.length;
      },
      getMemoryStats: () => readMemoryStats(workerMemory),
      scrollToPage: (pageNumber) => pagedEditorRef.current?.scrollToPage(pageNumber),
      scrollToPosition: (displayPosition) =>
        pagedEditorRef.current?.scrollToPosition(displayPosition),
      openPrintPreview: () => void commands.execute('print', null),
      print: () => void commands.execute('print', null),
      loadDocument: loadParsedDocument,
      loadDocumentBuffer: loadBuffer,

      readParagraphs: (request) => {
        const main = async (input: typeof request) =>
          (await flushedSession(pagedEditorRef, experimentalWorkerOpen)).session.readParagraphs(input);
        const authority = proposalAuthority();
        return authority ? authority.readParagraphs(request, main) : main(request);
      },
      getParagraphIdentities: () => {
        const main = async () =>
          (await flushedSession(pagedEditorRef, experimentalWorkerOpen)).session.paragraphIdentities();
        return proposalAuthority()?.paragraphIdentities(main) ?? main();
      },
      resolveParagraphAnchors: (anchors) => {
        const main = async (input: typeof anchors) => {
          const { session } = await flushedSession(pagedEditorRef, experimentalWorkerOpen);
          return { version: session.version(), results: input.map((anchor) => session.resolveParagraphAnchor(anchor)) };
        };
        const authority = proposalAuthority();
        return authority ? authority.resolveParagraphAnchors(anchors, main) : main(anchors);
      },
      listContentControls: async (options) =>
        (await flushedSession(pagedEditorRef, experimentalWorkerOpen)).session.listContentControls(options),
      findContentControls: async (query, options) =>
        (await flushedSession(pagedEditorRef, experimentalWorkerOpen)).session.findContentControls(query, options),
      findText: async (request) => (await flushedSession(pagedEditorRef, experimentalWorkerOpen)).session.findText(request),
      validateEdits: async (request) => {
        const { session } = await flushedSession(pagedEditorRef, experimentalWorkerOpen);
        return modeRefusal(session, modeRef.current, request) ?? session.validateEdits(request);
      },
      applyEdits: async (request) => {
        const outcome = await applyEditBatch(
          pagedEditorRef, () => modeRef.current, request, undefined, undefined, experimentalWorkerOpen
        );
        if ('flush' in outcome) throw outcome.flush.error;
        return outcome.result;
      },

      proposeChanges: (request) =>
        routedProposalCall(request, (authority, main) => authority.propose(request, main), (session, input) =>
          session.proposeChanges(input)
        ),
      setProposalStates: (request) =>
        routedProposalCall(request, (authority, main) => authority.setStates(request, main), (session, input) =>
          session.setProposalStates(input)
        ),
      withdrawProposals: (request) =>
        routedProposalCall(request, (authority, main) => authority.withdraw(request, main), (session, input) =>
          session.withdrawProposals(input)
        ),
      getProposals: () => {
        const main = async () =>
          (await flushedSession(pagedEditorRef, experimentalWorkerOpen)).session.getProposals();
        return proposalAuthority()?.getProposals(main) ?? main();
      },

      exportStructuredWithPages: (options) => exportWithPages(pagedEditorRef, options, experimentalWorkerOpen),
      getPositionAtPoint: (clientX, clientY) =>
        pagedEditorRef.current?.getPositionAtPoint(clientX, clientY) ?? null,

      addComment: (options) => {
        const editor = pagedEditorRef.current;
        const session = editor?.getYrsSession();
        const located = session ? locateParagraph(session, options.paraId) : null;
        if (!editor || !session || !located || options.search === '') return null;
        const comment = createComment(commentIdAllocator, options.text, options.author);
        const result = session.commentTextTarget(
          helperTarget(located.story, options.paraId, options.search),
          {
            id: String(comment.id),
            author: options.author,
            date: comment.date ?? '',
            body: comment.content,
          }
        );
        if (!result.ok) return null;
        editor.syncYrsInputState(true, [located.story]);
        setComments((previous) => [...previous, comment]);
        setShowCommentsSidebar(true);
        return comment.id;
      },

      replyToComment: (commentId, text, authorName) => {
        if (opening() || !comments.some((comment) => comment.id === commentId)) return null;
        const reply = createComment(commentIdAllocator, text, authorName, commentId);
        setComments((previous) => [...previous, reply]);
        return reply.id;
      },

      resolveComment: (commentId) => {
        if (opening()) return;
        setComments((previous) =>
          previous.map((comment) =>
            comment.id === commentId ? { ...comment, done: true } : comment
          )
        );
      },

      proposeChange: (options) => {
        const editor = pagedEditorRef.current;
        const session = editor?.getYrsSession();
        if (!editor || !session || (!options.search && !options.replaceWith)) return false;
        const located = locateParagraph(session, options.paraId);
        if (!located) return false;
        const suggest = { author: options.author, date: new Date().toISOString() };
        const step: DocxEditStep = options.search
          ? {
              op: 'replaceText',
              target: helperTarget(located.story, options.paraId, options.search),
              text: options.replaceWith,
              suggest,
            }
          : {
              op: 'insertText',
              target: helperTarget(located.story, options.paraId),
              at: 'end',
              text: options.replaceWith,
              suggest,
            };
        const result = session.applyEdits({
          expectVersion: session.version(),
          source: 'agent',
          steps: [step],
        });
        if (!result.ok) return false;
        if (result.applied) {
          editor.syncYrsInputState(true, result.changedStories, { inWorker: true });
        }
        setShowCommentsSidebar(true);
        return true;
      },

      applyFormatting: (options) => {
        const editor = pagedEditorRef.current;
        const session = editor?.getYrsSession();
        const located = session ? locateParagraph(session, options.paraId) : null;
        if (!editor || !session || !located) return false;
        if (options.search !== '') {
          const result = session.formatTextTarget(
            helperTarget(located.story, options.paraId, options.search),
            formattingDelta(options.marks)
          );
          if (!result.ok) return false;
        }
        editor.syncYrsInputState(true, [located.story]);
        return true;
      },

      setParagraphStyle: (options) => {
        const editor = pagedEditorRef.current;
        const session = editor?.getYrsSession();
        const located = session ? locateParagraph(session, options.paraId) : null;
        if (!editor || !session || !located) return false;
        const at = { paraId: options.paraId, offset: 0 };
        const range: YrsStoryRange = { story: located.story, start: at, end: at };
        const currentDocument = historyStateRef.current;
        const resolver = currentDocument?.package.styles
          ? getCachedStyleResolver(currentDocument.package.styles)
          : null;
        if (resolver && !resolver.hasParagraphStyle(options.styleId)) return false;
        session.applyParagraphStyle(range, options.styleId);
        editor.syncYrsInputState(true);
        return true;
      },

      insertBreak: (options) => {
        const editor = pagedEditorRef.current;
        const session = editor?.getYrsSession();
        const located = session ? locateParagraph(session, options.paraId) : null;
        if (!editor || !session || !located || located.story !== 'body') return false;
        const span = session.locateParagraph(located.story, located.paragraph.paraId);
        const at = {
          story: located.story,
          paraId: located.paragraph.paraId,
          offset: span.end - span.start,
        };
        if (options.type === 'page') session.insertPageBreak(at);
        else if (options.type === 'sectionNextPage') {
          session.insertSectionBreak(at, 'nextPage');
        } else if (options.type === 'sectionContinuous') {
          session.insertSectionBreak(at, 'continuous');
        } else return false;
        editor.syncYrsInputState(true);
        return true;
      },

      getPageContent: (pageNumber) => {
        const editor = pagedEditorRef.current;
        const session = editor?.getYrsSession();
        const layout = editor?.getLayout();
        const page = layout && !layout.partial ? layout.pages[pageNumber - 1] : undefined;
        if (!editor || !session || !page) return null;
        const seen = new Set<string>();
        const paragraphs: Array<{ paraId: string; text: string; styleId?: string }> = [];
        for (const fragment of page.fragments) {
          if (fragment.kind !== 'paragraph' || fragment.pmStart == null) continue;
          const loc =
            editor.displayPositionToYrsLoc(fragment.pmStart) ??
            editor.displayPositionToYrsLoc(fragment.pmStart + 1);
          if (!loc || seen.has(loc.paraId)) continue;
          const paragraph = session
            .paragraphs(loc.story)
            .find((candidate) => candidate.paraId === loc.paraId);
          if (!paragraph) continue;
          seen.add(paragraph.paraId);
          const styleId = paragraph.properties.pStyle;
          paragraphs.push({
            paraId: paragraph.paraId,
            text: paragraph.text,
            ...(typeof styleId === 'string' ? { styleId } : {}),
          });
        }
        return {
          pageNumber,
          text: paragraphs.map((paragraph) => `[${paragraph.paraId}] ${paragraph.text}`).join('\n'),
          paragraphs,
        };
      },

      scrollToParaId: (paraId: string, options?: ScrollToParaIdOptions) =>
        pagedEditorRef.current?.scrollToParaId(paraId, options) ?? false,
      scrollToCommentId: (commentId) =>
        pagedEditorRef.current?.scrollToCommentId(commentId) ?? false,
      scrollToChangeId: (revisionId) =>
        pagedEditorRef.current?.scrollToChangeId(revisionId) ?? false,
      highlightRange: (from, to) => pagedEditorRef.current?.highlightRange(from, to),

      findInDocument: (query, options) => {
        const session = pagedEditorRef.current?.getYrsSession();
        if (!session || !query) return [];
        const caseSensitive = options?.caseSensitive ?? false;
        const needle = caseSensitive ? query : query.toLowerCase();
        const limit = options?.limit ?? 20;
        const results: ReturnType<DocxEditorRef['findInDocument']> = [];
        for (const story of bodyStoryIds(session)) {
          for (const paragraph of session.paragraphs(story)) {
            if (results.length >= limit) return results;
            const haystack = caseSensitive ? paragraph.text : paragraph.text.toLowerCase();
            const offset = haystack.indexOf(needle);
            if (offset < 0 || haystack.indexOf(needle, offset + 1) >= 0) continue;
            results.push({
              paraId: paragraph.paraId,
              match: paragraph.text.slice(offset, offset + query.length),
              before: paragraph.text.slice(Math.max(0, offset - 40), offset),
              after: paragraph.text.slice(offset + query.length, offset + query.length + 40),
            });
          }
        }
        return results;
      },

      getSelectionInfo: () => {
        const session = pagedEditorRef.current?.getYrsSession();
        const range = session ? normalizeSelection(session) : null;
        if (!session || !range) return null;
        try {
          return session.selectionText(range);
        } catch {
          return null;
        }
      },

      getComments: () => (opening() ? [] : comments),

      onContentChange: (listener) => {
        const subscribers = contentChangeSubscribersRef.current;
        subscribers.add(listener);
        onContentSubscribersChange?.(subscribers.size);
        return () => {
          const removed = subscribers.delete(listener);
          onContentSubscribersChange?.(subscribers.size);
          return removed;
        };
      },
      onSelectionChange: (listener) => {
        selectionChangeSubscribersRef.current.add(listener);
        return () => selectionChangeSubscribersRef.current.delete(listener);
      },
      ...hostSearch,
    }, pagedEditorRef, experimentalWorkerOpen),
    [
      document,
      documentFromYrs,
      zoom,
      scrollPageInfo,
      readCurrentPage,
      handleSave,
      loadParsedDocument,
      loadBuffer,
      comments,
      commands,
      workerMemory,
      settledDisplayList,
      awaitingDocument,
      experimentalWorkerOpen,
      hostSearch,
    ]
  );
}
