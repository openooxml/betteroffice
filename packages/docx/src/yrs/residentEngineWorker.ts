/// <reference lib="webworker" />

import { resolveReservations, validReservationsPayload, type DocxResolvedReservations } from './reservations';
import { readResidentToolbarMetadata } from './residentToolbarMetadata';
import { readResidentRangeHighlights } from './residentRangeHighlights';
import { readResidentReviewSnapshot, readReviewTarget } from './residentReview';
import { createResidentCollaboration, readResidentCollaborationState, readResidentPresence } from './residentCollaboration';
import { stateVectorClock } from '../collaboration/protocol';
import { LAYOUT_META_VERSION, type LayoutMetaV1, type RetainedLayoutMeta } from './layoutMeta';
import { isCommandOutcomeUnknownError } from '../../../../shared/office-session/commands/types';
import { isTransferResult, type TransferResult } from '../../../../shared/office-session/protocol';
import { SessionFailure } from '../../../../shared/office-session/types';
import { buildCommandFrame, exactFrameBuffer, packResidentFrame } from './residentCommandFrame';
import { createResidentCommandHost } from './residentCommandHost';
import { createResidentDocxCommandHooks } from './residentCommandActions';
import { releaseResidentSelectionReads } from './residentCommandSelection';
import { currentSelectionRange, readCommandStickyPosition } from './residentCommandAnchors';
import { projectCommandEchoContext } from './residentCommandViewProjection';
import { EchoGlyphSources, type EchoGlyphEngine } from './residentEchoGlyphSource';
import { echoBackgroundPage, stageCommandPresentation } from './residentCommandPresentation';
import { registerCommandFonts } from './residentCommandFonts';
import { residentSourceMetadata, validateResidentSource, type ResidentSourceOpenRequest, type ResidentSourcePreviewOptions } from './residentSource';
import { decodeDocxHostJson } from './yrsSessionFacade';
import { buildResidentRegionLayoutRequest } from '../editor/computeLayout';
import { residentMediaImages } from './residentMediaImages';
import { ResidentSessionFonts, type ResidentFontEnvironment } from './residentFontEnvironment';
import type {
  DocxCommandApplyContext,
  DocxCommandEnvelope,
  DocxCommandFramePayload,
  DocxCommandPageIdentity,
  DocxCommandSelection,
  DocxFrameReservations,
  DocxNavigationIntent,
  DocxSelectionPointIntent,
  ResidentCommandHooksFactory,
  ResidentCommandHost,
  ResidentCommandInitRequest,
  ResidentCommandViewportRequest,
  PreparedDocxCommand,
  ResidentDocxCommandEngine,
} from './residentCommandProtocol';
import type { CollaborationCursor } from '../collaboration/types';
import type { YrsLoc, YrsRegionLayoutProgress, YrsRenderEnv, YrsResidentCaretRect, YrsResidentWorkerSnapshot } from './index';
import { createDisplayListQueries } from '../layout/render/displayListQueries';
import {
  createResidentEngineSession,
  type ResidentEngineSession,
} from './residentEngineSession';
import { finalPreviewDisplayWindow, finalPreviewPageCount } from './previewDisplayWindow';
import { preloadEditWasm, preloadEditWasmFrom } from './wasm/index';
import {
  createProposalRegistry,
  proposalRevisionPreview,
  type DocxProposalRegistry,
  type DocxProposalRequest,
  type DocxProposalResult,
} from './proposals';
import {
  computeAnchorTargetGeometry,
  computeProposalGeometryMirror,
  hiddenRanges,
  resolveNavigationTarget,
} from './proposalGeometry';
import { findBodyMatches } from './findMatches';
import { readResidentSearch, residentBodyPositions } from './residentSearch';
import type { ResidentSaveRecord } from './residentSave';
import { proposalProjectionStories } from './dirtyProjectionStories';
import { findParagraphs } from './findParagraphs';
import { DisplayPositionIndex } from './displayPositionIndex';
import { resolveYrsPointPosition } from './pointPosition';
import { readResidentPointer } from './residentPointerRead';
import { readCommentInsertionRange } from './residentCommentProjection';
import { partMetadata } from './residentParts';
import { readControlHandles } from './residentControls';
import { readResidentProperties } from './residentProperties';
import { readResidentFrameGeometry, readResidentCellSelectionRects, resetResidentObjectHandles } from './residentObjectGeometry';
import { readCellSelection } from './residentCommandTableSelection';
import {
  resolveBookmarkPosition,
  resolveBookmarkTarget,
  resolveCommentTarget,
  resolveParagraphTarget,
  resolveRevisionTarget,
  resolveSelectionInfo,
  resolveSelectionText,
  resolveSelectionUnit,
} from './viewerSelection';
import { readAdjacentRevision, readSidebar, readOutlineHeadings } from './sidebarReads';
import { clipboardTablePayload, readClipboardSlice } from './clipboardFragment';
import { storyOffsetForLoc, tableAnchors } from './storyPlainText';
import { createResidentScheduler, type SchedulerMessage } from './residentScheduler';
import { hasCachedYrsSidebarProjection } from '../layout/render/yrsSidebarProjection';
import {
  presentOffscreenPageBackBuffer,
  presentOffscreenPageBackBufferWithCaret,
  rasterizeDisplayPageToBackBuffer,
  releaseOffscreenPageCanvas,
} from '../layout/render/canvasBackend';
import {
  applyFrameDelta,
  applyFrameDeltaOwned,
  decodeFrameDelta,
  retainedFramePageById,
  type RetainedFrame,
} from '../layout/render/frameDelta';
import { GlyphCache } from '../layout/render/glyphCache';
import { wasmModuleMemories } from '../wasm/loadWasmAsset';
import {
  encodeDisplayListFrameExtras,
  loadRustDisplayListQueryEngine,
  loadedRustDisplayListQueryEngine,
  type DisplayListBuildInputs,
} from '../layout/render/rustDisplayList';
import {
  RESIDENT_HOST_MODULE_WAIT_MS,
  type ResidentEngineWorkerHostModule,
  type ResidentEngineWorkerRequest,
  type ResidentEngineWorkerResponse,
} from './residentEngineWorkerProtocol';
import {
  residentCaretDeviceRect,
  residentCaretSnapshotForFrame,
  type ResidentCaretPaintStyle,
} from './residentCaret';

const scope = self as unknown as DedicatedWorkerGlobalScope;
let resolveHostEditModule: (module: WebAssembly.Module | null) => void;
let hostEditModule = nextHostEditModule();
let session: ResidentEngineSession | null = null;
let collaboration: ReturnType<typeof createResidentCollaboration> | null = null;
let commandHost: ResidentCommandHost | null = null;
let releaseCommandHost: (() => void) | null = null;
let commandQueries: ((frame: RetainedFrame, version: string) => ReturnType<typeof createDisplayListQueries>) | null = null;
let commandInit: ResidentCommandInitRequest | null = null;
/** The bound command session's layout owner, through which every command layout runs. */
interface CommandLayout {
  /** Pages of the certified prefix the whole layout has not replaced yet, or null. */
  certifiedPages(): number | null;
  /** The document version the certified prefix lays out, or null. */
  certifiedVersion(): string | null;
  begin(): YrsRegionLayoutProgress;
  resume(blocks: number): YrsRegionLayoutProgress;
  /** Finishes a certified prefix's layout now; false when there is none. */
  complete(): boolean;
}
let commandLayout: CommandLayout | null = null;
let sessionFonts: ResidentSessionFonts | null = null;
let commandHooksFactory: ResidentCommandHooksFactory | null = null;
let commandFrameChanges: {
  pages: Set<number>; updates: Uint8Array[];
  selectionContext?: boolean;
  /** The document changed outside a command and its layout has not followed yet. */
  layoutStale?: boolean;
  viewport?: ResidentCommandViewportRequest;
  window?: [number, number];
  reviewFrame?: { frameEpoch: number; previewVersion: number };
  echoRaster?: { pageId: string; pixelRatio: number };
  layoutPreviewVersion?: number;
} | null = null;

export function setResidentCommandHooksFactory(factory: ResidentCommandHooksFactory): void {
  commandHooksFactory = factory;
}

let queryParseFailures = 0;
const queryParseFailureVersions = new WeakMap<RetainedFrame, string>();

function reportQueryParseFailure(frame: RetainedFrame, version: string, error: unknown): void {
  const key = `${frame.docEpoch}:${frame.layoutEpoch}:${frame.frameEpoch}:${layoutRevision}:${version}`;
  if (queryParseFailureVersions.get(frame) === key) return;
  queryParseFailureVersions.set(frame, key);
  console.error('[docx] display-list query parse failure', {
    error: error instanceof Error ? error.message : String(error), queryParseFailures: ++queryParseFailures,
  });
}

function bindResidentCommands(init: ResidentCommandInitRequest): ResidentCommandHost {
  if (!session) throw new Error('Resident engine worker is not initialized');
  if (!commandHooksFactory) throw new Error('Resident command hooks are not installed');
  releaseCommandHost?.();
  releaseCommandHost = null;
  const engine = session;
  if (init.rasterScale !== undefined) setCommandRasterScale(init.rasterScale);
  init.layoutInput = resolveLayoutReservationInput(engine, init.layoutInput, true);
  const input = JSON.parse(init.layoutInput || '{}');
  const recoveredReservations = init.recoveryState?.layoutReservations;
  if (recoveredReservations) {
    input.renderEnv = { ...input.renderEnv, reservations: recoveredReservations.values };
    init.layoutInput = JSON.stringify(input);
  }
  let layoutInput = init.layoutWithRegions ? init.layoutInput : JSON.stringify({
    ...input, measured: [], retainOnly: true,
    workerOwned: { bodyStory: 'body', renderEnv: input.renderEnv, measurement: input.measurement },
    workerHeadersFooters: JSON.parse(init.extras).headersFooters,
  });
  if (!init.layoutWithRegions) engine.layoutDocumentJson(layoutInput);
  let layoutDirty = !init.layoutWithRegions;
  let appliedReservations: DocxFrameReservations | undefined = recoveredReservations?.set;
  let laidOutReservations: DocxFrameReservations | undefined;
  sessionFonts?.dispose();
  const fonts = new ResidentSessionFonts(engine, scope.fonts);
  sessionFonts = fonts;
  const preparedFonts = new WeakMap<PreparedDocxCommand, Awaited<ReturnType<ResidentSessionFonts['prepare']>>>();
  const registry = proposals ??= createProposalRegistry(engine.proposalEngine);
  let layoutPreviewVersion = registry.snapshot().previewVersion;
  let prefix: { pages: number; documentWidth?: number; version: string } | null = null;
  let certify = false;
  const dropPrefix = () => {
    prefix = null;
    incompleteLayout = null;
    clearProvisionalFinalPages();
    selectionLayout = undefined;
  };
  const layoutWithPreview = () => {
    laidOutReservations = appliedReservations;
    const input = JSON.parse(init.layoutWithRegions ? init.layoutInput : layoutInput);
    const owner = init.layoutWithRegions ? input : input.workerOwned;
    const snapshot = registry.snapshot();
    const preview = proposalRevisionPreview(snapshot);
    if (preview !== undefined || owner.renderEnv?.revisionPreview !== undefined) {
      owner.renderEnv = { ...owner.renderEnv };
      if (preview) owner.renderEnv.revisionPreview = preview;
      else delete owner.renderEnv.revisionPreview;
    }
    if (init.layoutWithRegions) engine.layoutDocumentWithRegionsAfterEdit(JSON.stringify(input));
    else engine.layoutDocumentJson(JSON.stringify(input));
    layoutPreviewVersion = snapshot.previewVersion;
    if (commandFrameChanges) {
      commandFrameChanges.layoutPreviewVersion = layoutPreviewVersion;
      commandFrameChanges.layoutStale = false;
    }
  };
  const relayout = (step?: 'initial' | 'complete' | 'begin' | 'resume', blocks = ALL_BLOCKS) => {
    if (step === 'begin' || step === 'resume') {
      if (step === 'begin') laidOutReservations = appliedReservations;
      const progress = step === 'begin'
        ? engine.beginRegionLayout(init.layoutInput) : engine.resumeRegionLayout(blocks);
      if (progress.layoutJson !== undefined) {
        dropPrefix();
        certify = true;
        layoutDirty = true;
      }
      return progress;
    }
    if (step === 'initial') {
      laidOutReservations = appliedReservations;
      const laid = engine.layoutDocumentWithRegionsExactPrefixRetained(init.layoutInput,
        Math.max(init.provisionalPages ?? 0, init.displayWindow?.[1] ?? 0));
      if (laid.provisional) {
        prefix = { pages: laid.pages, documentWidth: laid.documentWidth, version: engine.proposalEngine.version() };
        provisionalFinalPages = laid.pages;
        incompleteLayout = { layoutInput: init.layoutInput, workerAuthoritative: true, extras: init.extras };
      }
      retainedLayoutVersion = prefix ? null : engine.proposalEngine.version();
      layoutReady = true;
      layoutDirty = true;
      return;
    }
    if (prefix) supersedeSlicedCompletion();
    if (step === 'complete') {
      laidOutReservations = appliedReservations;
      if (init.layoutWithRegions) engine.layoutDocumentWithRegionsRetained(init.layoutInput);
      else engine.layoutDocumentJson(layoutInput);
      certify = true;
    } else {
      layoutWithPreview();
      retainedLayoutVersion = engine.proposalEngine.version();
    }
    if (prefix) dropPrefix();
    layoutDirty = true;
  };
  const presentLayout = (): Pick<DocxCommandFramePayload, 'layoutProvisional' | 'finalizedPages' | 'documentWidth'> => {
    if (prefix) {
      return { layoutProvisional: true, finalizedPages: prefix.pages,
        ...(prefix.documentWidth === undefined ? {} : { documentWidth: prefix.documentWidth }) };
    }
    if (certify) {
      retainedLayoutVersion = engine.proposalEngine.version();
      certify = false;
    }
    return {};
  };
  let selectionLayout: { version: string; frame: RetainedFrame } | undefined;
  let selectionPage: DocxCommandPageIdentity | undefined;
  let selectionNavigation: {
    focus: YrsLoc; direction: 'up' | 'down'; queryFrame?: { frame: RetainedFrame; version: string };
  } | undefined;
  let selectionQueries: {
    frame: RetainedFrame; frameEpoch: number; layoutRevision: number; version: string;
    queries: ReturnType<typeof createDisplayListQueries>;
  } | undefined;
  const queriesFor = (frame: RetainedFrame, version: string) => {
    if (selectionQueries?.frame !== frame || selectionQueries.frameEpoch !== frame.frameEpoch ||
      selectionQueries.layoutRevision !== layoutRevision || selectionQueries.version !== version) {
      selectionQueries?.queries.dispose();
      selectionQueries = {
        frame, frameEpoch: frame.frameEpoch, layoutRevision, version,
        queries: createDisplayListQueries(frame.displayList, loadedRustDisplayListQueryEngine() ?? undefined),
      };
    }
    return selectionQueries.queries;
  };
  const readSelectionFrame = () => {
    const version = engine.proposalEngine.version();
    if (layoutDirty || !retainedFrame) {
      if (!selectionLayout || selectionLayout.version !== version) {
        const extras = JSON.stringify({
          ...JSON.parse(init.extras), headersFooters: JSON.parse(engine.retainedHeadersFootersJson() ?? 'null'),
        });
        selectionLayout = { version, frame: applyFrameDelta(null, decodeFrameDelta(engine.buildDisplayListFrame(extras, 0))) };
      }
      return selectionLayout;
    }
    return { frame: retainedFrame, version };
  };
  const selectionPageOf = (selected: DocxCommandSelection['focus'], limit = Infinity): number | undefined => {
    const at = selected.sticky
      ? engine.resolveStickyPosition({ story: selected.story, encoded: selected.sticky }) : selected;
    if (!at || (at.story !== 'body' && !at.story.startsWith('body:'))) return undefined;
    const position = displayPositionIndex(engine).positionOf(at, 'body');
    if (position === null) return undefined;
    const { frame: selectedFrame, version } = readSelectionFrame();
    let frame = selectedFrame;
    if (prefix && limit === prefix.pages) {
      const pages = frame.pages.filter(({ page, pageIndex }) => pageIndex < limit && page.unbuilt &&
        page.positionSpan && position >= page.positionSpan[0] && position <= page.positionSpan[1])
        .map(({ pageIndex }) => pageIndex);
      if (pages.length) {
        frame = applyFrameDelta(frame, decodeFrameDelta(engine.buildDisplayPagesFrame(pages, frame.frameEpoch)));
        if (layoutDirty || !retainedFrame) selectionLayout = { version, frame };
        else retainedFrame = frame;
        for (const page of pages) commandFrameChanges?.pages.add(page);
      }
    }
    const queries = queriesFor(frame, version);
    const forward = queries.rangeRects(position, position + 1).find((rect) =>
      rect.height > 0 && rect.pageIndex < limit && frame.pages[rect.pageIndex] && !frame.pages[rect.pageIndex]!.page.unbuilt);
    if (forward) return forward.pageIndex;
    const paragraph = engine.geometryReader.locateParagraph(at.story, at.paraId);
    if (at.offset !== paragraph.end - paragraph.start) return undefined;
    const caret = queries.caretRect(position);
    return caret && caret.height > 0 && caret.pageIndex < limit && frame.pages[caret.pageIndex] &&
      !frame.pages[caret.pageIndex]!.page.unbuilt ? caret.pageIndex : undefined;
  };
  resetResidentObjectHandles(engine);
  const baseHooks = commandHooksFactory(engine, init, {
    getLayoutReservations() {
      return (JSON.parse(init.layoutInput) as { renderEnv?: YrsRenderEnv }).renderEnv?.reservations ?? null;
    },
    setRenderEnvironment(environment, set) {
      if (set !== undefined) appliedReservations = set ?? (appliedReservations && { version: appliedReservations.version, key: null });
      const input = JSON.parse(init.layoutInput);
      input.renderEnv = { ...input.renderEnv, ...environment };
      init.layoutInput = JSON.stringify(input);
      if (!init.layoutWithRegions) {
        const layout = JSON.parse(layoutInput);
        layout.workerOwned.renderEnv = { ...layout.workerOwned.renderEnv, ...environment };
        layoutInput = JSON.stringify(layout);
      }
    },
    supportsFontEnvironment: true,
    proposals: registry,
    collaboration: collaboration !== null,
    readSelectionLayout() {
      const { version, frame: selectedFrame } = readSelectionFrame();
      let frame = selectedFrame;
      const target = selectionPage;
      const pages = new Set<number>();
      const installSelectionFrame = (bytes: Uint8Array) => {
        frame = applyFrameDelta(frame, decodeFrameDelta(bytes));
        if (layoutDirty || !retainedFrame) selectionLayout = { version, frame };
        else {
          retainedFrame = frame;
          if (commandFrameChanges) commandFrameChanges.reviewFrame = { frameEpoch: frame.frameEpoch, previewVersion: layoutPreviewVersion };
        }
      };
      const buildPages = () => {
        if (!pages.size) return;
        installSelectionFrame(engine.buildDisplayPagesFrame([...pages].sort((a, b) => a - b), frame.frameEpoch));
        for (const page of pages) commandFrameChanges?.pages.add(page);
        pages.clear();
      };
      if (target) {
        const page = retainedPageByKey(frame, target.pageId)
          ?? frame.pages[Math.max(0, Math.min(frame.pages.length - 1, Math.trunc(target.pageIndex)))];
        if (page?.page.unbuilt) pages.add(page.pageIndex);
      }
      if (selectionNavigation) {
        const { focus, direction } = selectionNavigation;
        const root = focus.story === 'body' || focus.story.startsWith('body:') ? 'body'
          : focus.story.split(':').slice(0, 2).join(':');
        const position = displayPositionIndex(engine).positionOf(focus, root);
        if (position !== null && root === 'body') {
          const queryFailed = () => {
            const state = queriesFor(frame, version).sourceState?.();
            if (state?.status !== 'error') return false;
            reportQueryParseFailure(frame, version, state.error);
            return true;
          };
          let atTextEnd: boolean | undefined;
          const trailingCaretOnPage = (pageIndex: number) => {
            if (atTextEnd === undefined) {
              const paragraph = engine.geometryReader.locateParagraph(focus.story, focus.paraId);
              atTextEnd = focus.offset === paragraph.end - paragraph.start;
            }
            const queries = queriesFor(frame, version);
            let rect = atTextEnd && position > 0
              ? queries.rangeRectsOnPages?.(position - 1, position, pageIndex, pageIndex).at(-1)
              : undefined;
            if (!rect && atTextEnd && queryFailed()) {
              rect = queries.visualLinesOnPage(pageIndex).filter((line) =>
                line.paraId === focus.paraId && line.from < position && line.to === position).at(-1);
            }
            return rect && rect.height > 0 ? { ...rect, x: rect.x + rect.width, width: 0 } : null;
          };
          let caret = queriesFor(frame, version).caretRect(position);
          let failed = queryFailed();
          if (!caret || caret.height <= 0 || frame.pages[caret.pageIndex]?.page.unbuilt) {
            for (const { page, pageIndex } of frame.pages) {
              if (page.unbuilt) continue;
              const trailing = trailingCaretOnPage(pageIndex);
              if (trailing) { caret = trailing; break; }
            }
          }
          if (!caret || caret.height <= 0 || frame.pages[caret.pageIndex]?.page.unbuilt) {
            const candidates = new Set<number>();
            if (caret && frame.pages[caret.pageIndex]?.page.unbuilt) candidates.add(caret.pageIndex);
            for (const { page, pageIndex } of frame.pages) {
              if (page.unbuilt && page.positionSpan &&
                position >= page.positionSpan[0] && position <= page.positionSpan[1]) candidates.add(pageIndex);
            }
            const probePage = candidates.values().next().value ?? -1;
            failed = queryFailed() || failed;
            let probes = 0;
            for (const pageIndex of candidates) {
              if (failed && probes >= 2) break;
              if (failed && probes > 0 && pageIndex !== probePage + 1) continue;
              probes += 1;
              installSelectionFrame(engine.buildDisplayPagesFrame([pageIndex], frame.frameEpoch));
              caret = queriesFor(frame, version).caretRect(position);
              if (!caret || caret.height <= 0 || frame.pages[caret.pageIndex]?.page.unbuilt) {
                caret = trailingCaretOnPage(pageIndex) ?? caret;
              }
              failed = queryFailed() || failed;
              if (caret && caret.height > 0 && !frame.pages[caret.pageIndex]?.page.unbuilt && caret.pageIndex === pageIndex) {
                commandFrameChanges?.pages.add(pageIndex);
                break;
              }
              const released = engine.releaseDisplayPagesFrame([pageIndex], frame.frameEpoch);
              if (released === null) throw new Error('Caret probe page could not be released');
              installSelectionFrame(released);
              if (caret && caret.height > 0 && !frame.pages[caret.pageIndex]?.page.unbuilt) break;
            }
          }
          if (caret && caret.height > 0 && !frame.pages[caret.pageIndex]?.page.unbuilt) {
            const neighbour = caret.pageIndex + (direction === 'up' ? -1 : 1);
            if (frame.pages[neighbour]?.page.unbuilt) pages.add(neighbour);
          }
        }
      }
      buildPages();
      if (selectionNavigation) selectionNavigation.queryFrame = { frame, version };
      return { frame, layoutRevision };
    },
  });
  const changes: NonNullable<typeof commandFrameChanges> = {
    pages: new Set<number>(), updates: [], layoutPreviewVersion, layoutStale: init.recoveryState !== undefined,
  };
  const wrapped = collaboration?.wrap(baseHooks) ?? baseHooks;
  const hooks = { ...wrapped, ...(wrapped.recoveryState ? {
    recoveryState: () => ({
      ...wrapped.recoveryState!(),
      comments: engine.commentRecoveryState?.(),
      layoutReservations: {
        values: (JSON.parse(init.layoutInput) as { renderEnv?: YrsRenderEnv }).renderEnv?.reservations ?? {},
        set: appliedReservations,
      },
    }),
  } : {}) };
  commandFrameChanges = changes;
  const prefixNavigation = {
    character: { backward: 'result', forward: 'result', up: 'complete', down: 'complete' },
    word: { backward: 'result', forward: 'result', up: 'complete', down: 'complete' },
    line: { backward: 'result', forward: 'result', up: 'interior', down: 'interior' },
    paragraph: { backward: 'interior', forward: 'interior', up: 'complete', down: 'complete' },
    'document-boundary': { backward: 'result', forward: 'complete', up: 'complete', down: 'complete' },
  } satisfies Record<DocxNavigationIntent['unit'], Record<DocxNavigationIntent['direction'], 'result' | 'interior' | 'complete'>>;
  const geometryIndependent = new Set<DocxCommandEnvelope['kind']>(['propose', 'setProposalStates', 'withdrawProposals']);
  const canUsePrefix = (
    command: DocxCommandEnvelope, context: DocxCommandApplyContext, result?: DocxCommandSelection | null
  ): boolean => {
    if (!prefix || changes.layoutStale) return false;
    if (geometryIndependent.has(command.kind)) return true;
    if (result !== undefined) {
      if (!result || result.anchor.story !== result.focus.story || result.anchor.paraId !== result.focus.paraId ||
        result.anchor.offset !== result.focus.offset || result.focus.story !== 'body') return false;
      const pageIndex = selectionPageOf(result.focus, prefix.pages);
      if (pageIndex === undefined) return false;
      const { frame, version } = readSelectionFrame();
      const position = displayPositionIndex(engine).positionOf(result.focus, 'body');
      if (position === null) return false;
      const queries = queriesFor(frame, version);
      const caret = queries.caretRect(position);
      if (!caret || caret.height <= 0 || caret.pageIndex !== pageIndex || queries.imageByPos(position)) return false;
      if (!frame.pages[pageIndex]!.page.primitives.some(primitive =>
        (primitive.kind === 'glyphRun' || primitive.kind === 'text') && primitive.docStart !== undefined &&
        primitive.docEnd !== undefined && primitive.docStart <= position && position <= primitive.docEnd)) return false;
      if (command.kind === 'selectPoint' || command.kind === 'extendSelection') {
        const point = command.payload as DocxSelectionPointIntent;
        const hit = queries.hitTestRegions(point.page.pageIndex, point.x, point.y);
        if (hit?.region !== 'body' || hit.target !== 'text') return false;
      }
    }
    const base = context.predecessorSelection ?? hooks.selection();
    switch (command.kind) {
      case 'selectPoint': case 'extendSelection': {
        const point = command.payload as DocxSelectionPointIntent | null;
        const index = point?.page?.pageIndex;
        if (index === undefined || !Number.isSafeInteger(index) || index < 0 || index >= prefix.pages ||
          typeof point?.page?.pageId !== 'string') return false;
        const page = retainedPageByKey(readSelectionFrame().frame, point.page.pageId);
        if (!page || page.pageIndex !== index) return false;
        return (command.kind !== 'extendSelection' && !point.modifiers?.shift) ||
          (!!base && selectionPageOf(base.anchor, prefix.pages) !== undefined);
      }
      case 'navigate': {
        const navigation = command.payload as DocxNavigationIntent | null;
        const policy = navigation && prefixNavigation[navigation.unit]?.[navigation.direction];
        if (!policy || policy === 'complete' || !base) return false;
        const focus = selectionPageOf(base.focus, prefix.pages);
        if (focus === undefined || selectionPageOf(base.anchor, prefix.pages) === undefined) return false;
        return policy === 'result' || focus < prefix.pages - 1;
      }
      default: return false;
    }
  };
  let revision = 0;
  const host = createResidentCommandHost({
    init,
    registry: { handlers: {}, policies: {} },
    hooks: {
      ...hooks,
      async preflight(command, context) {
        if (command.kind === 'selectPoint' || command.kind === 'extendSelection' || command.kind === 'navigate') {
          await loadRustDisplayListQueryEngine();
        }
        if (prefix && !(await host.withFrameWork(async () => canUsePrefix(command, context)))) await completeCommandPrefix(host);
        selectionPage = command.kind === 'selectPoint' || command.kind === 'extendSelection'
          ? (command.payload as { page?: DocxCommandPageIdentity } | null)?.page : undefined;
        const navigation = command.kind === 'navigate' ? command.payload as DocxNavigationIntent : undefined;
        if (navigation?.unit === 'line' && (navigation.direction === 'up' || navigation.direction === 'down')) {
          const selected = context.predecessorSelection?.focus;
          const focus = selected?.sticky
            ? engine.resolveStickyPosition({ story: selected.story, encoded: selected.sticky })
            : selected ?? engine.selection()?.head;
          if (focus) selectionNavigation = { focus, direction: navigation.direction };
        }
        const queryEngine = selectionNavigation ? loadedRustDisplayListQueryEngine() : null;
        const verticalMove = queryEngine?.verticalMoveJson;
        if (queryEngine && verticalMove) {
          queryEngine.verticalMoveJson = (...args) => {
            try {
              return verticalMove.apply(queryEngine, args);
            } catch (error) {
              const queryFrame = selectionNavigation?.queryFrame;
              if (queryFrame) reportQueryParseFailure(queryFrame.frame, queryFrame.version, error);
              throw error;
            }
          };
        }
        try {
          let checked = await hooks.preflight(command, context);
          if (prefix && (checked.status !== 'ready' || !canUsePrefix(command, context, checked.prepared.selection))) {
            await completeCommandPrefix(host);
            checked = await hooks.preflight(command, context);
          }
          if (checked.status === 'ready' && checked.prepared.selection &&
            (command.kind === 'selectPoint' || command.kind === 'extendSelection' || command.kind === 'navigate')) {
            changes.selectionContext = true;
          }
          if (command.kind === 'setFontEnvironment' && checked.status === 'ready') {
            try {
              preparedFonts.set(checked.prepared, await fonts.prepare(command.payload as ResidentFontEnvironment));
            } catch (error) {
              if (error instanceof WebAssembly.RuntimeError) throw error;
              return { status: 'refused' as const, reason: 'invalid-font-environment', selection: hooks.selection(), editState: hooks.editState() };
            }
          }
          return checked;
        } finally {
          if (queryEngine && verticalMove) queryEngine.verticalMoveJson = verticalMove;
          selectionPage = undefined;
          selectionNavigation = undefined;
        }
      },
      async apply(command, context) {
        if (command.command.kind === 'setFontEnvironment') {
          const prepared = preparedFonts.get(command);
          if (!prepared) throw new Error('Resident font environment was not prepared');
          const snapshot = fonts.commit(prepared, (faces, chains) => {
            engine.clearFonts();
            glyphCache = null;
            const registered = registerCommandFonts(engine, faces);
            if (registered.ids.some(id => id === null)) throw new Error('Resident font registration failed');
            const input = JSON.parse(init.layoutInput);
            input.measurement = { ...input.measurement, fontChains: chains };
            init.layoutInput = registered.remap(JSON.stringify(input));
            init.extras = registered.remap(JSON.stringify({ ...JSON.parse(init.extras), fontChains: chains }));
            if (!init.layoutWithRegions) {
              const layout = JSON.parse(layoutInput);
              layout.workerOwned.measurement = JSON.parse(init.layoutInput).measurement;
              layoutInput = JSON.stringify(layout);
            }
          });
          preparedFonts.delete(command);
          reply({ id: 0, ok: true, fontEnvironmentEvent: {
            ...snapshot, phase: 'ready', session: init.session, incarnation: init.incarnation, epoch: init.epoch,
          } });
          fontsRevision += 1;
          layoutRevision += 1;
          retainedLayoutVersion = null;
          for (const pageId of activeOffscreenPageIds) pendingOffscreenPageIds.add(pageId);
          intactBackBuffers.clear();
          fontRequirements = null;
          forgetRequirementsCache();
        }
        const result = await hooks.apply(command, context);
        if (result.status === 'applied' || (result.status === 'noop' && ['setRenderEnvironment', 'setFontEnvironment'].includes(command.command.kind))) {
          try { relayout(); } catch (error) {
            if (error instanceof WebAssembly.RuntimeError) throw error;
            throw projectionFailure = new SessionFailure('message', `Projection failed after ${command.command.kind}: ${error instanceof Error ? error.message : String(error)}`);
          }
        }
        return result;
      },
      async frame({ coversThroughSeq, full }) {
        const viewport = changes.viewport;
        const selectionContext = changes.selectionContext;
        changes.selectionContext = undefined;
        const previousWindow = changes.window ?? commandDisplayWindow ?? init.displayWindow;
        changes.viewport = undefined;
        const pages = [...changes.pages];
        changes.pages.clear();
        if (viewport) {
          changes.window = [viewport.start, viewport.end];
          commandDisplayWindow = changes.window;
          engine.setDisplayWindow(...changes.window);
          engine.setDisplayRetainBuiltPages(false);
        } else if (pages.length && !changes.window) {
          commandDisplayWindow = [Math.min(...pages), Math.max(...pages) + 1];
          engine.setDisplayWindow(...commandDisplayWindow);
        }
        const updates = changes.updates.splice(0);
        for (const update of updates) engine.applyUpdate(update);
        if (updates.length || changes.layoutStale) relayout();
        const framePreviewVersion = layoutPreviewVersion;
        const relaid = layoutDirty;
        layoutDirty = false;
        const extras = JSON.stringify({
          ...JSON.parse(init.extras), headersFooters: JSON.parse(engine.retainedHeadersFootersJson() ?? 'null'),
        });
        const preceding: Uint8Array[] = [];
        const residency = changes as typeof changes & {
          builtPages?: Map<string, number>; caretPage?: number; caretPageId?: string;
          caretFocus?: Pick<YrsLoc, 'story' | 'paraId'>;
        };
        const builtPages = changes.window ? residency.builtPages ??= new Map(retainedFrame?.pages
          .filter(({ page }) => !page.unbuilt).map(({ pageId, pageIndex }) => [pageId.toString(), pageIndex])) : undefined;
        const trackBuilt = (delta: ReturnType<typeof decodeFrameDelta>) => {
          if (!builtPages) return;
          if (delta.full) builtPages.clear();
          for (const operation of delta.operations) {
            if (operation.kind === 'remove' || (operation.kind === 'upsert' && operation.page.unbuilt)) {
              builtPages.delete(operation.pageId.toString());
            } else if (operation.kind === 'upsert' || (operation.kind === 'move' && builtPages.has(operation.pageId.toString()))) {
              builtPages.set(operation.pageId.toString(), operation.pageIndex);
            }
          }
        };
        for (const index of pages) {
          const page = retainedFrame?.pages[index];
          if (page && !page.page.unbuilt) builtPages?.set(page.pageId.toString(), page.pageIndex);
        }
        const requestedBuilds = new Set(pages);
        if (viewport && retainedFrame) {
          for (let index = viewport.start; index < Math.min(viewport.end, retainedFrame.pages.length); index += 1) {
            if (retainedFrame.pages[index]?.page.unbuilt) requestedBuilds.add(index);
          }
        }
        let bytes = (pages.length || viewport) && !relaid
          ? engine.buildDisplayPagesFrame([...requestedBuilds], full ? 0 : revision)
          : engine.buildDisplayListFrame(extras, full ? 0 : revision);
        let first: ReturnType<typeof decodeFrameDelta> | undefined;
        if (pages.length && relaid) {
          first = applyWorkerFrame(bytes, framePreviewVersion);
          trackBuilt(first);
          preceding.push(bytes);
          bytes = engine.buildDisplayPagesFrame(pages, first.frameEpoch);
        }
        let finalized = await finalizeResidentFrame(bytes, false, true, true, framePreviewVersion);
        trackBuilt(finalized.delta);
        if (!retainedFrame) throw new Error('Command frame is unavailable');
        let caretPage = finalized.caret.caretRect?.pageIndex;
        let modelCaret: YrsResidentCaretRect | undefined;
        let probed = false;
        if (changes.window) {
          const [start, end] = changes.window;
          const keep = new Set(pages);
          const selection = hooks.selection();
          const selected = selection?.focus;
          const focus = selected?.sticky
            ? engine.resolveStickyPosition({ story: selected.story, encoded: selected.sticky }) : selected;
          const nativeCaret = finalized.caret.caretRect;
          if (focus && (focus.story === 'body' || focus.story.startsWith('body:')) &&
            (!nativeCaret || nativeCaret.height <= 0 || retainedFrame.pages[nativeCaret.pageIndex]?.page.unbuilt)) {
            const paragraph = engine.geometryReader.locateParagraph(focus.story, focus.paraId);
            const length = paragraph.end - paragraph.start;
            const previousPage = residency.caretPage === undefined ? undefined : retainedFrame.pages[residency.caretPage];
            let position: number | null = null;
            if (residency.caretFocus?.story === focus.story && residency.caretFocus.paraId === focus.paraId &&
              previousPage && previousPage.pageId.toString() === residency.caretPageId && !previousPage.page.unbuilt) {
              for (const primitive of previousPage.page.primitives) {
                if (primitive.paraId !== focus.paraId || primitive.fragmentDocStart === undefined ||
                  primitive.fragmentDocEnd === undefined) continue;
                const inlineLength = primitive.fragmentDocEnd - primitive.fragmentDocStart - 2;
                const leading = length - inlineLength;
                if (inlineLength < 0 || leading < 0) continue;
                const candidate = primitive.fragmentDocStart + 1 + Math.min(Math.max(0, focus.offset - leading), inlineLength);
                if (position !== null && position !== candidate) { position = null; break; }
                position = candidate;
              }
            }
            const caretPosition = position ?? displayPositionIndex(engine).positionOf(focus, 'body');
            if (caretPosition !== null) {
              const version = engine.proposalEngine.version();
              const trailingCaretOnPage = (pageIndex: number) => {
                const rect = focus.offset === length && caretPosition > 0
                  ? queriesFor(retainedFrame!, version).rangeRectsOnPages?.(caretPosition - 1, caretPosition, pageIndex, pageIndex).at(-1)
                  : undefined;
                return rect && rect.height > 0 ? { ...rect, x: rect.x + rect.width, width: 0 } : null;
              };
              const install = async (bytes: Uint8Array) => {
                first ??= finalized.delta;
                preceding.push(finalized.bytes);
                finalized = await finalizeResidentFrame(bytes, false, true, true, framePreviewVersion);
                trackBuilt(finalized.delta);
                probed = true;
              };
              let caret = queriesFor(retainedFrame, version).caretRect(caretPosition);
              if (!caret || caret.height <= 0 || retainedFrame.pages[caret.pageIndex]?.page.unbuilt) {
                for (const [pageId, pageIndex] of builtPages!) {
                  const page = retainedFrame.pages[pageIndex];
                  if (page?.pageId.toString() !== pageId || page.page.unbuilt) continue;
                  const trailing = trailingCaretOnPage(pageIndex);
                  if (trailing) { caret = trailing; break; }
                }
              }
              if (!caret || caret.height <= 0 || retainedFrame.pages[caret.pageIndex]?.page.unbuilt) {
                const candidates = new Set<number>();
                for (const { page, pageIndex } of retainedFrame.pages) {
                  if (page.unbuilt && page.positionSpan && caretPosition >= page.positionSpan[0] &&
                    caretPosition <= page.positionSpan[1]) candidates.add(pageIndex);
                }
                if (caret && retainedFrame.pages[caret.pageIndex]?.page.unbuilt) candidates.add(caret.pageIndex);
                for (const index of candidates) {
                  await install(engine.buildDisplayPagesFrame([index], retainedFrame!.frameEpoch));
                  caret = queriesFor(retainedFrame!, version).caretRect(caretPosition);
                  if (!caret || caret.height <= 0 || retainedFrame!.pages[caret.pageIndex]?.page.unbuilt) {
                    caret = trailingCaretOnPage(index) ?? caret;
                  }
                  if (caret && caret.height > 0 && !retainedFrame!.pages[caret.pageIndex]?.page.unbuilt && caret.pageIndex === index) break;
                  const released = engine.releaseDisplayPagesFrame([index], retainedFrame!.frameEpoch);
                  if (released === null) throw new Error('Caret probe page could not be released');
                  await install(released);
                  if (caret && caret.height > 0 && !retainedFrame!.pages[caret.pageIndex]?.page.unbuilt) break;
                }
              }
              const page = caret && retainedFrame!.pages[caret.pageIndex];
              if (caret && caret.height > 0 && page && !page.page.unbuilt) {
                caretPage = caret.pageIndex;
                const selectedAnchor = selection?.anchor;
                const anchor = selectedAnchor?.sticky
                  ? engine.resolveStickyPosition({ story: selectedAnchor.story, encoded: selectedAnchor.sticky }) : selectedAnchor;
                if (anchor?.story === focus.story && anchor.paraId === focus.paraId && anchor.offset === focus.offset) {
                  modelCaret = { pageId: page.pageId.toString(), pageIndex: caret.pageIndex,
                    x: caret.x, y: caret.y, height: caret.height };
                }
              }
            }
          }
          if (caretPage !== undefined) keep.add(caretPage);
          for (let index = start; index < Math.min(end, retainedFrame.pages.length); index += 1) keep.add(index);
          const entering = [...keep].filter((index) => retainedFrame!.pages[index]?.page.unbuilt);
          if (entering.length) {
            first ??= finalized.delta;
            preceding.push(finalized.bytes);
            finalized = await finalizeResidentFrame(engine.buildDisplayPagesFrame(entering, retainedFrame!.frameEpoch), false, true, true, framePreviewVersion);
            trackBuilt(finalized.delta);
          }
          const distant = [...builtPages!].flatMap(([pageId, index]) => {
            const page = retainedFrame!.pages[index];
            return page?.pageId.toString() === pageId && !page.page.unbuilt && !keep.has(index) ? [index] : [];
          });
          if (distant.length) {
            const released = engine.releaseDisplayPagesFrame(distant, retainedFrame!.frameEpoch);
            if (released === null) throw new Error('Viewport display pages could not be released');
            first ??= finalized.delta;
            preceding.push(finalized.bytes);
            finalized = await finalizeResidentFrame(released, false, true, true, framePreviewVersion);
            trackBuilt(finalized.delta);
          }
          residency.caretPage = caretPage;
          residency.caretPageId = caretPage === undefined ? undefined : retainedFrame!.pages[caretPage]?.pageId.toString();
          residency.caretFocus = focus ? { story: focus.story, paraId: focus.paraId } : undefined;
          if (modelCaret) finalized.caret = { frameEpoch: retainedFrame!.frameEpoch, caretRect: modelCaret };
          if (probed || entering.length || distant.length) {
            selectionQueries?.queries.dispose();
            selectionQueries = undefined;
          }
        }
        const layout = { frame: retainedFrame!, layoutRevision };
        const scaleOnlyViewport = viewport?.rasterScale !== undefined && !pages.length && !updates.length && !relaid &&
          previousWindow?.[0] === viewport.start && previousWindow?.[1] === viewport.end;
        if (viewport?.rasterScale !== undefined) {
          const indices = new Set<number>();
          for (let index = viewport.start; index < Math.min(viewport.end, layout.frame.pages.length); index += 1) indices.add(index);
          if (caretPage !== undefined) indices.add(caretPage);
          setCommandRasterScale(viewport.rasterScale, indices);
          init.rasterScale = viewport.rasterScale;
        }
        const requestedPages = new Set(pages);
        if (viewport && !scaleOnlyViewport) {
          for (let index = viewport.start; index < Math.min(viewport.end, layout.frame.pages.length); index += 1) requestedPages.add(index);
        }
        if (changes.window && caretPage !== undefined && !scaleOnlyViewport) requestedPages.add(caretPage);
        if (!glyphCache) glyphCache = new GlyphCache({ provider: (fontId, glyphId) => engine.outlineGlyphJson(fontId, glyphId) });
        const pixelRatio = offscreenDpr * offscreenZoom;
        const echoRaster = changes.echoRaster;
        if (echoRaster && echoRaster.pixelRatio !== pixelRatio) pendingOffscreenPageIds.add(echoRaster.pageId);
        const echoGlyphs = commandEchoGlyphs(glyphCache, engine, full);
        const projectedEcho = hooks.editState().storedMarks ? undefined : projectCommandEchoContext({
          engine: engine as ResidentDocxCommandEngine, selection: hooks.selection(), layout,
          glyphs: echoGlyphs.provider(pixelRatio),
        });
        if (selectionContext && projectedEcho?.glyphs) requestedPages.add(projectedEcho.page.pageIndex);
        const echoContext = projectedEcho?.glyphs && (full || requestedPages.has(projectedEcho.page.pageIndex) ||
          pendingOffscreenPageIds.has(projectedEcho.page.pageId))
          ? projectedEcho : undefined;
        const echoGlyphSets = echoGlyphs.unsent(echoContext);
        const images = new Map<string, ImageBitmap>();
        const failedImages = new Set<string>();
        const rasterOptions = {
          glyphCache,
          async resolveImage(source: string) {
            if (failedImages.has(source)) return null;
            let image = images.get(source);
            if (!image) {
              try {
                const resolved = /^media:(0|[1-9]\d*)$/.test(source) ? engine.mediaDataUrl(source) : source;
                if (!resolved || (!resolved.startsWith('data:') && !resolved.startsWith('blob:'))) return null;
                image = await createImageBitmap(await (await fetch(resolved)).blob());
                images.set(source, image);
              } catch {
                failedImages.add(source);
                return null;
              }
            }
            return image;
          },
        };
        const visiblePages = new Set<number>();
        const center = caretPage ?? 0;
        if (caretPage !== undefined) visiblePages.add(center);
        if (echoContext) visiblePages.add(echoContext.page.pageIndex);
        if (echoRaster && pendingOffscreenPageIds.has(echoRaster.pageId)) {
          const retained = retainedPageByKey(layout.frame, echoRaster.pageId);
          if (retained) visiblePages.add(retained.pageIndex);
        }
        for (const index of pages) visiblePages.add(index);
        if (changes.window) {
          for (let index = changes.window[0]; index < Math.min(changes.window[1], layout.frame.pages.length); index += 1) visiblePages.add(index);
        } else if (activeOffscreenPageIds.size) {
          for (const pageId of activeOffscreenPageIds) {
            const retained = retainedPageByKey(layout.frame, pageId);
            if (retained) visiblePages.add(retained.pageIndex);
          }
        } else {
          const start = Math.max(0, center - 1);
          const window = commandDisplayWindow ?? [start, start + 3];
          for (let index = window[0]; index < window[1]; index += 1) visiblePages.add(index);
        }
        const stagedPages = [...visiblePages].flatMap((index) => {
          const retained = layout.frame.pages[index];
          if (!retained || retained.page.unbuilt) return [];
          const pageId = retained.pageId.toString();
          if (!full && !requestedPages.has(index) && !pendingOffscreenPageIds.has(pageId)) return [];
          return [{ pageId, pageIndex: retained.pageIndex }];
        });
        const presentation = await stageCommandPresentation({
          layout, pixelRatio, echoContext,
          pages: stagedPages,
          async rasterizePage(page) {
            const retained = retainedPageByKey(layout.frame, page.pageId);
            if (!retained) throw new Error('Command page is unavailable');
            const buffer = new OffscreenCanvas(1, 1);
            await rasterizeDisplayPageToBackBuffer(buffer, retained.page, rasterOptions, pixelRatio);
            return buffer.transferToImageBitmap();
          },
          async rasterizeBackground(dependency) {
            const retained = retainedPageByKey(layout.frame, dependency.page.pageId);
            if (!retained) throw new Error('Command background is unavailable');
            const { rect } = dependency;
            const device = [rect.x, rect.y, rect.width, rect.height].map((value) => Math.round(value * pixelRatio)) as
              [number, number, number, number];
            const buffer = new OffscreenCanvas(device[2], device[3]);
            await rasterizeDisplayPageToBackBuffer(buffer, echoBackgroundPage(retained, dependency), rasterOptions,
              pixelRatio, 1, { x: device[0], y: device[1], width: device[2], height: device[3] });
            return buffer.transferToImageBitmap();
          },
        }).finally(() => { for (const image of images.values()) image.close(); });
        if (echoContext) changes.echoRaster = { pageId: echoContext.page.pageId, pixelRatio };
        else if (stagedPages.some((page) => page.pageId === echoRaster?.pageId)) changes.echoRaster = undefined;
        for (const page of stagedPages) pendingOffscreenPageIds.delete(page.pageId);
        const shown: ReturnType<typeof presentLayout> = init.provisionalPages === undefined ? {} : presentLayout();
        const selected = hooks.selection();
        const cells = readCellSelection(engine);
        const cellRects = !shown.layoutProvisional ? readResidentCellSelectionRects(engine, layout.frame, cells) : null;
        const selectionRects = cellRects ?? (!shown.layoutProvisional && selected && (selected.anchor.story !== selected.focus.story ||
          selected.anchor.paraId !== selected.focus.paraId || selected.anchor.offset !== selected.focus.offset)
          ? (await readResidentRangeHighlights(engine, layout.frame, [{
            story: selected.anchor.story, start: selected.anchor, end: selected.focus,
          }], layoutRevision)).rects[0] ?? [] : []);
        const built = buildCommandFrame({
          frame: finalized.bytes,
          metadata: first
            ? { ...finalized.delta, full: first.full, baseFrameEpoch: first.baseFrameEpoch }
            : finalized.delta,
          pageFrames: preceding, coversThroughSeq, caret: finalized.caret,
          layoutRevision, documentVersion: engine.proposalEngine.version(),
          selectionRects,
          ...(fonts.snapshot ? { fontEnvironment: fonts.snapshot } : {}),
          parts: shown.layoutProvisional ? [] : engine.partGeometry(retainedFrame!),
          geometry: hooks.editState().readOnly || shown.layoutProvisional ? undefined
            : await readResidentFrameGeometry(engine, retainedFrame!, hooks.selection(), layoutRevision),
          ...shown, echoContext, ...(echoGlyphSets.length ? { echoGlyphSets } : {}), stagedPresentation: presentation.value,
          ...(laidOutReservations ? { reservations: laidOutReservations } : {}),
        });
        echoGlyphs.sent(echoGlyphSets);
        revision = built.value.revision;
        changes.reviewFrame = { frameEpoch: retainedFrame!.frameEpoch, previewVersion: framePreviewVersion };
        pendingUpdates = [];
        return built;
      },
      snapshot: () => engine.encodeState(),
    },
  });
  commandInit = init;
  commandQueries = queriesFor;
  releaseCommandHost = () => {
    commandQueries = null;
    selectionQueries?.queries.dispose();
    selectionQueries = undefined;
    releaseResidentSelectionReads(engine);
  };
  commandHost = host;
  if (!viewportFirst(init)) return host;
  commandLayout = {
    certifiedPages: () => prefix?.pages ?? null,
    certifiedVersion: () => prefix?.version ?? null,
    begin: () => relayout('begin')!,
    resume: (blocks) => relayout('resume', blocks)!,
    complete() {
      if (!prefix) return false;
      const waiting = slicedCompletion;
      slicedCompletion = null;
      if (waiting?.begun) {
        try {
          if (relayout('resume', ALL_BLOCKS)?.layoutJson !== undefined) return true;
        } catch (error) {
          if (!abandonedRegionLayout(error)) throw error;
        }
      }
      relayout('complete');
      return true;
    },
  };
  relayout('initial');
  return host;
}

type EchoGlyphState = { cache: GlyphCache; engine: EchoGlyphEngine; sources: EchoGlyphSources };

/** Echo glyph sets live as long as the glyph cache of the font store they were shaped from. */
function commandEchoGlyphs(cache: GlyphCache, engine: EchoGlyphEngine, full: boolean): EchoGlyphSources {
  let state = echoGlyphState;
  if (!state || state.cache !== cache || state.engine !== engine) {
    state = { cache, engine, sources: new EchoGlyphSources(engine, ++echoGlyphGeneration) };
    echoGlyphState = state;
  }
  if (full) state.sources.resend();
  return state.sources;
}

/** Whether `init` opens on a certified prefix that completes in the background. */
function viewportFirst(init: ResidentCommandInitRequest): boolean {
  return init.layoutWithRegions && init.provisionalPages !== undefined;
}

/**
 * Finishes a certified prefix's layout after the frames already in flight; false when there is none.
 * Only for requests `runRequest` does not run inside the frame work, which complete directly.
 */
function completeCommandPrefix(host: ResidentCommandHost): Promise<boolean> {
  if (commandLayout?.certifiedPages() == null) return Promise.resolve(false);
  return host.withFrameWork(async () => commandLayout?.complete() ?? false);
}

/** A sliced region layout that a change in between abandoned. */
function abandonedRegionLayout(error: unknown): boolean {
  if (error instanceof WebAssembly.RuntimeError) return false;
  const message = error instanceof Error ? error.message : String(error);
  return message === 'no region layout to resume' ||
    message === 'the document or its fonts changed since the region layout began';
}

function replacementCommandInit(): ResidentCommandInitRequest | null {
  if (!commandInit) return null;
  return {
    ...commandInit, epoch: commandInit.epoch + 1,
    version: 0, throughSeq: 0, refusedSeqs: [], replayAppliedSeqs: [], selection: null,
  };
}

function requireCommandHost(): ResidentCommandHost {
  if (!commandHost) throw new Error('Resident commands are not initialized');
  return commandHost;
}

function replyCommand<K extends 'commandInitial' | 'commandAck' | 'checkpoint', T>(
  id: number,
  key: K,
  result: T | TransferResult<T>
): void {
  const value = isTransferResult(result) ? result.value : result;
  reply({ id, ok: true, [key]: value }, isTransferResult(result) ? result.transfer : []);
}

/** Answers `id` with the next command frame, carrying `answer` (and its `transfer`) along. */
function postCommandFrame(
  id: number,
  answer: Omit<ResidentEngineWorkerResponse & { ok: true }, 'id' | 'ok'> = {},
  transfer: Transferable[] = []
): void {
  const host = requireCommandHost();
  const init = commandInit!;
  void host.frame().then((result) => {
    if (commandHost !== host) {
      replyFailure(id, new SessionFailure('disposed', 'Resident command document was replaced'));
      return;
    }
    const frame = isTransferResult(result) ? result.value : result;
    reply({
      ...answer, id, ok: true, ...frame.payload, selection: session?.selection() ?? null,
      ...(frame.payload.pageFrames ? { pageFrames: [...frame.payload.pageFrames, frame.payload.frame] } : {}),
      commandFrame: { session: init.session, incarnation: init.incarnation, epoch: init.epoch, frame },
    }, [...transfer, ...(isTransferResult(result) ? result.transfer : [])]);
  }).catch((error: unknown) => {
    if (commandHost !== host) {
      replyFailure(id, new SessionFailure('disposed', 'Resident command document was replaced'));
      return;
    }
    reply({ id, ok: false, terminal: true, error: error instanceof Error ? error.message : String(error),
      failureCode: error instanceof SessionFailure ? error.code : 'message' });
  });
}
let positionIndex: { session: ResidentEngineSession; index: DisplayPositionIndex } | null = null;

function displayPositionIndex(current: ResidentEngineSession): DisplayPositionIndex {
  if (positionIndex?.session !== current) {
    positionIndex = {
      session: current,
      index: new DisplayPositionIndex({
        ...current.geometryReader,
        selectionText: current.selectionText,
        resolveComment: current.resolveComment,
      }),
    };
  }
  return positionIndex.index;
}
let proposals: DocxProposalRegistry | null = null;
let lastProposalMirrorVersion: string | null = null;
let fontRequirements: {
  version: string;
  layoutInput: string;
  requirementsJson: string;
} | null = null;
/** Font requirements by layout input, for one session at one document version. */
let requirementsCache: {
  owner: ResidentEngineSession;
  version: string;
  byInput: Map<string, string>;
  release: () => void;
} | null = null;
/** The layout input of the host's last font requirements request. */
let requestedRequirements: { owner: ResidentEngineSession; layoutInput: string } | null = null;
const REQUIREMENTS_CACHE_INPUTS = 8;
/** Set while the session holds the document `open` seeded, with the heap limit it used. */
let openedDocument: { heapLimitBytes?: number } | null = null;
let openedSource: { bytes: ArrayBuffer; hostJson: string } | null = null;
let openedViaSource = false;
let openedParsed = false;
let sourcePreviewing = false;
let editorSaves: ResidentSaveRecord = { full: false };
class SaveUnavailableError extends Error {}
class SourceCutKeptError extends Error {}
let openedVersion: string | null | undefined;
/** A change before the whole document's first frame leaves no frame as opened. */
function noteDocumentChange(): void {
  if (openedVersion === undefined) openedVersion = null;
}
// The opened document is a display-only preview that an `open` of the whole package replaces.
let previewing = false;
/** Whether the opened source is a cut that stops before the end of its body. */
let cutPartial = false;
/** The request that opened the current cut, re-opened when the whole source that replaces it fails. */
let openedCut: ResidentSourceOpenRequest | null = null;
/** Pages of a cut preview's layout that match the whole document's; null for a whole document. */
let previewFinalPages: number | null = null;
let provisionalFinalPages: number | null = null;
let provisionalDisplayWindow: {
  window: [number, number];
  retainBuiltPages: boolean;
} | null = null;
let unsubscribe: (() => void) | null = null;
let layoutReady = false;
let pendingUpdates: Uint8Array[] = [];
let layoutRevision = 0;
let retainedRegions = false;
let retainedLayoutVersion: string | null = null;
let headersFootersPayload: string | undefined;
let headersFootersEpoch = 0;
let sentHeadersFootersEpoch = 0;
// -1 = no fonts applied yet (fresh session); hydrate skips re-registration
// when the snapshot's revision matches what this session already holds.
let fontsRevision = -1;
let retainedFrame: RetainedFrame | null = null;
let glyphCache: GlyphCache | null = null;
let echoGlyphState: EchoGlyphState | null = null;
let echoGlyphGeneration = 0;
const offscreenCanvases = new Map<string, OffscreenCanvas>();
const offscreenBackBuffers = new Map<string, OffscreenCanvas>();
const pendingOffscreenPageIds = new Set<string>();
let activeOffscreenPageIds = new Set<string>();
let commandDisplayWindow: [number, number] | undefined;
let offscreenDpr = 1;
let offscreenZoom = 1;
// Present-time caret painting. `caretPaintRect` is what the current frame
// wants painted, `paintedCaretPageId/Key` what is on screen. Caret-composited
// pages keep their back-buffer raster (`intactBackBuffers`) so the line can be
// dropped by re-presenting without raster or engine work; plain presents
// detach their bitmap and stay zero-copy.
let caretStyle: ResidentCaretPaintStyle = { color: '#000', width: 2 };
let caretPaintRect: YrsResidentCaretRect | null = null;
let paintedCaretPageId: string | null = null;
let paintedCaretKey: string | null = null;
let caretStage: OffscreenCanvas | null = null;
const intactBackBuffers = new Set<string>();
let trap: WebAssembly.RuntimeError | null = null;
let projectionFailure: SessionFailure | null = null;
// A bootstrap laid out only the body's first pages: the request to finish
// it, then the finished layout until `completeLayout` hands it over.
interface LayoutRequest {
  extras: string;
  layoutExtras?: string;
}
let incompleteLayout:
  | (LayoutRequest & { layoutInput: string; workerAuthoritative?: boolean })
  | null = null;
let completedLayout:
  | (LayoutRequest & { layoutJson: string; headersFootersJson: string | undefined })
  | null = null;
// A `completeLayout` measured a few blocks at a time, as operations queued
// behind the requests that arrive meanwhile.
interface SlicedCompletion {
  id: number;
  expectedFrameEpoch: number;
  paintCaret: boolean;
  /** Body blocks per step. */
  blocks: number;
  begun: boolean;
  /** Times a change in between abandoned the pass. */
  restarts: number;
  /** The command session whose certified prefix this completes; its frames publish the result. */
  commandOwner?: ResidentCommandHost;
}
let slicedCompletion: SlicedCompletion | null = null;
const COMPLETION_RESTARTS = 3;
const COMMAND_COMPLETION_BLOCKS = 64;
const COMPLETION_IDLE_MS = 300;
const ALL_BLOCKS = 2 ** 32 - 1;

interface BackgroundPageBuild {
  request: Extract<ResidentEngineWorkerRequest, { type: 'buildPages' }>;
  owner: ResidentEngineSession;
  frameEpoch: number;
  frames: Uint8Array[];
  offset: number;
  started: number;
  engineMs: number;
  pagesPerSlice: number;
}
let backgroundPageBuild: BackgroundPageBuild | null = null;
const BACKGROUND_SLICE_PAGES = 4;

// The request being handled, and the requests answered with a trap.
let handlingId = 0;
const trappedIds = new Set<number>();

scope.onmessage = (
  event: MessageEvent<ResidentEngineWorkerRequest | ResidentEngineWorkerHostModule>
) => {
  const message = event.data;
  if (message.type === 'editModule') {
    // The queued warm waits for this message.
    resolveHostEditModule(message.module instanceof WebAssembly.Module ? message.module : null);
    return;
  }
  scheduler.submit(classify(message));
};
scope.onmessageerror = () => resolveHostEditModule(null);

function nextHostEditModule(): Promise<WebAssembly.Module | null> {
  return new Promise((resolve) => {
    resolveHostEditModule = resolve;
  });
}

function classify(request: ResidentEngineWorkerRequest): SchedulerMessage {
  const run = () => runRequest(request);
  switch (request.type) {
    case 'edit': {
      const admittedHost = commandHost;
      const policy = admittedHost?.policy(request.command) ?? {
        lane: 'input' as const, reorderable: false, mutates: true, reframes: true,
      };
      return {
        ...policy,
        ...(isUserInputCommand(request.command.kind) ? { userInput: true, holdsIdleTasks: true } : {}),
        run: () => {
          if (commandHost && commandHost !== admittedHost) commandHost.policy(request.command);
          return runRequest(request);
        },
      };
    }
    case 'checkpoint':
      return { ...(commandHost?.checkpointPolicy() ?? { lane: 'input', reorderable: false }), run };
    case 'commandInit':
      return { lane: 'input', reorderable: false, reframes: true, run };
    case 'requestFullFrame':
      return { lane: 'input', reorderable: false, reframes: true, run };
    case 'applyInput':
    case 'applyDelete':
      return { lane: 'input', userInput: true, holdsIdleTasks: true, mutates: true, run };
    case 'applyUpdate':
    case 'syncUpdate':
    case 'commandUpdate':
      return { lane: 'collab', userInput: true, mutates: true, run };
    case 'proposal':
      if (request.operation.kind === 'snapshot') return { lane: 'interactive', run };
      return {
        lane: 'input', mutates: true,
        ...(request.operation.kind === 'setStates' ? { userInput: true, holdsIdleTasks: true } : {}), run,
      };
    case 'open':
    case 'openSource':
    case 'previewSource':
    case 'bootstrap':
    case 'sync':
    case 'destroy':
      return { lane: 'input', mutates: true, run };
    case 'warm':
      return { lane: 'input', run };
    case 'documentRead':
      return {
        lane: 'interactive',
        ...(request.expectVersion !== undefined ? { reorderable: true } : {}), run,
      };
    case 'fontRequirements':
    case 'encodeState':
    case 'revisionCount':
    case 'layoutJson':
    case 'save':
      return { lane: 'interactive', run };
    case 'buildPages':
      return {
        lane: 'interactive', reframes: true, key: 'pages', run,
        ...(request.background ? {
          replaceableBy: 'pages',
          supersede: () => replyDropped({ id: request.id, ok: true, pageBuildSuperseded: true }),
        } : {}),
      };
    case 'buildFrame':
    case 'commandViewport':
    case 'releasePages':
    case 'completeLayout':
    case 'attachCanvases':
    case 'eraseCaret':
      return { lane: 'interactive', reframes: true, run };
  }
}

function isUserInputCommand(kind: string): boolean {
  return kind === 'resizeObject' || kind === 'moveObject' || kind === 'setSectionMargins' || kind === 'setTabStops' || kind === 'removeTabStop' || kind === 'setHyperlink' || kind === 'setObjectProperties' || kind === 'setSectionProperties' || kind === 'setNoteProperties' || kind === 'setWatermark' || kind === 'setContentControlValue' || kind === 'addRepeatingItem' || kind === 'removeRepeatingItem' || kind === 'openPart' || kind === 'ensurePart' || kind === 'removePart' || kind === 'replaceMatches' || kind === 'undo' || kind === 'redo' || kind === 'select' || kind === 'insertText' || kind === 'replaceSelection' ||
    kind === 'deleteBackward' || kind === 'deleteForward' || kind === 'paragraphBreak' || kind === 'lineBreak' ||
    kind === 'toggleMark' || kind === 'formatRange' || kind === 'clearFormatting' ||
    kind === 'setParagraphAttrs' || kind === 'applyParagraphStyle' || kind === 'formatSelection' || kind === 'toggleStrike' ||
    kind === 'toggleList' || kind === 'changeIndent' || kind === 'insertImage' || kind === 'insertTable' ||
    kind === 'pasteFragment' ||
    kind === 'insertPageBreak' || kind === 'insertSectionBreak' ||
    kind === 'setTableBorders' || kind === 'setTableProperties';
}

function isLegacyMutation(request: ResidentEngineWorkerRequest): boolean {
  return request.type === 'applyInput' || request.type === 'applyDelete' ||
    request.type === 'applyUpdate' || request.type === 'syncUpdate' || request.type === 'sync' ||
    (request.type === 'bootstrap' && request.opened === true) ||
    (request.type === 'proposal' && request.operation.kind !== 'snapshot');
}

async function runRequest(request: ResidentEngineWorkerRequest): Promise<void> {
  try {
    if (trap || projectionFailure) throw trap ?? projectionFailure;
    handlingId = request.id;
    if ((commandHost || openedViaSource) && isLegacyMutation(request)) {
      throw new Error('Legacy mutations are unavailable in a resident command session');
    }
    if (commandHost && request.type !== 'edit' && request.type !== 'checkpoint' &&
      request.type !== 'commandInit' && request.type !== 'requestFullFrame' &&
      request.type !== 'commandUpdate' && request.type !== 'buildPages' && request.type !== 'commandViewport') {
      await commandHost.withFrameWork(() => handle(request));
    } else await handle(request);
  } catch (error) {
    replyFailure(request.id, error);
  }
}

/** Answers a request that will not run; after a trap, as the trap's failure. */
function replyDropped(response: ResidentEngineWorkerResponse & { ok: true }): void {
  if (trap) replyFailure(response.id, trap);
  else reply(response);
}

function replyFailure(id: number, error: unknown): void {
  // Once trapped, every failure is the trap's, answered once per request.
  const failure = trap ?? error;
  if (!trap && projectionFailure) {
    reply({ id, ok: false, terminal: true, error: projectionFailure.message, failureCode: projectionFailure.code });
    return;
  }
  if (isCommandOutcomeUnknownError(failure) || failure instanceof SessionFailure) {
    const message = collaboration && isCommandOutcomeUnknownError(failure)
      ? `Command ${failure.refusal.seq} has an unknown outcome` : failure.message;
    reply({
      id, ok: false, error: message,
      commandError: {
        name: failure.name, message,
        ...(isCommandOutcomeUnknownError(failure) ? { refusal: failure.refusal } : {}),
      },
      ...(failure instanceof SessionFailure ? { failureCode: failure.code } : {}),
    });
    return;
  }
  if (failure instanceof WebAssembly.RuntimeError) {
    trapped(id, failure);
    return;
  }
  reply({
    id,
    ok: false,
    error: error instanceof Error ? error.message : String(error),
    ...(failure instanceof SaveUnavailableError ? { code: 'save-unavailable' as const } : {}),
    ...(failure instanceof SourceCutKeptError ? { code: 'cut-kept' as const } : {}),
  });
}

/** A trap leaves the module unusable: `id` is answered as terminal, and nothing succeeds after. */
function trapped(id: number, error: WebAssembly.RuntimeError): void {
  if (trappedIds.has(id)) return;
  trappedIds.add(id);
  trap = error;
  const failed = editFailedAllocationBytes();
  reply({
    id,
    ok: false,
    error:
      failed > 0
        ? `Resident engine worker ran out of memory allocating ${failed} bytes: ${error.message}`
        : `Resident engine worker trapped: ${error.message}`,
    terminal: true,
    ...(failed > 0 ? { outOfMemory: true } : {}),
  });
}

async function handle(request: ResidentEngineWorkerRequest): Promise<void> {
  if (session && request.type === 'sync') validateFontsBaseRevision(request.snapshot);
  if (
    request.type !== 'documentRead' && request.type !== 'fontRequirements' &&
    request.type !== 'encodeState' && request.type !== 'revisionCount' &&
    request.type !== 'save' &&
    request.type !== 'warm' &&
    !(request.type === 'proposal' && request.operation.kind === 'snapshot')
  ) supersedeBackgroundPageBuild();
  if (request.type === 'warm') {
    try {
      if (request.hostModule) {
        const timer = setTimeout(() => resolveHostEditModule(null), RESIDENT_HOST_MODULE_WAIT_MS);
        await preloadEditWasmFrom(
          hostEditModule.then((module) => {
            clearTimeout(timer);
            return module;
          })
        );
      } else {
        await preloadEditWasm();
      }
      reply({ id: request.id, ok: true });
    } catch (error) {
      // No session exists yet, so a failed load is retried by the next request.
      if (request.hostModule) hostEditModule = nextHostEditModule();
      reply({
        id: request.id,
        ok: false,
        error: error instanceof Error ? error.message : String(error),
      });
    }
    return;
  }
  if (request.type === 'destroy') {
    destroySession();
    return;
  }
  if (request.type === 'previewSource') {
    if (!session || !openedViaSource || commandHost) throw new Error('A source preview requires an opened source before commands');
    await replySourcePreview(request.id, session, request.preview, () => request.preview.layoutInput, performance.now());
    return;
  }
  if (request.type === 'commandInit') {
    if (!session || !openedDocument || previewing) {
      throw new Error('Resident commands require a whole opened document');
    }
    if (commandHost) throw new Error('Resident commands are already initialized');
    session.initializeComments(request.request.recoveryState?.comments);
    if (!commandHooksFactory) setResidentCommandHooksFactory(createResidentDocxCommandHooks);
    const init = { ...request.request };
    if (init.provisionalPages !== undefined &&
      (!Number.isSafeInteger(init.provisionalPages) || init.provisionalPages < 1)) {
      throw new TypeError('Invalid provisional pages');
    }
    if (openedViaSource) fontsRevision = -1;
    setFrameDisplayWindow(session, init.displayWindow, init.retainBuiltPages);
    hydrate({
      workerAuthoritative: true, clientId: 0, state: new Uint8Array(0), selection: null,
      fonts: init.fonts, fontsRevision: 0, renderInputs: [], measureInputs: [],
      layoutInput: init.layoutInput, layoutWithRegions: init.layoutWithRegions,
      layoutRevision: init.layoutRevision,
    }, undefined, false, false, undefined, undefined, init.fontRegistrationSlots ? init : undefined,
    !init.layoutWithRegions || init.provisionalPages === undefined);
    sourcePreviewing = false;
    if (init.selection) session.setSelection(init.selection.anchor, init.selection.focus);
    unsubscribe?.();
    unsubscribe = null;
    pendingUpdates = [];
    replyCommand(request.id, 'commandInitial', await bindResidentCommands(init).initial());
    return;
  }
  if (request.type === 'edit') {
    if (!commandLayout) await completeProvisionalLayout();
    noteDocumentChange();
    fontRequirements = null;
    completedLayout = null;
    const result = await requireCommandHost().handler(undefined, request.command);
    replyCommand(request.id, 'commandAck', result);
    return;
  }
  if (request.type === 'checkpoint') {
    replyCommand(request.id, 'checkpoint', await requireCommandHost().checkpoint(undefined));
    return;
  }
  if (request.type === 'requestFullFrame') {
    replyCommand(request.id, 'commandAck', await requireCommandHost().requestFullFrame());
    return;
  }
  if (request.type === 'commandUpdate') {
    requireCommandHost();
    if (!session) throw new Error('Resident engine worker is not initialized');
    if (collaboration) throw new Error('Collaboration merges use the ordered applyRemoteUpdate command');
    noteDocumentChange();
    fontRequirements = null;
    completedLayout = null;
    commandFrameChanges!.updates.push(request.update);
    postCommandFrame(request.id);
    return;
  }
  if (request.type === 'commandViewport') {
    const { start, end, frameEpoch } = request.request;
    if (!Number.isSafeInteger(start) || start < 0 || !Number.isSafeInteger(end) || end <= start ||
      !Number.isSafeInteger(frameEpoch) || frameEpoch < 0) throw new TypeError('Invalid command viewport');
    const host = requireCommandHost();
    if (end > (commandLayout?.certifiedPages() ?? end)) await completeCommandPrefix(host);
    commandFrameChanges!.viewport = request.request;
    postCommandFrame(request.id);
    if (frameEpoch > 0 && commandLayout?.certifiedPages() != null && !slicedCompletion) {
      slicedCompletion = {
        id: 0, expectedFrameEpoch: frameEpoch, paintCaret: false, blocks: COMMAND_COMPLETION_BLOCKS,
        begun: false, restarts: 0, commandOwner: host,
      };
      scheduleCompletion(slicedCompletion);
    }
    return;
  }
  if (request.type === 'openSource') {
    validateResidentSource(request.source);
    const source = request.source;
    if (session && (!previewing || commandHost || source.cut)) throw new Error('Resident engine worker already holds a document');
    const replaced = previewing ? openedCut : null;
    if (session) destroySession(true);
    const started = performance.now();
    let opening: ResidentEngineSession | undefined;
    try {
      const pinned = source.collaboration;
      const reseeded = pinned?.clientClock !== undefined && pinned.initialUpdate === undefined;
      opening = await createResidentEngineSession(source.heapLimitBytes, reseeded ? undefined : pinned?.clientId);
      let opened = adoptResidentSource(opening, source);
      if (opened && !reseeded && pinned?.clientClock !== undefined &&
        stateVectorClock(opening.encodeStateVector(), opening.clientId) < pinned.clientClock) {
        destroySession(true);
        opening = await createResidentEngineSession(source.heapLimitBytes);
        opened = adoptResidentSource(opening, source);
      }
      if (!opened) {
        opening.destroy();
        reply({ id: request.id, ok: true, previewRefused: true });
        return;
      }
      const { document, hostJson } = opened;
      if (source.collaboration) collaboration = createResidentCollaboration(opening);
      const metadata = residentSourceMetadata(document);
      const identity = { clientId: opening.clientId, clientClock: stateVectorClock(opening.encodeStateVector(), opening.clientId) };
      if (source.preview) {
        await replySourcePreview(request.id, opening, typeof source.preview === 'object' ? source.preview : {},
          () => JSON.stringify(buildResidentRegionLayoutRequest(metadata, 24, {})), started, { sourceDocument: metadata, hostJson, ...identity });
      } else {
        const stateVector = exactBuffer(opening.encodeStateVector());
        reply({ id: request.id, ok: true, sourceDocument: metadata, hostJson, stateVector, ...identity }, [stateVector]);
      }
    } catch (error) {
      if (error instanceof WebAssembly.RuntimeError) throw error;
      if (opening && session === opening) destroySession(true);
      else opening?.destroy();
      if (replaced && await reopenSourceCut(replaced)) throw new SourceCutKeptError(error instanceof Error ? error.message : String(error));
      throw error;
    }
    return;
  }
  if (request.type === 'open') {
    // One document per worker, so every queued request addresses the one it was sent for;
    // only a preview gives way, to the whole document.
    if (session && (!previewing || request.previewBlocks !== undefined)) {
      throw new Error('Resident engine worker already holds a document');
    }
    // The preview's memory goes before the whole package seeds; its pages stay painted.
    if (session) destroySession(true);
    openedSource = null;
    editorSaves = { full: false };
    const opening = await createResidentEngineSession(request.heapLimitBytes);
    let hostJson: string | null;
    try {
      if (request.previewBlocks === undefined) opening.setDirectBatches(true);
      hostJson =
        request.previewBlocks === undefined
          ? opening.openDocx(new Uint8Array(request.bytes), request.digest, request.generation)
          : opening.openDocxPreview(
              new Uint8Array(request.bytes),
              request.previewBlocks,
              request.previewParagraphBudget
            );
    } catch (error) {
      if (!(error instanceof WebAssembly.RuntimeError)) opening.destroy();
      throw error;
    }
    if (hostJson === null) {
      opening.destroy();
      reply({ id: request.id, ok: true, previewRefused: true });
      return;
    }
    proposals?.destroy();
    proposals = null;
    session = opening;
    if (request.previewBlocks === undefined) openedVersion = undefined;
    openedDocument = { heapLimitBytes: request.heapLimitBytes };
    previewing = request.previewBlocks !== undefined;
    if (!previewing) {
      openedSource = { bytes: request.bytes, hostJson };
    }
    const stateVector = exactBuffer(session.encodeStateVector());
    reply({ id: request.id, ok: true, hostJson, stateVector }, [stateVector]);
    return;
  }
  if (request.type === 'bootstrap') {
    const replacement = !request.opened ? replacementCommandInit() : null;
    const source = replacement ? openedSource : null;
    if (!request.opened) {
      destroySession(request.keepSurfaces === true);
      openedVersion = null;
      // The worker is a genuine yrs peer. Reusing the main replica's client id
      // makes a fast structural input race overlap one client's clock range and
      // corrupt the update; a fresh id lets yrs merge queued/local operations
      // safely while the main replica applies worker updates with local origin.
      session = await createResidentEngineSession(request.heapLimitBytes);
      if (source) session.openDocx(new Uint8Array(source.bytes), undefined, undefined, false);
    } else if (!session || !openedDocument) {
      throw new Error('Resident engine worker has no opened document');
    } else if (
      request.heapLimitBytes !== undefined &&
      request.heapLimitBytes !== openedDocument.heapLimitBytes
    ) {
      throw new Error('Resident engine worker opened its document under another heap limit');
    }
    unsubscribe?.();
    unsubscribe = null;
    previewFinalPages = null;
    clearProvisionalFinalPages();
    setFrameDisplayWindow(session, request.displayWindow, request.retainBuiltPages);
    const { layoutJson, layoutMeta, provisional } = hydrate(
      request.snapshot,
      request.provisionalPages,
      request.layoutExtras !== undefined,
      request.opened !== true,
      request.layoutReply,
      request.headersFootersEpoch,
      undefined,
      !replacement || !viewportFirst({ ...replacement, layoutWithRegions: request.snapshot.layoutWithRegions })
    );
    if (previewFinalPages !== null || provisionalFinalPages !== null) {
      setFrameDisplayWindow(session, request.displayWindow, request.retainBuiltPages);
    }
    if (provisional) {
      incompleteLayout = {
        layoutInput: request.snapshot.layoutInput,
        workerAuthoritative: request.snapshot.workerAuthoritative,
        extras: request.extras,
        layoutExtras: request.layoutExtras,
      };
    }
    if (replacement) {
      if (source) session.markProjectionStories(session.geometryReader.storyIds());
      openedSource = source;
      const host = bindResidentCommands({
        ...replacement,
        fonts: request.snapshot.fonts,
        layoutInput: request.snapshot.layoutInput,
        layoutWithRegions: request.snapshot.layoutWithRegions,
        layoutRevision: request.snapshot.layoutRevision,
        extras: frameExtras(request.extras, request.layoutExtras, layoutMeta ? headersFootersPayload : undefined),
        displayWindow: request.displayWindow,
        retainBuiltPages: request.retainBuiltPages,
      });
      const initial = await host.initial();
      const ack = isTransferResult(initial) ? initial.value : initial;
      reply({
        id: request.id, ok: true, commandInitial: ack,
        ...ack.frame.payload,
        selection: ack.selection ? { anchor: ack.selection.anchor, head: ack.selection.focus } : null,
        ...(layoutMeta ? { layoutMeta } : {}),
        ...(layoutJson ? { layoutJson } : {}),
        ...(provisional ? { layoutProvisional: true } : {}),
      }, [...(isTransferResult(initial) ? initial.transfer : []),
        ...(layoutMeta ? [layoutMeta.pageSizes.buffer as ArrayBuffer] : [])]);
      return;
    }
    subscribe();
    const started = performance.now();
    const frame = session.buildDisplayListFrame(
      frameExtras(request.extras, request.layoutExtras, layoutMeta ? headersFootersPayload : undefined),
      request.expectedFrameEpoch
    );
    await replyFrame(
      request.id,
      frame,
      performance.now() - started,
      pendingUpdates,
      undefined,
      started,
      false,
      false,
      request.layoutExtras === undefined ? undefined : (layoutMeta ?? layoutJson ?? undefined),
      provisional
    );
    return;
  }
  if (request.type === 'fontRequirements') {
    if (!session) throw new Error('Resident engine worker is not initialized');
    const requirementsJson = layoutFontRequirements(session, request.layoutInput);
    requestedRequirements = { owner: session, layoutInput: request.layoutInput };
    fontRequirements = {
      version: session.proposalEngine.version(),
      layoutInput: request.layoutInput,
      requirementsJson,
    };
    reply({ id: request.id, ok: true, requirementsJson });
    return;
  }
  if (request.type === 'encodeState') {
    if (!session) throw new Error('Resident engine worker is not initialized');
    const state = exactBuffer(session.encodeState());
    let peerMetadata: ArrayBuffer | undefined;
    let peerMetadataReason: string | undefined;
    if (request.peerMetadata) {
      if (!session.encodePeerMetadata) {
        peerMetadataReason = 'missing-capability: Worker cannot encode peer metadata';
      } else {
        try {
          peerMetadata = exactBuffer(session.encodePeerMetadata());
        } catch (error) {
          if (!(error instanceof Error) || error.name !== 'PeerMetadataError') throw error;
          const code = (error as Error & { code?: string }).code;
          peerMetadataReason = `${code ?? 'metadata-error'}: ${error.message}`;
        }
      }
    }
    reply(
      {
        id: request.id,
        ok: true,
        state,
        ...(request.peerMetadata ? { peerMetadata, peerMetadataReason } : {}),
        version: session.proposalEngine.version(),
        proposals: proposals?.exportState() ?? { previewVersion: 0, entries: [] },
      },
      peerMetadata ? [state, peerMetadata] : [state]
    );
    return;
  }
  if (request.type === 'save') {
    if (previewing) throw new SaveUnavailableError('Resident engine worker is still opening');
    if (!session || (!openedSource && !openedParsed)) {
      throw new SaveUnavailableError('Resident engine worker has no opened document to save');
    }
    const held = pendingUpdates;
    pendingUpdates = [];
    const before = session.proposalEngine.version();
    try {
      const projection = commandHost ? session.commentProjection(request.commentsSource !== 'host') : null;
      editorSaves.commentIds = projection?.sessionIds;
      const comments = request.commentsSource === 'host' ? request.comments : projection?.comments ?? request.comments;
      const saved = openedParsed ? await session.saveDocument(request.host, comments, editorSaves) : await session.save(
        new Uint8Array(openedSource!.bytes),
        openedSource!.hostJson,
        request.host,
        comments,
        editorSaves,
        request.stories
      );
      const bytes = saved.slice(0);
      // What the editor copy lacks: the paragraph IDs this save recorded and repairs its updates caused.
      const updates = request.stateVector
        ? [exactBuffer(session.encodeStateAsUpdate(request.stateVector))]
        : [];
      const stateVector = exactBuffer(session.encodeStateVector());
      const answer = { saved: bytes, updates, stateVector, version: session.proposalEngine.version() };
      if (commandHost && answer.version !== before) {
        // Identities the save assigned change the document as a remote update would; the answer waits for their frame.
        noteDocumentChange();
        fontRequirements = null;
        completedLayout = null;
        commandFrameChanges!.layoutStale = true;
        postCommandFrame(request.id, answer, [bytes, ...updates, stateVector]);
      } else reply({ id: request.id, ok: true, ...answer }, [bytes, ...updates, stateVector]);
    } finally {
      pendingUpdates = held;
    }
    return;
  }
  if (request.type === 'revisionCount') {
    if (!session) throw new Error('Resident engine worker is not initialized');
    // Host proposals' revisions are not the document's own.
    const proposed = new Set(proposals?.snapshot().proposals.flatMap((p) => p.revisionIds));
    reply({ id: request.id, ok: true, revisionCount: session.revisionCount(proposed) });
    return;
  }
  if (request.type === 'eraseCaret') {
    caretPaintRect = null;
    if (paintedCaretPageId !== null) await replayOffscreen(false);
    reply({ id: request.id, ok: true });
    return;
  }
  if (!session) throw new Error('Resident engine worker is not initialized');
  if (request.type === 'proposal') {
    if (!(commandHost ? layoutReady : unsubscribe)) {
      throw new Error('Resident engine worker has not laid out its document');
    }
    pendingUpdates = [];
    let committed = false;
    try {
      const registry = proposals ??= createProposalRegistry(session.proposalEngine);
      const previousVersion = session.proposalEngine.version();
      const known = new Set(registry.snapshot().proposals.map((proposal) => proposal.id));
      const since = session.storiesChangedSince(Number.MAX_SAFE_INTEGER).revision;
      let result: DocxProposalResult | undefined;
      let projectionStories: string[] = [];
      switch (request.operation.kind) {
        case 'propose':
          result = registry.propose(request.operation.peerStateVector
            ? { ...request.operation.request, expectVersion: session.proposalEngine.version() }
            : request.operation.request);
          break;
        case 'setStates':
          result = registry.setStates(request.operation.peerStateVector
            ? { ...request.operation.request, expectVersion: session.proposalEngine.version() }
            : request.operation.request);
          break;
        case 'withdraw':
          result = registry.withdraw(request.operation.peerStateVector
            ? { ...request.operation.request, expectVersion: session.proposalEngine.version() }
            : request.operation.request);
          break;
        case 'removeComment':
          try {
            const story = session.selection()?.head.story ?? 'body';
            session.applyRawOps(story, [{ op: 'removeComment', id: request.operation.id }]);
            projectionStories = [story];
          } catch {}
          break;
      }
      committed = request.operation.kind !== 'snapshot';
      const changedStories = session.storiesChangedSince(since).stories;
      if (result?.ok) {
        projectionStories = [...proposalProjectionStories(known, result, changedStories)];
      }
      session.markProjectionStories(projectionStories);
      if (changedStories.length > 0) completedLayout = null;
      const updates = pendingUpdates.map(exactBuffer);
      const peerDiff = 'peerStateVector' in request.operation && request.operation.peerStateVector
        ? exactBuffer(session.encodeStateAsUpdate(request.operation.peerStateVector))
        : undefined;
      const stateVector = exactBuffer(session.encodeStateVector());
      const snapshot = registry.snapshot();
      const version = snapshot.version;
      const geometry = computeProposalGeometryMirror(
        session.geometryReader,
        snapshot,
        version === previousVersion && (
          version === lastProposalMirrorVersion || hasCachedYrsSidebarProjection(session.geometryReader)
        )
      );
      lastProposalMirrorVersion = geometry.version;
      const unchangedFonts =
        fontRequirements?.version === previousVersion &&
        request.operation.kind === 'propose' &&
        result?.ok === true &&
        proposalRevisionPreview(snapshot) === undefined &&
        asciiProposalFontsUnchanged(session, request.operation.request, result);
      fontRequirements =
        fontRequirements && (unchangedFonts ||
          (request.operation.kind === 'snapshot' && fontRequirements.version === snapshot.version))
          ? { ...fontRequirements, version: snapshot.version }
          : null;
      // A decision changes only the preview: the host's next input is its last one previewing it.
      const preview = proposalRevisionPreview(snapshot);
      if (
        request.operation.kind === 'setStates' &&
        result?.ok === true &&
        changedStories.length === 0 &&
        requestedRequirements?.owner === session
      ) {
        fontRequirements = previewFontRequirements(
          session,
          requestedRequirements.layoutInput,
          preview,
          snapshot.version
        );
      }
      reply(
        {
          id: request.id,
          ok: true,
          proposal: {
            ...(result === undefined ? {} : { result }),
            mirror: { version, proposals: registry.exportState() },
            changedStories,
            projectionStories,
            updates,
            ...(peerDiff === undefined ? {} : { peerDiff }),
            stateVector,
            geometry,
            ...(fontRequirements
              ? { fontRequirements: {
                  layoutInput: fontRequirements.layoutInput,
                  requirementsJson: fontRequirements.requirementsJson,
                } }
              : {}),
          },
        },
        [...updates, stateVector, ...(peerDiff === undefined ? [] : [peerDiff])]
      );
    } catch (error) {
      if (trap) throw trap;
      if (error instanceof WebAssembly.RuntimeError) throw error;
      reply({
        id: request.id,
        ok: false,
        error: error instanceof Error ? error.message : String(error),
        ...(committed || pendingUpdates.length > 0 ? { terminal: true } : {}),
      });
    } finally {
      pendingUpdates = [];
    }
    return;
  }
  if (request.type === 'documentRead') {
    const engine = session.proposalEngine;
    if (request.expectVersion !== undefined && engine.version() !== request.expectVersion) {
      reply({ id: request.id, ok: true, superseded: true });
      return;
    }
    let value: unknown;
    let transfer: Transferable[] = [];
    switch (request.read.kind) {
      case 'commentInsertionRange':
        value = readCommentInsertionRange(session, request.read.paraId, request.read.search);
        break;
      case 'commentSelectionRange':
        value = currentSelectionRange(session, request.read.selection);
        break;
      case 'collabStateVector':
        value = session.encodeStateVector();
        break;
      case 'collabDiff':
        value = readResidentCollaborationState(session, request.read.stateVector);
        break;
      case 'collabDiffChunks': {
        const chunks = session.encodeStateAsUpdateChunks(request.read.stateVector, request.read.maxChunkBytes);
        transfer = chunks.map(chunk => chunk.buffer as ArrayBuffer);
        value = chunks;
        break;
      }
      case 'collabPresence':
        value = !retainedFrame || engine.version() !== request.read.expectVersion || retainedLayoutVersion !== engine.version() ||
          request.read.expectFrameEpoch !== retainedFrame.frameEpoch || request.read.expectLayoutRevision !== layoutRevision
          ? null : await readResidentPresence(session, retainedFrame, request.read.peers, request.read.pages, layoutRevision);
        break;
      case 'rawSelectionInfo': {
        const { anchor, focus } = request.read.selection;
        if (anchor.story !== focus.story) throw new Error('Cross-story selection is unavailable');
        value = session.selectionText({ story: anchor.story, start: anchor, end: focus });
        break;
      }
      case 'clipboardSlice': {
        const { anchor, focus } = request.read.selection;
        const cells = session.cellSelection();
        const multiCell = !!cells && (cells.anchor.story !== cells.head.story ||
          cells.anchor.tableIndex !== cells.head.tableIndex ||
          cells.anchor.row !== cells.head.row || cells.anchor.column !== cells.head.column);
        if (multiCell && cells) {
          const payload = clipboardTablePayload(
            session.tablePayload(cells.anchor.story, cells.anchor.tableIndex),
          );
          const holds = !!payload &&
            tableAnchors(payload).anchors.some(
              (cell) => focus.story === cell.story || focus.story.startsWith(`${cell.story}:`),
            );
          if (holds) {
            value = readClipboardSlice(session, { cells });
            break;
          }
        }
        if (anchor.story !== focus.story) throw new Error('Cross-story selection is unavailable');
        value = readClipboardSlice(session, {
          story: anchor.story,
          from: storyOffsetForLoc(session.geometryReader, anchor),
          to: storyOffsetForLoc(session.geometryReader, focus),
        });
        break;
      }
      case 'validateEdits':
        value = engine.validateEdits(request.read.request);
        break;
      case 'contentControlHandles': {
        const frame = retainedFrame && retainedLayoutVersion === engine.version() ? retainedFrame : undefined;
        let queries: ReturnType<typeof createDisplayListQueries> | undefined;
        try {
          value = readControlHandles(session, frame, layoutRevision,
            frame && (() => queries ??= createDisplayListQueries(frame.displayList, loadedRustDisplayListQueryEngine() ?? undefined)));
        } finally { queries?.dispose(); }
        break;
      }
      case 'partSnapshot':
        value = { metadata: partMetadata(session.partDocument()), geometry: retainedFrame && retainedLayoutVersion === engine.version() ? session.partGeometry(retainedFrame) : [],
          frameEpoch: retainedFrame && retainedLayoutVersion === engine.version() ? retainedFrame.frameEpoch : null, layoutRevision };
        break;
      case 'properties':
        value = readResidentProperties(session, request.read.target);
        break;
      case 'sectionProperties':
        value = session.sectionProperties();
        break;
      case 'exportStructuredWithPages': {
        if (commandLayout?.complete()) postCommandFrame(0);
        let currentRequest = request.read.currentRequest;
        if (request.read.workerFonts === true && commandInit?.layoutWithRegions) {
          try {
            const input = JSON.parse(currentRequest);
            const owned = JSON.parse(commandInit.layoutInput);
            if (input?.measurement && typeof input.measurement === 'object' && !Array.isArray(input.measurement) && owned.measurement?.fontChains) {
              input.measurement = { ...input.measurement, fontChains: owned.measurement.fontChains };
              currentRequest = JSON.stringify(input);
            }
          } catch {}
        }
        value = session.exportStructuredWithPagesJson(
          request.read.options,
          resolveLayoutReservationInput(session, currentRequest)
        );
        break;
      }
      case 'listContentControls':
        value = session.listContentControls(request.read.options);
        break;
      case 'findContentControls':
        value = session.findContentControls(request.read.query, request.read.options);
        break;
      case 'rangeHighlights':
        value = !retainedFrame || engine.version() !== request.read.expectVersion || retainedLayoutVersion !== engine.version() ||
          (request.read.expectFrameEpoch !== undefined && request.read.expectFrameEpoch !== retainedFrame.frameEpoch) ||
          (request.read.expectLayoutRevision !== undefined && request.read.expectLayoutRevision !== layoutRevision)
          ? null : await readResidentRangeHighlights(session, retainedFrame, request.read.ranges, layoutRevision);
        break;
      case 'reviewSnapshot': {
        const registry = proposals ??= createProposalRegistry(engine);
        const frame = retainedFrame;
        const layoutVersion = retainedLayoutVersion ?? commandLayout?.certifiedVersion() ?? null;
        const previewVersion = registry.snapshot().previewVersion;
        const published = commandFrameChanges?.reviewFrame;
        const review = !frame || layoutVersion !== engine.version() ||
          (commandHost && (published?.frameEpoch !== frame.frameEpoch || published?.previewVersion !== previewVersion)) ||
          (request.read.expectPreviewVersion !== undefined && request.read.expectPreviewVersion !== previewVersion)
          ? null : await readResidentReviewSnapshot(session, registry, frame, layoutRevision, {
            queries: (at, version) => commandQueries?.(at, version), index: displayPositionIndex(session),
          });
        value = commandHost && review && (review.previewVersion !== previewVersion || review.version !== layoutVersion ||
          registry.snapshot().previewVersion !== previewVersion || engine.version() !== layoutVersion || retainedFrame?.frameEpoch !== frame?.frameEpoch)
          ? null : review;
        break;
      }
      case 'toolbarMetadata':
        value = openedSource ? readResidentToolbarMetadata(openedSource.hostJson) : { styles: [], fontFamilies: [] };
        break;
      case 'paragraphIdentities':
        value = session.paragraphIdentities();
        break;
      case 'resolveParagraphAnchors':
        value = {
          results: request.read.anchors.map((anchor) => engine.resolveParagraphAnchor(anchor)),
        };
        break;
      case 'readParagraphs':
        value = engine.readParagraphs(request.read.request);
        break;
      case 'listStories':
        value = session.listStories();
        break;
      case 'readStories':
        value = session.readStories(request.read.request);
        break;
      case 'findText':
        value = engine.findText(request.read.request);
        break;
      case 'findMatches':
        value = engine.version() !== request.read.expectVersion ? null : findBodyMatches(
          session.geometryReader,
          residentBodyPositions(session.geometryReader),
          request.read.searchText,
          request.read.options
        );
        break;
      case 'navigationTarget':
        value = resolveNavigationTarget(
          session.geometryReader,
          request.read.story,
          request.read.paraId
        );
        break;
      case 'reviewTarget':
        value = readReviewTarget(session, request.read.target);
        break;
      case 'anchorTargets': {
        const reader = session.geometryReader;
        value = computeAnchorTargetGeometry(reader, request.read.targets, () =>
          hiddenRanges(reader, reader.version(), proposals?.snapshot() ?? null)
        );
        break;
      }
      case 'searchText':
        value = readResidentSearch(
          {
            ...session.geometryReader,
            searchText: session.searchText,
            resolveStickyPosition: session.resolveStickyPosition,
          },
          request.read.query,
          request.read.caseSensitive,
          request.read.carry
        );
        break;
      case 'pointPosition':
        value = request.read.nativePoint ? retainedFrame && retainedLayoutVersion === engine.version()
          ? await readResidentPointer(session, retainedFrame, layoutRevision, request.read.nativePoint, request.read.expectVersion) : null : resolveYrsPointPosition(
          displayPositionIndex(session),
          request.read.hit,
          request.read.expectVersion
        );
        break;
      case 'findParagraphs':
        value = findParagraphs(session.geometryReader, request.read.query, request.read);
        break;
      case 'selectionUnit':
        value = resolveSelectionUnit(
          displayPositionIndex(session),
          request.read.story,
          request.read.position,
          request.read.unit,
          request.read.expectVersion
        );
        break;
      case 'selectionText':
        value = resolveSelectionText(
          displayPositionIndex(session),
          request.read.story,
          request.read.anchor,
          request.read.head,
          request.read.expectVersion
        );
        break;
      case 'bookmarkTarget':
        value = resolveBookmarkTarget(displayPositionIndex(session), request.read.story, request.read.name, engine.version());
        break;
      case 'bookmarkPosition':
        value = resolveBookmarkPosition(
          displayPositionIndex(session),
          request.read.story,
          request.read.name,
          request.read.expectVersion
        );
        break;
      case 'selectionInfo':
        value = resolveSelectionInfo(
          displayPositionIndex(session),
          request.read.story,
          request.read.anchor,
          request.read.head,
          request.read.expectVersion
        );
        break;
      case 'paragraphTarget':
        value = resolveParagraphTarget(
          displayPositionIndex(session),
          request.read.story,
          request.read.paraId,
          request.read.expectVersion
        );
        break;
      case 'commentTarget':
        value = resolveCommentTarget(
          displayPositionIndex(session),
          request.read.story,
          request.read.commentId,
          request.read.expectVersion
        );
        break;
      case 'revisionTarget':
        value = resolveRevisionTarget(
          displayPositionIndex(session),
          request.read.story,
          request.read.revisionId,
          request.read.expectVersion
        );
        break;
      case 'sidebar':
        value = readSidebar(session.geometryReader, request.read.commentIds, request.read.expectVersion);
        break;
      case 'headings':
        value = readOutlineHeadings(session.geometryReader, request.read.expectVersion);
        break;
      case 'adjacentRevision':
        value = readAdjacentRevision(session.geometryReader, request.read.selection, request.read.direction,
          new Set(proposals?.snapshot().proposals.flatMap((proposal) => proposal.revisionIds)));
        break;
      case 'stickyAnchors': {
        const currentSession = session;
        value = request.read.locs.map((loc) => {
          try {
            return readCommandStickyPosition(currentSession, loc);
          } catch {
            return null;
          }
        });
        break;
      }
    }
    reply({ id: request.id, ok: true, read: { version: engine.version(), value } }, transfer);
    return;
  }
  if (request.type === 'layoutJson') {
    if (!retainedRegions || request.layoutRevision !== layoutRevision ||
      retainedLayoutVersion !== session.proposalEngine.version()) {
      reply({ id: request.id, ok: true, layoutJsonStatus: 'stale' });
    } else {
      reply({ id: request.id, ok: true, layoutJsonStatus: 'ok',
        layoutRevision, layoutJson: session.retainedLayoutJson() });
    }
    return;
  }
  if (request.type === 'sync') {
    unsubscribe?.();
    unsubscribe = null;
    previewFinalPages = null;
    clearProvisionalFinalPages();
    setFrameDisplayWindow(session, request.displayWindow, request.retainBuiltPages);
    const { layoutJson, layoutMeta, provisional } = hydrate(
      request.snapshot,
      request.provisionalPages,
      request.layoutExtras !== undefined,
      true,
      request.layoutReply,
      request.headersFootersEpoch
    );
    if (previewFinalPages !== null || provisionalFinalPages !== null) {
      setFrameDisplayWindow(session, request.displayWindow, request.retainBuiltPages);
    }
    if (provisional) {
      incompleteLayout = {
        layoutInput: request.snapshot.layoutInput,
        workerAuthoritative: request.snapshot.workerAuthoritative,
        extras: request.extras,
        layoutExtras: request.layoutExtras,
      };
    }
    subscribe();
    const started = performance.now();
    const frame = session.buildDisplayListFrame(
      frameExtras(request.extras, request.layoutExtras, layoutMeta ? headersFootersPayload : undefined),
      request.expectedFrameEpoch
    );
    await replyFrame(
      request.id,
      frame,
      performance.now() - started,
      pendingUpdates,
      undefined,
      started,
      false,
      request.paintCaret,
      request.layoutExtras === undefined ? undefined : (layoutMeta ?? layoutJson ?? undefined),
      provisional
    );
    return;
  }
  if (request.type === 'buildPages') {
    const limit = previewFinalPages === null
      ? provisionalFinalPages
      : Math.min(previewFinalPages, provisionalFinalPages ?? previewFinalPages);
    const pages = limit === null ? request.pages : request.pages.filter((index) => index < limit);
    if (commandHost) {
      for (const page of pages) commandFrameChanges!.pages.add(page);
      postCommandFrame(request.id);
      return;
    }
    if (request.background && pages.length > BACKGROUND_SLICE_PAGES) {
      const build: BackgroundPageBuild = {
        request: pages === request.pages ? request : { ...request, pages },
        owner: session, frameEpoch: request.expectedFrameEpoch,
        frames: [], offset: 0, started: performance.now(), engineMs: 0,
        pagesPerSlice: BACKGROUND_SLICE_PAGES,
      };
      backgroundPageBuild = build;
      scheduler.schedule({
        kind: 'pageBuild', version: scheduler.version, generation: scheduler.generation,
        onStale: () => 'cancel',
        cancel: () => {
          if (backgroundPageBuild === build) supersedeBackgroundPageBuild();
        },
        fail: (error) => {
          if (backgroundPageBuild === build) backgroundPageBuild = null;
          replyFailure(build.request.id, error);
        },
        run: async (budgetMs) => {
          if (backgroundPageBuild !== build) return 'done';
          if (trap) throw trap;
          handlingId = build.request.id;
          return (await backgroundPageSlice(build, budgetMs)) ? 'done' : 'yield';
        },
      });
      return;
    }
    setFrameDisplayWindow(session);
    // Pages of the provisional frame build between steps, as before a completion.
    pendingUpdates = [];
    const started = performance.now();
    const frame = session.buildDisplayPagesFrame(pages, request.expectedFrameEpoch);
    await replyFrame(
      request.id,
      frame,
      performance.now() - started,
      pendingUpdates,
      undefined,
      started,
      false,
      request.paintCaret
    );
    return;
  }
  if (request.type === 'releasePages') {
    if (
      !retainedFrame ||
      retainedFrame.frameEpoch !== request.expectedFrameEpoch ||
      request.pages.some(
        ({ index, pageId }) => retainedFrame!.pages[index]?.pageId.toString() !== pageId
      )
    ) {
      reply({ id: request.id, ok: true, superseded: true });
      return;
    }
    const started = performance.now();
    const frame = session.releaseDisplayPagesFrame(
      request.pages.map(({ index }) => index),
      request.expectedFrameEpoch
    );
    if (frame === null) {
      reply({ id: request.id, ok: true, superseded: true });
      return;
    }
    await replyFrame(
      request.id,
      frame,
      performance.now() - started,
      [],
      undefined,
      started,
      false,
      request.paintCaret
    );
    return;
  }
  if (request.type === 'completeLayout') {
    if (!commandLayout) setFrameDisplayWindow(session);
    if (incompleteLayout && request.sliceBlocks && !commandLayout) {
      supersedeSlicedCompletion();
      slicedCompletion = {
        id: request.id,
        expectedFrameEpoch: request.expectedFrameEpoch,
        paintCaret: request.paintCaret,
        blocks: request.sliceBlocks,
        begun: false,
        restarts: 0,
      };
      scheduleCompletion(slicedCompletion);
      return;
    }
    await completeProvisionalLayout();
    await replyCompletedLayout(request.id, request.expectedFrameEpoch, request.paintCaret);
    return;
  }
  if (request.type === 'buildFrame') {
    if (!incompleteLayout?.workerAuthoritative) await completeProvisionalLayout();
    setFrameDisplayWindow(session, request.displayWindow, request.retainBuiltPages);
    pendingUpdates = [];
    const started = performance.now();
    const frame = session.buildDisplayListFrame(request.extras, request.expectedFrameEpoch);
    await replyFrame(
      request.id,
      frame,
      performance.now() - started,
      pendingUpdates,
      undefined,
      started,
      false,
      request.paintCaret
    );
    return;
  }
  if (request.type === 'applyUpdate' || request.type === 'syncUpdate') {
    noteDocumentChange();
    fontRequirements = null;
    if (request.update.length > 0 || request.type === 'applyUpdate') session.applyUpdate(request.update);
    if (request.type === 'applyUpdate') {
      if (request.selection) session.setSelection(request.selection.anchor, request.selection.head);
    } else {
      const repair = exactBuffer(session.encodeStateAsUpdate(request.stateVector));
      const stateVector = exactBuffer(session.encodeStateVector());
      reply({
        id: request.id,
        ok: true,
        version: session.proposalEngine.version(),
        stateVector,
        repair,
      }, [stateVector, repair]);
    }
    return;
  }
  if (request.type === 'attachCanvases') {
    const environmentChanged =
      offscreenCanvases.size > 0 &&
      (offscreenDpr !== request.devicePixelRatio || offscreenZoom !== request.zoom);
    // Pages needing pixels: freshly transferred canvases plus retained
    // canvases re-entering the active window (their buffers were zeroed when
    // they left it).
    const forcedPageIds = new Set(request.pages.map((page) => page.pageId));
    for (const pageId of request.activePageIds) {
      if (!activeOffscreenPageIds.has(pageId)) forcedPageIds.add(pageId);
    }
    for (const { pageId, canvas } of request.pages) offscreenCanvases.set(pageId, canvas);
    activeOffscreenPageIds = new Set(request.activePageIds);
    offscreenDpr = request.devicePixelRatio;
    offscreenZoom = request.zoom;
    caretStyle = request.caretStyle;
    for (const [pageId, canvas] of offscreenCanvases) {
      if (!activeOffscreenPageIds.has(pageId)) {
        // out of the page window: release the bitmap but KEEP the canvas —
        // a transferred surface can never be re-transferred, so the element
        // must stay usable for re-entry
        releaseOffscreenPageCanvas(canvas);
        offscreenBackBuffers.delete(pageId);
        forgetOffscreenPagePixels(pageId);
      }
    }
    // A dpr/zoom change repaints every active page; otherwise only the forced
    // set needs pixels — surviving active pages keep their bitmaps.
    await replayOffscreen(environmentChanged ? true : forcedPageIds);
    reply({ id: request.id, ok: true });
    return;
  }
  await completeProvisionalLayout();
  noteDocumentChange();
  fontRequirements = null;
  // The edit replaces the pagination a cached completion's frame would paint.
  completedLayout = null;
  setFrameDisplayWindow(session, request.displayWindow, request.retainBuiltPages);
  session.setSelection(request.selection.anchor, request.selection.head);
  pendingUpdates = [];
  const started = performance.now();
  try {
    const applied =
      request.type === 'applyDelete'
        ? request.profile
          ? session.applyDeleteProfiled(
              request.direction,
              request.expectedFrameEpoch,
              request.count
            )
          : {
              frame: session.applyDelete(
                request.direction,
                request.expectedFrameEpoch,
                request.count
              ),
              profile: undefined,
            }
        : request.profile
          ? session.applyInputProfiled(request.text, request.expectedFrameEpoch)
          : {
              frame: session.applyInput(request.text, request.expectedFrameEpoch),
              profile: undefined,
            };
    // Typing inside a surrogate pair lands before it and settles the caret after the text; the
    // host's sticky caret stays inside the pair, so it takes this one.
    let settled: CollaborationCursor | null = null;
    if (request.type === 'applyInput') {
      const sent = request.selection.head;
      const head = session.selection()?.head;
      const expected = sent.offset + request.text.length;
      if (head && (head.paraId !== sent.paraId || head.offset !== expected)) {
        settled = session.encodeSelection();
      }
    }
    await replyFrame(
      request.id,
      applied.frame,
      performance.now() - started,
      pendingUpdates,
      applied.profile,
      started,
      request.selection.head.story === 'body',
      request.paintCaret,
      undefined,
      false,
      request.type === 'applyDelete' ? session.residentDeletedUnits() : undefined,
      [],
      settled
    );
  } catch (error) {
    if (trap) throw trap;
    if (error instanceof WebAssembly.RuntimeError) throw error;
    const message = error instanceof Error ? error.message : String(error);
    reply({
      id: request.id,
      ok: false,
      error: message,
      residentUnavailable: message.includes('resident input state is not ready'),
      // The edit committed here but never reached the host: this replica is
      // no longer the host's, so the host must replace it.
      ...(pendingUpdates.length > 0 ? { terminal: true } : {}),
    });
  } finally {
    pendingUpdates = [];
  }
}

function validateFontsBaseRevision(snapshot: YrsResidentWorkerSnapshot): void {
  if (snapshot.fontsBaseRevision !== undefined && snapshot.fontsBaseRevision !== fontsRevision) {
    throw new Error('Resident engine worker font base revision mismatch');
  }
}

const appliedLayoutReservations = new WeakMap<ResidentEngineSession, {
  version: string;
  key: string;
  reservations: DocxResolvedReservations;
}>();

function resolveLayoutReservationInput(
  engine: ResidentEngineSession,
  layoutInput: string,
  apply = false
): string {
  if (!layoutInput.includes('"reservationSets"')) {
    if (apply) appliedLayoutReservations.delete(engine);
    return layoutInput;
  }
  const input = JSON.parse(layoutInput) as { renderEnv?: YrsRenderEnv };
  const payload = input.renderEnv?.reservationSets;
  if (payload === undefined) {
    if (apply) appliedLayoutReservations.delete(engine);
    return layoutInput;
  }
  const version = engine.proposalEngine.version();
  const previous = appliedLayoutReservations.get(engine);
  const key = JSON.stringify(payload);
  let reservations = previous?.reservations ?? {};
  if (validReservationsPayload(payload) && payload.version === version) {
    reservations = previous?.version === version && previous.key === key
      ? previous.reservations
      : resolveReservations(engine.geometryReader, payload, version, previous?.reservations ?? null);
    if (apply) appliedLayoutReservations.set(engine, { version, key, reservations });
  }
  delete input.renderEnv!.reservationSets;
  input.renderEnv!.reservations = reservations;
  return JSON.stringify(input);
}

/**
 * Loads a snapshot and runs its layout, over the first `provisionalPages`
 * pages only when given; returns the region layout reply, which a full pass
 * serializes only when `reply` asks for it. Without `layout` the command
 * session it binds lays the document out.
 */
function hydrate(
  snapshot: YrsResidentWorkerSnapshot,
  provisionalPages?: number,
  reply = true,
  loadState = true,
  replyMode?: 'meta',
  knownHeadersFootersEpoch?: number,
  commandFonts?: ResidentCommandInitRequest,
  layout = true
): { layoutJson: string | null; layoutMeta?: LayoutMetaV1; provisional: boolean } {
  if (!session) throw new Error('Resident engine worker is not initialized');
  validateFontsBaseRevision(snapshot);
  layoutReady = false;
  supersedeSlicedCompletion();
  incompleteLayout = null;
  clearProvisionalFinalPages();
  completedLayout = null;
  if (loadState && !snapshot.workerAuthoritative) {
    noteDocumentChange();
    fontRequirements = null;
    session.loadState(snapshot.state);
  }
  session.setPartialDocument(snapshot.partialDocument === true);
  previewFinalPages = snapshot.partialDocument === true ? 0 : null;
  if (!snapshot.workerAuthoritative) {
    session.loadMediaSources(snapshot.mediaSources ?? '');
    session.loadNoteSeparators(snapshot.noteSeparators ?? new Uint8Array(0));
  }
  if (snapshot.fontsRevision !== fontsRevision) {
    if (snapshot.fontsBaseRevision === undefined) {
      session.clearFonts();
      glyphCache = null;
      for (const pageId of activeOffscreenPageIds) pendingOffscreenPageIds.add(pageId);
      intactBackBuffers.clear();
    }
    fontsRevision = -1;
    if (commandFonts) {
      const registered = registerCommandFonts(session, snapshot.fonts);
      const chains = typeof commandFonts.fontRegistrationSlots === 'object'
        ? commandFonts.fontRegistrationSlots.fontChains : undefined;
      snapshot.layoutInput = registered.remap(snapshot.layoutInput, chains);
      commandFonts.layoutInput = snapshot.layoutInput;
      commandFonts.extras = registered.remap(commandFonts.extras, chains);
      commandFonts.fontRegistrationSlots = undefined;
      commandFonts.fonts = [];
    } else {
      for (const font of snapshot.fonts) {
        if (font instanceof Uint8Array) session.registerFont(font);
        else session.registerSubstituteFont(font.substituteOf, font.family);
      }
    }
    fontsRevision = snapshot.fontsRevision;
  }
  snapshot.layoutInput = resolveLayoutReservationInput(session, snapshot.layoutInput, true);
  for (const { story, env } of snapshot.renderInputs) session.yrsBlocksForStory(story, env);
  for (const input of snapshot.measureInputs) session.measureParagraphJson(input);
  if (!layout) {
    layoutRevision = snapshot.layoutRevision;
    retainedRegions = snapshot.layoutWithRegions;
    retainedLayoutVersion = null;
    pendingUpdates = [];
    return { layoutJson: null, provisional: false };
  }
  let layoutJson: string | null = null;
  let metadata: RetainedLayoutMeta | undefined;
  let provisional = false;
  let pageCount: number | null = null;
  if (snapshot.layoutWithRegions && provisionalPages !== undefined) {
    layoutJson = session.layoutDocumentWithRegionsPrefixRetainedJson(
      snapshot.layoutInput,
      provisionalPages
    );
    const layout = JSON.parse(layoutJson) as {
      provisional?: boolean;
      layout: { pages: unknown[] };
    };
    provisional = layout.provisional === true;
    pageCount = layout.layout.pages.length;
  } else if (snapshot.layoutWithRegions && !reply) {
    session.layoutDocumentWithRegionsRetained(snapshot.layoutInput);
  } else if (snapshot.layoutWithRegions && replyMode === 'meta') {
    metadata = session.layoutDocumentWithRegionsRetainedMeta(snapshot.layoutInput);
    provisional = metadata.provisional;
    pageCount = metadata.pageCount;
  } else if (snapshot.layoutWithRegions) {
    // the retained reply leaves out the tens-of-MB measured arena
    layoutJson = session.layoutDocumentWithRegionsRetainedJson(snapshot.layoutInput);
    if (snapshot.partialDocument === true) {
      pageCount = (JSON.parse(layoutJson) as { layout: { pages: unknown[] } }).layout.pages.length;
    }
  } else {
    session.layoutDocumentJson(snapshot.layoutInput);
  }
  if (!snapshot.workerAuthoritative && snapshot.selection) {
    session.setSelection(snapshot.selection.anchor, snapshot.selection.head);
  }
  layoutRevision = snapshot.layoutRevision;
  retainedRegions = snapshot.layoutWithRegions;
  retainedLayoutVersion = session.proposalEngine.version();
  layoutReady = true;
  pendingUpdates = [];
  previewFinalPages = finalPreviewPageCount(
    snapshot.partialDocument,
    provisional,
    provisionalPages,
    pageCount
  );
  provisionalFinalPages = provisional && snapshot.partialDocument !== true
    ? provisionalPages ?? null
    : null;
  return {
    layoutJson,
    ...(metadata ? { layoutMeta: layoutMetaReply(metadata, knownHeadersFootersEpoch) } : {}),
    provisional,
  };
}

function retainedHeadersFootersPayload(): string {
  const payload = session?.retainedHeadersFootersJson() ?? 'null';
  if (payload !== headersFootersPayload) {
    headersFootersPayload = payload;
    headersFootersEpoch += 1;
  }
  return payload;
}

function layoutMetaReply(meta: RetainedLayoutMeta, knownEpoch = sentHeadersFootersEpoch): LayoutMetaV1 {
  const payload = retainedHeadersFootersPayload();
  const changed = knownEpoch !== headersFootersEpoch;
  return {
    v: LAYOUT_META_VERSION,
    layoutRevision,
    ...meta,
    headersFootersEpoch,
    ...(changed ? { headersFooters: payload } : {}),
  };
}

function forgetRequirementsCache(): void {
  requirementsCache?.release();
  requirementsCache = null;
}

function layoutFontRequirements(engine: ResidentEngineSession, layoutInput: string): string {
  const version = engine.proposalEngine.version();
  if (requirementsCache?.owner !== engine || requirementsCache.version !== version) {
    forgetRequirementsCache();
    requirementsCache = {
      owner: engine,
      version,
      byInput: new Map(),
      release: engine.onUpdate(forgetRequirementsCache),
    };
  }
  const cached = requirementsCache.byInput.get(layoutInput);
  if (cached !== undefined) return cached;
  const requirementsJson = engine.layoutFontRequirementsJson(resolveLayoutReservationInput(engine, layoutInput));
  if (requirementsCache.byInput.size >= REQUIREMENTS_CACHE_INPUTS) {
    requirementsCache.byInput.delete(requirementsCache.byInput.keys().next().value as string);
  }
  requirementsCache.byInput.set(layoutInput, requirementsJson);
  return requirementsJson;
}

/**
 * The font requirements of `layoutInput` previewing `preview`, keyed by the input the host builds
 * for it (its render environment with `revisionPreview` replaced), or null when they cannot be read.
 */
function previewFontRequirements(
  engine: ResidentEngineSession,
  layoutInput: string,
  preview: ReturnType<typeof proposalRevisionPreview>,
  version: string
): NonNullable<typeof fontRequirements> | null {
  try {
    if (engine.proposalEngine.version() !== version) return null;
    const request = JSON.parse(layoutInput) as { renderEnv?: Record<string, unknown> | null };
    if (!request.renderEnv || typeof request.renderEnv !== 'object') return null;
    if (preview === undefined) delete request.renderEnv.revisionPreview;
    else request.renderEnv.revisionPreview = preview;
    const next = JSON.stringify(request);
    return { version, layoutInput: next, requirementsJson: layoutFontRequirements(engine, next) };
  } catch (error) {
    // The host reads them itself and meets the failure there.
    if (trap) throw trap;
    if (error instanceof WebAssembly.RuntimeError) throw error;
    return null;
  }
}

function setFrameDisplayWindow(
  engine: ResidentEngineSession,
  window?: [number, number],
  retainBuiltPages?: boolean
): void {
  if (window) commandDisplayWindow = window;
  if (previewFinalPages !== null) {
    engine.setDisplayWindow(...finalPreviewDisplayWindow(window, previewFinalPages));
    engine.setDisplayRetainBuiltPages(window !== undefined && retainBuiltPages === true);
    engine.setWindowedIncrementalBuilds(true);
    return;
  }
  if (window) {
    if (provisionalFinalPages !== null) {
      provisionalDisplayWindow = { window, retainBuiltPages: retainBuiltPages === true };
      engine.setDisplayWindow(...finalPreviewDisplayWindow(window, provisionalFinalPages));
      engine.setDisplayRetainBuiltPages(false);
    } else {
      engine.setDisplayWindow(...window);
      engine.setDisplayRetainBuiltPages(retainBuiltPages === true);
    }
  }
  engine.setWindowedIncrementalBuilds(window !== undefined);
}

function clearProvisionalFinalPages(): void {
  provisionalFinalPages = null;
  if (session && provisionalDisplayWindow) {
    session.setDisplayWindow(...provisionalDisplayWindow.window);
    session.setDisplayRetainBuiltPages(provisionalDisplayWindow.retainBuiltPages);
  }
  provisionalDisplayWindow = null;
}

/**
 * Replaces a provisional layout with the full one before anything reads it,
 * finishing a sliced completion at once and answering its request first. A
 * command session's certified prefix completes inside the request's frame work
 * and publishes its completed frame instead.
 */
async function completeProvisionalLayout(): Promise<void> {
  if (commandLayout?.complete()) {
    postCommandFrame(0);
    return;
  }
  if (!session || !incompleteLayout) return;
  const waiting = slicedCompletion;
  slicedCompletion = null;
  const { layoutInput, ...request } = incompleteLayout;
  incompleteLayout = null;
  try {
    let layoutJson: string | undefined;
    if (waiting?.begun) {
      try {
        layoutJson = session.resumeRegionLayout(ALL_BLOCKS).layoutJson;
      } catch (error) {
        if (error instanceof WebAssembly.RuntimeError) throw error;
      }
    }
    completedLayout = {
      ...request,
      layoutJson: layoutJson ?? session.layoutDocumentWithRegionsRetainedJson(layoutInput),
      headersFootersJson: session.retainedHeadersFootersJson(),
    };
    clearProvisionalFinalPages();
    if (waiting) {
      await replyCompletedLayout(waiting.id, waiting.expectedFrameEpoch, waiting.paintCaret);
    }
  } catch (error) {
    // The waiting completion is answered too, and the request that finished it fails.
    if (waiting) replyFailure(waiting.id, error);
    throw error;
  }
}

/** Answers a `completeLayout` with the completed layout and its frame, or with none. */
async function replyCompletedLayout(
  id: number,
  expectedFrameEpoch: number,
  paintCaret: boolean
): Promise<void> {
  const completed = completedLayout;
  completedLayout = null;
  if (!session || !completed) {
    reply({ id, ok: true });
    return;
  }
  setFrameDisplayWindow(session);
  pendingUpdates = [];
  const started = performance.now();
  const frame = session.buildDisplayListFrame(
    frameExtras(completed.extras, completed.layoutExtras, completed.headersFootersJson),
    expectedFrameEpoch
  );
  await replyFrame(
    id,
    frame,
    performance.now() - started,
    pendingUpdates,
    undefined,
    started,
    false,
    paintCaret,
    completed.layoutJson
  );
}

// A zero-delay turn of the event loop, so requests that arrived meanwhile
// queue ahead of the next step.
const completionTurns = typeof MessageChannel === 'function' ? new MessageChannel() : null;
const turnCallbacks: Array<() => void> = [];
if (completionTurns) completionTurns.port1.onmessage = () => turnCallbacks.shift()?.();

function nextTurn(callback: () => void): void {
  if (!completionTurns) {
    setTimeout(callback, 0);
    return;
  }
  turnCallbacks.push(callback);
  completionTurns.port2.postMessage(null);
}

const scheduler = createResidentScheduler({
  now: () => performance.now(),
  turn: (callback) => nextTurn(callback),
  timer: (callback, ms) => {
    const handle = setTimeout(callback, ms);
    return () => clearTimeout(handle);
  },
  failed: () => {},
});

function supersedeBackgroundPageBuild(): void {
  const build = backgroundPageBuild;
  if (!build) return;
  backgroundPageBuild = null;
  replyDropped({ id: build.request.id, ok: true, pageBuildSuperseded: true });
}

async function backgroundPageSlice(build: BackgroundPageBuild, budgetMs: number): Promise<boolean> {
  if (session !== build.owner || (build.offset > 0 && retainedFrame?.frameEpoch !== build.frameEpoch)) {
    supersedeBackgroundPageBuild();
    return true;
  }
  setFrameDisplayWindow(build.owner);
  const started = performance.now();
  const pages = build.request.pages.slice(build.offset, build.offset + build.pagesPerSlice);
  const bytes = build.owner.buildDisplayPagesFrame(pages, build.frameEpoch);
  const elapsed = performance.now() - started;
  build.engineMs += elapsed;
  build.pagesPerSlice = Math.min(
    32,
    Math.max(BACKGROUND_SLICE_PAGES, Math.round((pages.length * budgetMs) / Math.max(1, elapsed)))
  );
  build.offset += pages.length;
  if (build.offset === build.request.pages.length) {
    backgroundPageBuild = null;
    await replyFrame(
      build.request.id, bytes, build.engineMs, [], undefined, build.started,
      false, build.request.paintCaret, undefined, false, undefined, build.frames
    );
    return true;
  }
  applyWorkerFrame(bytes);
  build.frameEpoch = retainedFrame!.frameEpoch;
  build.frames.push(bytes);
  return false;
}

function scheduleCompletion(completion: SlicedCompletion): void {
  scheduler.schedule({
    kind: 'completion', version: scheduler.version, generation: scheduler.generation,
    idleAfterInputMs: COMPLETION_IDLE_MS,
    onStale: () => 'continue',
    fail: (error) => {
      if (slicedCompletion === completion) slicedCompletion = null;
      if (!completion.commandOwner) replyFailure(completion.id, error);
      else if (completion.commandOwner === commandHost) {
        if (error instanceof WebAssembly.RuntimeError) replyFailure(0, error);
        else reply({ id: 0, ok: false, terminal: true, failureCode: 'message',
          error: error instanceof Error ? error.message : String(error) });
      }
    },
    run: async (budgetMs) => {
      if (slicedCompletion !== completion) return 'done';
      if (trap) throw trap;
      handlingId = completion.id;
      const owner = completion.commandOwner;
      if (owner) await owner.withFrameWork(() => completionSlice(completion, budgetMs));
      else await completionSlice(completion, budgetMs);
      return slicedCompletion === completion ? 'yield' : 'done';
    },
  });
}

async function completionSlice(completion: SlicedCompletion, budgetMs: number): Promise<void> {
  if (!session || slicedCompletion !== completion || !incompleteLayout) return;
  const owner = completion.commandOwner ? commandLayout : null;
  if (completion.commandOwner && (!owner || commandHost !== completion.commandOwner)) return;
  let progress;
  if (!completion.begun) {
    progress = owner ? owner.begin() : session.beginRegionLayout(incompleteLayout.layoutInput);
    completion.begun = true;
  } else {
    const started = performance.now();
    try {
      progress = owner ? owner.resume(completion.blocks) : session.resumeRegionLayout(completion.blocks);
    } catch (error) {
      if (error instanceof WebAssembly.RuntimeError) throw error;
      const message = error instanceof Error ? error.message : String(error);
      if (
        message !== 'no region layout to resume' &&
        message !== 'the document or its fonts changed since the region layout began'
      ) throw error;
      // A change in between abandoned the pass: begin again on the new state. Host
      // proposals the worker holds keep yielding to user requests; other changes
      // finish in one step once they keep coming.
      completion.begun = false;
      completion.restarts += 1;
      if (
        completion.restarts > COMPLETION_RESTARTS &&
        !incompleteLayout.workerAuthoritative &&
        !proposals
      ) {
        try {
          await completeProvisionalLayout();
        } catch {
          // Answered as the completion's failure.
        }
        return;
      }
      return;
    }
    const elapsed = Math.max(1, performance.now() - started);
    completion.blocks = Math.min(
      8192,
      Math.max(8, Math.round((completion.blocks * budgetMs) / elapsed))
    );
  }
  if (progress.layoutJson === undefined) {
    return;
  }
  if (owner) {
    slicedCompletion = null;
    postCommandFrame(0);
    return;
  }
  const { layoutInput: _input, ...request } = incompleteLayout;
  incompleteLayout = null;
  clearProvisionalFinalPages();
  slicedCompletion = null;
  completedLayout = {
    ...request,
    layoutJson: progress.layoutJson,
    headersFootersJson: session.retainedHeadersFootersJson(),
  };
  await replyCompletedLayout(completion.id, completion.expectedFrameEpoch, completion.paintCaret);
}

/** A snapshot or a new session replaces the layout a sliced completion was finishing. */
function supersedeSlicedCompletion(): void {
  const completion = slicedCompletion;
  slicedCompletion = null;
  if (completion && !completion.commandOwner) replyDropped({ id: completion.id, ok: true });
}

/**
 * The extras a frame is built with. For a layout this worker owns, the host
 * sends them without the header/footer payload, which only this layout has:
 * the session retains it after a region layout,
 * or a completed layout captured it.
 */
function frameExtras(
  extras: string,
  layoutExtras: string | undefined,
  headersFootersJson?: string
): string {
  if (layoutExtras === undefined) return extras;
  const retained =
    headersFootersJson ??
    (retainedRegions ? retainedHeadersFootersPayload() : undefined);
  const headersFooters =
    retained === undefined
      ? undefined
      : (JSON.parse(retained) as DisplayListBuildInputs['headersFooters']);
  return encodeDisplayListFrameExtras({
    ...(JSON.parse(layoutExtras) as DisplayListBuildInputs),
    ...(headersFooters ? { headersFooters } : {}),
  });
}

function subscribe(): void {
  if (!session) return;
  unsubscribe = session.onUpdate((update) => pendingUpdates.push(update.slice()));
}

function asciiProposalFontsUnchanged(
  engine: ResidentEngineSession,
  request: DocxProposalRequest,
  result: DocxProposalResult
): boolean {
  if (!result.ok || !fontRequirements) return false;
  const requirements = JSON.parse(fontRequirements.requirementsJson) as { scripts?: string[] }[];
  const input = JSON.parse(fontRequirements.layoutInput) as {
    renderEnv?: { revisionPreview?: Record<string, unknown> };
  };
  if (requirements.some(({ scripts }) => scripts && scripts.length > 0) ||
    Object.keys(input.renderEnv?.revisionPreview ?? {}).length > 0) return false;
  const stories = new Map<string, Set<string>>();
  for (const proposal of request.proposals) {
    const text = proposal.op === 'replaceText' ? proposal.replaceWith : proposal.text;
    if (!/^[\x20-\x7e]*$/.test(text)) return false;
    const record = result.snapshot.proposals.find(({ id }) => id === proposal.id);
    if (!record?.changed) continue;
    const { story, paraId } = record.paragraph;
    const paragraphs = stories.get(story) ?? new Set<string>();
    paragraphs.add(paraId);
    stories.set(story, paragraphs);
  }
  for (const [story, targets] of stories) {
    const found = new Map<string, number>();
    const spans = engine.geometryReader.paragraphSpans(story);
    for (let index = 0; index < spans.length; index += 1) {
      const { paraId } = spans[index]!;
      if (!targets.has(paraId)) continue;
      if (found.has(paraId)) return false;
      found.set(paraId, index);
    }
    if (found.size !== targets.size) return false;
    const indices = [...found.values()];
    for (const unit of engine.paragraphSegments(story, indices)) {
      let ascii = true;
      let existingText = false;
      for (const segment of unit) {
        if (segment.kind === 'text') {
          ascii &&= /^[\x20-\x7e]*$/.test(segment.text);
          existingText ||= segment.text.length > 0 && segment.attributes.ins == null;
        } else if (segment.kind !== 'pilcrow') {
          ascii = false;
        }
      }
      if (!ascii || !existingText) return false;
    }
  }
  return true;
}

/**
 * Drops the document. `keepSurfaces` keeps the attached page canvases, still
 * showing the old pages, for a document that replaces it page for page.
 */
function destroySession(keepSurfaces = false): void {
  collaboration?.dispose(); collaboration = null;
  supersedeBackgroundPageBuild();
  fontRequirements = null;
  forgetRequirementsCache();
  requestedRequirements = null;
  unsubscribe?.();
  unsubscribe = null;
  proposals?.destroy();
  layoutReady = false;
  proposals = null;
  releaseCommandHost?.();
  releaseCommandHost = null;
  session?.destroy();
  lastProposalMirrorVersion = null;
  session = null;
  commandHost = null;
  commandInit = null;
  commandLayout = null;
  sessionFonts?.dispose();
  sessionFonts = null;
  commandFrameChanges = null;
  positionIndex = null;
  commandDisplayWindow = undefined;
  openedDocument = null;
  openedSource = null;
  openedViaSource = false;
  openedParsed = false;
  sourcePreviewing = false;
  editorSaves = { full: false };
  openedVersion = undefined;
  previewing = false;
  cutPartial = false;
  openedCut = null;
  previewFinalPages = null;
  clearProvisionalFinalPages();
  pendingUpdates = [];
  layoutRevision = 0;
  retainedRegions = false;
  retainedLayoutVersion = null;
  sentHeadersFootersEpoch = 0;
  fontsRevision = -1;
  supersedeSlicedCompletion();
  incompleteLayout = null;
  completedLayout = null;
  retainedFrame = null;
  glyphCache = null;
  offscreenBackBuffers.clear();
  pendingOffscreenPageIds.clear();
  if (!keepSurfaces) {
    offscreenCanvases.clear();
    activeOffscreenPageIds.clear();
  }
  caretPaintRect = null;
  paintedCaretPageId = null;
  paintedCaretKey = null;
  caretStage = null;
  intactBackBuffers.clear();
}

function forgetOffscreenPagePixels(pageId: string): void {
  intactBackBuffers.delete(pageId);
  if (paintedCaretPageId === pageId) {
    paintedCaretPageId = null;
    paintedCaretKey = null;
  }
}

function retainedPageByKey(frame: RetainedFrame, pageId: string) {
  return /^\d+$/.test(pageId) ? retainedFramePageById(frame, BigInt(pageId)) : undefined;
}

function setCommandRasterScale(scale: number, indices?: ReadonlySet<number>): void {
  if (!Number.isFinite(scale) || scale <= 0) throw new Error('Invalid pixel ratio');
  const previous = offscreenDpr * offscreenZoom;
  if (previous === scale) return;
  const pages = retainedFrame?.pages.filter(({ page, pageIndex }) => !page.unbuilt &&
    (indices === undefined || indices.has(pageIndex))) ?? [];
  const resized = pages.some(({ page }) => Math.ceil(page.width * previous) !== Math.ceil(page.width * scale) ||
    Math.ceil(page.height * previous) !== Math.ceil(page.height * scale));
  offscreenDpr = scale;
  offscreenZoom = 1;
  offscreenBackBuffers.clear();
  intactBackBuffers.clear();
  caretStage = null;
  paintedCaretPageId = null;
  paintedCaretKey = null;
  if (resized) for (const { pageId } of pages) pendingOffscreenPageIds.add(pageId.toString());
}

function applyWorkerFrame(bytes: Uint8Array, previewVersion?: number) {
  const delta = decodeFrameDelta(bytes);
  const published = commandFrameChanges?.reviewFrame;
  const inherited = published?.frameEpoch === retainedFrame?.frameEpoch && delta.baseFrameEpoch === retainedFrame?.frameEpoch
    ? published?.previewVersion : commandFrameChanges?.layoutPreviewVersion;
  retainedFrame = applyFrameDeltaOwned(retainedFrame, delta);
  if (commandHost && commandFrameChanges && (previewVersion ?? inherited) !== undefined) {
    commandFrameChanges.reviewFrame = { frameEpoch: retainedFrame.frameEpoch, previewVersion: (previewVersion ?? inherited)! };
  }
  for (const pageId of retainedFrame.damagedPageIds) pendingOffscreenPageIds.add(pageId.toString());
  return delta;
}

async function finalizeResidentFrame(
  bytes: Uint8Array,
  requireCaret = false,
  paintCaret = false,
  stageCommand = false,
  previewVersion?: number
) {
  let delta = applyWorkerFrame(bytes, previewVersion);
  const limit = provisionalFinalPages;
  if (session && limit !== null && retainedFrame) {
    const pages = retainedFrame.displayList.pages
      .filter((page) => page.pageIndex >= limit && !page.unbuilt)
      .map((page) => page.pageIndex);
    if (pages.length > 0) {
      const released = session.releaseDisplayPagesFrame(pages, retainedFrame.frameEpoch);
      if (released === null) throw new Error('Provisional display pages could not be released');
      bytes = session.buildDisplayPagesFrame([], 0);
      delta = applyWorkerFrame(bytes, previewVersion);
    }
  }
  const caret = session?.residentCaretSnapshot();
  if (!caret || !retainedFrame || !residentCaretSnapshotForFrame(caret, retainedFrame)) {
    throw new Error('Resident caret snapshot does not match the produced frame');
  }
  if (requireCaret && !caret.caretRect) {
    throw new Error('Resident input frame omitted collapsed caret geometry');
  }
  for (const pageId of pendingOffscreenPageIds) {
    const page = retainedPageByKey(retainedFrame, pageId);
    if (!page || page.page.unbuilt === true) pendingOffscreenPageIds.delete(pageId);
  }
  if (stageCommand) return { bytes, delta, caret };
  caretPaintRect = paintCaret ? (caret.caretRect ?? null) : null;
  for (const pageId of new Set([...offscreenCanvases.keys(), ...offscreenBackBuffers.keys()])) {
    const page = retainedPageByKey(retainedFrame, pageId);
    const unbuilt = page ? page.page.unbuilt === true : undefined;
    if (unbuilt === false) continue;
    const canvas = offscreenCanvases.get(pageId);
    if (unbuilt === undefined) offscreenCanvases.delete(pageId);
    else if (canvas) releaseOffscreenPageCanvas(canvas);
    offscreenBackBuffers.delete(pageId);
    forgetOffscreenPagePixels(pageId);
  }
  await replayOffscreen(false);
  return { bytes, delta, caret };
}

/** Parses `source` into `opening` and makes it the worker's document; null when the worker declines a cut. */
function adoptResidentSource(
  opening: ResidentEngineSession,
  source: ResidentSourceOpenRequest
): { document: import('../types/document').Document; hostJson?: string } | null {
  let hostJson: string | undefined;
  let document: import('../types/document').Document;
  let partial = false;
  const options = { ...source, initialUpdate: source.collaboration?.initialUpdate };
  if (source.document) {
    opening.openDocument(source.document, options);
    document = source.document;
  } else if (source.cut) {
    const cut = opening.openDocxPreview(source.bytes!, source.cut.blocks, source.cut.paragraphBudget);
    if (cut === null) return null;
    hostJson = cut;
    const host = decodeDocxHostJson(hostJson, source.bytes!);
    document = host.document;
    partial = host.wholeBody !== true;
  } else {
    hostJson = opening.openSourceDocx(source.bytes!, options);
    document = decodeDocxHostJson(hostJson, source.bytes!).document;
  }
  if (!source.cut) opening.setDirectBatches(true);
  session = opening;
  openedDocument = { heapLimitBytes: source.heapLimitBytes };
  openedViaSource = true;
  previewing = source.cut !== undefined;
  cutPartial = partial;
  openedCut = source.cut ? source : null;
  openedParsed = source.document !== undefined;
  openedSource = source.bytes && hostJson && !source.cut ? { bytes: exactBuffer(source.bytes), hostJson } : null;
  editorSaves = { full: false };
  return { document, hostJson };
}

/** Restores a cut whose whole source failed to open, so the worker still holds the document its preview shows. */
async function reopenSourceCut(source: ResidentSourceOpenRequest): Promise<boolean> {
  const opening = await createResidentEngineSession(source.heapLimitBytes);
  try {
    if (adoptResidentSource(opening, source)) return true;
    opening.destroy();
  } catch (error) {
    if (error instanceof WebAssembly.RuntimeError) throw error;
    if (session === opening) destroySession(true);
    else opening.destroy();
  }
  return false;
}

async function replySourcePreview(
  id: number,
  opening: ResidentEngineSession,
  preview: ResidentSourcePreviewOptions,
  defaultLayoutInput: () => string,
  started: number,
  sourceOpened?: { sourceDocument: import('../types/document').Document; hostJson?: string; clientId: number; clientClock: number }
): Promise<void> {
  sourcePreviewing = true;
  const pages = preview.pages ?? 1;
  const init: ResidentCommandInitRequest = {
    session: '', incarnation: 0, epoch: 0, readOnly: true,
    fonts: preview.fonts ?? [], fontRegistrationSlots: preview.fontRegistrationSlots,
    layoutInput: preview.layoutInput ?? defaultLayoutInput(),
    layoutWithRegions: true, layoutRevision: 1, extras: preview.extras ?? '{}',
    displayWindow: preview.displayWindow ?? [0, pages],
  };
  setFrameDisplayWindow(opening, init.displayWindow, false);
  const layout = hydrate({
    workerAuthoritative: true, clientId: 0, state: new Uint8Array(0), selection: null,
    fonts: init.fonts, fontsRevision: 0, renderInputs: [], measureInputs: [],
    layoutInput: init.layoutInput, layoutWithRegions: true, layoutRevision: 1, partialDocument: cutPartial,
  }, pages, true, false, undefined, undefined, init.fontRegistrationSlots ? init : undefined);
  if (layout.provisional) incompleteLayout = {
    layoutInput: init.layoutInput, workerAuthoritative: true, extras: init.extras,
  };
  setFrameDisplayWindow(opening, init.displayWindow, false);
  const extras = JSON.stringify({ ...JSON.parse(init.extras), headersFooters: JSON.parse(opening.retainedHeadersFootersJson() ?? 'null') });
  const bytes = opening.buildDisplayListFrame(extras, 0);
  await replyFrame(id, bytes, performance.now() - started, [], undefined, started,
    false, false, layout.layoutJson ?? undefined, layout.provisional, undefined, [], null, sourceOpened);
}

async function replyFrame(
  id: number,
  bytes: Uint8Array,
  engineMs: number,
  updates = pendingUpdates,
  engineProfile?: import('./index').YrsEngineApplyProfile,
  requestStarted = performance.now(),
  requireCaret = false,
  paintCaret = false,
  layoutReply?: string | LayoutMetaV1,
  layoutProvisional = false,
  deletedUnits?: number,
  precedingPageFrames: Uint8Array[] = [],
  selectionCursor: CollaborationCursor | null = null,
  sourceOpened?: { sourceDocument: import('../types/document').Document; hostJson?: string; clientId: number; clientClock: number }
): Promise<void> {
  const layoutJson = typeof layoutReply === 'string' ? layoutReply : undefined;
  const layoutMeta = typeof layoutReply === 'object' ? layoutReply : undefined;
  const documentVersion = session?.proposalEngine.version();
  if (layoutReply !== undefined) retainedLayoutVersion = documentVersion ?? null;
  const documentPreview = previewing || sourcePreviewing;
  if (!previewing && openedVersion === undefined && documentVersion !== undefined) {
    openedVersion = documentVersion;
  }
  const documentAsOpened =
    !previewing && documentVersion !== undefined && documentVersion === openedVersion;
  applyWorkerFrame(bytes);
  const limit = provisionalFinalPages;
  if (session && limit !== null && retainedFrame) {
    const pages = retainedFrame.displayList.pages
      .filter((page) => page.pageIndex >= limit && !page.unbuilt)
      .map((page) => page.pageIndex);
    if (pages.length > 0) {
      const released = session.releaseDisplayPagesFrame(pages, retainedFrame.frameEpoch);
      if (released === null) throw new Error('Provisional display pages could not be released');
      bytes = session.buildDisplayPagesFrame([], 0);
      applyWorkerFrame(bytes);
    }
  }
  const caret = session?.residentCaretSnapshot();
  if (!caret || !retainedFrame || !residentCaretSnapshotForFrame(caret, retainedFrame)) {
    throw new Error('Resident caret snapshot does not match the produced frame');
  }
  if (requireCaret && !caret.caretRect) {
    throw new Error('Resident input frame omitted collapsed caret geometry');
  }
  const selection = session?.selection() ?? null;
  caretPaintRect = paintCaret ? (caret.caretRect ?? null) : null;
  // Pages no longer in the document drop their surfaces (their elements
  // unmounted main-side). An unbuilt page keeps its transferred canvas, which
  // can never be transferred again, and only loses its pixels.
  for (const pageId of pendingOffscreenPageIds) {
    const page = retainedPageByKey(retainedFrame, pageId);
    if (!page || page.page.unbuilt === true) pendingOffscreenPageIds.delete(pageId);
  }
  for (const pageId of new Set([...offscreenCanvases.keys(), ...offscreenBackBuffers.keys()])) {
    const page = retainedPageByKey(retainedFrame, pageId);
    const unbuilt = page ? page.page.unbuilt === true : undefined;
    if (unbuilt === false) continue;
    const canvas = offscreenCanvases.get(pageId);
    if (unbuilt === undefined) offscreenCanvases.delete(pageId);
    else if (canvas) releaseOffscreenPageCanvas(canvas);
    offscreenBackBuffers.delete(pageId);
    forgetOffscreenPagePixels(pageId);
  }
  const replayStarted = performance.now();
  const { replayedPages, caretPainted } = await replayOffscreen(false);
  const replayMs = performance.now() - replayStarted;
  const packed = packResidentFrame(bytes, precedingPageFrames);
  const { frame } = packed.value;
  const pageFrames = packed.value.pageFrames ?? [];
  const updateBuffers = updates.map(exactBuffer);
  const stateVector = session ? exactBuffer(session.encodeStateVector()) : undefined;
  if (layoutMeta) sentHeadersFootersEpoch = layoutMeta.headersFootersEpoch;
  reply(
    {
      id,
      ok: true,
      ...sourceOpened,
      frame,
      ...(pageFrames.length > 0 ? { pageFrames: [...pageFrames, frame] } : {}),
      updates: updateBuffers,
      engineMs,
      workerTotalMs: performance.now() - requestStarted,
      engineProfile,
      caret,
      selection,
      ...(selectionCursor ? { selectionCursor } : {}),
      caretPainted,
      replayMs,
      replayedPages,
      layoutRevision,
      ...(documentVersion === undefined ? {} : { documentVersion }),
      ...(documentPreview ? { documentPreview: true } : {}),
      ...(documentAsOpened ? { documentAsOpened: true } : {}),
      ...(deletedUnits === undefined ? {} : { deletedUnits }),
      ...(stateVector ? { stateVector } : {}),
      ...(layoutJson !== undefined ? { layoutJson } : {}),
      ...(layoutMeta ? { layoutMeta } : {}),
      ...(layoutProvisional ? { layoutProvisional } : {}),
    },
    [...packed.transfer, ...updateBuffers, ...(stateVector ? [stateVector] : []),
      ...(layoutMeta ? [layoutMeta.pageSizes.buffer as ArrayBuffer] : [])]
  );
}

async function replayOffscreen(
  force: boolean | Set<string>
): Promise<{ replayedPages: number; caretPainted: boolean }> {
  const forcedPageIds = force === true ? activeOffscreenPageIds : force;
  if (forcedPageIds) {
    for (const { pageId, page } of retainedFrame?.pages ?? []) {
      const key = pageId.toString();
      if (!page.unbuilt && forcedPageIds.has(key)) pendingOffscreenPageIds.add(key);
    }
  }
  if (!retainedFrame || offscreenCanvases.size === 0) {
    return { replayedPages: 0, caretPainted: false };
  }
  if (!glyphCache && session) {
    glyphCache = new GlyphCache({
      provider: (fontId, glyphId) => {
        try {
          return session!.outlineGlyphJson(fontId, glyphId);
        } catch (error) {
          // The raster paints on with browser text, so running out of memory is answered here.
          if (error instanceof WebAssembly.RuntimeError && editFailedAllocationBytes() > 0) {
            trapped(handlingId, error);
          }
          throw error;
        }
      },
    });
  }
  const caretTarget =
    caretPaintRect && activeOffscreenPageIds.has(caretPaintRect.pageId) ? caretPaintRect : null;
  const caretDevice = caretTarget
    ? residentCaretDeviceRect(caretTarget, caretStyle, offscreenDpr, offscreenZoom)
    : null;
  const caretKey =
    caretTarget && caretDevice
      ? `${caretTarget.pageId}|${caretDevice.x}|${caretDevice.y}|${caretDevice.width}|${caretDevice.height}|${caretStyle.color}`
      : null;
  const preparations: Array<
    Promise<{
      canvas: OffscreenCanvas;
      buffer: OffscreenCanvas;
      pageId: string;
    }>
  > = [];
  const images = openedViaSource && session ? residentMediaImages(session) : undefined;
  for (let index = 0; index < retainedFrame.pages.length; index += 1) {
    const retainedPage = retainedFrame.pages[index];
    const pageIdString = retainedPage.pageId.toString();
    // Off-window pages hold no pixels; they re-raster through the forced set
    // when they re-enter the window.
    if (!activeOffscreenPageIds.has(pageIdString)) continue;
    const damaged = pendingOffscreenPageIds.has(pageIdString);
    // Beyond damage, a page presents only for caret compositing: the page
    // gaining the painted line and the page losing it.
    const gainsCaret =
      pageIdString === caretTarget?.pageId &&
      !(paintedCaretPageId === pageIdString && paintedCaretKey === caretKey);
    const losesCaret =
      pageIdString === paintedCaretPageId && pageIdString !== caretTarget?.pageId;
    if (!damaged && !gainsCaret && !losesCaret) continue;
    const canvas = offscreenCanvases.get(pageIdString);
    const page = retainedFrame.displayList.pages[index];
    if (!canvas || !page || page.unbuilt) continue;
    const pageId = pageIdString;
    let buffer = offscreenBackBuffers.get(pageId);
    if (!buffer) {
      buffer = new OffscreenCanvas(1, 1);
      offscreenBackBuffers.set(pageId, buffer);
    }
    const resolvedBuffer = buffer;
    if (!damaged && intactBackBuffers.has(pageId)) {
      preparations.push(Promise.resolve({ canvas, buffer: resolvedBuffer, pageId }));
      continue;
    }
    preparations.push(
      rasterizeDisplayPageToBackBuffer(
        resolvedBuffer,
        page,
        { glyphCache: glyphCache ?? undefined, ...(images ? { resolveImage: images.resolve } : {}) },
        offscreenDpr,
        offscreenZoom
      ).then(() => ({ canvas, buffer: resolvedBuffer, pageId }))
    );
  }
  const prepared = await Promise.all(preparations).catch(async (error) => {
    await Promise.allSettled(preparations);
    throw error;
  }).finally(() => images?.close());
  let caretPainted =
    caretTarget !== null && paintedCaretPageId === caretTarget.pageId && paintedCaretKey === caretKey;
  // Present only after the entire damaged frame is ready. This loop is
  // synchronous, so the compositor can never observe the renderer's clears.
  for (const { canvas, buffer, pageId } of prepared) {
    if (caretTarget && caretDevice && pageId === caretTarget.pageId) {
      if (!caretStage) caretStage = new OffscreenCanvas(1, 1);
      presentOffscreenPageBackBufferWithCaret(canvas, buffer, caretStage, {
        ...caretDevice,
        color: caretStyle.color,
      });
      intactBackBuffers.add(pageId);
      paintedCaretPageId = pageId;
      paintedCaretKey = caretKey;
      caretPainted = true;
    } else {
      presentOffscreenPageBackBuffer(canvas, buffer);
      intactBackBuffers.delete(pageId);
      if (paintedCaretPageId === pageId) {
        paintedCaretPageId = null;
        paintedCaretKey = null;
      }
    }
    pendingOffscreenPageIds.delete(pageId);
  }
  return { replayedPages: prepared.length, caretPainted };
}

function exactBuffer(bytes: Uint8Array): ArrayBuffer {
  return exactFrameBuffer(bytes);
}

function editFailedAllocationBytes(): number {
  return wasmModuleMemories().find((module) => module.label === 'docx-edit')?.failedAllocationBytes ?? 0;
}

function reply(response: ResidentEngineWorkerResponse, transfer: Transferable[] = []): void {
  // A trap the raster painted past fails the request that would succeed, and
  // every request waiting on it, through the paths that answer failures.
  if (trap && response.ok) throw trap;
  const collabUpdates = collaboration?.drain();
  scope.postMessage({ ...response, ...(collabUpdates ? { collabUpdates: { ...collabUpdates,
    ...(commandInit ? { session: commandInit.session, incarnation: commandInit.incarnation, epoch: commandInit.epoch } : {}) } } : {}),
    memory: wasmModuleMemories() }, [...transfer, ...(collabUpdates?.updates ?? [])]);
}
