import { readRetainedLayoutMeta } from './layoutMeta';
import type {
  YrsChangeTarget,
  YrsCommentInfo,
  YrsCommentReceipt,
  YrsDocxHost,
  YrsEngineApplyProfile,
  YrsLoc,
  YrsOpeningOptions,
  YrsParagraph,
  YrsParagraphLength,
  YrsParagraphSpan,
  YrsRegionLayoutProgress,
  YrsRenderEnv,
  YrsReplaceReceipt,
  YrsResidentCaretSnapshot,
  YrsResidentFontRegistration,
  YrsResolveReceipt,
  YrsResolvedCommentAnchor,
  YrsRevisionInfo,
  YrsRevisionReceipt,
  YrsSelection,
  YrsSelectionContext,
  YrsSelectionText,
  YrsSession,
  YrsSplitReceipt,
  YrsStoryRange,
  YrsStorySegment,
  YrsTableRange,
  YrsTableReceipt,
  YrsTargetEditResult,
  YrsTextMatch,
  YrsUndoCaptureMode,
} from './index';
import type { EditSession } from './wasm/index';
import { resolveCommentMedia } from './hostMedia';
import { registerSessionInternals } from './sessionInternals';
import { editorSaveKeys } from './editorSaveKeys';
import { noteYrsStoriesDirty } from './yrsToDocument';
import type {
  DocxParagraphAnchorResult,
  DocxParagraphIdentityReceipt,
  DocxParagraphIdentitySnapshot,
  DocxParagraphSavePlan,
} from './paragraphIdentity';
import { decodeS9Envelope, decodeS9EnvelopeValue } from '../docx/rustParseFacade';
import { decodeEncodedSelection } from './encodedSelection';
import type { CollaborationTextInsertion, CollaborationUpdateOrigin } from '../collaboration/types';
import type {
  DocxEditResult,
  DocxFindTextResult,
  DocxReadParagraphsResult,
  DocxTextTarget,
  DocxValidationResult,
} from './edits';
import type { DocxParagraphHeading } from './readTypes';
import { createProposalRegistry } from './proposals';
import type { DocxContentControlsResult } from './contentControls';
import type {
  DocxExportFailure,
  DocxExportResult,
  DocxMarkdownContent,
  DocxStructuredContent,
} from './structuredExport';
import type {
  DocxLayoutMap,
  DocxPagedStructuredContent,
  DocxSnapshotLayoutMap,
} from './pagedExport';

/** SHA-256 digests of the byte copies {@link prepareDocxBytes} made, by copy. */
/** @internal */
export const preparedDigests = new WeakMap<Uint8Array, string>();

/** @internal */
export function randomClientId(): number {
  if (typeof crypto !== 'undefined' && typeof crypto.getRandomValues === 'function') {
    const buffer = new Uint32Array(1);
    crypto.getRandomValues(buffer);
    return buffer[0];
  }
  return Math.floor(Math.random() * 0xffffffff);
}

function wireChangeTarget(target: YrsChangeTarget): string {
  return JSON.stringify(
    'revisionId' in target
      ? { revisionId: target.revisionId }
      : {
          story: target.story,
          startPara: target.start.paraId,
          startOffset: target.start.offset,
          endPara: target.end.paraId,
          endOffset: target.end.offset,
        }
  );
}

function wireRanges(ranges: readonly YrsStoryRange[]): string {
  return JSON.stringify(
    ranges.map((range) => ({
      story: range.story,
      startPara: range.start.paraId,
      startOffset: range.start.offset,
      endPara: range.end.paraId,
      endOffset: range.end.offset,
    }))
  );
}

function targetStory(target: DocxTextTarget): string {
  return target.kind === 'search' ? target.within.story : target.story;
}

function docxSourceBuffer(bytes: Uint8Array): ArrayBuffer {
  if (
    bytes.buffer instanceof ArrayBuffer &&
    bytes.byteOffset === 0 &&
    bytes.byteLength === bytes.buffer.byteLength
  ) {
    return bytes.buffer;
  }
  return new Uint8Array(bytes).buffer as ArrayBuffer;
}

/**
 * Decodes the host metadata a resident worker's `open` replied with, for the
 * package `source` it opened. @internal
 */
export function decodeDocxHostJson(json: string, source: Uint8Array): YrsDocxHost {
  return decodeDocxHost(json, source);
}

function decodeDocxHost(json: string, source: Uint8Array): YrsDocxHost {
  const value: unknown = JSON.parse(json);
  if (typeof value !== 'object' || value === null || Array.isArray(value)) {
    throw new TypeError('DOCX host metadata must be an object');
  }
  const wire = value as Record<string, unknown>;
  if (
    !Array.isArray(wire.referencedFonts) ||
    !wire.referencedFonts.every((name) => typeof name === 'string')
  ) {
    throw new TypeError('DOCX host referencedFonts must be a string array');
  }
  const unusedScriptFonts = wire.unusedScriptFonts ?? [];
  if (
    !Array.isArray(unusedScriptFonts) ||
    !unusedScriptFonts.every((name) => typeof name === 'string')
  ) {
    throw new TypeError('DOCX host unusedScriptFonts must be a string array');
  }
  const result = decodeS9EnvelopeValue(wire.envelope, docxSourceBuffer(source));
  return {
    document: result.document,
    referencedFonts: wire.referencedFonts,
    unusedScriptFonts,
    ...(wire.wholeBody === true ? { wholeBody: true as const } : {}),
    embeddedFonts: result.embeddedFonts,
    ...(result.fontTableRelationshipsXml === undefined
      ? {}
      : { fontTableRelationshipsXml: result.fontTableRelationshipsXml }),
  };
}

/** A lone UTF-16 surrogate, which crossing into Wasm would turn into U+FFFD. */
const LONE_SURROGATE = /[\uD800-\uDBFF](?![\uDC00-\uDFFF])|(?<![\uD800-\uDBFF])[\uDC00-\uDFFF]/;

/** @internal */
export function wrapSession(
  session: EditSession,
  clientId: number,
  opened?: { source: Uint8Array; host: YrsDocxHost }
): YrsSession {
  const listeners = new Map<
    number,
    (update: Uint8Array, origin: CollaborationUpdateOrigin) => void
  >();
  const pendingUpdates: Array<{
    update: Uint8Array;
    origin: CollaborationUpdateOrigin;
  }> = [];
  let observing = false;
  let destroyed = false;
  let nextListenerId = 0;
  let wasmCallDepth = 0;
  let flushingUpdates = false;
  let undoTracked = false;
  let cachedSelection: YrsSelection | null | undefined;
  let cachedSelectionContext: { key: string; json: string } | null = null;
  const residentFonts: YrsResidentFontRegistration[] = [];
  const residentRenderInputs = new Map<string, YrsRenderEnv>();
  const residentMeasureInputs = new Map<string, string>();
  let residentLayoutInput: string | null = null;
  let residentLayoutWithRegions = false;
  let residentLayoutRevision = 0;
  // The current resident layout ran only in a resident worker replica.
  let layoutRanInWorker = false;
  // The request of a region layout begun a step at a time.
  let steppedLayoutInput: string | null = null;
  const completedRegionLayout = (progress: YrsRegionLayoutProgress): YrsRegionLayoutProgress => {
    if (progress.layoutJson !== undefined && steppedLayoutInput !== null) {
      residentLayoutInput = steppedLayoutInput;
      residentLayoutWithRegions = true;
      residentLayoutRevision += 1;
      layoutRanInWorker = false;
      steppedLayoutInput = null;
    }
    return progress;
  };
  let residentFontsRevision = 0;
  let residentFontsClearRevision = 0;
  let docxSource: Uint8Array | null = opened?.source ?? null;
  let docxSourceKeys: ReturnType<typeof editorSaveKeys> | null = opened
    ? editorSaveKeys(opened.host.document)
    : null;

  const invalidateReadCaches = (): void => {
    cachedSelection = undefined;
    cachedSelectionContext = null;
  };

  const flushUpdates = (): void => {
    if (destroyed || flushingUpdates || wasmCallDepth !== 0) return;
    flushingUpdates = true;
    try {
      while (!destroyed && pendingUpdates.length > 0) {
        const event = pendingUpdates.shift();
        if (!event) break;
        for (const [id, listener] of [...listeners]) {
          if (destroyed) return;
          if (listeners.get(id) !== listener) continue;
          try {
            listener(event.update.slice(), event.origin);
          } catch {}
        }
      }
    } finally {
      flushingUpdates = false;
      if (destroyed) pendingUpdates.length = 0;
    }
  };

  // A preview session refuses every change to its document.
  let displayOnly = false;
  // A preview's cut of a package, whose layouts count only its own pages.
  let partialDocument = false;
  const mutate = <T>(operation: () => T): T => {
    if (displayOnly) throw new Error('A document preview is display-only');
    return mutateAlways(operation);
  };
  const mutateAlways = <T>(operation: () => T): T => {
    invalidateReadCaches();
    wasmCallDepth += 1;
    try {
      return operation();
    } finally {
      wasmCallDepth -= 1;
      if (wasmCallDepth === 0) flushUpdates();
    }
  };

  const cloneSelection = (value: YrsSelection | null): YrsSelection | null =>
    value
      ? {
          anchor: { ...value.anchor },
          head: { ...value.head },
        }
      : null;

  const ensureUndo = (targetStory?: string): void => {
    if (!undoTracked) {
      session.track_undo();
      undoTracked = true;
    }
    if (targetStory !== undefined) {
      session.select_story(targetStory);
      markDirty(targetStory);
    }
  };

  const markDirty = (stories: 'all' | string | Iterable<string>): void => {
    noteYrsStoriesDirty(facade, stories);
  };

  const markReceiptStories = (receipt: YrsTableReceipt): YrsTableReceipt => {
    markDirty(receipt.createdStoryIds);
    markDirty(receipt.deletedStoryIds);
    markDirty(receipt.changedStoryIds);
    return receipt;
  };

  const selectionStory = (): 'all' | string =>
    (
      (cachedSelection !== undefined
        ? cachedSelection
        : (JSON.parse(session.selection()) as YrsSelection | null))?.head.story ?? 'all'
    );

  const ensureObserver = () => {
    if (observing) return;
    session.set_update_observer((update: Uint8Array, origin: number) => {
      if (origin !== 0 && origin !== 1) return;
      pendingUpdates.push({
        update: update.slice(),
        origin: origin === 0 ? 'local' : 'remote',
      });
      flushUpdates();
    });
    observing = true;
  };

  const clearUnusedObserver = (): void => {
    if (!observing || listeners.size > 0 || destroyed) return;
    pendingUpdates.length = 0;
    session.clear_update_observer();
    observing = false;
  };

  // `data:` URLs of the opened package's `media:{n}` sources.
  const mediaDataUrls = new Map<string, string | null>();
  let mediaScope = 0;
  const resetMedia = (): void => {
    mediaDataUrls.clear();
    mediaScope += 1;
  };

  const openDocx = (
    bytes: Uint8Array,
    seedStories: boolean,
    options: YrsOpeningOptions = {}
  ): YrsDocxHost => {
    const source = bytes.slice();
    markDirty('all');
    resetMedia();
    const json = mutate(() => {
      session.set_media_tokens(options.mediaTokens === true);
      const opened = session.open_docx(
        source,
        seedStories,
        options.generation,
        preparedDigests.get(bytes)
      );
      proposals.reset();
      return opened;
    });
    const host = withHostMedia(decodeDocxHost(json, source));
    docxSource = source;
    docxSourceKeys = editorSaveKeys(host.document);
    partialDocument = false;
    return host;
  };

  const mediaDataUrl = (token: string): string | null => {
    if (destroyed || !token.startsWith('media:')) return null;
    let url = mediaDataUrls.get(token);
    if (url === undefined) {
      url = session.media_data_url(token) ?? null;
      mediaDataUrls.set(token, url);
    }
    return url;
  };

  const withHostMedia = (host: YrsDocxHost): YrsDocxHost => {
    resolveCommentMedia(host.document.package.document.comments, mediaDataUrl);
    return host;
  };

  let workerDocumentVersion: string | null = null;
  const proposals = createProposalRegistry({
    version: () => session.version(),
    resolveParagraphAnchor: (anchor) => facade.resolveParagraphAnchor(anchor),
    findText: (request) => facade.findText(request),
    readParagraphs: (request) => facade.readParagraphs(request),
    applyEdits: (request) => facade.applyEdits(request),
    listRevisions: () => facade.listRevisions(),
    revisionStamps: (ids) => JSON.parse(session.revision_stamps_json(JSON.stringify(ids))),
    settleRevisions: (accept, reject) => {
      const since = facade.storiesChangedSince(Number.MAX_SAFE_INTEGER).revision;
      session.settle_revisions_json(JSON.stringify({ accept, reject }));
      markDirty(facade.storiesChangedSince(since).stories);
    },
    ...(typeof session.begin_shared_reads === 'function' &&
    typeof session.end_shared_reads === 'function'
      ? {
          sharedReads: <R>(read: () => R): R => {
            session.begin_shared_reads();
            try {
              return read();
            } finally {
              session.end_shared_reads();
            }
          },
        }
      : {}),
  });

  const facade: YrsSession = {
    clientId,
    openDocxPreview: (bytes, blocks, paragraphBudget) => {
      markDirty('all');
      resetMedia();
      const json = mutateAlways(() =>
        paragraphBudget === undefined
          ? session.open_docx_preview(bytes, blocks)
          : session.open_docx_preview_with_budget(bytes, blocks, paragraphBudget)
      );
      if (json === undefined) return null;
      displayOnly = true;
      partialDocument = true;
      return withHostMedia(decodeDocxHost(json, bytes));
    },
    isDisplayOnly: () => displayOnly,
    markDisplayOnly: () => {
      displayOnly = true;
      partialDocument = true;
    },
    setPartialDocument: (partial) => {
      partialDocument = partial;
      session.set_partial_document(partial);
    },
    layoutDocumentWithRegionsPrefixRetainedJson: (input, pages) => {
      const output = session.layout_document_with_regions_prefix_retained_json(input, pages);
      residentLayoutInput = input;
      residentLayoutWithRegions = true;
      residentLayoutRevision += 1;
      layoutRanInWorker = false;
      return output;
    },

    registerFont: (bytes) => {
      const id = session.register_measure_font(bytes);
      residentFonts.push(bytes.slice());
      residentFontsRevision += 1;
      return id;
    },
    registerSubstituteFont: (base, family) => {
      const id = session.register_substitute_measure_font(base, family);
      residentFonts.push({ substituteOf: base, family });
      residentFontsRevision += 1;
      return id;
    },
    clearFonts: () => {
      session.clear_measure_fonts();
      residentFonts.length = 0;
      residentMeasureInputs.clear();
      residentFontsRevision += 1;
      residentFontsClearRevision = residentFontsRevision;
    },
    measureParagraphJson: (input) => {
      const output = session.measure_paragraph_json(input);
      residentMeasureInputs.set(input, input);
      return output;
    },
    layoutDocumentJson: (input) => {
      const output = session.layout_document_json(input);
      residentLayoutInput = input;
      residentLayoutWithRegions = false;
      residentLayoutRevision += 1;
      layoutRanInWorker = false;
      return output;
    },
    layoutFontRequirementsJson: (input) => session.layout_font_requirements_json(input),
    layoutDocumentWithRegionsJson: (input) => {
      const output = session.layout_document_with_regions_json(input);
      residentLayoutInput = input;
      residentLayoutWithRegions = true;
      residentLayoutRevision += 1;
      layoutRanInWorker = false;
      return output;
    },
    layoutDocumentWithRegionsRetainedJson: (input) => {
      const output = session.layout_document_with_regions_retained_json(input);
      residentLayoutInput = input;
      residentLayoutWithRegions = true;
      residentLayoutRevision += 1;
      layoutRanInWorker = false;
      return output;
    },
    layoutDocumentWithRegionsRetainedMeta: (input) => {
      const output = readRetainedLayoutMeta(session.layout_document_with_regions_retained_meta(input));
      residentLayoutInput = input;
      residentLayoutWithRegions = true;
      residentLayoutRevision += 1;
      layoutRanInWorker = false;
      return output;
    },
    retainedLayoutJson: () => session.retained_layout_json(),
    beginRegionLayout: (input) => {
      steppedLayoutInput = input;
      return completedRegionLayout(
        JSON.parse(session.begin_region_layout(input)) as YrsRegionLayoutProgress
      );
    },
    resumeRegionLayout: (blocks) =>
      completedRegionLayout(
        JSON.parse(session.resume_region_layout(blocks)) as YrsRegionLayoutProgress
      ),
    adoptResidentWorkerLayout: (input) => {
      residentLayoutInput = input;
      residentLayoutWithRegions = true;
      residentLayoutRevision += 1;
      layoutRanInWorker = true;
      return residentLayoutRevision;
    },
    residentLayoutInWorker: () => layoutRanInWorker,
    retainedKernelInputsJson: (expectedLayoutRevision) => {
      if (layoutRanInWorker) {
        throw new Error('the retained layout was computed in the resident worker');
      }
      if (expectedLayoutRevision !== residentLayoutRevision) {
        throw new Error(
          `retained layout revision mismatch: expected ${expectedLayoutRevision}, current ${residentLayoutRevision}`
        );
      }
      return session.retained_kernel_inputs_json();
    },
    buildDisplayListJson: (input) => session.build_display_list_json(input),
    resetFrameBase: () => {
      if (!destroyed) session.reset_frame_base();
    },
    buildDisplayListFrame: (input, expectedFrameEpoch) =>
      session.build_display_list_frame(input, expectedFrameEpoch),
    setDisplayWindow: (start, end) => session.set_display_window(start, end),
    setDisplayRetainBuiltPages: (retain) => session.set_display_retain_built_pages(retain),
    setWindowedIncrementalBuilds: (enabled) => session.set_windowed_incremental_builds(enabled),
    buildDisplayPagesFrame: (pages, expectedFrameEpoch) =>
      session.build_display_pages_frame(Uint32Array.from(pages), expectedFrameEpoch),
    releaseDisplayPagesFrame: (pages, expectedFrameEpoch) => {
      const frame = session.release_display_pages_frame(Uint32Array.from(pages), expectedFrameEpoch);
      return frame.length === 0 ? null : frame;
    },
    residentCaretSnapshot: () =>
      JSON.parse(session.resident_caret_snapshot_json()) as YrsResidentCaretSnapshot,
    applyInput: (text, expectedFrameEpoch) => {
      ensureUndo();
      markDirty(selectionStory());
      return mutate(() => session.apply_input(text, expectedFrameEpoch));
    },
    applyDelete: (direction, expectedFrameEpoch, count = 1) => {
      ensureUndo();
      markDirty(selectionStory());
      return mutate(() => session.apply_delete(direction, expectedFrameEpoch, count));
    },
    residentDeletedUnits: () => session.resident_deleted_units(),
    applyInputProfiled: (text, expectedFrameEpoch) => {
      ensureUndo();
      markDirty(selectionStory());
      const frame = mutate(() => session.apply_input_profiled(text, expectedFrameEpoch));
      const profile = JSON.parse(session.apply_input_profile_json()) as YrsEngineApplyProfile;
      return { frame, profile };
    },
    applyDeleteProfiled: (direction, expectedFrameEpoch, count = 1) => {
      ensureUndo();
      markDirty(selectionStory());
      const frame = mutate(() =>
        session.apply_delete_profiled(direction, expectedFrameEpoch, count)
      );
      const profile = JSON.parse(session.apply_input_profile_json()) as YrsEngineApplyProfile;
      return { frame, profile };
    },
    residentWorkerSnapshot: (options) => {
      if (!residentLayoutInput) return null;
      if (!residentLayoutWithRegions && residentRenderInputs.size === 0) return null;
      const mirrored = workerDocumentVersion !== null;
      const selectionJson = mirrored ? 'null' : session.selection();
      const mediaSources = mirrored ? undefined : session.media_sources_json();
      const noteSeparators = mirrored ? undefined : session.note_separators_state();
      const knownFontsRevision = options?.knownFontsRevision;
      const fontsBaseRevision =
        knownFontsRevision != null &&
        Number.isInteger(knownFontsRevision) &&
        knownFontsRevision >= residentFontsClearRevision &&
        knownFontsRevision <= residentFontsRevision
          ? knownFontsRevision
          : undefined;
      const fonts =
        fontsBaseRevision === undefined
          ? residentFonts
          : residentFonts.slice(residentFonts.length - (residentFontsRevision - fontsBaseRevision));
      let state: Uint8Array | null = null;
      if (!mirrored && options?.knownStateVector) {
        try {
          state = session.encode_diff(options.knownStateVector.slice());
        } catch {
          state = null;
        }
      }
      return {
        clientId,
        ...(mirrored ? { workerAuthoritative: true as const } : {}),
        state: mirrored ? new Uint8Array(0) : (state ?? session.encode_state()),
        selection: JSON.parse(selectionJson) as YrsSelection | null,
        fonts: fonts.map((font) => (font instanceof Uint8Array ? font.slice() : { ...font })),
        fontsRevision: residentFontsRevision,
        ...(fontsBaseRevision === undefined ? {} : { fontsBaseRevision }),
        renderInputs: [...residentRenderInputs].map(([story, env]) => ({
          story,
          env: structuredClone(env),
        })),
        measureInputs: [...residentMeasureInputs.values()],
        layoutInput: residentLayoutInput,
        layoutWithRegions: residentLayoutWithRegions,
        layoutRevision: residentLayoutRevision,
        ...(partialDocument ? { partialDocument: true } : {}),
        ...(mediaSources ? { mediaSources } : {}),
        ...(noteSeparators?.length ? { noteSeparators } : {}),
      };
    },
    residentWorkerProbe: () => {
      if (!residentLayoutInput) return null;
      if (!residentLayoutWithRegions && residentRenderInputs.size === 0) return null;
      return { layoutRevision: residentLayoutRevision };
    },
    displayHitTestRegionsJson: (pageIndex, x, y) =>
      session.display_hit_test_regions_json(pageIndex, x, y),
    displayVerticalMoveJson: (position, direction, goalX) =>
      session.display_vertical_move_json(position, direction, goalX),
    displayRangeRectsJson: (from, to) => session.display_range_rects_json(from, to),
    displayRangeRectsRegionJson: (region, rId, from, to) =>
      session.display_range_rects_region_json(region, rId, from, to),
    outlineGlyphJson: (fontId, glyphId) => session.outline_glyph_json(fontId, glyphId),

    loadState: (update) => {
      markDirty('all');
      mutate(() => {
        session.load(update);
        proposals.reset();
      });
    },
    seedFromDocx: (bytes, options) => openDocx(bytes, true, options),
    openDocx,
    beginOpening: (generation) => {
      mutate(() => {
        session.begin_opening(generation);
        proposals.reset();
      });
    },
    seedOpaqueSequences: (names) => {
      markDirty('all');
      mutate(() => session.seed_opaque_sequences(JSON.stringify(names)));
    },
    mediaSource: (token) => {
      if (destroyed || !token.startsWith('media:')) return null;
      const bytes = session.media_bytes(token);
      const mimeType = bytes && session.media_type(token);
      return bytes && mimeType ? { bytes, mimeType } : null;
    },
    loadMediaSources: (json) => session.load_media_sources(json),
    loadNoteSeparators: (state) => session.load_note_separators(state),
    mediaDataUrl,
    mediaScope: () => mediaScope,
    materializeDocx: () => {
      const source = docxSource;
      const json = session.materialize_docx();
      if (!source || json === undefined) return null;
      return decodeS9Envelope(json, docxSourceBuffer(source)).document;
    },
    loadStories: (stories) => {
      markDirty(stories.map((seed) => seed.storyId));
      return mutate(
        () => JSON.parse(session.load_json(JSON.stringify(stories))) as Record<string, string[]>
      );
    },
    encodeState: () => session.encode_state(),
    encodeStateVector: () => session.encode_state_vector(),
    encodeStateAsUpdate: (remoteStateVector) =>
      remoteStateVector === undefined
        ? session.encode_state()
        : session.encode_diff(remoteStateVector.slice()),
    applyUpdate: (update) => {
      markDirty('all');
      return mutate(
        () =>
          JSON.parse(
            session.apply_update_with_inference(update)
          ) as CollaborationTextInsertion | null
      );
    },
    applyLocalUpdate: (update) => {
      ensureUndo();
      markDirty('all');
      mutate(() => session.apply_local_update(update));
    },
    applyHostUpdate: (update, stories) => {
      markDirty(stories ?? 'all');
      mutate(() => session.apply_host_update(update));
    },
    onUpdate: (listener) => {
      if (destroyed) throw new Error('yrs session is destroyed');
      if (typeof listener !== 'function') throw new TypeError('update listener must be a function');
      const id = nextListenerId++;
      listeners.set(id, listener);
      ensureObserver();
      let subscribed = true;
      return () => {
        if (!subscribed) return;
        subscribed = false;
        listeners.delete(id);
        clearUnusedObserver();
      };
    },

    encodeStickyPosition: (loc) => ({
      story: loc.story,
      encoded: session.encode_sticky_position(loc.story, loc.paraId, loc.offset),
    }),
    resolveStickyPosition: (position) => {
      try {
        return JSON.parse(
          session.resolve_sticky_position(position.story, position.encoded)
        ) as YrsLoc;
      } catch {
        return null;
      }
    },
    setSelection: (anchor, head = anchor) => {
      if (anchor.story !== head.story) throw new Error('yrs selection must stay inside one story');
      session.set_selection(anchor.story, anchor.paraId, anchor.offset, head.paraId, head.offset);
      cachedSelection = {
        anchor: { ...anchor },
        head: { ...head },
      };
    },
    selection: () => {
      if (cachedSelection !== undefined) return cloneSelection(cachedSelection);
      cachedSelection = JSON.parse(session.selection()) as YrsSelection | null;
      return cloneSelection(cachedSelection);
    },
    encodeSelection: () => decodeEncodedSelection(session.encoded_selection()),
    resolveSelection: (cursor) => {
      try {
        return JSON.parse(
          session.resolve_encoded_selection(cursor.story, cursor.anchor, cursor.head)
        ) as YrsSelection;
      } catch {
        return null;
      }
    },
    setCellSelection: (range) => session.set_cell_selection(JSON.stringify(range)),
    cellSelection: () => JSON.parse(session.cell_selection()) as YrsTableRange | null,
    beginUndoCapture: ensureUndo,
    addUndoBoundary: () => session.add_undo_boundary(),
    setUndoCaptureMode: (mode) => session.set_undo_capture_mode(mode),
    undoCaptureMode: () => session.undo_capture_mode() as YrsUndoCaptureMode,
    historyStories: () => session.history_stories(),
    undo: () =>
      mutate(() => {
        const applied = session.undo();
        if (applied) markDirty(session.history_stories());
        return applied;
      }),
    redo: () =>
      mutate(() => {
        const applied = session.redo();
        if (applied) markDirty(session.history_stories());
        return applied;
      }),
    canUndo: () => session.can_undo(),
    canRedo: () => session.can_redo(),

    createStory: (storyId, initialText, pStyle = 'Normal', alignment = 'left') => {
      markDirty(storyId);
      return mutate(
        () =>
          JSON.parse(session.create_story(storyId, initialText, pStyle, alignment)) as {
            paraId: string;
          }
      );
    },
    deleteStory: (storyId) => {
      markDirty(storyId);
      return mutate(() => session.delete_story(storyId));
    },
    insertTable: (at, rows, columns, suggesting) => {
      ensureUndo(at.story);
      return mutate(
        () =>
          markReceiptStories(
            JSON.parse(
              session.insert_table(
                at.story,
                at.paraId,
                at.offset,
                rows,
                columns,
                suggesting?.name,
                suggesting?.date
              )
            ) as YrsTableReceipt
          )
      );
    },
    insertRow: (at, side, suggesting) => {
      ensureUndo(at.story);
      return mutate(
        () =>
          markReceiptStories(
            JSON.parse(
              session.insert_row(
                JSON.stringify(at),
                side === 'below',
                suggesting?.name,
                suggesting?.date
              )
            ) as YrsTableReceipt
          )
      );
    },
    insertColumn: (at, side) => {
      ensureUndo(at.story);
      return mutate(
        () =>
          markReceiptStories(
            JSON.parse(session.insert_column(JSON.stringify(at), side === 'right')) as YrsTableReceipt
          )
      );
    },
    deleteRow: (range, suggesting) => {
      ensureUndo(range.anchor.story);
      return mutate(
        () =>
          markReceiptStories(
            JSON.parse(
              session.delete_row(JSON.stringify(range), suggesting?.name, suggesting?.date)
            ) as YrsTableReceipt
          )
      );
    },
    deleteColumn: (range) => {
      ensureUndo(range.anchor.story);
      return mutate(
        () => markReceiptStories(JSON.parse(session.delete_column(JSON.stringify(range))) as YrsTableReceipt)
      );
    },
    deleteTable: (table) => {
      ensureUndo(table.story);
      return mutate(
        () => markReceiptStories(JSON.parse(session.delete_table(JSON.stringify(table))) as YrsTableReceipt)
      );
    },
    mergeCells: (range) => {
      ensureUndo(range.anchor.story);
      return mutate(
        () => markReceiptStories(JSON.parse(session.merge_cells(JSON.stringify(range))) as YrsTableReceipt)
      );
    },
    splitCell: (at, rows, columns) => {
      ensureUndo(at.story);
      return mutate(
        () => markReceiptStories(JSON.parse(session.split_cell(JSON.stringify(at), rows, columns)) as YrsTableReceipt)
      );
    },
    setCellShading: (range, color) => {
      ensureUndo(range.anchor.story);
      return mutate(
        () =>
          markReceiptStories(
            JSON.parse(
              session.set_cell_shading(JSON.stringify(range), color ?? undefined)
            ) as YrsTableReceipt
          )
      );
    },
    setCellTextFormat: (range, patch) => {
      ensureUndo(range.anchor.story);
      return mutate(
        () =>
          markReceiptStories(
            JSON.parse(
              session.set_cell_text_format(JSON.stringify(range), JSON.stringify(patch))
            ) as YrsTableReceipt
          )
      );
    },
    setCellBorders: (range, borders) => {
      ensureUndo(range.anchor.story);
      return mutate(
        () =>
          markReceiptStories(
            JSON.parse(
              session.set_cell_borders(JSON.stringify(range), JSON.stringify(borders))
            ) as YrsTableReceipt
          )
      );
    },
    setColumnWidth: (at, widthTwips) => {
      ensureUndo(at.story);
      return mutate(
        () =>
          markReceiptStories(
            JSON.parse(session.set_column_width(JSON.stringify(at), widthTwips)) as YrsTableReceipt
          )
      );
    },
    setTableWidth: (table, widthTwips) => {
      ensureUndo(table.story);
      return mutate(
        () =>
          markReceiptStories(
            JSON.parse(session.set_table_width(JSON.stringify(table), widthTwips)) as YrsTableReceipt
          )
      );
    },
    insertText: (at, text, suggesting) => {
      ensureUndo(at.story);
      return mutate(
        () =>
          JSON.parse(
            session.insert_text(
              at.story,
              at.paraId,
              at.offset,
              text,
              suggesting?.name,
              suggesting?.date
            )
          ) as YrsReplaceReceipt
      );
    },
    deleteRange: (range, suggesting) => {
      ensureUndo(range.story);
      return mutate(
        () =>
          JSON.parse(
            session.delete_range(
              range.story,
              range.start.paraId,
              range.start.offset,
              range.end.paraId,
              range.end.offset,
              suggesting?.name,
              suggesting?.date
            )
          ) as YrsReplaceReceipt
      );
    },
    replaceRange: (range, text, suggesting) => {
      ensureUndo(range.story);
      return mutate(
        () =>
          JSON.parse(
            session.replace_range(
              range.story,
              range.start.paraId,
              range.start.offset,
              range.end.paraId,
              range.end.offset,
              text,
              suggesting?.name,
              suggesting?.date
            )
          ) as YrsReplaceReceipt
      );
    },
    splitParagraph: (at, suggesting) => {
      ensureUndo(at.story);
      return mutate(
        () =>
          JSON.parse(
            session.split_paragraph(
              at.story,
              at.paraId,
              at.offset,
              suggesting?.name,
              suggesting?.date
            )
          ) as YrsSplitReceipt
      );
    },
    mergeParagraphs: (story, paraId, suggesting) => {
      ensureUndo(story);
      return mutate(
        () =>
          JSON.parse(
            session.merge_paragraphs(story, paraId, suggesting?.name, suggesting?.date)
          ) as YrsRevisionReceipt
      );
    },
    toggleMark: (range, mark) => {
      ensureUndo(range.story);
      mutate(() =>
        session.toggle_mark(
          range.story,
          range.start.paraId,
          range.start.offset,
          range.end.paraId,
          range.end.offset,
          JSON.stringify(mark)
        )
      );
    },
    formatRange: (range, delta) => {
      ensureUndo(range.story);
      mutate(() =>
        session.format_range(
          range.story,
          range.start.paraId,
          range.start.offset,
          range.end.paraId,
          range.end.offset,
          JSON.stringify(delta)
        )
      );
    },
    setHyperlink: (range, hyperlink) => {
      ensureUndo(range.story);
      mutate(() =>
        session.set_hyperlink(
          range.story,
          range.start.paraId,
          range.start.offset,
          range.end.paraId,
          range.end.offset,
          JSON.stringify(hyperlink)
        )
      );
    },
    clearFormatting: (range) => {
      ensureUndo(range.story);
      mutate(() =>
        session.clear_formatting(
          range.story,
          range.start.paraId,
          range.start.offset,
          range.end.paraId,
          range.end.offset
        )
      );
    },
    applyParagraphStyle: (range, styleId, suggesting) => {
      ensureUndo(range.story);
      mutate(() =>
        session.apply_paragraph_style(
          range.story,
          range.start.paraId,
          range.start.offset,
          range.end.paraId,
          range.end.offset,
          styleId,
          suggesting?.name,
          suggesting?.date
        )
      );
    },
    setParagraphAttrs: (range, attrs, suggesting) => {
      ensureUndo(range.story);
      mutate(() =>
        session.set_paragraph_attrs(
          range.story,
          range.start.paraId,
          range.start.offset,
          range.end.paraId,
          range.end.offset,
          JSON.stringify(attrs),
          suggesting?.name,
          suggesting?.date
        )
      );
    },
    insertImage: (at, image, suggesting) => {
      ensureUndo(at.story);
      return mutate(
        () =>
          JSON.parse(
            session.insert_image(
              at.story,
              at.paraId,
              at.offset,
              JSON.stringify(image),
              suggesting?.name,
              suggesting?.date
            )
          ) as YrsReplaceReceipt
      );
    },
    setContentControlValue: (embedId, value) => {
      ensureUndo();
      markDirty('all');
      mutate(() => session.set_content_control_value(embedId, JSON.stringify(value)));
    },
    setContentControlValueAt: (at, value) => {
      ensureUndo(at.story);
      if (typeof value === 'string') markDirty('all');
      mutate(() =>
        session.set_content_control_value_at(at.story, at.paraId, at.offset, JSON.stringify(value))
      );
    },
    clearContentControlValue: (embedId) => {
      ensureUndo();
      markDirty('all');
      mutate(() => session.clear_content_control_value(embedId));
    },
    setImageGeometry: (embedId, geometry) => {
      ensureUndo();
      markDirty('all');
      mutate(() => session.set_image_geometry(embedId, JSON.stringify(geometry)));
    },
    setImageGeometryAt: (at, geometry) => {
      ensureUndo(at.story);
      mutate(() =>
        session.set_image_geometry_at(at.story, at.paraId, at.offset, JSON.stringify(geometry))
      );
    },
    insertPageBreak: (at) => {
      ensureUndo(at.story);
      mutate(() => session.insert_page_break(at.story, at.paraId, at.offset));
    },
    insertSectionBreak: (at, type) => {
      ensureUndo(at.story);
      mutate(() => session.insert_section_break(at.story, at.paraId, at.offset, type));
    },
    insertWatermark: (at, watermark) => {
      ensureUndo(at.story);
      mutate(() =>
        session.insert_watermark(at.story, at.paraId, at.offset, JSON.stringify(watermark))
      );
    },
    applyRawOps: (story, ops) => {
      const rekeys = ops.some(
        (op) =>
          (op.op === 'insertEmbed' && op.kind === 'pilcrow') ||
          (op.op === 'setEmbedAttr' && op.key === 'paraId')
      );
      markDirty(rekeys ? 'all' : story);
      mutate(() => session.apply_raw_ops(story, JSON.stringify(ops)));
    },
    applySeedRawOps: (story, ops) => {
      markDirty(story);
      mutate(() => session.apply_seed_raw_ops(story, JSON.stringify(ops)));
    },
    setParagraphAttr: (paraId, key, value) => {
      markDirty('all');
      mutate(() => session.set_paragraph_attr(paraId, key, JSON.stringify(value ?? null)));
    },
    addComment: (ranges, commentAuthor, date, body) => {
      markDirty(ranges.map((range) => range.story));
      return mutate(
        () =>
          JSON.parse(
            session.add_comment(
              wireRanges(ranges),
              commentAuthor,
              date,
              JSON.stringify(body ?? null)
            )
          ) as YrsCommentReceipt
      );
    },
    setCommentRanges: (commentId, ranges) => {
      ensureUndo();
      mutate(() => {
        session.set_comment_ranges(commentId, wireRanges(ranges));
        markDirty('all');
      });
    },
    acceptChange: (target) => {
      markDirty('all');
      return mutate(
        () => JSON.parse(session.accept_change(wireChangeTarget(target))) as YrsResolveReceipt
      );
    },
    rejectChange: (target) => {
      markDirty('all');
      return mutate(
        () => JSON.parse(session.reject_change(wireChangeTarget(target))) as YrsResolveReceipt
      );
    },

    selectionContext: (range) => {
      const key = JSON.stringify(range);
      if (cachedSelectionContext?.key === key) {
        return JSON.parse(cachedSelectionContext.json) as YrsSelectionContext;
      }
      const json = session.selection_context(
        range.story,
        range.start.paraId,
        range.start.offset,
        range.end.paraId,
        range.end.offset
      );
      const context = JSON.parse(json) as YrsSelectionContext;
      cachedSelectionContext = { key, json };
      return context;
    },
    listRevisions: () => JSON.parse(session.list_revisions()) as YrsRevisionInfo[],
    resolveComment: (commentId) =>
      JSON.parse(session.resolve_comment(commentId)) as YrsResolvedCommentAnchor[],
    listComments: () => JSON.parse(session.list_comments()) as YrsCommentInfo[],
    storyIds: () => session.story_ids(),
    hasStory: (story) => !LONE_SURROGATE.test(story) && session.has_story(story),
    storyLength: (story) => session.story_len(story),
    storyChecksum: (story) => BigInt(session.story_checksum(story)),
    yrsBlocksForStory: (story, env = {}) => {
      const json = session.yrs_blocks_for_story(story, JSON.stringify(env));
      const blocks = JSON.parse(json) as unknown[];
      residentRenderInputs.set(story, structuredClone(env));
      return blocks;
    },
    paragraphs: (story) => JSON.parse(session.paragraphs(story)) as YrsParagraph[],
    searchText: (query, options = {}) => {
      if (!query) return [];
      const limit = options.limit ?? Number.POSITIVE_INFINITY;
      if ((!Number.isSafeInteger(limit) && limit !== Number.POSITIVE_INFINITY) || limit < 0) {
        throw new RangeError('search limit must be a non-negative safe integer');
      }
      return JSON.parse(
        session.search_text(
          query,
          options.caseSensitive ?? false,
          Number.isFinite(limit) ? Math.min(limit, 0xffffffff) : undefined
        )
      ) as YrsTextMatch[];
    },
    paragraphSpans: (story) => JSON.parse(session.paragraph_spans(story)) as YrsParagraphLength[],
    storySegments: (story) => JSON.parse(session.story_segments(story)) as YrsStorySegment[],
    storiesChangedSince: (since) =>
      JSON.parse(session.stories_changed_since(since)) as { revision: number; stories: string[] },
    storySegmentUnitDigests: (story) =>
      JSON.parse(session.story_segment_unit_digests(story)) as string[],
    storySegmentUnits: (story, units) =>
      JSON.parse(session.story_segment_units(story, Uint32Array.from(units))) as YrsStorySegment[][],
    tablePayload: (story, tableIndex) => {
      // No table has an index the u32 boundary would wrap; the story must still exist.
      if (!Number.isInteger(tableIndex) || tableIndex < 0 || tableIndex > 0xffffffff) {
        session.story_len(story);
        return null;
      }
      const payload = session.table_payload(story, tableIndex);
      return payload === undefined ? null : (JSON.parse(payload) as Record<string, unknown>);
    },
    locateParagraph: (story, paraId) =>
      JSON.parse(session.locate_paragraph(story, paraId)) as YrsParagraphSpan,
    paragraphIdCount: (story, paraId) => session.paragraph_id_count(story, paraId),

    paragraphIdentities: () =>
      JSON.parse(session.paragraph_identities()) as DocxParagraphIdentitySnapshot,
    persistParagraphIds: () => {
      markDirty('all');
      return mutate(
        () => JSON.parse(session.persist_paragraph_ids()) as DocxParagraphIdentityReceipt
      );
    },
    resolveParagraphAnchor: (anchor) =>
      JSON.parse(
        session.resolve_paragraph_anchor(JSON.stringify(anchor))
      ) as DocxParagraphAnchorResult,
    storyParagraphIds: (story) =>
      JSON.parse(session.story_paragraph_ids(story)) as Array<string | null>,
    paragraphSavePlan: () => JSON.parse(session.paragraph_save_plan()) as DocxParagraphSavePlan,
    recordSavedParagraphIds: (saved) =>
      mutate(
        () =>
          JSON.parse(session.record_saved_paragraph_ids(JSON.stringify(saved))) as Array<
            [string, string]
          >
      ),
    writtenParagraphIds: (bytes) =>
      JSON.parse(session.written_paragraph_ids(bytes)) as Record<string, string[]>,
    exportStructured: (options) =>
      JSON.parse(
        session.export_structured_json(JSON.stringify(options))
      ) as DocxExportResult<DocxStructuredContent>,
    exportMarkdown: (options) =>
      JSON.parse(
        session.export_markdown_json(JSON.stringify(options))
      ) as DocxExportResult<DocxMarkdownContent>,
    exportStructuredWithPages: (options) =>
      JSON.parse(
        session.export_structured_with_pages_json(JSON.stringify(options), undefined)
      ) as DocxExportResult<DocxPagedStructuredContent<DocxLayoutMap>>,
    exportStructuredWithPagesFor: (options, currentRequest) =>
      JSON.parse(
        session.export_structured_with_pages_json(JSON.stringify(options), currentRequest)
      ) as DocxExportResult<DocxPagedStructuredContent<DocxLayoutMap>>,
    exportSnapshotWithPrivateFonts: (fonts, fontLengths, request, options) =>
      JSON.parse(
        session.export_snapshot_with_private_fonts_json(
          fonts,
          fontLengths,
          request,
          JSON.stringify(options)
        )
      ) as
        | { ok: true; content: DocxPagedStructuredContent<DocxSnapshotLayoutMap> }
        | { ok: false; failure: DocxExportFailure },
    headings: (story) => JSON.parse(session.headings_json(story)) as DocxParagraphHeading[],
    listContentControls: (options = {}) =>
      JSON.parse(
        session.list_content_controls_json(JSON.stringify(options))
      ) as DocxContentControlsResult,
    findContentControls: (query, options = {}) =>
      JSON.parse(
        session.find_content_controls_json(JSON.stringify(query), JSON.stringify(options))
      ) as DocxContentControlsResult,

    version: () => workerDocumentVersion ?? session.version(),
    readParagraphs: (request) =>
      JSON.parse(session.read_paragraphs_json(JSON.stringify(request))) as DocxReadParagraphsResult,
    findText: (request) =>
      JSON.parse(session.find_text_json(JSON.stringify(request))) as DocxFindTextResult,
    validateEdits: (request) =>
      JSON.parse(session.validate_edits_json(JSON.stringify(request))) as DocxValidationResult,
    applyEdits: (request) =>
      mutate(() => {
        const result = JSON.parse(
          session.apply_edits_json(JSON.stringify(request))
        ) as DocxEditResult;
        if (result.ok && result.applied) markDirty(result.changedStories);
        return result;
      }),
    proposeChanges: (request) => mutate(() => proposals.propose(request)),
    setProposalStates: (request) => proposals.setStates(request),
    withdrawProposals: (request) => mutate(() => proposals.withdraw(request)),
    getProposals: () => proposals.snapshot(),
    mirrorWorkerDocument: (mirror) => {
      workerDocumentVersion = mirror?.version ?? null;
      proposals.mirror(mirror);
    },
    workerDocumentMirrored: () => workerDocumentVersion !== null,
    onProposalChange: (listener) => {
      if (destroyed) throw new Error('yrs session is destroyed');
      return proposals.subscribe(listener);
    },
    formatTextTarget: (target, delta) => {
      ensureUndo(targetStory(target));
      return mutate(
        () =>
          JSON.parse(
            session.format_text_target_json(JSON.stringify(target), JSON.stringify(delta))
          ) as YrsTargetEditResult
      );
    },
    commentTextTarget: (target, comment) => {
      markDirty(targetStory(target));
      return mutate(
        () =>
          JSON.parse(
            session.comment_text_target_json(
              JSON.stringify(target),
              JSON.stringify({ ...comment, body: comment.body ?? null })
            )
          ) as YrsTargetEditResult
      );
    },
    selectionText: (range) =>
      JSON.parse(
        session.selection_text_json(
          range.story,
          range.start.paraId,
          range.start.offset,
          range.end.paraId,
          range.end.offset
        )
      ) as YrsSelectionText,

    destroy: () => {
      if (destroyed) return;
      destroyed = true;
      residentFonts.length = 0;
      docxSource = null;
      docxSourceKeys = null;
      resetMedia();
      listeners.clear();
      proposals.destroy();
      pendingUpdates.length = 0;
      if (observing) session.clear_update_observer();
      session.free();
    },
  };

  registerSessionInternals(facade, {
    sourcePackage: () =>
      docxSource && docxSourceKeys
        ? { buffer: docxSourceBuffer(docxSource), keys: docxSourceKeys }
        : null,
    compareDocx: (original, revised, options) => {
      markDirty('all');
      const json = mutate(() => {
        const compared = session.compare_docx_json(original, revised, options);
        proposals.reset();
        return compared;
      });
      docxSource = original.slice();
      docxSourceKeys = null;
      return json;
    },
    finishComparedDocx: (bytes) => session.finish_compared_docx_json(bytes),
    failComparedDocx: (message) => session.fail_compared_docx_json(message),
  });

  return facade;
}

/** Borrows an opened edit session for saving; never destroy the facade. @internal */
export function wrapOpenedEditSession(
  session: EditSession,
  clientId: number,
  source: Uint8Array,
  hostJson: string
): { session: YrsSession; host: YrsDocxHost } {
  const exact = docxSourceBuffer(source) === source.buffer ? source : source.slice();
  const host = decodeDocxHost(hostJson, exact);
  return { session: wrapSession(session, clientId, { source: exact, host }), host };
}
