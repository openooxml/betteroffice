import { decodeEncodedSelection } from './encodedSelection';
import type {
  YrsEngineApplyProfile,
  YrsParagraph,
  YrsParagraphLength,
  YrsParagraphSpan,
  YrsRegionLayoutProgress,
  YrsResidentCaretSnapshot,
  YrsResolvedCommentAnchor,
  YrsRevisionInfo,
  YrsSelection,
  YrsSelectionText,
  YrsSession,
  YrsStorySegment,
  YrsTextMatch,
  YrsLoc,
} from './index';
import type {
  CollaborationTextInsertion,
  CollaborationUpdateOrigin,
} from '../collaboration/types';
import type { DocxEditResult, DocxFindTextResult, DocxReadParagraphsResult } from './edits';
import type {
  DocxParagraphAnchorResult,
  DocxParagraphIdentitySnapshot,
} from './paragraphIdentity';
import type { DocxSidebarReader } from './sidebarReads';
import type { ProposalGeometryReader, ProposalGeometryRevision } from './proposalGeometry';
import type { DocxProposalSession } from './proposals';
import type { DocxPageExportOptions } from './pagedExport';
import type { DocxContentControlsResult } from './contentControls';
import type { YrsPositionOutline } from './yrsPositionProjection';
import { resolveHostJsonCommentMedia } from './hostMedia';
import { createEditSession, preloadEditWasm, setEditWasmHeapLimit } from './wasm/index';

export type ResidentEngineSession = Pick<
  YrsSession,
  | 'applyDelete'
  | 'applyDeleteProfiled'
  | 'applyInput'
  | 'applyInputProfiled'
  | 'applyRawOps'
  | 'applyUpdate'
  | 'beginRegionLayout'
  | 'buildDisplayListFrame'
  | 'buildDisplayPagesFrame'
  | 'releaseDisplayPagesFrame'
  | 'clearFonts'
  | 'destroy'
  | 'encodeSelection'
  | 'encodeStateVector'
  | 'findContentControls'
  | 'listContentControls'
  | 'layoutDocumentJson'
  | 'layoutFontRequirementsJson'
  | 'layoutDocumentWithRegionsRetainedJson'
  | 'loadMediaSources'
  | 'loadNoteSeparators'
  | 'loadState'
  | 'setPartialDocument'
  | 'measureParagraphJson'
  | 'onUpdate'
  | 'outlineGlyphJson'
  | 'registerFont'
  | 'registerSubstituteFont'
  | 'residentCaretSnapshot'
  | 'residentDeletedUnits'
  | 'resumeRegionLayout'
  | 'selection'
  | 'selectionText'
  | 'resolveComment'
  | 'searchText'
  | 'encodeStickyPosition'
  | 'resolveStickyPosition'
  | 'setDisplayRetainBuiltPages'
  | 'setDisplayWindow'
  | 'setSelection'
  | 'storiesChangedSince'
  | 'yrsBlocksForStory'
> & {
  /** @internal */
  proposalEngine: DocxProposalSession;
  /** @internal */
  geometryReader: ProposalGeometryReader & DocxSidebarReader;
  /** @internal The segments of the story's paragraphs at `indices`, each ending with its pilcrow. */
  paragraphSegments(story: string, indices: readonly number[]): YrsStorySegment[][];
  /** @internal */
  paragraphIdentities(): DocxParagraphIdentitySnapshot;
  /** @internal */
  exportStructuredWithPagesJson(options: DocxPageExportOptions, currentRequest: string): string;
  /** The region layout of only as much of the body as fills `pages` pages. */
  layoutDocumentWithRegionsPrefixRetainedJson(input: string, pages: number): string;
  /** Limit incremental rebuilds to the display window and caret pages. Off by default. */
  setWindowedIncrementalBuilds(enabled: boolean): void;
  /** @internal */
  setDirectBatches(enabled: boolean): void;
  /** @internal */
  directBatchesApplied(): number;
  /** Parses and seeds a DOCX; returns the host metadata JSON the main thread decodes. */
  openDocx(bytes: Uint8Array, digest?: string, generation?: string): string;
  /** Opens a display-only preview of the first `blocks` body blocks; null when it refuses. */
  openDocxPreview(bytes: Uint8Array, blocks: number): string | null;
  /** The whole document state as one yrs v1 update. */
  encodeState(): Uint8Array;
  /** Tracked changes in the document, leaving out the revisions in `excluding`. */
  revisionCount(excluding?: ReadonlySet<string>): number;
  /** The retained region layout pass without serializing its reply. */
  layoutDocumentWithRegionsRetained(input: string): void;
  /** The retained region layout's `headersFooters` JSON, when it has any. */
  retainedHeadersFootersJson(): string | undefined;
};

export async function createResidentEngineSession(
  heapLimitBytes?: number,
  clientId = randomClientId()
): Promise<ResidentEngineSession> {
  await preloadEditWasm();
  setEditWasmHeapLimit(heapLimitBytes);
  const session = createEditSession(clientId);
  session.set_local_lowering(true);
  const listeners = new Set<
    (update: Uint8Array, origin: CollaborationUpdateOrigin) => void
  >();
  let observing = false;
  let destroyed = false;
  let undoTracked = false;
  const geometryStories = new Map<string, {
    revision: number;
    segments: YrsStorySegment[];
  }>();

  const geometrySpans = new Map<string, { revision: number; spans: YrsParagraphLength[] }>();
  const geometryOutlines = new Map<string, { version: string; outline: YrsPositionOutline | null }>();
  const storyRevisions = new Map<string, number>();
  let nativeStoryRevision = 0;
  let storyRevision = 0;

  const syncStoryRevisions = (): number => {
    const changes = JSON.parse(session.stories_changed_since(nativeStoryRevision)) as {
      revision: number;
      stories: string[];
    };
    if (changes.revision !== nativeStoryRevision) {
      nativeStoryRevision = changes.revision;
      storyRevision += 1;
      for (const story of changes.stories) storyRevisions.set(story, storyRevision);
    }
    return storyRevision;
  };

  const storiesChangedSince = (since: number) => {
    const revision = syncStoryRevisions();
    if (since >= revision) return { revision, stories: [] };
    return {
      revision,
      stories: [...storyRevisions].filter(([, changed]) => changed > since)
        .map(([story]) => story).sort(),
    };
  };

  const storyChangedSince = (story: string, since: number): boolean =>
    (storyRevisions.get(story) ?? 0) > since;

  const geometryStory = (story: string) => {
    const revision = syncStoryRevisions();
    let cached = geometryStories.get(story);
    if (!cached || storyChangedSince(story, cached.revision)) {
      cached = {
        revision,
        segments: JSON.parse(session.story_segments(story)) as YrsStorySegment[],
      };
      geometryStories.set(story, cached);
    } else {
      cached.revision = revision;
    }
    return cached;
  };

  const ensureUndo = (): void => {
    if (undoTracked) return;
    session.track_undo();
    undoTracked = true;
  };

  const ensureObserver = (): void => {
    if (observing) return;
    session.set_update_observer((update: Uint8Array, origin: number) => {
      if (origin !== 0 && origin !== 1) return;
      for (const listener of [...listeners]) {
        listener(update, origin === 0 ? 'local' : 'remote');
      }
    });
    observing = true;
  };

  const proposalEngine: DocxProposalSession = {
    version: () => session.version(),
    resolveParagraphAnchor: (anchor) =>
      JSON.parse(
        session.resolve_paragraph_anchor(JSON.stringify(anchor))
      ) as DocxParagraphAnchorResult,
    findText: (request) =>
      JSON.parse(session.find_text_json(JSON.stringify(request))) as DocxFindTextResult,
    readParagraphs: (request) =>
      JSON.parse(session.read_paragraphs_json(JSON.stringify(request))) as DocxReadParagraphsResult,
    applyEdits: (request) =>
      JSON.parse(session.apply_edits_json(JSON.stringify(request))) as DocxEditResult,
    listRevisions: () =>
      JSON.parse(session.list_revisions()) as ReturnType<DocxProposalSession['listRevisions']>,
    revisionStamps: (ids) =>
      JSON.parse(session.revision_stamps_json(JSON.stringify(ids))) as ReturnType<
        NonNullable<DocxProposalSession['revisionStamps']>
      >,
    settleRevisions: (accept, reject) => {
      session.settle_revisions_json(JSON.stringify({ accept, reject }));
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
  };

  const geometryReader: ProposalGeometryReader & DocxSidebarReader = {
    resolveComment: (id) => JSON.parse(session.resolve_comment(id)),
    headings: (story) => JSON.parse(session.headings_json(story)),
    version: () => session.version(),
    hasStory: (story) => !LONE_SURROGATE.test(story) && session.has_story(story),
    storyIds: () => session.story_ids(),
    paragraphs: (story) => JSON.parse(session.paragraphs(story)) as YrsParagraph[],
    paragraphIdCount: (story, paraId) => session.paragraph_id_count(story, paraId),
    paragraphSpans: (story) => {
      const revision = syncStoryRevisions();
      const cached = geometrySpans.get(story);
      if (cached && !storyChangedSince(story, cached.revision)) {
        cached.revision = revision;
        return cached.spans;
      }
      const spans = JSON.parse(session.paragraph_spans(story)) as YrsParagraphLength[];
      geometrySpans.set(story, { revision, spans });
      return spans;
    },
    storySegments: (story) => geometryStory(story).segments,
    positionOutline: (root) => {
      if (LONE_SURROGATE.test(root)) return null;
      const version = session.version();
      const cached = geometryOutlines.get(root);
      if (cached?.version === version) return cached.outline;
      const parsed = JSON.parse(session.geometry_position_outline_json(root)) as
        YrsPositionOutline | 'legacy';
      const outline = parsed === 'legacy' ? null : parsed;
      geometryOutlines.set(root, { version, outline });
      return outline;
    },
    locateParagraph: (story, paraId) =>
      JSON.parse(session.locate_paragraph(story, paraId)) as YrsParagraphSpan,
    listRevisions: () => JSON.parse(session.list_revisions()) as YrsRevisionInfo[],
    resolveParagraphAnchor: proposalEngine.resolveParagraphAnchor,
    findText: proposalEngine.findText,
    proposalRevisions: (ids) => {
      const result = JSON.parse(session.proposal_revision_ranges_json(JSON.stringify(ids))) as
        ProposalGeometryRevision[] | 'fallback' | 'legacy';
      if (Array.isArray(result)) return result;
      if (result === 'fallback') {
        const owned = new Set(ids);
        return geometryReader.listRevisions().filter(({ revisionId }) => owned.has(revisionId));
      }
      return readLegacyProposalRevisions(geometryReader, ids);
    },
  };

  return {
    proposalEngine,
    geometryReader,
    resolveComment: (commentId) =>
      JSON.parse(session.resolve_comment(commentId)) as YrsResolvedCommentAnchor[],
    paragraphSegments: (story, indices) =>
      JSON.parse(session.story_segment_units(story, Uint32Array.from(indices))) as YrsStorySegment[][],
    searchText: (query, options = {}) => {
      if (!query) return [];
      const limit = options.limit ?? Number.POSITIVE_INFINITY;
      if ((!Number.isSafeInteger(limit) && limit !== Number.POSITIVE_INFINITY) || limit < 0) {
        throw new RangeError('search limit must be a non-negative safe integer');
      }
      return JSON.parse(session.search_text(
        query,
        options.caseSensitive ?? false,
        Number.isFinite(limit) ? Math.min(limit, 0xffffffff) : undefined
      )) as YrsTextMatch[];
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
    paragraphIdentities: () =>
      JSON.parse(session.paragraph_identities()) as DocxParagraphIdentitySnapshot,
    exportStructuredWithPagesJson: (options, currentRequest) =>
      session.export_structured_with_pages_json(JSON.stringify(options), currentRequest),
    listContentControls: (options = {}) =>
      JSON.parse(
        session.list_content_controls_json(JSON.stringify(options))
      ) as DocxContentControlsResult,
    findContentControls: (query, options = {}) =>
      JSON.parse(
        session.find_content_controls_json(JSON.stringify(query), JSON.stringify(options))
      ) as DocxContentControlsResult,
    storiesChangedSince,
    openDocx: (bytes, digest, generation) => {
      geometryStories.clear();
      return resolveHostJsonCommentMedia(
        session.open_docx(bytes, true, generation, digest),
        (token) => (token.startsWith('media:') ? (session.media_data_url(token) ?? null) : null)
      );
    },
    openDocxPreview: (bytes, blocks) => {
      geometryStories.clear();
      const json = session.open_docx_preview(bytes, blocks);
      return json === undefined
        ? null
        : resolveHostJsonCommentMedia(
            json,
            (token) => (token.startsWith('media:') ? (session.media_data_url(token) ?? null) : null)
          );
    },
    encodeState: () => session.encode_state(),
    revisionCount: (excluding) =>
      (JSON.parse(session.list_revisions()) as { revisionId: string }[]).filter(
        (revision) => !excluding?.has(revision.revisionId)
      ).length,
    registerFont: (bytes) => session.register_measure_font(bytes),
    registerSubstituteFont: (base, family) =>
      session.register_substitute_measure_font(base, family),
    clearFonts: () => session.clear_measure_fonts(),
    encodeStateVector: () => session.encode_state_vector(),
    measureParagraphJson: (input) => session.measure_paragraph_json(input),
    layoutDocumentJson: (input) => session.layout_document_json(input),
    layoutFontRequirementsJson: (input) => session.layout_font_requirements_json(input),
    layoutDocumentWithRegionsRetainedJson: (input) =>
      session.layout_document_with_regions_retained_json(input),
    beginRegionLayout: (input) =>
      JSON.parse(session.begin_region_layout(input)) as YrsRegionLayoutProgress,
    resumeRegionLayout: (blocks) =>
      JSON.parse(session.resume_region_layout(blocks)) as YrsRegionLayoutProgress,
    setPartialDocument: (partial) => session.set_partial_document(partial),
    layoutDocumentWithRegionsPrefixRetainedJson: (input, pages) =>
      session.layout_document_with_regions_prefix_retained_json(input, pages),
    layoutDocumentWithRegionsRetained: (input) =>
      session.layout_document_with_regions_retained(input),
    retainedHeadersFootersJson: () => session.retained_headers_footers_json(),
    buildDisplayListFrame: (input, expectedFrameEpoch) =>
      session.build_display_list_frame(input, expectedFrameEpoch),
    setDisplayWindow: (start, end) => session.set_display_window(start, end),
    setDisplayRetainBuiltPages: (retain) => session.set_display_retain_built_pages(retain),
    setWindowedIncrementalBuilds: (enabled) => session.set_windowed_incremental_builds(enabled),
    setDirectBatches: (enabled) => session.set_direct_batches(enabled),
    directBatchesApplied: () => session.direct_batches_applied(),
    buildDisplayPagesFrame: (pages, expectedFrameEpoch) =>
      session.build_display_pages_frame(Uint32Array.from(pages), expectedFrameEpoch),
    releaseDisplayPagesFrame: (pages, expectedFrameEpoch) => {
      const frame = session.release_display_pages_frame(Uint32Array.from(pages), expectedFrameEpoch);
      return frame.length === 0 ? null : frame;
    },
    residentCaretSnapshot: () =>
      JSON.parse(session.resident_caret_snapshot_json()) as YrsResidentCaretSnapshot,
    selection: () => JSON.parse(session.selection()) as YrsSelection | null,
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
    encodeSelection: () => decodeEncodedSelection(session.encoded_selection()),
    applyInput: (text, expectedFrameEpoch) => {
      ensureUndo();
      return session.apply_input(text, expectedFrameEpoch);
    },
    applyDelete: (direction, expectedFrameEpoch, count = 1) => {
      ensureUndo();
      return session.apply_delete(direction, expectedFrameEpoch, count);
    },
    residentDeletedUnits: () => session.resident_deleted_units(),
    applyRawOps: (story, ops) => {
      geometryStories.clear();
      const version = session.version();
      try {
        session.apply_raw_ops(story, JSON.stringify(ops));
      } finally {
        if (session.version() !== version) {
          syncStoryRevisions();
          storyRevision += 1;
          storyRevisions.set(story, storyRevision);
        }
      }
    },
    applyInputProfiled: (text, expectedFrameEpoch) => {
      ensureUndo();
      const frame = session.apply_input_profiled(text, expectedFrameEpoch);
      const profile = JSON.parse(session.apply_input_profile_json()) as YrsEngineApplyProfile;
      return { frame, profile };
    },
    applyDeleteProfiled: (direction, expectedFrameEpoch, count = 1) => {
      ensureUndo();
      const frame = session.apply_delete_profiled(direction, expectedFrameEpoch, count);
      const profile = JSON.parse(session.apply_input_profile_json()) as YrsEngineApplyProfile;
      return { frame, profile };
    },
    outlineGlyphJson: (fontId, glyphId) => session.outline_glyph_json(fontId, glyphId),
    loadMediaSources: (json) => session.load_media_sources(json),
    loadNoteSeparators: (state) => session.load_note_separators(state),
    loadState: (update) => {
      geometryStories.clear();
      session.load(update);
    },
    applyUpdate: (update) =>
      JSON.parse(
        session.apply_update_with_inference(update)
      ) as CollaborationTextInsertion | null,
    onUpdate: (listener) => {
      listeners.add(listener);
      ensureObserver();
      return () => listeners.delete(listener);
    },
    setSelection: (anchor, head = anchor) => {
      if (anchor.story !== head.story) throw new Error('yrs selection must stay inside one story');
      session.set_selection(anchor.story, anchor.paraId, anchor.offset, head.paraId, head.offset);
    },
    yrsBlocksForStory: (story, env = {}) =>
      JSON.parse(session.yrs_blocks_for_story(story, JSON.stringify(env))) as unknown[],
    destroy: () => {
      if (destroyed) return;
      destroyed = true;
      listeners.clear();
      geometryStories.clear();
      storyRevisions.clear();
      if (observing) session.clear_update_observer();
      session.free();
    },
  };
}

export function readLegacyProposalRevisions(
  reader: ProposalGeometryReader,
  ids: readonly string[]
): ProposalGeometryRevision[] {
  const owned = new Set(ids);
  const revisions: ProposalGeometryRevision[] = [];
  const fallback = () =>
    reader.listRevisions().filter(({ revisionId }) => owned.has(revisionId));
  for (const story of reader.storyIds().sort()) {
    let offset = 0;
    const paragraphs = new Set<string>();
    let changes: Array<{
      revisionId: string;
      kind: 'insertion' | 'deletion';
      start: number;
      end: number;
    }> = [];
    const previous = new Map<string, (typeof changes)[number]>();
    for (const segment of reader.storySegments(story)) {
      if (segment.kind === 'pilcrow') {
        if (hasRevisionProperties(segment.properties) || paragraphs.has(segment.paraId)) {
          return fallback();
        }
        paragraphs.add(segment.paraId);
        for (const change of changes.sort((a, b) => a.start - b.start)) {
          revisions.push({
            revisionId: change.revisionId,
            kind: change.kind,
            story,
            range: {
              story,
              start: { paraId: segment.paraId, offset: change.start },
              end: { paraId: segment.paraId, offset: change.end },
            },
          });
        }
        offset = 0;
        changes = [];
        previous.clear();
        continue;
      }
      if (segment.kind === 'embed' && hasRevisionProperties(segment.payload)) {
        return fallback();
      }
      const length = segment.kind === 'text' ? segment.text.length : 1;
      for (const [key, kind] of [['ins', 'insertion'], ['del', 'deletion']] as const) {
        const value = segment.attributes[key];
        if (!value || typeof value !== 'object' || Array.isArray(value)) continue;
        const attributes = value as Record<string, unknown>;
        const info = attributes.info;
        const stamp = info && typeof info === 'object' && !Array.isArray(info)
          ? info as Record<string, unknown>
          : attributes;
        const id = 'id' in stamp ? stamp.id : stamp.revisionId;
        if (typeof id !== 'string' && !(typeof id === 'number' && Number.isFinite(id))) {
          continue;
        }
        const revisionId = String(id);
        if (!owned.has(revisionId)) {
          previous.delete(kind);
          continue;
        }
        const last = previous.get(kind);
        if (last?.revisionId === revisionId && last.end === offset) {
          last.end += length;
        } else {
          const change = { revisionId, kind, start: offset, end: offset + length };
          changes.push(change);
          previous.set(kind, change);
        }
      }
      offset += length;
    }
    if (changes.length > 0) return fallback();
  }
  return revisions;
}

const LONE_SURROGATE = /[\uD800-\uDBFF](?![\uDC00-\uDFFF])|(?<![\uD800-\uDBFF])[\uDC00-\uDFFF]/;

function hasRevisionProperties(value: unknown): boolean {
  if (!value || typeof value !== 'object') return false;
  return Object.entries(value).some(([key, child]) =>
    key === 'id' || key === 'revisionId' || key === 'pPrIns' || key === 'pPrDel' ||
    key === 'pPrChange' || key === 'trIns' || key === 'trDel' ||
    key === 'tableIns' || key === 'tableDel' || hasRevisionProperties(child)
  );
}

function randomClientId(): number {
  if (typeof crypto !== 'undefined' && typeof crypto.getRandomValues === 'function') {
    const buffer = new Uint32Array(1);
    crypto.getRandomValues(buffer);
    return buffer[0];
  }
  return Math.floor(Math.random() * 0xffffffff);
}
