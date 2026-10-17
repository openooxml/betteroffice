import { readRetainedLayoutMeta } from './layoutMeta';
import { PeerMetadataError } from './peerMetadata';
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
  YrsStoryRange,
  YrsTextMatch,
  YrsLoc,
} from './index';
import type {
  CollaborationTextInsertion,
  CollaborationUpdateOrigin,
} from '../collaboration/types';
import { requireUnreservedClientId } from '../collaboration/protocol';
import type {
  DocxEditRefusal,
  DocxEditResult,
  DocxFindTextResult,
  DocxReadParagraphsResult,
  DocxReadStoriesResult,
} from './edits';
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
import type { ResidentSaveRecord } from './residentSave';
import type { Comment } from '../types/content';
import type { Document, SectionProperties } from '../types/document';
import type { DocxListStoriesResult } from './stories';
import type { DocxCommentPatch } from './residentReview';
import { projectResidentComments, type ResidentCommentProjection } from './residentCommentProjection';
import { DirtyProjectionStories } from './dirtyProjectionStories';
import { resolveHostJsonCommentMedia } from './hostMedia';
import { createEditSession, preloadEditWasm, setEditWasmHeapLimit } from './wasm/index';
import { decodeDocxHostJson, wrapSession } from './yrsSessionFacade';
import { documentToYrs } from './documentToYrs';
import { tokenizeResidentDocument } from './residentSource';
import { partMetadata, withPartMetadata, readPartGeometry, type ResidentPartMetadata, type ResidentPartGeometry } from './residentParts';
import type { RetainedFrame } from '../layout/render/frameDelta';
import type { DisplayListHitRegion, DisplayListRect } from '../layout/render/displayListQueries';
import type { ResidentControlTarget, ResidentControlPosition } from './residentControls';
import { withPackageMetadata, type ResidentPackageMetadata } from './residentProperties';
import { getDocumentWatermark } from '../docx/watermarkApi';
import { resolvedFinalSectionProperties, updateFinalSectionProperties } from '../editor/finalSection';
import { readSectionSnapshot, residentSectionPropertiesAt, sectionMarkKey, withSectionIds, withSectionNames, withSectionSnapshot, type DocxSectionMargins, type ResidentSectionSnapshot } from './residentSections';
import { pageBreakTarget, splitSection } from './sectionBreak';

/** @internal */
export interface ResidentExactPrefixLayout {
  provisional: boolean;
  pages: number;
  /** The widest page of the whole document, while only a prefix is laid out. */
  documentWidth?: number;
}

export type ResidentEngineSession = Pick<
  YrsSession,
  | 'applyDelete'
  | 'applyDeleteProfiled'
  | 'applyInput'
  | 'applyInputProfiled'
  | 'applyRawOps'
  | 'applyUpdate'
  | 'paragraphs'
  | 'selectionContext'
  | 'storySegments'
  | 'cellSelection'
  | 'setCellSelection'
  | 'insertRow'
  | 'insertColumn'
  | 'deleteRow'
  | 'deleteColumn'
  | 'deleteTable'
  | 'mergeCells'
  | 'splitCell'
  | 'setCellShading'
  | 'setCellBorders'
  | 'setTableProperties'
  | 'insertTable'
  | 'insertTableFragment'
  | 'validateTableFragment'
  | 'insertImage'
  | 'insertPageBreak'
  | 'insertSectionBreak'
  | 'tablePayload'
  | 'setHyperlink'
  | 'setImageGeometryAt'
  | 'setColumnWidth'
  | 'insertText'
  | 'replaceRange'
  | 'deleteRange'
  | 'splitParagraph'
  | 'toggleMark'
  | 'formatRange'
  | 'clearFormatting'
  | 'setParagraphAttrs'
  | 'applyParagraphStyle'
  | 'beginUndoCapture'
  | 'addUndoBoundary'
  | 'canUndo'
  | 'canRedo'
  | 'undo'
  | 'redo'
  | 'addComment'
  | 'setCommentRanges'
  | 'listComments'
  | 'acceptChange'
  | 'rejectChange'
  | 'beginRegionLayout'
  | 'buildDisplayListFrame'
  | 'buildDisplayPagesFrame'
  | 'releaseDisplayPagesFrame'
  | 'clearFonts'
  | 'destroy'
  | 'encodeSelection'
  | 'encodeStateAsUpdate'
  | 'encodeStateVector'
  | 'findContentControls'
  | 'listContentControls'
  | 'layoutDocumentJson'
  | 'layoutFontRequirementsJson'
  | 'layoutDocumentWithRegionsRetainedJson'
  | 'layoutDocumentWithRegionsRetainedMeta'
  | 'retainedLayoutJson'
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
  | 'readStories'
  | 'listStories'
  | 'storiesChangedSince'
  | 'yrsBlocksForStory'
> & {
  setUndoInputTime(span: readonly [first: number, last: number] | null): void;
  splitSection(at: YrsLoc, properties: SectionProperties, start: NonNullable<SectionProperties['sectionStart']>): YrsLoc;
  /**
   * Display ranges of the paragraphs, tables and block controls of `root` that `ranges` touch.
   * A paragraph's `length` counts story units; its inline `controls` give each node's story
   * offset and display range. `selection` maps both ends to display positions.
   */
  residentGeometry(root: string, ranges: Array<[number, number]>, selection?: [YrsLoc, YrsLoc]): {
    paragraphs: Array<{ story: string; paraId: string; index: number; from: number; to: number; leading: number; length: number;
      controls?: Array<{ offset: number; from: number; to: number }> }>;
    tables: Array<{ story: string; tableIndex: number; from: number; to: number; at: YrsLoc;
      cells: Array<{ story: string; paraId: string; row: number; column: number; rowspan: number; colspan: number; from: number; to: number }> }>;
    controls: Array<{ story: string; contentStory: string; from: number; to: number }>;
    selection: [number, number] | null;
  };
  /** Rects of the resident display list for each `[from, to)` range of one region, in order. */
  residentRangeRects(region: DisplayListHitRegion, partId: string, ranges: readonly (readonly [number, number])[]): DisplayListRect[][];
  clearCellSelection(): YrsLoc | null;
  controlTargets(): ResidentControlTarget[];
  repeatControlAt(at: ResidentControlPosition, remove: boolean): boolean;
  setControlValue(at: ResidentControlPosition, value: import('./index').YrsContentControlValue): void;
  partDocument(): Document;
  setPartDocument(document: Document, removeStories?: string[]): boolean;
  partStory(relationshipId: string): string;
  partGeometry(frame: RetainedFrame): ResidentPartGeometry[];
  propertyDocument(): Document;
  setPropertyDocument(document: Document): boolean;
  sectionProperties(): ResidentSectionSnapshot;
  setSectionMargins(section: string, margins: DocxSectionMargins): boolean;
  updateComment(id: string, patch: DocxCommentPatch): void;
  commentProjection(): ResidentCommentProjection | null;
  mediaDataUrl(token: string): string | null;
  /** @internal Typing echo model of one font at one shaping size; see `echo_glyph_set_json`. */
  echoGlyphSetJson(fontId: number, size: number, kern: boolean, alphabet: string): string;
  /** @internal */
  proposalEngine: DocxProposalSession & Pick<YrsSession, 'validateEdits'>;
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
  /**
   * @internal The retained region layout of exactly the first `pages` pages when nothing later
   * can change them (`provisional`), or of the whole document.
   */
  layoutDocumentWithRegionsExactPrefixRetained(input: string, pages: number): ResidentExactPrefixLayout;
  /** Limit incremental rebuilds to the display window and caret pages. Off by default. */
  setWindowedIncrementalBuilds(enabled: boolean): void;
  /** @internal */
  setDirectBatches(enabled: boolean): void;
  /** @internal */
  directBatchesApplied(): number;
  /** @internal */
  markProjectionStories(stories: readonly string[]): void;
  /** Whether a remote update is held until the one it depends on arrives. */
  hasPendingUpdate(): boolean;
  /** Parses and seeds a DOCX; returns the host metadata JSON the main thread decodes. */
  openDocx(bytes: Uint8Array, digest?: string, generation?: string, seedStories?: boolean): string;
  /** With `initialUpdate`, the shared state replaces the package's stories. */
  openSourceDocx(bytes: Uint8Array, options: { digest?: string; generation?: string; mediaTokens?: boolean; initialUpdate?: Uint8Array }): string;
  decodeFont(bytes: Uint8Array): Uint8Array;
  openDocument(document: Document, options: { generation?: string; mediaTokens?: boolean; initialUpdate?: Uint8Array }): void;
  saveDocument(host: Document | undefined, comments: Comment[], record: ResidentSaveRecord): Promise<ArrayBuffer>;
  /** Opens a display-only preview with a block limit and optional paragraph budget; null when refused. */
  openDocxPreview(bytes: Uint8Array, blocks: number, paragraphBudget?: number): string | null;
  /** The whole document state as one yrs v1 update. */
  encodeState(): Uint8Array;
  encodePeerMetadata?(): Uint8Array;
  /** Tracked changes in the document, leaving out the revisions in `excluding`. */
  revisionCount(excluding?: ReadonlySet<string>): number;
  /** The retained region layout pass without serializing its reply. */
  layoutDocumentWithRegionsRetained(input: string): void;
  /** The region pass after one edit, resident when that edit was a plain text edit. */
  layoutDocumentWithRegionsAfterEdit(input: string): void;
  /** The retained region layout's `headersFooters` JSON, when it has any. */
  retainedHeadersFootersJson(): string | undefined;
  /** @internal */
  save(
    source: Uint8Array,
    hostJson: string,
    host: Document | undefined,
    comments: Comment[],
    record: ResidentSaveRecord,
    stories?: readonly string[]
  ): Promise<ArrayBuffer>;
};

export async function createResidentEngineSession(
  heapLimitBytes?: number,
  clientId = randomClientId()
): Promise<ResidentEngineSession> {
  requireUnreservedClientId(clientId);
  await preloadEditWasm();
  setEditWasmHeapLimit(heapLimitBytes);
  const session = createEditSession(clientId);
  let partBase: Document | undefined;
  const commentAliases = new Map<string, number>();
  let partsTouched = false;
  let marginsTouched = false;
  let sectionsTouched = false;
  const seededSectionMarks = new Map<string, string>();
  let sectionNames = { host: new Map<string, string>(), resident: new Map<string, string>() };
  let sectionLayout: { revision: string; sections: ResidentSectionSnapshot['sections'] } | null = null;
  const parsedMedia = new Map<string, string>();
  let parsedFacade: YrsSession | undefined;
  const partDocument = (): Document => {
    if (!partBase) throw new Error('Part metadata requires an opened package');
    const stored = JSON.parse(session.part_metadata_json()) as ResidentPartMetadata;
    const properties = JSON.parse(session.document_properties_json()) as { metadata?: ResidentPackageMetadata };
    const document = stored.body ? withPackageMetadata(withPartMetadata(partBase, stored), stored)
      : properties.metadata ? withPackageMetadata(partBase, properties.metadata) : partBase;
    const margins = JSON.parse(session.section_margins_json()) as Record<string, DocxSectionMargins>;
    if (sectionsTouched) {
      const revision = JSON.stringify([session.resident_section_revision(),
        document.package.document.sections?.map(section => [section.id, section.properties]),
        document.package.document.finalSectionProperties, margins]);
      if (sectionLayout?.revision !== revision) sectionLayout = { revision,
        sections: readSectionSnapshot(resident, document, margins, seededSectionMarks, true).sections };
      const sections = sectionLayout.sections;
      return { ...document, package: { ...document.package, document: { ...document.package.document,
        sections: sections.map(section => ({ id: section.section, properties: section.properties, content: [] })),
        finalSectionProperties: sections.at(-1)!.properties,
      } } };
    }
    return Object.keys(margins).length ? withSectionSnapshot(document, readSectionSnapshot(resident, document, margins, seededSectionMarks)) : document;
  };
  const hostDocument = (): Document => withSectionNames(partDocument(), sectionNames.host);
  const partLayoutInput = (input: string): string => {
    if (!partsTouched && !marginsTouched && !sectionsTouched && session.part_metadata_json() === '{}' &&
      session.document_properties_json() === '{}' && session.section_margins_json() === '{}') return input;
    const document = hostDocument();
    const request = JSON.parse(input);
    const sections = document.package.document.sections?.map(section => ({ sectionId: section.id ?? section.properties.sectionId, properties: section.properties })) ?? [];
    if (!sectionsTouched) sections.push({ sectionId: sections.at(-1)?.sectionId, properties: document.package.document.finalSectionProperties ?? {} });
    const variants = sections.flatMap((section, sectionIndex) => (['header', 'footer'] as const).flatMap(kind =>
      (section.properties[kind === 'header' ? 'headerReferences' : 'footerReferences'] ?? []).map(ref => ({ kind, type: ref.type, rId: ref.rId, sectionIndex }))));
    const contents = [...(document.package.footnotes ?? []).map(note => ({ id: note.id, noteKind: 'footnote', height: 0 })),
      ...(document.package.endnotes ?? []).map(note => ({ id: note.id, noteKind: 'endnote', height: 0 }))];
    const watermark = getDocumentWatermark(document);
    return JSON.stringify({ ...request, regions: { ...request.regions, sections, watermark,
      headersFooters: request.regions?.headersFooters ? { ...request.regions.headersFooters, watermark } : undefined }, notes: { ...request.notes, contents },
      ...(request.workerOwned ? { workerHeadersFooters: { ...request.workerHeadersFooters, variants } } : {}) });
  };
  const keySections = (document: Document): Document => {
    sectionNames = { host: new Map(), resident: new Map() };
    seededSectionMarks.clear();
    if (!geometryReader.hasStory('body')) return document;
    const closing = resident.paragraphs('body')
      .filter(paragraph => paragraph.properties.sectPr ?? paragraph.properties._sectionProperties);
    for (const paragraph of closing) {
      seededSectionMarks.set(paragraph.paraId,
        sectionMarkKey(paragraph.properties.sectPr ?? paragraph.properties._sectionProperties));
    }
    if ((document.package.document.sections?.length ?? 0) < 2) return document;
    const keyed = withSectionIds(document, closing.map(paragraph => paragraph.paraId));
    document.package.document.sections!.forEach(({ id }, index) => {
      const key = keyed.package.document.sections![index]!.id;
      if (id === undefined || key === undefined || id === key) return;
      sectionNames.host.set(key, id);
      sectionNames.resident.set(id, key);
    });
    return keyed;
  };
  const setSectionMark = (paraId: string, sectPr: SectionProperties): boolean => {
    const version = session.version();
    const at = { paraId, offset: 0 };
    resident.setParagraphAttrs({ story: 'body', start: at, end: at }, { other: { sectPr, sectionBreakType: sectPr.sectionStart ?? null } });
    return session.version() !== version;
  };
  const sectionProperties = (): ResidentSectionSnapshot => {
    return readSectionSnapshot(resident, partDocument(), JSON.parse(session.section_margins_json()), seededSectionMarks, sectionsTouched);
  };
  session.set_local_lowering(true);
  const listeners = new Set<
    (update: Uint8Array, origin: CollaborationUpdateOrigin) => void
  >();
  let observing = false;
  let dispatchingUpdates = false;
  let destroyed = false;
  let undoTracked = false;
  const geometryStories = new Map<string, {
    revision: number;
    segments: YrsStorySegment[];
  }>();

  const geometrySpans = new Map<string, { revision: number; spans: YrsParagraphLength[] }>();
  const geometryOutlines = new Map<string, { version: string; outline: YrsPositionOutline | null }>();
  const storyRevisions = new Map<string, number>();
  const projectionStories = new DirtyProjectionStories();
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

  const detachObserver = (): void => {
    if (destroyed || !observing || listeners.size > 0) return;
    session.clear_update_observer();
    observing = false;
  };

  const destroyNative = (): void => {
    if (observing) {
      session.clear_update_observer();
      observing = false;
    }
    session.free();
  };

  const ensureObserver = (): void => {
    if (observing) return;
    session.set_update_observer((update: Uint8Array, origin: number) => {
      if (origin !== 0 && origin !== 1) return;
      dispatchingUpdates = true;
      try {
        for (const listener of [...listeners]) {
          listener(update, origin === 0 ? 'local' : 'remote');
        }
      } finally {
        dispatchingUpdates = false;
      }
    });
    observing = true;
  };

  const markProjectionStories = (stories: readonly string[]): void => {
    for (const story of stories) projectionStories.add(story);
  };

  const applyResidentInput = <T>(apply: () => T): T => {
    ensureUndo();
    const result = apply();
    const selection = JSON.parse(session.selection()) as YrsSelection | null;
    projectionStories.add(selection?.head.story ?? 'body');
    return result;
  };

  const rangeArgs = (range: YrsStoryRange): [string, string, number, string, number] =>
    [range.story, range.start.paraId, range.start.offset, range.end.paraId, range.end.offset];

  const applyCommandMutation = <T>(story: string, apply: () => T): T => {
    ensureUndo();
    const version = session.version();
    try {
      return apply();
    } finally {
      if (session.version() !== version) projectionStories.add(story);
    }
  };

  const proposalEngine: ResidentEngineSession['proposalEngine'] = {
    validateEdits: request => JSON.parse(session.validate_edits_json(JSON.stringify(request))),
    version: () => session.version(),
    resolveParagraphAnchor: (anchor) =>
      JSON.parse(
        session.resolve_paragraph_anchor(JSON.stringify(anchor))
      ) as DocxParagraphAnchorResult,
    findText: (request) =>
      JSON.parse(session.find_text_json(JSON.stringify(request))) as DocxFindTextResult,
    readParagraphs: (request) =>
      JSON.parse(session.read_paragraphs_json(JSON.stringify(request))) as DocxReadParagraphsResult,
    applyEdits: (request) => {
      const result = JSON.parse(session.apply_edits_json(JSON.stringify(request))) as DocxEditResult;
      if (result.ok && result.applied) markProjectionStories(result.changedStories);
      return result;
    },
    listRevisions: () =>
      JSON.parse(session.list_revisions()) as ReturnType<DocxProposalSession['listRevisions']>,
    revisionStamps: (ids) =>
      JSON.parse(session.revision_stamps_json(JSON.stringify(ids))) as ReturnType<
        NonNullable<DocxProposalSession['revisionStamps']>
      >,
    settleRevisions: (accept, reject) => {
      const since = storiesChangedSince(Number.MAX_SAFE_INTEGER).revision;
      const settled = JSON.parse(
        session.settle_revisions_json(JSON.stringify({ accept, reject }))
      ) as { ok: true } | DocxEditRefusal;
      markProjectionStories(storiesChangedSince(since).stories);
      return settled.ok ? undefined : settled.failure;
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
    headerFooterAliases: () => {
      const value: unknown = JSON.parse(session.header_footer_aliases_json());
      return value !== null && typeof value === 'object' && !Array.isArray(value)
        ? new Map(Object.entries(value).filter((entry): entry is [string, string] => typeof entry[1] === 'string'))
        : new Map();
    },
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

  const resident: ResidentEngineSession = {
    splitSection: (at, properties, start) => {
      ensureUndo();
      const split = splitSection(resident, session, at, () => properties, start);
      sectionsTouched = true;
      sectionLayout = null;
      if (!split.closed) {
        const document = partDocument();
        const body = updateFinalSectionProperties(document.package.document, final => ({ ...final, sectionStart: start }));
        session.set_part_metadata(JSON.stringify(partMetadata({ ...document, package: { ...document.package, document: body } })), '[]');
        partsTouched = true;
        sectionLayout = null;
      }
      return split.at;
    },
    residentGeometry: (root, ranges, selection) => JSON.parse(session.resident_geometry_json(root, JSON.stringify({ ranges, selection }))),
    residentRangeRects: (region, partId, ranges) => {
      const flat = session.display_range_rects_batch(region, partId, Float64Array.from(ranges.flat()));
      const batches: DisplayListRect[][] = [];
      for (let at = 0; at < flat.length;) {
        const rects: DisplayListRect[] = [];
        for (let count = flat[at++]!; count > 0; count -= 1, at += 5) {
          rects.push({ pageIndex: flat[at]!, x: flat[at + 1]!, y: flat[at + 2]!, width: flat[at + 3]!, height: flat[at + 4]! });
        }
        batches.push(rects);
      }
      return batches;
    },
    clearCellSelection: () => {
      const at = session.clear_cell_selection();
      return at === undefined ? null : JSON.parse(at) as YrsLoc;
    },
    commentProjection: () => partBase ? projectResidentComments(resident, commentAliases) : null,
    sectionProperties,
    setSectionMargins: (section, margins) => {
      ensureUndo();
      const changed = session.set_section_margins(section, JSON.stringify(margins));
      if (changed) { marginsTouched = true; sectionLayout = null; projectionStories.add('body'); }
      return changed;
    },
    propertyDocument: hostDocument,
    setPropertyDocument: input => {
      sectionLayout = null;
      const document = withSectionNames(input, sectionNames.resident);
      const margins = JSON.parse(session.section_margins_json()) as Record<string, DocxSectionMargins>;
      let changed = false;
      for (const section of sectionProperties().sections.slice(0, -1)) {
        const desired = document.package.document.sections?.find(entry => entry.id === section.section)?.properties;
        if (desired && JSON.stringify(section.properties) !== JSON.stringify(desired)) changed = setSectionMark(section.section, desired) || changed;
      }
      for (const [section, current] of Object.entries(margins)) {
        const desired = section === 'final' ? resolvedFinalSectionProperties(document.package.document)
          : document.package.document.sections?.find(entry => entry.id === section)?.properties;
        const patch = Object.fromEntries(Object.entries(current).filter(([key, value]) => {
          const next = desired?.[key as keyof DocxSectionMargins];
          return next !== undefined && next !== value;
        }).map(([key]) => [key, desired![key as keyof DocxSectionMargins]])) as DocxSectionMargins;
        if (Object.keys(patch).length) changed = resident.setSectionMargins(section, patch) || changed;
      }
      const metadataChanged = resident.setPartDocument(document);
      return changed || metadataChanged;
    },
    controlTargets: () => JSON.parse(session.content_control_targets_json()),
    repeatControlAt: (at, remove) => applyCommandMutation(at.story, () =>
      session.repeat_content_control_at(at.story, at.index, remove)),
    setControlValue: (at, value) => applyCommandMutation(at.story, () =>
      session.set_content_control_value_index(at.story, at.index, JSON.stringify(value))),
    partDocument: hostDocument,
    partGeometry: frame => readPartGeometry(resident, frame),
    partStory: (id) => `hf:${(JSON.parse(session.header_footer_aliases_json()) as Record<string, string>)[id] ?? id}`,
    setPartDocument: (input, removeStories = []) => {
      sectionLayout = null;
      ensureUndo();
      const document = withSectionNames(input, sectionNames.resident);
      let changed = false;
      if (sectionsTouched) {
        const boundaries = resident.paragraphs('body').filter(paragraph => paragraph.properties.sectPr || paragraph.properties._sectionProperties);
        for (const paragraph of boundaries) {
          const desired = document.package.document.sections?.find(section => section.id === paragraph.paraId)?.properties;
          if (!desired || JSON.stringify(desired) === JSON.stringify(paragraph.properties.sectPr ?? paragraph.properties._sectionProperties)) continue;
          changed = setSectionMark(paragraph.paraId, desired) || changed;
        }
      }
      changed = session.set_part_metadata(JSON.stringify(partMetadata(document)), JSON.stringify(removeStories)) || changed;
      if (changed) { partsTouched = true; markProjectionStories(['body', ...session.story_ids()]); }
      return changed;
    },
    proposalEngine,
    geometryReader,
    paragraphs: geometryReader.paragraphs,
    storySegments: (story) => geometryStory(story).segments,
    selectionContext: (range) => JSON.parse(session.selection_context(...rangeArgs(range))),
    cellSelection: () => JSON.parse(session.cell_selection()),
    setCellSelection: (range) => session.set_cell_selection(JSON.stringify(range)),
    insertRow: (at, side, author) => applyCommandMutation(at.story, () => {
      const receipt = JSON.parse(session.insert_row(JSON.stringify(at), side === 'below', author?.name, author?.date));
      markProjectionStories([...receipt.changedStoryIds, ...receipt.createdStoryIds, ...receipt.deletedStoryIds]);
      return receipt;
    }),
    insertTable: (at, rows, columns, author) => applyCommandMutation(at.story, () => {
      const receipt = JSON.parse(session.insert_table(at.story, at.paraId, at.offset, rows, columns, author?.name, author?.date));
      markProjectionStories([...receipt.changedStoryIds, ...receipt.createdStoryIds, ...receipt.deletedStoryIds]);
      return receipt;
    }),
    insertColumn: (at, side) => applyCommandMutation(at.story, () => {
      const receipt = JSON.parse(session.insert_column(JSON.stringify(at), side === 'right'));
      markProjectionStories([...receipt.changedStoryIds, ...receipt.createdStoryIds, ...receipt.deletedStoryIds]);
      return receipt;
    }),
    deleteRow: (range, author) => applyCommandMutation(range.anchor.story, () => {
      const receipt = JSON.parse(session.delete_row(JSON.stringify(range), author?.name, author?.date));
      markProjectionStories([...receipt.changedStoryIds, ...receipt.createdStoryIds, ...receipt.deletedStoryIds]);
      return receipt;
    }),
    deleteColumn: (range) => applyCommandMutation(range.anchor.story, () => {
      const receipt = JSON.parse(session.delete_column(JSON.stringify(range)));
      markProjectionStories([...receipt.changedStoryIds, ...receipt.createdStoryIds, ...receipt.deletedStoryIds]);
      return receipt;
    }),
    deleteTable: (table) => applyCommandMutation(table.story, () => {
      const receipt = JSON.parse(session.delete_table(JSON.stringify(table)));
      markProjectionStories([...receipt.changedStoryIds, ...receipt.createdStoryIds, ...receipt.deletedStoryIds]);
      return receipt;
    }),
    mergeCells: (range) => applyCommandMutation(range.anchor.story, () => {
      const receipt = JSON.parse(session.merge_cells(JSON.stringify(range)));
      markProjectionStories([...receipt.changedStoryIds, ...receipt.createdStoryIds, ...receipt.deletedStoryIds]);
      return receipt;
    }),
    splitCell: (at, rows, columns) => applyCommandMutation(at.story, () => {
      const receipt = JSON.parse(session.split_cell(JSON.stringify(at), rows, columns));
      markProjectionStories([...receipt.changedStoryIds, ...receipt.createdStoryIds, ...receipt.deletedStoryIds]);
      return receipt;
    }),
    setCellShading: (range, color) => applyCommandMutation(range.anchor.story, () => {
      const receipt = JSON.parse(session.set_cell_shading(JSON.stringify(range), color ?? undefined));
      markProjectionStories([...receipt.changedStoryIds, ...receipt.createdStoryIds, ...receipt.deletedStoryIds]);
      return receipt;
    }),
    setCellBorders: (range, borders) => applyCommandMutation(range.anchor.story, () => {
      const receipt = JSON.parse(session.set_cell_borders(JSON.stringify(range), JSON.stringify(borders)));
      markProjectionStories([...receipt.changedStoryIds, ...receipt.createdStoryIds, ...receipt.deletedStoryIds]);
      return receipt;
    }),
    setTableProperties: (table, patch) => applyCommandMutation(table.story, () => {
      const receipt = JSON.parse(session.set_table_properties(JSON.stringify(table), JSON.stringify(patch)));
      markProjectionStories([...receipt.changedStoryIds, ...receipt.createdStoryIds, ...receipt.deletedStoryIds]);
      return receipt;
    }),
    insertTableFragment: (at, table, author) => applyCommandMutation(at.story, () => {
      const receipt = JSON.parse(session.insert_table_fragment(at.story, at.paraId, at.offset, JSON.stringify(table), author?.name, author?.date));
      markProjectionStories([...receipt.changedStoryIds, ...receipt.createdStoryIds, ...receipt.deletedStoryIds]);
      return receipt;
    }),
    validateTableFragment: (table) => { session.validate_table_fragment(JSON.stringify(table)); },
    insertImage: (at, image, author) => applyCommandMutation(at.story, () => JSON.parse(session.insert_image(at.story, at.paraId, at.offset, JSON.stringify(image), author?.name, author?.date))),
    insertPageBreak: at => applyCommandMutation(at.story, () => {
      const target = pageBreakTarget(resident, session, at);
      session.insert_page_break(target.story, target.paraId, target.offset);
      return target;
    }),
    insertSectionBreak: (at, type) => {
      const properties = residentSectionPropertiesAt(resident, at);
      if (!properties) throw new Error('Section break paragraph is unavailable');
      resident.splitSection(at, properties, type);
    },
    tablePayload: (story, index) => {
      const payload = session.table_payload(story, index);
      return payload === undefined ? null : JSON.parse(payload);
    },
    setHyperlink: (range, hyperlink) => applyCommandMutation(range.story, () =>
      session.set_hyperlink(...rangeArgs(range), JSON.stringify(hyperlink))),
    setImageGeometryAt: (at, geometry) => applyCommandMutation(at.story, () =>
      session.set_image_geometry_at(at.story, at.paraId, at.offset, JSON.stringify(geometry))),
    setColumnWidth: (at, width) => applyCommandMutation(at.story, () =>
      JSON.parse(session.set_column_width(JSON.stringify(at), width))),
    beginUndoCapture: ensureUndo,
    setUndoInputTime: (span) => {
      if (span) session.set_undo_input_time(span[0], span[1]);
      else session.clear_undo_input_time();
    },
    addUndoBoundary: () => session.add_undo_boundary(),
    undo: () => {
      const changed = session.undo();
      if (changed) markProjectionStories(session.history_stories());
      return changed;
    },
    redo: () => {
      const changed = session.redo();
      if (changed) markProjectionStories(session.history_stories());
      return changed;
    },
    canUndo: () => session.can_undo(),
    canRedo: () => session.can_redo(),
    insertText: (at, text, author) => applyCommandMutation(at.story, () => JSON.parse(author
      ? session.insert_text(at.story, at.paraId, at.offset, text, author.name, author.date)
      : session.edit_text_resident(at.story, at.paraId, at.offset, at.paraId, at.offset, text))),
    replaceRange: (range, text, author) => applyCommandMutation(range.story, () =>
      JSON.parse(session.replace_range(...rangeArgs(range), text, author?.name, author?.date))),
    deleteRange: (range, author) => applyCommandMutation(range.story, () => JSON.parse(author
      ? session.delete_range(...rangeArgs(range), author.name, author.date)
      : session.edit_text_resident(...rangeArgs(range), undefined))),
    splitParagraph: (at, author) => applyCommandMutation(at.story, () =>
      JSON.parse(session.split_paragraph(at.story, at.paraId, at.offset, author?.name, author?.date))),
    toggleMark: (range, mark) => applyCommandMutation(range.story, () =>
      session.toggle_mark(...rangeArgs(range), JSON.stringify(mark))),
    formatRange: (range, delta) => applyCommandMutation(range.story, () =>
      session.format_range(...rangeArgs(range), JSON.stringify(delta))),
    clearFormatting: (range) => applyCommandMutation(range.story, () =>
      session.clear_formatting(...rangeArgs(range))),
    setParagraphAttrs: (range, attrs, author) => applyCommandMutation(range.story, () =>
      session.set_paragraph_attrs(...rangeArgs(range), JSON.stringify(attrs), author?.name, author?.date)),
    applyParagraphStyle: (range, styleId, author) => applyCommandMutation(range.story, () =>
      session.apply_paragraph_style(...rangeArgs(range), styleId, author?.name, author?.date)),
    addComment: (ranges, author, date, body) => {
      ensureUndo();
      const receipt = JSON.parse(session.add_comment(JSON.stringify(ranges.map(range => ({
        story: range.story, startPara: range.start.paraId, startOffset: range.start.offset,
        endPara: range.end.paraId, endOffset: range.end.offset,
      }))), author, date, JSON.stringify(body)));
      markProjectionStories(ranges.map(range => range.story));
      return receipt;
    },
    setCommentRanges: (id, ranges) => {
      ensureUndo();
      session.set_comment_ranges(id, JSON.stringify(ranges.map(range => ({
        story: range.story, startPara: range.start.paraId, startOffset: range.start.offset,
        endPara: range.end.paraId, endOffset: range.end.offset,
      }))));
      markProjectionStories(ranges.map(range => range.story));
    },
    listComments: () => JSON.parse(session.list_comments()),
    updateComment: (id, patch) => {
      ensureUndo();
      session.update_comment_metadata_json(id, JSON.stringify(patch));
    },
    acceptChange: target => { ensureUndo(); return JSON.parse(session.accept_change(JSON.stringify(target))); },
    rejectChange: target => { ensureUndo(); return JSON.parse(session.reject_change(JSON.stringify(target))); },
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
    readStories: (request) =>
      JSON.parse(session.read_stories_json(JSON.stringify(request))) as DocxReadStoriesResult,
    listStories: () => JSON.parse(session.list_stories_json()) as DocxListStoriesResult,
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
    decodeFont: bytes => session.decode_font(bytes),
    openSourceDocx: (bytes, options) => {
      session.set_media_tokens(options.mediaTokens === true);
      const json = resident.openDocx(bytes, options.digest, options.generation, !options.initialUpdate);
      if (options.initialUpdate) {
        resident.loadState(options.initialUpdate);
        partBase = keySections(partBase!);
        markProjectionStories(session.story_ids());
      }
      return json;
    },
    openDocument: (document, options) => {
      commentAliases.clear();
      geometryStories.clear();
      projectionStories.clear();
      parsedMedia.clear();
      partBase = document;
      parsedFacade = wrapSession(session, clientId);
      const nativeMediaDataUrl = parsedFacade.mediaDataUrl;
      parsedFacade.mediaDataUrl = token => parsedMedia.get(token) ?? nativeMediaDataUrl(token);
      if (options.initialUpdate) {
        resident.loadState(options.initialUpdate);
        partBase = keySections(document);
      } else {
        documentToYrs(parsedFacade, options.mediaTokens ? tokenizeResidentDocument(document, parsedMedia) : document, options);
        partBase = keySections(document);
        session.set_part_metadata(JSON.stringify(partMetadata(partBase)), '[]');
      }
      partsTouched = false;
      marginsTouched = false;
      sectionsTouched = false;
      sectionLayout = null;
      markProjectionStories(session.story_ids());
    },
    saveDocument: async (host, comments, record) => {
      if (!parsedFacade || !partBase) throw new Error('The resident worker holds no parsed document');
      const { saveEditorDocument, mergeDocxHostMetadata } = await import('./editorSave');
      const { noteYrsStoriesDirty, yrsToDocument } = await import('./yrsToDocument');
      const pending = projectionStories.capture();
      noteYrsStoriesDirty(parsedFacade, pending.stories);
      const document = yrsToDocument(parsedFacade, mergeDocxHostMetadata(record.base ?? partBase, host ?? partDocument()));
      const saved = await saveEditorDocument(parsedFacade, document, comments, record);
      document.originalBuffer = saved;
      record.base = document;
      pending.clear();
      return saved;
    },
    openDocx: (bytes, digest, generation, seedStories = true) => {
      commentAliases.clear();
      geometryStories.clear();
      projectionStories.clear();
      const json = resolveHostJsonCommentMedia(
        session.open_docx(bytes, seedStories, generation, digest),
        (token) => (token.startsWith('media:') ? (session.media_data_url(token) ?? null) : null)
      );
      partBase = keySections(decodeDocxHostJson(json, bytes).document);
      partsTouched = false;
      marginsTouched = false;
      sectionsTouched = false;
      sectionLayout = null;
      return json;
    },
    openDocxPreview: (bytes, blocks, paragraphBudget) => {
      commentAliases.clear();
      geometryStories.clear();
      projectionStories.clear();
      const json =
        paragraphBudget === undefined
          ? session.open_docx_preview(bytes, blocks)
          : session.open_docx_preview_with_budget(bytes, blocks, paragraphBudget);
      return json === undefined
        ? null
        : resolveHostJsonCommentMedia(
            json,
            (token) => (token.startsWith('media:') ? (session.media_data_url(token) ?? null) : null)
          );
    },
    mediaDataUrl: (token) => parsedMedia.get(token) ?? session.media_data_url(token) ?? null,
    encodeState: () => session.encode_state(),
    encodePeerMetadata: () => {
      const peerSession = session as typeof session & { encode_peer_metadata?: () => Uint8Array };
      if (!peerSession.encode_peer_metadata) {
        throw new PeerMetadataError('missing-capability', 'Editing wasm cannot encode peer metadata');
      }
      return peerSession.encode_peer_metadata();
    },
    encodeStateAsUpdate: (remoteStateVector) =>
      remoteStateVector === undefined
        ? session.encode_state()
        : session.encode_diff(remoteStateVector.slice()),
    save: async (source, hostJson, host, comments, record, stories) => {
      const { saveResidentDocument } = await import('./residentSave');
      let projectedHost = partsTouched || session.part_metadata_json() !== '{}'
        ? withPartMetadata(host ?? partDocument(), partMetadata(partDocument())) : host;
      if (marginsTouched || sectionsTouched || session.section_margins_json() !== '{}') projectedHost = withSectionSnapshot(projectedHost ?? partDocument(), sectionProperties());
      return saveResidentDocument(
        session, clientId, source, hostJson, projectedHost, comments, record, projectionStories, stories
      );
    },
    markProjectionStories,
    revisionCount: (excluding) =>
      (JSON.parse(session.list_revisions()) as { revisionId: string }[]).filter(
        (revision) => !excluding?.has(revision.revisionId)
      ).length,
    registerFont: (bytes) => session.register_measure_font(bytes),
    registerSubstituteFont: (base, family) =>
      session.register_substitute_measure_font(base, family),
    clearFonts: () => session.clear_measure_fonts(),
    encodeStateVector: () => session.encode_state_vector(),
    hasPendingUpdate: () => session.has_pending_update(),
    measureParagraphJson: (input) => session.measure_paragraph_json(input),
    layoutDocumentJson: (input) => session.layout_document_json(partLayoutInput(input)),
    layoutFontRequirementsJson: (input) => session.layout_font_requirements_json(input),
    layoutDocumentWithRegionsRetainedJson: (input) =>
      session.layout_document_with_regions_retained_json(partLayoutInput(input)),
    layoutDocumentWithRegionsRetainedMeta: (input) =>
      readRetainedLayoutMeta(session.layout_document_with_regions_retained_meta(partLayoutInput(input))),
    retainedLayoutJson: () => session.retained_layout_json(),
    beginRegionLayout: (input) =>
      JSON.parse(session.begin_region_layout(partLayoutInput(input))) as YrsRegionLayoutProgress,
    resumeRegionLayout: (blocks) =>
      JSON.parse(session.resume_region_layout(blocks)) as YrsRegionLayoutProgress,
    setPartialDocument: (partial) => session.set_partial_document(partial),
    layoutDocumentWithRegionsPrefixRetainedJson: (input, pages) =>
      session.layout_document_with_regions_prefix_retained_json(partLayoutInput(input), pages),
    layoutDocumentWithRegionsExactPrefixRetained: (input, pages) => JSON.parse(
      session.layout_document_with_regions_exact_prefix_retained(partLayoutInput(input), pages)
    ) as ResidentExactPrefixLayout,
    layoutDocumentWithRegionsRetained: (input) =>
      session.layout_document_with_regions_retained(partLayoutInput(input)),
    layoutDocumentWithRegionsAfterEdit: (input) =>
      session.layout_document_with_regions_after_edit(partLayoutInput(input)),
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
    applyInput: (text, expectedFrameEpoch) =>
      applyResidentInput(() => session.apply_input(text, expectedFrameEpoch)),
    applyDelete: (direction, expectedFrameEpoch, count = 1) =>
      applyResidentInput(() => session.apply_delete(direction, expectedFrameEpoch, count)),
    residentDeletedUnits: () => session.resident_deleted_units(),
    applyRawOps: (story, ops) => {
      geometryStories.clear();
      const version = session.version();
      try {
        session.apply_raw_ops(story, JSON.stringify(ops));
        projectionStories.add(story);
      } finally {
        if (session.version() !== version) {
          projectionStories.add(story);
          syncStoryRevisions();
          storyRevision += 1;
          storyRevisions.set(story, storyRevision);
        }
      }
    },
    applyInputProfiled: (text, expectedFrameEpoch) => {
      const frame = applyResidentInput(() => session.apply_input_profiled(text, expectedFrameEpoch));
      const profile = JSON.parse(session.apply_input_profile_json()) as YrsEngineApplyProfile;
      return { frame, profile };
    },
    applyDeleteProfiled: (direction, expectedFrameEpoch, count = 1) => {
      const frame = applyResidentInput(() => session.apply_delete_profiled(direction, expectedFrameEpoch, count));
      const profile = JSON.parse(session.apply_input_profile_json()) as YrsEngineApplyProfile;
      return { frame, profile };
    },
    outlineGlyphJson: (fontId, glyphId) => session.outline_glyph_json(fontId, glyphId),
    echoGlyphSetJson: (fontId, size, kern, alphabet) => session.echo_glyph_set_json(fontId, size, kern, alphabet),
    loadMediaSources: (json) => session.load_media_sources(json),
    loadNoteSeparators: (state) => session.load_note_separators(state),
    loadState: (update) => {
      geometryStories.clear();
      sectionLayout = null;
      session.load(update);
    },
    applyUpdate: (update) => {
      sectionLayout = null;
      const before = session.resident_section_revision().split(':')[1];
      const raw = session.apply_update_with_inference(update);
      if (session.resident_section_revision().split(':')[1] !== before) sectionsTouched = true;
      return JSON.parse(raw) as CollaborationTextInsertion | null;
    },
    onUpdate: (listener) => {
      listeners.add(listener);
      ensureObserver();
      return () => {
        listeners.delete(listener);
        if (!destroyed && observing && listeners.size === 0) {
          if (dispatchingUpdates) queueMicrotask(detachObserver);
          else detachObserver();
        }
      };
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
      projectionStories.clear();
      parsedMedia.clear();
      parsedFacade = undefined;
      partBase = undefined;
      if (dispatchingUpdates) queueMicrotask(destroyNative);
      else destroyNative();
    },
  };
  return resident;
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
