// @betteroffice/docx declaration yrs/residentEngineSession.d.ts
// Regenerate: bun scripts/api-surface.ts --write docx
import type { YrsReplaceReceipt, YrsSession, YrsStorySegment, YrsLoc } from './index';
import type { DocxParagraphIdentitySnapshot } from './paragraphIdentity';
import type { DocxSidebarReader } from './sidebarReads';
import type { ProposalGeometryReader, ProposalGeometryRevision } from './proposalGeometry';
import type { DocxProposalSession } from './proposals';
import type { DocxPageExportOptions } from './pagedExport';
import type { ResidentSaveRecord } from './residentSave';
import type { Comment } from '../types/content';
import type { Document, SectionProperties } from '../types/document';
import type { DocxCommentPatch } from './residentReview';
import type { DocxSvgMedia } from './residentCommandProtocol';
import { type ResidentCommentRecoveryState, type ResidentCommentProjection } from './residentCommentProjection';
import { type ResidentPartGeometry } from './residentParts';
import type { RetainedFrame } from '../layout/render/frameDelta';
import type { DisplayListHitRegion, DisplayListRect } from '../layout/render/displayListQueries';
import type { ResidentControlTarget, ResidentControlPosition } from './residentControls';
import { type DocxSectionMargins, type ResidentSectionSnapshot } from './residentSections';
/** @internal */
export interface ResidentExactPrefixLayout {
    provisional: boolean;
    pages: number;
    /** The widest page of the whole document, while only a prefix is laid out. */
    documentWidth?: number;
}
export type ResidentEngineSession = Pick<YrsSession, 'clientId' | 'applyDelete' | 'applyDeleteProfiled' | 'applyInput' | 'applyInputProfiled' | 'applyRawOps' | 'applyUpdate' | 'applyUpdateChunks' | 'paragraphs' | 'selectionContext' | 'storySegments' | 'cellSelection' | 'setCellSelection' | 'insertRow' | 'insertColumn' | 'deleteRow' | 'deleteColumn' | 'deleteTable' | 'mergeCells' | 'splitCell' | 'setCellShading' | 'setCellBorders' | 'setTableProperties' | 'insertTable' | 'insertTableFragment' | 'validateTableFragment' | 'insertImage' | 'insertPageBreak' | 'insertSectionBreak' | 'tablePayload' | 'setHyperlink' | 'setImageGeometryAt' | 'setColumnWidth' | 'setRowHeight' | 'insertText' | 'insertLineBreak' | 'replaceRange' | 'deleteRange' | 'splitParagraph' | 'toggleMark' | 'formatRange' | 'clearFormatting' | 'setParagraphAttrs' | 'applyParagraphStyle' | 'beginUndoCapture' | 'addUndoBoundary' | 'canUndo' | 'canRedo' | 'undo' | 'redo' | 'addComment' | 'setCommentRanges' | 'listComments' | 'acceptChange' | 'rejectChange' | 'beginRegionLayout' | 'buildDisplayListFrame' | 'buildDisplayPagesFrame' | 'releaseDisplayPagesFrame' | 'clearFonts' | 'destroy' | 'encodeSelection' | 'encodeStateAsUpdate' | 'encodeStateAsUpdateChunks' | 'encodeStateVector' | 'findContentControls' | 'listContentControls' | 'layoutDocumentJson' | 'layoutFontRequirementsJson' | 'layoutDocumentWithRegionsRetainedJson' | 'layoutDocumentWithRegionsRetainedMeta' | 'retainedLayoutJson' | 'loadMediaSources' | 'loadNoteSeparators' | 'loadState' | 'setPartialDocument' | 'measureParagraphJson' | 'onUpdate' | 'outlineGlyphJson' | 'registerFont' | 'updateFontOutlines' | 'registerSubstituteFont' | 'residentCaretSnapshot' | 'residentDeletedUnits' | 'resumeRegionLayout' | 'selection' | 'selectionText' | 'resolveComment' | 'searchText' | 'encodeStickyPosition' | 'resolveStickyPosition' | 'setDisplayRetainBuiltPages' | 'setDisplayWindow' | 'setSelection' | 'readStories' | 'listStories' | 'storiesChangedSince' | 'yrsBlocksForStory'> & {
    /** Moves an inline embed in one transaction. Internal: concurrent moves would duplicate it, so collaboration never reaches it. */
    moveEmbed(at: YrsLoc, to: YrsLoc): YrsReplaceReceipt;
    setUndoInputTime(span: readonly [first: number, last: number] | null): void;
    splitSection(at: YrsLoc, properties: SectionProperties, start: NonNullable<SectionProperties['sectionStart']>): YrsLoc;
    /**
     * Display ranges of the paragraphs, tables and block controls of `root` that `ranges` touch.
     * A paragraph's `length` counts story units; its inline `controls` give each node's story
     * offset and display range. `selection` maps both ends to display positions.
     */
    storyGeometry(root: string, ranges: Array<[number, number]>, selection?: [YrsLoc, YrsLoc]): {
        paragraphs: Array<{
            story: string;
            paraId: string;
            index: number;
            from: number;
            to: number;
            leading: number;
            length: number;
            controls?: Array<{
                offset: number;
                from: number;
                to: number;
            }>;
        }>;
        tables: Array<{
            story: string;
            tableIndex: number;
            from: number;
            to: number;
            at: YrsLoc;
            cells: Array<{
                story: string;
                paraId: string;
                row: number;
                column: number;
                rowspan: number;
                colspan: number;
                from: number;
                to: number;
            }>;
        }>;
        controls: Array<{
            story: string;
            contentStory: string;
            from: number;
            to: number;
        }>;
        selection: [number, number] | null;
    };
    /** Rects of the resident display list for each `[from, to)` range of one region, in order. */
    storyRangeRects(region: DisplayListHitRegion, partId: string, ranges: readonly (readonly [number, number])[]): DisplayListRect[][];
    clearCellSelection(): YrsLoc | null;
    controlTargets(): ResidentControlTarget[];
    repeatControlAt(at: ResidentControlPosition, remove: boolean): boolean;
    setControlValue(at: ResidentControlPosition, value: import('./index').YrsContentControlValue): void;
    partDocument(): Document;
    setPartDocument(document: Document, removeStories?: string[]): boolean;
    partStory(relationshipId: string): string;
    partGeometry(frame: RetainedFrame): ResidentPartGeometry[];
    propertyDocument(): Document;
    /** The `w:next` style a paragraph break at `at` gives the new paragraph, or null to keep the split's own. */
    nextParagraphStyle(at: YrsLoc): string | null;
    setPropertyDocument(document: Document): boolean;
    sectionProperties(): ResidentSectionSnapshot;
    setSectionMargins(section: string, margins: DocxSectionMargins): boolean;
    updateComment(id: string, patch: DocxCommentPatch): void;
    commentProjection(fresh?: boolean): ResidentCommentProjection | null;
    /** Seeds the opened package's comment metadata, unless the session joined from a peer's state. */
    initializeComments(recovery?: ResidentCommentRecoveryState): void;
    commentRecoveryState?(): ResidentCommentRecoveryState;
    mediaDataUrl(token: string): string | null;
    /** The pictures whose display copy is SVG, which a worker cannot decode. */
    svgMedia?(): DocxSvgMedia[];
    /** @internal Typing echo model of one font at one shaping size; see `echo_glyph_set_json`. */
    echoGlyphSetJson(fontId: number, size: number, kern: boolean, alphabet: string): string;
    /** @internal */
    proposalEngine: DocxProposalSession & Pick<YrsSession, 'validateEdits'>;
    /** @internal */
    pageTextsJson(pages: Uint32Array): string;
    /** @internal */
    geometryReader: ProposalGeometryReader & DocxSidebarReader;
    /** @internal The segments of the story's paragraphs at `indices`, each ending with its pilcrow. */
    paragraphSegments(story: string, indices: readonly number[]): YrsStorySegment[][];
    /** @internal Index of the first paragraph of `story` with `paraId`, in `paragraphSpans` order, or -1. */
    paragraphUnit(story: string, paraId: string): number;
    /** @internal Ids of the tracked changes of `at.story` whose range holds `at`, in `listRevisions` order. */
    revisionIdsAt(at: YrsLoc): string[];
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
    openSourceDocx(bytes: Uint8Array, options: {
        digest?: string;
        generation?: string;
        mediaTokens?: boolean;
        initialUpdate?: Uint8Array;
        state?: Uint8Array;
    }): string;
    decodeFont(bytes: Uint8Array): Uint8Array;
    takeBaseLatinMiss?(): boolean;
    openDocument(document: Document, options: {
        generation?: string;
        mediaTokens?: boolean;
        initialUpdate?: Uint8Array;
    }): void;
    saveDocument(host: Document | undefined, comments: Comment[], record: ResidentSaveRecord): Promise<ArrayBuffer>;
    /** Caches save seed states while unchanged, within a paragraph and inflated package budget. */
    primeSaveSeedStates(maxParagraphs: number, maxInflatedBytes: number): boolean;
    /** One slice of {@link primeSaveSeedStates} digesting about `maxBytes` bytes: 0 known, 1 more to do, 2 declined. */
    primeSaveSeedStatesSlice(maxParagraphs: number, maxInflatedBytes: number, maxBytes: number): number;
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
    save(source: Uint8Array, hostJson: string, host: Document | undefined, comments: Comment[], record: ResidentSaveRecord, stories?: readonly string[], yieldSave?: () => Promise<void>): Promise<ArrayBuffer>;
};
export declare function createResidentEngineSession(heapLimitBytes?: number, clientId?: number): Promise<ResidentEngineSession>;
export declare function createResidentSaveWarmupSession(): ResidentEngineSession;
export declare function readLegacyProposalRevisions(reader: ProposalGeometryReader, ids: readonly string[]): ProposalGeometryRevision[];
