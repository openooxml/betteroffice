// @betteroffice/docx declaration yrs/proposalGeometry.d.ts
// Regenerate: bun scripts/api-surface.ts --write docx
import type { YrsLoc, YrsSession, YrsStorySegment } from './index';
import type { DocxTextRange } from './edits';
import type { DocxParagraphAnchor } from './paragraphIdentity';
import { proposalRevisionPreview, type DocxOccurrence, type DocxProposalSnapshot } from './proposals';
import { type YrsPositionOutline } from './yrsPositionProjection';
/** @internal */
export interface AnchorResolutionFailure {
    ok: false;
    failure: {
        code: 'stale-version' | 'missing-target' | 'ambiguous-target' | 'layout-unavailable' | 'unsupported' | 'unknown-proposal';
        message: string;
    };
}
/** @internal */
export type AnchorReader = Pick<YrsSession, 'resolveParagraphAnchor' | 'hasStory' | 'paragraphSpans' | 'storySegments' | 'findText' | 'listRevisions'>;
/** @internal */
export type ProposalGeometryReader = AnchorReader & Pick<YrsSession, 'storyIds' | 'paragraphs' | 'paragraphIdCount' | 'locateParagraph' | 'version'> & {
    proposalRevisions?(ids: readonly string[]): readonly ProposalGeometryRevision[];
    positionOutline?(root: string): YrsPositionOutline | null;
    /** Header and footer alias relationship ids, each mapped to its canonical one. */
    headerFooterAliases?(): ReadonlyMap<string, string>;
};
/** @internal */
export type ProposalGeometryRevision = Pick<ReturnType<AnchorReader['listRevisions']>[number], 'revisionId' | 'kind' | 'story' | 'range'>;
/** @internal */
export type AnchorGeometryTarget = {
    kind: 'proposal';
    id: string;
} | {
    kind: 'revision';
    revisionId: string;
} | {
    kind: 'paragraph';
    paragraph: DocxParagraphAnchor;
} | {
    kind: 'search';
    paragraph: DocxParagraphAnchor;
    text: string;
    occurrence?: DocxOccurrence;
} | {
    kind: 'range';
    version: string;
    range: DocxTextRange;
} | {
    kind: 'drawing';
    paragraph: DocxParagraphAnchor;
    offset: number;
    index?: number;
};
/** @internal */
export interface RawAnchorRange {
    start: YrsLoc;
    end: YrsLoc;
}
/** @internal */
export type AnchorResolution = {
    ok: true;
    ranges: RawAnchorRange[];
    paragraph: YrsLoc;
    drawing?: boolean;
    drawingOffset?: number;
} | AnchorResolutionFailure;
/** @internal */
export declare function anchorFailure(code: AnchorResolutionFailure['failure']['code'], message: string): AnchorResolutionFailure;
/** @internal */
export declare function isBodyStory(story: string): boolean;
/**
 * The story whose display positions `story` uses: `body` for the body and its cells, `hf:<rId>`
 * for a header or footer and its cells, else null. @internal
 */
export declare function anchorDisplayRoot(story: string): string | null;
/** @internal The block embeds `paraId` opens with, which its Loc offsets count before its first inline unit. */
export declare function paragraphInlineStart(session: AnchorReader, version: string, story: string, paraId: string): number;
/** @internal */
export declare function textRangeToRaw(segments: readonly YrsStorySegment[], range: DocxTextRange): {
    ok: true;
    range: RawAnchorRange;
} | AnchorResolutionFailure;
/** @internal */
export declare function hiddenRanges(session: AnchorReader, version: string, snapshot: DocxProposalSnapshot | null, revisions?: readonly ProposalGeometryRevision[]): RawAnchorRange[];
/** @internal */
export declare function hiddenRangesForPreview(session: AnchorReader, version: string, preview: ReturnType<typeof proposalRevisionPreview>, revisions?: readonly ProposalGeometryRevision[]): RawAnchorRange[];
/** @internal */
export declare function resolveAnchorTarget(session: AnchorReader, target: AnchorGeometryTarget, version: string, snapshot?: DocxProposalSnapshot | null, proposalRevisions?: readonly ProposalGeometryRevision[]): AnchorResolution;
/** @internal */
export interface ProposalDisplayTarget {
    ranges: {
        from: number;
        to: number;
    }[];
    paragraph: number | null;
    /** The ranges are one drawing atom; answer with its drawn frame box. */
    drawing?: boolean;
    /** A header or footer target's story, whose positions these are; absent for the body. */
    root?: string;
    /** With `root`, that story's hidden ranges. */
    hidden?: {
        from: number;
        to: number;
    }[];
    /** With `root`, the relationship ids whose bands show it, when the reader knows its aliases. */
    rIds?: string[];
}
/** @internal */
export type ProposalGeometryTarget = ({
    ok: true;
} & ProposalDisplayTarget) | AnchorResolutionFailure;
/** @internal */
export interface ProposalGeometryMirror {
    version: string;
    previewVersion: number;
    proposals: string;
    targets: Record<string, ProposalGeometryTarget>;
    navigationTargets?: Record<string, ReturnType<typeof resolveNavigationTarget>>;
    hidden: {
        from: number;
        to: number;
    }[];
}
/** @internal */
export declare function proposalSetIdentity(snapshot: DocxProposalSnapshot): string;
/** @internal */
export declare function computeProposalGeometryMirror(reader: ProposalGeometryReader, snapshot: DocxProposalSnapshot, includeNavigationTargets?: boolean): ProposalGeometryMirror;
/**
 * @internal Display geometry of non-proposal targets, each in the story it is painted from;
 * `hidden` gives the raw hidden ranges a header or footer answer carries for its story.
 */
export declare function computeAnchorTargetGeometry(reader: ProposalGeometryReader, targets: readonly Exclude<AnchorGeometryTarget, {
    kind: 'proposal';
}>[], hidden?: () => readonly RawAnchorRange[]): ProposalGeometryTarget[];
/** @internal {@link computeAnchorTargetGeometry} of targets already resolved at `version`. */
export declare function resolvedTargetGeometry(reader: ProposalGeometryReader, version: string, hidden?: () => readonly RawAnchorRange[]): (resolved: AnchorResolution) => ProposalGeometryTarget;
/** @internal */
export declare function resolveMirroredNavigationTarget(mirror: ProposalGeometryMirror | null, snapshot: DocxProposalSnapshot, story: string, paraId: string): ReturnType<typeof resolveNavigationTarget> | null;
/** @internal */
export declare function resolveNavigationTarget(reader: ProposalGeometryReader, story: string, paraId: string): {
    loc: YrsLoc;
    position: number;
} | 'missing-target' | 'ambiguous-target' | 'unsupported';
