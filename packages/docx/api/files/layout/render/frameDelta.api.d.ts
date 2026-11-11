// @betteroffice/docx declaration layout/render/frameDelta.d.ts
// Regenerate: bun scripts/api-surface.ts --write docx
import type { DisplayList, DisplayPage } from './displayList';
export declare const FRAME_DELTA_VERSION = 1;
export declare const FRAME_DELTA_HEADER_BYTES = 80;
export declare const FRAME_DELTA_PAGE_OP_BYTES = 48;
export type FrameDeltaErrorCode = 'invalid-frame' | 'unsupported-version' | 'stale-frame' | 'base-mismatch';
export declare class FrameDeltaError extends Error {
    readonly code: FrameDeltaErrorCode;
    constructor(code: FrameDeltaErrorCode, message: string);
}
export interface FramePageUpsert {
    readonly kind: 'upsert';
    readonly pageIndex: number;
    readonly pageId: bigint;
    readonly fingerprint: bigint;
    /** Zero-copy view into {@link DecodedFrameDelta.bytes}. */
    readonly primitiveIds: BigUint64Array;
    readonly page: DisplayPage;
}
export interface FramePageRemove {
    readonly kind: 'remove';
    readonly pageIndex: number;
    readonly pageId: bigint;
}
export interface FramePageMove {
    readonly kind: 'move';
    readonly pageIndex: number;
    readonly pageId: bigint;
    readonly fingerprint: bigint;
}
export interface FramePrimitivePositionPatch {
    readonly primitiveId: bigint;
    readonly docStart?: number | null;
    readonly docEnd?: number | null;
    readonly fragmentDocStart?: number | null;
    readonly fragmentDocEnd?: number | null;
    readonly inlineWidgetPos?: number | null;
}
export interface FramePagePositionPatch {
    readonly kind: 'patch-positions';
    readonly pageIndex: number;
    readonly pageId: bigint;
    readonly fingerprint: bigint;
    readonly patches: readonly FramePrimitivePositionPatch[];
}
export interface FramePositionShiftRun {
    readonly start: number;
    readonly count: number;
    readonly changedMask: number;
    readonly delta: number;
}
/** A note region note's new backlink anchor; null clears the field. */
export interface FrameNoteAnchor {
    readonly area: number;
    readonly note: number;
    readonly start: number | null;
    readonly end: number | null;
}
export interface FramePagePositionShift {
    readonly kind: 'shift-positions';
    readonly pageIndex: number;
    readonly pageId: bigint;
    readonly fingerprint: bigint;
    readonly runs: readonly FramePositionShiftRun[];
    readonly anchors: readonly FrameNoteAnchor[];
    readonly spanDelta?: number;
}
/** Shift the body positions and fingerprints of consecutive retained pages. */
export interface FramePageShiftRange {
    readonly kind: 'shift-range';
    readonly pageIndex: number;
    readonly pageId: bigint;
    readonly count: number;
    readonly delta: number;
    readonly salt: bigint;
}
/** One applied position shift: its runs, then its note anchors. */
export interface DisplayPageShift {
    readonly runs: readonly FramePositionShiftRun[];
    readonly anchors: readonly FrameNoteAnchor[];
    readonly spanDelta?: number;
}
export type FramePageOperation = FramePageUpsert | FramePageRemove | FramePageMove | FramePagePositionPatch | FramePagePositionShift | FramePageShiftRange;
export interface DecodedFrameDelta {
    readonly protocolVersion: typeof FRAME_DELTA_VERSION;
    readonly full: boolean;
    readonly docEpoch: number;
    readonly layoutEpoch: number;
    readonly frameEpoch: number;
    readonly baseFrameEpoch: number;
    readonly pageCount: number;
    readonly contractVersion?: number;
    readonly operations: readonly FramePageOperation[];
    /** Complete binary frame; safe to transfer when the caller owns its buffer. */
    readonly bytes: Uint8Array;
}
export interface RetainedFramePage {
    readonly pageIndex: number;
    readonly pageId: bigint;
    readonly fingerprint: bigint;
    readonly primitiveIds: BigUint64Array;
    readonly page: DisplayPage;
}
export interface RetainedFrame {
    readonly protocolVersion: typeof FRAME_DELTA_VERSION;
    readonly docEpoch: number;
    readonly layoutEpoch: number;
    readonly frameEpoch: number;
    readonly contractVersion?: number;
    readonly pages: readonly RetainedFramePage[];
    /** IDs whose canvas/mirror/interactive projections must be replaced. */
    readonly damagedPageIds: ReadonlySet<bigint>;
    /** Removed surfaces, retained separately because they are absent from pages. */
    readonly removedPageIds: ReadonlySet<bigint>;
    readonly displayList: DisplayList;
}
/** @internal Find a retained page by its stable identity. */
export declare function retainedFramePageById(frame: RetainedFrame, pageId: bigint): RetainedFramePage | undefined;
/**
 * Decode and fully validate one binary FrameDelta v1 without mutating retained
 * browser state. Page payloads are typed values, not embedded JSON strings.
 */
export declare function decodeFrameDelta(input: Uint8Array | ArrayBuffer): DecodedFrameDelta;
/** @internal Decode in resumable slices without publishing partial state. */
export declare function decodeFrameDeltaSteps(input: Uint8Array | ArrayBuffer): Generator<void, DecodedFrameDelta>;
/** Apply one already-validated frame atomically, rejecting stale generations. */
export declare function applyFrameDelta(previous: RetainedFrame | null, delta: DecodedFrameDelta): RetainedFrame;
/**
 * Apply a frame while reusing position-only page arenas owned exclusively by
 * the caller. Worker/main resident-engine hot paths replace their outer frame
 * atomically and have no concurrent readers; avoiding deep suffix clones keeps
 * geometry shifts from creating periodic multi-megabyte GC stalls.
 */
export declare function applyFrameDeltaOwned(previous: RetainedFrame | null, delta: DecodedFrameDelta): RetainedFrame;
/** @internal Apply with optional general storage for equivalence checks. */
export declare function applyFrameDeltaInternal(previous: RetainedFrame | null, delta: DecodedFrameDelta, reusePositionPages: boolean, usePageIndex?: boolean): RetainedFrame;
/** Position shift revision, omitted from serialization. */
export declare function displayPageRevision(page: DisplayPage): number;
/** @internal Whether `page` has only copied position shifts since `base`. */
export declare function displayPageIsShiftOf(page: DisplayPage, base: DisplayPage): boolean;
/** Counts shifts touching this page's note anchors. */
export declare function displayPageNoteAnchorRevision(page: DisplayPage): number;
/**
 * Ordered shifts advancing the page from `sinceRevision` to its current
 * revision, or null when the bounded log no longer covers that span.
 */
export declare function displayPageShiftsSince(page: DisplayPage, sinceRevision: number): ReadonlyArray<DisplayPageShift> | null;
