// @betteroffice/docx declaration yrs/reservations.d.ts
// Regenerate: bun scripts/api-surface.ts --write docx
import type { DocxTextRange } from './edits';
import type { DocxParagraphAnchor } from './paragraphIdentity';
import type { DocxOccurrence } from './proposals';
import { type AnchorReader } from './proposalGeometry';
/** Text a reservation anchors to: a range read at `version`, or matches of `text` in a paragraph. */
export type DocxReservationAnchor = {
    kind: 'range';
    version: string;
    range: DocxTextRange;
} | {
    kind: 'search';
    paragraph: DocxParagraphAnchor;
    text: string;
    occurrence?: DocxOccurrence;
};
/** A whole paragraph, mark included. */
export interface DocxReservationParagraph {
    kind: 'paragraph';
    paragraph: DocxParagraphAnchor;
}
/**
 * Layout space reserved without changing the document, in CSS pixels at zoom 1. An `inline` box
 * sits at its anchor's start on the baseline, takes no document positions and is never split;
 * `hideAnchor` also hides the anchored text. A `block` sits beside the paragraph its anchor starts
 * in, spans the column unless `width` is set and is clipped to the column. `hide` removes its
 * anchor from render only.
 */
export type DocxReservation = {
    id: string;
    kind: 'inline';
    anchor: DocxReservationAnchor;
    width: number;
    height: number;
    hideAnchor?: boolean;
} | {
    id: string;
    kind: 'block';
    anchor: DocxReservationAnchor | DocxReservationParagraph;
    side: 'before' | 'after';
    height: number;
    width?: number;
} | {
    id: string;
    kind: 'hide';
    anchor: DocxReservationAnchor | DocxReservationParagraph;
};
/** @internal Every owner's reservations, computed at `version`. */
export interface DocxReservationsPayload {
    version: string;
    /** Names this payload in the frames laid out with it. */
    key?: string;
    sets: readonly {
        owner: string;
        version?: string;
        items: readonly DocxReservation[];
    }[];
}
/**
 * @internal One reservation resolved to units counted from its paragraph's first inline unit, the
 * mark at the inline length.
 */
export interface DocxResolvedReservation {
    owner: string;
    id: string;
    story: string;
    paraId: string;
    start: number;
    end: number;
    kind: 'inline' | 'block' | 'hide';
    width?: number;
    height?: number;
    side?: 'before' | 'after';
    hide?: boolean;
}
/** @internal Resolved reservations by key, as `renderEnv.reservations` carries them. */
export type DocxResolvedReservations = Record<string, DocxResolvedReservation>;
/** @internal Retains applied entries for the payload's owners. */
export declare function retainReservationOwners(payload: DocxReservationsPayload, previous: DocxResolvedReservations | null): DocxResolvedReservations;
/** @internal Whether `value` is a well-formed payload: ids unique per owner, owners unique. */
export declare function validReservationsPayload(value: unknown): value is DocxReservationsPayload;
/**
 * @internal Resolves a payload against `reader` at `readerVersion`, which a range anchor at the
 * payload's version is read at. Anchors that do not resolve are left out.
 */
export declare function resolveReservations(reader: AnchorReader, payload: DocxReservationsPayload, readerVersion?: string, previous?: DocxResolvedReservations | null): DocxResolvedReservations;
