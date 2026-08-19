/**
 * The one non-body part the editor can have open. Header/footer bands and
 * notes are alternatives, never neighbours: opening one closes the other, so
 * they share a single value rather than a flag each.
 */

import type {
  DisplayListImageRegion,
  DisplayListRegionHit,
} from '@betteroffice/docx/layout/render';

/**
 * What the user opened, before the package resolves it. A band names the
 * variant to look up (the first-page one wins on a `titlePg` page 1) and the
 * page it was opened from; a note is already addressed by its id.
 */
export type PartEditTarget =
  | { kind: 'header' | 'footer'; isFirstPage: boolean; pageIndex: number }
  | { kind: 'footnote' | 'endnote'; noteId: number };

/**
 * The resolved part: a band with the relationship id of its part (null until
 * a newly materialised one registers), or a note with its id.
 */
export type PartEdit =
  | { kind: 'header' | 'footer'; rId: string | null }
  | { kind: 'footnote' | 'endnote'; noteId: number };

/** The note half of {@link PartEdit} — one note, on one page, by id. */
export type NoteEdit = Extract<PartEdit, { noteId: number }>;

function isNote(part: PartEdit): part is NoteEdit {
  return part.kind === 'footnote' || part.kind === 'endnote';
}

/** yrs root story the open part types into; the body when none is typeable. */
export function partEditStory(part: PartEdit | null): string {
  if (!part) return 'body';
  if (isNote(part)) return `${part.kind === 'footnote' ? 'fn' : 'en'}:${part.noteId}`;
  return part.rId ? `hf:${part.rId}` : 'body';
}

/** Whether a hit lands in the open part — what separates typing from leaving. */
export function hitBelongsToPart(
  part: PartEdit | null,
  hit: DisplayListRegionHit | null
): boolean {
  if (!part || !hit || hit.region !== part.kind) return false;
  return isNote(part) ? hit.noteId === part.noteId : true;
}

/**
 * Display-list image scope of the open part. Null for a note: the list indexes
 * images by band, so the only answer available there would be a body image
 * from behind the note area.
 */
export function partImageRegion(part: PartEdit | null): DisplayListImageRegion | null {
  if (!part) return 'body';
  return isNote(part) ? null : part.kind;
}

export function isNoteAreaHit(
  hit: DisplayListRegionHit | null
): hit is DisplayListRegionHit & { region: 'footnote' | 'endnote' } {
  return hit?.region === 'footnote' || hit?.region === 'endnote';
}

/** The note a hit opens for editing, or null when it names none. */
export function noteEditFromHit(hit: DisplayListRegionHit | null): NoteEdit | null {
  if (!isNoteAreaHit(hit)) return null;
  return hit.noteId == null ? null : { kind: hit.region, noteId: hit.noteId };
}
