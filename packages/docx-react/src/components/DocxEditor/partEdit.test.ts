import { describe, expect, test } from 'bun:test';
import type { DisplayListRegionHit } from '@betteroffice/docx/layout/render';
import {
  hitBelongsToPart,
  isNoteAreaHit,
  noteEditFromHit,
  partEditStory,
  partImageRegion,
  type PartEdit,
} from './partEdit';

const hit = (over: Partial<DisplayListRegionHit>): DisplayListRegionHit => ({
  region: 'body',
  pos: 12,
  ...over,
});

describe('isNoteAreaHit', () => {
  test('recognises both note regions without requiring an attributed note', () => {
    expect(isNoteAreaHit(hit({ region: 'footnote', noteId: undefined, pos: null }))).toBe(true);
    expect(isNoteAreaHit(hit({ region: 'endnote', noteId: undefined, pos: null }))).toBe(true);
    expect(isNoteAreaHit(hit({ region: 'body' }))).toBe(false);
    expect(isNoteAreaHit(null)).toBe(false);
  });
});

describe('partEditStory', () => {
  test('the body owns typing while no part is open', () => {
    expect(partEditStory(null)).toBe('body');
  });

  test('a header or footer addresses its relationship story', () => {
    expect(partEditStory({ kind: 'header', rId: 'rId7' })).toBe('hf:rId7');
    expect(partEditStory({ kind: 'footer', rId: 'rId8' })).toBe('hf:rId8');
  });

  // The band is opened before its part exists, and an unresolved one has no
  // story to type into — the body keeps the caret until the rId lands.
  test('a header opened without a resolved part still addresses the body', () => {
    expect(partEditStory({ kind: 'header', rId: null })).toBe('body');
  });

  test('a note addresses its own numbered story', () => {
    expect(partEditStory({ kind: 'footnote', noteId: 2 })).toBe('fn:2');
    expect(partEditStory({ kind: 'endnote', noteId: 2 })).toBe('en:2');
  });
});

describe('hitBelongsToPart', () => {
  test('nothing belongs to the body', () => {
    expect(hitBelongsToPart(null, hit({ region: 'footnote', noteId: 1 }))).toBe(false);
  });

  test('an open band takes its own region and no other', () => {
    const header: PartEdit = { kind: 'header', rId: 'rId7' };
    expect(hitBelongsToPart(header, hit({ region: 'header' }))).toBe(true);
    expect(hitBelongsToPart(header, hit({ region: 'footer' }))).toBe(false);
    expect(hitBelongsToPart(header, hit({ region: 'body' }))).toBe(false);
    expect(hitBelongsToPart(header, null)).toBe(false);
  });

  // Notes of one kind share a region, so the id is what separates them.
  test('an open note takes only its own id', () => {
    const note: PartEdit = { kind: 'footnote', noteId: 2 };
    expect(hitBelongsToPart(note, hit({ region: 'footnote', noteId: 2 }))).toBe(true);
    expect(hitBelongsToPart(note, hit({ region: 'footnote', noteId: 3 }))).toBe(false);
    expect(hitBelongsToPart(note, hit({ region: 'endnote', noteId: 2 }))).toBe(false);
    expect(hitBelongsToPart(note, hit({ region: 'footnote' }))).toBe(false);
  });
});

describe('partImageRegion', () => {
  test('the body and the bands index their own images', () => {
    expect(partImageRegion(null)).toBe('body');
    expect(partImageRegion({ kind: 'header', rId: 'rId7' })).toBe('header');
    expect(partImageRegion({ kind: 'footer', rId: 'rId8' })).toBe('footer');
  });

  // The display list indexes images by band, never by note, so a note has no
  // scope to look one up in — and answering `body` would hand back an image
  // from behind the note area.
  test('an open note has no image scope', () => {
    expect(partImageRegion({ kind: 'footnote', noteId: 2 })).toBeNull();
    expect(partImageRegion({ kind: 'endnote', noteId: 2 })).toBeNull();
  });
});

describe('noteEditFromHit', () => {
  test('a note hit names the note it opens', () => {
    expect(noteEditFromHit(hit({ region: 'footnote', noteId: 4 }))).toEqual({
      kind: 'footnote',
      noteId: 4,
    });
    expect(noteEditFromHit(hit({ region: 'endnote', noteId: 4 }))).toEqual({
      kind: 'endnote',
      noteId: 4,
    });
  });

  test('a hit that names no note opens none', () => {
    expect(noteEditFromHit(hit({ region: 'body' }))).toBeNull();
    expect(noteEditFromHit(hit({ region: 'header' }))).toBeNull();
    expect(noteEditFromHit(hit({ region: 'footnote' }))).toBeNull();
    expect(noteEditFromHit(null)).toBeNull();
  });
});
