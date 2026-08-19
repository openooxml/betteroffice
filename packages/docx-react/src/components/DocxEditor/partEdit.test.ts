import { describe, expect, test } from 'bun:test';
import type { DisplayListRegionHit } from '@betteroffice/docx/layout/render';
import { hitBelongsToPart, partEditStory, partImageRegion, type PartEdit } from './partEdit';

const hit = (over: Partial<DisplayListRegionHit>): DisplayListRegionHit => ({
  region: 'body',
  pos: 12,
  ...over,
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
});

describe('hitBelongsToPart', () => {
  test('nothing belongs to the body', () => {
    expect(hitBelongsToPart(null, hit({ region: 'header' }))).toBe(false);
  });

  test('an open band takes its own region and no other', () => {
    const header: PartEdit = { kind: 'header', rId: 'rId7' };
    expect(hitBelongsToPart(header, hit({ region: 'header' }))).toBe(true);
    expect(hitBelongsToPart(header, hit({ region: 'footer' }))).toBe(false);
    expect(hitBelongsToPart(header, hit({ region: 'body' }))).toBe(false);
    expect(hitBelongsToPart(header, null)).toBe(false);
  });
});

describe('partImageRegion', () => {
  test('the body and the bands index their own images', () => {
    expect(partImageRegion(null)).toBe('body');
    expect(partImageRegion({ kind: 'header', rId: 'rId7' })).toBe('header');
    expect(partImageRegion({ kind: 'footer', rId: 'rId8' })).toBe('footer');
  });
});
