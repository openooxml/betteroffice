import { describe, expect, test } from 'bun:test';
import {
  computeViewportAnchoredScrollTop,
  mergeLayoutUpdateOrigin,
  shouldScrollCaretIntoView,
  type ViewportAnchorSnapshot,
} from './viewportAnchoring';

const anchor: ViewportAnchorSnapshot = {
  viewportOffset: 140,
  scrollTopSnapshot: 600,
};

describe('computeViewportAnchoredScrollTop', () => {
  test('compensates for content height added above the viewport', () => {
    expect(computeViewportAnchoredScrollTop(anchor, 840, 2_000)).toBe(700);
  });

  test('does not move for an edit below the viewport', () => {
    expect(computeViewportAnchoredScrollTop(anchor, 740, 2_000)).toBe(600);
  });

  test('keeps the prior position when an edit spanning the viewport removes the anchor', () => {
    expect(computeViewportAnchoredScrollTop(anchor, null, 2_000)).toBe(600);
  });
});

describe('shouldScrollCaretIntoView', () => {
  test('preserves local caret scrolling', () => {
    expect(shouldScrollCaretIntoView('local', false)).toBe(true);
  });

  test('does not scroll to a sticky caret after a remote relayout', () => {
    expect(shouldScrollCaretIntoView('remote', false)).toBe(false);
  });

  test('allows a local selection action after a remote relayout', () => {
    expect(shouldScrollCaretIntoView('remote', true)).toBe(true);
  });
});

describe('mergeLayoutUpdateOrigin', () => {
  test('preserves a remote origin through coalescing', () => {
    expect(mergeLayoutUpdateOrigin(null, 'remote')).toBe('remote');
    expect(mergeLayoutUpdateOrigin('remote', 'remote')).toBe('remote');
  });

  test('lets a newer local update supersede a pending remote restore', () => {
    expect(mergeLayoutUpdateOrigin('remote', 'local')).toBe('local');
  });
});
