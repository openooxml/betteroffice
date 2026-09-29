import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { afterAll, expect, test } from 'bun:test';
import type { DisplayPage, DisplayPrimitive } from './displayList';
import { buildInteractiveOverlayPage, interactiveOverlayHasTabStops } from './interactiveOverlay';
import { buildMirrorPage, displayPageHoldsMirrorId, mirrorPageHasTabStops } from './mirrorDom';

const ownsDom = !GlobalRegistrator.isRegistered;
if (ownsDom) GlobalRegistrator.register();
afterAll(async () => {
  if (ownsDom) await GlobalRegistrator.unregister();
});

const text = (extra: Partial<DisplayPrimitive>): DisplayPrimitive =>
  ({ kind: 'text', x: 10, y: 10, text: '1', font: '11px sans-serif', color: '#000', ...extra }) as DisplayPrimitive;

const pages: DisplayPage[] = [
  { pageIndex: 0, width: 100, height: 100, primitives: [text({})] },
  {
    pageIndex: 1,
    width: 100,
    height: 100,
    primitives: [text({ noteRef: { id: 7, kind: 'endnote' } })],
    noteAreas: [
      {
        kind: 'footnote',
        y: 50,
        height: 20,
        noteIds: [3],
        notes: [{ id: 3, label: '3' }],
        primitives: [text({ groupId: 'footnote-3' })],
      },
    ],
  },
  { pageIndex: 2, width: 100, height: 100, primitives: [text({ href: 'https://example.com' })] },
  {
    pageIndex: 3,
    width: 100,
    height: 100,
    primitives: [text({ inlineSdtWidget: { kind: 'checkbox', groupId: 'g', pos: 1 } })],
  },
  {
    pageIndex: 4,
    width: 100,
    height: 100,
    primitives: [
      text({ inlineSdtWidget: { kind: 'checkbox', groupId: 'l', pos: 1, locked: true } }),
    ],
    noteAreas: [
      {
        kind: 'endnote',
        y: 50,
        height: 20,
        noteIds: [9],
        primitives: [text({ groupId: 'endnote-9' })],
      },
    ],
  },
];

test('a page reports the note ids and tab stops its chrome builds', () => {
  for (const page of pages) {
    const mirror = buildMirrorPage(page);
    const overlay = buildInteractiveOverlayPage(page);
    const ids = Array.from(mirror.querySelectorAll('[id]'), (element) => element.id);
    for (const id of ids) expect(displayPageHoldsMirrorId(page, id)).toBe(true);
    for (const id of ['oox-footnote-3', 'oox-endnote-7', 'oox-noteref-endnote-7']) {
      if (!ids.includes(id)) expect(displayPageHoldsMirrorId(page, id)).toBe(false);
    }
    const stops = (root: HTMLElement) =>
      Array.from(root.querySelectorAll<HTMLButtonElement>('a[href], button')).filter(
        (element) => !element.disabled
      ).length;
    expect(mirrorPageHasTabStops(page)).toBe(stops(mirror) > 0);
    expect(interactiveOverlayHasTabStops(page)).toBe(stops(overlay) > 0);
  }
});
