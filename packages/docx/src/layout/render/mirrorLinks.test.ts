import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { afterAll, expect, test } from 'bun:test';
import type { DisplayPage, DisplayPrimitive } from './displayList';
import {
  buildMirrorPage,
  buildMirrorPageLinks,
  mirrorPageHasTabStops,
  updateMirrorPage,
} from './mirrorDom';

const ownsDom = !GlobalRegistrator.isRegistered;
if (ownsDom) GlobalRegistrator.register();
afterAll(async () => {
  if (ownsDom) await GlobalRegistrator.unregister();
});

const text = (extra: Partial<DisplayPrimitive>): DisplayPrimitive =>
  ({ kind: 'text', x: 10, y: 10, text: 'link', font: '11px sans-serif', color: '#000', ...extra }) as DisplayPrimitive;
const image = (href: string): DisplayPrimitive =>
  ({ kind: 'image', x: 0, y: 0, w: 10, h: 10, relId: 'rId1', src: '', href }) as unknown as DisplayPrimitive;
const page = (primitives: DisplayPrimitive[]): DisplayPage => ({ pageIndex: 0, width: 100, height: 100, primitives });

const HOSTILE = [
  'javascript:alert(1)',
  'JavaScript:alert(1)',
  ' javascript:alert(1)',
  'java\tscript:alert(1)',
  '\x01javascript:alert(1)',
  'data:text/html,<script>alert(1)</script>',
  'vbscript:msgbox(1)',
  'file:///etc/passwd',
];

const hrefs = (root: HTMLElement) =>
  Array.from(root.querySelectorAll('*'), (element) => [
    element.getAttribute('href'),
    (element as HTMLElement).dataset.href,
  ]).flat().filter((value): value is string => typeof value === 'string');

test('a hostile hyperlink target never reaches the mirror as a link', () => {
  for (const href of HOSTILE) {
    for (const primitive of [text({ href }), text({ href, noteRef: { id: 4 } }), image(href)]) {
      const target = page([primitive]);
      const mirror = buildMirrorPage(target);
      expect(hrefs(mirror).filter((value) => !value.startsWith('#oox-'))).toEqual([]);
      expect(mirror.querySelector('a[target]')).toBeNull();
      expect(hrefs(buildMirrorPageLinks(target))).not.toContain(href);
    }
    expect(mirrorPageHasTabStops(page([text({ href })]))).toBe(false);
    expect(mirrorPageHasTabStops(page([image(href)]))).toBe(false);
  }
});

test('a note reference whose hyperlink is hostile stays a note reference', () => {
  const mirror = buildMirrorPage(page([text({ href: 'javascript:alert(1)', noteRef: { id: 4 } })]));
  const anchor = mirror.querySelector('a');
  expect(anchor?.getAttribute('href')).toBe('#oox-footnote-4');
  expect(anchor?.getAttribute('role')).toBe('doc-noteref');
});

test('safe hyperlinks keep their target', () => {
  for (const href of ['https://example.com/a', 'mailto:a@example.com', '#bookmark', 'relative/doc.docx']) {
    const anchor = buildMirrorPage(page([text({ href })])).querySelector('a');
    expect(anchor?.getAttribute('href')).toBe(href);
    expect(anchor?.dataset.href).toBe(href);
  }
});

test('an update that turns a safe link hostile drops the link', () => {
  const mirror = buildMirrorPage(page([text({ href: 'https://example.com/a', blockId: 1 })]));
  expect(mirror.querySelector('a[href]')).not.toBeNull();
  updateMirrorPage(mirror, page([text({ href: 'javascript:alert(1)', blockId: 1 })]));
  expect(hrefs(mirror)).toEqual([]);
});
