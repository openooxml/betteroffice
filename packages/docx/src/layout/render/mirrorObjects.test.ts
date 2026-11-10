import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { afterAll, expect, test } from 'bun:test';
import type { DisplayPage, DisplayPrimitive } from './displayList';
import { buildMirrorPage, buildMirrorPageText, updateMirrorPage } from './mirrorDom';

const ownsDom = !GlobalRegistrator.isRegistered;
if (ownsDom) GlobalRegistrator.register();
afterAll(async () => {
  if (ownsDom) await GlobalRegistrator.unregister();
});

const page = (primitives: object[]): DisplayPage => ({ pageIndex: 0, width: 400, height: 200, primitives: primitives as DisplayPrimitive[] });
const image = (extra: object = {}) => ({ kind: 'image', x: 0, y: 0, w: 50, h: 50, src: 'x', relId: 'r1', blockKey: 'i', ...extra });
const shape = (extra: object = {}) => ({ kind: 'shape', x: 0, y: 0, w: 100, h: 50, blockKey: 's', ...extra });
const text = (value: string, blockKey: string) => ({
  kind: 'text', x: 5, y: 5, baselineY: 15, width: 40, text: value, font: '11px sans-serif', color: '#000', blockKey,
});
const labels = { picture: 'Bild' };

test('a picture without alternative text reads as the localised word for a picture', () => {
  const el = buildMirrorPage(page([image()]), { labels }).querySelector('.layout-run-image')!;
  expect([el.getAttribute('role'), el.getAttribute('aria-label')]).toEqual(['img', 'Bild']);
});

test('a picture keeps its own alternative text, and a decorative one stays hidden', () => {
  const named = buildMirrorPage(page([image({ altText: 'Logo' })]), { labels }).querySelector('.layout-run-image')!;
  expect(named.getAttribute('aria-label')).toBe('Logo');
  const decorative = buildMirrorPage(page([image({ decorative: true })]), { labels }).querySelector('.layout-run-image')!;
  expect(decorative.getAttribute('aria-hidden')).toBe('true');
  expect(decorative.hasAttribute('aria-label')).toBe(false);
});

test('with no word to use, a picture without alternative text is not an unnamed image', () => {
  expect(buildMirrorPage(page([image()])).querySelector('[role="img"]')).toBeNull();
});

test('updating a mirror in another language renames an unnamed picture', () => {
  const list = page([image()]);
  const mirror = buildMirrorPage(list, { labels: { picture: 'Picture' } });
  expect(updateMirrorPage(mirror, list, { labels })).toBe(true);
  expect(mirror.querySelector('.layout-run-image')!.getAttribute('aria-label')).toBe('Bild');
});

test('a linked picture stays a link named by its alternative text, or the word for a picture', () => {
  const link = (extra: object) => buildMirrorPage(page([image({ href: 'https://example.com', ...extra })]), { labels }).querySelector('a')!;
  expect([link({}).hasAttribute('role'), link({}).getAttribute('aria-label')]).toEqual([false, 'Bild']);
  expect(link({ altText: 'Home' }).getAttribute('aria-label')).toBe('Home');
});

test('a picture with only a description keeps it, as its name', () => {
  for (const options of [{ labels }, {}]) {
    const el = buildMirrorPage(page([image({ ariaDescription: 'Quarterly sales' })]), options).querySelector('[role="img"]')!;
    expect(el.getAttribute('aria-label')).toBe('Quarterly sales');
  }
  expect(buildMirrorPageText(page([image({ ariaDescription: 'Quarterly sales' })])).querySelector('[role="img"]')).not.toBeNull();
});

test('changing the word for a picture leaves a picture that does not use it in place', () => {
  const list = page([image({ altText: 'Logo' }), image({ ariaLabel: 'Seal', blockKey: 'j', relId: 'r2' })]);
  const mirror = buildMirrorPage(list, { labels: { picture: 'Picture' } });
  const before = Array.from(mirror.querySelectorAll('.layout-run-image'));
  updateMirrorPage(mirror, list, { labels });
  expect(Array.from(mirror.querySelectorAll('.layout-run-image'))).toEqual(before);
});

test('a shape with no name is not an image, and its text still reads', () => {
  for (const build of [buildMirrorPage, buildMirrorPageText]) {
    const mirror = build(page([shape(), text('Boxed note', 's')]), { labels });
    expect(mirror.querySelector('[role="img"]')).toBeNull();
    expect(mirror.textContent).toContain('Boxed note');
  }
});

test('a shape with a name or a description keeps its image role and a name', () => {
  const name = (extra: object) => {
    const el = buildMirrorPage(page([shape(extra)]), { labels }).querySelector('[role="img"]')!;
    return el.getAttribute('aria-label');
  };
  expect(name({ ariaLabel: 'Flow chart' })).toBe('Flow chart');
  expect(name({ ariaDescription: 'Three steps' })).toBe('Three steps');
  expect(buildMirrorPage(page([shape({ decorative: true })]), { labels }).querySelector('.layout-run-shape')!.getAttribute('aria-hidden')).toBe('true');
});
