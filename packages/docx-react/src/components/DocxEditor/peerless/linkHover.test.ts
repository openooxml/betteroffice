import { expect, test } from 'bun:test';
import type { DisplayPage } from '@betteroffice/docx/layout/render';
import { hyperlinkHoverTitle, popupWithLink } from './linkHover';

type Primitive = DisplayPage['primitives'][number];

const link = (extra: Partial<Primitive> = {}): Primitive =>
  ({ kind: 'rect', x: 10, y: 10, w: 40, h: 12, fill: 'transparent', href: 'https://example.com/', ...extra }) as Primitive;

const page = (primitives: Primitive[], footer: Primitive[] = []): DisplayPage => ({
  pageIndex: 0, width: 200, height: 300, primitives,
  ...(footer.length ? { footer: { kind: 'footer', primitives: footer } as DisplayPage['footer'] } : {}),
});

test('hovering a link with a ScreenTip shows the ScreenTip', () => {
  expect(hyperlinkHoverTitle(page([link({ tooltip: 'Open the report' })]), 20, 15)).toBe('Open the report');
});

test('the ScreenTip written as the link title wins over the popup tooltip', () => {
  expect(hyperlinkHoverTitle(page([link({ tooltip: 'popup', linkTitle: 'Screen tip' })]), 20, 15)).toBe('Screen tip');
});

test('a link without a ScreenTip, or no link at all, shows nothing', () => {
  expect(hyperlinkHoverTitle(page([link()]), 20, 15)).toBeNull();
  expect(hyperlinkHoverTitle(page([link({ tooltip: '   ' })]), 20, 15)).toBeNull();
  expect(hyperlinkHoverTitle(page([link({ tooltip: 'Open the report' })]), 120, 150)).toBeNull();
  expect(hyperlinkHoverTitle(null, 20, 15)).toBeNull();
});

test('a link in the footer shows its ScreenTip too', () => {
  expect(hyperlinkHoverTitle(page([], [link({ y: 280, tooltip: 'Footer link' })]), 20, 285)).toBe('Footer link');
});

const hit = { href: 'https://example.com/', displayText: '比亚迪', position: { top: 1, left: 2 } };

test('the popup shows the whole link text, not the run that was clicked', () => {
  expect(popupWithLink(hit, { url: 'https://example.com/', displayText: '比亚迪 2025 年度简报', tooltip: 'Report' }))
    .toEqual({ ...hit, displayText: '比亚迪 2025 年度简报', tooltip: 'Report' });
});

test('the clicked run stands when the link cannot be read', () => {
  expect(popupWithLink(hit, null)).toBe(hit);
  expect(popupWithLink(hit, { displayText: 'not a link' })).toBe(hit);
  expect(popupWithLink({ ...hit, tooltip: 'From paint' }, { url: 'https://example.com/', displayText: '' }))
    .toEqual({ ...hit, tooltip: 'From paint' });
});
