import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { afterAll, expect, test } from 'bun:test';
import type { DocxClipboardFragment } from '@betteroffice/docx/yrs';
import { clipboardFragmentHtml, parseClipboardHtml } from './clipboardHtml';

const ownsDom = !GlobalRegistrator.isRegistered;
if (ownsDom) GlobalRegistrator.register();
afterAll(async () => {
  if (ownsDom) await GlobalRegistrator.unregister();
});

test('the editor\'s own copy of a large picture pastes as the fragment with the picture', () => {
  const src = `data:image/png;base64,${'A'.repeat(1_500_000)}`;
  const fragment: DocxClipboardFragment = { paragraphs: [{ text: 'ab', inlines: [
    { text: 'a' }, { image: { src, width: 640, height: 480 } }, { text: 'b' },
  ] }] };
  const html = clipboardFragmentHtml(fragment);
  expect(html.length).toBeGreaterThan(3_000_000);
  const parsed = parseClipboardHtml(html);
  expect(parsed !== null).toBe(true);
  expect(parsed).toEqual(fragment);
}, 60_000);

test('a valid payload among more than fifty thousand elements pastes as its fragment', () => {
  const fragment: DocxClipboardFragment = { paragraphs: [{ text: 'Bold', runs: [{ text: 'Bold', format: { other: { bold: true } } }] }] };
  const filler = '<span>x</span>'.repeat(25_001);
  const html = `<html><body>${filler}${clipboardFragmentHtml(fragment)}${filler}</body></html>`;
  expect(html.split('<span>').length - 1).toBeGreaterThan(50_000);
  expect(parseClipboardHtml(html)).toEqual(fragment);
}, 60_000);
