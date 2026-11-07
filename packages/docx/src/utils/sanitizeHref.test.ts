import { expect, test } from 'bun:test';
import { sanitizeHref } from './sanitizeHref';

test('document links keep relative and fragment hrefs', () => {
  for (const href of ['foo.html', '../chapter.xml', '#section', '//example.test/path']) {
    expect(sanitizeHref(href)).toBe(href);
  }
});

test('document links preserve allowed protocols and reject unsafe protocols after URL preprocessing', () => {
  for (const href of ['https://example.test', 'http://example.test', 'mailto:a@example.test', 'tel:+123', 'ftp://example.test']) {
    expect(sanitizeHref(href)).toBe(href);
  }
  for (const href of ['javascript:alert(1)', 'JaVaScRiPt:alert(1)', 'java\tscript:alert(1)', '\x01javascript:alert(1)', 'data:text/html,bad', 'vbscript:bad', 'file:///tmp/file']) {
    expect(sanitizeHref(href)).toBeUndefined();
  }
  expect(sanitizeHref(null)).toBeUndefined();
  expect(sanitizeHref('')).toBeUndefined();
});
