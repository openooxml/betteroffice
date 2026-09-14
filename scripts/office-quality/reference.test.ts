import { expect, test } from 'bun:test';
import { validateReferenceMetadata } from './reference.mjs';

const hash = 'a'.repeat(64);
const metadata = (pages = 179) => ({
  source: { sha256: hash },
  reference: { status: 'ok', dpi: 150, sha256: hash, pages },
  reference_pages: Array.from({ length: pages }, (_, index) => ({
    url: `https://corpus.betteroffice.dev/example/page_${index + 1}.png`,
    bytes: 100,
    sha256: hash,
  })),
});

test('accepts complete references through the 250-page boundary without changing metadata', () => {
  for (const pages of [1, 100, 101, 179, 182, 250]) {
    const input = metadata(pages);
    const before = structuredClone(input);
    expect(() => validateReferenceMetadata(input, 'example')).not.toThrow();
    expect(input).toEqual(before);
  }
});

test('rejects empty, excessive, noninteger, and mismatched reference counts', () => {
  for (const input of [
    metadata(0),
    metadata(251),
    ...[-1, 1.5, NaN, Infinity, '179', null, undefined].map((pages) => ({
      ...metadata(),
      reference: { ...metadata().reference, pages },
    })),
    { ...metadata(), reference_pages: [] },
    { ...metadata(), reference_pages: metadata(178).reference_pages },
    { ...metadata(), reference_pages: metadata(180).reference_pages },
    { ...metadata(), reference_pages: null },
    { ...metadata(), reference_pages: { length: 179 } },
    { ...metadata(), reference_pages: undefined },
  ]) {
    expect(() => validateReferenceMetadata(input, 'example')).toThrow('Invalid Office reference');
  }
});

test('retains successful-export, source-identity, and 150-DPI requirements', () => {
  for (const input of [
    null,
    {},
    { ...metadata(), reference: undefined },
    { ...metadata(), reference: { ...metadata().reference, status: 'timeout' } },
    { ...metadata(), reference: { ...metadata().reference, dpi: 96 } },
    { ...metadata(), reference: { ...metadata().reference, sha256: 'b'.repeat(64) } },
    ...['', 'invalid', undefined].map((sha256) => ({
      ...metadata(),
      reference: { ...metadata().reference, sha256 },
      source: { sha256 },
    })),
    { ...metadata(), source: undefined },
  ]) {
    expect(() => validateReferenceMetadata(input, 'example')).toThrow('Invalid Office reference');
  }
});
