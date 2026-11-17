import { expect, test } from 'bun:test';
import { MAX_SAMPLES, corpusOrigin, selectSamples } from './samples.mjs';

const samples = Array.from({ length: 174 }, (_, index) => `sample-${index}`);

function collection(ids = samples, id = 'office-quality') {
  return JSON.stringify({ schema_version: 1, id, samples: ids });
}

test('loads the default collection from the canonical origin with a bounded download', async () => {
  const requests: unknown[][] = [];
  const selected = await selectSamples({}, async (...args: unknown[]) => {
    requests.push(args);
    return collection();
  });
  expect(selected).toEqual(samples);
  expect(requests).toEqual([
    ['https://corpus.betteroffice.dev/collections/office-quality.json', 2 * 1024 * 1024],
  ]);
});

test('explicit samples override the canonical collection without downloading it', async () => {
  const selected = await selectSamples(
    {
      QUALITY_SAMPLES: '["betteroffice-workbook"]',
      QUALITY_COLLECTION: 'office-quality',
    },
    async () => {
      throw new Error('Unexpected download');
    }
  );
  expect(selected).toEqual(['betteroffice-workbook']);
});

test('empty sample overrides use the canonical collection', async () => {
  for (const value of ['', ' \n ']) {
    const selected = await selectSamples(
      { QUALITY_SAMPLES: value, QUALITY_COLLECTION: 'office-quality' },
      async (url: string) => {
        expect(url).toBe(
          'https://corpus.betteroffice.dev/collections/office-quality.json'
        );
        return collection(['betteroffice-demo']);
      }
    );
    expect(selected).toEqual(['betteroffice-demo']);
  }
});

test('loads a named collection from a configured origin', async () => {
  const requests: unknown[] = [];
  const selected = await selectSamples(
    { QUALITY_COLLECTION: 'full-corpus-v1', QUALITY_CORPUS_ORIGIN: 'http://127.0.0.1:8099' },
    async (url: string) => {
      requests.push(url);
      return collection(['govuk-1'], 'full-corpus-v1');
    }
  );
  expect(selected).toEqual(['govuk-1']);
  expect(requests).toEqual(['http://127.0.0.1:8099/collections/full-corpus-v1.json']);
});

test('rejects collection names that are not folder names before downloading', async () => {
  for (const name of ['../office-quality', 'Office', 'a/b', 'x?y']) {
    await expect(
      selectSamples({ QUALITY_COLLECTION: name }, async () => {
        throw new Error('Unexpected download');
      })
    ).rejects.toThrow('QUALITY_COLLECTION must be a collection folder name');
  }
});

test('accepts https and loopback http corpus origins only', () => {
  expect(corpusOrigin({})).toBe('https://corpus.betteroffice.dev');
  expect(corpusOrigin({ QUALITY_CORPUS_ORIGIN: ' ' })).toBe('https://corpus.betteroffice.dev');
  expect(corpusOrigin({ QUALITY_CORPUS_ORIGIN: 'https://bench.example.org/' })).toBe('https://bench.example.org');
  expect(corpusOrigin({ QUALITY_CORPUS_ORIGIN: 'http://localhost:8080' })).toBe('http://localhost:8080');
  for (const value of [
    'http://bench.example.org',
    'https://bench.example.org/prefix',
    'https://bench.example.org?x=1',
    'file:///tmp/corpus',
    'not a url',
  ])
    expect(() => corpusOrigin({ QUALITY_CORPUS_ORIGIN: value })).toThrow(
      'QUALITY_CORPUS_ORIGIN must be an https origin or a loopback http origin'
    );
});

test('requires the canonical collection identity and schema', async () => {
  for (const manifest of [
    null,
    {},
    { schema_version: 1, id: 'other', samples },
    { schema_version: '1', id: 'office-quality', samples },
    { schema_version: 2, id: 'office-quality', samples },
  ]) {
    await expect(selectSamples({}, async () => JSON.stringify(manifest))).rejects.toThrow(
      'Invalid corpus collection'
    );
  }
});

test('validates sample IDs, uniqueness, and the sample limit in both selection paths', async () => {
  const thousand = Array.from({ length: MAX_SAMPLES }, (_, index) => `sample-${index}`);
  for (const count of [174, 1000, MAX_SAMPLES]) {
    const ids = thousand.slice(0, count);
    expect(
      await selectSamples({ QUALITY_SAMPLES: JSON.stringify(ids) }, async () => '')
    ).toEqual(ids);
    expect(await selectSamples({}, async () => collection(ids))).toEqual(ids);
  }

  for (const ids of [
    [],
    null,
    'sample',
    ['sample', 'sample'],
    [''],
    [42],
    ['../source'],
    ['nested/sample'],
    ['sample?x=1'],
    ['Sample'],
    ['sample\n'],
    [...thousand, 'one-too-many'],
  ]) {
    await expect(
      selectSamples({ QUALITY_SAMPLES: JSON.stringify(ids) }, async () => '')
    ).rejects.toThrow(`1–${MAX_SAMPLES} unique sample folder names`);
    await expect(
      selectSamples({}, async () =>
        JSON.stringify({ schema_version: 1, id: 'office-quality', samples: ids })
      )
    ).rejects.toThrow(`1–${MAX_SAMPLES} unique sample folder names`);
  }
});

test('rejects malformed explicit JSON without falling back to a collection', async () => {
  await expect(
    selectSamples({ QUALITY_SAMPLES: 'not-json' }, async () => {
      throw new Error('Unexpected download');
    })
  ).rejects.toBeInstanceOf(SyntaxError);
});
