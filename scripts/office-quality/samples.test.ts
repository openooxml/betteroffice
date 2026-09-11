import { expect, test } from 'bun:test';
import { selectSamples } from './samples.mjs';

const samples = Array.from({ length: 53 }, (_, index) => `sample-${index}`);

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

test('explicit samples override the collection without downloading it', async () => {
  const selected = await selectSamples(
    { QUALITY_SAMPLES: '["betteroffice-workbook"]', QUALITY_COLLECTION: '../ignored' },
    async () => {
      throw new Error('Unexpected download');
    }
  );
  expect(selected).toEqual(['betteroffice-workbook']);
});

test('empty sample overrides use the selected collection', async () => {
  for (const value of ['', ' \n ']) {
    const selected = await selectSamples(
      { QUALITY_SAMPLES: value, QUALITY_COLLECTION: 'demos' },
      async (url: string) => {
        expect(url).toBe('https://corpus.betteroffice.dev/collections/demos.json');
        return collection(['betteroffice-demo'], 'demos');
      }
    );
    expect(selected).toEqual(['betteroffice-demo']);
  }
});

test('rejects invalid collection names before downloading', async () => {
  for (const id of [
    '../other',
    '/other',
    'https://example.com',
    'other?x=1',
    'Docs',
    'docs\n',
  ]) {
    await expect(
      selectSamples({ QUALITY_COLLECTION: id }, async () => {
        throw new Error('Unexpected download');
      })
    ).rejects.toThrow('QUALITY_COLLECTION');
  }
});

test('requires the requested collection identity and schema', async () => {
  for (const manifest of [
    null,
    {},
    { schema_version: 1, id: 'other', samples },
    { schema_version: '1', id: 'office-quality', samples },
    { schema_version: 2, id: 'office-quality', samples },
  ]) {
    await expect(
      selectSamples({}, async () => JSON.stringify(manifest))
    ).rejects.toThrow('Invalid corpus collection');
  }
});

test('validates sample IDs, uniqueness, and the 100-sample limit in both selection paths', async () => {
  const hundred = Array.from({ length: 100 }, (_, index) => `sample-${index}`);
  expect(
    await selectSamples({ QUALITY_SAMPLES: JSON.stringify(hundred) }, async () => '')
  ).toEqual(hundred);
  expect(await selectSamples({}, async () => collection(hundred))).toEqual(hundred);

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
    [...hundred, 'one-too-many'],
  ]) {
    await expect(
      selectSamples({ QUALITY_SAMPLES: JSON.stringify(ids) }, async () => '')
    ).rejects.toThrow('1–100 unique sample folder names');
    await expect(
      selectSamples({}, async () =>
        JSON.stringify({ schema_version: 1, id: 'office-quality', samples: ids })
      )
    ).rejects.toThrow('1–100 unique sample folder names');
  }
});

test('rejects malformed explicit JSON without falling back to a collection', async () => {
  await expect(
    selectSamples({ QUALITY_SAMPLES: 'not-json' }, async () => {
      throw new Error('Unexpected download');
    })
  ).rejects.toBeInstanceOf(SyntaxError);
});
