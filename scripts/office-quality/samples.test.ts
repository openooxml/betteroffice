import { expect, test } from 'bun:test';
import { selectSamples } from './samples.mjs';

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

test('rejects legacy custom collections with guidance before downloading', async () => {
  await expect(
    selectSamples({ QUALITY_COLLECTION: 'demos' }, async () => {
      throw new Error('Unexpected download');
    })
  ).rejects.toThrow(
    'QUALITY_COLLECTION is fixed to office-quality; use QUALITY_FORMAT or QUALITY_SAMPLES'
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

test('validates sample IDs, uniqueness, and the 1000-sample limit in both selection paths', async () => {
  const thousand = Array.from({ length: 1000 }, (_, index) => `sample-${index}`);
  for (const count of [174, 1000]) {
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
    ).rejects.toThrow('1–1000 unique sample folder names');
    await expect(
      selectSamples({}, async () =>
        JSON.stringify({ schema_version: 1, id: 'office-quality', samples: ids })
      )
    ).rejects.toThrow('1–1000 unique sample folder names');
  }
});

test('rejects malformed explicit JSON without falling back to a collection', async () => {
  await expect(
    selectSamples({ QUALITY_SAMPLES: 'not-json' }, async () => {
      throw new Error('Unexpected download');
    })
  ).rejects.toBeInstanceOf(SyntaxError);
});
