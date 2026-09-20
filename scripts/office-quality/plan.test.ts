import { expect, test } from 'bun:test';
import {
  createPlan,
  mapPool,
  preparePlan,
  selectPlanSamples,
  validatePlan,
} from './plan.mjs';

const hash = 'a'.repeat(40);
const assetHash = 'b'.repeat(64);
const metadata = (format = 'pptx') => ({
  format,
  source: { sha256: assetHash },
  reference: { status: 'ok', dpi: 150, sha256: assetHash, pages: 1 },
  reference_pages: [
    {
      url: 'https://corpus.betteroffice.dev/example/page_0001.png',
      bytes: 1,
      sha256: assetHash,
    },
  ],
});
const input = () => ({
  source_sha: hash,
  commit: hash,
  versions: { docx: '1.2.3', pptx: '2.3.4', xlsx: '3.4.5' },
  react_version: '4.5.6',
  samples: [
    { id: 'pptarena-001-original', format: 'pptx', metadata: metadata() },
    { id: 'docx-example', format: 'docx', metadata: metadata('docx') },
  ],
});

test('creates a canonical plan with only nonempty formats', () => {
  const plan = createPlan(input());
  expect(plan.schema_version).toBe(1);
  expect(plan.formats).toEqual(['docx', 'pptx']);
  expect(validatePlan(plan)).toEqual(plan);
});

test('selects a worker format from frozen metadata before asset work', () => {
  const plan = createPlan(input());
  expect(selectPlanSamples(plan, 'pptx')).toMatchObject({
    formats: ['pptx'],
    samples: [{ id: 'pptarena-001-original', format: 'pptx' }],
  });
  expect(() => selectPlanSamples(plan, 'xlsx')).toThrow('no samples');
  expect(() => selectPlanSamples(plan, 'vsdx')).toThrow('QUALITY_FORMAT');
});

test('rejects bad revisions, metadata, duplicate samples, and format lists', () => {
  for (const plan of [
    { ...input(), source_sha: 'bad' },
    {
      ...input(),
      samples: [{ id: 'pptarena-001-original', format: 'docx', metadata: metadata() }],
    },
    {
      ...input(),
      samples: [
        {
          id: 'pptarena-001-original',
          format: 'pptx',
          metadata: { ...metadata(), id: 'other' },
        },
      ],
    },
    { ...input(), samples: [input().samples[0], input().samples[0]] },
    {
      ...input(),
      samples: Array.from({ length: 1001 }, (_, index) => ({
        id: `sample-${index}`,
        format: 'pptx',
        metadata: metadata(),
      })),
    },
  ])
    expect(() => createPlan(plan)).toThrow();
  expect(() => validatePlan({ ...createPlan(input()), formats: ['pptx'] })).toThrow(
    'Invalid fidelity plan formats'
  );
});

test('bounds metadata work while preserving canonical input order', async () => {
  let active = 0;
  let maximum = 0;
  const values = await mapPool([0, 1, 2, 3, 4], 2, async (value) => {
    active++;
    maximum = Math.max(maximum, active);
    await Promise.resolve();
    active--;
    return value * 2;
  });
  expect(values).toEqual([0, 2, 4, 6, 8]);
  expect(maximum).toBeLessThanOrEqual(2);
});

test('freezes canonical metadata and every package version in one plan', async () => {
  const plan = await preparePlan(
    {},
    {
      download: async (url: string) => {
        if (url.endsWith('/collections/office-quality.json'))
          return JSON.stringify({
            schema_version: 1,
            id: 'office-quality',
            samples: ['docx-example', 'pptarena-001-original'],
          });
        if (url.endsWith('/docx-example/metadata.json'))
          return JSON.stringify(metadata('docx'));
        if (url.endsWith('/pptarena-001-original/metadata.json'))
          return JSON.stringify(metadata());
        throw new Error(`Unexpected URL: ${url}`);
      },
      command: async (_program: string, args: string[]) =>
        args[0] === 'rev-parse' ? `${hash}\n` : `${'c'.repeat(40)}\n`,
      registry: async (name: string) => ({
        version: name.endsWith('docx-react') ? '4.5.6' : '1.2.3',
        gitHead: 'd'.repeat(40),
      }),
      poolSize: 2,
    }
  );
  expect(plan.source_sha).toBe(hash);
  expect(plan.commit).toBe('c'.repeat(40));
  expect(plan.versions).toEqual({ docx: '1.2.3', pptx: '1.2.3', xlsx: '1.2.3' });
  expect(plan.react_version).toBe('4.5.6');
  expect(plan.docx_published_source_sha).toBe('d'.repeat(40));
  expect(plan.samples.map((sample) => sample.id)).toEqual([
    'docx-example',
    'pptarena-001-original',
  ]);
});
