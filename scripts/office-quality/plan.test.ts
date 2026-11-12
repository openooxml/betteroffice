import { expect, test } from 'bun:test';
import {
  createPlan,
  mapPool,
  preparePlan,
  selectPlanSamples,
  validatePlan,
} from './plan.mjs';
import { MAX_SAMPLES } from './samples.mjs';

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
    { ...input(), pptx_published_source_sha: 'bad' },
    { ...input(), xlsx_published_source_sha: 'bad' },
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
      samples: Array.from({ length: MAX_SAMPLES + 1 }, (_, index) => ({
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
      packument: async () => ({ versions: { '1.0.0': { gitHead: 'd'.repeat(40) }, '2.0.0': { gitHead: 'e'.repeat(40) } } }),
      poolSize: 2,
    }
  );
  expect(plan.source_sha).toBe(hash);
  expect(plan.commit).toBe('c'.repeat(40));
  expect(plan.versions).toEqual({ docx: '1.2.3', pptx: '1.2.3', xlsx: '1.2.3' });
  expect(plan.react_version).toBe('4.5.6');
  expect(plan.published_fonts).toEqual({ docx: '1.0.0', pptx: '1.0.0' });
  expect(plan.published_fonts_cjk).toEqual({ docx: '1.0.0', pptx: '1.0.0' });
  expect(plan.published_fonts_intl).toEqual({ docx: '1.0.0' });
  expect(plan.docx_published_source_sha).toBe('d'.repeat(40));
  expect(plan.pptx_published_source_sha).toBe('d'.repeat(40));
  expect(plan.samples.map((sample) => sample.id)).toEqual([
    'docx-example',
    'pptarena-001-original',
  ]);
});

test('the published channel uses the fonts release from the format release commit, else the latest stable fonts published before it', async () => {
  const release = 'a'.repeat(40);
  const formatTime = '2026-10-27T11:28:38.352Z';
  const dependencies = (fonts: any, cjk: any) => ({
    download: async (url: string) => {
      if (url.endsWith('/collections/office-quality.json'))
        return JSON.stringify({ schema_version: 1, id: 'office-quality', samples: ['pptarena-001-original'] });
      if (url.endsWith('/pptarena-001-original/metadata.json')) return JSON.stringify(metadata());
      throw new Error(`Unexpected URL: ${url}`);
    },
    command: async (_program: string, args: string[]) => (args[0] === 'rev-parse' ? `${hash}\n` : `${'c'.repeat(40)}\n`),
    registry: async (name: string) => ({ version: name.endsWith('docx-react') ? '4.5.6' : '0.5.1', gitHead: release }),
    packument: async (name: string) =>
      name === '@betteroffice/fonts' ? fonts : name === '@betteroffice/fonts-cjk' ? cjk : { time: { '0.5.1': formatTime } },
    poolSize: 1,
  });
  const fonts = (versions: Record<string, [string, string]>) => ({
    'dist-tags': { latest: Object.keys(versions).at(-1) },
    versions: Object.fromEntries(Object.entries(versions).map(([version, [gitHead]]) => [version, { gitHead }])),
    time: Object.fromEntries(Object.entries(versions).map(([version, [, time]]) => [version, time])),
  });
  const cjkVersions = {
    '0.3.10': ['4'.repeat(40), '2026-10-01T00:00:00Z'],
    '0.3.5': ['5'.repeat(40), '2026-10-20T00:00:00Z'],
    '0.4.2': [release, '2026-10-27T11:29:10Z'],
    '0.6.0': ['6'.repeat(40), '2026-10-30T00:00:00Z'],
  } satisfies Record<string, [string, string]>;
  const plan = async (versions: Record<string, [string, string]>) =>
    preparePlan({}, dependencies(fonts(versions), fonts(cjkVersions)) as any);
  const sameCommit = await plan({
    '0.4.1': ['b'.repeat(40), '2026-10-26T00:00:00Z'],
    '0.4.2': [release, '2026-10-27T11:29:00Z'],
    '0.9.0': ['f'.repeat(40), '2026-11-01T00:00:00Z'],
  });
  expect(sameCommit.published_fonts).toEqual({ pptx: '0.4.2' });
  expect(sameCommit.fonts_basis).toEqual({ pptx: 'release-commit' });
  expect(sameCommit.published_fonts_cjk).toEqual({ pptx: '0.4.2' });
  expect(sameCommit.fonts_cjk_basis).toEqual({ pptx: 'release-commit' });
  const formatOnly = await plan({
    '0.0.0': ['0'.repeat(40), '2026-10-27T11:00:00Z'],
    '0.0.0-stage': ['1'.repeat(40), '2026-10-27T11:10:00Z'],
    '0.3.0': ['2'.repeat(40), '2026-09-01T00:00:00Z'],
    '0.4.1': ['b'.repeat(40), '2026-10-26T00:00:00Z'],
    '0.4.3-next.0': ['3'.repeat(40), '2026-10-27T11:20:00Z'],
    '0.9.0': ['f'.repeat(40), '2026-11-01T00:00:00Z'],
  });
  expect(formatOnly.published_fonts).toEqual({ pptx: '0.4.1' });
  expect(formatOnly.fonts_basis).toEqual({ pptx: 'latest-before-release' });
  expect(formatOnly.published_fonts_cjk).toEqual({ pptx: '0.3.5' });
  expect(formatOnly.fonts_cjk_basis).toEqual({ pptx: 'latest-before-release' });
  const backport = await plan({
    '0.4.1': ['b'.repeat(40), '2026-10-20T00:00:00Z'],
    '0.3.9': ['9'.repeat(40), '2026-10-25T00:00:00Z'],
  });
  expect(backport.published_fonts).toEqual({ pptx: '0.3.9' });
  await expect(plan({ '0.0.0': ['0'.repeat(40), '2026-10-01T00:00:00Z'], '0.9.0': ['f'.repeat(40), '2026-11-01T00:00:00Z'] }))
    .rejects.toThrow('No @betteroffice/fonts release from the release commit of @betteroffice/pptx@0.5.1 or published before it');
  await expect(plan({ '0.4.2': [release, '2026-10-27T11:00:00Z'], '0.4.3': [release, '2026-10-27T11:01:00Z'] }))
    .rejects.toThrow('Expected at most one @betteroffice/fonts release from the release commit of @betteroffice/pptx@0.5.1, found 2');
});

test('a DOCX plan pairs the published docx release with the fonts and fonts-cjk releases of its release commit', async () => {
  const release = 'a'.repeat(40);
  const requested: string[] = [];
  const plan = await preparePlan({}, {
    download: async (url: string) => {
      if (url.endsWith('/collections/office-quality.json'))
        return JSON.stringify({ schema_version: 1, id: 'office-quality', samples: ['docx-example'] });
      if (url.endsWith('/docx-example/metadata.json')) return JSON.stringify(metadata('docx'));
      throw new Error(`Unexpected URL: ${url}`);
    },
    command: async (_program: string, args: string[]) => (args[0] === 'rev-parse' ? `${hash}\n` : `${'c'.repeat(40)}\n`),
    registry: async (name: string) => ({ version: name.endsWith('docx-react') ? '0.5.1' : '0.5.1', gitHead: release }),
    packument: async (name: string) => {
      requested.push(name);
      if (name === '@betteroffice/fonts')
        return { versions: { '0.4.1': { gitHead: 'b'.repeat(40) }, '0.4.2': { gitHead: release } }, time: { '0.4.1': '2026-10-26T00:00:00Z', '0.4.2': '2026-10-27T11:27:19Z' } };
      if (name === '@betteroffice/fonts-cjk')
        return { versions: { '0.4.1': { gitHead: 'b'.repeat(40) }, '0.4.2': { gitHead: release }, '0.5.0': { gitHead: 'f'.repeat(40) } }, time: { '0.4.2': '2026-10-27T11:27:27Z' } };
      if (name === '@betteroffice/fonts-intl')
        return { versions: { '0.4.1': { gitHead: 'b'.repeat(40) }, '0.4.2': { gitHead: release }, '0.5.0': { gitHead: 'f'.repeat(40) } }, time: { '0.4.2': '2026-10-27T11:27:23Z' } };
      return { time: { '0.5.1': '2026-10-27T11:28:38Z' } };
    },
    poolSize: 1,
  } as any);
  expect(plan.published_fonts).toEqual({ docx: '0.4.2' });
  expect(plan.fonts_basis).toEqual({ docx: 'release-commit' });
  expect(plan.published_fonts_cjk).toEqual({ docx: '0.4.2' });
  expect(plan.fonts_cjk_basis).toEqual({ docx: 'release-commit' });
  expect(plan.published_fonts_intl).toEqual({ docx: '0.4.2' });
  expect(plan.fonts_intl_basis).toEqual({ docx: 'release-commit' });
  expect(requested).toEqual(expect.arrayContaining(['@betteroffice/docx', '@betteroffice/fonts', '@betteroffice/fonts-cjk', '@betteroffice/fonts-intl']));
  expect(validatePlan(plan)).toEqual(plan);
  expect(() => createPlan({ ...input(), published_fonts: { pptx: '0.4.2' }, published_fonts_cjk: { docx: '0.4.2' } })).toThrow('Invalid fidelity plan versions');
  expect(() => createPlan({ ...input(), published_fonts: { docx: '0.4.2' }, published_fonts_cjk: { docx: 'latest' } })).toThrow('Invalid fidelity plan versions');
  expect(() => createPlan({ ...input(), published_fonts: { docx: '0.4.2', pptx: '0.4.2' }, published_fonts_intl: { pptx: '0.4.2' } })).toThrow('Invalid fidelity plan versions');
  expect(() => createPlan({ ...input(), published_fonts: { docx: '0.4.2' }, published_fonts_intl: { docx: 'latest' } })).toThrow('Invalid fidelity plan versions');
});

test('the local runner plans PPTX and XLSX collections through the registry loaders it shares with planning', async () => {
  const { registryLoaders } = await import('./plan.mjs');
  const release = 'a'.repeat(40);
  const served: string[] = [];
  const download = async (url: string) => {
    served.push(url);
    if (url.endsWith('/collections/office-quality.json'))
      return JSON.stringify({ schema_version: 1, id: 'office-quality', samples: ['pptarena-001-original', 'sheet-001'] });
    if (url.endsWith('/pptarena-001-original/metadata.json')) return JSON.stringify(metadata());
    if (url.endsWith('/sheet-001/metadata.json')) return JSON.stringify(metadata('xlsx'));
    if (url.endsWith('/latest')) return JSON.stringify({ version: url.includes('docx-react') ? '4.5.6' : '0.5.1', gitHead: release });
    if (url === 'https://registry.npmjs.org/%40betteroffice%2Ffonts' || url === 'https://registry.npmjs.org/%40betteroffice%2Ffonts-cjk')
      return JSON.stringify({ versions: { '0.4.2': { gitHead: release } }, time: { '0.4.2': '2026-10-27T11:27:19Z' } });
    if (/^https:\/\/registry\.npmjs\.org\/%40betteroffice%2F(pptx|xlsx)$/.test(url))
      return JSON.stringify({ time: { '0.5.1': '2026-10-27T11:28:38Z' } });
    throw new Error(`Unexpected URL: ${url}`);
  };
  const plan = await preparePlan({}, {
    download,
    command: async (_program: string, args: string[]) => (args[0] === 'rev-parse' ? `${hash}\n` : `${'c'.repeat(40)}\n`),
    ...(registryLoaders as any)(download),
    poolSize: 1,
  } as any);
  expect(plan.published_fonts).toEqual({ pptx: '0.4.2', xlsx: '0.4.2' });
  expect(served).toContain('https://registry.npmjs.org/%40betteroffice%2Ffonts');
  expect(served).toContain('https://registry.npmjs.org/%40betteroffice%2Ffonts-cjk');
  const runner = await Bun.file(new URL('./run.mjs', import.meta.url)).text();
  expect(runner).toContain('...registryLoaders(download)');
});
