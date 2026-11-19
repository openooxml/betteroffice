import { spawnSync } from 'node:child_process';
import { mkdtemp, readFile, rm, symlink, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { expect, test } from 'bun:test';
import { aggregateReport, renderSection, updateReadme } from './readme.mjs';

const row = (label: string, published: string, current: string) =>
  `<tr><td>${label}</td><td align="right">${published}</td><td align="right">${current}</td></tr>`;

const commit = 'a'.repeat(40);
function report() {
  const comparison = {
    source_verified: true,
    reference: { status: 'ok', sha256: 'source' },
    actual: { status: 'ok', sha256: 'source' },
    penalized_ssim: 0.8,
  };
  return {
    commit,
    versions: { docx: '0.1.0', pptx: '0.0.4', xlsx: '0.1.0' },
    samples: [
      {
        id: 'demo',
        format: 'docx',
        metadata_url: 'https://corpus.betteroffice.dev/demo/metadata.json',
        comparisons: [
          { ...comparison, channel: 'published', version: '0.1.0' },
          { ...comparison, channel: 'commit', renderer_source_commit: commit },
        ],
      },
    ],
  };
}

test('generates every format and never reuses scores for a different release or commit', () => {
  const input = report();
  expect(renderSection(input)).toContain('0.8000');
  input.versions.docx = '0.2.0';
  input.commit = 'b'.repeat(40);
  const section = renderSection(input);
  expect(section).not.toContain('0.8000');
  expect(section).toContain('### PPTX');
  expect(section).toContain('pptx/v/0.0.4"');
  expect(section).toContain('xlsx/v/0.1.0"');
  expect(section).toContain(row('SSIM', '—', '—'));
  expect(section).toContain(row('Scored/total', '0/1', '0/1'));
});

test('reports DOCX page agreement apart from SSIM, and only where pages are known', () => {
  const input = report();
  const paged = (reference: number, actual: number, id: string) => ({
    id,
    format: 'docx',
    metadata_url: `https://corpus.betteroffice.dev/${id}/metadata.json`,
    comparisons: [
      {
        source_verified: true,
        reference: { status: 'ok', sha256: 'source' },
        actual: { status: 'ok', sha256: 'source' },
        penalized_ssim: 0.8,
        reference_pages: reference,
        actual_pages: actual,
        channel: 'commit',
        renderer_source_commit: commit,
      },
    ],
  });
  // two exact, one two pages short, one page over: 2/4 exact, error 3
  input.samples = [paged(4, 4, 'a'), paged(9, 9, 'b'), paged(5, 3, 'c'), paged(2, 3, 'd')] as any;
  const section = renderSection(input);
  expect(section).toContain(row('Exact page counts', '—', '2/4'));
  expect(section).toContain(row('Absolute page error', '—', '3'));
  // a comparison without page counts contributes to SSIM but not to page agreement
  input.samples = [paged(4, 4, 'a'), report().samples[0]] as any;
  const mixed = renderSection(input);
  expect(mixed).toContain(row('Exact page counts', '—', '1/1'));
});

test('refuses a README without generated markers and leaves every section in it untouched', async () => {
  const readme =
    '# Project\n\n## Packages\n\n### DOCX\n\nPackage list.\n\n## Development\n\nCommands.\n\n## Visual fidelity\n\nOld text.\n\n## Contributing\n\nPolicy.\n';
  for (const section of [renderSection(report()), renderSection(report(), { privateCorpus: true, date: DATE })])
    expect(() => updateReadme(readme, section)).toThrow('README has no generated block');
  const dir = await mkdtemp(join(tmpdir(), 'readme-no-markers-'));
  try {
    const path = join(dir, 'README.md');
    await writeFile(path, readme);
    await writeFile(join(dir, 'public.md'), renderSection(report()));
    refused(cli(['--section', 'public.md', 'README.md'], dir, false), 'README has no generated block');
    expect(await readFile(path, 'utf8')).toBe(readme);
  } finally {
    await rm(dir, { recursive: true, force: true });
  }
});

test('refuses failed or mismatched render records', () => {
  const input = report();
  input.samples[0].comparisons[0].actual.sha256 = 'wrong';
  expect(() => renderSection(input)).toThrow('mismatched');
});

test('keeps each format tied to its own published version and comparison', () => {
  const input = report();
  for (const [format, version, score] of [
    ['pptx', '0.0.4', 0.91],
    ['xlsx', '0.1.0', 0.87],
  ] as const) {
    input.samples.push({
      id: format,
      format,
      metadata_url: `https://corpus.betteroffice.dev/${format}/metadata.json`,
      comparisons: input.samples[0].comparisons.map((comparison) => ({
        ...comparison,
        version,
        penalized_ssim: score,
      })),
    });
  }
  const section = renderSection(input);
  const pptx = section.slice(section.indexOf('### PPTX'), section.indexOf('### XLSX'));
  expect(pptx).toContain(row('SSIM', '0.9100', '0.9100'));
  expect(pptx).toContain(row('Scored/total', '1/1', '1/1'));
  const xlsx = section.slice(section.indexOf('### XLSX'));
  expect(xlsx).toContain(row('SSIM', '0.8700', '0.8700'));
  expect(xlsx).toContain(row('Scored/total', '1/1', '1/1'));
  expect(section).toContain(row('SSIM', '0.8000', '0.8000'));
});

test('all failures show no score and cannot carry an invented zero', () => {
  const input = report();
  const failure = {
    channel: 'published',
    version: '0.1.0',
    status: 'failed',
    stage: 'capture',
    error: 'Could not render',
  };
  const failedReport = {
    ...input,
    samples: [{ ...input.samples[0], comparisons: [failure] }],
  };
  const section = renderSection(failedReport);
  expect(section).not.toContain('0.0000');
  expect(section).toContain(row('SSIM', '—', '—'));
  expect(section).toContain(row('Exact page counts', '—', '—'));
  expect(section).not.toContain('Could not render');
  expect(() =>
    renderSection({
      ...failedReport,
      samples: [
        { ...input.samples[0], comparisons: [{ ...failure, penalized_ssim: 0 }] },
      ],
    })
  ).toThrow('Invalid failed comparison');
});

test('omits failure diagnostics from the README', () => {
  const input = report();
  const failedReport = {
    ...input,
    samples: [
      {
        ...input.samples[0],
        comparisons: [
          {
            channel: 'published',
            version: '0.1.0',
            status: 'failed',
            stage: 'capture',
            error:
              '<script>alert(1)</script>|[link](https://example.com/private)\nextra row',
          },
        ],
      },
    ],
  };
  const section = renderSection(failedReport);
  expect(section).not.toContain('<script>');
  expect(section).not.toContain('https://example.com');
  expect(section).not.toContain('extra row');
  expect(section).not.toContain('Failed or missing');
  expect(section).not.toContain('<td>Sample</td>');
});

test('duplicate and invalid successful comparisons cannot be published', () => {
  const input = report();
  input.samples[0].comparisons.push(input.samples[0].comparisons[0]);
  expect(() => renderSection(input)).toThrow('Duplicate');
  for (const changed of [
    { source_verified: false },
    { actual: { status: 'error', sha256: 'source' } },
    { reference: { status: 'ok' }, actual: { status: 'ok' } },
    { penalized_ssim: NaN },
    { penalized_ssim: 2 },
    { resized: true },
  ]) {
    const invalid = report();
    Object.assign(invalid.samples[0].comparisons[0], changed);
    expect(() => renderSection(invalid)).toThrow('mismatched');
  }
});

test('the section embeds the benchmark chart flush left, over collapsed data, with the numbers as alt text', () => {
  const section = renderSection(report());
  for (const name of ['benchmarks']) {
    expect(section).toContain(`<source media="(prefers-color-scheme: dark)" srcset="https://betteroffice.dev/api/charts/${name}.svg?theme=dark">`);
    expect(section).toContain(`<img src="https://betteroffice.dev/api/charts/${name}.svg" alt="`);
  }
  expect(section).toContain('width="100%">');
  expect(section).not.toContain('<p align="center">');
  expect(section.indexOf('<picture>')).toBeLessThan(section.indexOf('<details>\n<summary>Data</summary>'));
  expect(section.indexOf('<summary>Data</summary>')).toBeLessThan(section.indexOf('### DOCX'));
  expect(section.indexOf('### XLSX')).toBeLessThan(section.indexOf('</details>'));
  expect(section).not.toContain('perf.svg');
  expect(section.lastIndexOf('</details>')).toBeLessThan(section.indexOf('For how scores'));
  expect(section.startsWith('<!-- BEGIN GENERATED VISUAL FIDELITY -->')).toBe(true);
});

test('the aggregate carries per-format counts and means and nothing that names a document', () => {
  const input = report();
  input.samples.push({
    ...input.samples[0],
    id: 'second',
    metadata_url: 'https://corpus.betteroffice.dev/second/metadata.json',
    comparisons: input.samples[0].comparisons.map((entry) => ({ ...entry, penalized_ssim: 0.6, reference_pages: 2, actual_pages: 3 })),
  });
  const aggregate = aggregateReport(input);
  expect(aggregate.schema_version).toBe(1);
  expect(aggregate.commit).toBe(commit);
  expect(aggregate.formats.docx).toEqual({
    samples: 2,
    published_version: '0.1.0',
    libreoffice_version: null,
    fidelity: {
      published: { ssim: 0.7, scored: 2, total: 2, exact_pages: 0, paged: 1, page_error: 1 },
      commit: { ssim: 0.7, scored: 2, total: 2, exact_pages: 0, paged: 1, page_error: 1 },
      libreoffice: null,
    },
    failures: [],
  });
  expect(aggregate.formats.pptx.samples).toBe(0);
  expect(aggregate.formats.pptx.fidelity.commit).toEqual({ ssim: null, scored: 0, total: 0, exact_pages: 0, paged: 0, page_error: 0 });
  const text = JSON.stringify(aggregate);
  for (const leak of ['demo', 'second', 'metadata', 'corpus.betteroffice.dev', 'sha256'])
    expect(text.includes(leak) && leak).toBe(false);
});

test('the aggregate refuses the same invalid reports the README refuses', () => {
  const input = report();
  input.commit = 'short';
  expect(() => aggregateReport(input)).toThrow('Expected a full commit SHA');
});

test('the aggregate counts failures by area, channel and stage without their messages', () => {
  const input = report();
  input.samples[0].comparisons[1] = {
    channel: 'commit', renderer_source_commit: commit, status: 'failed', stage: 'capture',
    error: 'navigation to https://corpus.betteroffice.dev/demo/source.docx failed',
  } as any;
  const aggregate = aggregateReport(input);
  expect(aggregate.formats.docx.failures).toEqual([{ area: 'render', channel: 'commit', stage: 'capture', count: 1 }]);
  const text = JSON.stringify(aggregate);
  for (const leak of ['demo', 'navigation', 'corpus.betteroffice.dev']) expect(text.includes(leak) && leak).toBe(false);
});

test('a private run headline names no per-document site and lists no document', () => {
  const input = report();
  const shown = renderSection(input, { privateCorpus: true });
  expect(shown).not.toContain('benchmarks.betteroffice.dev');
  expect(shown).toContain('aggregates only; documents are not public');
  for (const sample of input.samples) expect(shown).not.toContain(sample.id);
  expect(renderSection(input)).toContain('https://benchmarks.betteroffice.dev');
});

const three = (label: string, ...cells: string[]) =>
  `<tr><td>${label}</td>${cells.map((cell) => `<td align="right">${cell}</td>`).join('')}</tr>`;

test('reports PPTX slide agreement like DOCX pages, and no page rows for XLSX print ranges', () => {
  const input = report();
  const comparison = (channel: string, reference: number, actual: number) => ({
    source_verified: true, reference: { status: 'ok', sha256: 'source' }, actual: { status: 'ok', sha256: 'source' },
    penalized_ssim: 0.9, reference_pages: reference, actual_pages: actual, channel,
    ...(channel === 'commit' ? { renderer_source_commit: commit } : { version: channel === 'published' ? '0.0.4' : '0.1.0' }),
  });
  input.samples = [
    { id: 'p1', format: 'pptx', comparisons: [comparison('published', 12, 12), comparison('commit', 12, 12)] },
    { id: 'p2', format: 'pptx', comparisons: [comparison('published', 5, 7), comparison('commit', 5, 5)] },
    { id: 'x1', format: 'xlsx', comparisons: [comparison('published', 2, 2), comparison('commit', 2, 2)] },
  ] as any;
  const section = renderSection(input);
  const pptx = section.slice(section.indexOf('### PPTX'), section.indexOf('### XLSX'));
  expect(pptx).toContain(row('Exact page counts', '1/2', '2/2'));
  expect(pptx).toContain(row('Absolute page error', '2', '0'));
  const xlsx = section.slice(section.indexOf('### XLSX'));
  expect(xlsx).not.toContain('Exact page counts');
  expect(xlsx).not.toContain('Absolute page error');
});

test('the corpus line is generated from document and page counts, aggregates only', () => {
  const input: any = { ...report(), corpus: { docx: { documents: 1776, pages: 20204 }, pptx: { documents: 656, pages: 12477 }, xlsx: { documents: 3297, pages: 5384 } } };
  const section = renderSection(input, { privateCorpus: true });
  expect(section).toContain('Corpus: 1,776 DOCX documents (20,204 pages), 656 PPTX presentations (12,477 slides) and 3,297 XLSX workbooks (5,384 printed pages), each with a Microsoft Office reference.');
  expect(section).toContain('XLSX is compared per Excel print range, one range per page, so its page counts always match and are not reported.');
  const visual = renderSection(input);
  expect(visual.indexOf('Corpus: ')).toBeLessThan(visual.indexOf('<picture>'));
  expect(section).not.toContain('<picture>');
  expect(renderSection(report())).not.toContain('Corpus: ');
  expect(aggregateReport(input).formats.pptx.corpus).toEqual({ documents: 656, pages: 12477 });
  expect(() => renderSection({ ...input, corpus: { docx: { documents: 2, pages: 1 } } })).toThrow('Invalid corpus summary');
});

const sourceSha = 'c'.repeat(40);
const lo = '26.2.3.2';
function calculated(fully: Record<string, boolean>) {
  const result = (ok: boolean, ms: number) => ({ status: 'ok', correct: ok ? 4 : 3, total: 4, elapsed_ms: Array(5).fill(ms) });
  return { published: result(fully.published, 10), commit: result(fully.commit, 10), libreoffice: result(fully.libreoffice, 30) };
}
function xlsxReport() {
  const input: any = report();
  input.source_sha = sourceSha;
  input.xlsx_benchmark = {
    libreoffice_version: lo, published_version: '0.1.0', builds: { commit: { source_sha: sourceSha } },
    reference_coverage: { total: 4, measured: 2, no_formulas: ['n'], missing_reference: ['m'] },
  };
  input.samples.push(
    { id: 'w1', format: 'xlsx', comparisons: [], calculations: calculated({ published: true, commit: true, libreoffice: false }) },
    { id: 'w2', format: 'xlsx', comparisons: [], calculations: calculated({ published: false, commit: true, libreoffice: true }) },
  );
  return input;
}

test('the recalc note and the aggregate carry per-channel fully matching workbooks, generated from the data', () => {
  const section = renderSection(xlsxReport());
  expect(section).toContain(`Recalc rows: 2 workbooks with formulas (8 formula results), each recalculated from cleared caches and compared with the values Excel saved. Workbooks fully matching: BetterOffice 0.1.0 1/2, BetterOffice at ${commit.slice(0, 8)} 2/2, LibreOffice ${lo} 1/2. 1 further workbook has no Excel calculation reference and is not counted.`);
  expect(section).toContain(three('Recalc accuracy', '87.50%', '100.00%', '87.50%'));
  expect(section).not.toContain('Recalc workbooks fully matching');
  const recalc = aggregateReport(xlsxReport()).formats.xlsx.recalc;
  expect(recalc.commit).toEqual({ correct: 8, total: 8, workbooks: 2, fully_matching: 2, mean_ms: null });
  expect(recalc.libreoffice.fully_matching).toBe(1);
  expect(JSON.stringify(recalc)).not.toContain('w1');
});

function timed(mean: Record<string, number>) {
  const channel = (value: number) => ({ successful: 2, mean_ms: value, median_ms: value, round_means_ms: [value, value, value], round_spread: 0 });
  return {
    method: 'sequential-paired-sample-v1', rounds: 3, sample_size: 200, runner: { processors: 4 },
    summary: { format: 'xlsx', total: 2, common: 2,
      channels: { published: channel(mean.published), commit: channel(mean.commit), libreoffice: channel(mean.libreoffice) },
      ratio_to_libreoffice: { published: { geometric_mean: 0.5, ci95: [0.4, 0.6] }, commit: { geometric_mean: 0.5, ci95: [0.4, 0.6] } } },
    samples: [{ id: 'w1' }],
  };
}

test('private runs show time rows only from the sequential timing job, as the mean per document', () => {
  expect(renderSection(xlsxReport())).toContain('<td>Recalc time (avg)</td>');
  const untimed = renderSection(xlsxReport(), { privateCorpus: true });
  expect(untimed).not.toContain('Recalc time (avg)');
  expect(untimed).not.toContain('Render time (avg)');
  expect(untimed).not.toContain('average time per file');
  expect(untimed).not.toContain('Times: ');
  const input = xlsxReport();
  input.timing_benchmark = { xlsx: timed({ published: 120.4, commit: 99.6, libreoffice: 410 }) };
  const shown = renderSection(input, { privateCorpus: true });
  expect(shown).toContain(three('Recalc time (avg)', '120 ms', '100 ms', '410 ms'));
  expect(shown).not.toContain('<picture>');
  expect(shown).toContain(`Times: mean per document over a seeded sample of up to 200 documents per format. A document counts only when every engine timed it in every round: XLSX 2 of 2 sampled (timed per engine: BetterOffice 0.1.0 2, BetterOffice at ${commit.slice(0, 8)} 2, LibreOffice 2). They run one after another on one machine per format`);
  expect(shown).not.toContain('w1');
  const timing = aggregateReport(input).formats.xlsx.timing;
  expect(timing.channels.commit).toEqual({ successful: 2, mean_ms: 100, median_ms: 100, round_means_ms: [100, 100, 100], round_spread: 0 });
  expect(timing.ratio_to_libreoffice.commit).toEqual({ geometric_mean: 0.5, ci95: [0.4, 0.6] });
  expect(JSON.stringify(timing)).not.toContain('w1');
});

test('the timing note states how many timed documents had notes whose layout did not converge', () => {
  const input = xlsxReport();
  const docx = timed({ published: 300, commit: 200, libreoffice: 500 });
  docx.summary = { ...docx.summary, format: 'docx', notes_not_converged: 1 } as any;
  input.timing_benchmark = { docx };
  const shown = renderSection(input, { privateCorpus: true });
  expect(shown).toContain(`DOCX 2 of 2 sampled (timed per engine: BetterOffice 0.1.0 2, BetterOffice at ${commit.slice(0, 8)} 2, LibreOffice 2, 1 of them with notes whose layout did not converge, timed as the product renders them)`);
  expect(aggregateReport(input).formats.docx.timing.notes_not_converged).toBe(1);
});

function fontsReport() {
  const input = report();
  const published = { package: '@betteroffice/fonts@0.4.2', cjk: '@betteroffice/fonts-cjk@0.4.2', copy: 'local release copy' };
  const workspace = { package: '@betteroffice/fonts@0.5.0', copy: 'workspace', base_latin: 4 };
  const second = {
    ...input.samples[0],
    id: 'second',
    metadata_url: 'https://corpus.betteroffice.dev/second/metadata.json',
    comparisons: input.samples[0].comparisons.map((entry) => ({ ...entry })),
  };
  input.samples.push(second);
  const [first] = input.samples;
  Object.assign(first.comparisons[0], { fonts: { ...published, advertised: 90, loaded_faces: ['Carlito-Regular.ttf', 'Caladea-Bold.ttf'], failed_faces: { 'NotoSansJP-Regular.otf': 'http 404' } } });
  Object.assign(second.comparisons[0], { fonts: { ...published, advertised: 90, loaded_faces: ['Carlito-Regular.ttf', 'NotoSansJP-Regular.otf'], failed_faces: { 'Gelasio-Bold.ttf': 'blocked', 'Inter-Regular.ttf': 'error' } } });
  Object.assign(first.comparisons[1], { fonts: { ...workspace, advertised: 120, loaded_faces: ['Carlito-Regular.ttf.zst'], failed_faces: {} } });
  Object.assign(second.comparisons[1], { fonts: { ...workspace, advertised: 120, loaded_faces: ['Carlito-Regular.ttf.zst', 'Caladea-Bold.ttf.zst'], failed_faces: {} } });
  return input;
}

const noFontCounts = (section: string) => {
  for (const text of ['did not load', 'M = ', 'by reason', 'not measured', 'nothing was measured', 'unattributed', 'face loads'])
    expect(section.includes(text) && text).toBe(false);
};

test('the section names each arm font source only, without face counts', () => {
  const section = renderSection(fontsReport());
  expect(section).toContain('\n- Published fonts: @betteroffice/fonts@0.4.2 and @betteroffice/fonts-cjk@0.4.2 (local release copy)\n');
  expect(section).toContain('\n- Commit fonts: @betteroffice/fonts@0.5.0 (workspace)\n');
  noFontCounts(section);
  expect(renderSection(report())).not.toContain('fonts:');
});

test('PPTX and XLSX arms name their font source only', () => {
  const input = report();
  input.samples.push({
    ...input.samples[0],
    id: 'deck',
    format: 'pptx',
    comparisons: [
      { ...input.samples[0].comparisons[0], version: '0.0.4', fonts: { package: '@betteroffice/fonts@0.4.2', cjk: '@betteroffice/fonts-cjk@0.4.2', copy: 'local release copy' } },
      { ...input.samples[0].comparisons[1], fonts: { package: '@betteroffice/fonts@0.5.0', copy: 'workspace' } },
    ],
  } as any);
  const section = renderSection(input);
  expect(section).toContain('\n- Published fonts: @betteroffice/fonts@0.4.2 and @betteroffice/fonts-cjk@0.4.2 (local release copy)\n');
  expect(section).toContain('\n- Commit fonts: @betteroffice/fonts@0.5.0 (workspace)\n');
  noFontCounts(section);
});

test('the aggregate counts advertised, loaded and failed faces per arm by reason, without names', () => {
  const { fonts } = aggregateReport(fontsReport()).formats.docx;
  expect(fonts).toEqual({
    published: {
      sources: ['@betteroffice/fonts@0.4.2 and @betteroffice/fonts-cjk@0.4.2 (local release copy)'],
      captures: 2,
      missing: 0,
      advertised: 90,
      base_latin: 0,
      requested: 5,
      loaded: 2,
      failed: 3,
      reasons: { 'http 404': 1, blocked: 1, error: 1 },
      unattributed: { loaded: 0, failed: 0 },
      not_measured: { timed_out: 0, errored: 0 },
    },
    commit: {
      sources: ['@betteroffice/fonts@0.5.0 (workspace)'],
      captures: 2,
      missing: 0,
      advertised: 120,
      base_latin: 4,
      requested: 2,
      loaded: 2,
      failed: 0,
      reasons: {},
      unattributed: { loaded: 0, failed: 0 },
      not_measured: { timed_out: 0, errored: 0 },
    },
  });
  const text = JSON.stringify(aggregateReport(fontsReport()));
  for (const leak of ['Carlito', 'Gelasio', 'NotoSansJP', 'second', 'demo'])
    expect(text.includes(leak) && leak).toBe(false);
  expect(aggregateReport(report()).formats.docx.fonts).toBeUndefined();
});

test('the aggregate counts a roundtrip stage or a font reason outside the fixed sets as other', async () => {
  const { fontArmLine, fontSummary } = await import('./readme.mjs');
  const input = fontsReport();
  Object.assign(input.samples[0], { roundtrip: { published: { parse: { status: 'ok' }, roundtrip: { status: 'failed', stage: '/tmp/x/y.ttf', error: 'failed' } } } });
  (input.samples[1] as any).comparisons[0].fonts.failed_faces['Inter-Regular.ttf'] = 'host-a /tmp/x/y.ttf';
  const aggregate = aggregateReport(input);
  expect(aggregate.formats.docx.fonts.published.reasons).toEqual({ 'http 404': 1, blocked: 1, other: 1 });
  expect(aggregate.formats.docx.failures).toEqual([{ area: 'roundtrip', channel: 'published', stage: 'other', count: 1 }]);
  const text = JSON.stringify([aggregate, renderSection(input), fontArmLine('published', fontSummary(input.samples, 'published'))]);
  for (const leak of ['host-a', '/tmp/x/y.ttf']) expect(text.includes(leak) && leak).toBe(false);
});

test('a font count that is not a non-negative integer fails naming the job and the field, never the value', () => {
  for (const [field, value] of [['loaded', 'host-a'], ['failed', -3]] as const) {
    const input = fontsReport();
    (input.samples[1] as any).comparisons[0].fonts.unattributed = { loaded: 0, failed: 0, [field]: value };
    let message = '';
    try {
      aggregateReport(input);
    } catch (error) {
      message = (error as Error).message;
    }
    expect(message).toContain('docx measure job');
    expect(message).toContain(`fonts.unattributed.${field}`);
    expect(message.includes(String(value)) && String(value)).toBe(false);
  }
});

test('the arm line states how many advertised faces did not load and why', async () => {
  const { fontArmLine, fontSummary } = await import('./readme.mjs');
  const { samples } = fontsReport();
  expect(fontArmLine('published', fontSummary(samples, 'published'))).toBe(
    "published arm: 3 of 5 faces did not load (M = faces the documents requested that the arm's provider offers, base latin: none in this release), by reason: blocked 1, error 1, http 404 1; sliced faces are observed only through their whole-face fallback at open; a chunk that fails or times out after open paints blank glyphs and is not observed"
  );
  expect(fontArmLine('commit', fontSummary(samples, 'commit'))).toBe(
    "commit arm: 0 of 2 faces did not load (M = faces the documents requested that the arm's provider offers, base latin included), by reason: none; sliced faces are observed only through their whole-face fallback at open; a chunk that fails or times out after open paints blank glyphs and is not observed"
  );
});

const VISUAL_BEGIN = '<!-- BEGIN GENERATED VISUAL FIDELITY -->';
const VISUAL_END = '<!-- END GENERATED VISUAL FIDELITY -->';
const LARGE_BEGIN = '<!-- BEGIN GENERATED LARGE SET -->';
const LARGE_END = '<!-- END GENERATED LARGE SET -->';
const DATE = '2026-11-17';
const span = (text: string, begin: string, end: string) => text.slice(text.indexOf(begin), text.indexOf(end) + end.length);
const shell = (blocks: string) => `# Project\n\n## Packages\n\nList.\n\n## Fonts\n\nFonts.\n\n${blocks}\n\n## Evals\n\nEvals.\n`;
function later() {
  const input = report();
  input.commit = 'b'.repeat(40);
  (input.samples[0].comparisons[1] as any).renderer_source_commit = input.commit;
  return input;
}

test('a private run replaces only the large set block and leaves the public block byte-identical', () => {
  const visual = renderSection(report());
  const readme = shell(`${visual}\n\n${LARGE_BEGIN}\nold large set\n${LARGE_END}`);
  const large = renderSection(later(), { privateCorpus: true, date: DATE });
  const updated = updateReadme(readme, large);
  expect(span(updated, VISUAL_BEGIN, VISUAL_END)).toBe(visual);
  expect(updated).toBe(shell(`${visual}\n\n${large}`));
});

test('a public run replaces its block in place and leaves the large set block byte-identical', () => {
  const large = renderSection(report(), { privateCorpus: true, date: DATE });
  const readme = shell(`${renderSection(report())}\n\n${large}`);
  const visual = renderSection(later());
  const updated = updateReadme(readme, visual);
  expect(span(updated, LARGE_BEGIN, LARGE_END)).toBe(large);
  expect(updated).toBe(shell(`${visual}\n\n${large}`));
});

test('a public run refuses a README that has only the large set block and leaves it unchanged', async () => {
  const large = renderSection(report(), { privateCorpus: true, date: DATE });
  const readme = shell(large);
  expect(() => updateReadme(readme, renderSection(later()))).toThrow('README has no generated block');
  const dir = await mkdtemp(join(tmpdir(), 'readme-public-without-pair-'));
  try {
    const path = join(dir, 'README.md');
    await writeFile(path, readme);
    await writeFile(join(dir, 'public.md'), renderSection(later()));
    refused(cli(['--section', 'public.md', 'README.md'], dir, false), 'README has no generated block');
    expect(await readFile(path, 'utf8')).toBe(readme);
  } finally {
    await rm(dir, { recursive: true, force: true });
  }
});

test('a README without a large set block gets it right after the public block, and a second private run replaces it', async () => {
  const visual = renderSection(report());
  const first = renderSection(report(), { privateCorpus: true, date: DATE });
  const inserted = updateReadme(shell(visual), first);
  expect(inserted).toBe(shell(`${visual}\n\n${first}`));
  const second = renderSection(later(), { privateCorpus: true, date: '2026-11-18' });
  const replaced = updateReadme(inserted, second);
  expect(replaced).toBe(shell(`${visual}\n\n${second}`));
  expect(replaced.split(LARGE_BEGIN).length).toBe(2);
});

test('a README shaped like the published one gets the large set block right after its badged public block, and a later run replaces it', () => {
  const visual = badged(renderSection(report()), DATE);
  const first = renderSection(report(), { privateCorpus: true, date: DATE });
  const second = renderSection(later(), { privateCorpus: true, date: '2026-11-18' });
  const head = '# BetterOffice\n\n## Packages\n\n| Package | Version |\n| --- | --- |\n| docx | 0.1.0 |\n\n## Fonts\n\nFonts.\n\n';
  const tail = '\n\n## Development\n\nCommands.\n\n## Evals\n\nEvals.\n';
  const fixture = `${head}${visual}${tail}`;
  expect(updateReadme(fixture, first)).toBe(`${head}${visual}\n\n${first}${tail}`);
  expect(updateReadme(`${head}${visual}\n\n${first}${tail}`, second)).toBe(`${head}${visual}\n\n${second}${tail}`);
});

test('refuses a README with a repeated or broken pair of either kind', () => {
  const visual = renderSection(report());
  const large = renderSection(report(), { privateCorpus: true, date: DATE });
  expect(() => updateReadme(shell(`${visual}\n\n${large}\n\n${large}`), large)).toThrow('Broken generated markers');
  expect(() => updateReadme(shell(`${LARGE_BEGIN}\n${visual}`), visual)).toThrow('Broken generated markers');
});

test('the large set section carries one generated line and aggregates only: no ids, site link or chart', () => {
  const input: any = report();
  input.samples.push({
    id: 'hidden-name', format: 'pptx', metadata_url: 'https://corpus.betteroffice.dev/hidden-name/metadata.json',
    comparisons: [{ channel: 'commit', renderer_source_commit: commit, status: 'failed', stage: 'capture',
      error: 'navigation to https://corpus.betteroffice.dev/hidden-name/source.pptx failed' }],
  });
  const section = renderSection(input, { privateCorpus: true, date: DATE });
  expect(section.startsWith(LARGE_BEGIN)).toBe(true);
  expect(section.endsWith(LARGE_END)).toBe(true);
  expect(section).toContain('\n## Benchmarks: large set\n');
  expect(section.split('\n').filter((line) => line.startsWith('*Generated'))).toEqual([
    `*Generated by the [Benchmarks action](.github/workflows/visual-fidelity.yml) from 1 DOCX, 1 PPTX and 0 XLSX documents at <a href="https://github.com/openooxml/betteroffice/commit/${commit}">${commit.slice(0, 8)}</a> on ${DATE}; aggregates only; documents are not public.*`,
  ]);
  for (const leak of ['demo', 'hidden-name', 'metadata', 'navigation', 'benchmarks.betteroffice.dev', 'api/charts', '<picture>', 'GENERATED VISUAL FIDELITY'])
    expect(section.includes(leak) && leak).toBe(false);
  expect(() => renderSection(input, { privateCorpus: true, date: '17.11.2026' })).toThrow('Invalid date');
});

const script = join(import.meta.dir, 'readme.mjs');
const cli = (args: string[], cwd: string, privateRun: boolean) =>
  spawnSync(process.execPath, [script, ...args], { cwd, encoding: 'utf8', env: { ...process.env, QUALITY_PRIVATE: privateRun ? 'true' : 'false' } });
const refused = (result: { status: number | null; stderr: string }, text: string) =>
  expect({ refused: result.status !== 0, stderr: result.stderr }).toEqual({ refused: true, stderr: expect.stringContaining(text) });

const largeRenderer = async () => ((await import('./readme.mjs')) as any).renderLargeSection;
const PRIVATE_SECTION = 'a private run writes the large set block only from aggregate.json';

test('the CLI writes the public block from --section and the large set block from aggregate.json, each in its own run kind', async () => {
  const dir = await mkdtemp(join(tmpdir(), 'readme-cli-'));
  try {
    const renderLargeSection = await largeRenderer();
    const visual = renderSection(report());
    const next = renderSection(later());
    const aggregate = merged(fullReport());
    const large = renderLargeSection(aggregate, { date: DATE });
    const readme = shell(visual);
    const path = join(dir, 'README.md');
    await writeFile(path, readme);
    await writeFile(join(dir, 'public.md'), next);
    await writeFile(join(dir, 'large.md'), renderSection(fullReport(), { privateCorpus: true, date: DATE }));
    await writeFile(join(dir, 'aggregate.json'), `${JSON.stringify(aggregate, null, 2)}\n`);
    refused(cli(['--section', 'public.md', 'README.md'], dir, true), PRIVATE_SECTION);
    refused(cli(['--section', 'large.md', 'README.md'], dir, true), PRIVATE_SECTION);
    refused(cli(['--section', 'large.md', 'README.md'], dir, false), 'The generated section does not match this run');
    expect(await readFile(path, 'utf8')).toBe(readme);
    expect(cli(['--as', 'large', '--date', DATE, 'aggregate.json', 'README.md'], dir, true).status).toBe(0);
    expect(await readFile(path, 'utf8')).toBe(shell(`${visual}\n\n${large}`));
    expect(cli(['--section', 'public.md', 'README.md'], dir, false).status).toBe(0);
    expect(await readFile(path, 'utf8')).toBe(shell(`${next}\n\n${large}`));
    await symlink('README.md', join(dir, 'link.md'));
    refused(cli(['--as', 'large', '--date', DATE, 'aggregate.json', 'link.md'], dir, true), 'must be a regular file');
  } finally {
    await rm(dir, { recursive: true, force: true });
  }
});

test('--public is refused in either run kind and the README is unchanged', async () => {
  const dir = await mkdtemp(join(tmpdir(), 'readme-no-public-option-'));
  try {
    const visual = renderSection(report());
    const stale = `${VISUAL_BEGIN}\n## Benchmarks\n\n*Generated by the [Benchmarks action](.github/workflows/visual-fidelity.yml); per-document results are not published.*\n\nOlder private numbers.\n\n${VISUAL_END}`;
    const path = join(dir, 'README.md');
    await writeFile(path, shell(stale));
    await writeFile(join(dir, 'main.md'), shell(visual));
    await writeFile(join(dir, 'aggregate.json'), `${JSON.stringify(merged(later()), null, 2)}\n`);
    refused(cli(['--as', 'small', 'aggregate.json', 'README.md'], dir, false), 'usage: readme.mjs');
    for (const privateRun of [false, true])
      refused(cli(['--as', 'large', '--public', 'main.md', '--date', DATE, 'aggregate.json', 'README.md'], dir, privateRun), '--public');
    expect(await readFile(path, 'utf8')).toBe(shell(stale));
  } finally {
    await rm(dir, { recursive: true, force: true });
  }
});

function fullReport() {
  const input: any = xlsxReport();
  input.corpus = { docx: { documents: 1776, pages: 20204 }, xlsx: { documents: 3297, pages: 5384 } };
  input.xlsx_fidelity_benchmark = { libreoffice_version: lo, source_sha: sourceSha };
  input.roundtrip_benchmark = { xlsx: { libreoffice_version: lo, published_version: '0.1.0', builds: { commit: { source_sha: sourceSha } } } };
  const scored = (penalized_ssim: number) => ({ source_verified: true, reference: { status: 'ok', sha256: 'source' }, actual: { status: 'ok', sha256: 'source' }, penalized_ssim });
  const parsed = { parse: { status: 'ok' }, roundtrip: { status: 'failed', error: 'changed parts' } };
  const unreadable = { parse: { status: 'failed', error: 'unreadable' }, roundtrip: { status: 'failed', error: 'unreadable' } };
  input.samples.filter((sample: any) => sample.format === 'xlsx').forEach((sample: any, index: number) => {
    sample.comparisons = [
      { ...scored(0.91 + index / 100), channel: 'published', version: '0.1.0' },
      { ...scored(0.95), channel: 'commit', renderer_source_commit: commit },
      { ...scored(0.85005), channel: 'libreoffice', version: lo },
    ];
    sample.roundtrip = { published: parsed, commit: parsed, libreoffice: index ? unreadable : parsed };
  });
  const docx = timed({ published: 300.5, commit: 199.49, libreoffice: 500 });
  docx.summary = { ...docx.summary, format: 'docx', notes_not_converged: 1 } as any;
  input.timing_benchmark = { docx, xlsx: timed({ published: 120.4, commit: 99.6, libreoffice: 410 }) };
  return input;
}

const written = (input: any) => JSON.parse(`${JSON.stringify(aggregateReport(input), null, 2)}\n`);
const merged = (input: any) => ({ ...written(input), run_kind: 'private' });

test('the large set block renders byte-identical from the aggregate.json the merge writes and from the full report', async () => {
  const { renderLargeSection } = (await import('./readme.mjs')) as any;
  const failedPptx: any = report();
  failedPptx.samples.push({ id: 'hidden-name', format: 'pptx', comparisons: [{ channel: 'commit', renderer_source_commit: commit, status: 'failed', stage: 'capture', error: 'navigation failed' }] });
  for (const input of [report(), later(), fontsReport(), xlsxReport(), failedPptx, fullReport()])
    expect(renderLargeSection(written(input), { date: DATE })).toBe(renderSection(input, { privateCorpus: true, date: DATE }));
  const shown = renderSection(fullReport(), { privateCorpus: true, date: DATE });
  for (const text of ['Corpus: 1,776 DOCX documents', `<th width="230" align="right">LibreOffice (${lo})</th>`, '<td>Parse success</td>', '<td>Render time (avg)</td>',
    '<td>Recalc time (avg)</td>', '1 further workbook has no Excel calculation reference', 'up to 200 documents', '1 of them with notes', '<td align="right">0.9150</td>'])
    expect(shown.includes(text) && text).toBe(text);
  expect(() => renderLargeSection(written(fullReport()), { date: '17.11.2026' })).toThrow('Invalid date');
});

test('an aggregate.json without a field the large set block shows is refused, naming the field', async () => {
  const { renderLargeSection } = (await import('./readme.mjs')) as any;
  const unsized = written(fullReport());
  delete unsized.formats.docx.timing.sample_size;
  expect(() => renderLargeSection(unsized, { date: DATE })).toThrow('docx timing.sample_size');
  expect(() => renderLargeSection(unsized, { date: DATE })).toThrow('--timing-sample-size');
  const uncovered = written(fullReport());
  delete uncovered.formats.xlsx.recalc_missing_reference;
  expect(() => renderLargeSection(uncovered, { date: DATE })).toThrow('xlsx recalc_missing_reference');
  expect(() => renderLargeSection(uncovered, { date: DATE })).toThrow('--missing-reference');
  expect(() => renderLargeSection({ ...written(report()), schema_version: 2 }, { date: DATE })).toThrow('schema version 1');
});

test('--as large renders the block from the run artifact aggregate.json and refuses a report.json', async () => {
  const dir = await mkdtemp(join(tmpdir(), 'readme-aggregate-'));
  try {
    const visual = renderSection(report());
    const stored = fullReport();
    const path = join(dir, 'README.md');
    await writeFile(path, shell(visual));
    await writeFile(join(dir, 'report.json'), JSON.stringify({ ...stored, run_kind: 'private' }));
    await writeFile(join(dir, 'aggregate.json'), `${JSON.stringify(aggregateReport(stored), null, 2)}\n`);
    for (const privateRun of [false, true])
      refused(cli(['--as', 'large', '--date', DATE, 'report.json', 'README.md'], dir, privateRun), 'only from an aggregate.json');
    expect(await readFile(path, 'utf8')).toBe(shell(visual));
    expect(cli(['--as', 'large', '--date', DATE, 'aggregate.json', 'README.md'], dir, false).status).toBe(0);
    expect(await readFile(path, 'utf8')).toBe(shell(`${visual}\n\n${renderSection(stored, { privateCorpus: true, date: DATE })}`));
  } finally {
    await rm(dir, { recursive: true, force: true });
  }
});

test('--as large takes the timing sample size and missing reference count as arguments for an aggregate.json without them', async () => {
  const dir = await mkdtemp(join(tmpdir(), 'readme-old-aggregate-'));
  try {
    const visual = renderSection(report());
    const stored = fullReport();
    const expected = shell(`${visual}\n\n${renderSection(stored, { privateCorpus: true, date: DATE })}`);
    const old = written(stored);
    for (const format of ['docx', 'xlsx']) delete old.formats[format].timing.sample_size;
    delete old.formats.xlsx.recalc_missing_reference;
    await writeFile(join(dir, 'old.json'), `${JSON.stringify(old, null, 2)}\n`);
    await writeFile(join(dir, 'new.json'), `${JSON.stringify(aggregateReport(stored), null, 2)}\n`);
    const path = join(dir, 'README.md');
    const run = (args: string[]) => cli(['--as', 'large', '--date', DATE, ...args, 'README.md'], dir, false);

    await writeFile(path, shell(visual));
    refused(run(['old.json']), '--timing-sample-size');
    refused(run(['--missing-reference', '1', 'old.json']), '--timing-sample-size');
    refused(run(['--timing-sample-size', '200', 'old.json']), '--missing-reference');
    expect(await readFile(path, 'utf8')).toBe(shell(visual));
    expect(run(['--timing-sample-size', '200', '--missing-reference', '1', 'old.json']).status).toBe(0);
    expect(await readFile(path, 'utf8')).toBe(expected);

    for (const args of [['--timing-sample-size', '199'], ['--missing-reference', '2']]) {
      await writeFile(path, shell(visual));
      refused(run([...args, 'new.json']), args[0]);
      expect(await readFile(path, 'utf8')).toBe(shell(visual));
    }
    for (const args of [['--timing-sample-size', '200'], ['--missing-reference', '1'], ['--timing-sample-size', '200', '--missing-reference', '1']]) {
      await writeFile(path, shell(visual));
      expect(run([...args, 'new.json']).status).toBe(0);
      expect(await readFile(path, 'utf8')).toBe(expected);
    }

    await writeFile(path, shell(visual));
    for (const value of ['-1', '1.5', 'abc', '', '1e2', ' 1'])
      for (const [name, other] of [['--timing-sample-size', ['--missing-reference=1']], ['--missing-reference', ['--timing-sample-size=200']]] as const)
        refused(run([`${name}=${value}`, ...other, 'old.json']), name);
    expect(await readFile(path, 'utf8')).toBe(shell(visual));
  } finally {
    await rm(dir, { recursive: true, force: true });
  }
});

const RUN = 'https://github.com/openooxml/betteroffice/actions/runs/44216627357';
const runBadge = (date: string) =>
  `<a href="${RUN}"><img src="https://betteroffice.dev/api/badge?label=benchmark%20run&amp;message=${date}&amp;color=0a0a0a" alt="benchmark run ${date}"></a>`;
const badged = (section: string, date: string) => {
  const lines = section.split('\n');
  const generated = lines.findIndex((line) => line.startsWith('*Generated by'));
  return [...lines.slice(0, generated + 1), '', runBadge(date), ...lines.slice(generated + 1)].join('\n');
};
const utcDay = () => new Date().toISOString().slice(0, 10);

test('the CLI puts a badge linking the run into the block it writes when given the run URL, and none without it', async () => {
  const dir = await mkdtemp(join(tmpdir(), 'readme-run-badge-'));
  try {
    const visual = renderSection(report());
    const next = renderSection(later());
    const large = renderSection(report(), { privateCorpus: true, date: DATE });
    const path = join(dir, 'README.md');
    await writeFile(join(dir, 'public.md'), next);
    await writeFile(join(dir, 'aggregate.json'), `${JSON.stringify(merged(report()), null, 2)}\n`);

    await writeFile(path, shell(visual));
    expect(cli(['--as', 'large', '--date', DATE, 'aggregate.json', 'README.md'], dir, true).status).toBe(0);
    const plain = await readFile(path, 'utf8');
    expect(plain).toBe(shell(`${visual}\n\n${large}`));
    for (const absent of ['actions/runs', 'benchmark%20run']) expect(plain.includes(absent) && absent).toBe(false);

    await writeFile(path, shell(visual));
    expect(cli(['--as', 'large', '--date', DATE, '--run-url', RUN, 'aggregate.json', 'README.md'], dir, true).status).toBe(0);
    const largeBadged = badged(large, DATE);
    expect(largeBadged.split(RUN).length).toBe(2);
    expect(largeBadged.indexOf(RUN)).toBeLessThan(largeBadged.indexOf('<details>'));
    expect(await readFile(path, 'utf8')).toBe(shell(`${visual}\n\n${largeBadged}`));

    const before = utcDay();
    expect(cli(['--section', 'public.md', '--run-url', RUN, 'README.md'], dir, false).status).toBe(0);
    const after = utcDay();
    const updated = await readFile(path, 'utf8');
    const day = updated.match(/benchmark run (\d{4}-\d{2}-\d{2})/)?.[1] ?? '';
    expect([before, after]).toContain(day);
    expect(updated).toBe(shell(`${badged(next, day)}\n\n${largeBadged}`));

    await writeFile(path, shell(visual));
    refused(cli(['--as', 'large', '--date', DATE, '--run-url', 'https://example.com/actions/runs/1', 'aggregate.json', 'README.md'], dir, true), '--run-url must be');
    expect(await readFile(path, 'utf8')).toBe(shell(visual));
  } finally {
    await rm(dir, { recursive: true, force: true });
  }
});

test('--as large takes --run-url for the badge and refuses a URL that is not a GitHub actions run', async () => {
  const dir = await mkdtemp(join(tmpdir(), 'readme-repair-badge-'));
  try {
    const visual = renderSection(report());
    const stored = fullReport();
    const expected = shell(`${visual}\n\n${badged(renderSection(stored, { privateCorpus: true, date: DATE }), DATE)}`);
    await writeFile(join(dir, 'aggregate.json'), `${JSON.stringify(aggregateReport(stored), null, 2)}\n`);
    const path = join(dir, 'README.md');
    await writeFile(path, shell(visual));
    expect(cli(['--as', 'large', '--date', DATE, '--run-url', RUN, 'aggregate.json', 'README.md'], dir, false).status).toBe(0);
    expect(await readFile(path, 'utf8')).toBe(expected);
    await writeFile(path, shell(visual));
    for (const url of [
      '',
      'http://github.com/openooxml/betteroffice/actions/runs/44216627357',
      'https://github.com/openooxml/betteroffice/actions/runs/',
      'https://github.com/openooxml/betteroffice/actions/runs/44216627357/job/1',
      'https://github.com/openooxml/betteroffice/actions/runs/abc',
      'https://github.com/openooxml/betteroffice/pull/1',
      'https://github.com.example.com/openooxml/betteroffice/actions/runs/1',
      'https://example.com/openooxml/betteroffice/actions/runs/1',
      'https://github.com/openooxml/betteroffice/actions/runs/1"><img src=x>',
      ` ${RUN}`,
      `${RUN}\n`,
    ])
      refused(cli(['--as', 'large', '--date', DATE, `--run-url=${url}`, 'aggregate.json', 'README.md'], dir, false), '--run-url must be');
    expect(await readFile(path, 'utf8')).toBe(shell(visual));
  } finally {
    await rm(dir, { recursive: true, force: true });
  }
});

test('updateReadme refuses a section that carries markers other than its own outer pair', () => {
  const visual = renderSection(report());
  const large = renderSection(report(), { privateCorpus: true, date: DATE });
  const readme = shell(`${visual}\n\n${large}`);
  for (const section of [
    `${LARGE_BEGIN}\n${visual}\n${LARGE_END}`,
    large.replace('<details>', `${LARGE_END}\n\n${LARGE_BEGIN}\n<details>`),
    visual.replace('<details>', `${LARGE_BEGIN}\nOther block.\n${LARGE_END}\n\n<details>`),
    visual.replace('<details>', `${VISUAL_END}\n\n${VISUAL_BEGIN}\n<details>`),
    `${VISUAL_BEGIN}\n${large}\n${VISUAL_END}`,
  ])
    expect(() => updateReadme(readme, section)).toThrow('Invalid generated section');
});

test('the CLI --section refuses a public section that carries the large set pair and leaves the README unchanged', async () => {
  const dir = await mkdtemp(join(tmpdir(), 'readme-nested-'));
  try {
    const visual = renderSection(report());
    const readme = shell(`${visual}\n\n${renderSection(report(), { privateCorpus: true, date: DATE })}`);
    const path = join(dir, 'README.md');
    await writeFile(path, readme);
    await writeFile(join(dir, 'nested.md'), `${VISUAL_BEGIN}\n${renderSection(later(), { privateCorpus: true, date: DATE })}\n${VISUAL_END}`);
    refused(cli(['--section', 'nested.md', 'README.md'], dir, false), 'Invalid generated section');
    expect(await readFile(path, 'utf8')).toBe(readme);
  } finally {
    await rm(dir, { recursive: true, force: true });
  }
});

test('either run kind refuses a README without the generated block it needs and leaves it unchanged', async () => {
  const dir = await mkdtemp(join(tmpdir(), 'readme-unmarked-'));
  try {
    const readme = '# P\n\n## Packages\n\nL.\n\n## Visual fidelity\n\nHand text.\n\n## Evals\n';
    const path = join(dir, 'README.md');
    await writeFile(path, readme);
    await writeFile(join(dir, 'public.md'), renderSection(later()));
    await writeFile(join(dir, 'aggregate.json'), `${JSON.stringify(merged(report()), null, 2)}\n`);
    for (const [args, privateRun] of [[['--as', 'large', '--date', DATE, 'aggregate.json'], true], [['--section', 'public.md'], false]] as const) {
      refused(cli([...args, 'README.md'], dir, privateRun), 'README has no generated block');
      expect(await readFile(path, 'utf8')).toBe(readme);
    }
  } finally {
    await rm(dir, { recursive: true, force: true });
  }
});

const STORED_PATH = '/private/acme.bin';
const assign = (target: any, path: string, value: unknown) => {
  const keys = path.split('.');
  keys.slice(0, -1).reduce((node, key) => node[key], target)[keys.at(-1)!] = value;
};
const refusesNaming = (render: () => unknown, field: string, value: unknown) => {
  let message = '';
  try {
    render();
  } catch (error) {
    message = String((error as Error)?.message ?? error);
  }
  expect({ field, named: message.includes(field), leaked: typeof value === 'string' && message.includes(value) })
    .toEqual({ field, named: true, leaked: false });
};

test('the large set renderers refuse a stored count, round or font label that is not valid, naming the field and not the value', async () => {
  const { renderLargeSection } = (await import('./readme.mjs')) as any;
  const fields: [string, string, unknown][] = [
    ['docx', 'samples', STORED_PATH],
    ['docx', 'fidelity.published.scored', STORED_PATH],
    ['docx', 'fidelity.published.scored', -1],
    ['docx', 'fidelity.published.total', STORED_PATH],
    ['docx', 'fidelity.commit.exact_pages', STORED_PATH],
    ['docx', 'fidelity.commit.paged', STORED_PATH],
    ['docx', 'fidelity.commit.page_error', STORED_PATH],
    ['docx', 'timing.rounds', STORED_PATH],
    ['docx', 'timing.rounds', 2.5],
    ['docx', 'timing.sample', STORED_PATH],
    ['docx', 'timing.timed_by_every_engine', STORED_PATH],
    ['docx', 'timing.notes_not_converged', STORED_PATH],
    ['docx', 'timing.channels.commit.successful', STORED_PATH],
    ['docx', 'timing.channels.commit.mean_ms', STORED_PATH],
    ['xlsx', 'timing.rounds', STORED_PATH],
    ['xlsx', 'recalc.commit.workbooks', STORED_PATH],
    ['xlsx', 'recalc.commit.correct', STORED_PATH],
    ['xlsx', 'recalc.commit.total', STORED_PATH],
    ['xlsx', 'recalc.commit.fully_matching', STORED_PATH],
    ['xlsx', 'roundtrip.commit.parsed', STORED_PATH],
    ['xlsx', 'roundtrip.commit.total', STORED_PATH],
  ];
  for (const [format, path, value] of fields) {
    const aggregate = written(fullReport());
    assign(aggregate, `formats.${format}.${path}`, value);
    refusesNaming(() => renderLargeSection(aggregate, { date: DATE }), `${format} ${path}`, value);
  }
  for (const label of [`${STORED_PATH} (workspace)`, `@betteroffice/fonts@0.5.0 (${STORED_PATH})`, `@betteroffice/fonts@0.4.2 and ${STORED_PATH} (local release copy)`]) {
    const aggregate = written(fontsReport());
    aggregate.formats.docx.fonts.published.sources = [label];
    refusesNaming(() => renderLargeSection(aggregate, { date: DATE }), 'docx fonts.published.sources', STORED_PATH);
  }
  for (const path of ['rounds', 'sample_size', 'summary.total', 'summary.common', 'summary.channels.commit.successful', 'summary.channels.commit.mean_ms']) {
    const input = fullReport();
    assign(input.timing_benchmark.docx, path, STORED_PATH);
    refusesNaming(() => renderSection(input, { privateCorpus: true, date: DATE }), `timing_benchmark.docx.${path}`, STORED_PATH);
  }
  for (const key of ['package', 'cjk', 'copy']) {
    const input: any = fontsReport();
    input.samples[0].comparisons[0].fonts[key] = STORED_PATH;
    refusesNaming(() => renderSection(input, { privateCorpus: true, date: DATE }), `fonts.${key}`, STORED_PATH);
  }
});

test('--as public is refused for a report or an aggregate in either run kind, and the README is unchanged', async () => {
  const dir = await mkdtemp(join(tmpdir(), 'readme-no-public-positional-'));
  try {
    const visual = renderSection(report());
    const path = join(dir, 'README.md');
    await writeFile(path, shell(visual));
    await writeFile(join(dir, 'report.json'), JSON.stringify(later()));
    await writeFile(join(dir, 'tagged.json'), JSON.stringify({ ...later(), run_kind: 'public' }));
    await writeFile(join(dir, 'aggregate.json'), `${JSON.stringify({ ...written(later()), run_kind: 'public' }, null, 2)}\n`);
    for (const privateRun of [false, true])
      for (const file of ['report.json', 'tagged.json', 'aggregate.json'])
        refused(cli(['--as', 'public', file, 'README.md'], dir, privateRun), 'the public block is written only by --section');
    expect(await readFile(path, 'utf8')).toBe(shell(visual));
  } finally {
    await rm(dir, { recursive: true, force: true });
  }
});

test('the positional CLI refuses a call without --as and leaves the README unchanged', async () => {
  const dir = await mkdtemp(join(tmpdir(), 'readme-positional-as-'));
  try {
    const visual = renderSection(report());
    const path = join(dir, 'README.md');
    await writeFile(path, shell(visual));
    await writeFile(join(dir, 'report.json'), JSON.stringify(later()));
    await writeFile(join(dir, 'aggregate.json'), `${JSON.stringify(aggregateReport(fullReport()), null, 2)}\n`);
    for (const privateRun of [false, true])
      for (const args of [['report.json'], ['report.json', 'README.md'], ['--date', DATE, 'aggregate.json', 'README.md']])
        refused(cli(args, dir, privateRun), 'the positional command requires --as large');
    expect(await readFile(path, 'utf8')).toBe(shell(visual));
  } finally {
    await rm(dir, { recursive: true, force: true });
  }
});

const crlf = (text: string) => text.replace(/\r?\n/g, '\r\n');

test('a block written into a CRLF README takes its CRLF line endings, on first insertion and on replacement', () => {
  const visual = renderSection(report());
  const large = renderSection(report(), { privateCorpus: true, date: DATE });
  const next = renderSection(later());
  const inserted = updateReadme(crlf(shell(visual)), large);
  expect(inserted).toBe(crlf(shell(`${visual}\n\n${large}`)));
  expect(updateReadme(inserted, next)).toBe(crlf(shell(`${next}\n\n${large}`)));
  expect(updateReadme(shell(visual), large)).toBe(shell(`${visual}\n\n${large}`));
});

test('--timing-sample-size below a format sampled count is refused', async () => {
  const { renderLargeSection } = (await import('./readme.mjs')) as any;
  const aggregate = written(fullReport());
  for (const format of ['docx', 'xlsx']) delete aggregate.formats[format].timing.sample_size;
  expect([aggregate.formats.docx.timing.sample, aggregate.formats.xlsx.timing.sample]).toEqual([2, 2]);
  for (const size of [0, 1])
    expect(() => renderLargeSection(aggregate, { date: DATE, timingSampleSize: size })).toThrow('--timing-sample-size');
  expect(renderLargeSection(aggregate, { date: DATE, timingSampleSize: 200 })).toBe(renderSection(fullReport(), { privateCorpus: true, date: DATE }));
});

test('a mean just below zero renders 0.0000 byte-identically on the report and the aggregate path', async () => {
  const { renderLargeSection } = (await import('./readme.mjs')) as any;
  const input = report();
  for (const comparison of input.samples[0].comparisons) comparison.penalized_ssim = -0.00001;
  const shown = renderSection(input, { privateCorpus: true, date: DATE });
  expect(renderLargeSection(written(input), { date: DATE })).toBe(shown);
  expect(shown).toContain(row('SSIM', '0.0000', '0.0000'));
});

test('--run-url refuses dot segments', async () => {
  const { withRunBadge } = (await import('./readme.mjs')) as any;
  const visual = renderSection(report());
  for (const url of [
    'https://github.com/openooxml/../actions/runs/1',
    'https://github.com/openooxml/./actions/runs/1',
    'https://github.com/../betteroffice/actions/runs/1',
    'https://github.com/./betteroffice/actions/runs/1',
    'https://github.com/openooxml/.../actions/runs/1',
    'https://github.com/openooxml/tmp/../betteroffice/actions/runs/1',
    'https://github.com/openooxml/./betteroffice/actions/runs/1',
  ])
    expect(() => withRunBadge(visual, url, DATE)).toThrow('--run-url must be');
  expect(withRunBadge(visual, RUN, DATE)).toBe(badged(visual, DATE));
});

test('withRunBadge refuses a section that already carries a run badge', async () => {
  const { withRunBadge } = (await import('./readme.mjs')) as any;
  const visual = renderSection(report());
  const large = renderSection(report(), { privateCorpus: true, date: DATE });
  expect(() => withRunBadge(withRunBadge(visual, RUN, DATE), RUN, DATE)).toThrow('Invalid generated section');
  expect(() => withRunBadge(withRunBadge(large, RUN, DATE), RUN, DATE)).toThrow('Invalid generated section');
});

test('withRunBadge finds an existing run badge by parsing its image URL, not by its spelling', async () => {
  const { withRunBadge } = (await import('./readme.mjs')) as any;
  const visual = renderSection(report());
  const canonical = 'https://betteroffice.dev/api/badge?';
  for (const spelling of ['https://BETTEROFFICE.dev/api/badge?', 'https://betteroffice.dev:443/api/badge?', 'https://BetterOffice.DEV:443/api/badge?']) {
    const section = badged(visual, DATE).replace(canonical, spelling);
    expect({ spelling, changed: section !== badged(visual, DATE) }).toEqual({ spelling, changed: true });
    expect(() => withRunBadge(section, RUN, DATE)).toThrow('Invalid generated section');
  }
});

test('the large set badge takes its date from --date, the date its block was rendered with', async () => {
  const { withRunBadge } = (await import('./readme.mjs')) as any;
  const renderLargeSection = await largeRenderer();
  const dated = '2026-01-02';
  const large = renderLargeSection(merged(fullReport()), { date: dated });
  expect(withRunBadge(large, RUN, dated)).toBe(badged(large, dated));
});

test('the public block badge takes its date from --date, and two runs write identical bytes', async () => {
  const dir = await mkdtemp(join(tmpdir(), 'readme-badge-date-'));
  try {
    const visual = renderSection(report());
    const next = renderSection(later());
    const dated = '2026-01-02';
    const expected = shell(badged(next, dated));
    const path = join(dir, 'README.md');
    await writeFile(path, shell(visual));
    await writeFile(join(dir, 'public.md'), next);
    const run = () => cli(['--section', 'public.md', '--date', dated, '--run-url', RUN, 'README.md'], dir, false);
    const first = run();
    expect({ status: first.status, readme: await readFile(path, 'utf8') }).toEqual({ status: 0, readme: expected });
    const second = run();
    expect({ status: second.status, readme: await readFile(path, 'utf8') }).toEqual({ status: 0, readme: expected });
    refused(cli(['--section', 'public.md', '--date', '2026-1-2', '--run-url', RUN, 'README.md'], dir, false), 'Invalid date');
    expect(await readFile(path, 'utf8')).toBe(expected);
  } finally {
    await rm(dir, { recursive: true, force: true });
  }
});

test('the website reads the writer output: the badged public block only, with the large set block beside it', async () => {
  const { withRunBadge } = (await import('./readme.mjs')) as any;
  const { parseFidelity } = await import('../../apps/web/lib/fidelity.ts');
  const visual = withRunBadge(renderSection(fullReport()), RUN, DATE);
  const large = withRunBadge(renderSection(fullReport(), { privateCorpus: true, date: DATE }), RUN, DATE);
  const readme = updateReadme(updateReadme(shell(renderSection(report())), visual), large);
  const formats = parseFidelity(readme);
  expect(formats.map((format) => format.format)).toEqual(['xlsx']);
  expect(formats).toEqual(parseFidelity(renderSection(fullReport())));
  const config = await readFile(join(import.meta.dir, '..', '..', 'apps', 'web', 'next.config.ts'), 'utf8');
  const pattern = config.match(/README\.match\(\/(.+?)\/\)/)![1];
  expect(readme.match(new RegExp(pattern))?.[0]).toBe(visual);
});

const messageOf = (render: () => unknown) => {
  try {
    render();
    return '';
  } catch (error) {
    return String((error as Error)?.message ?? error);
  }
};

test('the large set block takes only an aggregate that fits its schema, naming the field and never the value', async () => {
  const renderLargeSection = await largeRenderer();
  const cases: [string, (aggregate: any) => void][] = [
    ['commit', (aggregate) => { aggregate.commit = [aggregate.commit]; }],
    ['formats', (aggregate) => { aggregate.formats[STORED_PATH] = aggregate.formats.docx; }],
    ['docx published_version', (aggregate) => { aggregate.formats.docx.published_version = ['0.1.0']; }],
    ['docx published_version', (aggregate) => { aggregate.formats.docx.published_version = { toString: () => STORED_PATH }; }],
    ['xlsx libreoffice_version', (aggregate) => { aggregate.formats.xlsx.libreoffice_version = [lo]; }],
    ['docx corpus.documents', (aggregate) => { aggregate.formats.docx.corpus.documents = { length: STORED_PATH }; }],
    ['docx corpus.pages', (aggregate) => { aggregate.formats.docx.corpus.pages = STORED_PATH; }],
    ['docx fidelity.published.scored', (aggregate) => { aggregate.formats.docx.fidelity.published.scored = { length: STORED_PATH }; }],
    ['docx fidelity.published.ssim', (aggregate) => { aggregate.formats.docx.fidelity.published.ssim = Infinity; }],
    ['docx timing.channels.commit.mean_ms', (aggregate) => { aggregate.formats.docx.timing.channels.commit.mean_ms = Number.NaN; }],
    ['xlsx recalc_missing_reference', (aggregate) => { aggregate.formats.xlsx.recalc_missing_reference = { length: STORED_PATH }; }],
  ];
  for (const [field, change] of cases) {
    const aggregate = merged(fullReport());
    change(aggregate);
    refusesNaming(() => renderLargeSection(aggregate, { date: DATE }), field, STORED_PATH);
  }
  for (const label of [
    '@betteroffice/fonts@0.5.0+acme-payroll (workspace)',
    '@betteroffice/fonts@0.5.0-acme-payroll (workspace)',
    '@betteroffice/fonts@0.4.2 and @betteroffice/fonts-cjk@0.4.2+acme-payroll (local release copy)',
  ]) {
    const aggregate = merged(fontsReport());
    aggregate.formats.docx.fonts.published.sources = [label];
    refusesNaming(() => renderLargeSection(aggregate, { date: DATE }), 'docx fonts.published.sources', 'acme-payroll');
  }
  expect(renderLargeSection(merged(fontsReport()), { date: DATE })).toBe(renderSection(fontsReport(), { privateCorpus: true, date: DATE }));
});

test('a version that is not a string is refused by name on the report path', () => {
  for (const value of [['0.1.0'], { toString: () => '0.1.0' }]) {
    const input: any = report();
    input.versions.docx = value;
    refusesNaming(() => renderSection(input), 'docx version', STORED_PATH);
  }
});

test('a timing sample size below the documents sampled is refused on the report and the aggregate path alike, naming the field and no number', async () => {
  const renderLargeSection = await largeRenderer();
  const input = fullReport();
  input.timing_benchmark.docx.sample_size = 1;
  const aggregate = merged(fullReport());
  aggregate.formats.docx.timing.sample_size = 1;
  const shown = (text: string, field: string) => ({ field, named: text.includes(field), number: /\d/.test(text) });
  expect([
    shown(messageOf(() => renderSection(input, { privateCorpus: true, date: DATE })), 'timing_benchmark.docx.sample_size'),
    shown(messageOf(() => renderLargeSection(aggregate, { date: DATE })), 'docx timing.sample_size'),
  ]).toEqual([
    { field: 'timing_benchmark.docx.sample_size', named: true, number: false },
    { field: 'docx timing.sample_size', named: true, number: false },
  ]);
});

test('refuses a README whose marker shares its line with other text', async () => {
  const visual = renderSection(report());
  const large = renderSection(report(), { privateCorpus: true, date: DATE });
  const next = renderSection(later());
  const cases: [string, string][] = [
    [shell(`See ${visual}\n\n${large}`), next],
    [shell(`${visual} (generated)\n\n${large}`), next],
    [shell(`${visual}\n\nSee ${large}`), large],
    [shell(`${visual}\n\n${large} end`), large],
    [shell(visual).replace(VISUAL_BEGIN, `${VISUAL_BEGIN}<p>`), large],
  ];
  expect(cases.map(([readme, section]) => messageOf(() => updateReadme(readme, section))))
    .toEqual(cases.map(() => 'Broken generated markers'));
  const dir = await mkdtemp(join(tmpdir(), 'readme-inline-marker-'));
  try {
    const path = join(dir, 'README.md');
    await writeFile(path, cases[0][0]);
    await writeFile(join(dir, 'public.md'), next);
    refused(cli(['--section', 'public.md', 'README.md'], dir, false), 'Broken generated markers');
    expect(await readFile(path, 'utf8')).toBe(cases[0][0]);
  } finally {
    await rm(dir, { recursive: true, force: true });
  }
});

test('refuses a README with a marker twice or an end marker before its begin marker', () => {
  const visual = renderSection(report());
  const large = renderSection(report(), { privateCorpus: true, date: DATE });
  const next = renderSection(later());
  const cases: [string, string][] = [
    [shell(`${visual}\n\n${large}\n\n\`\`\`html\n${VISUAL_BEGIN}\n\`\`\``), next],
    [shell(`${visual}\n\n${large}\n\n${LARGE_END}`), large],
    [shell(`${visual}\n\n${visual}`), large],
    [shell(`${VISUAL_END}\n\nText.\n\n${VISUAL_BEGIN}\n\n${large}`), next],
    [shell(`${visual}\n\n${LARGE_END}\n\nText.\n\n${LARGE_BEGIN}`), large],
  ];
  expect(cases.map(([readme, section]) => messageOf(() => updateReadme(readme, section))))
    .toEqual(cases.map(() => 'Broken generated markers'));
});

const ok = (result: { status: number | null; stderr: string }) =>
  expect({ status: result.status, stderr: result.status ? result.stderr : '' }).toEqual({ status: 0, stderr: '' });
const outsideBytes = (bytes: Buffer, begin: string, end: string) => {
  const start = bytes.indexOf(begin);
  const stop = bytes.indexOf(end) + Buffer.byteLength(end);
  return { before: bytes.subarray(0, start).toString('hex'), after: bytes.subarray(stop).toString('hex') };
};

test('a README that is not valid UTF-8 is refused by either run kind and left byte-identical', async () => {
  const dir = await mkdtemp(join(tmpdir(), 'readme-invalid-utf8-'));
  try {
    const renderLargeSection = await largeRenderer();
    const valid = shell(`${renderSection(report())}\n\n${renderLargeSection(merged(report()), { date: DATE })}`);
    const path = join(dir, 'README.md');
    await writeFile(join(dir, 'public.md'), renderSection(later()));
    await writeFile(join(dir, 'aggregate.json'), `${JSON.stringify(merged(fullReport()), null, 2)}\n`);
    for (const original of [
      Buffer.concat([Buffer.from(valid), Buffer.from([0x46, 0x6f, 0xff, 0x6f, 0x0a])]),
      Buffer.concat([Buffer.from('# P '), Buffer.from([0xe2, 0x82]), Buffer.from(`\n\n${valid}`)]),
    ])
      for (const [args, privateRun] of [[['--section', 'public.md'], false], [['--as', 'large', '--date', DATE, 'aggregate.json'], true]] as const) {
        await writeFile(path, original);
        refused(cli([...args, 'README.md'], dir, privateRun), 'is not valid UTF-8');
        expect((await readFile(path)).toString('hex')).toBe(original.toString('hex'));
      }
  } finally {
    await rm(dir, { recursive: true, force: true });
  }
});

test('a write keeps every byte outside its own block, in a CRLF README and with trailing bytes', async () => {
  const dir = await mkdtemp(join(tmpdir(), 'readme-outside-bytes-'));
  try {
    const renderLargeSection = await largeRenderer();
    const path = join(dir, 'README.md');
    await writeFile(join(dir, 'public.md'), renderSection(later()));
    await writeFile(join(dir, 'aggregate.json'), `${JSON.stringify(merged(fullReport()), null, 2)}\n`);
    for (const eol of ['\n', '\r\n']) {
      const lines = (text: string) => Buffer.from(eol === '\n' ? text : crlf(text));
      const original = Buffer.concat([
        Buffer.from(`﻿# Projekt é 😀${eol}${eol}## Packages${eol}${eol}`),
        lines(renderSection(report())),
        Buffer.from(`${eol}${eol}Zwischen ü${eol}${eol}`),
        lines(renderLargeSection(merged(report()), { date: DATE })),
        Buffer.from(`${eol}${eol}## Evals${eol}${eol}Ende ß \t `),
      ]);
      for (const [args, privateRun, begin, end] of [
        [['--section', 'public.md'], false, VISUAL_BEGIN, VISUAL_END],
        [['--as', 'large', 'aggregate.json'], true, LARGE_BEGIN, LARGE_END],
      ] as const) {
        await writeFile(path, original);
        ok(cli([...args, '--date', DATE, '--run-url', RUN, 'README.md'], dir, privateRun));
        const bytes = await readFile(path);
        expect({ eol, begin, ...outsideBytes(bytes, begin, end) }).toEqual({ eol, begin, ...outsideBytes(original, begin, end) });
      }
    }
  } finally {
    await rm(dir, { recursive: true, force: true });
  }
});

test('a public and a private run on a CRLF README each write their block with CRLF line endings', async () => {
  const dir = await mkdtemp(join(tmpdir(), 'readme-crlf-runs-'));
  try {
    const renderLargeSection = await largeRenderer();
    const visual = renderSection(report());
    const next = renderSection(later());
    const aggregate = merged(fullReport());
    const oldLarge = renderLargeSection(merged(report()), { date: DATE });
    const path = join(dir, 'README.md');
    await writeFile(join(dir, 'public.md'), next);
    await writeFile(join(dir, 'aggregate.json'), `${JSON.stringify(aggregate, null, 2)}\n`);
    await writeFile(path, crlf(shell(`${visual}\n\n${oldLarge}`)));
    ok(cli(['--section', 'public.md', '--date', DATE, '--run-url', RUN, 'README.md'], dir, false));
    expect(await readFile(path, 'utf8')).toBe(crlf(shell(`${badged(next, DATE)}\n\n${oldLarge}`)));
    await writeFile(path, crlf(shell(visual)));
    ok(cli(['--as', 'large', '--date', DATE, '--run-url', RUN, 'aggregate.json', 'README.md'], dir, true));
    expect(await readFile(path, 'utf8')).toBe(crlf(shell(`${visual}\n\n${badged(renderLargeSection(aggregate, { date: DATE }), DATE)}`)));
  } finally {
    await rm(dir, { recursive: true, force: true });
  }
});

const outsideText = (text: string, begin: string, end: string) =>
  [text.slice(0, text.indexOf(begin)), text.slice(text.indexOf(end) + end.length)];
const EARLIER = '2026-11-01';

test('end to end: a public run changes only the public block of a README with both blocks', async () => {
  const dir = await mkdtemp(join(tmpdir(), 'readme-e2e-public-'));
  try {
    const renderLargeSection = await largeRenderer();
    const readme = shell(`${badged(renderSection(report()), EARLIER)}\n\n${badged(renderLargeSection(merged(report()), { date: EARLIER }), EARLIER)}`);
    const path = join(dir, 'README.md');
    await writeFile(path, readme);
    await writeFile(join(dir, 'public.md'), renderSection(later()));
    ok(cli(['--section', 'public.md', '--date', DATE, '--run-url', RUN, 'README.md'], dir, false));
    const updated = await readFile(path, 'utf8');
    expect(outsideText(updated, VISUAL_BEGIN, VISUAL_END)).toEqual(outsideText(readme, VISUAL_BEGIN, VISUAL_END));
    expect(span(updated, VISUAL_BEGIN, VISUAL_END)).toBe(badged(renderSection(later()), DATE));
  } finally {
    await rm(dir, { recursive: true, force: true });
  }
});

test('end to end: a private run changes only the large set block of a README with both blocks, dated by --date', async () => {
  const dir = await mkdtemp(join(tmpdir(), 'readme-e2e-private-'));
  try {
    const renderLargeSection = await largeRenderer();
    const readme = shell(`${badged(renderSection(report()), EARLIER)}\n\n${badged(renderLargeSection(merged(report()), { date: EARLIER }), EARLIER)}`);
    const aggregate = merged(fullReport());
    const dated = '2026-01-02';
    const path = join(dir, 'README.md');
    await writeFile(path, readme);
    await writeFile(join(dir, 'aggregate.json'), `${JSON.stringify(aggregate, null, 2)}\n`);
    ok(cli(['--as', 'large', '--date', dated, '--run-url', RUN, 'aggregate.json', 'README.md'], dir, true));
    const updated = await readFile(path, 'utf8');
    expect(outsideText(updated, LARGE_BEGIN, LARGE_END)).toEqual(outsideText(readme, LARGE_BEGIN, LARGE_END));
    const block = span(updated, LARGE_BEGIN, LARGE_END);
    expect(block).toBe(badged(renderLargeSection(aggregate, { date: dated }), dated));
    expect([block.includes(` on ${dated}; aggregates only;`), block.includes(`message=${dated}&amp;`)]).toEqual([true, true]);
  } finally {
    await rm(dir, { recursive: true, force: true });
  }
});

test('end to end: a first private run inserts the large set block right after the public block and changes nothing else', async () => {
  const dir = await mkdtemp(join(tmpdir(), 'readme-e2e-first-private-'));
  try {
    const renderLargeSection = await largeRenderer();
    const visual = badged(renderSection(report()), EARLIER);
    const aggregate = merged(fullReport());
    const path = join(dir, 'README.md');
    await writeFile(path, shell(visual));
    await writeFile(join(dir, 'aggregate.json'), `${JSON.stringify(aggregate, null, 2)}\n`);
    ok(cli(['--as', 'large', '--date', DATE, '--run-url', RUN, 'aggregate.json', 'README.md'], dir, true));
    expect(await readFile(path, 'utf8')).toBe(shell(`${visual}\n\n${badged(renderLargeSection(aggregate, { date: DATE }), DATE)}`));
  } finally {
    await rm(dir, { recursive: true, force: true });
  }
});

const SCHEMA_ONE_COMMIT = 'd'.repeat(40);

function schemaOneAggregate() {
  const side = (ssim: number, scored: number, total: number, exact_pages: number, paged: number, page_error: number) =>
    ({ ssim, scored, total, exact_pages, paged, page_error });
  const kept = (total: number, parsed: number, preserved: number) => ({ parsed, preserved, total });
  const timed = (sample: number, [published, commit, libreoffice]: number[]) => ({
    method: 'sequential-paired-sample-v1',
    sample,
    timed_by_every_engine: sample,
    rounds: 3,
    processors: 4,
    channels: Object.fromEntries(Object.entries({ published, commit, libreoffice }).map(([channel, mean]) =>
      [channel, { successful: sample, mean_ms: mean, median_ms: mean, round_means_ms: [mean, mean, mean], round_spread: 0.0125 }])),
    ratio_to_libreoffice: { published: { geometric_mean: 0.42, ci95: [0.4, 0.44] }, commit: { geometric_mean: 0.39, ci95: [0.37, 0.41] } },
  });
  return {
    schema_version: 1,
    commit: SCHEMA_ONE_COMMIT,
    source_sha: 'e'.repeat(40),
    formats: {
      docx: {
        samples: 5,
        published_version: '0.1.0',
        libreoffice_version: lo,
        fidelity: { published: side(0.8123, 5, 5, 4, 5, 1), commit: side(0.8345, 5, 5, 5, 5, 0), libreoffice: side(0.7012, 5, 5, 3, 5, 2) },
        corpus: { documents: 5, pages: 9 },
        render_ms: { published: 410, commit: 380, libreoffice: 900 },
        roundtrip: { published: kept(5, 5, 4), commit: kept(5, 5, 5), libreoffice: kept(5, 4, 3) },
        timing: timed(3, [400, 370, 880]),
        failures: [],
      },
      pptx: {
        samples: 4,
        published_version: '0.0.4',
        libreoffice_version: lo,
        fidelity: { published: side(0.7811, 4, 4, 3, 4, 1), commit: side(0.7922, 4, 4, 4, 4, 0), libreoffice: side(0.6633, 3, 4, 2, 3, 3) },
        corpus: { documents: 4, pages: 30 },
        render_ms: { published: 520, commit: 500, libreoffice: 1300 },
        roundtrip: { published: kept(4, 4, 4), commit: kept(4, 4, 4), libreoffice: kept(4, 4, 2) },
        timing: timed(3, [510, 490, 1250]),
        failures: [{ area: 'render', channel: 'libreoffice', stage: 'capture', count: 1 }],
      },
      xlsx: {
        samples: 6,
        published_version: '0.1.0',
        libreoffice_version: lo,
        fidelity: { published: side(0.9012, 6, 6, 6, 6, 0), commit: side(0.9134, 6, 6, 6, 6, 0), libreoffice: side(0.8456, 6, 6, 6, 6, 0) },
        corpus: { documents: 6, pages: 8 },
        recalc: {
          published: { correct: 95, total: 100, workbooks: 4, fully_matching: 3, mean_ms: 60 },
          commit: { correct: 98, total: 100, workbooks: 4, fully_matching: 4, mean_ms: 55 },
          libreoffice: { correct: 90, total: 100, workbooks: 4, fully_matching: 2, mean_ms: 400 },
        },
        roundtrip: { published: kept(6, 6, 5), commit: kept(6, 6, 6), libreoffice: kept(6, 5, 4) },
        timing: timed(2, [64, 58, 405]),
        failures: [],
      },
    },
  };
}

test('end to end: a repair from a schema-1 aggregate as written at 68a523172 changes only the large set block', async () => {
  const dir = await mkdtemp(join(tmpdir(), 'readme-e2e-schema-one-'));
  try {
    const renderLargeSection = await largeRenderer();
    const readme = shell(`${badged(renderSection(report()), EARLIER)}\n\n${badged(renderLargeSection(merged(report()), { date: EARLIER }), EARLIER)}`);
    const path = join(dir, 'README.md');
    await writeFile(path, readme);
    await writeFile(join(dir, 'aggregate.json'), `${JSON.stringify(schemaOneAggregate(), null, 2)}\n`);
    ok(cli(['--as', 'large', '--timing-sample-size', '200', '--missing-reference', '2', '--date', DATE, '--run-url', RUN, 'aggregate.json', 'README.md'], dir, false));
    const updated = await readFile(path, 'utf8');
    expect(outsideText(updated, LARGE_BEGIN, LARGE_END)).toEqual(outsideText(readme, LARGE_BEGIN, LARGE_END));
    const block = span(updated, LARGE_BEGIN, LARGE_END);
    const short = SCHEMA_ONE_COMMIT.slice(0, 8);
    const docx = block.slice(block.indexOf('### DOCX'), block.indexOf('### PPTX'));
    const xlsx = block.slice(block.indexOf('### XLSX'));
    for (const text of [
      `*Generated by the [Benchmarks action](.github/workflows/visual-fidelity.yml) from 5 DOCX, 4 PPTX and 6 XLSX documents at <a href="https://github.com/openooxml/betteroffice/commit/${SCHEMA_ONE_COMMIT}">${short}</a> on ${DATE}; aggregates only; documents are not public.*`,
      runBadge(DATE),
      'Corpus: 5 DOCX documents (9 pages), 4 PPTX presentations (30 slides) and 6 XLSX workbooks (8 printed pages), each with a Microsoft Office reference.',
      `Recalc rows: 4 workbooks with formulas (100 formula results), each recalculated from cleared caches and compared with the values Excel saved. Workbooks fully matching: BetterOffice 0.1.0 3/4, BetterOffice at ${short} 4/4, LibreOffice ${lo} 2/4. 2 further workbooks have no Excel calculation reference and are not counted.`,
      `Times: mean per document over a seeded sample of up to 200 documents per format. A document counts only when every engine timed it in every round: DOCX 3 of 3 sampled (timed per engine: BetterOffice 0.1.0 3, BetterOffice at ${short} 3, LibreOffice 3); PPTX 3 of 3 sampled (timed per engine: BetterOffice 0.0.4 3, BetterOffice at ${short} 3, LibreOffice 3); XLSX 2 of 2 sampled (timed per engine: BetterOffice 0.1.0 2, BetterOffice at ${short} 2, LibreOffice 2).`,
    ])
      expect(block.includes(text) && text).toBe(text);
    for (const text of [three('SSIM', '0.8123', '0.8345', '0.7012'), three('Exact page counts', '4/5', '5/5', '3/5'), three('Render time (avg)', '400 ms', '370 ms', '880 ms'), three('Parse success', '100.00%', '100.00%', '80.00%')])
      expect(docx.includes(text) && text).toBe(text);
    for (const text of [three('Recalc accuracy', '95.00%', '98.00%', '90.00%'), three('Recalc time (avg)', '64 ms', '58 ms', '405 ms')])
      expect(xlsx.includes(text) && text).toBe(text);
    expect(block.includes('of them with notes')).toBe(false);
    expect(span(updated, VISUAL_BEGIN, VISUAL_END)).toBe(badged(renderSection(report()), EARLIER));
  } finally {
    await rm(dir, { recursive: true, force: true });
  }
});
