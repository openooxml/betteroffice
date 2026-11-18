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

test('replaces and moves the generated block without touching other sections', () => {
  const section = renderSection(report());
  const readme =
    '# Project\n\n## Packages\n\n### DOCX\n\nPackage list.\n\n## Development\n\nCommands.\n\n## Visual fidelity\n\nOld text.\n\n## Contributing\n\nPolicy.\n';
  const updated = updateReadme(readme, section);
  expect(updated).not.toContain('Old text.');
  expect(updated).toContain('Package list.\n\n<!-- BEGIN GENERATED');
  expect(updated).toContain('<!-- END GENERATED VISUAL FIDELITY -->\n\n## Development');
  expect(updated).toContain('Commands.\n\n## Contributing');
  expect(updated.endsWith('Policy.\n')).toBe(true);
  expect(updateReadme(updated, section)).toBe(updated);
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
  const head = '# Project\n\n## Packages\n\nList.\n\n';
  expect(updateReadme(`${head}${large}\n\n## Evals\n`, visual)).toBe(`${head}${visual}\n\n${large}\n\n## Evals\n`);
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
  const today = await readFile(join(import.meta.dir, '..', '..', 'README.md'), 'utf8');
  const end = today.indexOf(VISUAL_END) + VISUAL_END.length;
  expect(updateReadme(today, first)).toBe(`${today.slice(0, end)}\n\n${first}${today.slice(end)}`);
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

test('the CLI writes a section only into the block of its own run kind', async () => {
  const dir = await mkdtemp(join(tmpdir(), 'readme-cli-'));
  try {
    const visual = renderSection(report());
    const next = renderSection(later());
    const large = renderSection(report(), { privateCorpus: true, date: DATE });
    const readme = shell(visual);
    const path = join(dir, 'README.md');
    await writeFile(path, readme);
    await writeFile(join(dir, 'public.md'), next);
    await writeFile(join(dir, 'large.md'), large);
    expect(cli(['--section', 'public.md', 'README.md'], dir, true).status).not.toBe(0);
    expect(cli(['--section', 'large.md', 'README.md'], dir, false).status).not.toBe(0);
    expect(await readFile(path, 'utf8')).toBe(readme);
    expect(cli(['--section', 'large.md', 'README.md'], dir, true).status).toBe(0);
    expect(await readFile(path, 'utf8')).toBe(shell(`${visual}\n\n${large}`));
    expect(cli(['--section', 'public.md', 'README.md'], dir, false).status).toBe(0);
    expect(await readFile(path, 'utf8')).toBe(shell(`${next}\n\n${large}`));
    await symlink('README.md', join(dir, 'link.md'));
    expect(cli(['--section', 'large.md', 'link.md'], dir, true).status).not.toBe(0);
  } finally {
    await rm(dir, { recursive: true, force: true });
  }
});

test('--as large repairs a README whose public block an older private run overwrote', async () => {
  const dir = await mkdtemp(join(tmpdir(), 'readme-repair-'));
  try {
    const visual = renderSection(report());
    const stale = `${VISUAL_BEGIN}\n## Benchmarks\n\n*Generated by the [Benchmarks action](.github/workflows/visual-fidelity.yml); per-document results are not published.*\n\nOlder private numbers.\n\n${VISUAL_END}`;
    const stored = later();
    const path = join(dir, 'README.md');
    await writeFile(path, shell(stale));
    await writeFile(join(dir, 'main.md'), shell(visual));
    await writeFile(join(dir, 'report.json'), JSON.stringify(stored));
    expect(cli(['--as', 'small', 'report.json', 'README.md'], dir, false).status).not.toBe(0);
    expect(cli(['--as', 'large', '--public', 'main.md', '--date', DATE, 'report.json', 'README.md'], dir, false).status).toBe(0);
    expect(await readFile(path, 'utf8')).toBe(shell(`${visual}\n\n${renderSection(stored, { privateCorpus: true, date: DATE })}`));
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

test('--as large renders the same block from the run artifact aggregate.json as from report.json', async () => {
  const dir = await mkdtemp(join(tmpdir(), 'readme-aggregate-'));
  try {
    const visual = renderSection(report());
    const stored = fullReport();
    await writeFile(join(dir, 'from-report.md'), shell(visual));
    await writeFile(join(dir, 'from-aggregate.md'), shell(visual));
    await writeFile(join(dir, 'report.json'), JSON.stringify(stored));
    await writeFile(join(dir, 'aggregate.json'), `${JSON.stringify(aggregateReport(stored), null, 2)}\n`);
    expect(cli(['--as', 'large', '--date', DATE, 'report.json', 'from-report.md'], dir, false).status).toBe(0);
    expect(cli(['--as', 'large', '--date', DATE, 'aggregate.json', 'from-aggregate.md'], dir, false).status).toBe(0);
    const expected = shell(`${visual}\n\n${renderSection(stored, { privateCorpus: true, date: DATE })}`);
    expect(await readFile(join(dir, 'from-report.md'), 'utf8')).toBe(expected);
    expect(await readFile(join(dir, 'from-aggregate.md'), 'utf8')).toBe(expected);
    await writeFile(join(dir, 'public.md'), shell(visual));
    expect(cli(['aggregate.json', 'public.md'], dir, false).status).not.toBe(0);
    expect(await readFile(join(dir, 'public.md'), 'utf8')).toBe(shell(visual));
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
    const bare = run(['old.json']);
    expect(bare.status).not.toBe(0);
    expect(bare.stderr).toContain('--timing-sample-size');
    const unsized = run(['--missing-reference', '1', 'old.json']);
    expect(unsized.status).not.toBe(0);
    expect(unsized.stderr).toContain('--timing-sample-size');
    const uncovered = run(['--timing-sample-size', '200', 'old.json']);
    expect(uncovered.status).not.toBe(0);
    expect(uncovered.stderr).toContain('--missing-reference');
    expect(await readFile(path, 'utf8')).toBe(shell(visual));
    expect(run(['--timing-sample-size', '200', '--missing-reference', '1', 'old.json']).status).toBe(0);
    expect(await readFile(path, 'utf8')).toBe(expected);

    for (const args of [['--timing-sample-size', '199'], ['--missing-reference', '2']]) {
      await writeFile(path, shell(visual));
      const refused = run([...args, 'new.json']);
      expect(refused.status).not.toBe(0);
      expect(refused.stderr).toContain(args[0]);
      expect(await readFile(path, 'utf8')).toBe(shell(visual));
    }
    for (const args of [['--timing-sample-size', '200'], ['--missing-reference', '1'], ['--timing-sample-size', '200', '--missing-reference', '1']]) {
      await writeFile(path, shell(visual));
      expect(run([...args, 'new.json']).status).toBe(0);
      expect(await readFile(path, 'utf8')).toBe(expected);
    }

    await writeFile(path, shell(visual));
    for (const value of ['-1', '1.5', 'abc', '', '1e2', ' 1'])
      for (const [name, other] of [['--timing-sample-size', ['--missing-reference=1']], ['--missing-reference', ['--timing-sample-size=200']]]) {
        const refused = run([`${name}=${value}`, ...other, 'old.json']);
        expect(refused.status).not.toBe(0);
        expect(refused.stderr).toContain(name);
      }
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
    await writeFile(join(dir, 'large.md'), large);

    await writeFile(path, shell(visual));
    expect(cli(['--section', 'large.md', 'README.md'], dir, true).status).toBe(0);
    const plain = await readFile(path, 'utf8');
    expect(plain).toBe(shell(`${visual}\n\n${large}`));
    for (const absent of ['actions/runs', 'benchmark%20run']) expect(plain.includes(absent) && absent).toBe(false);

    await writeFile(path, shell(visual));
    expect(cli(['--section', 'large.md', '--run-url', RUN, 'README.md'], dir, true).status).toBe(0);
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

    await writeFile(join(dir, 'report.json'), JSON.stringify(later()));
    await writeFile(path, shell(visual));
    expect(cli(['report.json', '--run-url', RUN, 'README.md'], dir, false).status).toBe(0);
    const fromReport = await readFile(path, 'utf8');
    const reportDay = fromReport.match(/benchmark run (\d{4}-\d{2}-\d{2})/)?.[1] ?? '';
    expect([before, utcDay()]).toContain(reportDay);
    expect(fromReport).toBe(shell(badged(next, reportDay)));

    await writeFile(path, shell(visual));
    const refused = cli(['--section', 'large.md', '--run-url', 'https://example.com/actions/runs/1', 'README.md'], dir, true);
    expect(refused.status).not.toBe(0);
    expect(refused.stderr).toContain('--run-url');
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
    await writeFile(join(dir, 'report.json'), JSON.stringify(stored));
    await writeFile(join(dir, 'aggregate.json'), `${JSON.stringify(aggregateReport(stored), null, 2)}\n`);
    const path = join(dir, 'README.md');
    for (const input of ['report.json', 'aggregate.json']) {
      await writeFile(path, shell(visual));
      expect(cli(['--as', 'large', '--date', DATE, '--run-url', RUN, input, 'README.md'], dir, false).status).toBe(0);
      expect(await readFile(path, 'utf8')).toBe(expected);
    }
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
    ]) {
      const refused = cli(['--as', 'large', '--date', DATE, `--run-url=${url}`, 'aggregate.json', 'README.md'], dir, false);
      expect(refused.status).not.toBe(0);
      expect(refused.stderr).toContain('--run-url');
    }
    expect(await readFile(path, 'utf8')).toBe(shell(visual));
  } finally {
    await rm(dir, { recursive: true, force: true });
  }
});
