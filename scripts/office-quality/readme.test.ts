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
  expect(shown).toContain('per-document results are not published');
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
  expect(section.indexOf('Corpus: ')).toBeLessThan(section.indexOf('<picture>'));
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
  expect(shown).toContain('XLSX Recalc time (avg) 100 ms vs 410 ms');
  expect(shown).toContain('Times: mean per document over a seeded sample of up to 200 documents per format (timed by every engine: XLSX 2), run one after another on one machine per format');
  expect(shown).not.toContain('w1');
  const timing = aggregateReport(input).formats.xlsx.timing;
  expect(timing.channels.commit).toEqual({ successful: 2, mean_ms: 100, median_ms: 100, round_means_ms: [100, 100, 100], round_spread: 0 });
  expect(timing.ratio_to_libreoffice.commit).toEqual({ geometric_mean: 0.5, ci95: [0.4, 0.6] });
  expect(JSON.stringify(timing)).not.toContain('w1');
});
