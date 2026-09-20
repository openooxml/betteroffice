import { expect, test } from 'bun:test';
import { renderSection, updateReadme } from './readme.mjs';

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
