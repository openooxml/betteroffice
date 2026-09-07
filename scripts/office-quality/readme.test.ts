import { expect, test } from 'bun:test';
import { renderSection, updateReadme } from './readme.mjs';

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
  expect(section).toContain('| PPTX | [0.0.4]');
  expect(section).toContain('| XLSX | [0.1.0]');
  expect(section).toContain('— = not measured');
});

test('replaces and moves the generated block without touching other sections', () => {
  const section = renderSection(report());
  const readme =
    '# Project\n\n## Visual fidelity\n\nOld text.\n\n## Development\n\nCommands.\n\n## Contributing\n\nPolicy.\n';
  const updated = updateReadme(readme, section);
  expect(updated).not.toContain('Old text.');
  expect(updated).toContain('Commands.\n\n<!-- BEGIN GENERATED');
  expect(updated).toContain('<!-- END GENERATED VISUAL FIDELITY -->\n\n## Contributing');
  expect(updated.endsWith('Policy.\n')).toBe(true);
  expect(updateReadme(updated, section)).toBe(updated);
});

test('refuses failed or mismatched render records', () => {
  const input = report();
  input.samples[0].comparisons[0].actual.sha256 = 'wrong';
  expect(() => renderSection(input)).toThrow('mismatched');
});
