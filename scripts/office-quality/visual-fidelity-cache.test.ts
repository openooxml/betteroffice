import { readFile } from 'node:fs/promises';
import { expect, test } from 'bun:test';

const workflow = await readFile(
  new URL('../../.github/workflows/visual-fidelity.yml', import.meta.url),
  'utf8'
);

test('fidelity asset cache restores latest prefix and saves only changed digests', () => {
  expect(workflow).toContain(
    'actions/cache/restore@0057852bfaa89a56745cba8c7296529d2fc39830'
  );
  expect(workflow).toContain(
    'actions/cache/save@0057852bfaa89a56745cba8c7296529d2fc39830'
  );
  expect(workflow).toContain('path: ${{ runner.temp }}/fidelity-assets');
  expect(workflow).toContain('fidelity-assets-v2-');
  expect(workflow).toContain(
    'key: fidelity-assets-v2-${{ runner.os }}-${{ matrix.format }}-'
  );
  expect(workflow).toContain(
    'fidelity-assets-v2-${{ runner.os }}-${{ matrix.format }}-${{ steps.asset-cache.outputs.digest }}'
  );
  expect(workflow).toContain('cache-matched-key');
  expect(workflow).toContain("steps.asset-cache.outputs.digest != ''");
  expect(workflow).not.toContain('github.run_id');
  expect(workflow).not.toContain('hashFiles');
});

test('asset cache digest comes from cached blob names without extra manifest fetch', () => {
  expect(workflow).toContain(
    'node scripts/office-quality/asset-cache.mjs "${RUNNER_TEMP}/fidelity-assets"'
  );
  expect(workflow).toContain('echo "digest=${digest}"');
  expect(workflow).not.toContain('corpus.betteroffice.dev/collections');
  expect(workflow).not.toContain('steps.corpus.outputs.sha');
  expect(workflow).toContain('QUALITY_ASSET_CACHE: ${{ runner.temp }}/fidelity-assets');
});

test('digest step and save run even on measurement failure', () => {
  const digest = workflow.slice(workflow.indexOf('Compute fidelity asset cache digest'));
  expect(digest).toContain('id: asset-cache');
  expect(digest).toContain('if: always()');
  const save = workflow.slice(workflow.indexOf('Save fidelity asset cache'));
  expect(save).toContain('if: always()');
});

test('branch-head guard, least privilege, and score artifacts are unchanged', () => {
  expect(workflow).toContain("if: github.ref == 'refs/heads/main'");
  expect(workflow).toContain('contents: read');
  const artifacts = workflow.slice(workflow.indexOf('actions/upload-artifact@'));
  expect(artifacts).toContain('.source/office-quality/ci/report.json');
  expect(artifacts).toContain('${{ runner.temp }}/fidelity-report/section.md');
  expect(artifacts).toContain('name: visual-fidelity-report-${{ matrix.format }}');
  expect(artifacts).toContain('name: visual-fidelity-renders-${{ matrix.format }}');
});
