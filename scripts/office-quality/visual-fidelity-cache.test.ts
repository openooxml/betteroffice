import { readFile } from 'node:fs/promises';
import { expect, test } from 'bun:test';

const workflow = await readFile(
  new URL('../../.github/workflows/visual-fidelity.yml', import.meta.url),
  'utf8'
);

const SHOWCASE = "!inputs.private_corpus && inputs.corpus_origin == '' && inputs.collection == 'office-quality'";
const parsed = Bun.YAML.parse(workflow) as any;
const steps = Object.entries<any>(parsed.jobs).flatMap(([job, definition]) => (definition.steps ?? []).map((step: any) => ({ job, step })));

test('only showcase runs from the default corpus touch the Actions cache for corpus assets', () => {
  const cacheSteps = steps.filter(({ step }) =>
    (String(step.uses ?? '').startsWith('actions/cache') && String(step.with?.path ?? '').includes('fidelity-assets')) || step.id === 'asset-cache');
  expect(cacheSteps.map(({ job }) => job)).toEqual(['measure', 'measure', 'measure', 'docx-benchmark']);
  for (const { job, step } of cacheSteps)
    expect({ job, name: step.name, if: String(step.if) }).toEqual({ job, name: step.name, if: expect.stringContaining(SHOWCASE) });
  const users = steps.filter(({ step }) => step.env?.QUALITY_ASSET_CACHE !== undefined);
  expect(users.map(({ job }) => job)).toEqual(['measure', 'docx-benchmark', 'xlsx-benchmark', 'pptx-benchmark', 'xlsx-fidelity', 'roundtrip-a', 'roundtrip-b']);
  for (const { step } of users) expect(step.env.QUALITY_ASSET_CACHE).toBe(`\${{ ${SHOWCASE} && format('{0}/fidelity-assets', runner.temp) || '' }}`);
  const benchmarkRestore = cacheSteps[3].step;
  expect(benchmarkRestore.with.key).toBe('fidelity-assets-v3-${{ runner.os }}-docx-');
});

test('fidelity asset cache restores latest prefix and saves only changed digests', () => {
  expect(workflow).toContain(
    'actions/cache/restore@55cc8345863c7cc4c66a329aec7e433d2d1c52a9'
  );
  expect(workflow).toContain(
    'actions/cache/save@55cc8345863c7cc4c66a329aec7e433d2d1c52a9'
  );
  expect(workflow).toContain('path: ${{ runner.temp }}/fidelity-assets');
  expect(workflow).not.toContain('fidelity-assets-v2-');
  expect(workflow).toContain(
    'key: fidelity-assets-v3-${{ runner.os }}-${{ matrix.format }}-${{ matrix.shard }}-'
  );
  expect(workflow).toContain(
    'fidelity-assets-v3-${{ runner.os }}-${{ matrix.format }}-${{ matrix.shard }}-${{ steps.asset-cache.outputs.digest }}'
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
  expect(workflow).toContain(`QUALITY_ASSET_CACHE: \${{ ${SHOWCASE} && format('{0}/fidelity-assets', runner.temp) || '' }}`);
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
  expect(artifacts).toContain('name: visual-fidelity-report-${{ matrix.format }}-${{ matrix.shard }}');
  expect(artifacts).toContain('name: visual-fidelity-renders-${{ matrix.format }}-${{ matrix.shard }}');
});
