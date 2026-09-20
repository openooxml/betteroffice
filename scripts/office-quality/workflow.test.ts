import { readFile } from 'node:fs/promises';
import { expect, test } from 'bun:test';

const workflow = Bun.YAML.parse(
  await readFile(
    new URL('../../.github/workflows/visual-fidelity.yml', import.meta.url),
    'utf8'
  )
) as any;
const { prepare, measure, publish } = workflow.jobs;

test('one frozen plan feeds independent format jobs at the same source revision', () => {
  expect(workflow.on.workflow_dispatch.inputs.collection).toBeUndefined();
  expect(prepare.outputs.formats).toBe('${{ steps.plan.outputs.formats }}');
  expect(measure.needs).toBe('prepare');
  expect(measure.strategy).toEqual({
    'fail-fast': false,
    matrix: { format: '${{ fromJSON(needs.prepare.outputs.formats) }}' },
  });
  expect(measure.steps[0].with.ref).toBe('${{ needs.prepare.outputs.source-sha }}');
  const evaluate = measure.steps.find(
    (step: any) => step.run === 'node scripts/office-quality/run.mjs'
  );
  expect(evaluate.env.QUALITY_PLAN).toBe('${{ runner.temp }}/fidelity-plan/plan.json');
  expect(evaluate.env.QUALITY_FORMAT).toBe('${{ matrix.format }}');
});

test('only the reconciler publishes reports and renders after every format succeeds', () => {
  expect(publish.needs).toEqual(['prepare', 'measure', 'docx-benchmark']);
  expect(publish.if).not.toContain('always()');
  expect(
    measure.steps.some((step: any) => step.run?.includes('publish-renders.mjs'))
  ).toBe(false);
  const merge = publish.steps.findIndex((step: any) => step.run?.includes('merge.mjs'));
  const renders = publish.steps.findIndex((step: any) =>
    step.run?.includes('publish-renders.mjs')
  );
  const install = publish.steps.findIndex((step: any) =>
    step.run?.includes('bun install --frozen-lockfile')
  );
  const readme = publish.steps.findIndex(
    (step: any) => step.name === 'Replace the generated README section'
  );
  expect(merge).toBeGreaterThan(-1);
  expect(install).toBeGreaterThan(merge);
  expect(install).toBeLessThan(renders);
  expect(publish.steps[install].if).toBe(publish.steps[renders].if);
  expect(publish.steps[install].run).toContain('bunx --no-install wrangler --version');
  expect(renders).toBeGreaterThan(merge);
  expect(readme).toBeGreaterThan(renders);
  expect(publish.steps[renders]['continue-on-error']).toBeUndefined();
  expect(publish.steps[1].env.SOURCE_SHA).toBe('${{ needs.prepare.outputs.source-sha }}');
});

test('artifact directories remain stable when a run selects only one format', () => {
  const downloads = publish.steps.filter(
    (step: any) => step.uses?.startsWith('actions/download-artifact@')
  );
  expect(downloads.filter((step: any) => step.with.pattern).map((step: any) => step.with.pattern))
    .toEqual(['docx-benchmark-report-*']);
  for (const format of ['docx', 'pptx', 'xlsx']) {
    for (const kind of ['report', 'renders']) {
      const name = `visual-fidelity-${kind}-${format}`;
      const matches = downloads.filter((step: any) => step.with.name === name);
      expect(matches).toHaveLength(1);
      expect(matches[0].with.path).toBe(`\${{ runner.temp }}/fidelity-parts/${name}`);
      expect(matches[0].if).toBe(
        (kind === 'renders' ? 'inputs.publish_renders && ' : '') +
          `contains(fromJSON(needs.prepare.outputs.formats), '${format}')`
      );
    }
  }
});

test('native builds and paired DOCX shards are isolated from browser measurements', () => {
  const build = workflow.jobs['native-build'];
  const benchmark = workflow.jobs['docx-benchmark'];
  expect(build.needs).toBe('prepare');
  expect(build.strategy.matrix.channel).toEqual(['published', 'commit']);
  expect(build.steps[1].with.ref).toContain('needs.prepare.outputs.docx-published-source');
  expect(benchmark.needs).toEqual(['prepare', 'native-build']);
  expect(benchmark.strategy.matrix.shard).toBe('${{ fromJSON(needs.prepare.outputs.docx-shards) }}');
  expect(benchmark['runs-on']).toBe('ubuntu-24.04');
  expect(benchmark.steps.some((step: any) => step.run?.includes('playwright'))).toBe(false);
  expect(benchmark.steps.some((step: any) => step.run?.includes('chmod +x'))).toBe(true);
  expect(publish.if).toContain("needs.docx-benchmark.result == 'success'");
  const merge = publish.steps.find((step: any) => step.run?.includes('merge.mjs'));
  expect(merge.env.QUALITY_REQUIRE_DOCX_BENCHMARK).toBe('true');
});
