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
  expect(publish.needs).toEqual(['prepare', 'measure']);
  expect(publish.if).not.toContain('always()');
  expect(
    measure.steps.some((step: any) => step.run?.includes('publish-renders.mjs'))
  ).toBe(false);
  const merge = publish.steps.findIndex((step: any) => step.run?.includes('merge.mjs'));
  const renders = publish.steps.findIndex((step: any) =>
    step.run?.includes('publish-renders.mjs')
  );
  const readme = publish.steps.findIndex(
    (step: any) => step.name === 'Replace the generated README section'
  );
  expect(merge).toBeGreaterThan(-1);
  expect(renders).toBeGreaterThan(merge);
  expect(readme).toBeGreaterThan(renders);
  expect(publish.steps[renders]['continue-on-error']).toBeUndefined();
  expect(publish.steps[1].env.SOURCE_SHA).toBe('${{ needs.prepare.outputs.source-sha }}');
});
