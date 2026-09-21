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
  expect(publish.needs).toEqual(['prepare', 'measure', 'docx-benchmark', 'xlsx-benchmark', 'pptx-benchmark', 'xlsx-fidelity', 'roundtrip']);
  expect(publish.if).not.toContain('always()');
  expect(
    measure.steps.some((step: any) => step.run?.includes('publish-renders.mjs'))
  ).toBe(false);
  const merge = publish.steps.findIndex((step: any) => step.run?.includes('merge.mjs'));
  const renders = publish.steps.findIndex((step: any) =>
    step.run?.includes('publish-renders.mjs')
  );
  const runtime = publish.steps.findIndex((step: any) =>
    step.uses?.startsWith('oven-sh/setup-bun@')
  );
  const readme = publish.steps.findIndex(
    (step: any) => step.name === 'Replace the generated README section'
  );
  expect(merge).toBeGreaterThan(-1);
  expect(runtime).toBeGreaterThan(merge);
  expect(runtime).toBeLessThan(renders);
  expect(publish.steps[runtime].if).toBe(publish.steps[renders].if);
  expect(publish.steps[renders].run).toBe('bun scripts/office-quality/publish-renders.mjs');
  expect(publish.steps[renders].env.CLOUDFLARE_API_TOKEN).toBe('${{ secrets.CLOUDFLARE_API_TOKEN }}');
  expect(publish.steps.some((step: any) => step.run?.includes('wrangler'))).toBe(false);
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
    .toEqual(['docx-benchmark-report-*', 'xlsx-benchmark-report-*', 'xlsx-fidelity-report-*', 'roundtrip-report-*']);
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


test('XLSX calculation and PPTX LibreOffice run independently and gate publication', () => {
  const xlsx = workflow.jobs['xlsx-benchmark'];
  const pptx = workflow.jobs['pptx-benchmark'];
  expect(xlsx.needs).toEqual(['prepare', 'xlsx-native-build']);
  expect(xlsx.strategy.matrix.shard).toBe('${{ fromJSON(needs.prepare.outputs.xlsx-shards) }}');
  expect(pptx.needs).toEqual(['prepare', 'pptx-native-build']);
  expect(workflow.jobs['pptx-native-build'].steps[1].with.ref).toContain('needs.prepare.outputs.pptx-published-source');
  expect(pptx.steps.find((step: any) => step.run?.includes('pptx_benchmark.py')).env.QUALITY_NATIVE).toBe('${{ runner.temp }}/native');
  expect(publish.if).toContain("needs.xlsx-benchmark.result == 'success'");
  expect(publish.if).toContain("needs.pptx-benchmark.result == 'success'");
  const merge = publish.steps.find((step: any) => step.run?.includes('merge.mjs'));
  expect(merge.env.QUALITY_REQUIRE_PPTX_BENCHMARK).toBe('true');
  expect(merge.env.QUALITY_REQUIRE_XLSX_BENCHMARK).toBe('true');
  const builds = workflow.jobs['xlsx-native-build'];
  expect(builds.strategy.matrix.channel).toEqual(['published','commit']);
  expect(builds.steps[1].with.ref).toContain('needs.prepare.outputs.xlsx-published-source');
});


test('XLSX fidelity covers every planned workbook independently of calculation workers', () => {
  const fidelity = workflow.jobs['xlsx-fidelity'];
  expect(fidelity.needs).toBe('prepare');
  expect(fidelity.strategy.matrix.shard).toBe('${{ fromJSON(needs.prepare.outputs.xlsx-fidelity-shards) }}');
  expect(fidelity.if).toBe("contains(fromJSON(needs.prepare.outputs.formats), 'xlsx')");
  expect(publish.if).toContain("needs.xlsx-fidelity.result == 'success'");
  expect(publish.steps.find((step: any) => step.run?.includes('merge.mjs')).env.QUALITY_REQUIRE_XLSX_FIDELITY).toBe('true');
});


test('native parse/edit/save probes run independently and gate the final report', () => {
  const build = workflow.jobs['roundtrip-build'];
  const probe = workflow.jobs.roundtrip;
  expect(build.needs).toBe('prepare');
  expect(build.strategy.matrix.include).toBe('${{ fromJSON(needs.prepare.outputs.roundtrip-builds) }}');
  expect(build.steps[1].with.ref).toBe('${{ matrix.source }}');
  expect(probe.needs).toEqual(['prepare', 'roundtrip-build']);
  expect(probe.strategy.matrix.include).toBe('${{ fromJSON(needs.prepare.outputs.roundtrip-shards) }}');
  expect(probe.env.LO_PYTHON).toBe('/opt/libreoffice26.2/program/python');
  expect(probe.steps.some((step: any) => step.run?.includes('install-libreoffice.sh'))).toBe(true);
  expect(publish.if).toContain("needs.roundtrip.result == 'success'");
  expect(publish.steps.find((step: any) => step.run?.includes('merge.mjs')).env.QUALITY_REQUIRE_ROUNDTRIP).toBe('true');
});
