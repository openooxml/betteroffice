import { readFile } from 'node:fs/promises';
import { expect, test } from 'bun:test';
import { CHANNELS, ENGINE_TIMEOUT, HELPER_TIMEOUT, PARALLELISM, SHARD_SIZE } from './roundtrip.mjs';

const workflow = Bun.YAML.parse(
  await readFile(
    new URL('../../.github/workflows/visual-fidelity.yml', import.meta.url),
    'utf8'
  )
) as any;
const { prepare, measure, publish } = workflow.jobs;

test('one frozen plan feeds independent format jobs at the same source revision', () => {
  expect(workflow.on.workflow_dispatch.inputs.collection.default).toBe('office-quality');
  expect(workflow.on.workflow_dispatch.inputs.corpus_origin.default).toBe('');
  expect(workflow.env.QUALITY_CORPUS_ORIGIN).toBe('${{ inputs.corpus_origin }}');
  expect(prepare.steps.find((step: any) => step.id === 'plan').env.QUALITY_COLLECTION).toBe('${{ inputs.collection }}');
  expect(prepare.outputs.formats).toBe('${{ steps.plan.outputs.formats }}');
  expect(measure.needs).toBe('prepare');
  expect(measure.strategy).toEqual({
    'fail-fast': false,
    matrix: { format: '${{ fromJSON(needs.prepare.outputs.formats) }}' },
  });
  const checkout = measure.steps.find((step: any) => step.uses?.startsWith('actions/checkout@'));
  expect(checkout.with.ref).toBe('${{ needs.prepare.outputs.source-sha }}');
  const evaluate = measure.steps.find(
    (step: any) => step.run === 'node scripts/office-quality/private-step.mjs measure -- node scripts/office-quality/run.mjs'
  );
  expect(evaluate.env.QUALITY_PLAN).toBe('${{ runner.temp }}/fidelity-plan/plan.json');
  expect(evaluate.env.QUALITY_FORMAT).toBe('${{ matrix.format }}');
});

test('every job feeding the report runs on the runner image the plan recorded', () => {
  expect(prepare['runs-on']).toBe('ubuntu-24.04');
  expect(prepare.outputs['runner-image']).toBe('${{ steps.runner.outputs.image }}');
  for (const name of publish.needs.filter((name: string) => name !== 'prepare')) {
    const [guard] = workflow.jobs[name].steps;
    expect(workflow.jobs[name]['runs-on']).toBe('ubuntu-24.04');
    expect(guard.env.PLANNED_IMAGE).toBe('${{ needs.prepare.outputs.runner-image }}');
    expect(guard.run).toContain('"$ImageVersion" != "$PLANNED_IMAGE"');
  }
});

test('only the reconciler publishes reports and renders after every format succeeds', () => {
  expect(publish.needs).toEqual(['prepare', 'measure', 'docx-benchmark', 'xlsx-benchmark', 'pptx-benchmark', 'xlsx-fidelity', 'roundtrip-a', 'roundtrip-b']);
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
        (kind === 'renders' ? "inputs.publish_renders && inputs.corpus_origin == '' && " : '') +
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
  expect(build.needs).toBe('prepare');
  expect(build.strategy.matrix.include).toBe('${{ fromJSON(needs.prepare.outputs.roundtrip-builds) }}');
  expect(build.steps[1].with.ref).toBe('${{ matrix.source }}');
  expect(workflow.jobs.roundtrip).toBeUndefined();
  const [first, second] = ['a', 'b'].map(half => {
    const probe = workflow.jobs[`roundtrip-${half}`];
    expect(prepare.outputs[`roundtrip-shards-${half}`]).toBe(`\${{ steps.plan.outputs.roundtrip-shards-${half} }}`);
    expect(probe.needs).toEqual(['prepare', 'roundtrip-build']);
    expect(probe.strategy).toEqual({'fail-fast': false, matrix: {include: `\${{ fromJSON(needs.prepare.outputs.roundtrip-shards-${half}) }}`}});
    expect(probe.env.LO_PYTHON).toBe('/opt/libreoffice26.2/program/python');
    expect(probe.steps.some((step: any) => step.run?.includes('install-libreoffice.sh'))).toBe(true);
    return probe;
  });
  expect(prepare.outputs['roundtrip-shards']).toBeUndefined();
  expect(prepare.steps.find((step: any) => step.id === 'plan').run).toContain('roundtripMatrices(plan)');
  expect(first.if).toBeUndefined();
  expect(second.if).toBe("needs.prepare.outputs.roundtrip-shards-b != '[]'");
  expect(second.steps).toEqual(first.steps);
  expect(second.env).toEqual(first.env);
  expect(second['timeout-minutes']).toBe(first['timeout-minutes']);
  expect(publish.if).toContain("needs.roundtrip-a.result == 'success'");
  expect(publish.if).toContain("(needs.prepare.outputs.roundtrip-shards-b == '[]' || needs.roundtrip-b.result == 'success')");
  expect(publish.steps.find((step: any) => step.run?.includes('merge.mjs')).env.QUALITY_REQUIRE_ROUNDTRIP).toBe('true');
});

test('a shard of the slowest possible samples still finishes inside the job timeout', () => {
  const probe = workflow.jobs['roundtrip-a'];
  const worstSeconds =
    Math.ceil(SHARD_SIZE / PARALLELISM) *
    (CHANNELS.length * (ENGINE_TIMEOUT + HELPER_TIMEOUT) + HELPER_TIMEOUT);
  expect(worstSeconds).toBe(44 * 60);
  expect(probe['timeout-minutes']).toBeGreaterThan(worstSeconds / 60);
  expect(worstSeconds).toBeLessThan(probe['timeout-minutes'] * 60 * 0.75);
});

test('a custom corpus origin never publishes renders, and a dispatch asking for both fails before planning', () => {
  const RENDERS = "inputs.publish_renders && inputs.corpus_origin == ''";
  const guard = prepare.steps.find((step: any) => step.name === 'Keep custom-origin runs off the public render viewer');
  expect(guard.if).toBe("inputs.corpus_origin != '' && inputs.publish_renders");
  expect(guard.run).toContain('exit 1');
  expect(prepare.steps.indexOf(guard)).toBeLessThan(prepare.steps.findIndex((step: any) => step.id === 'plan'));
  const gated = Object.values<any>(workflow.jobs).flatMap((job) => job.steps ?? [])
    .filter((step: any) => step !== guard && String(step.if ?? '').includes('publish_renders'));
  expect(gated.length).toBe(6);
  for (const step of gated) expect(String(step.if).startsWith(RENDERS)).toBe(true);
  const publishing = publish.steps.find((step: any) => String(step.run ?? '').includes('publish-renders.mjs'));
  expect(publishing.if).toBe(`${RENDERS} && !inputs.private_corpus && github.repository == 'openooxml/betteroffice'`);
  const reconcile = publish.steps.find((step: any) => step.name === 'Reconcile the format reports');
  expect(reconcile.env.QUALITY_REQUIRE_RENDERS).toBe(`\${{ ${RENDERS} }}`);
});

const steps = Object.entries<any>(workflow.jobs).flatMap(([job, definition]) =>
  (definition.steps ?? []).map((step: any) => ({ job, step })));
const CORPUS_FREE = new Set(['build-native.mjs', 'build-format.mjs', 'build-roundtrip.mjs', 'native-fonts.ts',
  'install-libreoffice.sh', 'asset-cache.mjs', 'publish-renders.mjs']);
const BINARY_ARTIFACTS = /^(docx-native-\$\{\{ matrix\.channel \}\}|docx-native-fonts|xlsx-native-\$\{\{ matrix\.channel \}\}|pptx-native-\$\{\{ matrix\.channel \}\}|roundtrip-native-\$\{\{ matrix\.format \}\}-\$\{\{ matrix\.channel \}\})$/;
const PRIVATE_OFF = /(^|&& )!inputs\.private_corpus( &&|$)|^\$\{\{ !inputs\.private_corpus \}\}$/;

test('private corpus runs are an explicit input that refuses sample lists and render publishing', () => {
  expect(workflow.on.workflow_dispatch.inputs.private_corpus).toEqual({
    description: 'Private corpus run; logs, artifacts and caches carry no per-document data',
    default: false, required: false, type: 'boolean',
  });
  expect(workflow.env.QUALITY_PRIVATE).toBe("${{ inputs.private_corpus && 'true' || 'false' }}");
  const guard = prepare.steps.find((step: any) => step.name === 'Keep private runs aggregate-only');
  expect(prepare.steps.indexOf(guard)).toBeLessThan(prepare.steps.findIndex((step: any) => step.id === 'plan'));
  expect(guard.if).toBe('inputs.private_corpus');
  expect(guard.env).toEqual({ SAMPLES_GIVEN: "${{ inputs.samples != '' }}", PUBLISH_RENDERS: '${{ inputs.publish_renders }}' });
  expect(guard.run).toContain('exit 1');
});

test('every step that reads the corpus runs through the private-step wrapper', () => {
  const labels: string[] = [];
  for (const { job, step } of steps) {
    for (const line of String(step.run ?? '').split('\n')) {
      const invoked = [...line.matchAll(/\b(?:node|python3?|bun|bash) scripts\/office-quality\/([\w.-]+\.(?:mjs|py|ts|sh))/g)].map((match) => match[1]);
      if (!invoked.length || invoked.every((name) => CORPUS_FREE.has(name))) continue;
      const wrapped = line.trim().match(/^node scripts\/office-quality\/private-step\.mjs ([a-z0-9-]+) -- (node|python) scripts\/office-quality\/[\w.-]+$/);
      expect({ job, line: line.trim(), wrapped: Boolean(wrapped) }).toEqual({ job, line: line.trim(), wrapped: true });
      labels.push(`${job}/${wrapped![1]}`);
    }
  }
  expect(labels.length).toBe(15);
  expect(new Set(labels).size).toBe(labels.length);
});

test('private runs upload no per-document artifact and touch no asset cache, render bucket or README', () => {
  for (const { job, step } of steps) {
    const uses = String(step.uses ?? '');
    if (uses.startsWith('actions/upload-artifact@')) {
      if (BINARY_ARTIFACTS.test(step.with.name)) continue;
      if (step.if === 'inputs.private_corpus') {
        expect({ job, name: step.with.name, path: step.with.path }).toEqual({ job, name: 'visual-fidelity', path: '${{ runner.temp }}/fidelity-report/section.md' });
        continue;
      }
      expect({ job, name: step.with.name, if: step.if }).toEqual({ job, name: step.with.name, if: expect.stringMatching(PRIVATE_OFF) });
    }
    if (uses.startsWith('actions/cache') && step.with.path.includes('fidelity-assets') || step.id === 'asset-cache' || String(step.run ?? '').includes('publish-renders.mjs') ||
        ['oven-sh/setup-bun@0c5077e51419868618aeaa5fe8019c62421857d6'].includes(uses) && job === 'publish' ||
        ['Replace the generated README section', 'Detect a changed report'].includes(step.name))
      expect({ job, name: step.name ?? uses, if: step.if }).toEqual({ job, name: step.name ?? uses, if: expect.stringMatching(PRIVATE_OFF) });
    const cache = step.env?.QUALITY_ASSET_CACHE;
    if (cache !== undefined) expect(cache).toBe("${{ !inputs.private_corpus && format('{0}/fidelity-assets', runner.temp) || '' }}");
  }
  const mint = publish.steps.find((step: any) => step.name === 'Mint the bot token');
  const commit = publish.steps.find((step: any) => step.name === 'Commit only the generated README');
  expect([mint.if, commit.if]).toEqual(["steps.changed.outputs.value == 'true'", "steps.changed.outputs.value == 'true'"]);
});
