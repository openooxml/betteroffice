import { spawnSync } from 'node:child_process';
import { mkdtemp, readFile, rm } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { expect, test } from 'bun:test';
import { JOB_TIMEOUT_MINUTES } from './shards.mjs';
import { CHANNELS, ENGINE_TIMEOUT, HELPER_TIMEOUT, PARALLELISM, SHARD_SIZE, ROUNDTRIP_MATRICES, MAX_MATRIX_JOBS, roundtripMatrices, roundtripShards } from './roundtrip.mjs';

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
  expect(prepare.outputs['measure-shards']).toBe('${{ steps.plan.outputs.measure-shards }}');
  expect(prepare.steps.find((step: any) => step.id === 'plan').run)
    .toContain('measure-shards=${JSON.stringify(measureMatrix(plan))}');
  expect(measure.strategy).toEqual({
    'fail-fast': false,
    'max-parallel': "${{ fromJSON(needs.prepare.outputs.parallel)['measure'] }}",
    matrix: { include: '${{ fromJSON(needs.prepare.outputs.measure-shards) }}' },
  });
  const checkout = measure.steps.find((step: any) => step.uses?.startsWith('actions/checkout@'));
  expect(checkout.with.ref).toBe('${{ needs.prepare.outputs.source-sha }}');
  const evaluate = measure.steps.find(
    (step: any) => step.run === 'node scripts/office-quality/private-step.mjs measure -- node scripts/office-quality/run.mjs'
  );
  expect(evaluate.env.QUALITY_PLAN).toBe('${{ runner.temp }}/fidelity-plan/plan.json');
  expect(evaluate.env.QUALITY_FORMAT).toBe('${{ matrix.format }}');
  expect(evaluate.env.QUALITY_SHARD).toBe('${{ matrix.shard }}');
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
  expect(publish.needs).toEqual(['prepare', 'measure', 'docx-benchmark', 'xlsx-benchmark', 'pptx-benchmark', 'xlsx-fidelity', 'roundtrip-a', 'roundtrip-b', 'timing']);
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

test('every measure shard has its own artifacts, and the reconciler collects them by pattern', () => {
  const downloads = publish.steps.filter(
    (step: any) => step.uses?.startsWith('actions/download-artifact@')
  );
  expect(downloads.filter((step: any) => step.with.pattern).map((step: any) => [step.with.pattern, step.with.path, step.if])).toEqual([
    ['visual-fidelity-report-*', '${{ runner.temp }}/fidelity-parts', '${{ !inputs.private_corpus }}'],
    ['visual-fidelity-renders-*', '${{ runner.temp }}/fidelity-parts', "inputs.publish_renders && inputs.corpus_origin == ''"],
    ['docx-benchmark-report-*', '${{ runner.temp }}/fidelity-parts', expect.any(String)],
    ['xlsx-benchmark-report-*', '${{ runner.temp }}/fidelity-parts', expect.any(String)],
    ['pptx-benchmark-report-*', '${{ runner.temp }}/fidelity-parts', expect.any(String)],
    ['xlsx-fidelity-report-*', '${{ runner.temp }}/fidelity-parts', expect.any(String)],
    ['roundtrip-report-*', '${{ runner.temp }}/fidelity-parts', expect.any(String)],
  ]);
  expect(downloads.some((step: any) => String(step.with.name ?? '').startsWith('visual-fidelity-re'))).toBe(false);
  for (const step of downloads.filter((step: any) => step.with.pattern)) expect(step.with['merge-multiple']).toBeUndefined();
  const uploads = measure.steps.filter((step: any) => step.uses?.startsWith('actions/upload-artifact@'));
  expect(uploads.map((step: any) => step.with.name)).toEqual([
    'visual-fidelity-report-${{ matrix.format }}-${{ matrix.shard }}',
    'visual-fidelity-renders-${{ matrix.format }}-${{ matrix.shard }}',
  ]);
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
  expect(pptx.strategy).toEqual({ 'fail-fast': false, 'max-parallel': "${{ fromJSON(needs.prepare.outputs.parallel)['pptx-benchmark'] }}",
    matrix: { shard: '${{ fromJSON(needs.prepare.outputs.pptx-shards) }}' } });
  expect(prepare.outputs['pptx-shards']).toBe('${{ steps.plan.outputs.pptx-shards }}');
  expect(pptx.steps.find((step: any) => String(step.run ?? '').includes('pptx-benchmark.mjs')).env.QUALITY_SHARD).toBe('${{ matrix.shard }}');
  expect(pptx.steps.filter((step: any) => step.uses?.startsWith('actions/upload-artifact@')).map((step: any) => step.with.name))
    .toEqual(['pptx-benchmark-report-${{ matrix.shard }}', 'pptx-benchmark-diagnostics-${{ matrix.shard }}']);
  expect(publish.steps.find((step: any) => step.with?.pattern === 'pptx-benchmark-report-*')).toBeDefined();
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
    expect(probe.strategy).toEqual({'fail-fast': false, 'max-parallel': `\${{ fromJSON(needs.prepare.outputs.parallel)['roundtrip-${half}'] }}`,
      matrix: {include: `\${{ fromJSON(needs.prepare.outputs.roundtrip-shards-${half}) }}`}});
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
    .filter((step: any) => step !== guard && step.name !== 'Keep report-only runs off the public render viewer' &&
      String(step.if ?? '').includes('publish_renders'));
  expect(gated.length).toBe(4);
  for (const step of gated) expect(String(step.if).startsWith(RENDERS)).toBe(true);
  const publishing = publish.steps.find((step: any) => String(step.run ?? '').includes('publish-renders.mjs'));
  expect(publishing.if).toBe(`${RENDERS} && !inputs.private_corpus && github.repository == 'openooxml/betteroffice'`);
  const reconcile = publish.steps.find((step: any) => step.name === 'Reconcile the format reports');
  expect(reconcile.env.QUALITY_REQUIRE_RENDERS).toBe(`\${{ ${RENDERS} }}`);
});

const steps = Object.entries<any>(workflow.jobs).flatMap(([job, definition]) =>
  (definition.steps ?? []).map((step: any) => ({ job, step })));
const CORPUS_FREE = new Set(['build-native.mjs', 'build-format.mjs', 'build-roundtrip.mjs', 'native-fonts.ts',
  'install-libreoffice.sh', 'asset-cache.mjs', 'publish-renders.mjs', 'private-store.mjs', 'shards.mjs']);
const BINARY_ARTIFACTS = /^(docx-native-\$\{\{ matrix\.channel \}\}|docx-native-fonts|xlsx-native-\$\{\{ matrix\.channel \}\}|pptx-native-\$\{\{ matrix\.channel \}\}|roundtrip-native-\$\{\{ matrix\.format \}\}-\$\{\{ matrix\.channel \}\})$/;
const README = "${{ (!inputs.private_corpus && !inputs.report_only && inputs.corpus_origin == '' && inputs.samples == '' && inputs.collection == 'office-quality') || (inputs.private_corpus && inputs.publish_readme && !inputs.report_only) }}";
const PRIVATE_OFF = /(^|&& |^\$\{\{ )!inputs\.private_corpus( &&|$| \}\}$)/;

test('private corpus runs are an explicit input that refuses sample lists and render publishing', () => {
  expect(workflow.on.workflow_dispatch.inputs.private_corpus).toEqual({
    description: 'Private corpus run; logs, artifacts and caches carry no per-document data',
    default: false, required: false, type: 'boolean',
  });
  expect(workflow.env.QUALITY_PRIVATE).toBe("${{ inputs.private_corpus && 'true' || 'false' }}");
  const guard = prepare.steps.find((step: any) => step.name === 'Keep private runs aggregate-only');
  expect(prepare.steps.indexOf(guard)).toBeLessThan(prepare.steps.findIndex((step: any) => step.id === 'plan'));
  expect(guard.if).toBe('inputs.private_corpus');
  expect(guard.env).toEqual({
    SAMPLES_GIVEN: "${{ inputs.samples != '' }}",
    PUBLISH_RENDERS: '${{ inputs.publish_renders }}',
    ORIGIN_GIVEN: "${{ inputs.corpus_origin != '' }}",
  });
  expect(guard.run).toContain('[ "$ORIGIN_GIVEN" != false ]');
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
  expect(labels.length).toBe(17);
  expect(new Set(labels).size).toBe(labels.length);
});

test('private runs upload no per-document artifact and touch no asset cache, render bucket or README', () => {
  for (const { job, step } of steps) {
    const uses = String(step.uses ?? '');
    if (uses.startsWith('actions/upload-artifact@')) {
      if (BINARY_ARTIFACTS.test(step.with.name)) continue;
      if (step.if === 'inputs.private_corpus') {
        expect({ job, name: step.with.name, path: step.with.path }).toEqual({ job, name: 'visual-fidelity', path: AGGREGATE_ONLY });
        continue;
      }
      expect({ job, name: step.with.name, if: step.if }).toEqual({ job, name: step.with.name, if: expect.stringMatching(PRIVATE_OFF) });
    }
    if (uses.startsWith('actions/cache') && step.with.path.includes('fidelity-assets') || step.id === 'asset-cache' || String(step.run ?? '').includes('publish-renders.mjs') ||
        ['oven-sh/setup-bun@0c5077e51419868618aeaa5fe8019c62421857d6'].includes(uses) && job === 'publish' ||
        false)
      expect({ job, name: step.name ?? uses, if: step.if }).toEqual({ job, name: step.name ?? uses, if: expect.stringMatching(PRIVATE_OFF) });
    if (['Replace the generated README section', 'Detect a changed report'].includes(step.name)) expect(step.if).toBe(README);
    const cache = step.env?.QUALITY_ASSET_CACHE;
    if (cache !== undefined) expect(cache).toStartWith('${{ !inputs.private_corpus && ');
  }
  const mint = publish.steps.find((step: any) => step.name === 'Mint the bot token');
  const commit = publish.steps.find((step: any) => step.name === 'Commit only the generated README');
  expect([mint.if, commit.if]).toEqual(["steps.changed.outputs.value == 'true'", "steps.changed.outputs.value == 'true'"]);
});

const ENVIRONMENT = { name: "${{ inputs.private_corpus && 'private-corpus' || '' }}", deployment: false };
const READ = {
  PRIVATE_R2_ENDPOINT: '${{ secrets.PRIVATE_R2_ENDPOINT }}',
  PRIVATE_R2_READ_KEY_ID: '${{ secrets.PRIVATE_R2_READ_KEY_ID }}',
  PRIVATE_R2_READ_SECRET: '${{ secrets.PRIVATE_R2_READ_SECRET }}',
};
const WRITE = {
  PRIVATE_R2_ENDPOINT: '${{ secrets.PRIVATE_R2_ENDPOINT }}',
  PRIVATE_R2_WRITE_KEY_ID: '${{ secrets.PRIVATE_R2_WRITE_KEY_ID }}',
  PRIVATE_R2_WRITE_SECRET: '${{ secrets.PRIVATE_R2_WRITE_SECRET }}',
};
const PRIVATE_ON = /^(always\(\) && )?inputs\.private_corpus$/;
const privateKeys = (env: any) => Object.fromEntries(Object.entries(env ?? {}).filter(([key]) => key.startsWith('PRIVATE_R2_')));
const stores = steps.flatMap(({ job, step }) => [...String(step.run ?? '').matchAll(/(?:private-store\.mjs"?|"\$store") (put|get) (?:--optional )?('[^']+'|(?:[a-z0-9-]|\$\{\{ [a-z.]+ \}\})+)(?: "([^"]+)")?/g)]
  .map((match) => ({ job, step, mode: match[1], name: match[2].replace(/'/g, ''), path: match[3] })));

test('private runs take their secrets from a main-only environment without deployment records', () => {
  const corpusJobs = Object.entries<any>(workflow.jobs).filter(([, job]) =>
    (job.steps ?? []).some((step: any) => /private-(step|store)\.mjs/.test(String(step.run ?? ''))));
  expect(corpusJobs.map(([name]) => name).sort()).toEqual(
    ['docx-benchmark', 'measure', 'prepare', 'publish', 'pptx-benchmark', 'roundtrip-a', 'roundtrip-b', 'timing', 'xlsx-benchmark', 'xlsx-fidelity'].sort());
  for (const [, job] of corpusJobs) expect(job.environment).toEqual(ENVIRONMENT);
  for (const [name, job] of Object.entries<any>(workflow.jobs))
    if (!corpusJobs.some(([corpus]) => corpus === name)) expect({ name, environment: job.environment }).toEqual({ name, environment: undefined });
  expect(prepare.if).toBe("github.ref == 'refs/heads/main'");
  expect(JSON.stringify(workflow.env)).not.toContain('PRIVATE_R2');
  for (const job of Object.values<any>(workflow.jobs)) expect(JSON.stringify(job.env ?? {})).not.toContain('PRIVATE_R2');
});

test('read keys reach only corpus readers and bucket reads, write keys only bucket writes', () => {
  let readers = 0;
  for (const { job, step } of steps) {
    const keys = privateKeys(step.env);
    const run = String(step.run ?? '');
    const label = { job, name: step.name ?? step.uses };
    if (!Object.keys(keys).length) {
      expect({ ...label, store: run.includes('private-store.mjs') }).toEqual({ ...label, store: false });
      continue;
    }
    if (run.includes('private-store.mjs')) {
      expect({ ...label, if: step.if }).toEqual({ ...label, if: expect.stringMatching(PRIVATE_ON) });
      const modes = new Set(stores.filter((store) => store.step === step).map((store) => store.mode));
      expect([...modes]).toHaveLength(1);
      expect(keys).toEqual(modes.has('put') ? WRITE : READ);
    } else {
      expect({ ...label, run }).toEqual({ ...label, run: expect.stringMatching(/private-step\.mjs (plan|measure|[a-z-]+-stage) -- node /) });
      expect(keys).toEqual(READ);
      readers += 1;
    }
  }
  expect(readers).toBe(9);
  const checks = measure.steps.find((step: any) => step.name === 'Check benchmark scripts');
  expect(checks.env).toEqual({ QUALITY_PRIVATE: 'false' });
});

test('every per-document artifact has a private bundle of the same name and directory', () => {
  const normalize = (path: string) => path.replace('${{ runner.temp }}', '$RUNNER_TEMP').replace('$parts', '$RUNNER_TEMP/fidelity-parts');
  for (const { job, step } of steps) {
    const uses = String(step.uses ?? '');
    if (uses.startsWith('actions/upload-artifact@') && PRIVATE_OFF.test(step.if ?? '') && !step.with.name.startsWith('visual-fidelity-renders-'))
      expect({ job, name: step.with.name, put: stores.some((store) => store.job === job && store.mode === 'put' && store.name === step.with.name) })
        .toEqual({ job, name: step.with.name, put: true });
    if (uses.startsWith('actions/download-artifact@') && PRIVATE_OFF.test(step.if ?? '') && !String(step.with.name ?? '').startsWith('visual-fidelity-renders-')) {
      const name = step.with.name ?? step.with.pattern;
      const matches = stores.filter((store) => store.job === job && store.mode === 'get' && store.name === name);
      expect({ job, name, paths: matches.map((store) => normalize(store.path!)) }).toEqual({ job, name, paths: [normalize(step.with.path)] });
    }
  }
  const puts = stores.filter((store) => store.mode === 'put').map((store) => `${store.job}/${store.name}`);
  expect(new Set(puts).size).toBe(puts.length);
  for (const job of ['prepare', 'measure', 'docx-benchmark', 'xlsx-benchmark', 'pptx-benchmark', 'xlsx-fidelity', 'roundtrip-a', 'timing', 'publish'])
    expect(stores.some((store) => store.job === job && store.mode === 'put' && store.name.startsWith('logs-') &&
      store.step.if === 'always() && inputs.private_corpus')).toBe(true);
});

test('private runs check out only the main commit they were dispatched from', () => {
  const checkout = prepare.steps.find((step: any) => step.uses?.startsWith('actions/checkout@'));
  expect(checkout.with.ref).toBe('${{ inputs.private_corpus && github.sha || inputs.branch }}');
  const pin = prepare.steps.find((step: any) => step.id === 'source');
  expect(pin.env).toEqual({ TARGET_BRANCH: '${{ inputs.branch }}', PRIVATE_CORPUS: '${{ inputs.private_corpus }}' });
  const publishCheckout = publish.steps.find((step: any) => step.uses?.startsWith('actions/checkout@'));
  expect(publishCheckout.with.ref).toBe('${{ inputs.private_corpus && needs.prepare.outputs.source-sha || inputs.branch }}');
  for (const [name, job] of Object.entries<any>(workflow.jobs)) {
    if (job.environment === undefined || name === 'prepare' || name === 'publish') continue;
    for (const step of job.steps ?? [])
      if (step.uses?.startsWith('actions/checkout@'))
        expect({ name, ref: step.with?.ref }).toEqual({ name, ref: '${{ needs.prepare.outputs.source-sha }}' });
  }
});

test('the pin step refuses private runs on another branch or commit', async () => {
  const pin = prepare.steps.find((step: any) => step.id === 'source');
  const root = await mkdtemp(join(tmpdir(), 'pin-step-'));
  const git = (...args: string[]) => spawnSync('git', ['-c', 'user.name=t', '-c', 'user.email=t@example.com', ...args], { cwd: root, encoding: 'utf8' });
  try {
    git('init', '-q');
    git('commit', '-q', '--allow-empty', '-m', 'one');
    const first = git('rev-parse', 'HEAD').stdout.trim();
    git('commit', '-q', '--allow-empty', '-m', 'two');
    const head = git('rev-parse', 'HEAD').stdout.trim();
    const run = (branch: string, sha: string) => {
      const output = join(root, 'output');
      const result = spawnSync('bash', ['--noprofile', '--norc', '-eo', 'pipefail', '-c', pin.run], {
        cwd: root,
        encoding: 'utf8',
        env: { PATH: process.env.PATH, TARGET_BRANCH: branch, PRIVATE_CORPUS: 'true', GITHUB_SHA: sha, GITHUB_OUTPUT: output },
      });
      const written = spawnSync('cat', [output], { encoding: 'utf8' }).stdout;
      spawnSync('rm', ['-f', output]);
      return { status: result.status, refused: result.stdout.includes('::error::'), written };
    };
    expect(run('main', head)).toEqual({ status: 0, refused: false, written: `sha=${head}\n` });
    expect(run('', head)).toEqual({ status: 0, refused: false, written: `sha=${head}\n` });
    expect(run('feature', head)).toEqual({ status: 1, refused: true, written: '' });
    expect(run('main-copy', head)).toEqual({ status: 1, refused: true, written: '' });
    expect(run('main', first)).toEqual({ status: 1, refused: true, written: '' });
  } finally {
    await rm(root, { recursive: true, force: true });
  }
});

const AGGREGATE_ONLY = '${{ runner.temp }}/fidelity-report/section.md\n${{ runner.temp }}/fidelity-report/aggregate.json\n';
test('a private run commits the README only when publish_readme asks for it', () => {
  expect(workflow.on.workflow_dispatch.inputs.publish_readme).toEqual({
    description: "Private corpus run: commit this run's aggregate-only section to the README",
    default: false, required: false, type: 'boolean',
  });
  const guard = prepare.steps.find((step: any) => step.name === 'Keep README publishing to private runs');
  expect(guard.if).toBe('inputs.publish_readme && (!inputs.private_corpus || inputs.report_only)');
  expect(guard.run).toContain('exit 1');
  expect(prepare.steps.indexOf(guard)).toBeLessThan(prepare.steps.findIndex((step: any) => step.id === 'plan'));
  const merge = publish.steps.find((step: any) => String(step.run ?? '').includes('merge.mjs'));
  expect(merge.env?.QUALITY_PRIVATE ?? workflow.env.QUALITY_PRIVATE).toBe("${{ inputs.private_corpus && 'true' || 'false' }}");
});

test('only a default showcase run commits the README; report-only runs publish nothing', () => {
  expect(workflow.on.workflow_dispatch.inputs.report_only).toEqual({
    description: "Keep the results in this run's artifacts; never commit the README or publish renders",
    default: false, required: false, type: 'boolean',
  });
  const guard = prepare.steps.find((step: any) => step.name === 'Keep report-only runs off the public render viewer');
  expect(guard.if).toBe('inputs.report_only && inputs.publish_renders');
  expect(guard.run).toContain('exit 1');
  expect(prepare.steps.indexOf(guard)).toBeLessThan(prepare.steps.findIndex((step: any) => step.id === 'plan'));
  const readme = publish.steps.filter((step: any) => ['Replace the generated README section', 'Detect a changed report'].includes(step.name));
  expect(readme.map((step: any) => step.if)).toEqual([README, README]);
  for (const step of publish.steps.filter((step: any) => /README|bot token/.test(step.name ?? '') && !readme.includes(step)))
    expect(step.if).toBe("steps.changed.outputs.value == 'true'");
});

test('the reconciled artifact always carries the aggregate, and the private bundle keeps it too', () => {
  const uploads = publish.steps.filter((step: any) => step.uses?.startsWith('actions/upload-artifact@') && step.with.name === 'visual-fidelity');
  expect(uploads.map((step: any) => [step.if, step.with.path])).toEqual([
    ['${{ !inputs.private_corpus }}', '${{ runner.temp }}/fidelity-report/report.json\n' + AGGREGATE_ONLY + '${{ runner.temp }}/fidelity-report/failures.json\n'],
    ['inputs.private_corpus', AGGREGATE_ONLY],
  ]);
  const keep = publish.steps.find((step: any) => step.name === 'Keep the reconciled report in the private bucket');
  expect(keep.run).toContain('private-store.mjs" put visual-fidelity report.json section.md aggregate.json failures.json');
});

test('private runs time a seeded sample per format, sequentially on one runner each, and gate the report on it', () => {
  const timing = workflow.jobs.timing;
  expect(timing.needs).toEqual(['prepare', 'native-build', 'xlsx-native-build', 'pptx-native-build']);
  expect(timing.if).toContain('inputs.private_corpus');
  expect(timing.if).toContain("needs.prepare.outputs.timing-formats != '[]'");
  expect(timing.if).toStartWith('!cancelled()');
  expect(prepare.outputs['timing-formats']).toBe('${{ steps.plan.outputs.timing-formats }}');
  expect(prepare.steps.find((step: any) => step.id === 'plan').run).toContain('timing-formats=${JSON.stringify(timingFormats(plan))}');
  expect(timing.strategy).toEqual({ 'fail-fast': false, 'max-parallel': "${{ fromJSON(needs.prepare.outputs.parallel)['timing'] }}",
    matrix: { format: '${{ fromJSON(needs.prepare.outputs.timing-formats) }}' } });
  const downloads = timing.steps.filter((step: any) => step.uses?.startsWith('actions/download-artifact@')).map((step: any) => step.with.name);
  expect(downloads).toEqual(['${{ matrix.format }}-native-published', '${{ matrix.format }}-native-commit']);
  const measured = timing.steps.find((step: any) => String(step.run ?? '').includes('timing.py'));
  expect(measured.run).toBe('node scripts/office-quality/private-step.mjs timing -- python scripts/office-quality/timing.py');
  expect(timing.steps.some((step: any) => step.uses?.startsWith('actions/upload-artifact@'))).toBe(false);
  expect(publish.if).toContain("(!inputs.private_corpus || needs.prepare.outputs.timing-formats == '[]' || needs.timing.result == 'success')");
  const merge = publish.steps.find((step: any) => step.run?.includes('merge.mjs'));
  expect(merge.env.QUALITY_REQUIRE_TIMING).toBe('${{ inputs.private_corpus }}');
  const fetch = publish.steps.find((step: any) => step.name === 'Fetch the format reports from the private bucket');
  expect(fetch.run).toContain(`node "$store" get 'timing-report-*' "$parts"`);
  expect(stores.some((store) => store.job === 'timing' && store.mode === 'put' && store.name === 'timing-report-${{ matrix.format }}')).toBe(true);
});

test('the roundtrip matrices cover every file of a corpus-sized plan for all three channels', () => {
  const ids = (format: string, count: number) => Array.from({ length: count }, (_, index) => ({ id: `${format}-${index}`, format }));
  for (const sizes of [{ docx: 1776, pptx: 656, xlsx: 3297 }, { docx: 1823, pptx: 674, xlsx: 3526 }, { docx: 2667, pptx: 2667, xlsx: 2666 }]) {
    const samples = Object.entries(sizes).flatMap(([format, count]) => ids(format, count));
    const plan = { formats: Object.keys(sizes), samples };
    const shards = roundtripShards(plan);
    expect(shards.every((shard: any) => shard.ids.length >= 1 && shard.ids.length <= SHARD_SIZE)).toBe(true);
    expect(shards.flatMap((shard: any) => shard.ids).sort()).toEqual(samples.map((sample) => sample.id).sort());
    const matrices = roundtripMatrices(plan);
    expect(matrices).toHaveLength(ROUNDTRIP_MATRICES);
    for (const matrix of matrices) expect(matrix.length).toBeLessThanOrEqual(MAX_MATRIX_JOBS);
    const jobs = matrices.flat().map(({ format, shard }: any) => `${format}/${shard}`);
    expect(new Set(jobs).size).toBe(shards.length);
    expect(jobs.sort()).toEqual(shards.map(({ format, shard }: any) => `${format}/${shard}`).sort());
  }
  expect(CHANNELS).toEqual(['published', 'commit', 'libreoffice']);
});

test('planning prints every benchmark job shard count and planned minutes against the timeout', () => {
  const plan = prepare.steps.find((step: any) => step.id === 'plan').run;
  expect(plan).toContain('benchmarkShardPlan(plan)');
  expect(plan).toContain('GITHUB_STEP_SUMMARY');
  expect(plan).toContain('pptx-shards=${JSON.stringify(pptxShards(plan).map((_, index) => index))}');
  for (const name of ['docx-benchmark', 'xlsx-benchmark', 'pptx-benchmark', 'xlsx-fidelity'])
    expect({ name, timeout: workflow.jobs[name]['timeout-minutes'] }).toEqual({ name, timeout: JOB_TIMEOUT_MINUTES });
});

test('every benchmark job prints its wall time next to its planned work, even when a step fails', () => {
  for (const name of ['docx-benchmark', 'xlsx-benchmark', 'pptx-benchmark', 'xlsx-fidelity']) {
    const steps = workflow.jobs[name].steps;
    expect(steps[1]).toEqual({ name: 'Record the job start', run: 'echo "QUALITY_JOB_STARTED=$(date +%s)" >> "$GITHUB_ENV"' });
    expect(steps.at(-1)).toEqual({ name: 'Compare the wall time with the planned work', if: 'always()',
      env: { QUALITY_PLAN: '${{ runner.temp }}/fidelity-plan/plan.json' }, run: `node scripts/office-quality/shards.mjs wall ${name} "\${{ matrix.shard }}"` });
  }
});

test('a private run bounds its concurrent benchmark jobs by max_parallel; public runs are unbounded', () => {
  expect(workflow.on.workflow_dispatch.inputs.max_parallel).toEqual({
    description: "Private corpus run: most benchmark jobs held at once, leaving the rest of the plan's concurrent jobs to pull request CI",
    default: 12, required: false, type: 'number',
  });
  const plan = prepare.steps.find((step: any) => step.id === 'plan');
  expect(plan.env.MAX_PARALLEL).toBe('${{ inputs.max_parallel }}');
  expect(plan.run).toContain("parallelPlan(plan, { privateRun: process.env.QUALITY_PRIVATE === 'true', total: Number(process.env.MAX_PARALLEL) })");
  expect(plan.run).toContain('parallelSummary(plan, parallel)');
  expect(prepare.outputs.parallel).toBe('${{ steps.plan.outputs.parallel }}');
  for (const name of ['measure', 'docx-benchmark', 'xlsx-benchmark', 'pptx-benchmark', 'xlsx-fidelity', 'roundtrip-a', 'roundtrip-b', 'timing'])
    expect({ name, cap: workflow.jobs[name].strategy['max-parallel'] }).toEqual({ name, cap: `\${{ fromJSON(needs.prepare.outputs.parallel)['${name}'] }}` });
});

test('the reconciler fails on a channel that scored nothing, and the summary leads with scored counts even when it fails', () => {
  const reconcile = publish.steps.find((step: any) => step.name === 'Reconcile the format reports');
  expect(reconcile.env.QUALITY_REQUIRE_COVERAGE).toBe('true');
  expect(reconcile.env.QUALITY_COVERAGE).toBe('${{ runner.temp }}/coverage.md');
  expect(reconcile.run).toBe('node scripts/office-quality/private-step.mjs merge -- node scripts/office-quality/merge.mjs');
  const summary = publish.steps[publish.steps.indexOf(reconcile) + 1];
  expect(summary.name).toBe('Summarise scored documents and the generated section');
  expect(summary.if).toBe('always()');
  expect(summary.run.indexOf('coverage.md')).toBeLessThan(summary.run.indexOf('section.md'));
});

test('planning states which font package each channel renders with', () => {
  const plan = prepare.steps.find((step: any) => step.id === 'plan').run;
  expect(plan).toContain("published channel: @betteroffice/fonts@${version}, ${plan.fonts_basis?.[format] === 'release-commit' ? 'the fonts release from the same release commit as' : 'the latest fonts release published before'} @betteroffice/${format}@${plan.versions[format]}");
  expect(plan).toContain('commit channel: the font packages of this checkout');
  expect(plan).toContain('### Font sources');
});
