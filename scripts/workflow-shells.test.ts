import { describe, expect, test } from 'bun:test';
import { mkdtempSync, readFileSync, readdirSync, rmSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';

const workflows = new URL('../.github/workflows/', import.meta.url);
const actions = new URL('../.github/actions/', import.meta.url);
const toolchainFile = new URL('../rust-toolchain.toml', import.meta.url);
const nativeJob = 'native-dist.yml bindings';
const upstream = /^dtolnay\/rust-toolchain@/i;

// These steps also run on windows-latest, where a bash default would replace pwsh.
const WINDOWS_STEPS = [
  'native-dist.yml bindings: Install dependencies',
  'native-dist.yml bindings: Build binding',
  'native-dist.yml bindings: Build the JS entries',
  'native-dist.yml bindings: Test binding',
];

// GitHub runs `shell: bash` as `bash --noprofile --norc -eo pipefail {0}`; with no shell at all it runs `bash -e {0}`.
const pipefail = (shell: unknown) => shell === 'bash' || (typeof shell === 'string' && /\bpipefail\b/.test(shell));
const label = (step: any) => step.name ?? String(step.run).split('\n')[0];

const jobs = readdirSync(workflows)
  .filter((file) => /\.ya?ml$/.test(file))
  .flatMap((file) => {
    const workflow = Bun.YAML.parse(readFileSync(new URL(file, workflows), 'utf8')) as any;
    return Object.entries<any>(workflow.jobs).map(([id, job]) => ({
      name: `${file} ${id}`,
      job,
      shell: workflow.defaults?.run?.shell,
    }));
  });

const compositeActions = readdirSync(actions, { withFileTypes: true })
  .filter((entry) => entry.isDirectory())
  .map((entry) => {
    const action = Bun.YAML.parse(readFileSync(new URL(`${entry.name}/action.yml`, actions), 'utf8')) as any;
    return { name: entry.name, steps: action.runs?.steps ?? [] };
  });

const runSteps = jobs.flatMap(({ name, job, shell }) =>
  (job.steps ?? [])
    .filter((step: any) => step.run !== undefined)
    .map((step: any) => ({
      name: `${name}: ${label(step)}`,
      job,
      shell: step.shell ?? job.defaults?.run?.shell ?? shell,
    }))
);

const rustCommand = /(?:^|[\s(;&|=])(?:cargo|rustc|rustup|wasm-pack|napi[ \t]+build)(?=\s|[);&|]|$)/;
const bunScript = /\bbun[ \t]+run[ \t]+(?:(?:--(?:cwd|filter)(?:=|[ \t]+)(?:"[^"\n]+"|'[^'\n]+'|\S+)|--[\w-]+)[ \t]+)*["']?([\w:./-]+)/g;

function needsRust(job: string, step: { uses?: string; run?: string }): boolean {
  if (step.uses === './.github/actions/wasm-toolchain') return true;
  const run = (step.run ?? '').replace(/^[ \t]*#.*$/gm, '').replace(/\\\r?\n/g, ' ');
  if (rustCommand.test(run)) return true;
  if (job === 'ci.yml python' && /\bpython -m pip install\b[^\n]* -e\b/.test(run)) return true;
  const scripts = run.replace(/\$\{\{.*?\}\}/g, 'expression').matchAll(bunScript);
  return [...scripts].some(([, script]) => /wasm|native/.test(script));
}

describe('workflow shells', () => {
  test('every run step runs under bash with pipefail, except the native binding steps that also run on Windows', () => {
    expect(runSteps.length).toBeGreaterThan(0);
    const without = runSteps.filter(({ shell }) => !pipefail(shell));
    expect(without.map(({ name }) => name).sort()).toEqual([...WINDOWS_STEPS].sort());
    for (const { name, job } of without) {
      const systems = (job.strategy?.matrix?.platform ?? []).map((platform: any) => platform.os);
      expect(systems, name).toContain('windows-latest');
      expect(systems.some((os: string) => !os.startsWith('windows-')), name).toBe(true);
    }
  });

  test('every composite action run step declares shell: bash', () => {
    const steps = compositeActions.flatMap(({ name, steps }) =>
      steps
        .filter((step: any) => step.run !== undefined)
        .map((step: any) => ({ name: `${name}: ${label(step)}`, shell: step.shell }))
    );
    expect(steps.length).toBeGreaterThan(0);
    for (const { name, shell } of steps) expect(shell, name).toBe('bash');
  });
});

test('the Rust pin readers require exact channels and handle missing files', () => {
  const source = readFileSync(toolchainFile, 'utf8');
  const pin = Bun.TOML.parse(source) as { toolchain: { channel: string } };
  expect(pin.toolchain.channel).toMatch(/^\d+\.\d+\.\d+$/);
  const composite = compositeActions.find(({ name }) => name === 'rust-toolchain');
  const native = jobs.find(({ name }) => name === nativeJob);
  const readers = [
    {
      name: 'composite pin',
      run: composite?.steps.find((step: any) => step.id === 'pin')?.run,
      missingChannel: null,
    },
    {
      name: 'native-dist rust-pin',
      run: native?.job.steps.find((step: any) => step.id === 'rust-pin')?.run,
      missingChannel: 'stable',
    },
  ];
  const valid = '[toolchain]\nprofile = "minimal"\nchannel = "1.2.3"\ncomponents = ["clippy"]\n';
  const cases = [
    { name: 'LF', contents: valid, channel: '1.2.3' },
    { name: 'CRLF', contents: valid.replace(/\n/g, '\r\n'), channel: '1.2.3' },
    { name: 'stable', contents: valid.replace('1.2.3', 'stable'), channel: null },
    { name: 'minor only', contents: valid.replace('1.2.3', '1.2'), channel: null },
    { name: 'missing file', contents: undefined, channel: null },
    { name: 'repository pin', contents: source, channel: pin.toolchain.channel },
  ];
  for (const { name, contents, channel } of cases) {
    const cwd = mkdtempSync(join(tmpdir(), 'workflow-rust-pin-'));
    try {
      const file = join(cwd, 'rust-toolchain.toml');
      const output = join(cwd, 'output');
      if (contents !== undefined) writeFileSync(file, contents);
      for (const reader of readers) {
        const context = `${reader.name}: ${name}`;
        expect(typeof reader.run, context).toBe('string');
        writeFileSync(output, '');
        const result = Bun.spawnSync(['bash', '--noprofile', '--norc', '-eo', 'pipefail', '-c', reader.run], {
          cwd,
          env: { ...process.env, RUST_TOOLCHAIN_FILE: file, GITHUB_OUTPUT: output },
          stdout: 'pipe',
          stderr: 'pipe',
        });
        const expected = contents === undefined ? reader.missingChannel : channel;
        const statusContext = `${context}\n${result.stdout.toString()}${result.stderr.toString()}`;
        if (expected === null) {
          expect(result.exitCode, statusContext).toBeGreaterThan(0);
        } else {
          expect(result.exitCode, statusContext).toBe(0);
        }
        expect(readFileSync(output, 'utf8'), context).toBe(expected === null ? '' : `toolchain=${expected}\n`);
      }
    } finally {
      rmSync(cwd, { recursive: true, force: true });
    }
  }
});

test('CI installs Rust from the repository toolchain pin', () => {
  const installers: { name: string; toolchain: unknown }[] = [];
  for (const { name, steps } of [
    ...jobs.map(({ name, job }) => ({ name, steps: [job, ...(job.steps ?? [])] })),
    ...compositeActions,
  ]) {
    for (const step of steps) {
      expect(step.uses ?? '', name).not.toMatch(/^dtolnay\/rust-toolchain@stable$/i);
      expect(step.with?.toolchain, name).not.toBe('stable');
      if (upstream.test(step.uses ?? '')) {
        expect(step.uses, name).toMatch(/@[0-9a-f]{40}$/i);
        installers.push({ name, toolchain: step.with?.toolchain });
      }
    }
  }
  expect(installers).toEqual([
    { name: nativeJob, toolchain: '${{ steps.rust-pin.outputs.toolchain }}' },
    { name: 'rust-toolchain', toolchain: '${{ steps.pin.outputs.toolchain }}' },
  ]);

  const rustJobs = jobs.flatMap(({ name, job }) => {
    const steps = job.steps ?? [];
    const firstUse = steps.findIndex((step: any) => needsRust(name, step));
    return firstUse < 0 ? [] : [{ name, steps, firstUse }];
  });
  expect(rustJobs.map(({ name }) => name)).toEqual(expect.arrayContaining([
    'ci.yml native-hooks',
    'ci.yml python',
    nativeJob,
  ]));

  for (const { name, steps, firstUse } of rustJobs) {
    const install = steps.findIndex((step: any) => name === nativeJob
      ? upstream.test(step.uses ?? '')
      : step.uses === './.github/actions/rust-toolchain');
    expect(install, name).toBeGreaterThanOrEqual(0);
    expect(install, `${name}: ${label(steps[firstUse])}`).toBeLessThan(firstUse);
    if (name === nativeJob) {
      const readPin = steps.findIndex((step: any) => step.id === 'rust-pin' && step.run);
      expect(readPin, name).toBeGreaterThanOrEqual(0);
      expect(install, name).toBeGreaterThan(readPin);
    }
  }
});
