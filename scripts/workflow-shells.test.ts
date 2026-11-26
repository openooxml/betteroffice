import { describe, expect, test } from 'bun:test';
import { readFileSync, readdirSync } from 'node:fs';

const workflows = new URL('../.github/workflows/', import.meta.url);
const actions = new URL('../.github/actions/', import.meta.url);

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

test('CI installs Rust from the repository toolchain pin', () => {
  const pin = Bun.TOML.parse(readFileSync(new URL('../rust-toolchain.toml', import.meta.url), 'utf8')) as any;
  expect(pin.toolchain.channel).toMatch(/^\d+\.\d+\.\d+$/);

  const upstream = /^dtolnay\/rust-toolchain@/i;
  const nativeJob = 'native-dist.yml bindings';
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

  for (const { name, job } of jobs) {
    const steps = job.steps ?? [];
    if (name === nativeJob) {
      const readPin = steps.findIndex((step: any) => step.id === 'rust-pin' && step.run);
      const install = steps.findIndex((step: any) => upstream.test(step.uses ?? ''));
      expect(readPin, name).toBeGreaterThanOrEqual(0);
      expect(install, name).toBeGreaterThan(readPin);
      continue;
    }
    const wasm = steps.some((step: any) => step.uses === './.github/actions/wasm-toolchain');
    const cargo = steps.some((step: any) => /\bcargo[ \t]/.test(step.run ?? ''));
    const maturin = steps.some((step: any) => /^PyO3\/maturin-action@/i.test(step.uses ?? ''));
    if (wasm || (cargo && !maturin)) {
      expect(steps.map((step: any) => step.uses), name).toContain('./.github/actions/rust-toolchain');
    }
  }
});
