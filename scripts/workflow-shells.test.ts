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

const runSteps = readdirSync(workflows)
  .filter((file) => /\.ya?ml$/.test(file))
  .flatMap((file) => {
    const workflow = Bun.YAML.parse(readFileSync(new URL(file, workflows), 'utf8')) as any;
    return Object.entries<any>(workflow.jobs).flatMap(([id, job]) =>
      (job.steps ?? [])
        .filter((step: any) => step.run !== undefined)
        .map((step: any) => ({
          name: `${file} ${id}: ${label(step)}`,
          job,
          shell: step.shell ?? job.defaults?.run?.shell ?? workflow.defaults?.run?.shell,
        }))
    );
  });

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
    const steps = readdirSync(actions, { withFileTypes: true })
      .filter((entry) => entry.isDirectory())
      .flatMap((entry) => {
        const action = Bun.YAML.parse(readFileSync(new URL(`${entry.name}/action.yml`, actions), 'utf8')) as any;
        return (action.runs?.steps ?? [])
          .filter((step: any) => step.run !== undefined)
          .map((step: any) => ({ name: `${entry.name}: ${label(step)}`, shell: step.shell }));
      });
    expect(steps.length).toBeGreaterThan(0);
    for (const { name, shell } of steps) expect(shell, name).toBe('bash');
  });
});
