import { summarize } from './harness';
import type { OpTiming, RecordedRun, ScenarioRun } from './harness';

export function scenario(
  ops: OpTiming[] = [{ op: 'edit', actor: 'a', e2eMs: 1 }]
): ScenarioRun {
  return {
    format: 'xlsx',
    scenario: 'editing',
    sample: 'fixture',
    sampleSha256: 'a'.repeat(64),
    description: 'fixture',
    participants: ['a', 'b'],
    status: 'passed',
    loadMs: 0,
    ops,
    summary: summarize(ops),
  };
}

export function recorded(run = scenario()): RecordedRun {
  return {
    schemaVersion: 3,
    format: run.format,
    commit: 'fixture',
    dirty: false,
    recordedAt: '2026-01-01T00:00:00Z',
    environment: {
      platform: 'linux',
      arch: 'x64',
      cpu: 'fixture',
      cpus: 4,
      bun: '1.3.14',
      rustc: 'fixture',
      python: '3.13',
    },
    scenarios: [run],
  };
}
