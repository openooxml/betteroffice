/**
 * Turns a list of scenarios into bun tests: one test per scenario and pinned
 * sample, every scenario recorded through the harness, and the format's run
 * closed out (printed, published, recorded or compared) after the last one.
 */

import { afterAll, beforeAll, describe, it } from 'bun:test';

import type { Format, PinnedSample } from './corpus';
import { loadSample, samplesFor } from './corpus';
import { ScenarioRecorder, e2eEnabled, finishFormat } from './harness';
import type { ScenarioRun } from './harness';

export interface Scenario<Ctx> {
  name: string;
  description: string;
  /** Who takes part: `web`, `web:a`, `python`, ... as the ops are stamped. */
  participants: string[];
  /** Run on every pinned sample (default) or only the first, for slow scenarios. */
  samples?: 'all' | 'first';
  /** A reason to skip, e.g. a missing Python environment. */
  requires?: () => string | undefined;
  run(ctx: Ctx): void | Promise<void>;
}

export interface SuiteOptions<Ctx> {
  /** Runs once before the scenarios; idempotent wasm init goes here. */
  setup?: () => void | Promise<void>;
  /** Builds the per-run context; `dispose` runs after the scenario. */
  context(sample: PinnedSample, bytes: Uint8Array, recorder: ScenarioRecorder): Promise<Ctx & { dispose?(): void }> | (Ctx & { dispose?(): void });
  /** Per-test timeout; cross-SDK scenarios boot an interpreter. */
  timeoutMs?: number;
}

export function defineSuite<Ctx>(format: Format, scenarios: Scenario<Ctx>[], options: SuiteOptions<Ctx>): void {
  const suite = e2eEnabled() ? describe : describe.skip;
  const runs: ScenarioRun[] = [];
  suite(`${format} end-to-end`, () => {
    if (options.setup) beforeAll(options.setup);
    afterAll(() => finishFormat(format, runs));
    for (const scenario of scenarios) {
      const samples = scenario.samples === 'first' ? samplesFor(format).slice(0, 1) : samplesFor(format);
      const reason = scenario.requires?.();
      for (const sample of samples) {
        const meta = { scenario: scenario.name, sample: sample.id, description: scenario.description, participants: scenario.participants };
        if (reason) {
          it.skip(`${scenario.name} on ${sample.id} (${reason})`, () => {});
          runs.push(new ScenarioRecorder(format, meta).skipped(reason));
          continue;
        }
        it(
          `${scenario.name} on ${sample.id}`,
          async () => {
            const recorder = new ScenarioRecorder(format, meta);
            const bytes = await loadSample(sample);
            const ctx = await options.context(sample, bytes, recorder);
            try {
              await scenario.run(ctx);
            } finally {
              ctx.dispose?.();
            }
            runs.push(recorder.finish());
          },
          options.timeoutMs ?? 120_000
        );
      }
    }
  });
}
