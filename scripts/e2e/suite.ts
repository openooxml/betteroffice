import { afterAll, beforeAll, describe, it } from 'bun:test';

import type { Format, PinnedSample } from './corpus';
import { loadSample, samplesFor } from './corpus';
import {
  ScenarioRecorder,
  e2eEnabled,
  errorText,
  finishFormat,
} from './harness';
import type { ScenarioRun } from './harness';

export interface Scenario<Ctx> {
  name: string;
  description: string;
  participants: string[];
  samples?: 'all' | 'first';
  requires?: () => string | undefined;
  run(ctx: Ctx): void | Promise<void>;
}

export interface SuiteOptions<Ctx> {
  setup?: () => void | Promise<void>;
  context(
    sample: PinnedSample,
    bytes: Uint8Array,
    recorder: ScenarioRecorder
  ):
    | Promise<Ctx & { dispose?(): void | Promise<void> }>
    | (Ctx & { dispose?(): void | Promise<void> });
  timeoutMs?: number;
}

export async function executeScenario<Ctx>(
  recorder: ScenarioRecorder,
  create: () => Promise<Ctx & { dispose?(): void | Promise<void> }>,
  run: (ctx: Ctx) => void | Promise<void>
): Promise<ScenarioRun> {
  let ctx: (Ctx & { dispose?(): void | Promise<void> }) | undefined;
  const errors: unknown[] = [];
  try {
    ctx = await create();
    await run(ctx);
  } catch (error) {
    errors.push(error);
  } finally {
    try {
      await ctx?.dispose?.();
    } catch (error) {
      errors.push(error);
    }
  }
  if (errors.length === 1) return recorder.failed(errors[0]);
  return errors.length
    ? recorder.failed(
        new AggregateError(errors, errors.map(errorText).join('\n'))
      )
    : recorder.finish();
}

export function defineSuite<Ctx>(
  format: Format,
  scenarios: Scenario<Ctx>[],
  options: SuiteOptions<Ctx>
): void {
  const suite = e2eEnabled() ? describe : describe.skip;
  const cases = scenarios.flatMap((scenario) => {
    const samples =
      scenario.samples === 'first'
        ? samplesFor(format).slice(0, 1)
        : samplesFor(format);
    return samples.map((sample) => ({
      scenario,
      sample,
      recorder: new ScenarioRecorder(format, {
        scenario: scenario.name,
        sample: sample.id,
        sampleSha256: sample.sha256,
        description: scenario.description,
        participants: scenario.participants,
      }),
    }));
  });
  const runs = cases.map(({ recorder }) =>
    recorder.failed('scenario did not complete (setup, filter, or timeout)')
  );
  suite(format + ' corpus integration', () => {
    if (options.setup)
      beforeAll(async () => {
        try {
          await options.setup!();
        } catch (error) {
          for (let i = 0; i < cases.length; i++)
            runs[i] = cases[i].recorder.failed(error);
          throw error;
        }
      });
    afterAll(() => finishFormat(format, runs));
    for (const [index, { scenario, sample, recorder }] of cases.entries()) {
      const reason = e2eEnabled() ? scenario.requires?.() : undefined;
      if (reason) {
        runs[index] = recorder.skipped(reason);
        it.skip(
          scenario.name + ' on ' + sample.id + ' (' + reason + ')',
          () => {}
        );
        continue;
      }
      it(
        scenario.name + ' on ' + sample.id,
        async () => {
          runs[index] = await executeScenario(
            recorder,
            async () =>
              options.context(sample, await loadSample(sample), recorder),
            (ctx) => scenario.run(ctx)
          );
          if (runs[index].status !== 'passed')
            throw new Error(runs[index].reason);
        },
        options.timeoutMs ?? 120_000
      );
    }
  });
}
