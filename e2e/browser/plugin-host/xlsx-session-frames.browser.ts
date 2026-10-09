import { test, expect } from 'playwright/test';
import type { SessionFramesProbe } from './xlsx-session-frames-harness';

test('xlsx: worker session frames paint the same pixels as a local handle', async ({
  page,
}) => {
  const consoleErrors: string[] = [];
  const pageErrors: string[] = [];
  page.on('console', (message) => {
    if (message.type() === 'error') consoleErrors.push(message.text());
  });
  page.on('pageerror', (error) => pageErrors.push(error.message));
  await page.goto('/xlsx-session-frames.html');
  await page.waitForFunction(() => '__xlsxSessionFrames' in window);
  const { results, workers } = await page.evaluate(async () => {
    const probe = (window as unknown as { __xlsxSessionFrames: SessionFramesProbe })
      .__xlsxSessionFrames;
    return { results: await probe.done, workers: probe.workers };
  });
  const failures = results.filter(
    (result) =>
      result.differingPixels !== 0 || result.maxChannelDelta !== 0 || !result.versionMatches
  );
  expect.soft(results).toHaveLength(12);
  expect.soft(failures, `Failing frame cases:\n${JSON.stringify(failures, null, 2)}`).toEqual([]);
  expect.soft(workers, 'Expected one default session worker per fixture').toHaveLength(2);
  expect.soft(
    workers.every((worker) => worker.type === 'module' && worker.url.includes('xlsxSessionWorker')),
    `Expected default module workers: ${JSON.stringify(workers)}`
  ).toBe(true);
  expect.soft(consoleErrors, 'Console errors while rendering frames').toEqual([]);
  expect.soft(pageErrors, 'Uncaught errors while rendering frames').toEqual([]);
});
