import { test, expect } from 'playwright/test';
import type { SessionFramesProbe } from './pptx-session-frames-harness';

test('pptx: worker session frames paint the same pixels as a local handle', async ({ page }) => {
  const consoleErrors: string[] = [];
  const pageErrors: string[] = [];
  page.on('console', (message) => {
    if (message.type() === 'error') consoleErrors.push(message.text());
  });
  page.on('pageerror', (error) => pageErrors.push(error.message));
  await page.goto('/pptx-session-frames.html');
  await page.waitForFunction(() => '__pptxSessionFrames' in window);
  const { results, workers } = await page.evaluate(async () => {
    const probe = (window as unknown as { __pptxSessionFrames: SessionFramesProbe })
      .__pptxSessionFrames;
    return { results: await probe.done, workers: probe.workers };
  });
  const failures = results.filter(
    (result) =>
      result.differingPixels !== 0 || result.maxChannelDelta !== 0 || !result.versionMatches
  );
  expect.soft(results).toHaveLength(12);
  expect.soft(failures, `Failing frame cases:\n${JSON.stringify(failures, null, 2)}`).toEqual([]);
  for (const [fixture, slideCount] of [['demo', 3], ['tiff', 1]] as const) {
    for (let slideIndex = 0; slideIndex < slideCount; slideIndex += 1) {
      const cases = results.filter(
        (result) => result.fixture === fixture && result.slideIndex === slideIndex
      );
      expect.soft(
        cases.map((result) => `${result.dpr}/${result.zoom}`).sort(),
        `${fixture} slide=${slideIndex} scale coverage`
      ).toEqual(['1/1', '1/1.5', '2/1']);
    }
    const imageCases = results.filter(
      (result) => result.fixture === fixture && result.slideIndex === 0
    );
    expect.soft(imageCases, `${fixture} image cases`).toHaveLength(3);
    expect.soft(
      imageCases.filter((result) => result.mainImagesDrawn === 0 || result.sessionImagesDrawn === 0),
      `${fixture} must draw decoded media in both arms at every scale`
    ).toEqual([]);
    expect.soft(
      workers.filter((worker) => worker.fixture === fixture),
      `${fixture} must use one default session worker`
    ).toHaveLength(1);
  }
  expect.soft(workers, 'Expected one default session worker per fixture').toHaveLength(2);
  expect.soft(
    workers.every((worker) => worker.type === 'module' && worker.url.includes('pptxSessionWorker')),
    `Expected default module workers: ${JSON.stringify(workers)}`
  ).toBe(true);
  expect.soft(consoleErrors, 'Console errors while rendering frames').toEqual([]);
  expect.soft(pageErrors, 'Uncaught errors while rendering frames').toEqual([]);
});
