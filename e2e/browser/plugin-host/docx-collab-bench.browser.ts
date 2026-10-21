import { expect, test } from 'playwright/test';
import type {} from './docx-collab-bench-probe';

const RUNS = Number(process.env.COLLAB_BENCH_RUNS ?? 3);

test.skip(process.env.COLLAB_BENCH !== '1', 'collaboration bench runs only with COLLAB_BENCH=1');

test('docx collaboration bench: 200-edit join and steady remote typing', async ({ page }) => {
  test.setTimeout(600_000);
  await page.goto('/docx-collab-bench.html');
  await page.waitForFunction(() => '__docxCollabBench' in window);
  const cdp = await page.context().newCDPSession(page);
  await cdp.send('Performance.enable');
  const mainMs = async () => {
    const { metrics } = await cdp.send('Performance.getMetrics');
    return 1000 * (metrics.find((metric) => metric.name === 'TaskDuration')?.value ?? 0);
  };
  const run = async (label: string, scenario: () => Promise<{ frames: number; ms: number; paints: number; dispatchMs: number }>) => {
    const before = await mainMs();
    const result = await scenario();
    const main = (await mainMs()) - before;
    console.log(`[collab-bench] ${label} ${JSON.stringify({ ...result, mainMs: Math.round(main), mainMsPerFrame: Math.round(100 * main / result.frames) / 100,
      framesPerSec: Math.round(1000 * result.frames / result.ms) })}`);
  };
  await run('join200', () => page.evaluate(() => window.__docxCollabBench.join(200)));
  for (let index = 0; index < RUNS; index++) {
    await run(`type200-burst#${index}`, () => page.evaluate(() => window.__docxCollabBench.type(200, 0)));
    await run(`type100-paced16#${index}`, () => page.evaluate(() => window.__docxCollabBench.type(100, 16)));
  }
  expect(await page.evaluate(() => window.__docxCollabBench.errors)).toEqual([]);
});
