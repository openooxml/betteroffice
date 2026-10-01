import { defineConfig } from 'playwright/test';
import { resolve } from 'node:path';
import base from './playwright.config';

const root = resolve(import.meta.dirname, '../../..');
const port = Number(process.env.TYPING_PORT ?? 4191);

/** Benchmarks against the plugin-host harness: `bun run bench:docx-typing`. */
export default defineConfig({
  ...base,
  testMatch: '*.bench.ts',
  testIgnore: [],
  outputDir: resolve(root, '.source/e2e/perf/results'),
  timeout: 0,
  use: {
    ...base.use,
    channel: 'chromium',
    baseURL: `http://127.0.0.1:${port}`,
    trace: 'off',
  },
  webServer: process.env.TYPING_ARMS
    ? undefined
    : {
        command: `bunx vite --config vite.config.ts --host 127.0.0.1 --port ${port} --strictPort`,
        cwd: import.meta.dirname,
        url: `http://127.0.0.1:${port}/docx-typing.html`,
        reuseExistingServer: false,
        timeout: 120_000,
      },
});
