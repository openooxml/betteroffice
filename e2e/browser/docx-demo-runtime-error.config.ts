import { defineConfig } from 'playwright/test';
import { resolve } from 'node:path';

const root = resolve(import.meta.dirname, '../..');

export default defineConfig({
  testDir: resolve(root, 'e2e/browser'),
  testMatch: 'docx-demo-runtime-error.browser.ts',
  outputDir: process.env.BETTEROFFICE_BROWSER_OUTPUT
    ? resolve(process.env.BETTEROFFICE_BROWSER_OUTPUT)
    : resolve(root, '.source/e2e/docx-demo-runtime-error/results'),
  reporter: [['list']],
  fullyParallel: false,
  workers: 1,
  retries: 0,
  timeout: 120_000,
  expect: { timeout: 20_000 },
  use: {
    browserName: 'chromium',
    headless: true,
    baseURL: 'http://127.0.0.1:4187',
    viewport: { width: 1440, height: 1000 },
    locale: 'en-US',
    timezoneId: 'UTC',
    trace: 'retain-on-failure',
    screenshot: 'only-on-failure',
  },
  webServer: {
    command:
      'node ../../node_modules/next/dist/bin/next dev --hostname 127.0.0.1 --port 4187',
    cwd: resolve(root, 'apps/demo'),
    url: 'http://127.0.0.1:4187/docx',
    reuseExistingServer: false,
    timeout: 120_000,
    env: { NEXT_TELEMETRY_DISABLED: '1' },
  },
});
