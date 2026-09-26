import { defineConfig } from 'playwright/test';
import { resolve } from 'node:path';

const root = resolve(import.meta.dirname, '../../..');
const output = process.env.BETTEROFFICE_BROWSER_OUTPUT ?? resolve(root, '.source/e2e/browser');

export default defineConfig({
  testDir: import.meta.dirname,
  testMatch: '*.browser.ts',
  outputDir: resolve(output, 'plugin-host-results'),
  reporter: [['list']],
  fullyParallel: false,
  workers: 1,
  retries: 0,
  forbidOnly: Boolean(process.env.CI),
  timeout: 180_000,
  expect: { timeout: 30_000 },
  use: {
    browserName: 'chromium',
    headless: true,
    actionTimeout: 15_000,
    navigationTimeout: 120_000,
    baseURL: 'http://127.0.0.1:4187',
    viewport: { width: 1440, height: 1000 },
    locale: 'en-US',
    timezoneId: 'UTC',
    trace: 'retain-on-failure',
  },
  webServer: {
    command: 'bunx vite --config vite.config.ts --host 127.0.0.1 --port 4187 --strictPort',
    cwd: import.meta.dirname,
    url: 'http://127.0.0.1:4187',
    reuseExistingServer: false,
    timeout: 60_000,
  },
});
