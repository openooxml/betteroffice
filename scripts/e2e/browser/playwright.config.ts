import { defineConfig } from 'playwright/test';
import { resolve } from 'node:path';

const root = resolve(import.meta.dirname, '../../..');
const output =
  process.env.BETTEROFFICE_BROWSER_OUTPUT ??
  resolve(root, '.source/e2e/browser');

export default defineConfig({
  testDir: import.meta.dirname,
  testMatch: '**/*.browser.ts',
  outputDir: resolve(output, 'results'),
  reporter: [
    ['list'],
    ['html', { outputFolder: resolve(output, 'report'), open: 'never' }],
  ],
  fullyParallel: false,
  workers: 1,
  retries: 0,
  forbidOnly: Boolean(process.env.CI),
  timeout: 120_000,
  expect: { timeout: 15_000 },
  use: {
    browserName: 'chromium',
    headless: true,
    actionTimeout: 15_000,
    navigationTimeout: 30_000,
    baseURL: 'http://127.0.0.1:4186',
    viewport: { width: 1440, height: 1000 },
    locale: 'en-US',
    timezoneId: 'UTC',
    serviceWorkers: 'block',
    trace: 'retain-on-failure',
    screenshot: 'only-on-failure',
  },
  webServer: {
    command: 'bunx vite preview --host 127.0.0.1 --port 4186 --strictPort',
    cwd: resolve(root, 'apps/desktop'),
    url: 'http://127.0.0.1:4186',
    reuseExistingServer: false,
    timeout: 30_000,
  },
});
