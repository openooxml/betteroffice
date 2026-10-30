import { defineConfig } from "playwright/test";
import { resolve } from "node:path";

const port = 4190;

export default defineConfig({
  testDir: import.meta.dirname,
  testMatch: "**/*.browser.ts",
  fullyParallel: false,
  workers: 1,
  retries: 0,
  timeout: 180_000,
  expect: { timeout: 20_000 },
  use: {
    browserName: "chromium",
    headless: true,
    actionTimeout: 20_000,
    navigationTimeout: 30_000,
    baseURL: `http://127.0.0.1:${port}`,
    viewport: { width: 1600, height: 1100 },
    locale: "en-US",
    serviceWorkers: "block",
    trace: "retain-on-failure",
    screenshot: "only-on-failure",
  },
  webServer: {
    command: "bun run build:web && bun e2e/server.ts",
    cwd: resolve(import.meta.dirname, ".."),
    url: `http://127.0.0.1:${port}/healthz`,
    reuseExistingServer: false,
    timeout: 180_000,
    env: { WOPI_E2E_PORT: String(port) },
  },
});
