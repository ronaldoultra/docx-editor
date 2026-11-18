import { defineConfig, devices } from '@playwright/test';

// Browser tests for the Angular example. Build first:
//   bun run build:packages && bun run --filter './examples/angular' build
export default defineConfig({
  testDir: '.',
  testMatch: '*.e2e.ts',
  timeout: 60_000,
  fullyParallel: false,
  workers: 1,
  reporter: 'list',
  use: {
    baseURL: 'http://localhost:4300',
    ...devices['Desktop Chrome'],
    viewport: { width: 1280, height: 900 },
    // Set PLAYWRIGHT_CHANNEL=chrome to use an installed Google Chrome instead of the
    // Playwright Chromium download.
    ...(process.env.PLAYWRIGHT_CHANNEL ? { channel: process.env.PLAYWRIGHT_CHANNEL } : {}),
  },
  webServer: {
    command: 'node serve.mjs',
    cwd: import.meta.dirname,
    url: 'http://localhost:4300',
    reuseExistingServer: false,
    timeout: 30_000,
  },
});
