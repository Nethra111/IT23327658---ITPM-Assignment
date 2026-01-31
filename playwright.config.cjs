const { defineConfig } = require("@playwright/test");

module.exports = defineConfig({
  testDir: "./tests",
  testMatch: /.*\.spec\.js/,
  timeout: 120000,
  expect: { timeout: 30000 },
  fullyParallel: false,
  workers: 1,
  reporter: [["html"], ["list"]],
  use: {
    headless: true,
    viewport: { width: 1280, height: 720 },
    actionTimeout: 30000,
    navigationTimeout: 60000,
    trace: "on-first-retry"
  }
});
