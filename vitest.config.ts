import { defineConfig } from "vitest/config";
export default defineConfig({
  resolve: {
    alias: {
      "@zana-ai/zcc-plugin-sdk/app": new URL(
        "./test-runtime.ts",
        import.meta.url,
      ).pathname,
    },
  },
  test: {
    environment: "jsdom",
    include: ["*.test.ts", "*.test.tsx"],
    maxWorkers: 2,
    coverage: {
      provider: "v8",
      include: [
        "model.ts",
        "server.ts",
        "app.tsx",
        "ui.tsx",
        "parser.ts",
        "quote.ts",
        "worker.ts",
      ],
      thresholds: { statements: 80, branches: 80, functions: 80, lines: 80 },
    },
  },
});
