import { readFile, writeFile, unlink } from "node:fs/promises";
import { resolve, join, dirname } from "node:path";
import { fileURLToPath } from "node:url";
import { spawn } from "node:child_process";
const root = process.env.ZANA_REPO_ROOT;
if (!root) throw new Error("Set ZANA_REPO_ROOT to a Zana source checkout");
const here = dirname(fileURLToPath(import.meta.url));
const name = `office-viewer-settings-${process.pid}.spec.ts`;
const staged = join(root, "e2e", name);
await writeFile(staged, await readFile(join(here, "settings-live.spec.txt")));
try {
  const child = spawn(process.execPath, [
    join(root, "node_modules/@playwright/test/cli.js"), "test", `e2e/${name}`,
    ...(process.argv.includes("--office-only") ? [] : ["e2e/plugin-authoring-live.spec.ts", "e2e/packaged-provider-startup.spec.ts"]),
    "--workers=1",
  ], {
    cwd: root, stdio: "inherit",
    env: { ...process.env, ZCC_E2E_BUILD: process.argv.includes("--existing-build") ? "0" : "1", OFFICE_VIEWER_ROOT: resolve(here, "..") },
  });
  process.exitCode = await new Promise((done, fail) => {
    child.once("error", fail); child.once("exit", code => done(code ?? 1));
  });
} finally { await unlink(staged); }
