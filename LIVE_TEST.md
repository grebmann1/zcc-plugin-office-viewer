# Live check: Office Viewer

From this repository, set `ZANA_REPO_ROOT` to a Zana source checkout and run:

```sh
ZANA_REPO_ROOT=/path/to/zana-command-center node integration/run-live.mjs
```

The runner stages its spec in that checkout, runs the standard isolated built-Electron fixture, and removes the temporary spec afterward. It installs **only Office Viewer** into a disposable HOME. It verifies the Settings sidebar link, native grouped rows, the usage disclosure, narrow-layout overflow, visible rebuild/reload while Settings stays open, real DOCX/XLSX/PPTX file previews and quoting, and rejected outside-project access. It also runs Zana's authoring/runtime baseline suites.

`--office-only` runs just the Office scenario. `--existing-build` requires a compatible complete production build. `OFFICE_VIEWER_TOOLCHAIN_CACHE` may point to an existing public Zana compiler dependency cache; it is copied read-only into the disposable HOME to avoid a redundant download. No tests alter the active profile or need live provider credentials.
