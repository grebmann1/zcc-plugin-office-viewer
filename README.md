# Office Viewer for Zana

Optional Zana Command Center plugin, version 0.1.0. This is a focused first release inspired by [Office Viewer for BB](https://github.com/VKirill/bb-plugin-office-viewer/tree/dcaaae50f84e8c701db521501d357b2fec59ffbf), not full upstream feature parity. The researched original source is preserved on the [`upstream/bb`](https://github.com/grebmann1/zcc-plugin-office-viewer/tree/upstream/bb/source) branch.

Open .docx, .xlsx and .pptx files through Zana’s project/thread file opener. Select a sheet or slide and quote its text to the current conversation draft.

## Install

Requires a current Zana 2.6.1 build with the thread-list and project-file SDK APIs. In **Plugins → Install**, paste:

```text
https://github.com/grebmann1/zcc-plugin-office-viewer.git
```

Or run:

```sh
zcc plugin install 'git:https://github.com/grebmann1/zcc-plugin-office-viewer.git@v0.1.0'
```

The repository includes compiled `app.js` and `server.mjs` entries and editable TypeScript sources. No global tools or paid analysis model are required. Plugins execute with full server trust after installation. This plugin sends no telemetry or document/trace content to an external service.

## Use

Open an Office file from the project Explorer or a thread attachment; Zana selects the Office file opener. The plugin’s page also explains supported formats. Use the refresh control to reload recorded data. Installation is optional; this repository does not change Zana core.

## Limits

Read-only content extraction, not pixel-perfect Office rendering. Images, charts, styling, date/number formatting, formula calculation and document editing are not included. Spreadsheet previews show stored values, or formula text when no cached value exists. Limits: 8 MiB input, 32 MiB expanded archive, 3,000 archive entries, 8-second parse deadline, two concurrent workers, 1,000 Word paragraphs, 100 slides, or 30 sheets × 200 rows × 50 columns. Password-protected and legacy binary formats are unsupported.

## Develop and verify

```sh
npm ci
npm run check
```

`check` runs TypeScript, meaningful parser/model/server/UI tests with an enforced 80% minimum for statements, branches, functions and lines, and rebuilds the committed artifacts. Build scripts use Zana’s shared React/SDK runtime instead of shipping another React instance. CI repeats these checks and rejects stale compiled artifacts. SDK ambient types are vendored; the small `types/current-host.d.ts` and `host.ts` additions reflect APIs in the current source host that the CLI facade has not yet incorporated.

Use an isolated built-Electron profile for installation, file access and reload checks; see `LIVE_TEST.md`. The Agent Traces repository contains the cross-plugin integration runner and expects sibling `agent-traces`, `office-viewer` and `burn` checkouts. No test changes your active Zana profile.

## Attribution

Upstream: [https://github.com/VKirill/bb-plugin-office-viewer](https://github.com/VKirill/bb-plugin-office-viewer/tree/dcaaae50f84e8c701db521501d357b2fec59ffbf), researched revision `dcaaae50f84e8c701db521501d357b2fec59ffbf`. See `UPSTREAM.md`, `upstream-source.json` and `LICENSE.upstream`. Zana adaptation: MIT, Guillaume Rebmann. This repository is independently maintained and is not an official BB or Zana release.
