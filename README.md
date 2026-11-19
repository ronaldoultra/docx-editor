# docx-editor (fork)

This repository is a fork of [eigenpal/docx-editor](https://github.com/eigenpal/docx-editor), a WYSIWYG editor and rendering engine for DOCX files in the browser. The fork keeps only the Apache-2.0 engine and the packages that an Angular interface needs.

## Packages

| Package | Description |
| --- | --- |
| [`packages/angular`](packages/angular) | The Angular component `<docx-editor>`: load, show pages, edit text, and save. It has no toolbar, menus, or dialogs. |
| [`packages/core`](packages/core) | The engine: OOXML read and write, the document tree, layout, paint, and editing. It has no framework dependency. It also holds the interface strings (`@docx-editor.dev/core/i18n`) and open-licensed substitutes for common DOCX default fonts (`@docx-editor.dev/core/fonts`). |
| [`packages/vue`](packages/vue) | The upstream Vue 3 adapter. It stays only as a reference for the Angular port. |

The fork removed the upstream packages under the EigenPal Pro License (`pro`, `editor-api`, `docx-to-pdf`, and `telemetry`). It also removed `react`, `nuxt`, `docx-to-markdown`, `fonts-cjk`, the examples, and the release tooling.

## Development

The repository uses [Bun](https://bun.sh/) workspaces. Use the Bun version in `.bun-version` and Node.js `^20.16.0 || >=22.3.0`. Run these commands from the repository root:

```bash
bun install
bun run build:packages
bun run typecheck
bun run lint
bun run test
```

To run only the engine tests, pass the `core` test files to the test runner, or run `bun test packages/core`.

## Angular example

The example in [`examples/angular`](examples/angular) opens a `.docx` file, edits it, and downloads the result. Build the packages, then start it:

```bash
bun run build:packages
bun run dev:angular
```

Open `http://localhost:4200/`. Add `?zone` to the URL to start the app with zone.js. Without it, the app uses zoneless change detection.

The browser tests for the example use Playwright:

```bash
bun run --filter './examples/angular' build
bun run test:e2e:angular
```

To use an installed Google Chrome instead of the Playwright Chromium download, set `PLAYWRIGHT_CHANNEL=chrome`.

## Translations

The `i18n` lane of `core` has catalogs for English and Portuguese (Brazil). To add a locale, see the [i18n contribution guide](docs/i18n.md).

## License

The code in this repository uses the [Apache License 2.0](LICENSE). Bundled fonts in `packages/core/src/fonts/assets` keep their own licenses, in `packages/core/licenses/fonts`.
