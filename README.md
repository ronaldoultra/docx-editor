# docx-editor (fork)

This repository is a fork of [eigenpal/docx-editor](https://github.com/eigenpal/docx-editor), a WYSIWYG editor and rendering engine for DOCX files in the browser. The fork keeps only the Apache-2.0 engine and the packages that an Angular interface needs.

## Packages

| Package | Description |
| --- | --- |
| [`packages/core`](packages/core) | The engine: OOXML read and write, the document tree, layout, paint, and editing. It has no framework dependency. |
| [`packages/i18n`](packages/i18n) | Interface strings and locale types. |
| [`packages/fonts`](packages/fonts) | Open-licensed substitutes for common DOCX default fonts. |
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

## Translations

The `i18n` package has catalogs for 15 locales. To add a locale, see the [i18n contribution guide](docs/i18n.md).

## License

The code in this repository uses the [Apache License 2.0](LICENSE). Bundled fonts in `packages/fonts` keep their own licenses.
