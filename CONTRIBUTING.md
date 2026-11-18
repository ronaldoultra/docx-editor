# Contributing

This guide covers setup, tests, and changes for this fork of docx-editor.

## Prerequisites

- [Bun](https://bun.sh/), at the version in `.bun-version`.
- [Node.js](https://nodejs.org/) 24, the CI version.

## Development setup

Install the dependencies and build the packages:

```bash
bun install
bun run build:packages
```

If you work on parsing, serialization, or layout, run `bun run reference:fetch` to download the ECMA-376 reference files. The download includes about 58 MB of PDFs and ZIP archives. Git ignores these files. The repository includes reference summaries and XML schemas under `reference/`.

## Run tests

Run one test file during development:

```bash
bun test path/to/test.ts
```

Run the whole suite with `bun run test`. The runner starts one process per file and limits the number of workers by the free memory.

## Code style

The project uses [oxlint](https://oxc.rs/docs/guide/usage/linter.html) and Prettier. The pre-commit hook runs both checks. The lint configuration is `.oxlintrc.json`, and repository-specific rules live in `scripts/oxlint/docx-rules.mjs`.

```bash
bun run lint:fix
bun run format
```

## Make changes

1. Create a branch from `main`.
2. Read the [engine architecture](docs/architecture/production-engine-packages.md).
3. Keep each change focused.
4. Add or update package tests.
5. Verify the change:
   ```bash
   bun run typecheck
   bun run lint
   bun run format:check
   bun run test
   bun run api:check
   bun run i18n:validate
   ```

## Architecture overview

The canonical OOXML tree is the only editing state. Layout reads that tree and produces painted pages.

The painted pages are the editable surface. Browser mutations become tree operations instead of document markup.

ProseMirror is a projection inside `packages/core/src/binding/`. It never reconstructs the canonical tree.

For more information, see the [engine architecture](docs/architecture/production-engine-packages.md).

## Write documentation

Follow the [Google developer documentation style guide](https://developers.google.com/style/highlights) for READMEs and guides.

- Lead with the task or behavior. Use active voice and address the reader as "you."
- Keep sentences short, and give each sentence one idea.
- Use sentence-case headings, descriptive links, and exact identifiers in code formatting.
- State prerequisites, defaults, units, and limits near the relevant example.

## License

Code contributions use the [Apache License 2.0](LICENSE). Do not copy code from the upstream packages under the EigenPal Pro License into this repository.
