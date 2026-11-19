// Single source of truth for every published package's API Extractor
// configuration. Consumed by `scripts/api-extractor.mjs` (the root
// driver behind `api:extract` / `api:check`) and `scripts/build-docs-json.mjs`
// (the docs JSON orchestrator).
//
// Adding a new published package means adding one entry here. The
// per-package wrappers under `packages/*/scripts/` are gone — each
// package's `package.json` just calls the root driver with
// `--package <name>`.

import path from 'node:path';

export const PACKAGES = [
  {
    name: '@docx-editor.dev/i18n',
    root: 'packages/i18n',
    pkgSlug: 'docx-editor-i18n',
  },
  {
    name: '@docx-editor.dev/vue',
    root: 'packages/vue',
    pkgSlug: 'docx-editor-vue',
    tsconfigPath: 'packages/vue/tsconfig.api.json',
  },
  {
    name: '@docx-editor.dev/core',
    root: 'packages/core',
    pkgSlug: 'docx-editor-core',
    tsconfigPath: 'packages/core/tsconfig.api.json',
    // Unsupported-content signatures reference types owned by the layout/store subpaths.
    // Keep those graphs at their owning entry points, as with the existing facade escape hatches.
    forgottenExports: {
      logLevel: 'warning',
      allowlist: {
        editor: ['AnchoredDrawingRecord', 'EquationSpanRecord', 'InlineDrawingRecord', 'LayoutBox'],
        index: ['UnsupportedElementRule'],
        layout: ['UnsupportedContentRecord', 'UnsupportedContentResolver'],
        output: [
          'AnchoredDrawingRecord',
          'EquationSpanRecord',
          'InlineDrawingRecord',
          'LayoutBox',
          'OoxmlElement',
        ],
      },
    },
  },
  {
    name: '@docx-editor.dev/fonts',
    root: 'packages/fonts',
    pkgSlug: 'docx-editor-fonts',
    tsconfigPath: 'packages/fonts/tsconfig.api.json',
  },
  {
    name: '@docx-editor.dev/angular',
    // ng-packagr writes the typed `exports` map into the built package.json, so the
    // extractor reads the built package, not the source one.
    root: 'packages/angular/dist',
    pkgSlug: 'docx-editor-angular',
    tsconfigPath: 'packages/angular/tsconfig.api.json',
  },
];

// Derived: build invocation hint shown in `api:check` drift error
// output. Every package builds via the same `bun run --filter` shape,
// so it's computed from `name` rather than duplicated per entry.
export function buildHintFor(pkg) {
  return `bun run --filter '${pkg.name}' build`;
}

// Derived: where API Extractor writes (and reads-for-drift-check) the
// committed `<slug>.api.md` snapshots. Same path for all packages — one
// directory per package under `docs/api/`. Co-located with the rest of
// the docs tree, rather than the API Extractor default
// `<packageRoot>/etc/`.
export function reportDirFor(pkg, repoRoot) {
  return path.join(repoRoot, 'docs', 'api', pkg.pkgSlug);
}

export function packageByName(name) {
  return PACKAGES.find((p) => p.name === name);
}
