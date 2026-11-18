# Proposal

## Why

This fork ships Angular as its framework adapter, but `@docx-editor.dev/angular` covers only load, edit, mode, and save. Everything a user sees around the pages — toolbar, menus, context menu, dialogs, popups, navigation pane, rulers, review — exists only in `packages/vue`, which the repository keeps solely as a reference for the port. Until the Angular adapter reaches parity, the Vue package cannot leave, and the repository carries a second adapter with its own build, lint rules, test preload, and API snapshot.

## What Changes

- Grow `packages/angular` from one component to a full adapter over the same engine API that the Vue adapter uses. The adapter keeps owning no editing state: every control reads the engine snapshot and dispatches engine commands.
- Add a composition layer: a root that owns one engine instance per editor and provides it through Angular dependency injection, a viewport and content mount, and standalone components for each chrome part.
- Add injectable services and signal-returning functions that replace the Vue composables (`useEditorState`, `useEditorCommand`, `useZoom`, `usePageSetup`, and the rest), named for Angular conventions.
- Port the chrome: toolbar, menu bar, context menu, navigation pane, hyperlink and content-control popups, every dialog, rulers, page number, loading and open-failure states, font notice, image, shape, chart, and object controls, header/footer and notes chrome, and review markup.
- Extend the packaged `<docx-editor>` component with the inputs and outputs the Vue `DocxEditor` accepts (`fonts`, `locale`, `zoom`, `author`, `modules`, `colorMode`, `chrome`, `menu`, `rulers`, `navigation`, `revisionMarkup`, and the rest), the `'suggesting'` mode, and instance methods (`load`, `focus`, `exec`, `snapshot`, `getEditor`).
- Add an Angular `LocaleProvider` equivalent: a translation function from `@docx-editor.dev/i18n`, the `locale` input, and right-to-left chrome.
- Add Vitest component tests for the Angular adapter, and port each Vue test's claims before the Vue test leaves.
- Do not port the deprecated Vue shell: `Toolbar`, `ToolbarButton`, `ToolbarGroup`, `TitleBar`, `MenuBar`, `DocumentName`, `Logo`, `TitleBarRight`, `DocxEditorShell`, `PaginatedDocxEditor`, and `PaginatedDocxEditorShell`.
- **BREAKING**: Remove `packages/vue`, its API snapshots under `docs/api/docx-editor-vue/`, and every build, lint, test, and documentation reference to it. Consumers of `@docx-editor.dev/vue` move to `@docx-editor.dev/angular`.
- **BREAKING**: The `<docx-editor>` default for chrome changes from "no chrome" to the packaged toolbar and menus, as the Vue `DocxEditor` does. Hosts that want only pages set `chrome` to `false`.

## Capabilities

### New Capabilities

- `angular-adapter-composition`: The editor root, viewport, and content mount; how one engine instance is created, shared through dependency injection, attached, moved, and destroyed; zone behavior.
- `angular-adapter-state`: The injectable services and signal-returning functions that expose engine state, commands, events, and document features to Angular code.
- `angular-adapter-chrome`: The toolbar, menus, context menu, navigation, popups, dialogs, panes, rulers, and review chrome, including localization, right-to-left layout, and enabled-state rules.
- `angular-adapter-packaging`: The packaged `<docx-editor>` component, the public API surface, the package build, peers, styles, and the requirement that the workspace holds one framework adapter.

### Modified Capabilities

- `header-footer-editing`: The requirement "React and Vue adapter parity" names two adapters that this fork does not ship. It is removed; the Angular header/footer chrome requirement in `angular-adapter-chrome` replaces it.

## Impact

- **Code**: `packages/angular/src` grows from one file to a full adapter. `packages/vue` is deleted at the end.
- **Tooling**: `package.json` (`build:packages`, a new `test:angular` script), `bunfig.toml` (the Vue JSX preload), `.oxlintrc.json` (Vue overrides), `scripts/lib/packages.mjs`, `scripts/api-forgotten-exports.baseline.json`, `scripts/__tests__/*` that name `vue`, and `packages/core/src/store/__tests__/adapter-authority.test.ts`.
- **Public API**: New `@public` exports in `@docx-editor.dev/angular` with an API Extractor snapshot. The `@docx-editor.dev/vue` package is no longer published.
- **Dependencies**: Vitest and the Angular Vitest integration as dev dependencies of `packages/angular`. `vue` leaves the lockfile.
- **Docs**: `README.md`, `AGENTS.md`, `packages/angular/README.md`, and `examples/angular`.
- **Other changes**: `vue-composable-adapter-parity` and `vue-drawing-authoring-parity` are superseded and leave with the Vue package.
