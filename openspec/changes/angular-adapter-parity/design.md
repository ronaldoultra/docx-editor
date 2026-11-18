# Design

## Context

See `proposal.md` for the motivation. The current state that shapes the approach:

- `packages/angular` has one standalone `OnPush` component (`docx-editor.component.ts`, 163 lines). It creates the engine with `createDocxEditor`, attaches it, and runs it outside the zone.
- `packages/vue/src` has about 41,400 lines. Most of it is UI: toolbar (7,000), top-level editor parts and dialogs (14,900), menu (3,300), components (3,100), navigation (2,700), images (2,600), context menu (1,900), objects (1,600). The engine logic it calls lives in `@docx-editor.dev/core/editor` (182 imports) and `@docx-editor.dev/core/contracts/editor` (77 imports).
- The engine already owns the shared logic that both adapters need: `CHROME_GROUPS`, `CHROME_MENUS`, `toolbarCommandState`, `runToolbarCommand`, `commandForSlot`, the snapshot contract, fonts, compare, refresh, and popups state.
- `packages/vue/test` holds 97 test files. They are the most precise record of the behavior that the Angular adapter must keep.
- `bun run test` (`scripts/test/run-parallel.mjs`) runs every `*.test.*` and `*.spec.*` file under `packages/` with Bun. `bunfig.toml` preloads a Vue JSX plugin.
- The Angular package compiles with TypeScript 5.9 through `ng-packagr`. Angular 20 does not support TypeScript 6 or 7.

## Goals / Non-Goals

**Goals:**

- Port behavior, not code. Each Angular part calls the same engine API as its Vue part and keeps the claims of the Vue tests.
- Keep the Vue package working and unchanged until the removal phase, so that it stays a runnable reference.
- Land the adapter in phases that each leave `main` releasable.

**Non-Goals:**

- Moving logic from `packages/vue` into `core`. If a Vue file holds engine logic that Angular also needs, the move is a separate, small change before the Angular port of that part.
- Porting the deprecated Vue shell (`Toolbar`, `TitleBar`, `DocxEditorShell`, `PaginatedDocxEditor`, and their siblings).
- Angular versions outside the current peer range (`^19 || ^20`).
- A visual redesign. The chrome keeps the core stylesheet's classes and look.

## Decisions

### D1. Map Vue concepts to Angular concepts one-to-one

| Vue adapter | Angular adapter |
| --- | --- |
| `DocxEditorRoot` + `provide(docxEditorKey)` | `DocxEditorRootDirective` (`docxEditorRoot`) with a component-level provider of `DocxEditorContext` |
| `ShallowRef<DocxEditorInstance \| null>` | `Signal<DocxEditorInstance \| null>` on `DocxEditorContext` |
| Composable `useX()` | Function `injectX()` that calls `inject()` and returns signals and functions; it runs in an injection context |
| `ref` / `computed` / `watch` | `signal` / `computed` / `effect` |
| `onScopeDispose` | `DestroyRef.onDestroy` |
| Named slots and compound namespaces (`DocxEditor.Toolbar.Bold`) | Standalone components plus `ng-template` override directives keyed by slot id |
| `definePopup` and popup config | A `DOCX_EDITOR_POPUPS` injection token that maps popup ids to component types, rendered with `NgComponentOutlet` |
| `LocaleProvider` + `useTranslation` | `provideDocxEditorI18n()` + `injectTranslation()`; the packaged component's `i18n` input feeds the same token |
| Packaged `DocxEditor` | Packaged `<docx-editor>` that composes the root, chrome, viewport, and content |

Why: a one-to-one map lets each Vue file and test map to one Angular file and test, which makes the parity ledger (D6) mechanical. Alternative considered: RxJS observables for state. Rejected, because signals match the engine's synchronous snapshot model, need no subscription management in templates, and work without zone.js.

### D2. One subscription per root; selectors memoize

`DocxEditorContext` holds one engine subscription to `change`, `selectionChange`, and `error`. It raises a `tick` signal after the engine's deferred notifier yields, as the Vue `deferred-notifier.ts` does, so that reads happen after layout publishes. `injectEditorState(selector, isEqual)` returns a `computed` over `tick` that keeps the last `(snapshot, slice)` pair and returns the old slice when `isEqual` holds. The tick update runs through `NgZone.run` only when zone.js is present.

Why: this keeps the cost of fifty controls at one listener, and it keeps the Vue behavior that a page-number control sleeps through a bold toggle. Alternative: one engine listener per control. Rejected, because it scales with control count and reorders reads before layout.

### D3. Overrides through slot templates, defaults from the registry

Each compound (toolbar, menu, context menu, navigation, hyperlink, content control) renders its default set from `CHROME_GROUPS` or `CHROME_MENUS`. A host overrides one slot with `<ng-template docxSlot="text.bold">…</ng-template>` inside the compound. The compound reads overrides with `contentChildren` and renders the override in place of the default. An input such as `[defaults]="false"` drops the default set.

Why: this keeps the Vue "default set plus in-place override" semantics without a copied arrangement. Alternative: a configuration array input. Rejected, because it cannot carry host templates and it duplicates the registry.

### D4. Packaged component composes the public parts

`<docx-editor>` is a thin template over the same public components that a composing host uses. It holds no logic that a composing host cannot reach. Inputs that the engine accepts live (`mode`, `zoom`, `locale`, `author`, and the rest) go through engine setters. Construction-time inputs (`modules`, `fontShaping`, `imageDecodePort`) recreate the engine and reopen the current document.

Why: if the packaged editor used private paths, composing hosts would drift from it. This also matches the Vue sugar component.

### D5. Vitest with the Angular unit-test builder

Angular tests run with Vitest through the Angular CLI unit-test builder (`@angular/build:unit-test`, runner `vitest`), configured in an `angular.json` inside `packages/angular`. The test environment is jsdom or happy-dom, whichever the builder supports for the engine's DOM needs. Tests are `packages/angular/src/**/*.spec.ts`. A root script `test:angular` runs them.

`scripts/test/run-parallel.mjs` and `scripts/test/changed.mjs` skip `packages/angular`, because Bun cannot compile Angular components. The Verify list in `AGENTS.md` gains `bun run test:angular`.

Why: the user chose Vitest. The builder compiles components the same way the app build does. Alternative: `@analogjs/vitest-angular` with a plain `vitest.config.ts`. It is the fallback if the builder is unusable with Angular 20 in this workspace (the builder is experimental in Angular 20). The first task proves one or the other.

### D6. A parity ledger drives the order and the removal gate

`openspec/changes/angular-adapter-parity/parity-ledger.md` lists every Vue public export and every Vue test file. Each row states its Angular counterpart (export name or spec file) and a status: `todo`, `ported`, `dropped` (with a reason, such as "deprecated shell" or "Vue-only mechanism"). A small script, `scripts/check-angular-parity.mjs`, reads the Vue API snapshot and the ledger and fails when a non-deprecated Vue export or a Vue test file has no row. The Vue package can leave only when the ledger has no `todo` row.

Why: the Vue surface is large, and a name-only check lets an incomplete port pass. The ledger records the claim of each Vue test in its Angular port, so a reviewer can check behavior, not names. Alternative: delete Vue tests as each part ports. Rejected, because the deletion would hide gaps.

### D7. Phase order

1. Test harness and ledger.
2. Base: context, root, viewport, content, state and command helpers, events, i18n, color mode.
3. Packaged component on the base, with the inputs that need no chrome.
4. Toolbar and color picker.
5. Menu bar, context menu, keyboard shortcuts.
6. Popups and dialogs.
7. Navigation pane and outline.
8. Images, objects, shapes, charts, side panel.
9. Header, footer, notes, content controls, page furniture (rulers, page number, loading, open failure, font notice, unsupported elements).
10. Review: suggesting mode, authors, markup, gutter, compare, refresh.
11. Removal of `packages/vue`.

Why: each phase depends only on earlier phases. The packaged component appears in phase 3 so that the example and the browser tests exercise each later phase as it lands.

### D8. Line caps and file layout

Angular files follow the repository line caps (1000 lines). Each component keeps its template inline when the template is short, and in a sibling `.html` file when it is long. Directory layout mirrors the Vue `editor/` subdirectories (`toolbar/`, `menu/`, `navigation/`, …) so that the ledger rows map by path.

## Risks / Trade-offs

- [The Angular unit-test builder is experimental in Angular 20] → The harness task proves it first. The fallback is the Analog Vitest plugin, with the same spec files.
- [Two test runners: Bun for the engine, Vitest for Angular] → `test:angular` is in the Verify list, and `test:changed` prints a note to run it when a change touches `packages/angular` or `packages/core/src/editor`.
- [Engine copies: Vitest resolves `@docx-editor.dev/core` from source while the app resolves `dist`] → Tests resolve one copy through one alias, and a spec asserts one shaper instance.
- [Change detection cost of many signals] → D2 memoization, and a spec that counts notifications for unrelated changes.
- [Ported tests lose precision] → The ledger keeps each Vue test's claims next to its Angular spec. A review checks the claims, not only the file count.
- [The adapter grows for months while Vue still exists] → Vue stays untouched and runnable until phase 11, so any behavior question has a live reference.
- [Removing Vue removes lint and test coverage of shared rules] → Phase 11 moves the oxlint rule fixtures in `scripts/__tests__/oxlint-rules.test.ts` from Vue paths to Angular paths before the Vue files leave.

## Migration Plan

- Phases 1–10 add code to `packages/angular` only, plus small tooling edits. Each phase ships behind no flag, because the new parts are additive and the packaged component keeps its current inputs working.
- Phase 11 deletes `packages/vue`, its snapshots, and its tooling references in one commit, after `scripts/check-angular-parity.mjs` passes with no `todo` rows.
- Rollback for phases 1–10 is a revert of the phase. Rollback for phase 11 is a revert of its single commit.
- Consumers of `@docx-editor.dev/vue` get a final release note that names `@docx-editor.dev/angular` as the replacement.

## Open Questions

- Which DOM environment (jsdom or happy-dom) the Angular unit-test builder uses best with the engine's layout code. The harness task decides it; the specs and tasks do not depend on the answer.
- Whether the Angular adapter keeps the Vue export names (`DocxEditorToolbar`) or uses Angular-style names (`DocxEditorToolbarComponent`). The ledger records the chosen name per row; the tasks do not change.
