# Tasks

## 1. Test harness and parity ledger

- [x] 1.1 Add the Angular unit-test builder (`@angular/build:unit-test`, runner `vitest`) with an `angular.json` in `packages/angular`, and a root `test:angular` script; verify a trivial `docx-editor.component.spec.ts` that mounts `<docx-editor>` with a fixture document passes under `bun run test:angular`
- [x] 1.2 If task 1.1 cannot run engine layout under the builder, switch to `@analogjs/vitest-angular` with the same spec file, record the reason in `design.md` D5, and verify the same spec passes (not needed: task 1.1 runs engine layout under the builder; see D5)
- [x] 1.3 Make `scripts/test/run-parallel.mjs` and `scripts/test/changed.mjs` skip `packages/angular`, and print a `test:angular` reminder from `test:changed` when `packages/angular` or `packages/core/src/editor` changed; verify with `bun run test:changed --list` and a script test in `scripts/__tests__/`
- [x] 1.4 Add a spec that asserts the Angular tests load one copy of `@docx-editor.dev/core`, and verify it fails when a second copy is aliased in
- [x] 1.5 Write `parity-ledger.md` with one row per non-deprecated `@docx-editor.dev/vue` export and per `packages/vue/test` file, every row `todo`; verify the row count matches the API snapshot and the test directory
- [x] 1.6 Add `scripts/check-angular-parity.mjs` and a `check:angular-parity` script that fail on a missing ledger row or on an Angular export without a TSDoc release tag; verify with a script test that covers a missing row and a `dropped` row without a reason
- [x] 1.7 Add `bun run test:angular` and `bun run check:angular-parity` to the Verify list in `AGENTS.md`, and verify the commands run as written

## 2. Base layer

- [x] 2.1 Add `DocxEditorContext` and the `docxEditorRoot` directive that create one engine, provide it, and destroy it; verify specs for two independent roots and for release on destroy
- [x] 2.2 Add the viewport and content components with attach and detach; verify a spec that detach keeps undo history and selection, and a spec that the scroller has `docx-editor__scroll-container` and `role="region"`
- [x] 2.3 Add the single engine subscription and `injectEditorState(selector, isEqual)` (D2); verify specs that an unrelated change does not notify, that a page-count read sees the post-layout count, and that one listener serves many consumers
- [x] 2.4 Add `injectEditorCommand`, `injectEditorValueCommand`, and `injectEditorEvent` over `toolbarCommandState` and `runToolbarCommand`; port the claims of `use-editor-command.test.ts` and verify the specs pass
- [x] 2.5 Add `provideDocxEditorI18n`, `injectTranslation`, chrome direction, and the `locale` setter; port the claims of `toolbar-direction.test.ts`, `form-field-i18n.test.ts`, and `paragraph-style-labels.test.ts`, and verify the specs pass (the toolbar and packaged-editor claims of these files move to tasks 3.1, 4.2, and 4.5)
- [x] 2.6 Add color mode and scope class handling; port the claims of `root-color-mode.test.ts` and `scope-dispose.test.ts`, and verify the specs pass
- [x] 2.7 Add zone handling for zone.js and zoneless apps; verify one spec with zone.js and one zoneless spec that controls update after a committed change
- [x] 2.8 Add host-window resolution through `editor/host-realm.ts` helpers; port the `DocxEditorWindow` claims of `host-window-chrome.test.ts` and verify the spec passes (the picker claim moves to task 4.2)
- [x] 2.9 Document the composition API in `packages/angular/README.md` with a root-viewport-content example, update the ledger rows, and verify `bun run check:angular-parity` and `bun run api:check` pass

## 3. Packaged component

- [x] 3.1 Rebuild `<docx-editor>` on the base layer (D4) with `document`, `mode` including `suggesting`, `locale`, `i18n`, `colorMode`, `zoom`, `zoomMode`, `author`, `fieldResults`, `contentLockPolicy`, and `lockedControlNavigation`; port the claims of `live-props.test.ts`, `field-results-prop.test.ts`, `content-lock-policy-prop.test.ts`, `locked-control-navigation-prop.test.ts`, and `editing-mode-author.test.ts`, and the packaged-editor claims of `form-field-i18n.test.ts` and `toolbar-direction.test.ts`
- [x] 3.2 Add construction-time inputs `modules`, `fonts`, `fontShaping`, and `imageDecodePort` that recreate the engine; port the claims of `fonts-prop.test.ts`, `font-origins.test.ts`, `font-shaping-lifecycle.test.ts`, and `image-decode-port-prop.test.ts`
- [x] 3.3 Add outputs `change`, `selectionChange`, `openFailure`, `error`, and methods `load`, `save`, `focus`, `exec`, `snapshot`, `getEditor`; port the claims of `lifecycle-regression.test.ts`, `sugar-host.test.ts`, and `paginated-form-save.test.ts`
- [x] 3.4 Update `examples/angular` to use the new inputs, extend `examples/angular/e2e/editor.e2e.ts` for live zoom and save, and verify `bun run test:e2e:angular` passes
- [x] 3.5 Re-extract the API snapshot, update the ledger and README, and verify `bun run api:check` and `bun run check:angular-parity` pass

## 4. Toolbar

- [ ] 4.1 Add the toolbar compound with registry-driven groups and `docxSlot` overrides (D3); port the claims of `chrome-composition.test.ts` and `toolbar-host-groups.test.ts`
- [ ] 4.2 Add history, text formatting, font family and size, and paragraph style controls; port the claims of `toolbar-picker-accessibility.test.ts` and `toolbar-dev-warnings.test.ts`, the claim of `paragraph-style-labels.test.ts`, and the picker claim of `host-window-chrome.test.ts`
- [ ] 4.3 Add the color picker, highlight, alignment, lists, indent, and line spacing controls, and verify specs for each control's enabled state and value
- [ ] 4.4 Add table, table border, and table cell fill controls; verify specs that cover border target labels and disabled reasons outside a table
- [ ] 4.5 Add link and direction controls; port the claims of `toolbar-link-platform.test.ts`, and the direction button claims of `toolbar-direction.test.ts`
- [ ] 4.6 Show the toolbar in `<docx-editor>` with the `chrome` input (`false` hides it), extend the browser test to bold a word from the toolbar, and verify `bun run test:e2e:angular` passes
- [ ] 4.7 Document toolbar composition and overrides in the README, update the ledger and API snapshot, and verify `bun run api:check` and `bun run check:angular-parity` pass

## 5. Menu bar, context menu, and keyboard

- [ ] 5.1 Add the menu bar compound from `CHROME_MENUS`; port the claims of `menu-composition.test.ts`, `menu-opening.test.ts`, and `menu-download-name.test.ts`
- [ ] 5.2 Add the context menu compound with clipboard, table, and TOC items; port the claims of `context-menu.test.ts`, `context-menu-target.test.ts`, `context-menu-start.test.ts`, `context-menu-placement.test.ts`, `context-menu-link.test.ts`, and `context-menu-chart.test.ts`
- [ ] 5.3 Add keyboard navigation and dismissal for menus and popups; port the claims of `context-menu-keyboard.test.ts` and `chrome-popup-dismissal.test.ts`
- [ ] 5.4 Add the `menu` and `contextMenu` inputs to `<docx-editor>`, extend the browser test to open the context menu on a table cell, and verify `bun run test:e2e:angular` passes
- [ ] 5.5 Update the README, ledger, and API snapshot, and verify `bun run api:check` and `bun run check:angular-parity` pass

## 6. Popups and dialogs

- [ ] 6.1 Add the `DOCX_EDITOR_POPUPS` token, the popup renderer, and dialog parts; port the claims of `popup-config.test.ts`, `engine-popup-renderers.test.ts`, and `dialog-customization.test.ts`
- [ ] 6.2 Add the hyperlink popup; port the claims of `hyperlink-as-child.test.ts` and `link-row-chrome.test.ts`, and the `javascript:` URL claim of `security.test.ts`
- [ ] 6.3 Add the page setup, paragraph, and watermark dialogs; port the claims of `paragraph-dialog.test.ts`, `paragraph-hooks-lock-state.test.ts`, and `watermark-dialog.test.ts`
- [ ] 6.4 Add the bookmark, split cell, text form field, and invalid text form field dialogs; port the claims of `bookmark-dialog.test.ts`
- [ ] 6.5 Add the export and print dialogs; port the claims of `export-popup.test.ts` and `print-popup.test.ts`
- [ ] 6.6 Add the equation popup; port the claims of `equation-popup.test.ts`
- [ ] 6.7 Add the `popups` and `hyperlinkPopup` inputs to `<docx-editor>`, update the README, ledger, and API snapshot, and verify `bun run api:check` and `bun run check:angular-parity` pass

## 7. Navigation

- [ ] 7.1 Add the navigation pane compound with find, go to, and headings tabs, and the find shortcut; port the claims of `navigation-pane-host.test.ts`, `navigation-goto.test.ts`, and `navigation-heading-page-break.test.ts`
- [ ] 7.2 Add search, outline, and go-to state helpers and the document outline; port the claims of `text-highlights.test.ts` and `highlights-background-layout.test.ts`
- [ ] 7.3 Add the `navigation` input to `<docx-editor>`, extend the browser test to find a word with Ctrl+F, and verify `bun run test:e2e:angular` passes
- [ ] 7.4 Update the README, ledger, and API snapshot, and verify `bun run api:check` and `bun run check:angular-parity` pass

## 8. Images, objects, shapes, and charts

- [ ] 8.1 Add image insert, wrap, alt text, and properties controls; port the claims of `image-authoring.test.ts`, `image-insert-refusal.test.ts`, `image-overlay-history.test.ts`, and `image-panel-rows.test.ts`
- [ ] 8.2 Add object controls and the side panel; port the claims of the object control types in `object-control-types.ts` as compile-time specs
- [ ] 8.3 Add the shape pane and the chart pane and chart insert; verify specs that a chart type change and a shape fill change commit through the engine
- [ ] 8.4 Update the README, ledger, and API snapshot, and verify `bun run api:check` and `bun run check:angular-parity` pass

## 9. Header, footer, notes, content controls, and page furniture

- [ ] 9.1 Add header and footer chrome; port the claims of `header-footer-chrome.test.ts`
- [ ] 9.2 Add notes chrome, note preview, note properties dialog, and the notes context menu; port the claims of `notes-context-menu.test.ts`
- [ ] 9.3 Add the content control compound and widget; port the claims of `content-control-widget-parts.test.ts`
- [ ] 9.4 Add the rulers and page number; port the claims of `page-number.test.ts` and `page-number-opening.test.ts`
- [ ] 9.5 Add the loading, open-failure, font notice, and unsupported-element parts; port the claims of `loading-indicator.test.ts`, `loading-continued.test.ts`, `open-failure-part.test.ts`, and `unsupported-elements.test.ts`
- [ ] 9.6 Add the `rulers` and `openFailure` inputs to `<docx-editor>`, update the README, ledger, and API snapshot, and verify `bun run api:check` and `bun run check:angular-parity` pass

## 10. Review

- [ ] 10.1 Add review author styles and the review gutter; port the claims of `review-gutter.test.ts`, `viewport-rail.test.ts`, and `use-comment-draft-request.test.ts`
- [ ] 10.2 Add revision markup display and its dialog; port the claims of `revision-markup.test.ts` and `review-overflow-navigation.test.ts`
- [ ] 10.3 Add the compare documents dialog and document refresh; port the claims of `document-refresh-edit.test.ts`
- [ ] 10.4 Add the `revisionMarkup` and `reviewDisplayMode` inputs to `<docx-editor>`, extend the browser test to type in suggesting mode, and verify `bun run test:e2e:angular` passes
- [ ] 10.5 Port the remaining claims of `editor-accessibility.test.ts`, `editor-composition.test.ts`, `runtime-audit.test.ts`, `selection-range-state.test.ts`, `scoped-chrome-anchor.test.ts`, `slot-ref.test.ts`, `props-primitives.test.ts`, `merge-arrangement.test.ts`, `deferred-notifier.test.ts`, and `hydration-renderer-slots.test.ts`, or mark them `dropped` with a reason
- [ ] 10.6 Verify the ledger has no `todo` row with `bun run check:angular-parity`

## 11. Remove the Vue adapter

- [ ] 11.1 Move the oxlint fixtures in `scripts/__tests__/oxlint-rules.test.ts` from `packages/vue` paths to `packages/angular` paths and remove the Vue overrides from `.oxlintrc.json`; verify `bun test scripts/__tests__/oxlint-rules.test.ts` and `bun run lint` pass
- [ ] 11.2 Remove the Vue entries from `scripts/lib/packages.mjs`, `scripts/api-forgotten-exports.baseline.json`, `scripts/__tests__/build-core-declarations.test.ts`, `scripts/__tests__/build-declarations.test.ts`, `packages/core/src/store/__tests__/adapter-authority.test.ts` (add the Angular package), and `packages/core/src/styles/__tests__/stylesheet-parses.test.ts`; verify those tests pass
- [ ] 11.3 Remove the Vue JSX settings and preload from `bunfig.toml`, and the Vue build from `build:packages` in `package.json`; verify `bun run build:packages` and `bun run test` pass
- [ ] 11.4 Delete `packages/vue`, `docs/api/docx-editor-vue/`, and the superseded changes `vue-composable-adapter-parity` and `vue-drawing-authoring-parity`, then run `bun install`; verify `vue` leaves `bun.lock`
- [ ] 11.5 Update `README.md`, `AGENTS.md`, and `packages/i18n/package.json` text that names the Vue adapter, and drop the Vue rows and the check from the ledger and `check-angular-parity.mjs`; verify the repository scan in `angular-adapter-packaging` finds no match
- [ ] 11.6 Run the full Verify list (`typecheck`, `lint`, `test`, `test:angular`, `test:e2e:angular`, `api:check`, `i18n:validate`, and `openspec validate angular-adapter-parity --strict`), compare with the baseline, and verify no new failure

## Workflow follow-up

- Archive this change after phase 11 merges, so that the `header-footer-editing` removal and the four new capability specs move into `openspec/specs/`.
- Publish a final `@docx-editor.dev/vue` release note that names `@docx-editor.dev/angular` as the replacement, if the Vue package was ever published from this fork.
