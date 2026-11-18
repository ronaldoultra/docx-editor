# Spec Delta

## Purpose

Defines the packaged Angular editor component, the public surface and build of `@docx-editor.dev/angular`, and the rule that the workspace holds one framework adapter.

## ADDED Requirements

### Requirement: Packaged editor component

The `<docx-editor>` component SHALL render a complete editor — root, chrome, viewport, and content — from a `document` input, with no other setup.

#### Scenario: Minimal host

- **WHEN** a host renders `<docx-editor [document]="bytes" />`
- **THEN** the document opens with the default toolbar and menus, and the user can edit and save it

### Requirement: Packaged editor inputs

The packaged component SHALL accept `document`, `mode` (`edit`, `view`, `suggesting`), `chrome`, `menu`, `contextMenu`, `navigation`, `rulers`, `hyperlinkPopup`, `openFailure`, `popups`, `title`, `zoom`, `zoomMode`, `fonts`, `fontShaping`, `imageDecodePort`, `colorMode`, `locale`, `i18n`, `author`, `modules`, `revisionMarkup`, `reviewDisplayMode`, `fieldResults`, `contentLockPolicy`, and `lockedControlNavigation`.

#### Scenario: Live input change

- **WHEN** the host changes the `zoom` input from 100 to 150 after the document opens
- **THEN** the pages repaint at 150 percent without a reload

#### Scenario: Construction-time input

- **WHEN** the host changes the `modules` input after the editor exists
- **THEN** the editor recreates the engine with the new modules and opens the current document again

#### Scenario: No chrome

- **WHEN** the host sets `chrome` to `false`
- **THEN** only the viewport and pages render

### Requirement: Packaged editor outputs and methods

The packaged component SHALL emit `change`, `selectionChange`, `openFailure`, and `error`, and SHALL expose `load`, `save`, `focus`, `exec`, `snapshot`, and `getEditor`.

#### Scenario: Save

- **WHEN** the host calls `save()` after an edit
- **THEN** it receives DOCX bytes that reopen with the edit

#### Scenario: Save without a document

- **WHEN** the host calls `save()` before a document opens
- **THEN** the returned promise rejects

### Requirement: Public API is versioned

Every public export of `@docx-editor.dev/angular` SHALL carry a TSDoc release tag, and the API Extractor snapshot under `docs/api/` SHALL match the build.

#### Scenario: Export added without a snapshot

- **WHEN** a change adds a public export and does not update the snapshot
- **THEN** `bun run api:check` fails

### Requirement: Engine stays a single peer

`@docx-editor.dev/core` SHALL stay a peer dependency of the Angular package and SHALL NOT be bundled into its output. The package SHALL NOT import ProseMirror or private engine lanes.

#### Scenario: Built package

- **WHEN** the package builds
- **THEN** its output imports `@docx-editor.dev/core` subpaths and contains no copy of engine code

### Requirement: Supported Angular versions

The package SHALL support the Angular major versions in its peer range, with and without zone.js, and SHALL build with the TypeScript version that those Angular versions support.

#### Scenario: Oldest supported major

- **WHEN** the example application builds against the oldest Angular major in the peer range
- **THEN** the build and the browser tests pass

### Requirement: The workspace holds one framework adapter

The workspace SHALL contain the Angular adapter as its only framework adapter. No build, lint, test, or API step SHALL refer to a Vue adapter.

#### Scenario: Repository scan

- **WHEN** a search runs for `@docx-editor.dev/vue` and `packages/vue` outside archived changes and changelogs
- **THEN** it finds no match
