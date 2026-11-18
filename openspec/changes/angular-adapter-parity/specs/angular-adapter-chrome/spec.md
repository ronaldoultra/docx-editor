# Spec Delta

## Purpose

Defines the editor controls that the Angular adapter ships around the painted pages, and the rules every control follows for arrangement, enabled state, localization, direction, styling, and security.

## ADDED Requirements

### Requirement: Arrangements come from the engine registry

The toolbar, menu bar, and context menu SHALL derive their default groups, items, and order from the engine's chrome registry. Slot identifiers SHALL be the engine's public `ChromeSlotId` values.

#### Scenario: Default toolbar

- **WHEN** a toolbar renders with no custom content
- **THEN** its groups and controls follow the engine's chrome groups in order, and each control carries its slot identifier

### Requirement: Hosts can override and extend chrome parts

A host SHALL be able to replace one control in a default set, add its own controls, and render a part without its defaults, without copying the default arrangement.

#### Scenario: Replace one toolbar control

- **WHEN** a host provides its own control for the `text.bold` slot
- **THEN** the toolbar renders the host control in that slot and keeps every other default control

### Requirement: Unwired controls render disabled with a reason

A control whose command the engine does not support in the current state SHALL render disabled, with the localized engine reason as its accessible description.

#### Scenario: Control without engine support

- **WHEN** the selection is inside a locked content control and the user opens the toolbar
- **THEN** formatting controls are disabled and expose the lock reason

### Requirement: Toolbar covers the shipped control set

The toolbar SHALL provide history, font family and size, text formatting, colors and highlight, paragraph style, alignment, lists and indent, line spacing, links, images, tables and table borders, charts, review authors, and direction controls.

#### Scenario: Apply a paragraph style

- **WHEN** the user picks Heading 1 in the paragraph style control
- **THEN** the paragraph takes the Heading 1 style and the control shows Heading 1

### Requirement: Menu bar, context menu, and keyboard

The adapter SHALL provide a menu bar and a context menu with the engine's commands. Both SHALL open from the keyboard, move focus with arrow keys, and close with Escape.

#### Scenario: Context menu on a table cell

- **WHEN** the user opens the context menu inside a table cell
- **THEN** the menu shows the clipboard items and the table row, column, merge, and split items

#### Scenario: Keyboard dismissal

- **WHEN** a menu is open and the user presses Escape
- **THEN** the menu closes and focus returns to the element that opened it

### Requirement: Navigation pane

The adapter SHALL provide a navigation pane with find, go to, and headings tabs, opened by the find shortcut unless the host turns the shortcut off.

#### Scenario: Find shortcut

- **WHEN** the user presses Ctrl+F, or Cmd+F on macOS, in the editor
- **THEN** the navigation pane opens on the find tab with focus in the search field

### Requirement: Popups and dialogs

The adapter SHALL provide the hyperlink and content-control popups, and the page setup, paragraph, watermark, bookmark, split cell, text form field, export, print, revision markup, compare documents, note properties, image properties, and alt text dialogs.

#### Scenario: Page setup dialog

- **WHEN** the user opens page setup, changes the orientation to landscape, and confirms
- **THEN** the section becomes landscape and the pages repaint

#### Scenario: Cancel a dialog

- **WHEN** the user changes a field in a dialog and cancels
- **THEN** the document does not change

### Requirement: Object, image, shape, and chart controls

The adapter SHALL provide image insert, wrap, and properties controls, and shape, chart, and object panes for the selected drawing.

#### Scenario: Insert an image

- **WHEN** the user inserts a PNG file from the image control
- **THEN** the image appears at the caret and undo removes it

### Requirement: Header, footer, and notes chrome

The adapter SHALL provide header and footer chrome and notes chrome, and SHALL show header and footer content through the engine's painted pages, never through a separate editable view.

#### Scenario: Edit a header

- **WHEN** the user double-clicks a header
- **THEN** the header enters edit mode with its labels visible, and the body content is dimmed

### Requirement: Page furniture

The adapter SHALL provide horizontal and vertical rulers, a page number indicator, a loading state, an open-failure state, a font notice, and unsupported-element notices.

#### Scenario: Document fails to open

- **WHEN** the host passes bytes that are not a DOCX package
- **THEN** the open-failure state shows a localized message and the host receives the failure

### Requirement: Review chrome

The adapter SHALL provide suggesting mode, review author styles, revision markup display, and the review gutter.

#### Scenario: Suggest an edit

- **WHEN** the editor is in suggesting mode and the user types a word
- **THEN** the word appears as a tracked insertion with the author's style

### Requirement: Localized and direction-aware chrome

Every user-facing string in the chrome SHALL come from `@docx-editor.dev/i18n` keys. Missing translations SHALL fall back to English. A right-to-left catalog or direction input SHALL mirror the chrome and SHALL NOT change paragraph direction in the document.

#### Scenario: Portuguese catalog

- **WHEN** a host provides the `pt-BR` catalog
- **THEN** the toolbar labels show Portuguese text

#### Scenario: Right-to-left catalog

- **WHEN** a host provides a catalog with `_dir` set to `rtl`
- **THEN** the chrome has `dir="rtl"` and the document paragraphs keep their own direction

### Requirement: Chrome styling comes from the core stylesheet

Chrome SHALL use only the classes and `--doc-*` tokens of the core stylesheet. The Angular package SHALL NOT add component styles that the core stylesheet does not hold.

#### Scenario: Color mode

- **WHEN** a host sets the color mode to dark
- **THEN** the chrome uses the dark `--doc-*` tokens and the document canvas keeps its Word colors

### Requirement: Chrome treats document values as untrusted

Chrome SHALL NOT build HTML from strings, SHALL pass every URL through the engine's URL sanitizer before use, and SHALL NOT fetch external resources without an explicit user action.

#### Scenario: Script URL in a hyperlink

- **WHEN** the user opens the hyperlink popup on a link with a `javascript:` target
- **THEN** the popup shows no active link and opening it does nothing
