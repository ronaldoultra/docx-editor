# Spec Delta

## Purpose

Defines how Angular code reads engine state and runs engine commands through the adapter, so that custom controls behave the same as the packaged chrome.

## ADDED Requirements

### Requirement: Selected state is a read-only signal

The adapter SHALL expose engine state as read-only signals derived from the engine snapshot by a selector. A signal SHALL change only when its selected value changes under its equality function.

#### Scenario: Unrelated change

- **WHEN** a control selects the current page number and the user toggles bold
- **THEN** the page-number signal does not notify its consumers

#### Scenario: Related change

- **WHEN** a control selects the bold state and the user toggles bold
- **THEN** the signal holds the new bold state after the change commits

### Requirement: One engine subscription per root

All state signals under one root SHALL share one subscription to the engine's change, selection, and error events. Notifications SHALL be coalesced and delivered after the engine publishes its layout.

#### Scenario: Many consumers

- **WHEN** fifty controls read state under one root
- **THEN** the engine has one change listener for that root

#### Scenario: State reads after layout

- **WHEN** an edit changes the page count
- **THEN** a page-count signal reads the new count, not the count before layout

### Requirement: Commands report whether they can run

Each command helper SHALL return its enabled state, its active state where it has one, a localized reason when it is disabled, and a run function. Enabled state SHALL come only from the engine's toolbar command state.

#### Scenario: Disabled command

- **WHEN** the editor is in view mode and a control reads the bold command
- **THEN** the command is disabled, its reason is the localized engine reason, and running it changes nothing

### Requirement: Value commands carry their current value

Each value command helper SHALL expose the current value from the snapshot and a function that applies a new value through the engine.

#### Scenario: Font size

- **WHEN** the caret is in 11-point text and a control sets the font size value to 14
- **THEN** the engine applies 14 points and the value signal reads 14

### Requirement: Engine events reach Angular code

The adapter SHALL let Angular code subscribe to engine events with automatic removal when the subscribing component or injector is destroyed.

#### Scenario: Subscriber destroyed

- **WHEN** a component that listens to the `change` event is destroyed
- **THEN** the listener is removed from the engine

### Requirement: Feature state services cover the engine features

The adapter SHALL provide state and actions for zoom, page setup, paragraph format and indent, fonts, document source, header and footer, notes, hyperlinks, content controls, bookmarks, watermarks, highlights, caret, search, outline, go-to, review authors, history grouping, and comment drafts.

#### Scenario: Zoom

- **WHEN** a control calls the zoom-in action at 100 percent
- **THEN** the zoom signal reads the next zoom level and the pages repaint at that level

#### Scenario: Search

- **WHEN** a control searches for a word that occurs three times
- **THEN** the search state reports three matches and the engine highlights them

### Requirement: State outside a root is empty and safe

A state signal or command helper used outside an editor root SHALL return the engine's typed empty values and SHALL NOT throw.

#### Scenario: Toolbar without a root

- **WHEN** a toolbar renders with no editor root above it
- **THEN** every control renders disabled and no error is thrown
