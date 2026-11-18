# Spec Delta

## Purpose

Defines how the Angular adapter creates one engine instance per editor, shares it with every control in that editor, mounts its pages, and releases it, so that hosts can compose their own editor layout.

## ADDED Requirements

### Requirement: One engine instance per editor root

An editor root SHALL create exactly one engine instance and provide it to every descendant component and service through Angular dependency injection. Two roots on one page SHALL have two independent instances.

#### Scenario: Descendants share the root's instance

- **WHEN** a toolbar and a navigation pane render inside one editor root
- **THEN** both read and command the same engine instance

#### Scenario: Two roots stay independent

- **WHEN** a page renders two editor roots and the user types in the first
- **THEN** the second root's snapshot does not change

### Requirement: Content mount attaches the painted pages

The content component SHALL attach the root's engine instance to its own element and detach it when the element leaves the view. The painted pages SHALL be the editable surface.

#### Scenario: Pages paint inside the content element

- **WHEN** a root with a document and a content component renders
- **THEN** the painted pages appear inside the content element and accept typing

#### Scenario: Content leaves the view

- **WHEN** the content component is destroyed while the root stays
- **THEN** the engine detaches and keeps its undo history, selection, and document

### Requirement: Viewport carries the scroll container

The viewport component SHALL render the scroll container that the engine stylesheet and geometry expect, with an accessible region name from the active translations.

#### Scenario: Viewport renders the scroller

- **WHEN** a root renders a viewport and content
- **THEN** the content sits inside an element with the `docx-editor__scroll-container` class and `role="region"`

### Requirement: Engine work runs outside the Angular zone

Engine creation, input handling, layout, and paint SHALL run outside the Angular zone. Only adapter outputs and state that templates read SHALL enter the zone. The adapter SHALL work with and without zone.js.

#### Scenario: Typing does not start change detection

- **WHEN** the user types one character in an app that uses zone.js
- **THEN** Angular change detection runs only for the state changes that the typed character publishes, not for each DOM event

#### Scenario: Zoneless application

- **WHEN** the adapter runs in an application with zoneless change detection
- **THEN** controls update after each committed change

### Requirement: Root releases the engine on destroy

When an editor root is destroyed, it SHALL remove its engine listeners and destroy its engine instance.

#### Scenario: Root destroyed

- **WHEN** the component that contains an editor root is destroyed
- **THEN** the engine instance is destroyed and no listener of it stays on the document or window

### Requirement: Editor moves between windows

The adapter SHALL keep a live editor when its content element moves to another window or iframe. Controls SHALL resolve the document and window from an element the editor owns, at use time.

#### Scenario: Move to a pop-out window

- **WHEN** a host moves the editor into a pop-out window
- **THEN** typing, selection, and the open popups work in the pop-out window, and undo keeps the earlier edits

### Requirement: Server rendering is inert

On a non-browser platform, the adapter SHALL render its static structure and SHALL NOT create an engine instance.

#### Scenario: Server render

- **WHEN** an application renders the editor on the server
- **THEN** the output contains the viewport structure and no engine instance is created
