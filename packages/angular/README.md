# @docx-editor.dev/angular

An Angular component for the `@docx-editor.dev/core` DOCX engine. The engine reads the file, lays out and paints the pages, handles text input, and writes the file again. The component owns only the DOM element and the engine lifetime.

The package has two levels:

- `<docx-editor>` is a packaged component for a page that needs only the pages.
- The composition API is a set of parts that you put together in your own layout. Use it to add your own controls.

The package has a formatting toolbar. It has no menus, context menu, or dialogs yet. `<docx-editor>` shows the toolbar above the pages, in a `.docx-editor` wrapper that carries the color mode and the chrome direction.

## Requirements

- Angular 19 or 20, with or without zone.js.
- `@docx-editor.dev/core` 2.39.

## Install

```bash
npm install @docx-editor.dev/angular @docx-editor.dev/core
```

## Add the styles

Add the package stylesheet to the `styles` array in `angular.json`, or import it from your global stylesheet:

```css
@import '@docx-editor.dev/angular/styles.css';
```

The stylesheet imports the engine stylesheet from `@docx-editor.dev/core`. Give `<docx-editor>` a height, because its pages scroll inside it.

## Serve the text shaper

The engine measures text with HarfBuzz, which it loads from `harfbuzz.wasm`. The Angular builder does not copy that file. Copy it with an `assets` entry in `angular.json`:

```json
{
  "glob": "harfbuzz.wasm",
  "input": "node_modules/@docx-editor.dev/core/dist",
  "output": "/"
}
```

Then set its URL before the first editor starts, for example in `main.ts`:

```ts
import { setHarfBuzzWasmUrl } from '@docx-editor.dev/core/layout';

setHarfBuzzWasmUrl(new URL('harfbuzz.wasm', document.baseURI));
```

## Use the component

The following component opens a file that the user selects, and downloads the edited file:

```ts
import { Component, signal, viewChild } from '@angular/core';
import { DocxEditorComponent } from '@docx-editor.dev/angular';

@Component({
  selector: 'app-root',
  imports: [DocxEditorComponent],
  template: `
    <input type="file" accept=".docx" (change)="open($event)" />
    <button type="button" (click)="download()">Save</button>
    <docx-editor [document]="bytes()" mode="edit" (change)="dirty.set(true)" />
  `,
})
export class AppComponent {
  protected readonly bytes = signal<ArrayBuffer | null>(null);
  protected readonly dirty = signal(false);
  private readonly editor = viewChild.required(DocxEditorComponent);

  protected async open(event: Event) {
    const file = (event.target as HTMLInputElement).files?.[0];
    if (file) this.bytes.set(await file.arrayBuffer());
  }

  protected async download() {
    const buffer = await this.editor().save();
    // Write `buffer` to a Blob and download it.
  }
}
```

## Compose your own editor

The composition API has three parts:

- `docxEditorRoot` is a directive. It creates one engine and provides it to the elements inside it. It renders no element, so you can put it on an `ng-container`.
- `<docx-editor-viewport>` is the scroll container of the pages. It is a named `region`.
- `<docx-editor-content>` is the element that the engine paints the pages into.

The following component puts a custom control above the pages:

```ts
import { Component, signal } from '@angular/core';
import {
  DocxEditorContentComponent,
  DocxEditorRootDirective,
  DocxEditorViewportComponent,
} from '@docx-editor.dev/angular';
import { BoldButtonComponent } from './bold-button.component';

@Component({
  selector: 'app-editor',
  imports: [
    DocxEditorRootDirective,
    DocxEditorViewportComponent,
    DocxEditorContentComponent,
    BoldButtonComponent,
  ],
  template: `
    <ng-container docxEditorRoot [document]="bytes()" mode="edit" (change)="dirty.set(true)">
      <app-bold-button />
      <docx-editor-viewport style="height: 80vh">
        <docx-editor-content />
      </docx-editor-viewport>
    </ng-container>
  `,
})
export class EditorComponent {
  protected readonly bytes = signal<ArrayBuffer | null>(null);
  protected readonly dirty = signal(false);
}
```

Two roots on one page have two independent engines. When Angular destroys the root, the root destroys its engine. When only `<docx-editor-content>` leaves the view, the engine keeps its document, undo history, and selection. A new content element attaches the engine again.

The root has these inputs and outputs:

| Member | Kind | Description |
| --- | --- | --- |
| `document` | input | The DOCX bytes. A new value opens the new document in the same engine. |
| `mode` | input | `'edit'`, `'view'`, or `'suggesting'`. Omit it to let the document's tracking setting select the mode. |
| `colorMode` | input | `'light'`, `'dark'`, or `'system'` for the chrome. The pages keep the look that Word gives them. |
| `locale` | input | The BCP-47 locale for dates and labels that the engine makes. |
| `translate` | input | A translation function for labels that the engine paints. Defaults to the active translations. |
| `author` | input | The author of later comments, replies, and tracked changes. |
| `ready` | output | Emits once, when the first document is open. |
| `change` | output | Emits for each change that the engine commits. |

### Write a control

Controls read engine state through `inject…` functions. Call them in an injection context, such as a field initializer. Each function returns signals, so a template reads the state directly.

The following control toggles bold. Its enabled state comes from the engine, so it follows the editing mode and the selection:

```ts
import { Component } from '@angular/core';
import { injectEditorCommand } from '@docx-editor.dev/angular';

@Component({
  selector: 'app-bold-button',
  template: `
    <button
      type="button"
      [disabled]="!bold.isEnabled()"
      [attr.aria-pressed]="bold.isActive()"
      [title]="bold.disabledReason() ?? ''"
      (mousedown)="$event.preventDefault()"
      (click)="bold.execute()"
    >
      Bold
    </button>
  `,
})
export class BoldButtonComponent {
  protected readonly bold = injectEditorCommand('text.bold');
}
```

The control prevents the default `mousedown` action, so the press does not move the caret out of the pages.

The state functions are:

| Function | Returns |
| --- | --- |
| `injectEditorState(selector, isEqual?)` | A signal of one value from the engine snapshot. It notifies only when `isEqual` says that the value changed. |
| `injectEditorCommand(target)` | The enabled, active, and disabled-reason signals of a chrome slot or an engine command, and `execute`. |
| `injectEditorValueCommand(slotId)` | The current value and options of a value slot, such as `font.size`, and `execute(value)`. |
| `injectEditorEvent(event, handler)` | Nothing. It listens to an engine event until Angular destroys the caller. |
| `injectHistoryGroup(options)` | A binding that makes a run of changes from one control one undo step. |
| `injectDocxEditor()` | A signal of the engine instance, or `null`. |
| `injectEditorHostDocument()` | A signal of the document that the pages are in. Use it in place of the global `document`. |

All signals under one root share one set of engine listeners. They read the snapshot after the engine publishes its layout. Outside a root, they read the engine's empty values, and commands are disabled.

### Translate the chrome

The chrome uses the English strings from `@docx-editor.dev/i18n` by default. To use another catalog, add `provideDocxEditorI18n()` to the providers of a component:

```ts
import { ptBR } from '@docx-editor.dev/i18n';
import { provideDocxEditorI18n } from '@docx-editor.dev/angular';

@Component({ providers: [provideDocxEditorI18n(ptBR)], … })
```

To change the catalog at run time, use the `docxEditorI18n` directive on an element around the root. Missing keys fall back to the outer provider, then to English. A right-to-left catalog mirrors the chrome. The `docxEditorDir` input sets the direction explicitly:

```html
<section [docxEditorI18n]="catalog()" docxEditorDir="rtl">
  <ng-container docxEditorRoot [document]="bytes()">…</ng-container>
</section>
```

In your own controls, `injectTranslation()` returns the `t`, `lang`, and `dir` signals of the active catalog.

### Match the color mode

The viewport carries the `docx-editor` scope class, and the `dark` class under a dark root. A part that you build outside the viewport reads the same classes from `injectScopeClassName()`. `injectEditorColorMode()` returns the resolved `'light'` or `'dark'` theme.

### Move the editor to another window

`<docx-editor-window>` shows its panel in a pop-out window while `open` is `true`, and in place otherwise. Give the panel as an `ng-template`. The root stays in your window, so the editor keeps its document, undo history, and selection when it moves. When the user closes the window, the panel comes back, and the component emits `openChange` with `false`.

The following template adds a button that moves the pages to a pop-out window:

```html
<ng-container docxEditorRoot [document]="bytes()">
  <button type="button" (click)="popped.set(!popped())">Pop out</button>
  <docx-editor-window [(open)]="popped">
    <ng-template>
      <docx-editor-viewport><docx-editor-content /></docx-editor-viewport>
    </ng-template>
  </docx-editor-window>
</ng-container>
```

## Add the toolbar

`fieldset[docxEditorToolbar]` is the formatting toolbar. Put it inside an editor root. It renders the engine's default groups and controls in registry order, and each control carries its slot id, such as `text.bold`, in `data-slot`. The engine decides whether a control is enabled. A disabled control shows the localized engine reason as its title.

The following template puts the toolbar above the pages:

```html
<ng-container docxEditorRoot [document]="bytes()">
  <fieldset docxEditorToolbar></fieldset>
  <docx-editor-viewport style="height: 80vh">
    <docx-editor-content />
  </docx-editor-viewport>
</ng-container>
```

When the bar has no room, controls leave from the end of the groups in collapse order, into a "⋯" panel. Set `[overflow]="false"` to let the bar wrap instead.

The toolbar has these inputs:

| Input | Description |
| --- | --- |
| `preset` | `true` (default) renders the default arrangement. `false` renders only your content. |
| `overflow` | `true` (default) collapses the controls that do not fit into the "⋯" panel. |
| `t` | A function that returns your own label for a key, or `undefined` to use the active catalog. |

### Change the default controls

Put `ng-template` elements inside the toolbar to change the default arrangement. You do not copy the arrangement.

- `<ng-template docxSlot="text.bold">` replaces one control in its place. `[docxSlotOverflow]` gives the content of its "⋯" panel row.
- `<ng-template docxSlot="text.strike" docxSlotHidden />` removes one control.
- `<ng-template docxToolbarGroup="mine" label="Mine">` adds a group of your own. It is measured with the other groups and collapses into a labeled section of the "⋯" panel. `priority`, `pinned`, and `after` place it.
- `<ng-template docxToolbarGroup="font">` adds content to a built-in group. `hidden` removes the group.

Other content of the toolbar renders after the groups, in a block that does not collapse.

The following toolbar replaces the bold control, removes strikethrough, and adds a group with a host action:

```html
<fieldset docxEditorToolbar>
  <ng-template docxSlot="text.bold">
    <button docxToolbarButton="text.bold" class="my-bold"></button>
  </ng-template>
  <ng-template docxSlot="text.strike" docxSlotHidden />
  <ng-template docxToolbarGroup="review-nav" label="Review" [priority]="15">
    <button docxToolbarAction label="Next change" (select)="next()">
      <svg [docxChromeIcon]="nextIcon"></svg>
    </button>
  </ng-template>
</fieldset>
```

In the "⋯" panel, a `docxToolbarAction` button renders as a panel row with its icon and label.

### Toolbar parts

Each part is a standalone component that you can also place yourself, for example in a toolbar with `[preset]="false"`:

| Part | Slot |
| --- | --- |
| `button[docxToolbarButton]` | Any command slot, such as `history.undo`, `text.bold`, `list.bullet`, `text.link`, or `direction.rtl`. |
| `button[docxToolbarAction]` | Your own action, with `label` and a `select` output. |
| `div[docxToolbarSeparator]` | A separator. |
| `div[docxToolbarFontFamily]` | `font.family` |
| `span[docxToolbarFontSize]` | `font.size` |
| `div[docxToolbarParagraphStyle]` | `styles.style` |
| `span[docxToolbarZoom]` | `zoom.level` |
| `div[docxToolbarFontColor]`, `div[docxToolbarHighlight]` | `text.color`, `text.highlight` |
| `div[docxToolbarAlignment]` | The alignment dropdown. |
| `span[docxToolbarLineSpacing]` | `list.lineSpacing` |
| `div[docxToolbarEditingMode]` | `review.editingMode` |
| `span[docxToolbarTableInsert]` | `table.insert` |
| `div[docxToolbarTableBorderTarget]`, `div[docxToolbarTableBorderStyle]`, `div[docxToolbarTableBorderWidth]`, `div[docxToolbarTableBorderColor]`, `div[docxToolbarTableCellFill]` | The table border and fill slots. They render only while the selection is in a table. |

To build a picker of your own, read its state with `injectFontFamily`, `injectParagraphStyle`, or `injectZoom`, and replace the slot with `docxSlot`.

## Load a document and its fonts

`injectDocxSource` loads a document from bytes or from a URL, and holds it back until its eager fonts settle. An eager font origin is a configuration, a fragment, a promise, or a loader with no argument. A resolver from `defineFontResolver` needs the open document, so the document opens at once and the resolver answers for each document. Call `injectDocxSource` in an injection context:

```ts
import { injectDocxSource } from '@docx-editor.dev/angular';

export class AppComponent {
  readonly url = signal('/files/report.docx');
  readonly source = injectDocxSource(() => this.url(), { fonts: myFonts });
}
```

Bind both values to the editor:

```html
<docx-editor [document]="source.document()" [fonts]="source.fonts()" />
```

`injectFonts` composes several font origins into one resolver. The first origin wins for a face that two origins supply. A signal origin is read again for each document. Any other function is a resolver, so to build an origin lazily, pass a `computed`:

```ts
readonly fonts = injectFonts(brandFonts, computed(() => this.extraFonts()));
```

## Packaged component API

`<docx-editor>` composes the same public parts that the previous sections describe: `docxEditorRoot`, `fieldset[docxEditorToolbar]`, `<docx-editor-viewport>`, and `<docx-editor-content>`. The root inputs have the same names and the same behavior. Set `[chrome]="false"` to show only the pages.

### Inputs that apply live

A change to one of these inputs applies to the open document without a reload. The engine, the edits, and the undo history stay.

| Input | Description |
| --- | --- |
| `document` | `ArrayBuffer` or `Uint8Array` with the DOCX bytes. A new value opens the new document in the same engine. `null` destroys the engine and its document. |
| `chrome` | `true` (default) shows the formatting toolbar above the pages. `false` shows only the pages. |
| `mode` | `'edit'` (default), `'view'`, or `'suggesting'`. Suggesting needs a review module in `modules` and an `author`. Without an author, the editor stays in editing until `author` has a value. |
| `author` | The author of later comments, replies, and tracked changes. |
| `colorMode` | `'light'` (default), `'dark'`, or `'system'`, which follows the operating system. |
| `i18n` | A translation catalog. Missing keys fall back to an outer `docxEditorI18n` provider, then to English. A right-to-left catalog mirrors the chrome. |
| `translate` | A function for the labels that the engine paints. Defaults to the active catalog. |
| `locale` | The BCP-47 locale for date input and the labels that the engine makes. Defaults to `en-US`. |
| `zoom` | The zoom factor, where `1` is 100 percent. |
| `zoomMode` | A fixed zoom, or a zoom that fits the page or its width to the viewport. |
| `fonts` | Font sources: a configuration, fragments, or a resolver. A new value replaces the fonts of the same engine. See `injectFonts` and `injectDocxSource`. |

### Inputs that create a new engine

A change to one of these inputs creates a new engine. The new engine opens the `document` input again, so edits that you did not save are lost.

| Input | Description |
| --- | --- |
| `modules` | Engine modules, such as a review module. Keep one array for the life of the editor. |
| `fontShaping` | Font shaping limits. Only a changed `maxFontBytes` creates a new engine. |
| `imageDecodePort` | The decoder for raster images. Destroying the engine aborts the decodes that still run. |

### Inputs that the engine reads once

The engine reads these inputs when it is created. A later change has no effect, and a development build warns once.

| Input | Description |
| --- | --- |
| `fieldResults` | `'atomic'` (default) keeps each field one unit. `'editable'` allows edits inside the saved result. |
| `contentLockPolicy` | `{ allowFormattingInLockedContent: true }` lets readers format locked content controls. |
| `lockedControlNavigation` | `'characters'` (default) or `'atomic'`: how the caret moves through a locked inline control. |

### Outputs

| Output | Description |
| --- | --- |
| `ready` | Emits once for each engine, when it finished opening its first document. |
| `change` | The engine `change` event: an edit, undo, redo, or load. |
| `selectionChange` | The engine snapshot after each selection change. |
| `error` | The errors that the engine publishes, such as a refused command or document. |
| `openFailure` | The reason when the engine refuses to open a document: damaged, not DOCX, or too large. |
| `fontError` | A font that failed to load. The engine keeps measuring with its fallback. |

### Methods

Get the component with `viewChild(DocxEditorComponent)`, then call these methods:

| Method | Description |
| --- | --- |
| `save()` | Returns `Promise<ArrayBuffer>` with the DOCX bytes. Rejects without a document, and when the engine refuses, for example for a form field value that its format does not accept. |
| `load(document)` | Opens another document in the same engine. The `document` input does not change. |
| `exec(command, options)` | Runs an engine command. Without an engine, returns a `notFound` refusal. |
| `snapshot(options)` | The engine snapshot. Without an engine, the empty loading snapshot. |
| `focus()` | Moves the keyboard focus to the pages. |
| `getEditor()` | The engine instance, or `null`. |
| `getDocumentHandle()` | A handle on the open document, or `null`. |

## Zone.js

The engine runs outside the Angular zone. Only the outputs and the state signals that templates read enter the zone again, so a DOM event on the pages does not start change detection. Without zone.js, the parts use signals and need no zone.

## Build

```bash
bun run --filter '@docx-editor.dev/angular' build
```

The build uses `ng-packagr` and writes the Angular Package Format to `dist/`.
