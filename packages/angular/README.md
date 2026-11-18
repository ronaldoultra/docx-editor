# @docx-editor.dev/angular

An Angular component for the `@docx-editor.dev/core` DOCX engine. The engine reads the file, lays out and paints the pages, handles text input, and writes the file again. The component owns only the DOM element and the engine lifetime.

This package is a minimal adapter. It has no toolbar, menus, context menu, or dialogs.

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

## API

| Member | Kind | Description |
| --- | --- | --- |
| `document` | input | `ArrayBuffer` or `Uint8Array` with the DOCX bytes. A new value loads the new document into the same engine instance. `null` releases the document. |
| `mode` | input | `'edit'` (default) or `'view'`. A change applies without a reload. |
| `viewportLabel` | input | The accessible name of the page region. |
| `change` | output | The engine `change` event: an edit, undo, redo, or load. |
| `save()` | method | Returns `Promise<ArrayBuffer>` with the DOCX bytes. |

The engine has no `ready` event, so the component has no `ready` output.

## Zone.js

The engine runs outside the Angular zone. Only the `change` output enters the zone again, so typing starts change detection only when your app listens to `change`. Without zone.js, the component uses signals and needs no zone.

## Build

```bash
bun run --filter '@docx-editor.dev/angular' build
```

The build uses `ng-packagr` and writes the Angular Package Format to `dist/`.
