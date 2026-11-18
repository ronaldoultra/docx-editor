import { isPlatformBrowser } from '@angular/common';
import {
  ChangeDetectionStrategy,
  Component,
  DestroyRef,
  ElementRef,
  NgZone,
  PLATFORM_ID,
  ViewEncapsulation,
  effect,
  inject,
  input,
  output,
  untracked,
  viewChild,
} from '@angular/core';
import type { DocumentChange } from '@docx-editor.dev/core/contracts/editor';
import { createDocxEditor, type DocxEditorInstance } from '@docx-editor.dev/core/editor';
import en from '@docx-editor.dev/i18n/en';

/** The DOCX bytes that `<docx-editor>` accepts. */
export type DocxEditorDocument = ArrayBuffer | Uint8Array;

/**
 * The host editing mode. `'edit'` allows text edits. `'view'` shows the pages read-only.
 * The engine also has a `'suggesting'` mode, which needs the review module, so this
 * component does not offer it.
 */
export type DocxEditorMode = 'edit' | 'view';

/**
 * A paginated DOCX editor. The engine in `@docx-editor.dev/core` reads, lays out, paints,
 * and edits the document. This component only owns the DOM element and the lifetime of
 * the engine instance.
 *
 * The engine runs outside the Angular zone when zone.js is present. Only the `change`
 * output enters the zone again, so typing does not start change detection by itself.
 *
 * @example
 * ```html
 * <docx-editor [document]="bytes()" mode="edit" (change)="dirty.set(true)" />
 * ```
 */
@Component({
  selector: 'docx-editor',
  changeDetection: ChangeDetectionStrategy.OnPush,
  encapsulation: ViewEncapsulation.None,
  template: `
    <div
      class="docx-editor docx-editor-one-surface docx-editor-one-surface__viewport docx-editor__scroll-container"
      data-testid="docx-editor-scroll"
      tabindex="0"
      role="region"
      [attr.aria-label]="viewportLabel()"
    >
      <div class="docx-content-mount">
        <div #surface class="docx-paginated-surface"></div>
      </div>
    </div>
  `,
})
export class DocxEditorComponent {
  /**
   * The DOCX bytes to show. A new value loads the new document into the same engine
   * instance with `load()`. The engine does not copy the bytes, so do not change them
   * after you pass them. `null` releases the current document.
   */
  readonly document = input<DocxEditorDocument | null | undefined>(undefined);

  /** The host editing mode. A change applies to the open document without a reload. */
  readonly mode = input<DocxEditorMode>('edit');

  /** The accessible name of the scrolling page region. */
  readonly viewportLabel = input<string>(en.editor.documentViewport);

  /**
   * Emits for each change that the engine commits: an edit, undo, redo, or a load.
   * This is the engine's `change` event without changes to its payload.
   */
  readonly change = output<DocumentChange>();

  private readonly surface = viewChild<ElementRef<HTMLDivElement>>('surface');
  private readonly zone = inject(NgZone);
  private readonly isBrowser = isPlatformBrowser(inject(PLATFORM_ID));

  private editor: DocxEditorInstance | null = null;
  private loaded: DocxEditorDocument | null = null;
  private appliedMode: DocxEditorMode | null = null;
  private readonly releases: Array<() => void> = [];

  constructor() {
    effect(() => {
      const host = this.surface()?.nativeElement;
      const source = this.document() ?? null;
      const mode = this.mode();
      if (!host || !this.isBrowser) return;
      untracked(() => this.zone.runOutsideAngular(() => this.sync(host, source, mode)));
    });
    inject(DestroyRef).onDestroy(() => this.release());
  }

  /**
   * Serializes the current document to DOCX bytes.
   *
   * Rejects when no document is loaded, or when the engine refuses the save.
   */
  save(): Promise<ArrayBuffer> {
    const editor = this.editor;
    if (!editor) return Promise.reject(new Error('docx-editor: no document is loaded.'));
    return this.zone.runOutsideAngular(() => editor.save());
  }

  private sync(host: HTMLDivElement, source: DocxEditorDocument | null, mode: DocxEditorMode) {
    if (source === null) {
      this.release();
      return;
    }
    const editor = this.editor;
    if (!editor) {
      this.create(host, source, mode);
      return;
    }
    if (source !== this.loaded) {
      this.loaded = source;
      editor.load(source);
    }
    if (mode !== this.appliedMode) {
      this.appliedMode = mode;
      editor.setMode(mode);
    }
  }

  private create(host: HTMLDivElement, source: DocxEditorDocument, mode: DocxEditorMode) {
    const editor = createDocxEditor({ document: source, mode });
    this.editor = editor;
    this.loaded = source;
    this.appliedMode = mode;
    this.releases.push(
      editor.on('change', (change) => this.zone.run(() => this.change.emit(change)))
    );
    // A file dropped on the pages must not make the browser open it in place of the app.
    const preventDrop = (event: DragEvent) => event.preventDefault();
    host.addEventListener('drop', preventDrop);
    this.releases.push(() => host.removeEventListener('drop', preventDrop));
    try {
      editor.attach(host);
    } catch (error) {
      // A document that fails to open is a refusal on the snapshot, not a crash of the
      // host. Other errors are defects, so they continue to the caller.
      if (editor.snapshot().openFailure?.code !== 'open.failed') throw error;
      console.error('[docx-editor] The document failed to open.', error);
    }
  }

  private release() {
    for (const off of this.releases.splice(0)) off();
    const editor = this.editor;
    this.editor = null;
    this.loaded = null;
    this.appliedMode = null;
    if (editor) this.zone.runOutsideAngular(() => editor.destroy());
  }
}
