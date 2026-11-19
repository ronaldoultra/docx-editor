import {
  ChangeDetectionStrategy,
  Component,
  NgZone,
  ViewEncapsulation,
  computed,
  inject,
  input,
  output,
  viewChild,
  type InputSignal,
  type Signal,
} from '@angular/core';
import type {
  DocumentChange,
  DocumentHandle,
  DocumentOpenFailure,
  DocumentSource,
  Editor,
  EditorCommand,
  EditorError,
  EditorExecOptions,
  EditorFontError,
  EditorScope,
  EditorSnapshot,
  ExecResult,
  ZoomMode,
} from '@docx-editor.dev/core/contracts/editor';
import {
  LOADING_SNAPSHOT,
  type ContentLockPolicy,
  type DocxEditorInstance,
  type EditorModule,
  type FieldResultsMode,
  type FontShapingOptions,
  type ImageDecodePort,
  type LockedControlNavigation,
} from '@docx-editor.dev/core/editor';
import type { Translations } from '@docx-editor.dev/i18n';
import type { DocxEditorColorMode } from './color-mode';
import { DocxEditorContentComponent } from './content.component';
import { DOCX_EDITOR_SCOPED } from './context';
import { DocxEditorI18n, injectChromeDir, type DocxEditorDirection } from './i18n';
import {
  DocxEditorRootDirective,
  type DocxEditorFonts,
  type DocxEditorRootMode,
  type DocxEditorTranslate,
} from './root.directive';
import { DocxEditorViewportComponent } from './viewport.component';

/**
 * The DOCX bytes that `<docx-editor>` accepts.
 *
 * @public
 */
export type DocxEditorDocument = ArrayBuffer | Uint8Array;

/**
 * The host editing mode. `'edit'` allows edits. `'view'` shows the pages read-only.
 * `'suggesting'` records edits as tracked changes by the `author`.
 *
 * @public
 */
export type DocxEditorMode = DocxEditorRootMode;

/**
 * A complete paginated DOCX editor. It composes the public parts — `docxEditorRoot`,
 * `<docx-editor-viewport>`, and `<docx-editor-content>` — so a host that composes its own
 * editor reaches the same behavior. The engine in `@docx-editor.dev/core` reads, lays out,
 * paints, and edits the document.
 *
 * The engine runs outside the Angular zone when zone.js is present. Only the outputs and the
 * state that templates read enter the zone again.
 *
 * @example
 * ```html
 * <docx-editor [document]="bytes()" mode="edit" [zoom]="1.5" (change)="dirty.set(true)" />
 * ```
 *
 * @public
 */
@Component({
  selector: 'docx-editor',
  changeDetection: ChangeDetectionStrategy.OnPush,
  encapsulation: ViewEncapsulation.None,
  imports: [DocxEditorRootDirective, DocxEditorViewportComponent, DocxEditorContentComponent],
  providers: [DocxEditorI18n],
  // The wrapper below carries the scope class, so the parts inside do not add it again.
  viewProviders: [{ provide: DOCX_EDITOR_SCOPED, useValue: true }],
  styles: `
    docx-editor {
      display: block;
    }
  `,
  template: `
    <div
      class="docx-editor"
      data-testid="docx-editor"
      [class.dark]="dark()"
      [attr.dir]="chromeDir() ?? null"
      style="display: flex; flex-direction: column; height: 100%; width: 100%; background-color: var(--doc-bg)"
    >
      @if (document() !== null) {
        <ng-container
          docxEditorRoot
          [document]="document() ?? undefined"
          [mode]="mode()"
          [colorMode]="colorMode()"
          [locale]="locale()"
          [translate]="translate()"
          [author]="author()"
          [zoom]="zoom()"
          [zoomMode]="zoomMode()"
          [fieldResults]="fieldResults()"
          [contentLockPolicy]="contentLockPolicy()"
          [lockedControlNavigation]="lockedControlNavigation()"
          [fonts]="fonts()"
          [fontShaping]="fontShaping()"
          [modules]="modules()"
          [imageDecodePort]="imageDecodePort()"
          (ready)="ready.emit($event)"
          (change)="change.emit($event)"
          (selectionChange)="selectionChange.emit($event)"
          (error)="error.emit($event)"
          (openFailure)="openFailure.emit($event)"
          (fontError)="fontError.emit($event)"
        >
          <docx-editor-viewport style="flex: 1; min-height: 0; min-width: 0; overflow-anchor: none">
            <docx-editor-content />
          </docx-editor-viewport>
        </ng-container>
      }
    </div>
  `,
})
export class DocxEditorComponent {
  /**
   * The DOCX bytes to show. A new value opens the new document in the same engine with
   * `load()`. The engine does not copy the bytes, so do not change them after you pass them.
   * `null` destroys the engine and its document.
   */
  readonly document = input<DocxEditorDocument | null | undefined>(undefined);

  /** The host editing mode. A change applies to the open document without a reload. */
  readonly mode = input<DocxEditorMode>('edit');

  /** The color theme of the chrome. `'system'` follows the operating system. */
  readonly colorMode = input<DocxEditorColorMode>('light');

  /**
   * The translation catalog of the chrome. Missing keys fall back to an outer
   * `docxEditorI18n` provider, then to English. A right-to-left catalog mirrors the chrome.
   */
  readonly i18n: InputSignal<Translations | undefined> = input<Translations | undefined>(undefined);

  /** Labels that the engine paints. Defaults to the active translations. */
  readonly translate = input<DocxEditorTranslate | undefined>(undefined);

  /** The BCP-47 locale for date input and labels that the engine makes. Defaults to `en-US`. */
  readonly locale = input<string | undefined>(undefined);

  /** The author of later comments, replies, and tracked changes. */
  readonly author = input<string | undefined>(undefined);

  /** The zoom factor, where `1` is 100 percent. A change applies without a reload. */
  readonly zoom = input<number | undefined>(undefined);

  /** A fixed zoom, or a zoom that fits the page or its width to the viewport. */
  readonly zoomMode = input<ZoomMode | 'auto' | undefined>(undefined);

  /** How the reader edits saved field results. Read once, when the engine is created. */
  readonly fieldResults = input<FieldResultsMode | undefined>(undefined);

  /** Relaxes content-control locks. Read once, when the engine is created. */
  readonly contentLockPolicy = input<ContentLockPolicy | undefined>(undefined);

  /** How the caret moves through a locked inline control. Read once, at creation. */
  readonly lockedControlNavigation = input<LockedControlNavigation | undefined>(undefined);

  /** Font sources. A new value replaces the fonts in place, so edits and history stay. */
  readonly fonts = input<DocxEditorFonts | undefined>(undefined);

  /** Font shaping limits. A changed `maxFontBytes` creates a new engine. */
  readonly fontShaping = input<FontShapingOptions | undefined>(undefined);

  /** Engine modules, such as the review module. A new array creates a new engine. */
  readonly modules = input<readonly EditorModule[] | undefined>(undefined);

  /** The decoder for raster images. A new port creates a new engine. */
  readonly imageDecodePort = input<ImageDecodePort | undefined>(undefined);

  /**
   * Emits for each change that the engine commits: an edit, undo, redo, or a load.
   * This is the engine's `change` event without changes to its payload.
   */
  readonly change = output<DocumentChange>();

  /** Emits once for each engine, when it finished opening its first document. */
  readonly ready = output<Editor>();

  /** Emits the snapshot after each selection change that the engine publishes. */
  readonly selectionChange = output<EditorSnapshot>();

  /** Emits the errors that the engine publishes, such as a refused command or document. */
  readonly error = output<EditorError>();

  /** Emits when the engine refuses to open a document. */
  readonly openFailure = output<DocumentOpenFailure>();

  /** Emits when a font fails to load. The engine keeps measuring with its fallback. */
  readonly fontError = output<EditorFontError>();

  private readonly zone = inject(NgZone);
  private readonly root = viewChild(DocxEditorRootDirective);

  /** @internal */
  protected readonly chromeDir: Signal<DocxEditorDirection | undefined> = injectChromeDir();
  /** @internal */
  protected readonly dark = computed(() => this.root()?.context.dark() ?? false);

  constructor() {
    inject(DocxEditorI18n, { self: true }).use(() => this.i18n());
  }

  /** The engine instance, or `null` before it exists and after `document` is `null`. */
  getEditor(): DocxEditorInstance | null {
    return this.root()?.context.editor() ?? null;
  }

  /**
   * Open another document in the same engine. The `document` input does not change, so a later
   * change of the input opens that value. Does nothing before the engine exists.
   */
  load(document: DocumentSource): void {
    const editor = this.getEditor();
    if (editor) this.zone.runOutsideAngular(() => editor.load(document));
  }

  /**
   * Serializes the current document to DOCX bytes.
   *
   * Rejects when no document is loaded, or when the engine refuses the save, for example
   * because a form field holds a value that its format does not accept.
   */
  save(): Promise<ArrayBuffer> {
    const editor = this.getEditor();
    if (!editor) return Promise.reject(new Error('docx-editor: no document is loaded.'));
    return this.zone.runOutsideAngular(() => editor.save());
  }

  /** Move the keyboard focus to the pages. */
  focus(): void {
    const editor = this.getEditor();
    if (editor) this.zone.runOutsideAngular(() => editor.focus());
  }

  /** Run an engine command. Before the engine exists, the result is a `notFound` refusal. */
  exec(command: EditorCommand, options?: EditorExecOptions): ExecResult {
    const editor = this.getEditor();
    if (!editor) return { ok: false, code: 'notFound', reason: 'no editor is mounted' };
    return this.zone.runOutsideAngular(() => editor.exec(command, options));
  }

  /** The engine snapshot. Before the engine exists, the empty loading snapshot. */
  snapshot(options?: { scope?: EditorScope }): EditorSnapshot {
    return this.getEditor()?.snapshot(options) ?? LOADING_SNAPSHOT;
  }

  /** A handle on the open document, or `null` before the engine exists. */
  getDocumentHandle(): DocumentHandle | null {
    return this.getEditor()?.getDocumentHandle() ?? null;
  }
}
