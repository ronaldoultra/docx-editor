import { isPlatformBrowser } from '@angular/common';
import {
  DestroyRef,
  Directive,
  ElementRef,
  NgZone,
  PLATFORM_ID,
  computed,
  effect,
  inject,
  input,
  output,
  signal,
  untracked,
  type OnInit,
} from '@angular/core';
import type {
  DocumentChange,
  DocumentSource,
  Editor,
  ZoomMode,
} from '@docx-editor.dev/core/contracts/editor';
import {
  createDocxEditor,
  hostWindowOf,
  type ContentLockPolicy,
  type DocxEditorInstance,
  type FieldResultsMode,
  type LockedControlNavigation,
} from '@docx-editor.dev/core/editor';
import { resolveIsDark, subscribeSystemDark, type DocxEditorColorMode } from './color-mode';
import { DocxEditorContext } from './context';
import { injectTranslation } from './i18n';
import {
  applyZoomInputs,
  warnCreationOptionChanged,
  type AppliedZoom,
  type CreationOption,
} from './root-options';

/**
 * The host editing mode. `'edit'` allows edits. `'view'` shows the pages read-only.
 * `'suggesting'` records edits as tracked changes.
 *
 * @public
 */
export type DocxEditorRootMode = 'edit' | 'view' | 'suggesting';

/** A translation function for labels that the engine paints, such as form controls. @public */
export type DocxEditorTranslate = (key: string, params?: Record<string, string | number>) => string;

/**
 * Create one editor engine for the content of this element, and provide it to every part
 * inside: the viewport, the content, and the chrome. Two roots on one page have two independent
 * engines. The root destroys its engine when Angular destroys the root.
 *
 * The root renders no element of its own, so put it on an element or on an `ng-container`.
 * Engine work runs outside the Angular zone. Only the outputs and the state that templates read
 * enter the zone again. On the server, the root creates no engine.
 *
 * @example
 * ```html
 * <ng-container docxEditorRoot [document]="bytes()" (change)="dirty.set(true)">
 *   <docx-editor-viewport>
 *     <docx-editor-content />
 *   </docx-editor-viewport>
 * </ng-container>
 * ```
 *
 * @public
 */
@Directive({
  selector: '[docxEditorRoot]',
  exportAs: 'docxEditorRoot',
  providers: [DocxEditorContext],
})
export class DocxEditorRootDirective implements OnInit {
  /**
   * The document to open. A new value opens the new document in the same engine with
   * `load()`. The engine does not copy bytes, so do not change them after you pass them.
   */
  readonly document = input<DocumentSource | null | undefined>(undefined);

  /** The host editing mode. Omit it to let the document's tracking setting select the mode. */
  readonly mode = input<DocxEditorRootMode | undefined>(undefined);

  /** The color theme of the chrome. A change applies without a new engine. */
  readonly colorMode = input<DocxEditorColorMode | undefined>(undefined);

  /**
   * The BCP-47 locale for date input and labels that the engine makes. Defaults to `en-US`.
   * For chrome translations, use `docxEditorI18n` or `provideDocxEditorI18n()`.
   */
  readonly locale = input<string | undefined>(undefined);

  /** Labels that the engine paints. Defaults to the active translations. */
  readonly translate = input<DocxEditorTranslate | undefined>(undefined);

  /** The author of later comments, replies, and tracked changes. */
  readonly author = input<string | undefined>(undefined);

  /** The zoom factor, where `1` is 100 percent. A change applies without a new engine. */
  readonly zoom = input<number | undefined>(undefined);

  /** A fixed zoom, or a zoom that fits the page or its width to the viewport. */
  readonly zoomMode = input<ZoomMode | 'auto' | undefined>(undefined);

  /**
   * How the reader edits saved field results. `'atomic'` (the default) keeps each field one unit.
   * `'editable'` allows edits inside the saved result. The engine reads it once, when the root
   * creates it. A later change has no effect, and a development build warns once.
   */
  readonly fieldResults = input<FieldResultsMode | undefined>(undefined);

  /**
   * Relaxes content-control locks. With `allowFormattingInLockedContent: true`, readers can
   * format locked controls while their text stays protected. Read once, as `fieldResults` is.
   */
  readonly contentLockPolicy = input<ContentLockPolicy | undefined>(undefined);

  /**
   * How the caret moves through an inline content control whose content is locked.
   * `'characters'` (the default) steps through its characters, as Word does. `'atomic'` makes
   * the control one caret unit. Read once, as `fieldResults` is.
   */
  readonly lockedControlNavigation = input<LockedControlNavigation | undefined>(undefined);

  /** Emits once, when the engine finished opening its first document. */
  readonly ready = output<Editor>();

  /** Emits for each change that the engine commits: an edit, undo, redo, or a load. */
  readonly change = output<DocumentChange>();

  /** The context that this root provides. */
  readonly context = inject(DocxEditorContext, { self: true });

  private readonly zone = inject(NgZone);
  private readonly host = inject(ElementRef<Node>);
  private readonly isBrowser = isPlatformBrowser(inject(PLATFORM_ID));
  private readonly translation = injectTranslation();
  private readonly systemDark = signal(false);
  private readonly releases: Array<() => void> = [];
  private loaded: DocumentSource | null | undefined;
  private appliedMode: DocxEditorRootMode | undefined;
  private readonly appliedZoom: AppliedZoom = { zoom: undefined, mode: undefined };
  private readonly created: Partial<Record<CreationOption, unknown>> = {};

  /** The function that the engine uses for its labels. */
  private readonly translateResolver = computed<DocxEditorTranslate>(() => {
    const custom = this.translate();
    if (custom) return custom;
    const t = this.translation.t();
    return (key, params) => t(key as Parameters<typeof t>[0], params);
  });

  constructor() {
    this.context.useDark(() => resolveIsDark(this.colorMode() ?? 'light', this.systemDark()));
    effect((onCleanup) => {
      if (this.colorMode() !== 'system' || !this.isBrowser) return;
      const view = hostWindowOf(this.host.nativeElement);
      onCleanup(subscribeSystemDark(view, (dark) => this.systemDark.set(dark)));
    });
    this.whenEditor((editor) => {
      const source = this.document();
      if (source === undefined || source === null || source === this.loaded) return;
      this.loaded = source;
      editor.load(source);
    });
    this.whenEditor((editor) => {
      const mode = this.mode();
      if (mode === this.appliedMode) return;
      this.appliedMode = mode;
      editor.setMode(mode);
    });
    this.whenEditor((editor) => editor.setLocale(this.locale()));
    this.whenEditor((editor) => editor.setAuthor(this.author()));
    this.whenEditor((editor) => editor.setTranslate(this.translateResolver()));
    this.whenEditor((editor) =>
      applyZoomInputs(editor, this.appliedZoom, this.zoom(), this.zoomMode())
    );
    for (const option of [
      'fieldResults',
      'contentLockPolicy',
      'lockedControlNavigation',
    ] as const) {
      this.whenEditor(() =>
        warnCreationOptionChanged(option, this.created[option], this[option]())
      );
    }
    inject(DestroyRef).onDestroy(() => this.release());
  }

  ngOnInit(): void {
    if (!this.isBrowser) return;
    this.zone.runOutsideAngular(() => this.create());
  }

  /** Run `apply` outside the zone with the engine, each time it or a signal it reads changes. */
  private whenEditor(apply: (editor: DocxEditorInstance) => void): void {
    effect(() => {
      const editor = this.context.editor();
      if (!editor) return;
      // Read the inputs here, so the effect runs again when they change.
      this.zone.runOutsideAngular(() => apply(editor));
    });
  }

  private create(): void {
    const source = untracked(this.document);
    const mode = untracked(this.mode);
    const locale = untracked(this.locale);
    const author = untracked(this.author);
    const zoom = untracked(this.zoom);
    const zoomMode = untracked(this.zoomMode);
    const fieldResults = untracked(this.fieldResults);
    const contentLockPolicy = untracked(this.contentLockPolicy);
    const lockedControlNavigation = untracked(this.lockedControlNavigation);
    const editor = createDocxEditor({
      ...(source !== undefined && source !== null ? { document: source } : {}),
      ...(mode !== undefined ? { mode } : {}),
      ...(locale !== undefined ? { locale } : {}),
      ...(author !== undefined ? { author } : {}),
      ...(zoom !== undefined ? { zoom } : {}),
      ...(zoomMode !== undefined ? { zoomMode } : {}),
      ...(fieldResults !== undefined ? { fieldResults } : {}),
      ...(contentLockPolicy !== undefined ? { contentLockPolicy } : {}),
      ...(lockedControlNavigation !== undefined ? { lockedControlNavigation } : {}),
      translate: untracked(this.translateResolver),
    });
    this.loaded = source;
    this.appliedMode = mode;
    this.appliedZoom.zoom = zoom;
    this.appliedZoom.mode = zoomMode;
    Object.assign(this.created, { fieldResults, contentLockPolicy, lockedControlNavigation });
    this.releases.push(
      editor.on('change', (change) => this.zone.run(() => this.change.emit(change)))
    );
    this.releases.push(this.context.connect(editor));
    this.emitReadyWhenOpen(editor);
  }

  private emitReadyWhenOpen(editor: DocxEditorInstance): void {
    const fire = () => this.zone.run(() => this.ready.emit(editor));
    if (!editor.snapshot().isOpening) {
      // After the constructor and the first effects, as for a document that opens later.
      queueMicrotask(() => {
        if (this.context.editor() === editor) fire();
      });
      return;
    }
    const off = editor.on('change', () => {
      off();
      fire();
    });
    this.releases.push(off);
  }

  private release(): void {
    for (const off of this.releases.splice(0)) off();
    const editor = untracked(this.context.editor);
    this.context.connect(null);
    if (editor) this.zone.runOutsideAngular(() => editor.destroy());
  }
}
