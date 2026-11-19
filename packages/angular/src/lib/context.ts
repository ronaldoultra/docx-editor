import {
  Injectable,
  InjectionToken,
  NgZone,
  computed,
  inject,
  signal,
  type Signal,
} from '@angular/core';
import type { DocxEditorInstance } from '@docx-editor.dev/core/editor';
import { deferredTick } from './deferred-notifier';

/**
 * The engine instance of one editor root, and the state that its controls share.
 *
 * `docxEditorRoot` provides one `DocxEditorContext` for its content. Each control under the root
 * reads the same engine instance from it. Inject it to write a control of your own, or use
 * `injectEditorState`, `injectEditorCommand`, and the other `inject…` functions, which read it
 * for you.
 *
 * @public
 */
@Injectable()
export class DocxEditorContext {
  private readonly zone = inject(NgZone);
  private readonly editorSignal = signal<DocxEditorInstance | null>(null);
  private readonly revisionSignal = signal(0);
  private readonly contentElementSignal = signal<HTMLElement | null>(null);
  private readonly darkSource = signal<() => boolean>(() => false);
  private engineListeners = 0;

  /** The engine instance, or `null` before the root creates it and after it destroys it. */
  readonly editor: Signal<DocxEditorInstance | null> = this.editorSignal.asReadonly();

  /**
   * A counter that goes up after each change, selection change, or error that the engine
   * publishes. Several engine events in one task give one increment, after the engine publishes
   * its layout. Read it in a `computed` to read the snapshot again after each change.
   */
  readonly revision: Signal<number> = this.revisionSignal.asReadonly();

  /** The element that the engine paints pages into, while a content component holds one. */
  readonly contentElement: Signal<HTMLElement | null> = this.contentElementSignal.asReadonly();

  /** True when the root's `colorMode` resolves to dark. */
  readonly dark: Signal<boolean> = computed(() => this.darkSource()());

  /**
   * The number of listeners that the root holds on the engine. All state signals under the root
   * share these listeners, so the number does not grow with the number of controls.
   *
   * @internal
   */
  get engineListenerCount(): number {
    return this.engineListeners;
  }

  /**
   * Connect an engine instance, or disconnect with `null`. Returns a function that removes the
   * listeners. Only the root calls it.
   *
   * @internal
   */
  connect(editor: DocxEditorInstance | null): () => void {
    this.editorSignal.set(editor);
    if (!editor) return () => {};
    const notify = deferredTick(() =>
      this.zone.run(() => this.revisionSignal.update((n) => n + 1))
    );
    const offs = [
      editor.on('change', notify),
      editor.on('selectionChange', notify),
      editor.on('error', notify),
      editor.on('revisionMarkupChange', notify),
    ];
    this.engineListeners += offs.length;
    return () => {
      this.engineListeners -= offs.length;
      for (const off of offs.splice(0)) off();
    };
  }

  /** @internal */
  setContentElement(element: HTMLElement | null): void {
    this.contentElementSignal.set(element);
  }

  /** Read the dark state from `source`, which can read signals. @internal */
  useDark(source: () => boolean): void {
    this.darkSource.set(source);
  }
}

/** The context, or `null` outside an editor root. @internal */
export function injectDocxEditorContext(): DocxEditorContext | null {
  return inject(DocxEditorContext, { optional: true });
}

const NO_EDITOR: Signal<DocxEditorInstance | null> = signal(null).asReadonly();

/**
 * The engine instance of the enclosing editor root, as a signal. Outside a root, or before the
 * root creates the instance, the signal holds `null`. Call it in an injection context.
 *
 * @public
 */
export function injectDocxEditor(): Signal<DocxEditorInstance | null> {
  return injectDocxEditorContext()?.editor ?? NO_EDITOR;
}

/**
 * True for the content of an element that already carries the `docx-editor` scope class, such
 * as the viewport. Parts below it do not add the scope class again.
 *
 * @internal
 */
export const DOCX_EDITOR_SCOPED = new InjectionToken<boolean>('DOCX_EDITOR_SCOPED', {
  factory: () => false,
});
