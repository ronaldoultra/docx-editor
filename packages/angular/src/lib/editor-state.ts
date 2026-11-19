import { NgZone, computed, effect, inject, signal, untracked, type Signal } from '@angular/core';
import type { EditorEvents, EditorSnapshot } from '@docx-editor.dev/core/contracts/editor';
import { LOADING_SNAPSHOT, editorHostDocument } from '@docx-editor.dev/core/editor';
import { injectDocxEditorContext } from './context';

/**
 * A read-only signal of one value from the engine snapshot.
 *
 * The signal reads the snapshot again after each change that the engine publishes, and after the
 * engine publishes its layout. It notifies its consumers only when `isEqual` says that the
 * selected value changed. All signals under one root share one set of engine listeners.
 * Outside a root, the selector reads the engine's empty loading snapshot. Call it in an
 * injection context.
 *
 * @example
 * ```ts
 * readonly page = injectEditorState(
 *   (snapshot) => snapshot.page,
 *   (a, b) => a.current === b.current && a.total === b.total
 * );
 * ```
 *
 * @public
 */
export function injectEditorState<T>(
  selector: (snapshot: EditorSnapshot) => T,
  isEqual: (a: T, b: T) => boolean = Object.is
): Signal<T> {
  const context = injectDocxEditorContext();
  return computed(
    () => {
      context?.revision();
      const editor = context?.editor() ?? null;
      const snapshot = editor ? untracked(() => editor.snapshot()) : LOADING_SNAPSHOT;
      return selector(snapshot);
    },
    { equal: isEqual }
  );
}

/**
 * Listen to an engine event while the calling component or service exists. The listener moves
 * to a new engine when the root creates one, and leaves the engine when Angular destroys the
 * caller. The handler runs in the Angular zone. Call it in an injection context.
 *
 * @example
 * ```ts
 * injectEditorEvent('change', (change) => this.saved.set(false));
 * ```
 *
 * @public
 */
export function injectEditorEvent<E extends keyof EditorEvents>(
  event: E,
  handler: EditorEvents[E]
): void {
  const context = injectDocxEditorContext();
  const zone = inject(NgZone);
  if (!context) return;
  effect((onCleanup) => {
    const editor = context.editor();
    if (!editor) return;
    const forward = ((...args: unknown[]) =>
      zone.run(() => (handler as (...forwarded: unknown[]) => void)(...args))) as EditorEvents[E];
    onCleanup(editor.on(event, forward));
  });
}

/**
 * The document that the editor's pages are in: the document of a pop-out window or an iframe
 * when the editor moved there, otherwise the page's own. `null` on the server.
 *
 * Use it in place of the global `document` for listeners, focus reads, and injected styles.
 * The signal changes when the editor moves to another window. Its `defaultView` is the editor's
 * window. Call it in an injection context.
 *
 * @public
 */
export function injectEditorHostDocument(): Signal<Document | null> {
  const context = injectDocxEditorContext();
  // Any window will do before an editor exists, so the global document is correct here.
  const fallback = typeof globalThis.document === 'undefined' ? null : globalThis.document;
  const current = signal<Document | null>(fallback);
  if (!context) return current.asReadonly();
  effect((onCleanup) => {
    const editor = context.editor();
    if (!editor) {
      current.set(fallback);
      return;
    }
    const read = () => current.set(editorHostDocument(editor));
    read();
    onCleanup(editor.on('hostChange', read));
  });
  return current.asReadonly();
}
