import { effect, untracked } from '@angular/core';
import { injectDocxEditor } from '../context';
import { injectEditorState } from '../editor-state';

/**
 * Tie an open dialog to the document that it opened on. The dialog closes when the engine
 * mounts another document. The function that this returns is a guard for Apply: it returns
 * `true` while the same document is mounted, and closes the dialog and returns `false` when it
 * is not, also before the next render. Call it in an injection context.
 *
 * @internal
 */
export function injectDialogGeneration(open: () => boolean, close: () => void): () => boolean {
  const editor = injectDocxEditor();
  const generation = injectEditorState(() => editor()?.mountGeneration ?? 0);
  let captured: number | null = null;
  effect(() => {
    const isOpen = open();
    const current = generation();
    untracked(() => {
      if (!isOpen) {
        captured = null;
        return;
      }
      if (captured !== null && captured !== current) {
        close();
        return;
      }
      if (editor()?.surface) captured = current;
    });
  });
  return () => {
    const instance = untracked(editor);
    if (captured !== null && instance?.surface && captured === instance.mountGeneration) {
      return true;
    }
    close();
    return false;
  };
}
