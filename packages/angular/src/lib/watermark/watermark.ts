import { computed, untracked, type Signal } from '@angular/core';
import type {
  EditorSnapshot,
  ExecResult,
  Watermark,
  WatermarkState,
} from '@docx-editor.dev/core/contracts/editor';
import { injectDocxEditor } from '../context';
import { injectEditorCommand } from '../editor-command';
import { injectEditorState } from '../editor-state';
import { watermarkStateEqual } from './watermark-dialog-model';

/**
 * What {@link injectWatermark} returns.
 *
 * @public
 */
export interface WatermarkCommandState {
  /** The document's watermark, or `null` when it has none. The same object while unchanged. */
  readonly watermark: Signal<WatermarkState | null>;
  /** True when the engine can write a watermark now. */
  readonly isEnabled: Signal<boolean>;
  /** The full current watermark, with the picture bytes. */
  read(): Watermark | null;
  /** Set the watermark as one undo step. `null` removes it. */
  apply(watermark: Watermark | null): ExecResult;
}

const selectEditable = (snapshot: EditorSnapshot): boolean => snapshot.editable;
const NOT_READY: ExecResult = { ok: false, code: 'unsupported', reason: 'editor is not ready' };

/**
 * The watermark of the enclosing editor root, and the command that changes it. Call it in an
 * injection context.
 *
 * @public
 */
export function injectWatermark(): WatermarkCommandState {
  const editor = injectDocxEditor();
  const watermark = injectEditorState(
    () => untracked(() => editor()?.getWatermark() ?? null),
    watermarkStateEqual
  );
  const editable = injectEditorState(selectEditable);
  // A short text is always a valid payload, so the answer follows only the mount and the mode.
  const command = injectEditorCommand({ type: 'setWatermark', watermark: { text: 'DRAFT' } });
  return {
    watermark,
    isEnabled: computed(() => editable() && command.isEnabled()),
    read: () => untracked(editor)?.query({ type: 'watermark' }) ?? null,
    apply: (next) => {
      const instance = untracked(editor);
      return instance ? instance.exec({ type: 'setWatermark', watermark: next }) : NOT_READY;
    },
  };
}
