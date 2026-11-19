import { computed, untracked, type Signal } from '@angular/core';
import type { EditorSnapshot, PageSetup } from '@docx-editor.dev/core/contracts/editor';
import { injectDocxEditor } from '../context';
import { injectEditorCommand } from '../editor-command';
import { injectEditorState } from '../editor-state';

/**
 * A change of the page setup, in twips. An omitted field keeps its value.
 *
 * @public
 */
export interface PageSetupUpdate {
  readonly pageWidthTwips?: number;
  readonly pageHeightTwips?: number;
  readonly orientation?: 'portrait' | 'landscape';
  readonly marginTopTwips?: number;
  readonly marginRightTwips?: number;
  readonly marginBottomTwips?: number;
  readonly marginLeftTwips?: number;
  /** Change the whole document, or only the section at the caret. */
  readonly scope?: 'document' | 'section';
}

/**
 * What {@link injectPageSetup} returns.
 *
 * @public
 */
export interface PageSetupState {
  /** The page setup of the section at the caret, or `null` before a document is open. */
  readonly pageSetup: Signal<PageSetup | null>;
  /** True when the engine can change the page setup now. */
  readonly isEnabled: Signal<boolean>;
  /** Change the page setup as one undo step. Returns `false` when the engine refuses it. */
  apply(update: PageSetupUpdate): boolean;
}

const selectPageSetup = (snapshot: EditorSnapshot): PageSetup | null => snapshot.pageSetup ?? null;
const selectEditable = (snapshot: EditorSnapshot): boolean => snapshot.editable;

/**
 * The page setup of the enclosing editor root, and the command that changes it. Call it in an
 * injection context.
 *
 * @public
 */
export function injectPageSetup(): PageSetupState {
  const editor = injectDocxEditor();
  const pageSetup = injectEditorState(selectPageSetup);
  const editable = injectEditorState(selectEditable);
  const command = injectEditorCommand({ type: 'setPageSetup', marginTop: 0 });
  const isEnabled = computed(() => editable() && command.isEnabled());
  return {
    pageSetup,
    isEnabled,
    apply: (update) => {
      const instance = untracked(editor);
      if (!instance) return false;
      return instance.exec({
        ...(update.pageWidthTwips !== undefined ? { pageWidth: update.pageWidthTwips } : {}),
        ...(update.pageHeightTwips !== undefined ? { pageHeight: update.pageHeightTwips } : {}),
        ...(update.orientation !== undefined ? { orientation: update.orientation } : {}),
        ...(update.marginTopTwips !== undefined ? { marginTop: update.marginTopTwips } : {}),
        ...(update.marginRightTwips !== undefined ? { marginRight: update.marginRightTwips } : {}),
        ...(update.marginBottomTwips !== undefined
          ? { marginBottom: update.marginBottomTwips }
          : {}),
        ...(update.marginLeftTwips !== undefined ? { marginLeft: update.marginLeftTwips } : {}),
        ...(update.scope !== undefined ? { scope: update.scope } : {}),
        type: 'setPageSetup',
      }).ok;
    },
  };
}
