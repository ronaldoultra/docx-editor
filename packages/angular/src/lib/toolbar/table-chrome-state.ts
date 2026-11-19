// The state that the contextual table controls share: whether the selection is in a table, the
// border draft (target, style, width, color) that the next border command applies, and the
// engine's answer for each table slot under that draft.

import {
  Injectable,
  NgZone,
  computed,
  inject,
  signal,
  untracked,
  type Signal,
} from '@angular/core';
import type { EditorSnapshot } from '@docx-editor.dev/core/contracts/editor';
import {
  DEFAULT_TABLE_CHROME_DRAFT,
  TABLE_CHROME_SLOT_IDS,
  runTableChromeCommand,
  tableChromeToolbarState,
  type TableChromeDraft,
  type TableChromeSlotId,
} from '@docx-editor.dev/core/editor';
import { localizeDisabledReason } from '@docx-editor.dev/core/i18n';
import { injectDocxEditor } from '../context';
import { injectEditorState } from '../editor-state';
import { injectTranslation } from '../i18n';

/** The engine's answer for one table slot. @internal */
export interface TableSlotState {
  readonly enabled: boolean;
  /** The localized reason when the slot is disabled, otherwise `null`. */
  readonly disabledReason: string | null;
}

type SlotStates = Readonly<Record<TableChromeSlotId, TableSlotState>>;

const NOT_READY: TableSlotState = { enabled: false, disabledReason: null };
const sameStates = (a: SlotStates, b: SlotStates): boolean =>
  TABLE_CHROME_SLOT_IDS.every(
    (slot) =>
      a[slot]?.enabled === b[slot]?.enabled && a[slot]?.disabledReason === b[slot]?.disabledReason
  );
const selectTable = (snapshot: EditorSnapshot) => snapshot.table !== null;

/**
 * The table chrome state of one toolbar. The toolbar provides it.
 *
 * @internal
 */
@Injectable()
export class DocxTableChromeState {
  private readonly editor = injectDocxEditor();
  private readonly zone = inject(NgZone);
  private readonly translation = injectTranslation();
  private readonly draftSignal = signal<TableChromeDraft>(DEFAULT_TABLE_CHROME_DRAFT);

  /** True while the selection is in a table. */
  readonly visible: Signal<boolean> = injectEditorState(selectTable);

  /** The border draft. */
  readonly draft: Signal<TableChromeDraft> = this.draftSignal.asReadonly();

  // Read on each engine tick, because the engine's answer follows the cell selection too.
  private readonly states = injectEditorState((): SlotStates => {
    const editor = this.editor();
    const draft = this.draftSignal();
    const t = this.translation.t();
    const states = {} as Record<TableChromeSlotId, TableSlotState>;
    for (const slot of TABLE_CHROME_SLOT_IDS) {
      const state = untracked(() => tableChromeToolbarState(editor, slot, draft));
      states[slot] = {
        enabled: state.enabled,
        disabledReason: state.enabled ? null : localizeDisabledReason(state.disabledReason, t),
      };
    }
    return states;
  }, sameStates);

  /** The state of one slot, as a signal. */
  slot(slot: TableChromeSlotId): Signal<TableSlotState> {
    return computed(() => (this.visible() ? this.states()[slot] : undefined) ?? NOT_READY);
  }

  /** Run a table slot with a value. A border command that the engine applies updates the draft. */
  apply(slot: TableChromeSlotId, value: unknown): void {
    const editor = untracked(this.editor);
    const draft = untracked(this.draftSignal);
    const { result, nextDraft } = this.zone.runOutsideAngular(() =>
      runTableChromeCommand(editor, slot, value, draft)
    );
    if (result.ok && nextDraft) this.draftSignal.set(nextDraft);
  }
}
