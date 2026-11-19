import {
  ChangeDetectionStrategy,
  Component,
  ElementRef,
  InjectionToken,
  NgZone,
  computed,
  inject,
  signal,
  untracked,
} from '@angular/core';
import type { EditorCommand, EditorSnapshot } from '@docx-editor.dev/core/contracts/editor';
import {
  isHTMLElementNode,
  lineSpacingMenuRowLabel,
  lineSpacingMenuRows,
  type ChromeSlotId,
  type LineSpacingMenuRow,
} from '@docx-editor.dev/core/editor';
import { injectDocxEditor } from '../context';
import { injectEditorCommand, type EditorCommandState } from '../editor-command';
import { injectEditorState } from '../editor-state';
import { injectTranslation } from '../i18n';
import { DocxChromeIconComponent } from './chrome-icon.component';
import { injectToolbarLabel } from './toolbar-context';
import { injectPopupDismiss } from './toolbar-popup';
import { chromeControlForSlot, guardToolbarMousedown } from './toolbar-registry';

const ALIGNMENT_SLOTS = [
  'alignment.left',
  'alignment.center',
  'alignment.right',
  'alignment.justify',
] as const satisfies readonly ChromeSlotId[];

/**
 * The alignment dropdown (`alignment`): a trigger that shows the alignment of the selection, and
 * the left, center, right, and justify buttons.
 *
 * @example
 * ```html
 * <div docxToolbarAlignment></div>
 * ```
 *
 * @public
 */
@Component({
  selector: 'div[docxToolbarAlignment]',
  changeDetection: ChangeDetectionStrategy.OnPush,
  imports: [DocxChromeIconComponent],
  host: { class: 'docx-toolbar__alignment', 'data-slot': 'alignment' },
  template: `<button
      type="button"
      class="docx-toolbar__button docx-toolbar__alignment-trigger"
      aria-haspopup="true"
      [disabled]="!enabled()"
      [attr.data-disabled]="enabled() ? null : ''"
      [attr.aria-expanded]="open()"
      [attr.aria-label]="current().text"
      [attr.title]="
        enabled() ? current().text : (current().state.disabledReason() ?? current().text)
      "
      (mousedown)="guard($event)"
      (click)="open.set(!open())"
    >
      @if (current().paths; as icon) {
        <svg [docxChromeIcon]="icon"></svg>
      }
      <span class="docx-toolbar__picker-caret" aria-hidden="true">▾</span>
    </button>
    @if (open()) {
      <div class="docx-toolbar__menu docx-toolbar__alignment-popup">
        @for (option of options(); track option.slot) {
          <button
            type="button"
            class="docx-toolbar__button docx-toolbar__alignment-option"
            [attr.data-slot]="option.slot"
            [disabled]="!option.state.isEnabled()"
            [attr.data-active]="option.state.isActive() ? '' : null"
            [attr.aria-pressed]="option.state.isActive()"
            [attr.aria-label]="option.text"
            [attr.title]="option.state.disabledReason() ?? option.text"
            (mousedown)="guard($event)"
            (click)="choose(option.state)"
          >
            @if (option.paths; as icon) {
              <svg [docxChromeIcon]="icon"></svg>
            }
          </button>
        }
      </div>
    }`,
})
export class DocxToolbarAlignmentComponent {
  private readonly host: HTMLElement = inject(ElementRef).nativeElement;
  private readonly states = ALIGNMENT_SLOTS.map((slot) => injectEditorCommand(slot));
  private readonly label = injectToolbarLabel();

  /** @internal */
  protected readonly open = signal(false);
  /** @internal */
  protected readonly guard: (event: Event) => void = guardToolbarMousedown;
  /** @internal */
  protected readonly options = computed(() => {
    const label = this.label();
    return ALIGNMENT_SLOTS.map((slot, index) => {
      const control = chromeControlForSlot(slot);
      return {
        slot,
        paths: control?.paths ?? null,
        text: label(control?.labelKey ?? slot),
        state: this.states[index]!,
      };
    });
  });
  /** @internal */
  protected readonly current = computed(
    () => this.options().find((option) => option.state.isActive()) ?? this.options()[0]!
  );
  /** @internal */
  protected readonly enabled = computed(() =>
    this.options().some((option) => option.state.isEnabled())
  );

  constructor() {
    injectPopupDismiss({
      open: this.open,
      root: () => this.host,
      close: (fromInside) => {
        this.open.set(false);
        if (fromInside) this.host.querySelector<HTMLElement>('[aria-haspopup]')?.focus();
      },
    });
  }

  /** @internal */
  protected choose(state: { readonly execute: () => boolean }): void {
    state.execute();
    this.open.set(false);
  }
}

/**
 * Opens the paragraph dialog from a toolbar control. The trigger receives the focus back when
 * the dialog closes. Provide it to show "Line spacing options…" in the line spacing menu.
 *
 * @public
 */
export const DOCX_PARAGRAPH_DIALOG = new InjectionToken<{
  readonly open: (trigger: HTMLElement | null) => void;
}>('DOCX_PARAGRAPH_DIALOG');

const DEFAULT_PARAGRAPH_SPACE_PT = 10;
// Remove writes an explicit zero, not nothing: without the attribute, the style's own space
// comes back, so on a Word default document Remove gave the space straight back.
const REMOVED_PARAGRAPH_SPACE_PT = 0;

interface SpacingSlice {
  readonly lineSpacing: Exclude<
    NonNullable<EditorSnapshot['formatting']>['lineSpacing'],
    undefined
  > | null;
  readonly spaceBeforePt: number | null;
  readonly spaceAfterPt: number | null;
}

const selectSpacing = (snapshot: EditorSnapshot): SpacingSlice => ({
  lineSpacing: snapshot.formatting?.lineSpacing ?? null,
  spaceBeforePt: snapshot.formatting?.spaceBeforePt ?? null,
  spaceAfterPt: snapshot.formatting?.spaceAfterPt ?? null,
});

const sameSpacing = (a: SpacingSlice, b: SpacingSlice) =>
  a.spaceBeforePt === b.spaceBeforePt &&
  a.spaceAfterPt === b.spaceAfterPt &&
  a.lineSpacing?.rule === b.lineSpacing?.rule &&
  a.lineSpacing?.value === b.lineSpacing?.value;

/**
 * The line spacing menu (`list.lineSpacing`): the spacing presets, with the spacing in force
 * always ticked, even one that no preset matches, and the space before and after rows.
 *
 * @example
 * ```html
 * <span docxToolbarLineSpacing></span>
 * ```
 *
 * @public
 */
@Component({
  selector: 'span[docxToolbarLineSpacing]',
  changeDetection: ChangeDetectionStrategy.OnPush,
  imports: [DocxChromeIconComponent],
  host: { class: 'docx-toolbar__line-spacing', 'data-slot': 'list.lineSpacing' },
  template: `<button
      type="button"
      class="docx-toolbar__button docx-toolbar__line-spacing-trigger"
      aria-haspopup="menu"
      [disabled]="!command.isEnabled()"
      [attr.data-disabled]="command.isEnabled() ? null : ''"
      [attr.aria-expanded]="open()"
      [attr.aria-label]="label()('lineSpacing.label')"
      [attr.title]="command.disabledReason() ?? label()('lineSpacing.label')"
      (mousedown)="guard($event)"
      (click)="open.set(!open())"
    >
      @if (paths; as icon) {
        <svg [docxChromeIcon]="icon"></svg>
      }
      <span class="docx-toolbar__picker-caret" aria-hidden="true">▾</span>
    </button>
    @if (open() && command.isEnabled()) {
      <div class="docx-toolbar__menu docx-toolbar__line-spacing-menu" role="menu">
        @for (row of rows(); track row.key) {
          <button
            type="button"
            role="menuitemradio"
            class="docx-toolbar__menu-item"
            [attr.aria-checked]="row.row.selected"
            [attr.data-selected]="row.row.selected ? '' : null"
            (mousedown)="guard($event)"
            (click)="applyRow(row.row)"
          >
            {{ row.text }}
            @if (row.row.selected) {
              <span class="docx-toolbar__menu-check" aria-hidden="true">✓</span>
            }
          </button>
        }
        @if (dialog) {
          <div class="docx-toolbar__menu-separator" role="separator"></div>
          <button
            type="button"
            role="menuitem"
            class="docx-toolbar__menu-item"
            data-slot="paragraph.dialog"
            [disabled]="!dialogCommand.isEnabled()"
            [attr.title]="dialogCommand.disabledReason()"
            (mousedown)="guard($event)"
            (click)="openDialog()"
          >
            {{ label()('lineSpacing.options') }}
          </button>
        }
        <div class="docx-toolbar__menu-separator" role="separator"></div>
        <button
          type="button"
          role="menuitem"
          class="docx-toolbar__menu-item"
          (mousedown)="guard($event)"
          (click)="toggleSpace('beforePt')"
        >
          {{
            label()(hasBefore() ? 'lineSpacing.removeSpaceBefore' : 'lineSpacing.addSpaceBefore')
          }}
        </button>
        <button
          type="button"
          role="menuitem"
          class="docx-toolbar__menu-item"
          (mousedown)="guard($event)"
          (click)="toggleSpace('afterPt')"
        >
          {{ label()(hasAfter() ? 'lineSpacing.removeSpaceAfter' : 'lineSpacing.addSpaceAfter') }}
        </button>
      </div>
    }`,
})
export class DocxToolbarLineSpacingComponent {
  private readonly editor = injectDocxEditor();
  private readonly zone = inject(NgZone);
  private readonly host: HTMLElement = inject(ElementRef).nativeElement;
  private readonly spacing = injectEditorState(selectSpacing, sameSpacing);
  private readonly translation = injectTranslation();

  /** @internal */
  protected readonly command: EditorCommandState = injectEditorCommand('list.lineSpacing');
  /**
   * The options row opens a `setParagraphFormat` editor, so that command decides whether the
   * row works, through the `paragraph.dialog` slot.
   *
   * @internal
   */
  protected readonly dialogCommand: EditorCommandState = injectEditorCommand('paragraph.dialog');
  /** @internal */
  protected readonly dialog = inject(DOCX_PARAGRAPH_DIALOG, { optional: true });
  /** @internal */
  protected readonly label = injectToolbarLabel();
  /** @internal */
  protected readonly open = signal(false);
  /** @internal */
  protected readonly guard: (event: Event) => void = guardToolbarMousedown;
  /** @internal */
  protected readonly paths = chromeControlForSlot('list.lineSpacing')?.paths ?? null;
  /** @internal */
  protected readonly hasBefore = computed(() => (this.spacing().spaceBeforePt ?? 0) > 0);
  /** @internal */
  protected readonly hasAfter = computed(() => (this.spacing().spaceAfterPt ?? 0) > 0);
  /** @internal */
  protected readonly rows = computed(() => {
    const label = this.label();
    const lang = this.translation.lang();
    return lineSpacingMenuRows(this.spacing().lineSpacing).map((row) => ({
      key: `${row.spacing.rule}:${row.spacing.value}`,
      row,
      text: lineSpacingMenuRowLabel(row, label, lang),
    }));
  });

  constructor() {
    injectPopupDismiss({
      open: this.open,
      root: () => this.host,
      close: (fromInside) => {
        this.open.set(false);
        if (fromInside) this.host.querySelector<HTMLElement>('[aria-haspopup]')?.focus();
      },
    });
  }

  /** @internal */
  protected applyRow(row: LineSpacingMenuRow): void {
    this.open.set(false);
    // The selected row has no command: it already holds.
    if (row.command) this.run(row.command);
  }

  /** @internal */
  protected toggleSpace(field: 'beforePt' | 'afterPt'): void {
    this.open.set(false);
    const present = field === 'beforePt' ? this.hasBefore() : this.hasAfter();
    const points = present ? REMOVED_PARAGRAPH_SPACE_PT : DEFAULT_PARAGRAPH_SPACE_PT;
    this.run({ type: 'setParagraphSpacing', [field]: points });
  }

  /** @internal */
  protected openDialog(): void {
    this.open.set(false);
    // The trigger, not the row: the menu closes in this same gesture.
    const trigger = this.host.querySelector('.docx-toolbar__line-spacing-trigger');
    this.dialog?.open(isHTMLElementNode(trigger) ? trigger : null);
  }

  private run(command: EditorCommand): void {
    const editor = untracked(this.editor);
    if (!editor) return;
    this.zone.runOutsideAngular(() => {
      if (editor.can(command).ok) editor.exec(command);
    });
  }
}
