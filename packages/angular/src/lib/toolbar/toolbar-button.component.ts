import { ChangeDetectionStrategy, Component, computed, inject, input } from '@angular/core';
import { chromeSlotIsToggle, type ChromeSlotId } from '@docx-editor.dev/core/editor';
import { platformShortcut } from '@docx-editor.dev/core/i18n';
import { injectEditorCommand, type EditorCommandState } from '../editor-command';
import { DocxChromeIconComponent } from './chrome-icon.component';
import { DocxToolbarState, injectToolbarLabel } from './toolbar-context';
import { chromeControlForSlot, guardToolbarMousedown } from './toolbar-registry';

/**
 * A toolbar button that runs the engine command of one chrome slot.
 *
 * The engine decides whether the button is enabled (`toolbarCommandState`). A disabled button
 * shows the localized engine reason as its title. Toggle slots, such as `text.bold`, carry
 * `aria-pressed`. The icon comes from the registry, unless you put your own content in the
 * button.
 *
 * @example
 * ```html
 * <button docxToolbarButton="text.bold"></button>
 * <button docxToolbarButton="history.undo"><my-undo-icon /></button>
 * ```
 *
 * @public
 */
@Component({
  selector: 'button[docxToolbarButton]',
  changeDetection: ChangeDetectionStrategy.OnPush,
  imports: [DocxChromeIconComponent],
  host: {
    type: 'button',
    class: 'docx-toolbar__button',
    '[attr.data-slot]': 'docxToolbarButton()',
    '[disabled]': '!command.isEnabled()',
    '[attr.data-active]': "command.isActive() ? '' : null",
    '[attr.data-disabled]': "command.isEnabled() ? null : ''",
    '[attr.aria-pressed]': 'toggle() ? command.isActive() : null',
    '[attr.data-value]': 'command.value()',
    '[attr.aria-label]': 'text()',
    '[attr.title]': 'command.disabledReason() ?? text()',
    '(mousedown)': 'guard($event)',
    '(click)': 'command.execute()',
  },
  template: `<ng-content>
    @if (paths(); as icon) {
      <svg [docxChromeIcon]="icon"></svg>
    }
  </ng-content>`,
})
export class DocxToolbarButtonComponent {
  /** The chrome slot that the button drives, such as `'text.bold'`. */
  readonly docxToolbarButton = input.required<ChromeSlotId>();

  private readonly label = injectToolbarLabel();

  /** @internal */
  protected readonly command: EditorCommandState = injectEditorCommand(() =>
    this.docxToolbarButton()
  );
  /** @internal */
  protected readonly toggle = computed(() => chromeSlotIsToggle(this.docxToolbarButton()));
  /** @internal */
  protected readonly paths = computed(
    () => chromeControlForSlot(this.docxToolbarButton())?.paths ?? null
  );
  /**
   * A registry label often names its shortcut ("Bold (Ctrl+B)"). The text names the modifier
   * keys of this keyboard.
   *
   * @internal
   */
  protected readonly text = computed(() => {
    const slot = this.docxToolbarButton();
    return platformShortcut(this.label()(chromeControlForSlot(slot)?.labelKey ?? slot));
  });
  /** @internal */
  protected readonly guard: (event: Event) => void = guardToolbarMousedown;
}

/**
 * A vertical separator between toolbar controls.
 *
 * @example
 * ```html
 * <div docxToolbarSeparator></div>
 * ```
 *
 * @public
 */
@Component({
  selector: 'div[docxToolbarSeparator]',
  changeDetection: ChangeDetectionStrategy.OnPush,
  host: {
    role: 'separator',
    'aria-orientation': 'vertical',
    class: 'docx-toolbar__separator',
  },
  template: '',
})
export class DocxToolbarSeparatorComponent {}

/**
 * One command row of the "⋯" panel: the same command, state, and label as the bar button.
 *
 * @internal
 */
@Component({
  selector: 'button[docxToolbarMoreItem]',
  changeDetection: ChangeDetectionStrategy.OnPush,
  imports: [DocxChromeIconComponent],
  host: {
    type: 'button',
    class: 'docx-toolbar__more-command',
    '[attr.data-slot]': 'docxToolbarMoreItem()',
    '[disabled]': '!command.isEnabled()',
    '[attr.title]': 'command.disabledReason()',
    '[attr.aria-pressed]': 'toggle() ? command.isActive() : null',
    '[attr.data-active]': "command.isActive() ? '' : null",
    '[attr.data-value]': 'command.value()',
    '(mousedown)': 'guard($event)',
    '(click)': 'run($event)',
  },
  template: `<span class="docx-toolbar__more-command-icon" aria-hidden="true">
      @if (paths(); as icon) {
        <svg [docxChromeIcon]="icon"></svg>
      }
    </span>
    <span class="docx-toolbar__more-command-label">{{ text() }}</span>`,
})
export class DocxToolbarMoreItemComponent {
  /** The chrome slot of the row. */
  readonly docxToolbarMoreItem = input.required<ChromeSlotId>();

  private readonly label = injectToolbarLabel();
  private readonly state = inject(DocxToolbarState, { optional: true });

  /** @internal */
  protected readonly command: EditorCommandState = injectEditorCommand(() =>
    this.docxToolbarMoreItem()
  );
  /** @internal */
  protected readonly toggle = computed(() => chromeSlotIsToggle(this.docxToolbarMoreItem()));
  /** @internal */
  protected readonly paths = computed(
    () => chromeControlForSlot(this.docxToolbarMoreItem())?.paths ?? null
  );
  /** @internal */
  protected readonly text = computed(() => {
    const slot = this.docxToolbarMoreItem();
    return platformShortcut(this.label()(chromeControlForSlot(slot)?.labelKey ?? slot));
  });
  /** @internal */
  protected readonly guard: (event: Event) => void = guardToolbarMousedown;

  /** @internal */
  protected run(event: MouseEvent): void {
    this.command.execute();
    // A keyboard activation has no click count. The row leaves the page, so the focus goes
    // back to the trigger.
    this.state?.closeMore(event.detail === 0);
  }
}
