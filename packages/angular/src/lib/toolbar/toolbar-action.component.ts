import {
  ChangeDetectionStrategy,
  Component,
  booleanAttribute,
  computed,
  inject,
  input,
  output,
} from '@angular/core';
import { DOCX_TOOLBAR_PLACEMENT, DocxToolbarState } from './toolbar-context';
import { guardToolbarMousedown } from './toolbar-registry';

const optionalBoolean = (value: unknown): boolean | undefined =>
  value === undefined || value === null ? undefined : booleanAttribute(value);

/**
 * A toolbar button that runs your own action. Put your icon in the button.
 *
 * In the bar, the button has the look of the built-in buttons. When its group collapses, the
 * button renders as a row of the "⋯" panel, with the icon and the label, and a click closes the
 * panel.
 *
 * @example
 * ```html
 * <button docxToolbarAction label="Next change" (select)="next()">
 *   <svg [docxChromeIcon]="nextPaths"></svg>
 * </button>
 * ```
 *
 * @public
 */
@Component({
  selector: 'button[docxToolbarAction]',
  changeDetection: ChangeDetectionStrategy.OnPush,
  host: {
    type: 'button',
    '[class.docx-toolbar__button]': '!panel',
    '[class.docx-toolbar__more-command]': 'panel',
    '[disabled]': 'disabled()',
    '[attr.data-active]': "active() ? '' : null",
    '[attr.data-disabled]': "!panel && disabled() ? '' : null",
    '[attr.aria-pressed]': 'active() ?? null',
    '[attr.aria-label]': 'panel ? null : label()',
    '[attr.title]': 'title()',
    '(mousedown)': 'guard($event)',
    '(click)': 'run($event)',
  },
  template: `<span
      [class.docx-toolbar__more-command-icon]="panel"
      [attr.aria-hidden]="panel ? 'true' : null"
      [style.display]="panel ? null : 'contents'"
      ><ng-content
    /></span>
    @if (panel) {
      <span class="docx-toolbar__more-command-label">{{ label() }}</span>
    }`,
})
export class DocxToolbarActionComponent {
  /** The accessible name and the title of the button, and the text of its panel row. */
  readonly label = input.required<string>();
  /** The pressed state. Set it for a toggle; leave it out for a plain action. */
  readonly active = input<boolean | undefined, unknown>(undefined, { transform: optionalBoolean });
  /** Disable the button. */
  readonly disabled = input(false, { transform: booleanAttribute });
  /** The title of a disabled button. */
  readonly disabledReason = input<string | undefined>(undefined);
  /** Emits when the user activates the button. */
  readonly select = output<void>();

  private readonly state = inject(DocxToolbarState, { optional: true });

  /** @internal */
  protected readonly panel = inject(DOCX_TOOLBAR_PLACEMENT) === 'panel';
  /** @internal */
  protected readonly title = computed(() => {
    if (!this.disabled()) return this.panel ? null : this.label();
    return this.disabledReason() ?? (this.panel ? null : this.label());
  });
  /** @internal */
  protected readonly guard: (event: Event) => void = guardToolbarMousedown;

  /** @internal */
  protected run(event: MouseEvent): void {
    if (this.disabled()) return;
    this.select.emit();
    if (this.panel) this.state?.closeMore(event.detail === 0);
  }
}
