import { ChangeDetectionStrategy, Component, computed, input } from '@angular/core';
import { injectEditorCommand, injectTranslation } from '@docx-editor.dev/angular';
import { CHROME_GROUPS, type ChromeSlotId } from '@docx-editor.dev/core/editor';

/** The label key and icon of each chrome slot, from the engine's chrome registry. */
const CONTROLS = new Map<string, { labelKey: string; paths: readonly string[] | null }>(
  CHROME_GROUPS.flatMap((group) =>
    group.controls.map((control) => [`${group.id}.${control.id}`, control] as const)
  )
);

/** The registry label key of a chrome slot, or the slot id when the registry has none. */
export function slotLabelKey(slot: ChromeSlotId): string {
  return CONTROLS.get(slot)?.labelKey ?? slot;
}

/**
 * One toolbar button for a chrome slot, such as `text.bold`. The engine decides if the button
 * is enabled and pressed. The label comes from the active translations.
 */
@Component({
  selector: 'app-command-button',
  changeDetection: ChangeDetectionStrategy.OnPush,
  template: `
    <button
      type="button"
      class="command"
      [attr.data-slot]="slot()"
      [disabled]="!command.isEnabled()"
      [attr.aria-pressed]="toggle() ? command.isActive() : null"
      [attr.aria-label]="label()"
      [title]="command.disabledReason() ?? label()"
      (mousedown)="$event.preventDefault()"
      (click)="command.execute()"
    >
      @if (paths().length) {
        <svg viewBox="0 -960 960 960" width="18" height="18" aria-hidden="true" focusable="false">
          @for (d of paths(); track $index) {
            <path [attr.d]="d" />
          }
        </svg>
      } @else {
        {{ label() }}
      }
    </button>
  `,
})
export class CommandButtonComponent {
  /** The chrome slot that the button runs. */
  readonly slot = input.required<ChromeSlotId>();
  /** Show the pressed state, for toggles such as bold or an alignment. */
  readonly toggle = input(false);

  protected readonly command = injectEditorCommand(() => this.slot());
  private readonly t = injectTranslation().t;

  protected readonly label = computed(() => {
    const key = CONTROLS.get(this.slot())?.labelKey;
    return key ? (this.t() as (key: string) => string)(key) : this.slot();
  });
  protected readonly paths = computed(() => CONTROLS.get(this.slot())?.paths ?? []);
}
