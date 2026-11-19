import {
  ChangeDetectionStrategy,
  Component,
  computed,
  viewChild,
  type ElementRef,
} from '@angular/core';
import {
  injectChromeDir,
  injectEditorValueCommand,
  injectHistoryGroup,
  injectScopeClassName,
  injectTranslation,
} from '@docx-editor.dev/angular';
import type { ChromeSlotId } from '@docx-editor.dev/core/editor';
import { CommandButtonComponent, slotLabelKey } from './command-button.component';

const FONT_FAMILIES = ['Calibri', 'Arial', 'Times New Roman', 'Georgia', 'Courier New'];
const FONT_SIZES_PT = [8, 9, 10, 11, 12, 14, 16, 18, 20, 24, 28, 36, 48, 72];
const LINE_SPACINGS = [1, 1.15, 1.5, 2, 3];
const HIGHLIGHTS = ['none', 'yellow', 'green', 'cyan', 'magenta', 'lightGray'];

/**
 * A formatting bar made from the composition API. Each control reads its state from the
 * engine, so it follows the selection, the editing mode, and undo.
 */
@Component({
  selector: 'app-format-bar',
  imports: [CommandButtonComponent],
  changeDetection: ChangeDetectionStrategy.OnPush,
  host: {
    '[class]': "scope() + 'format-bar'",
    '[attr.dir]': 'chromeDir()',
    role: 'toolbar',
    'aria-label': 'Formatação',
  },
  template: `
    <div class="group">
      <app-command-button slot="history.undo" />
      <app-command-button slot="history.redo" />
    </div>
    <div class="group">
      <select
        [attr.aria-label]="label('font.family')"
        [disabled]="!family.isEnabled()"
        (change)="setFamily($event)"
      >
        @if (family.value() === null) {
          <option value="" selected>—</option>
        }
        @for (name of families(); track name) {
          <option [value]="name" [selected]="name === family.value()">{{ name }}</option>
        }
      </select>
      <select
        class="narrow"
        [attr.aria-label]="label('font.size')"
        [disabled]="!size.isEnabled()"
        (change)="setSize($event)"
      >
        @if (sizePt() === null) {
          <option value="" selected>—</option>
        }
        @for (pt of sizes(); track pt) {
          <option [value]="pt" [selected]="pt === sizePt()">{{ pt }}</option>
        }
      </select>
    </div>
    <div class="group">
      <app-command-button slot="text.bold" [toggle]="true" />
      <app-command-button slot="text.italic" [toggle]="true" />
      <app-command-button slot="text.underline" [toggle]="true" />
      <app-command-button slot="text.strike" [toggle]="true" />
      <app-command-button slot="script.super" [toggle]="true" />
      <app-command-button slot="script.sub" [toggle]="true" />
    </div>
    <div class="group">
      <input
        #colorInput
        type="color"
        [attr.aria-label]="label('text.color')"
        [disabled]="!color.isEnabled()"
        [value]="'#' + (color.value() ?? '000000')"
        (input)="setColor($event)"
      />
      <select
        [attr.aria-label]="label('text.highlight')"
        [disabled]="!highlight.isEnabled()"
        (change)="setHighlight($event)"
      >
        @for (name of highlights; track name) {
          <option [value]="name" [selected]="name === (highlight.value() ?? 'none')">
            {{ name }}
          </option>
        }
      </select>
    </div>
    <div class="group">
      <app-command-button slot="alignment.left" [toggle]="true" />
      <app-command-button slot="alignment.center" [toggle]="true" />
      <app-command-button slot="alignment.right" [toggle]="true" />
      <app-command-button slot="alignment.justify" [toggle]="true" />
      <app-command-button slot="direction.ltr" [toggle]="true" />
      <app-command-button slot="direction.rtl" [toggle]="true" />
    </div>
    <div class="group">
      <app-command-button slot="list.bullet" [toggle]="true" />
      <app-command-button slot="list.numbered" [toggle]="true" />
      <app-command-button slot="list.outdent" />
      <app-command-button slot="list.indent" />
      <select
        class="narrow"
        [attr.aria-label]="label('list.lineSpacing')"
        [disabled]="!spacing.isEnabled()"
        (change)="setSpacing($event)"
      >
        @if (spacing.value() === null) {
          <option value="" selected>—</option>
        }
        @for (value of spacings; track value) {
          <option [value]="value" [selected]="value === spacing.value()">{{ value }}</option>
        }
      </select>
    </div>
    <div class="group">
      <app-command-button slot="format.clear" />
    </div>
  `,
})
export class FormatBarComponent {
  protected readonly scope = injectScopeClassName();
  // Set when the catalog or the docxEditorDir input asks for a direction.
  protected readonly chromeDir = injectChromeDir();
  private readonly t = injectTranslation().t;

  protected readonly family = injectEditorValueCommand('font.family');
  protected readonly size = injectEditorValueCommand('font.size');
  protected readonly color = injectEditorValueCommand('text.color');
  protected readonly highlight = injectEditorValueCommand('text.highlight');
  protected readonly spacing = injectEditorValueCommand('list.lineSpacing');

  protected readonly highlights = HIGHLIGHTS;
  protected readonly spacings = LINE_SPACINGS;

  // A drag in the color picker sends many values. The history group makes them one undo step.
  private readonly colorInput = viewChild<ElementRef<HTMLInputElement>>('colorInput');
  private readonly colorGesture = injectHistoryGroup({
    kind: 'native-color',
    element: this.colorInput,
  });

  /** The translated label of a value slot. */
  protected label(slot: ChromeSlotId): string {
    return (this.t() as (key: string) => string)(slotLabelKey(slot));
  }

  /** The document's font, when it is not in the list, comes first. */
  protected readonly families = computed(() => {
    const current = this.family.value();
    return current && !FONT_FAMILIES.includes(current)
      ? [current, ...FONT_FAMILIES]
      : FONT_FAMILIES;
  });

  /** The engine gives font sizes in half-points. The list shows points. */
  protected readonly sizePt = computed(() => {
    const halfPoints = this.size.value();
    return halfPoints === null ? null : halfPoints / 2;
  });
  protected readonly sizes = computed(() => {
    const current = this.sizePt();
    return current !== null && !FONT_SIZES_PT.includes(current)
      ? [...FONT_SIZES_PT, current].sort((a, b) => a - b)
      : FONT_SIZES_PT;
  });

  protected setFamily(event: Event) {
    const value = (event.target as HTMLSelectElement).value;
    if (value) this.family.execute(value);
  }

  protected setSize(event: Event) {
    const value = Number((event.target as HTMLSelectElement).value);
    if (value > 0) this.size.execute(value * 2);
  }

  protected setColor(event: Event) {
    const value = (event.target as HTMLInputElement).value.replace(/^#/, '').toUpperCase();
    this.color.execute(value, this.colorGesture.options());
  }

  protected setHighlight(event: Event) {
    this.highlight.execute((event.target as HTMLSelectElement).value);
  }

  protected setSpacing(event: Event) {
    const value = Number((event.target as HTMLSelectElement).value);
    if (value > 0) this.spacing.execute(value);
  }
}
