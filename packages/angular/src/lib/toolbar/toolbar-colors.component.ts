import {
  ChangeDetectionStrategy,
  Component,
  Directive,
  ElementRef,
  NgZone,
  computed,
  inject,
  input,
  output,
  signal,
  untracked,
  type Signal,
  type WritableSignal,
} from '@angular/core';
import type { EditorSnapshot } from '@docx-editor.dev/core/contracts/editor';
import {
  colorPickerGrid,
  commandForSlotValue,
  type ChromeSlotId,
} from '@docx-editor.dev/core/editor';
import { injectDocxEditor } from '../context';
import { injectEditorCommand, type EditorCommandState } from '../editor-command';
import { injectEditorState } from '../editor-state';
import { DocxChromeIconComponent } from './chrome-icon.component';
import { injectToolbarLabel } from './toolbar-context';
import { injectPopupDismiss } from './toolbar-popup';
import { chromeControlForSlot, guardToolbarMousedown } from './toolbar-registry';

interface Swatch {
  readonly value: string;
  readonly title: string;
  readonly css: string;
  readonly light: boolean;
}

interface SwatchDef {
  readonly value: string;
  readonly labelKey: string;
  readonly css: string;
}

const STANDARD_COLOR_SWATCHES: readonly SwatchDef[] = [
  { value: 'C00000', labelKey: 'colorPicker.colors.darkRed', css: '#c00000' },
  { value: 'FF0000', labelKey: 'colorPicker.colors.red', css: '#ff0000' },
  { value: 'FFC000', labelKey: 'colorPicker.colors.orange', css: '#ffc000' },
  { value: 'FFFF00', labelKey: 'colorPicker.colors.yellow', css: '#ffff00' },
  { value: '92D050', labelKey: 'colorPicker.colors.lightGreen', css: '#92d050' },
  { value: '00B050', labelKey: 'colorPicker.colors.green', css: '#00b050' },
  { value: '00B0F0', labelKey: 'colorPicker.colors.lightBlue', css: '#00b0f0' },
  { value: '0070C0', labelKey: 'colorPicker.colors.blue', css: '#0070c0' },
  { value: '002060', labelKey: 'colorPicker.colors.darkBlue', css: '#002060' },
  { value: '7030A0', labelKey: 'colorPicker.colors.purple', css: '#7030a0' },
];

const HIGHLIGHT_SWATCHES: readonly SwatchDef[] = [
  { value: 'yellow', labelKey: 'colorPicker.colors.yellow', css: '#ffff00' },
  { value: 'green', labelKey: 'colorPicker.colors.brightGreen', css: '#00ff00' },
  { value: 'cyan', labelKey: 'colorPicker.colors.cyan', css: '#00ffff' },
  { value: 'magenta', labelKey: 'colorPicker.colors.magenta', css: '#ff00ff' },
  { value: 'blue', labelKey: 'colorPicker.colors.blue', css: '#0000ff' },
  { value: 'red', labelKey: 'colorPicker.colors.red', css: '#ff0000' },
  { value: 'darkBlue', labelKey: 'colorPicker.colors.darkBlue', css: '#000080' },
  { value: 'darkCyan', labelKey: 'colorPicker.colors.darkCyan', css: '#008080' },
  { value: 'darkGreen', labelKey: 'colorPicker.colors.darkGreen', css: '#008000' },
  { value: 'darkMagenta', labelKey: 'colorPicker.colors.darkMagenta', css: '#800080' },
  { value: 'darkRed', labelKey: 'colorPicker.colors.darkRed', css: '#800000' },
  { value: 'darkYellow', labelKey: 'colorPicker.colors.darkYellow', css: '#808000' },
  { value: 'darkGray', labelKey: 'colorPicker.colors.darkGray', css: '#808080' },
  { value: 'lightGray', labelKey: 'colorPicker.colors.lightGray', css: '#c0c0c0' },
  { value: 'black', labelKey: 'colorPicker.colors.black', css: '#000000' },
  { value: 'white', labelKey: 'colorPicker.colors.white', css: '#ffffff' },
];
const HIGHLIGHT_GRID = HIGHLIGHT_SWATCHES.filter((swatch) => swatch.value !== 'white');
const HIGHLIGHT_CSS = new Map(HIGHLIGHT_SWATCHES.map((swatch) => [swatch.value, swatch.css]));

const DEFAULT_THEME_HEXES: readonly string[] = [
  'FFFFFF',
  '000000',
  'E7E6E6',
  '44546A',
  '4472C4',
  'ED7D31',
  'A5A5A5',
  'FFC000',
  '5B9BD5',
  '70AD47',
];

const THEME_COLUMN_KEYS: readonly string[] = [
  'colorPicker.theme.background1',
  'colorPicker.theme.text1',
  'colorPicker.theme.background2',
  'colorPicker.theme.text2',
  'colorPicker.theme.accent1',
  'colorPicker.theme.accent2',
  'colorPicker.theme.accent3',
  'colorPicker.theme.accent4',
  'colorPicker.theme.accent5',
  'colorPicker.theme.accent6',
];

const HEX_VALUE = /^[0-9A-Fa-f]{6}$/;

/** `accent1:40` is Lighter 40%, `accent1:-25` Darker 25%. */
function variantLabelKey(pick: string): string {
  const shift = Number(pick.split(':')[1] ?? 0);
  return shift > 0 ? `colorPicker.theme.lighter${shift}` : `colorPicker.theme.darker${-shift}`;
}

function isLightHex(hex: string): boolean {
  const r = parseInt(hex.slice(0, 2), 16);
  const g = parseInt(hex.slice(2, 4), 16);
  const b = parseInt(hex.slice(4, 6), 16);
  return 0.299 * r + 0.587 * g + 0.114 * b > 230;
}

const swatchOf = (value: string, title: string, css: string): Swatch => ({
  value,
  title,
  css,
  light: isLightHex(css.replace('#', '')),
});

/** The theme colors of the document, or Office's when the document has no complete theme. */
function themeHexesOf(editor: ReturnType<ReturnType<typeof injectDocxEditor>>): readonly string[] {
  const entries = editor?.getDocumentThemeColors() ?? [];
  // A theme color comes from the file, so it paints only as six hex digits.
  return entries.length === THEME_COLUMN_KEYS.length &&
    entries.every((entry) => HEX_VALUE.test(entry.hex))
    ? entries.map((entry) => entry.hex.toUpperCase())
    : DEFAULT_THEME_HEXES;
}

const SWATCH_BUTTON = `<button
    type="button"
    class="docx-toolbar__swatch"
    [style.background-color]="swatch.css"
    [attr.aria-label]="swatch.title"
    [attr.aria-pressed]="swatch.value === current()"
    [attr.title]="swatch.title"
    [attr.data-value]="swatch.value"
    [attr.data-selected]="swatch.value === current() ? '' : null"
    [attr.data-light]="swatch.light ? '' : null"
    (mousedown)="guard($event)"
    (click)="apply.emit(swatch.value)"
  ></button>`;

/**
 * The body of a hex color picker: the theme colors with their tints and shades, the standard
 * colors, and a custom hex field. The font color and the table border color use it.
 *
 * @internal
 */
@Component({
  selector: 'docx-hex-color-picker',
  changeDetection: ChangeDetectionStrategy.OnPush,
  host: { style: 'display: contents' },
  template: `<div class="docx-toolbar__swatch-section">
      <div class="docx-toolbar__swatch-heading">{{ label()('colorPicker.themeColors') }}</div>
      <div class="docx-toolbar__swatch-grid docx-toolbar__swatch-grid--theme" role="group">
        @for (swatch of theme(); track $index) {
          ${SWATCH_BUTTON}
        }
      </div>
    </div>
    <div class="docx-toolbar__swatch-section">
      <div class="docx-toolbar__swatch-heading">{{ label()('colorPicker.standardColors') }}</div>
      <div class="docx-toolbar__swatch-grid" role="group">
        @for (swatch of standard(); track swatch.value) {
          ${SWATCH_BUTTON}
        }
      </div>
    </div>
    <div class="docx-toolbar__swatch-section">
      <div class="docx-toolbar__swatch-heading">{{ label()('colorPicker.customColor') }}</div>
      <div class="docx-toolbar__swatch-custom">
        <span class="docx-toolbar__swatch-hash" aria-hidden="true">#</span>
        <input
          #hex
          type="text"
          class="docx-toolbar__swatch-hex"
          maxlength="6"
          spellcheck="false"
          [value]="draft()"
          [attr.aria-label]="label()('colorPicker.customColor')"
          (input)="onDraft(hex)"
          (keydown.enter)="valid() && apply.emit(draft().toUpperCase())"
        />
        <button
          type="button"
          class="docx-toolbar__swatch-apply"
          [disabled]="!valid()"
          (mousedown)="guard($event)"
          (click)="apply.emit(draft().toUpperCase())"
        >
          {{ label()('colorPicker.apply') }}
        </button>
      </div>
    </div>`,
})
export class DocxHexColorPickerComponent {
  /** The color at the selection, as six hex digits, or `null`. */
  readonly current = input<string | null>(null);
  /** Emits the picked color as six upper-case hex digits. */
  readonly apply = output<string>();

  private readonly editor = injectDocxEditor();
  private readonly hexes = computed(() => themeHexesOf(this.editor()));

  /** @internal */
  protected readonly label = injectToolbarLabel();
  /** @internal */
  protected readonly draft = signal('');
  /** @internal */
  protected readonly valid = computed(() => HEX_VALUE.test(this.draft()));
  /** @internal */
  protected readonly guard: (event: Event) => void = guardToolbarMousedown;
  /** @internal */
  protected readonly theme = computed(() => {
    const label = this.label();
    const hexes = this.hexes();
    const base = hexes.map((hex, column) =>
      swatchOf(hex, label(THEME_COLUMN_KEYS[column]!), `#${hex.toLowerCase()}`)
    );
    const variants = colorPickerGrid(hexes)
      .theme.slice(1)
      .flatMap((row) =>
        row.map((swatch, column) => {
          const hex = swatch.hex.slice(1).toUpperCase();
          return swatchOf(
            hex,
            `${label(THEME_COLUMN_KEYS[column]!)}, ${label(variantLabelKey(swatch.pick))}`,
            `#${hex.toLowerCase()}`
          );
        })
      );
    return [...base, ...variants];
  });
  /** @internal */
  protected readonly standard = computed(() => {
    const label = this.label();
    return STANDARD_COLOR_SWATCHES.map((swatch) =>
      swatchOf(swatch.value, label(swatch.labelKey), swatch.css)
    );
  });

  /** @internal */
  protected onDraft(field: HTMLInputElement): void {
    const clean = field.value.replace(/[^0-9A-Fa-f]/g, '').slice(0, 6);
    field.value = clean;
    this.draft.set(clean);
  }
}

/**
 * The highlight swatches.
 *
 * @internal
 */
@Component({
  selector: 'docx-highlight-picker',
  changeDetection: ChangeDetectionStrategy.OnPush,
  host: { style: 'display: contents' },
  template: `<div class="docx-toolbar__swatch-section">
    <div class="docx-toolbar__swatch-heading">{{ label()('colorPicker.highlightColors') }}</div>
    <div class="docx-toolbar__swatch-grid docx-toolbar__swatch-grid--highlight" role="group">
      @for (swatch of swatches(); track swatch.value) {
        ${SWATCH_BUTTON}
      }
    </div>
  </div>`,
})
export class DocxHighlightPickerComponent {
  /** The highlight at the selection, or `null`. */
  readonly current = input<string | null>(null);
  /** Emits the picked highlight name. */
  readonly apply = output<string>();

  /** @internal */
  protected readonly label = injectToolbarLabel();
  /** @internal */
  protected readonly guard: (event: Event) => void = guardToolbarMousedown;
  /** @internal */
  protected readonly swatches = computed(() => {
    const label = this.label();
    return HIGHLIGHT_GRID.map((swatch) =>
      swatchOf(swatch.value, label(swatch.labelKey), swatch.css)
    );
  });
}

/** The configuration of a color split button. @internal */
export interface ColorSplitConfig {
  readonly slot: ChromeSlotId;
  readonly defaultValue: string;
  readonly cssOf: (value: string) => string;
  readonly clear: { readonly value: string; readonly labelKey: string };
  readonly current: (snapshot: EditorSnapshot) => string | null;
}

/**
 * The shared state of a color split button: a main button and a swatch popup.
 *
 * @internal
 */
@Directive()
export abstract class ColorSplitBase {
  private readonly editor = injectDocxEditor();
  private readonly zone = inject(NgZone);
  private readonly host: HTMLElement = inject(ElementRef).nativeElement;
  private readonly lastValue: WritableSignal<string>;

  /** @internal */
  protected readonly command: EditorCommandState;
  /** @internal */
  protected readonly label = injectToolbarLabel();
  /** @internal */
  protected readonly open = signal(false);
  /** @internal */
  protected readonly current: Signal<string | null>;
  /** @internal */
  protected readonly guard: (event: Event) => void = guardToolbarMousedown;
  /** @internal */
  protected readonly paths: readonly string[] | null;
  /** @internal */
  protected readonly text: Signal<string>;
  /** @internal */
  protected readonly bar: Signal<string>;
  /** @internal */
  protected readonly clearValue: string;
  /** @internal */
  protected readonly clearKey: string;

  constructor(config: ColorSplitConfig) {
    this.lastValue = signal(config.defaultValue);
    this.command = injectEditorCommand(config.slot);
    this.current = injectEditorState(config.current);
    this.paths = chromeControlForSlot(config.slot)?.paths ?? null;
    this.text = computed(() =>
      this.label()(chromeControlForSlot(config.slot)?.labelKey ?? config.slot)
    );
    this.bar = computed(() => config.cssOf(this.lastValue()));
    this.clearValue = config.clear.value;
    this.clearKey = config.clear.labelKey;
    const slot = config.slot;
    this.applyValue = (value: string) => {
      this.open.set(false);
      const editor = untracked(this.editor);
      if (!editor) return;
      const command = commandForSlotValue(slot, value);
      this.zone.runOutsideAngular(() => {
        editor.focus();
        if (!command || !editor.can(command).ok) return;
        editor.exec(command);
        if (value !== config.clear.value) this.zone.run(() => this.lastValue.set(value));
      });
    };
    injectPopupDismiss({
      open: this.open,
      root: () => this.host,
      kind: 'picker',
      close: () => this.open.set(false),
    });
  }

  /** Apply a value and close the popup. @internal */
  protected readonly applyValue: (value: string) => void;

  /** Apply the last value that the main button applied. @internal */
  protected applyLast(): void {
    this.applyValue(untracked(this.lastValue));
  }
}

const SPLIT_START = `<button
    type="button"
    class="docx-toolbar__button docx-toolbar__colorsplit-main"
    [disabled]="!command.isEnabled()"
    [attr.data-disabled]="command.isEnabled() ? null : ''"
    [attr.aria-label]="text()"
    [attr.title]="command.disabledReason() ?? text()"
    (mousedown)="guard($event)"
    (click)="applyLast()"
  >
    <ng-content>
      @if (paths) {
        <svg [docxChromeIcon]="paths"></svg>
      }
    </ng-content>
    <span
      class="docx-toolbar__colorsplit-bar"
      aria-hidden="true"
      [style.background-color]="bar()"
    ></span>
  </button>
  <button
    type="button"
    class="docx-toolbar__colorsplit-caret"
    aria-haspopup="dialog"
    [disabled]="!command.isEnabled()"
    [attr.data-disabled]="command.isEnabled() ? null : ''"
    [attr.aria-expanded]="open()"
    [attr.aria-label]="text()"
    [attr.title]="command.disabledReason() ?? text()"
    (mousedown)="guard($event)"
    (click)="open.set(!open())"
  >
    ▾
  </button>
  @if (open()) {
    <div
      class="docx-toolbar__swatch-popup"
      role="dialog"
      [attr.aria-label]="text()"
      (mousedown)="guard($event)"
    >
      <button
        type="button"
        class="docx-toolbar__swatch-clear"
        (mousedown)="guard($event)"
        (click)="applyValue(clearValue)"
      >
        <span
          class="docx-toolbar__swatch-clear-chip"
          [class.docx-toolbar__swatch-clear-chip--none]="clearValue === 'none'"
          aria-hidden="true"
        ></span>
        {{ label()(clearKey) }}
      </button>
`;
const SPLIT_END = `
    </div>
  }`;

/**
 * The font color split button (`text.color`): the main button applies the last color, and the
 * caret opens the theme, standard, and custom colors. Put your own icon in it to replace the
 * registry icon.
 *
 * @example
 * ```html
 * <div docxToolbarFontColor></div>
 * ```
 *
 * @public
 */
@Component({
  selector: 'div[docxToolbarFontColor]',
  changeDetection: ChangeDetectionStrategy.OnPush,
  imports: [DocxChromeIconComponent, DocxHexColorPickerComponent],
  host: { class: 'docx-toolbar__colorsplit', 'data-slot': 'text.color' },
  template:
    SPLIT_START +
    '<docx-hex-color-picker [current]="current()" (apply)="applyValue($event)" />' +
    SPLIT_END,
})
export class DocxToolbarFontColorComponent extends ColorSplitBase {
  constructor() {
    super({
      slot: 'text.color',
      defaultValue: 'FF0000',
      cssOf: (value) => (value === 'auto' ? '#000000' : `#${value}`),
      clear: { value: 'auto', labelKey: 'colorPicker.automatic' },
      current: (snapshot) => {
        const color = snapshot.formatting?.color;
        return color && color.kind === 'hex' ? color.value.toUpperCase() : null;
      },
    });
  }
}

/**
 * The highlight split button (`text.highlight`): the main button applies the last highlight,
 * and the caret opens the highlight colors.
 *
 * @example
 * ```html
 * <div docxToolbarHighlight></div>
 * ```
 *
 * @public
 */
@Component({
  selector: 'div[docxToolbarHighlight]',
  changeDetection: ChangeDetectionStrategy.OnPush,
  imports: [DocxChromeIconComponent, DocxHighlightPickerComponent],
  host: { class: 'docx-toolbar__colorsplit', 'data-slot': 'text.highlight' },
  template:
    SPLIT_START +
    '<docx-highlight-picker [current]="current()" (apply)="applyValue($event)" />' +
    SPLIT_END,
})
export class DocxToolbarHighlightComponent extends ColorSplitBase {
  constructor() {
    super({
      slot: 'text.highlight',
      defaultValue: 'yellow',
      cssOf: (value) => HIGHLIGHT_CSS.get(value) ?? '#ffff00',
      clear: { value: 'none', labelKey: 'colorPicker.noColor' },
      current: (snapshot) => snapshot.formatting?.highlight ?? null,
    });
  }
}
