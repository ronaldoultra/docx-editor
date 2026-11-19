import { NgTemplateOutlet } from '@angular/common';
import {
  ChangeDetectionStrategy,
  Component,
  Directive,
  ElementRef,
  InjectionToken,
  afterRenderEffect,
  computed,
  inject,
  input,
  signal,
  untracked,
  viewChild,
  type Signal,
  type TemplateRef,
  type WritableSignal,
} from '@angular/core';
import {
  TABLE_BORDER_STYLE_OPTIONS,
  TABLE_BORDER_TARGET_OPTIONS,
  TABLE_BORDER_WIDTH_OPTIONS,
  editorInstanceScope,
  tableChromeIconPaths,
  tableChromeLabelKeyForTarget,
  type ChromeSlotId,
  type TableChromeSlotId,
} from '@docx-editor.dev/core/editor';
import { injectEditorCommand, type EditorCommandState } from '../editor-command';
import { DocxChromeIconComponent } from './chrome-icon.component';
import { DocxTableChromeState } from './table-chrome-state';
import { DocxHexColorPickerComponent } from './toolbar-colors.component';
import { DocxToolbarButtonComponent } from './toolbar-button.component';
import { injectToolbarLabel } from './toolbar-context';
import { injectPopupDismiss, popupId } from './toolbar-popup';
import { chromeControlForSlot, guardToolbarMousedown } from './toolbar-registry';

/**
 * Opens the editor dialogs that a toolbar control needs, such as the Split cells dialog. The
 * trigger receives the focus back when the dialog closes. Without a provider, the controls that
 * open a dialog do not render.
 *
 * @public
 */
export const DOCX_TOOLBAR_DIALOGS = new InjectionToken<{
  readonly open: (dialog: 'splitCell', trigger: HTMLElement | null) => void;
}>('DOCX_TOOLBAR_DIALOGS');

const MENU_ITEMS = '[role="menuitem"], [role="menuitemradio"], [role="menuitemcheckbox"]';

function injectTableState(): DocxTableChromeState {
  const state = inject(DocxTableChromeState, { optional: true });
  if (!state) throw new Error('docx-editor: table controls must render inside a toolbar.');
  return state;
}

/** Give the focus back to the painted pages after a table color applies. */
function focusPages(from: HTMLElement): void {
  const root = editorInstanceScope(from) ?? from.ownerDocument.body;
  root.querySelector<HTMLElement>('.docx-pages')?.focus();
}

/**
 * The trigger, the open state, and the keyboard of one table control.
 *
 * @internal
 */
@Directive()
export abstract class TableControlBase {
  protected readonly host: HTMLElement = inject(ElementRef).nativeElement;
  protected readonly chrome = injectTableState();
  /** @internal */
  protected readonly label = injectToolbarLabel();
  /** @internal */
  protected readonly open = signal(false);
  /** @internal */
  protected readonly guard: (event: Event) => void = guardToolbarMousedown;
  /** @internal */
  protected readonly reasonId = popupId('table-chrome-reason');
  /** @internal */
  protected readonly state: Signal<{ enabled: boolean; disabledReason: string | null }>;
  /** @internal */
  protected readonly text: Signal<string>;
  private readonly panel = viewChild<ElementRef<HTMLElement>>('panel');
  private readonly trigger = viewChild<ElementRef<HTMLElement>>('trigger');

  constructor(
    protected readonly slot: TableChromeSlotId,
    labelKey: string,
    popup: 'menu' | 'dialog'
  ) {
    this.state = this.chrome.slot(slot);
    this.text = computed(() => this.label()(chromeControlForSlot(slot)?.labelKey ?? labelKey));
    injectPopupDismiss({
      open: computed(() => this.open() && this.chrome.visible()),
      root: () => this.host,
      close: (fromInside) => {
        this.open.set(false);
        if (fromInside) this.trigger()?.nativeElement.focus();
      },
    });
    // An open popup takes the focus: a menu on its checked row, a dialog on its first button.
    afterRenderEffect(() => {
      const panel = this.panel()?.nativeElement;
      if (!panel) return;
      const rows = this.rows(panel, popup);
      const checked = rows.find(
        (row) => row.hasAttribute('data-selected') || row.hasAttribute('data-active')
      );
      (checked ?? rows[0])?.focus();
    });
  }

  /** @internal */
  protected toggle(): void {
    if (this.state().enabled) this.open.set(!this.open());
  }

  /** The keys of an open popup: arrows, Home, End, Enter, Space, and Escape. @internal */
  protected onPanelKey(event: KeyboardEvent, popup: 'menu' | 'dialog'): void {
    const panel = this.panel()?.nativeElement;
    if (!panel || event.defaultPrevented || event.isComposing) return;
    const rows = this.rows(panel, popup);
    const active = this.host.ownerDocument.activeElement;
    const at = rows.findIndex((row) => row === active);
    const focus = (index: number) => {
      event.preventDefault();
      rows[(index + rows.length) % rows.length]?.focus();
    };
    if (event.key === 'Escape') {
      event.preventDefault();
      this.open.set(false);
      this.trigger()?.nativeElement.focus();
    } else if (popup === 'menu' && event.key === 'ArrowDown') focus(at + 1);
    else if (popup === 'menu' && event.key === 'ArrowUp') focus(at < 0 ? -1 : at - 1);
    else if (popup === 'menu' && event.key === 'Home') focus(0);
    else if (popup === 'menu' && event.key === 'End') focus(rows.length - 1);
    else if ((event.key === 'Enter' || event.key === ' ') && at >= 0) {
      event.preventDefault();
      rows[at]!.click();
    }
  }

  /** @internal */
  protected pick(value: unknown): void {
    this.open.set(false);
    this.chrome.apply(this.slot, value);
  }

  private rows(panel: HTMLElement, popup: 'menu' | 'dialog'): HTMLElement[] {
    const selector = popup === 'menu' ? MENU_ITEMS : 'button:not([aria-disabled="true"])';
    return [...panel.querySelectorAll<HTMLElement>(selector)].filter(
      (row) => row.closest(popup === 'menu' ? '[role="menu"]' : '[role="dialog"]') === panel
    );
  }
}

/** The trigger of a table menu: focusable when disabled, with the reason announced. */
const MENU_TRIGGER = `<button
    #trigger
    type="button"
    class="docx-toolbar__button docx-table-chrome__trigger"
    aria-haspopup="menu"
    [attr.aria-expanded]="open()"
    [attr.aria-label]="text()"
    [attr.title]="state().disabledReason"
    [attr.aria-disabled]="state().enabled ? null : true"
    [attr.data-disabled]="state().enabled ? null : ''"
    [attr.aria-describedby]="state().enabled || !state().disabledReason ? null : reasonId"
    (mousedown)="guard($event)"
    (click)="toggle()"
  >
    @if (paths; as icon) {
      <svg [docxChromeIcon]="icon"></svg>
    }
    <span class="docx-toolbar__picker-caret" aria-hidden="true">▾</span>
    @if (!state().enabled && state().disabledReason) {
      <span class="docx-editor-sr-only" [id]="reasonId">{{ state().disabledReason }}</span>
    }
  </button>`;

const MENU_PANEL_START = `@if (open() && state().enabled) {
    <div
      #panel
      role="menu"
      class="docx-table-chrome__panel"
      [attr.aria-label]="text()"
      (mousedown)="guard($event)"
      (keydown)="onPanelKey($event, 'menu')"
    >`;
const MENU_PANEL_END = `</div>
  }`;

/**
 * The border target menu (`table.borderTarget`): which borders of the selected cells the next
 * border command changes. It renders only while the selection is in a table.
 *
 * @public
 */
@Component({
  selector: 'div[docxToolbarTableBorderTarget]',
  changeDetection: ChangeDetectionStrategy.OnPush,
  imports: [DocxChromeIconComponent],
  host: {
    style: 'position: relative; display: inline-flex; align-items: center; vertical-align: middle',
    '[style.display]': "chrome.visible() ? null : 'none'",
    class: 'docx-table-chrome docx-table-chrome--target',
    'data-slot': 'table.borderTarget',
  },
  template:
    '@if (chrome.visible()) {' +
    MENU_TRIGGER +
    MENU_PANEL_START +
    `<div class="docx-table-chrome__target-grid">
      @for (option of options(); track option.value) {
        <button
          type="button"
          role="menuitemradio"
          class="docx-table-chrome__target-btn"
          [attr.aria-checked]="option.selected"
          [attr.data-active]="option.selected ? '' : null"
          [attr.data-value]="option.value"
          [attr.aria-label]="option.text"
          [attr.title]="option.text"
          (mousedown)="guard($event)"
          (click)="pick(option.value)"
        >
          <svg [docxChromeIcon]="option.paths"></svg>
        </button>
      }
    </div>` +
    MENU_PANEL_END +
    '}',
})
export class DocxToolbarTableBorderTargetComponent extends TableControlBase {
  /** @internal */
  protected readonly paths = chromeControlForSlot('table.borderTarget')?.paths ?? null;
  /** @internal */
  protected readonly options = computed(() => {
    const label = this.label();
    const target = this.chrome.draft().activeTarget;
    return TABLE_BORDER_TARGET_OPTIONS.map((option) => ({
      value: option.value,
      paths: tableChromeIconPaths(option.icon),
      text: label(option.labelKey),
      selected: option.value !== 'none' && option.value === target,
    }));
  });

  constructor() {
    super('table.borderTarget', 'table.borders.tooltip', 'menu');
  }
}

/**
 * The border style menu (`table.borderStyle`).
 *
 * @public
 */
@Component({
  selector: 'div[docxToolbarTableBorderStyle]',
  changeDetection: ChangeDetectionStrategy.OnPush,
  imports: [DocxChromeIconComponent],
  host: {
    style: 'position: relative; display: inline-flex; align-items: center; vertical-align: middle',
    '[style.display]': "chrome.visible() ? null : 'none'",
    class: 'docx-table-chrome docx-table-chrome--style',
    'data-slot': 'table.borderStyle',
  },
  template:
    '@if (chrome.visible()) {' +
    MENU_TRIGGER +
    MENU_PANEL_START +
    `@for (option of options(); track option.value) {
      <button
        type="button"
        role="menuitemradio"
        class="docx-table-chrome__width-row"
        [attr.aria-checked]="option.selected"
        [attr.data-selected]="option.selected ? '' : null"
        (mousedown)="guard($event)"
        (click)="pick(option.value)"
      >
        <span [class]="'docx-table-line ' + option.previewClass" aria-hidden="true"></span>
        <span>{{ option.text }}</span>
      </button>
    }` +
    MENU_PANEL_END +
    '}',
})
export class DocxToolbarTableBorderStyleComponent extends TableControlBase {
  /** @internal */
  protected readonly paths = chromeControlForSlot('table.borderStyle')?.paths ?? null;
  /** @internal */
  protected readonly options = computed(() => {
    const label = this.label();
    const style = this.chrome.draft().spec.style;
    return TABLE_BORDER_STYLE_OPTIONS.map((option) => ({
      value: option.value,
      previewClass: option.previewClass,
      text: label(option.labelKey),
      selected: option.value === style,
    }));
  });

  constructor() {
    super('table.borderStyle', 'table.borders.styleAriaLabel', 'menu');
  }
}

/**
 * The border width menu (`table.borderWidth`).
 *
 * @public
 */
@Component({
  selector: 'div[docxToolbarTableBorderWidth]',
  changeDetection: ChangeDetectionStrategy.OnPush,
  imports: [DocxChromeIconComponent],
  host: {
    style: 'position: relative; display: inline-flex; align-items: center; vertical-align: middle',
    '[style.display]': "chrome.visible() ? null : 'none'",
    class: 'docx-table-chrome docx-table-chrome--width',
    'data-slot': 'table.borderWidth',
  },
  template:
    '@if (chrome.visible()) {' +
    MENU_TRIGGER +
    MENU_PANEL_START +
    `@for (option of options(); track option.size) {
      <button
        type="button"
        role="menuitemradio"
        class="docx-table-chrome__width-row"
        [attr.aria-checked]="option.selected"
        [attr.data-selected]="option.selected ? '' : null"
        (mousedown)="guard($event)"
        (click)="pick(option.size)"
      >
        <span
          class="docx-table-line docx-table-line--single"
          aria-hidden="true"
          [style.height.px]="option.thickness"
        ></span>
        <span>{{ option.text }}</span>
      </button>
    }` +
    MENU_PANEL_END +
    '}',
})
export class DocxToolbarTableBorderWidthComponent extends TableControlBase {
  /** @internal */
  protected readonly paths = chromeControlForSlot('table.borderWidth')?.paths ?? null;
  /** @internal */
  protected readonly options = computed(() => {
    const label = this.label();
    const size = this.chrome.draft().spec.size;
    return TABLE_BORDER_WIDTH_OPTIONS.map((option) => ({
      size: option.size,
      thickness: option.previewThickness,
      text: label(option.labelKey),
      selected: option.size === size,
    }));
  });

  constructor() {
    super('table.borderWidth', 'table.borderWidth', 'menu');
  }
}

const HEX = /^[0-9a-fA-F]{6}$/;

/**
 * A table color split button: the main button applies the bar color, the caret picks one.
 *
 * @internal
 */
@Directive()
export abstract class TableColorBase extends TableControlBase {
  /** @internal */
  protected readonly lastHex: WritableSignal<string>;
  /** @internal */
  protected readonly barHex: Signal<string>;
  /** @internal */
  protected readonly caretText: Signal<string>;

  constructor(slot: 'table.borderColor' | 'table.cellFill', labelKey: string, defaultHex: string) {
    super(slot, labelKey, 'dialog');
    this.lastHex = signal(defaultHex);
    this.caretText = computed(() => this.label()(labelKey));
    this.barHex = computed(() => {
      const color = this.chrome.draft().spec.color;
      const raw =
        slot === 'table.borderColor' && color.kind === 'hex' ? color.value : this.lastHex();
      return HEX.test(raw) ? raw.toUpperCase() : defaultHex;
    });
  }

  /** @internal */
  protected applyMain(): void {
    if (this.state().enabled)
      this.chrome.apply(this.slot, { kind: 'hex', value: untracked(this.barHex) });
  }

  /** @internal */
  protected applyHex(hex: string): void {
    this.open.set(false);
    this.lastHex.set(hex);
    this.chrome.apply(this.slot, { kind: 'hex', value: hex });
    focusPages(this.host);
  }

  /** @internal */
  protected clearFill(): void {
    this.open.set(false);
    this.chrome.apply(this.slot, null);
    focusPages(this.host);
  }
}

const COLOR_START = `@if (chrome.visible()) {
    <button
      type="button"
      class="docx-toolbar__button docx-toolbar__colorsplit-main"
      [attr.aria-label]="text()"
      [attr.title]="state().disabledReason"
      [attr.aria-disabled]="state().enabled ? null : true"
      [attr.data-disabled]="state().enabled ? null : ''"
      [attr.aria-describedby]="state().enabled || !state().disabledReason ? null : reasonId"
      (mousedown)="guard($event)"
      (click)="applyMain()"
    >
      @if (paths; as icon) {
        <svg [docxChromeIcon]="icon"></svg>
      }
      <span
        class="docx-toolbar__colorsplit-bar"
        aria-hidden="true"
        [style.background-color]="'#' + barHex().toLowerCase()"
      ></span>
      @if (!state().enabled && state().disabledReason) {
        <span class="docx-editor-sr-only" [id]="reasonId">{{ state().disabledReason }}</span>
      }
    </button>
    <button
      #trigger
      type="button"
      class="docx-toolbar__colorsplit-caret"
      aria-haspopup="dialog"
      [attr.aria-expanded]="open()"
      [attr.aria-label]="caretText()"
      [attr.title]="state().disabledReason"
      [attr.aria-disabled]="state().enabled ? null : true"
      [attr.data-disabled]="state().enabled ? null : ''"
      (mousedown)="guard($event)"
      (click)="toggle()"
    >
      ▾
    </button>
    @if (open() && state().enabled) {
      <div
        #panel
        role="dialog"
        class="docx-toolbar__swatch-popup docx-table-chrome__panel"
        [attr.aria-label]="caretText()"
        (mousedown)="guard($event)"
        (keydown)="onPanelKey($event, 'dialog')"
      >
`;
const COLOR_END = `        <docx-hex-color-picker [current]="barHex()" (apply)="applyHex($event)" />
      </div>
    }
  }`;

/**
 * The border color split button (`table.borderColor`).
 *
 * @public
 */
@Component({
  selector: 'div[docxToolbarTableBorderColor]',
  changeDetection: ChangeDetectionStrategy.OnPush,
  imports: [DocxChromeIconComponent, DocxHexColorPickerComponent],
  host: {
    style: 'position: relative; display: inline-flex; align-items: center; vertical-align: middle',
    '[style.display]': "chrome.visible() ? null : 'none'",
    class: 'docx-toolbar__colorsplit docx-table-chrome',
    'data-slot': 'table.borderColor',
  },
  template: COLOR_START + COLOR_END,
})
export class DocxToolbarTableBorderColorComponent extends TableColorBase {
  /** @internal */
  protected readonly paths = chromeControlForSlot('table.borderColor')?.paths ?? null;

  constructor() {
    super('table.borderColor', 'table.borderColor', '000000');
  }
}

/**
 * The cell fill split button (`table.cellFill`), with a row that clears the fill.
 *
 * @public
 */
@Component({
  selector: 'div[docxToolbarTableCellFill]',
  changeDetection: ChangeDetectionStrategy.OnPush,
  imports: [DocxChromeIconComponent, DocxHexColorPickerComponent],
  host: {
    style: 'position: relative; display: inline-flex; align-items: center; vertical-align: middle',
    '[style.display]': "chrome.visible() ? null : 'none'",
    class: 'docx-toolbar__colorsplit docx-table-chrome',
    'data-slot': 'table.cellFill',
  },
  template:
    COLOR_START +
    `<button
      type="button"
      class="docx-toolbar__swatch-clear"
      (mousedown)="guard($event)"
      (click)="clearFill()"
    >
      <span
        class="docx-toolbar__swatch-clear-chip docx-toolbar__swatch-clear-chip--none"
        aria-hidden="true"
      ></span>
      {{ label()('table.clearCellFill') }}
    </button>` +
    COLOR_END,
})
export class DocxToolbarTableCellFillComponent extends TableColorBase {
  /** @internal */
  protected readonly paths = chromeControlForSlot('table.cellFill')?.paths ?? null;

  constructor() {
    super('table.cellFill', 'table.cellFillColor', 'FFFF00');
  }
}

/**
 * The Split cells button (`table.splitCell`). It opens the Split cells dialog through
 * {@link DOCX_TOOLBAR_DIALOGS}, and it does not render without a provider.
 *
 * @public
 */
@Component({
  selector: 'span[docxToolbarTableSplitCell]',
  changeDetection: ChangeDetectionStrategy.OnPush,
  imports: [DocxChromeIconComponent],
  host: { style: 'display: contents', 'data-slot-host': 'table.splitCell' },
  template: `@if (dialogs) {
    <button
      type="button"
      class="docx-toolbar__button"
      data-slot="table.splitCell"
      aria-haspopup="dialog"
      [disabled]="!command.isEnabled()"
      [attr.data-disabled]="command.isEnabled() ? null : ''"
      [attr.aria-label]="text()"
      [attr.title]="command.disabledReason() ?? text()"
      (mousedown)="guard($event)"
      (click)="dialogs.open('splitCell', $any($event.currentTarget))"
    >
      @if (paths; as icon) {
        <svg [docxChromeIcon]="icon"></svg>
      }
    </button>
  }`,
})
export class DocxToolbarTableSplitCellComponent {
  private readonly label = injectToolbarLabel();

  /** @internal */
  protected readonly dialogs = inject(DOCX_TOOLBAR_DIALOGS, { optional: true });
  /** @internal */
  protected readonly command: EditorCommandState = injectEditorCommand('table.splitCell');
  /** @internal */
  protected readonly guard: (event: Event) => void = guardToolbarMousedown;
  /** @internal */
  protected readonly paths = chromeControlForSlot('table.splitCell')?.paths ?? null;
  /** @internal */
  protected readonly text = computed(() =>
    this.label()(chromeControlForSlot('table.splitCell')?.labelKey ?? 'table.splitCell')
  );
}

/** The parts of the contextual table group, in bar order. */
export const TABLE_GROUP_SLOTS: readonly ChromeSlotId[] = [
  'table.borderTarget',
  'table.borderColor',
  'table.borderStyle',
  'table.borderWidth',
  'table.cellFill',
  'table.mergeCells',
  'table.splitCell',
];

/**
 * The contextual table group: the border, fill, merge, and split controls, with the host's slot
 * templates in place of their defaults.
 *
 * @internal
 */
@Component({
  selector: 'docx-toolbar-table-chrome',
  changeDetection: ChangeDetectionStrategy.OnPush,
  imports: [
    NgTemplateOutlet,
    DocxToolbarButtonComponent,
    DocxToolbarTableBorderTargetComponent,
    DocxToolbarTableBorderColorComponent,
    DocxToolbarTableBorderStyleComponent,
    DocxToolbarTableBorderWidthComponent,
    DocxToolbarTableCellFillComponent,
    DocxToolbarTableSplitCellComponent,
  ],
  host: { style: 'display: contents' },
  template: `@if (chrome.visible()) {
      @for (slot of slots(); track slot) {
        <ng-container [ngTemplateOutlet]="part" [ngTemplateOutletContext]="{ $implicit: slot }" />
      }
    }
    <ng-template #part let-slot>
      @if (overrides().get(slot); as override) {
        <ng-container [ngTemplateOutlet]="override" />
      } @else {
        @switch (slot) {
          @case ('table.borderTarget') {
            <div docxToolbarTableBorderTarget></div>
          }
          @case ('table.borderColor') {
            <div docxToolbarTableBorderColor></div>
          }
          @case ('table.borderStyle') {
            <div docxToolbarTableBorderStyle></div>
          }
          @case ('table.borderWidth') {
            <div docxToolbarTableBorderWidth></div>
          }
          @case ('table.cellFill') {
            <div docxToolbarTableCellFill></div>
          }
          @case ('table.mergeCells') {
            <button docxToolbarButton="table.mergeCells"></button>
          }
          @case ('table.splitCell') {
            <span docxToolbarTableSplitCell></span>
          }
        }
      }
    </ng-template>`,
})
export class DocxToolbarTableChromeComponent {
  /** The host's templates for table slots. */
  readonly overrides = input<ReadonlyMap<string, TemplateRef<unknown>>>(new Map());
  /** The slots to render: all of them, or one row of the "⋯" panel. */
  readonly slots = input<readonly ChromeSlotId[]>(TABLE_GROUP_SLOTS);

  /** @internal */
  protected readonly chrome = injectTableState();
}

/**
 * The label of the border target in force, such as "All borders". Call it in an injection
 * context inside a toolbar.
 *
 * @public
 */
export function injectTableBorderTargetLabel(): Signal<string> {
  const chrome = injectTableState();
  const label = injectToolbarLabel();
  return computed(() => label()(tableChromeLabelKeyForTarget(chrome.draft().activeTarget)));
}
