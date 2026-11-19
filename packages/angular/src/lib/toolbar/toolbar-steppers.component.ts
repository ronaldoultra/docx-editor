import {
  ChangeDetectionStrategy,
  Component,
  ElementRef,
  NgZone,
  computed,
  inject,
  signal,
  untracked,
} from '@angular/core';
import {
  AUTO_ZOOM_MODE,
  FIT_WIDTH_ZOOM_MODE,
  commandForSlotValue,
  isElementNode,
  sameZoomMode,
} from '@docx-editor.dev/core/editor';
import { injectDocxEditor } from '../context';
import { injectEditorCommand, type EditorCommandState } from '../editor-command';
import { injectEditorState } from '../editor-state';
import { injectZoom, type ZoomState } from './picker-state';
import { injectToolbarLabel } from './toolbar-context';
import { injectPopupDismiss, popupId } from './toolbar-popup';
import { guardToolbarMousedown } from './toolbar-registry';

const MIN_HALF_POINTS = 2;
const MAX_HALF_POINTS = 3276;
const FONT_SIZE_PRESETS_PT: readonly number[] = [
  8, 9, 10, 11, 12, 14, 16, 18, 20, 24, 28, 36, 48, 72,
];

function nextPreset(current: number): number {
  for (const preset of FONT_SIZE_PRESETS_PT) if (preset > current) return preset;
  return Math.min(current + 1, MAX_HALF_POINTS / 2);
}

function previousPreset(current: number): number {
  for (let index = FONT_SIZE_PRESETS_PT.length - 1; index >= 0; index -= 1) {
    if (FONT_SIZE_PRESETS_PT[index]! < current) return FONT_SIZE_PRESETS_PT[index]!;
  }
  return Math.max(current - 1, MIN_HALF_POINTS / 2);
}

/** A typed size in half-points, or `null` when the text is not a size that Word accepts. */
function parseTypedSize(text: string): number | null {
  const trimmed = text.trim().replace(/pt$/i, '').trim();
  if (trimmed === '') return null;
  const points = Number(trimmed);
  if (!Number.isFinite(points)) return null;
  const halfPoints = Math.round(points * 2);
  if (halfPoints < MIN_HALF_POINTS || halfPoints > MAX_HALF_POINTS) return null;
  return halfPoints;
}

/**
 * The font size control (`font.size`): a decrease button, an editable size with a list of
 * preset sizes, and an increase button. The steps follow Word's preset sizes.
 *
 * @example
 * ```html
 * <span docxToolbarFontSize></span>
 * ```
 *
 * @public
 */
@Component({
  selector: 'span[docxToolbarFontSize]',
  changeDetection: ChangeDetectionStrategy.OnPush,
  host: {
    class: 'docx-toolbar__font-size',
    'data-slot': 'font.size',
    '[attr.title]': 'command.disabledReason()',
  },
  template: `<span
      class="docx-toolbar__stepper"
      data-slot="font.size"
      role="group"
      [attr.aria-label]="label()('fontSize.label')"
      [attr.title]="label()('fontSize.label')"
    >
      <button
        type="button"
        class="docx-toolbar__stepper-button"
        [disabled]="!canDecrease()"
        [attr.aria-label]="label()('fontSize.decrease')"
        (mousedown)="guard($event)"
        (click)="step('down')"
      >
        −
      </button>
      <input
        #input
        type="text"
        inputmode="decimal"
        class="docx-toolbar__stepper-value docx-toolbar__stepper-value--boxed docx-toolbar__font-size-input"
        role="combobox"
        aria-haspopup="listbox"
        autocomplete="off"
        [value]="shown()"
        [disabled]="!command.isEnabled()"
        [attr.aria-expanded]="open()"
        [attr.aria-label]="label()('fontSize.label')"
        [attr.aria-controls]="open() ? listId : null"
        (input)="onInput($any($event.target).value)"
        (focus)="input.select(); open.set(true)"
        (click)="open.set(true)"
        (keydown)="onKey($event, input)"
        (blur)="commitDraft(); draft.set(null)"
      />
      <button
        type="button"
        class="docx-toolbar__stepper-button"
        [disabled]="!canIncrease()"
        [attr.aria-label]="label()('fontSize.increase')"
        (mousedown)="guard($event)"
        (click)="step('up')"
      >
        +
      </button>
    </span>
    @if (open() && command.isEnabled()) {
      <div
        class="docx-toolbar__menu docx-toolbar__font-size-menu"
        role="listbox"
        tabindex="0"
        [id]="listId"
        [attr.aria-label]="label()('fontSize.listLabel')"
      >
        @for (preset of presets; track preset) {
          <button
            type="button"
            role="option"
            tabindex="-1"
            class="docx-toolbar__menu-item"
            [attr.aria-selected]="selected() === preset"
            [attr.data-selected]="selected() === preset ? '' : null"
            (mousedown)="guard($event)"
            (click)="apply(preset); dismiss(input, true)"
          >
            {{ preset }}
          </button>
        }
      </div>
    }`,
})
export class DocxToolbarFontSizeComponent {
  private readonly editor = injectDocxEditor();
  private readonly zone = inject(NgZone);
  private readonly host: HTMLElement = inject(ElementRef).nativeElement;
  private readonly sizePt = injectEditorState(
    (snapshot) => snapshot.formatting?.fontSizePt ?? null
  );

  /** @internal */
  protected readonly command: EditorCommandState = injectEditorCommand('font.size');
  /** @internal */
  protected readonly label = injectToolbarLabel();
  /** @internal */
  protected readonly open = signal(false);
  /** @internal */
  protected readonly draft = signal<string | null>(null);
  /** @internal */
  protected readonly listId = popupId('font-size-list');
  /** @internal */
  protected readonly presets = FONT_SIZE_PRESETS_PT;
  /** @internal */
  protected readonly guard: (event: Event) => void = guardToolbarMousedown;
  /** @internal */
  protected readonly selected = computed(() => {
    const size = this.sizePt();
    return size === null ? null : Math.round(size * 2) / 2;
  });
  /** @internal */
  protected readonly shown = computed(() => {
    const selected = this.selected();
    return this.draft() ?? (selected === null ? '—' : String(selected));
  });
  private readonly canStep = computed(() => this.command.isEnabled() && this.sizePt() !== null);
  /** @internal */
  protected readonly canDecrease = computed(
    () => this.canStep() && Math.round(this.sizePt()! * 2) > MIN_HALF_POINTS
  );
  /** @internal */
  protected readonly canIncrease = computed(
    () => this.canStep() && Math.round(this.sizePt()! * 2) < MAX_HALF_POINTS
  );

  constructor() {
    injectPopupDismiss({
      open: this.open,
      root: () => this.host,
      kind: 'picker',
      // The size input owns its draft and its Escape.
      skip: (event) =>
        isElementNode(event.target) &&
        this.host.contains(event.target) &&
        event.target.matches('input[role="combobox"]'),
      close: () => {
        this.open.set(false);
        this.draft.set(null);
      },
    });
  }

  /** @internal */
  protected step(direction: 'up' | 'down'): void {
    const size = untracked(this.sizePt) ?? 0;
    this.apply(direction === 'up' ? nextPreset(size) : previousPreset(size));
  }

  /** Apply a size in points. @internal */
  protected apply(points: number): void {
    this.applyHalfPoints(
      Math.min(MAX_HALF_POINTS, Math.max(MIN_HALF_POINTS, Math.round(points * 2)))
    );
  }

  /** @internal */
  protected onInput(value: string): void {
    this.draft.set(value);
    this.open.set(true);
  }

  /** @internal */
  protected onKey(event: KeyboardEvent, input: HTMLInputElement): void {
    if (event.key === 'Enter') {
      event.preventDefault();
      this.commitDraft();
      this.dismiss(input, true);
    } else if (event.key === 'Escape') {
      event.preventDefault();
      this.dismiss(input, true);
    } else if (event.key === 'ArrowDown' || event.key === 'ArrowUp') {
      event.preventDefault();
      this.open.set(true);
      if (untracked(this.sizePt) !== null) this.step(event.key === 'ArrowUp' ? 'up' : 'down');
    }
  }

  /** @internal */
  protected commitDraft(): void {
    const draft = untracked(this.draft);
    if (draft === null) return;
    const halfPoints = parseTypedSize(draft);
    if (halfPoints !== null) this.applyHalfPoints(halfPoints);
  }

  /** @internal */
  protected dismiss(input: HTMLInputElement, refocus: boolean): void {
    this.open.set(false);
    this.draft.set(null);
    input.blur();
    if (refocus) this.zone.runOutsideAngular(() => untracked(this.editor)?.focus());
  }

  private applyHalfPoints(halfPoints: number): void {
    const editor = untracked(this.editor);
    if (!editor) return;
    const command = commandForSlotValue('font.size', halfPoints);
    if (!command) return;
    this.zone.runOutsideAngular(() => {
      if (editor.can(command).ok) editor.exec(command);
    });
  }
}

/**
 * The zoom control (`zoom.level`): zoom out, a menu with the zoom steps and the fit modes, and
 * zoom in.
 *
 * @example
 * ```html
 * <span docxToolbarZoom></span>
 * ```
 *
 * @public
 */
@Component({
  selector: 'span[docxToolbarZoom]',
  changeDetection: ChangeDetectionStrategy.OnPush,
  host: {
    class: 'docx-toolbar__zoom',
    'data-slot': 'zoom.level',
  },
  template: `<span
      class="docx-toolbar__stepper"
      role="group"
      [attr.aria-label]="label()('zoom.zoomLevel')"
      [attr.title]="label()('zoom.zoomLevel')"
    >
      <button
        type="button"
        class="docx-toolbar__stepper-button"
        [disabled]="!zoom.canZoomOut()"
        [attr.aria-label]="label()('zoom.zoomOut')"
        (mousedown)="guard($event)"
        (click)="zoom.zoomOut()"
      >
        −
      </button>
      <button
        type="button"
        class="docx-toolbar__stepper-value docx-toolbar__stepper-value--menu"
        aria-haspopup="listbox"
        [disabled]="!hasEditor()"
        [attr.aria-expanded]="open()"
        [attr.aria-label]="label()('zoom.zoomLevel') + ': ' + display()"
        (mousedown)="guard($event)"
        (click)="open.set(!open())"
      >
        {{ display() }}<span class="docx-toolbar__picker-caret" aria-hidden="true">▾</span>
      </button>
      <button
        type="button"
        class="docx-toolbar__stepper-button"
        [disabled]="!zoom.canZoomIn()"
        [attr.aria-label]="label()('zoom.zoomIn')"
        (mousedown)="guard($event)"
        (click)="zoom.zoomIn()"
      >
        +
      </button>
    </span>
    @if (open()) {
      <div
        class="docx-toolbar__menu docx-toolbar__zoom-menu"
        role="listbox"
        tabindex="0"
        [attr.aria-label]="label()('zoom.zoomLevel')"
      >
        <button
          type="button"
          role="option"
          tabindex="-1"
          class="docx-toolbar__menu-item"
          [attr.aria-selected]="autoSelected()"
          [attr.data-selected]="autoSelected() ? '' : null"
          (mousedown)="guard($event)"
          (click)="pick('auto')"
        >
          {{ label()('zoom.automatic') }}
        </button>
        <button
          type="button"
          role="option"
          tabindex="-1"
          class="docx-toolbar__menu-item"
          [attr.aria-selected]="fitWidthSelected()"
          [attr.data-selected]="fitWidthSelected() ? '' : null"
          (mousedown)="guard($event)"
          (click)="pick('fit')"
        >
          {{ label()('zoom.fitWidth') }}
        </button>
        <hr class="docx-toolbar__menu-separator" role="presentation" />
        @for (level of zoom.levels; track level) {
          <button
            type="button"
            role="option"
            tabindex="-1"
            class="docx-toolbar__menu-item"
            [attr.aria-selected]="isLevel(level)"
            [attr.data-selected]="isLevel(level) ? '' : null"
            (mousedown)="guard($event)"
            (click)="pick(level)"
          >
            {{ percent(level) }}
          </button>
        }
      </div>
    }`,
})
export class DocxToolbarZoomComponent {
  private readonly editor = injectDocxEditor();
  private readonly host: HTMLElement = inject(ElementRef).nativeElement;

  /** @internal */
  protected readonly zoom: ZoomState = injectZoom();
  /** @internal */
  protected readonly label = injectToolbarLabel();
  /** @internal */
  protected readonly open = signal(false);
  /** @internal */
  protected readonly guard: (event: Event) => void = guardToolbarMousedown;
  /** @internal */
  protected readonly hasEditor = computed(() => this.editor() !== null);
  /** @internal */
  protected readonly display = computed(() => this.percent(this.zoom.zoom()));
  /** @internal */
  protected readonly autoSelected = computed(() => sameZoomMode(this.zoom.mode(), AUTO_ZOOM_MODE));
  /** @internal */
  protected readonly fitWidthSelected = computed(() =>
    sameZoomMode(this.zoom.mode(), FIT_WIDTH_ZOOM_MODE)
  );

  constructor() {
    injectPopupDismiss({
      open: this.open,
      root: () => this.host,
      kind: 'picker',
      close: () => this.open.set(false),
    });
  }

  /** @internal */
  protected percent(level: number): string {
    return `${Math.round(level * 100)}%`;
  }

  /** @internal */
  protected isLevel(level: number): boolean {
    return !this.zoom.isFit() && Math.abs(level - this.zoom.zoom()) < 0.001;
  }

  /** @internal */
  protected pick(choice: 'auto' | 'fit' | number): void {
    this.host.querySelector<HTMLElement>('[aria-haspopup]')?.focus();
    this.open.set(false);
    if (choice === 'auto') this.zoom.auto();
    else if (choice === 'fit') this.zoom.fitToWidth();
    else this.zoom.setZoom(choice);
  }
}
