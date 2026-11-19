import {
  ChangeDetectionStrategy,
  Component,
  ElementRef,
  NgZone,
  computed,
  inject,
  signal,
} from '@angular/core';
import { injectDocxEditor } from '../context';
import {
  injectFontFamily,
  injectParagraphStyle,
  type FontFamilyState,
  type ParagraphStyleOption,
  type ParagraphStyleState,
} from './picker-state';
import { injectToolbarLabel } from './toolbar-context';
import { injectPopupDismiss } from './toolbar-popup';
import { chromeControlForSlot, guardToolbarMousedown } from './toolbar-registry';

const FONT_CATEGORY: ReadonlyMap<string, 'sansSerif' | 'serif' | 'monospace'> = new Map([
  ['arial', 'sansSerif'],
  ['calibri', 'sansSerif'],
  ['helvetica', 'sansSerif'],
  ['verdana', 'sansSerif'],
  ['open sans', 'sansSerif'],
  ['roboto', 'sansSerif'],
  ['times new roman', 'serif'],
  ['georgia', 'serif'],
  ['cambria', 'serif'],
  ['garamond', 'serif'],
  ['courier new', 'monospace'],
  ['consolas', 'monospace'],
]);

const FONT_GROUPS = [
  { category: 'sansSerif', labelKey: 'font.sansSerif' },
  { category: 'serif', labelKey: 'font.serif' },
  { category: 'monospace', labelKey: 'font.monospace' },
  { category: 'other', labelKey: null },
] as const;

/** Close a picker and give the focus back to the pages. */
function focusPages(zone: NgZone, editor: ReturnType<typeof injectDocxEditor>): void {
  zone.runOutsideAngular(() => editor()?.focus());
}

/**
 * The font family picker (`font.family`): a trigger that shows the family at the selection, and
 * a list of the engine's families, in groups, with a search field.
 *
 * @example
 * ```html
 * <div docxToolbarFontFamily></div>
 * ```
 *
 * @public
 */
@Component({
  selector: 'div[docxToolbarFontFamily]',
  changeDetection: ChangeDetectionStrategy.OnPush,
  host: {
    class: 'docx-toolbar__font-family',
    'data-slot': 'font.family',
  },
  template: `<button
      type="button"
      class="docx-toolbar__font-family-trigger"
      aria-haspopup="listbox"
      [attr.aria-expanded]="open()"
      [attr.aria-label]="text() + ': ' + (state.value() ?? '—')"
      [attr.title]="text()"
      [disabled]="!state.isEnabled()"
      [attr.data-disabled]="state.isEnabled() ? null : ''"
      (mousedown)="guard($event)"
      (click)="open.set(!open())"
    >
      <span>{{ state.value() ?? '—' }}</span>
    </button>
    @if (open()) {
      <div class="docx-toolbar__menu docx-toolbar__font-family-content">
        <input
          type="search"
          class="docx-toolbar__font-search"
          [attr.aria-label]="label()('font.search')"
          [attr.placeholder]="label()('font.search')"
          [value]="query()"
          (input)="query.set($any($event.target).value)"
        />
        <div
          role="listbox"
          tabindex="0"
          [attr.aria-label]="label()('font.selectAriaLabel')"
          class="docx-toolbar__font-options"
        >
          @for (group of groups(); track group.category; let first = $first) {
            @if (!first) {
              <div class="docx-toolbar__menu-separator" aria-hidden="true"></div>
            }
            @if (group.labelKey) {
              <div class="docx-toolbar__menu-label">{{ label()(group.labelKey) }}</div>
            }
            @for (font of group.fonts; track font) {
              <button
                type="button"
                role="option"
                tabindex="-1"
                class="docx-toolbar__font-family-item"
                [attr.aria-selected]="font === state.value()"
                [attr.data-selected]="font === state.value() ? '' : null"
                (mousedown)="guard($event)"
                (click)="choose(font)"
              >
                <span [style.font-family]="font">{{ font }}</span>
                @if (font === state.value()) {
                  <span class="docx-toolbar__menu-check" aria-hidden="true">✓</span>
                }
              </button>
            }
          }
        </div>
      </div>
    }`,
})
export class DocxToolbarFontFamilyComponent {
  private readonly editor = injectDocxEditor();
  private readonly zone = inject(NgZone);
  private readonly host: HTMLElement = inject(ElementRef).nativeElement;

  /** @internal */
  protected readonly state: FontFamilyState = injectFontFamily();
  /** @internal */
  protected readonly label = injectToolbarLabel();
  /** @internal */
  protected readonly open = signal(false);
  /** @internal */
  protected readonly query = signal('');
  /** @internal */
  protected readonly guard: (event: Event) => void = guardToolbarMousedown;
  /** @internal */
  protected readonly text = computed(() =>
    this.label()(chromeControlForSlot('font.family')?.labelKey ?? 'font.family')
  );
  /** @internal */
  protected readonly groups = computed(() => {
    const query = this.query().toLowerCase();
    return FONT_GROUPS.map((group) => ({
      ...group,
      fonts: this.state
        .options()
        .filter(
          (option) =>
            option.toLowerCase().includes(query) &&
            (FONT_CATEGORY.get(option.toLowerCase()) ?? 'other') === group.category
        ),
    })).filter((group) => group.fonts.length > 0);
  });

  constructor() {
    injectPopupDismiss({
      open: this.open,
      root: () => this.host,
      kind: 'picker',
      close: () => this.close(),
    });
  }

  /** @internal */
  protected choose(font: string): void {
    this.state.setValue(font);
    this.close();
    focusPages(this.zone, this.editor);
  }

  private close(): void {
    this.open.set(false);
    this.query.set('');
  }
}

/** The inline preview of a style option. */
function previewStyle(preview: ParagraphStyleOption['preview']): Record<string, string> {
  const pt = preview.fontSizePt;
  return {
    ...(preview.fontFamily ? { 'font-family': preview.fontFamily } : {}),
    ...(pt ? { 'font-size': `${Math.min(Math.max(pt, 9), 20)}px` } : {}),
    ...(preview.bold ? { 'font-weight': '700' } : {}),
    ...(preview.italic ? { 'font-style': 'italic' } : {}),
    ...(preview.color && /^[0-9a-f]{6}$/i.test(preview.color)
      ? { color: `#${preview.color}` }
      : {}),
  };
}

/**
 * The paragraph style picker (`styles.style`): a trigger that shows the style of the
 * selection, and a list of the document's paragraph styles with a preview of each.
 *
 * @example
 * ```html
 * <div docxToolbarParagraphStyle></div>
 * ```
 *
 * @public
 */
@Component({
  selector: 'div[docxToolbarParagraphStyle]',
  changeDetection: ChangeDetectionStrategy.OnPush,
  host: {
    class: 'docx-toolbar__style',
    'data-slot': 'styles.style',
  },
  template: `<button
      type="button"
      class="docx-toolbar__style-trigger"
      aria-haspopup="listbox"
      [attr.aria-expanded]="open()"
      [attr.title]="text()"
      [disabled]="!state.isEnabled()"
      [attr.data-disabled]="state.isEnabled() ? null : ''"
      (mousedown)="guard($event)"
      (click)="open.set(!open())"
    >
      <span>{{ current() }}</span>
    </button>
    @if (open()) {
      <div
        role="listbox"
        tabindex="0"
        [attr.aria-label]="label()('styles.selectAriaLabel')"
        class="docx-toolbar__menu docx-toolbar__style-content"
      >
        @for (option of state.options(); track option.styleId) {
          <button
            type="button"
            role="option"
            tabindex="-1"
            class="docx-toolbar__style-item"
            [attr.aria-selected]="option.styleId === state.value()"
            [attr.data-selected]="option.styleId === state.value() ? '' : null"
            (mousedown)="guard($event)"
            (click)="choose(option.styleId)"
          >
            <span [style]="preview(option.preview)">{{ option.name }}</span>
            @if (option.styleId === state.value()) {
              <span class="docx-toolbar__menu-check" aria-hidden="true">✓</span>
            }
          </button>
        }
      </div>
    }`,
})
export class DocxToolbarParagraphStyleComponent {
  private readonly editor = injectDocxEditor();
  private readonly zone = inject(NgZone);
  private readonly host: HTMLElement = inject(ElementRef).nativeElement;

  /** @internal */
  protected readonly state: ParagraphStyleState = injectParagraphStyle();
  /** @internal */
  protected readonly label = injectToolbarLabel();
  /** @internal */
  protected readonly open = signal(false);
  /** @internal */
  protected readonly guard: (event: Event) => void = guardToolbarMousedown;
  /** @internal */
  protected readonly text = computed(() =>
    this.label()(chromeControlForSlot('styles.style')?.labelKey ?? 'styles.style')
  );
  /** @internal */
  protected readonly current = computed(() => {
    const value = this.state.value();
    if (value === null) {
      const key = chromeControlForSlot('styles.style')?.valueKey;
      return key ? this.label()(key) : '—';
    }
    return this.state.options().find((option) => option.styleId === value)?.name ?? value;
  });
  /** @internal */
  protected preview(preview: ParagraphStyleOption['preview']): Record<string, string> {
    return previewStyle(preview);
  }

  constructor() {
    injectPopupDismiss({
      open: this.open,
      root: () => this.host,
      kind: 'picker',
      close: () => this.open.set(false),
    });
  }

  /** @internal */
  protected choose(styleId: string): void {
    this.state.setValue(styleId);
    this.open.set(false);
    focusPages(this.zone, this.editor);
  }
}
