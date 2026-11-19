import {
  ChangeDetectionStrategy,
  Component,
  ElementRef,
  NgZone,
  afterRenderEffect,
  computed,
  inject,
  signal,
  untracked,
  viewChild,
} from '@angular/core';
import type { DocumentEditingMode, EditorSnapshot } from '@docx-editor.dev/core/contracts/editor';
import {
  hostDocumentOf,
  runToolbarCommand,
  type DocxEditorInstance,
} from '@docx-editor.dev/core/editor';
import { localizeDisabledReason } from '@docx-editor.dev/i18n';
import { injectDocxEditor } from '../context';
import { injectEditorCommand, type EditorCommandState } from '../editor-command';
import { injectEditorState } from '../editor-state';
import { injectTranslation } from '../i18n';
import { DocxChromeIconComponent } from './chrome-icon.component';
import { injectToolbarLabel } from './toolbar-context';
import { injectPopupDismiss } from './toolbar-popup';
import { chromeControlForSlot, guardToolbarMousedown } from './toolbar-registry';

// Keyboard travel skips a refused item.
const MENU_ITEMS = '[role="menuitemradio"]:not([disabled])';
const CHECK_PATHS: readonly string[] = ['M382-240 154-468l57-57 171 171 367-367 57 57-424 424Z'];

/** One mode of the editing mode menu. @internal */
export interface ModeOption {
  readonly mode: DocumentEditingMode;
  readonly labelKey: 'editingMode.editing' | 'editingMode.suggesting' | 'editingMode.viewing';
  readonly hintKey:
    'editingMode.editingHint' | 'editingMode.suggestingHint' | 'editingMode.viewingHint';
  readonly paths: readonly string[];
}

const MODE_OPTIONS: readonly ModeOption[] = [
  {
    mode: 'editing',
    labelKey: 'editingMode.editing',
    hintKey: 'editingMode.editingHint',
    paths: [
      'M200-200h57l391-391-57-57-391 391v57Zm-80 80v-170l528-527q12-11 26.5-17t30.5-6q16 0 31 6t26 18l55 56q12 11 17.5 26t5.5 30q0 16-5.5 30.5T817-647L290-120H120Zm640-584-56-56 56 56Zm-141 85-28-29 57 57-29-28Z',
    ],
  },
  {
    mode: 'suggesting',
    labelKey: 'editingMode.suggesting',
    hintKey: 'editingMode.suggestingHint',
    paths: [
      'M240-400h122l40-40H240v40Zm0-100h222l40-40H240v40Zm0-100h322l40-40H240v40ZM80-80v-720q0-33 23.5-56.5T160-880h640q33 0 56.5 23.5T880-800v320h-80v-320H160v525l46-45h274v80H240L80-80Zm520-80v-123l221-220q9-9 20-13t22-4q12 0 23 4.5t20 13.5l37 37q8 9 12.5 20t4.5 22q0 11-4 22.5T943-380L723-160H600Zm300-263-37-37 37 37ZM660-220h38l121-122-19-18-18-19-122 121v38Zm140-141-18-19 37 37-19-18Z',
    ],
  },
  {
    mode: 'viewing',
    labelKey: 'editingMode.viewing',
    hintKey: 'editingMode.viewingHint',
    paths: [
      'M480-320q75 0 127.5-52.5T660-500q0-75-52.5-127.5T480-680q-75 0-127.5 52.5T300-500q0 75 52.5 127.5T480-320Zm0-72q-45 0-76.5-31.5T372-500q0-45 31.5-76.5T480-608q45 0 76.5 31.5T588-500q0 45-31.5 76.5T480-392Zm0 192q-146 0-266-81.5T40-500q54-137 174-218.5T480-800q146 0 266 81.5T920-500q-54 137-174 218.5T480-200Z',
    ],
  },
];

type ItemReasons = readonly (string | null)[];
const NO_REASONS: ItemReasons = Object.freeze(MODE_OPTIONS.map(() => null));
const sameReasons = (a: ItemReasons, b: ItemReasons): boolean =>
  a.length === b.length && a.every((reason, index) => reason === b[index]);
const selectMode = (snapshot: EditorSnapshot): DocumentEditingMode =>
  snapshot.editingMode ?? 'editing';
const selectLoading = (snapshot: EditorSnapshot) =>
  snapshot.isLoading || snapshot.isOpening === true;

/** Each item's refusal from the engine, or `null` where its mode can be entered. */
function itemReasonsOf(editor: DocxEditorInstance | null): ItemReasons {
  if (!editor) return NO_REASONS;
  return MODE_OPTIONS.map((option) => {
    const probe = editor.can({ type: 'setEditingMode', mode: option.mode });
    return probe.ok ? null : probe.reason;
  });
}

/**
 * The editing mode menu (`review.editingMode`): Editing, Suggesting, or Viewing. A mode that the
 * engine refuses, such as Suggesting without the review module, renders disabled with the
 * engine's reason.
 *
 * @example
 * ```html
 * <div docxToolbarEditingMode></div>
 * ```
 *
 * @public
 */
@Component({
  selector: 'div[docxToolbarEditingMode]',
  changeDetection: ChangeDetectionStrategy.OnPush,
  imports: [DocxChromeIconComponent],
  host: {
    class: 'docx-toolbar__mode',
    'data-slot': 'review.editingMode',
  },
  template: `<button
      type="button"
      class="docx-toolbar__picker"
      data-testid="editing-mode-trigger"
      aria-haspopup="menu"
      [attr.data-mode]="mode()"
      [attr.aria-expanded]="open()"
      [attr.aria-label]="text()"
      [attr.title]="command.disabledReason()"
      [disabled]="loading() || !command.isEnabled()"
      (mousedown)="guard($event)"
      (click)="open.set(!open())"
    >
      <svg [docxChromeIcon]="current().paths"></svg>
      <span class="docx-toolbar__picker-value">{{ label()(current().labelKey) }}</span>
      <span class="docx-toolbar__picker-caret" aria-hidden="true">▾</span>
    </button>
    @if (open()) {
      <div
        #menu
        class="docx-toolbar__mode-menu"
        role="menu"
        data-testid="editing-mode-menu"
        [attr.aria-label]="text()"
        (keydown)="onMenuKey($event)"
      >
        @for (option of options; track option.mode; let index = $index) {
          <button
            type="button"
            role="menuitemradio"
            class="docx-toolbar__mode-item"
            [attr.data-testid]="'editing-mode-' + option.mode"
            [attr.aria-checked]="option.mode === mode()"
            [disabled]="reasons()[index] !== null"
            [attr.title]="reasons()[index]"
            (mousedown)="guard($event)"
            (click)="choose(option.mode)"
          >
            <svg [docxChromeIcon]="option.paths"></svg>
            <span class="docx-toolbar__mode-text">
              <span class="docx-toolbar__mode-label">{{ label()(option.labelKey) }}</span>
              <span class="docx-toolbar__mode-hint">{{ label()(option.hintKey) }}</span>
            </span>
            <span class="docx-toolbar__mode-check" aria-hidden="true">
              @if (option.mode === mode()) {
                <svg [docxChromeIcon]="checkPaths"></svg>
              }
            </span>
          </button>
        }
      </div>
    }`,
})
export class DocxToolbarEditingModeComponent {
  private readonly editor = injectDocxEditor();
  private readonly zone = inject(NgZone);
  private readonly host: HTMLElement = inject(ElementRef).nativeElement;
  private readonly menu = viewChild<ElementRef<HTMLElement>>('menu');
  private readonly translation = injectTranslation();
  // The engine's answer can change with no snapshot change, so this reads it on each tick.
  private readonly rawReasons = injectEditorState(
    () => untracked(() => itemReasonsOf(this.editor())),
    sameReasons
  );

  /** @internal */
  protected readonly options = MODE_OPTIONS;
  /** @internal */
  protected readonly checkPaths = CHECK_PATHS;
  /** @internal */
  protected readonly command: EditorCommandState = injectEditorCommand('review.editingMode');
  /** @internal */
  protected readonly mode = injectEditorState(selectMode);
  /** @internal */
  protected readonly loading = injectEditorState(selectLoading);
  /** @internal */
  protected readonly label = injectToolbarLabel();
  /** @internal */
  protected readonly open = signal(false);
  /** @internal */
  protected readonly guard: (event: Event) => void = guardToolbarMousedown;
  /** @internal */
  protected readonly text = computed(() =>
    this.label()(chromeControlForSlot('review.editingMode')?.labelKey ?? 'editingMode.label')
  );
  /** @internal */
  protected readonly current = computed(
    () => MODE_OPTIONS.find((option) => option.mode === this.mode()) ?? MODE_OPTIONS[0]!
  );
  /** @internal */
  protected readonly reasons = computed(() => {
    const t = this.translation.t();
    return this.rawReasons().map((reason) => localizeDisabledReason(reason, t));
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
    // An open menu takes the focus: the checked item when it is enabled, else the first one.
    afterRenderEffect(() => {
      const menu = this.menu()?.nativeElement;
      if (!menu) return;
      const checked = menu.querySelector<HTMLElement>(`${MENU_ITEMS}[aria-checked="true"]`);
      (checked ?? menu.querySelector<HTMLElement>(MENU_ITEMS))?.focus();
    });
  }

  /** @internal */
  protected onMenuKey(event: KeyboardEvent): void {
    const menu = this.menu()?.nativeElement;
    const items = [...(menu?.querySelectorAll<HTMLElement>(MENU_ITEMS) ?? [])];
    const active = hostDocumentOf(this.host).activeElement;
    const at = items.findIndex((item) => item === active);
    const move = (to: number) => {
      event.preventDefault();
      items[(to + items.length) % items.length]?.focus();
    };
    if (event.key === 'ArrowDown') move(at + 1);
    else if (event.key === 'ArrowUp') move(at - 1);
    else if (event.key === 'Home') move(0);
    else if (event.key === 'End') move(items.length - 1);
  }

  /** @internal */
  protected choose(mode: DocumentEditingMode): void {
    this.open.set(false);
    const editor = untracked(this.editor);
    this.zone.runOutsideAngular(() => {
      runToolbarCommand(editor, 'review.editingMode', mode);
      editor?.focus();
    });
  }
}
