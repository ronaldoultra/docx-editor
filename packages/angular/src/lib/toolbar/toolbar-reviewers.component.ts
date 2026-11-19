import {
  ChangeDetectionStrategy,
  Component,
  ElementRef,
  NgZone,
  afterRenderEffect,
  booleanAttribute,
  computed,
  inject,
  input,
  signal,
  untracked,
  viewChild,
} from '@angular/core';
import type { EditorSnapshot } from '@docx-editor.dev/core/contracts/editor';
import { hostDocumentOf, type ReviewAuthorInfo } from '@docx-editor.dev/core/editor';
import { injectDocxEditor } from '../context';
import { injectEditorCommand, type EditorCommandState } from '../editor-command';
import { injectEditorState } from '../editor-state';
import { DocxChromeIconComponent } from './chrome-icon.component';
import { injectToolbarLabel } from './toolbar-context';
import { injectPopupDismiss } from './toolbar-popup';
import { chromeControlForSlot, guardToolbarMousedown } from './toolbar-registry';

const MENU_ITEMS = '[role="menuitemcheckbox"]';
const EMPTY_AUTHORS: readonly ReviewAuthorInfo[] = [];
const selectHiddenAuthors = (snapshot: EditorSnapshot): readonly string[] =>
  snapshot.hiddenReviewAuthors ?? [];
const sameList = (a: readonly string[], b: readonly string[]) =>
  a.length === b.length && a.every((value, index) => value === b[index]);

/**
 * The reviewers menu (`review.authors`), as in Word for Mac: a checked row for all review
 * authors, then one checked row for each author. It changes the view only, not the document.
 * The trigger is disabled with the engine's reason when the document has no review authors.
 *
 * Put your own icon in the content. Without content, the trigger shows the registry icon.
 *
 * @example
 * ```html
 * <div docxToolbarReviewers></div>
 * ```
 *
 * @public
 */
@Component({
  selector: 'div[docxToolbarReviewers]',
  changeDetection: ChangeDetectionStrategy.OnPush,
  imports: [DocxChromeIconComponent],
  host: {
    '[class]': 'hostClass()',
    '[style.display]': 'hidden() ? "none" : null',
    'data-slot': 'review.authors',
  },
  template: `@if (!hidden()) {
    <button
      type="button"
      class="docx-toolbar__picker"
      data-testid="reviewers-trigger"
      aria-haspopup="menu"
      [attr.aria-expanded]="open()"
      [attr.aria-label]="text()"
      [attr.title]="command.disabledReason() ?? text()"
      [disabled]="!command.isEnabled()"
      (mousedown)="guard($event)"
      (click)="open.set(!open())"
    >
      <ng-content><svg [docxChromeIcon]="iconPaths"></svg></ng-content>
      <span class="docx-toolbar__picker-value">{{ text() }}</span>
      <span class="docx-toolbar__picker-caret" aria-hidden="true">▾</span>
    </button>
    @if (open()) {
      <div
        #menu
        class="docx-toolbar__reviewers-menu"
        role="menu"
        data-testid="reviewers-menu"
        [attr.aria-label]="text()"
        (keydown)="onMenuKey($event)"
      >
        <button
          type="button"
          role="menuitemcheckbox"
          class="docx-toolbar__reviewer-item"
          data-testid="reviewer-all"
          [attr.aria-checked]="allVisible()"
          (mousedown)="guard($event)"
          (click)="toggleAll()"
        >
          <span class="docx-toolbar__menu-check" aria-hidden="true">{{
            allVisible() ? '✓' : ''
          }}</span>
          <span class="docx-toolbar__reviewer-name">{{ label()('reviewers.all') }}</span>
        </button>
        <div class="docx-toolbar__menu-separator" role="separator"></div>
        @for (info of authors(); track info.author) {
          <button
            type="button"
            role="menuitemcheckbox"
            class="docx-toolbar__reviewer-item"
            [attr.data-testid]="'reviewer-' + info.slot"
            [attr.aria-checked]="!hiddenSet().has(info.author)"
            (mousedown)="guard($event)"
            (click)="toggle(info.author)"
          >
            <span class="docx-toolbar__menu-check" aria-hidden="true">{{
              hiddenSet().has(info.author) ? '' : '✓'
            }}</span>
            @if (info.color) {
              <span class="docx-toolbar__reviewer-swatch" [style.background]="info.color"></span>
            }
            <span class="docx-toolbar__reviewer-name">{{ info.author || '—' }}</span>
          </button>
        }
      </div>
    }
  }`,
})
export class DocxToolbarReviewersComponent {
  /** More classes for the root. */
  readonly className = input<string | undefined>(undefined);
  /** Render nothing. */
  readonly hidden = input(false, { transform: booleanAttribute });

  private readonly editor = injectDocxEditor();
  private readonly zone = inject(NgZone);
  private readonly host: HTMLElement = inject(ElementRef).nativeElement;
  private readonly menu = viewChild<ElementRef<HTMLElement>>('menu');
  private readonly hiddenAuthors = injectEditorState(selectHiddenAuthors, sameList);

  /** @internal */
  protected readonly iconPaths: readonly string[] =
    chromeControlForSlot('review.authors')?.paths ?? [];
  /** @internal */
  protected readonly command: EditorCommandState = injectEditorCommand('review.authors');
  /** @internal */
  protected readonly label = injectToolbarLabel();
  /** @internal */
  protected readonly open = signal(false);
  /** @internal */
  protected readonly guard: (event: Event) => void = guardToolbarMousedown;
  /** @internal */
  protected readonly hostClass = computed(() => {
    const extra = this.className();
    return `docx-toolbar__reviewers${extra ? ` ${extra}` : ''}`;
  });
  /** @internal */
  protected readonly text = computed(() =>
    this.label()(chromeControlForSlot('review.authors')?.labelKey ?? 'reviewers.label')
  );
  /** @internal */
  protected readonly authors = injectEditorState(() => {
    const editor = this.editor();
    return untracked(() => editor?.getReviewAuthors() ?? EMPTY_AUTHORS);
  });
  /** @internal */
  protected readonly hiddenSet = computed(() => new Set(this.hiddenAuthors()));
  /** @internal */
  protected readonly allVisible = computed(() => {
    const hidden = this.hiddenSet();
    return this.authors().every((info) => !hidden.has(info.author));
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
    // An open menu takes the focus on its first row.
    afterRenderEffect(() => {
      this.menu()?.nativeElement.querySelector<HTMLElement>(MENU_ITEMS)?.focus();
    });
  }

  /** @internal */
  protected onMenuKey(event: KeyboardEvent): void {
    const items = [...(this.menu()?.nativeElement.querySelectorAll<HTMLElement>(MENU_ITEMS) ?? [])];
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
    else if (event.key === 'Tab') this.open.set(false);
  }

  /** @internal */
  protected toggleAll(): void {
    const visible = untracked(this.allVisible);
    this.zone.runOutsideAngular(() => untracked(this.editor)?.setAllReviewAuthorsVisible(!visible));
  }

  /** @internal */
  protected toggle(author: string): void {
    const hidden = untracked(this.hiddenSet).has(author);
    this.zone.runOutsideAngular(() =>
      untracked(this.editor)?.setReviewAuthorVisible(author, hidden)
    );
  }
}
