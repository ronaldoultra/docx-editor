import {
  ChangeDetectionStrategy,
  Component,
  ElementRef,
  NgZone,
  afterRenderEffect,
  computed,
  inject,
  input,
  output,
  signal,
  untracked,
  viewChild,
} from '@angular/core';
import { isDomNode, toolbarPopupLeft, type ChromeSlotId } from '@docx-editor.dev/core/editor';
import { injectDocxEditor } from '../context';
import { injectEditorCommand, type EditorCommandState } from '../editor-command';
import { DocxChromeIconComponent } from './chrome-icon.component';
import { DocxToolbarState, injectToolbarLabel } from './toolbar-context';
import { injectPopupDismiss, popupId } from './toolbar-popup';
import { chromeControlForSlot, guardToolbarMousedown } from './toolbar-registry';

/** The insert-table grid is 6 columns by 6 rows. */
const GRID_COLUMNS = 6;
const GRID_ROWS = 6;
const range = (count: number) => Array.from({ length: count }, (_, index) => index + 1);

/**
 * A 6 × 6 grid that highlights as the pointer sweeps it, and inserts a table of the size under
 * it. One tab stop, with the arrows, Home, and End inside it.
 *
 * @internal
 */
@Component({
  selector: 'docx-table-size-grid',
  changeDetection: ChangeDetectionStrategy.OnPush,
  host: { style: 'display: contents' },
  template: `<div
    #grid
    role="grid"
    class="docx-menubar__grid"
    [attr.aria-label]="label()"
    (mouseleave)="hover.set(null)"
    (keydown)="onKey($event)"
  >
    <div class="docx-menubar__grid-cells">
      @for (row of rows; track row) {
        <div role="row" class="docx-menubar__grid-row">
          @for (col of cols; track col) {
            <button
              type="button"
              role="gridcell"
              class="docx-menubar__grid-cell"
              [attr.data-cell]="row + 'x' + col"
              [attr.tabindex]="cursor().rows === row && cursor().cols === col ? 0 : -1"
              [attr.data-filled]="filled(row, col) ? '' : null"
              [attr.aria-label]="col + ' × ' + row"
              (mousedown)="guard($event)"
              (mouseenter)="hover.set({ rows: row, cols: col })"
              (focus)="hover.set({ rows: row, cols: col })"
              (click)="insert(row, col)"
            ></button>
          }
        </div>
      }
    </div>
    <div class="docx-menubar__grid-caption" aria-hidden="true">
      {{ hover() ? hover()!.cols + ' × ' + hover()!.rows : '' }}
    </div>
  </div>`,
})
export class DocxTableSizeGridComponent {
  /** The slot whose enabled state gates the insert. */
  readonly slot = input<ChromeSlotId>('table.insert');
  /** The accessible name of the grid. */
  readonly label = input.required<string>();
  /** Emits after a table is inserted, to close the popup around the grid. */
  readonly inserted = output<void>();

  private readonly editor = injectDocxEditor();
  private readonly zone = inject(NgZone);
  private readonly command = injectEditorCommand(() => this.slot());
  private readonly grid = viewChild.required<ElementRef<HTMLElement>>('grid');

  /** @internal */
  protected readonly rows = range(GRID_ROWS);
  /** @internal */
  protected readonly cols = range(GRID_COLUMNS);
  /** @internal */
  protected readonly hover = signal<{ rows: number; cols: number } | null>(null);
  /** @internal */
  protected readonly cursor = signal({ rows: 1, cols: 1 });
  /** @internal */
  protected readonly guard: (event: Event) => void = guardToolbarMousedown;

  /** @internal */
  protected filled(row: number, col: number): boolean {
    const hover = this.hover();
    return !!hover && row <= hover.rows && col <= hover.cols;
  }

  /** @internal */
  protected insert(rows: number, cols: number): void {
    const editor = untracked(this.editor);
    if (!editor || !untracked(this.command.isEnabled)) return;
    // Can before exec: the selection can move while the grid is open.
    const command = { type: 'insertTable' as const, rows, cols };
    const inserted = this.zone.runOutsideAngular(() => {
      if (!editor.can(command).ok) return false;
      editor.exec(command);
      return true;
    });
    if (!inserted) return;
    this.inserted.emit();
    // The engine left the caret in the first cell, so the focus goes back to the document.
    this.zone.runOutsideAngular(() => editor.focus());
  }

  /** @internal */
  protected onKey(event: KeyboardEvent): void {
    const { rows, cols } = untracked(this.cursor);
    if (event.key === 'ArrowRight') this.move(rows, cols + 1);
    else if (event.key === 'ArrowLeft') this.move(rows, cols - 1);
    else if (event.key === 'ArrowDown') this.move(rows + 1, cols);
    else if (event.key === 'ArrowUp') this.move(rows - 1, cols);
    else if (event.key === 'Home') this.move(rows, 1);
    else if (event.key === 'End') this.move(rows, GRID_COLUMNS);
    else if (event.key === 'Enter' || event.key === ' ') this.insert(rows, cols);
    else return;
    // Stopped, so the arrows of the grid do not also move through the rows behind it.
    event.preventDefault();
    event.stopPropagation();
  }

  private move(rows: number, cols: number): void {
    const next = {
      rows: Math.min(GRID_ROWS, Math.max(1, rows)),
      cols: Math.min(GRID_COLUMNS, Math.max(1, cols)),
    };
    this.cursor.set(next);
    this.hover.set(next);
    queueMicrotask(() =>
      this.grid()
        .nativeElement.querySelector<HTMLElement>(`[data-cell="${next.rows}x${next.cols}"]`)
        ?.focus()
    );
  }
}

/**
 * Insert Table (`table.insert`): a button that opens the table size grid. The slot's own command
 * inserts a 1 × 1 table, so the button opens the grid, and the engine still decides whether the
 * button is enabled.
 *
 * @example
 * ```html
 * <span docxToolbarTableInsert></span>
 * ```
 *
 * @public
 */
@Component({
  selector: 'span[docxToolbarTableInsert]',
  changeDetection: ChangeDetectionStrategy.OnPush,
  imports: [DocxChromeIconComponent, DocxTableSizeGridComponent],
  host: { class: 'docx-toolbar__table-insert' },
  template: `<button
      #trigger
      type="button"
      class="docx-toolbar__button"
      data-slot="table.insert"
      aria-haspopup="grid"
      [disabled]="!command.isEnabled()"
      [attr.data-disabled]="command.isEnabled() ? null : ''"
      [attr.data-active]="shown() ? '' : null"
      [attr.aria-expanded]="shown()"
      [attr.aria-controls]="shown() ? gridId : null"
      [attr.aria-label]="text()"
      [attr.title]="command.disabledReason() ?? text()"
      (mousedown)="guard($event)"
      (click)="onTrigger()"
      (keydown.arrowdown)="$event.preventDefault(); open.set(true)"
    >
      <ng-content>
        @if (paths; as icon) {
          <svg [docxChromeIcon]="icon"></svg>
        }
      </ng-content>
    </button>
    @if (shown()) {
      <div
        #popup
        class="docx-toolbar__menu docx-toolbar__table-insert-menu"
        [id]="gridId"
        (focusout)="onFocusOut($event)"
        (keydown)="onPopupKey($event)"
      >
        <docx-table-size-grid [label]="label()('toolbar.insertTable')" (inserted)="close()" />
      </div>
    }`,
})
export class DocxToolbarTableInsertComponent {
  private readonly editor = injectDocxEditor();
  private readonly zone = inject(NgZone);
  private readonly host: HTMLElement = inject(ElementRef).nativeElement;
  private readonly toolbar = inject(DocxToolbarState, { optional: true });
  private readonly trigger = viewChild.required<ElementRef<HTMLElement>>('trigger');
  private readonly popup = viewChild<ElementRef<HTMLElement>>('popup');

  /** @internal */
  protected readonly command: EditorCommandState = injectEditorCommand('table.insert');
  /** @internal */
  protected readonly label = injectToolbarLabel();
  /** @internal */
  protected readonly open = signal(false);
  /** @internal */
  protected readonly shown = computed(() => this.open() && this.command.isEnabled());
  /** @internal */
  protected readonly gridId = popupId('table-insert');
  /** @internal */
  protected readonly guard: (event: Event) => void = guardToolbarMousedown;
  /** @internal */
  protected readonly paths = chromeControlForSlot('table.insert')?.paths ?? null;
  /** @internal */
  protected readonly text = computed(() =>
    this.label()(chromeControlForSlot('table.insert')?.labelKey ?? 'table.insert')
  );

  constructor() {
    injectPopupDismiss({
      open: this.shown,
      root: () => this.host,
      // The grid's own key handler takes Escape.
      skip: (event) => isDomNode(event.target) && this.host.contains(event.target),
      close: () => this.open.set(false),
    });
    afterRenderEffect(() => {
      const popup = this.popup()?.nativeElement;
      if (!popup) return;
      // In the bar, the grid opens under its button, moved in only as far as it must to stay
      // inside the viewport. In the "⋯" panel it is panel content.
      const view = this.host.ownerDocument.defaultView;
      if (view && !this.host.closest('.docx-toolbar__more-panel')) {
        const anchor = this.host.getBoundingClientRect();
        const left = toolbarPopupLeft({
          anchorLeft: anchor.left,
          popupWidth: popup.offsetWidth,
          viewportWidth: view.innerWidth,
        });
        popup.style.left = `${left - anchor.left}px`;
      }
      // Each open moves the focus into the grid's tab stop, so the grid hears Escape.
      popup.querySelector<HTMLElement>('[role="gridcell"][tabindex="0"]')?.focus();
    });
  }

  /** @internal */
  protected onTrigger(): void {
    if (!untracked(this.open)) {
      this.open.set(true);
      return;
    }
    // The focus is in the grid, which leaves the page: give it back to the document.
    this.open.set(false);
    this.zone.runOutsideAngular(() => untracked(this.editor)?.focus());
  }

  /** @internal */
  protected onFocusOut(event: FocusEvent): void {
    const next = event.relatedTarget;
    if (isDomNode(next) && !this.host.contains(next)) this.open.set(false);
  }

  /** @internal */
  protected onPopupKey(event: KeyboardEvent): void {
    if (event.key !== 'Escape') return;
    // Stopped, so an enclosing "⋯" panel stays open: Escape closes one layer.
    event.preventDefault();
    event.stopPropagation();
    this.open.set(false);
    this.trigger().nativeElement.focus();
  }

  /** @internal */
  protected close(): void {
    this.open.set(false);
    // Inside the "⋯" panel, an insert also closes the panel.
    if (this.host.closest('.docx-toolbar__more-panel')) this.toolbar?.closeMore(false);
  }
}
