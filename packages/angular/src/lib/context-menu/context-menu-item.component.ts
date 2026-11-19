import {
  ChangeDetectionStrategy,
  Component,
  Directive,
  ElementRef,
  TemplateRef,
  afterEveryRender,
  booleanAttribute,
  inject,
  input,
  output,
} from '@angular/core';
import type { ChromeSlotId } from '@docx-editor.dev/core/editor';
import { DocxMenuRowComponent } from '../menu/menu-row.component';
import { toolbarDevWarning } from '../toolbar/toolbar-registry';
import { DocxContextMenuState, type ContextMenuRowId } from './context-menu-state';

/**
 * A row of your own in the context menu. Choosing it closes the menu, gives the focus back to
 * the document, and then emits `activate`. A handler that moves the focus, such as one that
 * opens a dialog, keeps that focus.
 *
 * @example
 * ```html
 * <div docxEditorContextMenu>
 *   <docx-context-menu-item label="Look up" (activate)="lookUp()" />
 * </div>
 * ```
 *
 * @public
 */
@Component({
  selector: 'docx-context-menu-item',
  changeDetection: ChangeDetectionStrategy.OnPush,
  imports: [DocxMenuRowComponent],
  host: { style: 'display: contents' },
  template: `<button
    docxMenuRow
    [icon]="icon()"
    [shortcut]="shortcut()"
    [disabled]="disabled()"
    [reason]="disabled() ? reason() : undefined"
    [active]="active()"
    [rowClass]="className()"
    (activate)="run()"
  >
    {{ label() }}
  </button>`,
})
export class DocxContextMenuItemComponent {
  /** The row text. */
  readonly label = input.required<string>();
  /** Material Symbols paths for the icon column. */
  readonly icon = input<readonly string[] | undefined>(undefined);
  /** The shortcut text. Ctrl names the reader's own modifier key. */
  readonly shortcut = input<string | undefined>(undefined);
  /** Disable the row. It stays focusable. */
  readonly disabled = input(false, { transform: booleanAttribute });
  /** Why the row is disabled. */
  readonly reason = input<string | undefined>(undefined);
  /** The checked state of a row that toggles. Leave it unset on a row that only acts. */
  readonly active = input<boolean | undefined>(undefined);
  /** More classes for the row. */
  readonly className = input<string | undefined>(undefined);
  /** Emits after the menu closes, when the user chooses the row. */
  readonly activate = output<void>();

  private readonly menu = inject(DocxContextMenuState);

  /** @internal */
  protected run(): void {
    this.menu.close(true);
    this.activate.emit();
  }
}

/**
 * Rows of your own above the packaged rows of the context menu, with a separator below them.
 *
 * Use it for rows about the item under the pointer, such as spelling suggestions for the word
 * that was right-clicked. Read that item with `injectContextMenuTarget()` or
 * `injectContextMenuHighlights()`. A menu that the keyboard opened has no target:
 * `injectContextMenuAnchor()` reports `trigger: 'keyboard'`, and `editor.getSelectionRange()`
 * gives the caret.
 *
 * Put it directly in the context menu element. A start that a component of yours renders goes
 * after the packaged rows, and a development warning tells you so.
 *
 * @example
 * ```html
 * <div docxEditorContextMenu>
 *   <docx-context-menu-start>
 *     <docx-context-menu-item label="Ignore word" (activate)="ignore()" />
 *   </docx-context-menu-start>
 * </div>
 * ```
 *
 * @public
 */
@Component({
  selector: 'docx-context-menu-start',
  changeDetection: ChangeDetectionStrategy.OnPush,
  host: { style: 'display: contents' },
  template: '<ng-content />',
})
export class DocxContextMenuStartComponent {
  private readonly host: HTMLElement = inject(ElementRef).nativeElement;
  private readonly menu = inject(DocxContextMenuState, { optional: true });

  constructor() {
    afterEveryRender(() => {
      const parent = this.host.parentElement;
      if (!parent || !this.menu?.preset()) return;
      // A start in a start moves up with its parent.
      if (parent.getAttribute('role') === 'menu') return;
      if (parent.closest('docx-context-menu-start')) return;
      toolbarDevWarning(
        'A docx-context-menu-start is not a direct child of the context menu, so it renders ' +
          'after the packaged rows. Put it directly in the docxEditorContextMenu element.'
      );
    });
  }

  /** True when the start holds at least one element. @internal */
  hasRows(): boolean {
    return this.host.childElementCount > 0;
  }
}

/**
 * Replace one packaged row of the context menu. Put the template directly in the context menu.
 * The menu renders the template in the place of the row and keeps every other row. With
 * `docxContextMenuRowHidden`, the menu removes the row. A template for an id that the menu
 * does not have comes after the packaged rows.
 *
 * `ng-template[docxSlot]` does the same for a row whose id is a chrome slot, such as
 * `text.link`.
 *
 * @example
 * ```html
 * <div docxEditorContextMenu>
 *   <ng-template docxContextMenuRow="edit.delete" docxContextMenuRowHidden />
 *   <ng-template docxContextMenuRow="edit.cut">
 *     <docx-context-menu-row row="edit.cut" className="my-cut" />
 *   </ng-template>
 * </div>
 * ```
 *
 * @public
 */
@Directive({ selector: 'ng-template[docxContextMenuRow]' })
export class DocxContextMenuRowDirective {
  /** The row that the template replaces, such as `'edit.cut'`. */
  readonly docxContextMenuRow = input.required<ContextMenuRowId | ChromeSlotId>();
  /** Remove the row from the packaged rows. */
  readonly docxContextMenuRowHidden = input(false, { transform: booleanAttribute });

  /** The replacement content. */
  readonly template: TemplateRef<unknown> = inject(TemplateRef);
}
