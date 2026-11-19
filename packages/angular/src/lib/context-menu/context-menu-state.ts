// The state that one context menu shares with its rows: where it opened, what it opened on,
// and the clipboard refusal that disables Paste.

import {
  Injectable,
  computed,
  effect,
  inject,
  signal,
  untracked,
  type Signal,
} from '@angular/core';
import type { HighlightHit, HighlightRange } from '@docx-editor.dev/core/contracts/editor';
import { editorHostWindow } from '@docx-editor.dev/core/editor';
import { injectDocxEditor } from '../context';
import { DocxMenuBarState } from '../menu/menu-state';
import { readContextMenuHighlights, type ContextMenuAnchor } from './context-menu-target';

/** @internal */
@Injectable()
export class DocxContextMenuState {
  /** Where the open menu is anchored, or `null` while it is closed. */
  readonly anchor = signal<ContextMenuAnchor | null>(null);
  /** The id of the table of contents that the menu opened on, or `null`. */
  readonly tocId = signal<string | null>(null);
  /** The element that the opening right-click landed on, or `null`. */
  readonly target = signal<HTMLElement | null>(null);
  /** The reason why the last clipboard read failed, or `null`. Paste is disabled with it. */
  readonly clipboardRefusal = signal<string | null>(null);
  /** False when the menu renders only the host's content. */
  readonly preset = signal(true);
  /** Opens the split cells dialog. Without it, the menu has no Split cells row. */
  readonly splitCell = signal<(() => void) | undefined>(undefined);

  private closer: (restoreFocus: boolean) => void = () => {};

  /** Close the menu. With `restoreFocus`, the focus goes back to the document. */
  close(restoreFocus = false): void {
    this.closer(restoreFocus);
  }

  /** @internal */
  useClose(close: (restoreFocus: boolean) => void): void {
    this.closer = close;
  }
}

/**
 * The menu bar state inside a context menu. The menu rows, such as `docx-menu-item`, close
 * their bar through `setOpenMenu(null)`. Here that call closes the context menu and gives the
 * focus back to the document.
 *
 * @internal
 */
@Injectable()
export class DocxContextMenuBarState extends DocxMenuBarState {
  private readonly menu = inject(DocxContextMenuState);

  override setOpenMenu(id: string | null): void {
    if (id === null) this.menu.close(true);
  }
}

const CLOSED: Signal<null> = signal(null).asReadonly();

/** The state of the enclosing context menu, or `null` outside one. @internal */
export function injectContextMenuState(): DocxContextMenuState | null {
  return inject(DocxContextMenuState, { optional: true });
}

/**
 * Where the open context menu is anchored and how it was opened, or `null` while it is closed.
 *
 * Rows in `docx-context-menu-start` read `trigger` to tell a right-click from a keyboard open.
 * A keyboard open acts on the caret and not on a pointer target. Call it in an injection
 * context inside the context menu.
 *
 * @public
 */
export function injectContextMenuAnchor(): Signal<ContextMenuAnchor | null> {
  return injectContextMenuState()?.anchor.asReadonly() ?? CLOSED;
}

/**
 * The element that the opening right-click landed on, or `null` while the menu is closed. A
 * menu that the keyboard opened has no target: read the caret with
 * `editor.getSelectionRange()`. Call it in an injection context inside the context menu.
 *
 * @public
 */
export function injectContextMenuTarget(): Signal<HTMLElement | null> {
  return injectContextMenuState()?.target.asReadonly() ?? CLOSED;
}

/**
 * The text highlights that the open context menu is about, topmost first. Pass a set name to
 * read one set. `R` is your range type, so the fields that you added to your ranges come back
 * typed on `range`.
 *
 * For a right-click, these are the highlights under the pointer. For a keyboard open, these are
 * the highlights whose current range holds the caret. A document change while the menu is open
 * reads the highlights again. The signal is empty while the menu is closed. Call it in an
 * injection context inside the context menu.
 *
 * @example
 * ```ts
 * readonly hits = injectContextMenuHighlights<MisspellingRange>('spelling');
 * ```
 *
 * @public
 */
export function injectContextMenuHighlights<R extends HighlightRange = HighlightRange>(
  name?: string
): Signal<readonly HighlightHit<R>[]> {
  const editor = injectDocxEditor();
  const anchor = injectContextMenuAnchor();
  const hits = signal<readonly HighlightHit<R>[]>(readContextMenuHighlights<R>(null, null, name));
  effect((onCleanup) => {
    const instance = editor();
    const opened = anchor();
    const read = () => hits.set(readContextMenuHighlights<R>(instance, opened, name));
    untracked(read);
    if (!instance || !opened) return;
    // Read again after each document change: at once, for a change painted before it reports,
    // and one frame later, for a paint that input pressure deferred.
    const view = editorHostWindow(instance);
    let frame = 0;
    const off = instance.on('change', () => {
      read();
      if (frame === 0) {
        frame = view.requestAnimationFrame(() => {
          frame = 0;
          read();
        });
      }
    });
    onCleanup(() => {
      off();
      if (frame !== 0) view.cancelAnimationFrame(frame);
    });
  });
  return computed(() => hits());
}

/**
 * The ids of the rows of the context menu that are not chrome slots. A chrome slot id, such as
 * `text.link` or `table.mergeCells`, also names a row.
 *
 * @public
 */
export type ContextMenuRowId =
  | 'edit.cut'
  | 'edit.copy'
  | 'edit.paste'
  | 'edit.pasteWithoutFormatting'
  | 'edit.delete'
  | 'edit.selectAll'
  | 'format.copyFormatting'
  | 'format.pasteFormatting'
  | 'field.edit'
  | 'table.selectTable'
  | 'table.insertRowAbove'
  | 'table.insertRowBelow'
  | 'table.insertColumnLeft'
  | 'table.insertColumnRight'
  | 'table.mergeCells'
  | 'table.splitCell'
  | 'table.deleteRow'
  | 'table.deleteColumn'
  | 'table.deleteTable'
  | 'table.cellVerticalAlignment'
  | 'toc.refresh'
  | 'toc.refreshPageNumbers'
  | 'toc.remove';
