// Which right-clicks the document context menu answers, where its panel opens, and which text
// highlights it is about.
//
// The menu acts on the document selection, so it answers only presses on the document: the
// painted pages, the handles of a selected picture, and the empty canvas around them. Other
// content in the scroller is chrome, such as the link popup, dialogs, review cards, and host
// overlays. A press there keeps the browser's own menu, so a text field keeps its native rows,
// its focus, and its selection.

import type { Editor, HighlightHit, HighlightRange } from '@docx-editor.dev/core/contracts/editor';
import { isElementNode } from '@docx-editor.dev/core/editor';

/**
 * Where the panel opened, in client coordinates, and how it was opened.
 *
 * @public
 */
export interface ContextMenuAnchor {
  /** The horizontal client coordinate of the panel, in CSS pixels. */
  readonly x: number;
  /** The vertical client coordinate of the panel, in CSS pixels. */
  readonly y: number;
  /**
   * `pointer` for a right-click, which opens at the pointer. `keyboard` for Shift+F10 or the
   * Menu key, which opens below the caret, or at the top-left of the document area when the
   * caret is not on a painted page.
   */
  readonly trigger: 'pointer' | 'keyboard';
}

/** Form controls that own their text and selection. */
const FORM_CONTROL = 'input, textarea, select';

/** A drop-down form field painted on the page: document text, not a control of its own. */
const PAINTED_FORM_FIELD = '[data-docx-form-dropdown]';

/** The element that the engine paints the pages into. The loading placeholder shares its class. */
const SURFACE = '.docx-paginated-surface:not(.docx-editor__loading-surface)';

/** Document content beside the surface: the resize frame of a selected picture. */
const SURFACE_OVERLAY = '.docx-image-selection-overlay';

/**
 * True when a `contextmenu` event in `scroller` is on the document: the pages, the handles of
 * a selected picture, or the canvas around the pages. False on a form control and on chrome.
 */
export function isDocumentContextMenuTarget(
  scroller: HTMLElement,
  target: EventTarget | null
): boolean {
  if (!isElementNode(target)) return false;
  const control = target.closest(FORM_CONTROL);
  if (control && !control.matches(PAINTED_FORM_FIELD)) return false;
  // Without the content component, the scroller holds no chrome to tell apart.
  const surface = scroller.querySelector(SURFACE);
  if (!surface) return true;
  return surface.contains(target) || target.contains(surface) || !!target.closest(SURFACE_OVERLAY);
}

/** True when the keyboard opened the menu (Shift+F10 or the Menu key): no pointer position. */
export function isKeyboardContextMenu(event: MouseEvent): boolean {
  return event.button === -1 || (event.clientX === 0 && event.clientY === 0);
}

/**
 * Where the panel opens, in client coordinates. A pointer open uses the pointer. A keyboard
 * open uses the bottom-left corner of the caret. When the caret has no painted position, the
 * panel opens at the top-left of the scroller.
 */
export function contextMenuAnchorOf(
  event: MouseEvent,
  keyboard: boolean,
  editor: Pick<Editor, 'getCaretRect'> | null,
  scroller: HTMLElement
): ContextMenuAnchor {
  if (!keyboard) return { x: event.clientX, y: event.clientY, trigger: 'pointer' };
  const caret = editor?.getCaretRect() ?? null;
  if (caret) return { x: caret.left, y: caret.bottom, trigger: 'keyboard' };
  const box = scroller.getBoundingClientRect();
  return { x: box.left + 16, y: box.top + 16, trigger: 'keyboard' };
}

const NO_HITS: readonly never[] = Object.freeze([]);

/** How far beside the caret line the keyboard probe reads, in CSS pixels. */
const CARET_PROBE = 1;

function ofSet<R extends HighlightRange>(
  hits: readonly HighlightHit<R>[],
  name: string | undefined
): readonly HighlightHit<R>[] {
  return name === undefined ? hits : hits.filter((hit) => hit.name === name);
}

/**
 * The highlights whose current range holds the caret, topmost first. The probe reads just
 * before and just after the caret, so a caret at either edge of a word also finds the word.
 */
function highlightsAtCaret<R extends HighlightRange>(editor: Editor): readonly HighlightHit<R>[] {
  const caret = editor.getSelectionRange()?.head;
  const rect = editor.getCaretRect();
  if (!caret || !rect) return NO_HITS;
  const y = rect.top + rect.height / 2;
  const seen = new Set<string>();
  const hits: HighlightHit<R>[] = [];
  for (const x of [rect.left - CARET_PROBE, rect.left + CARET_PROBE]) {
    for (const hit of editor.getHighlightsAt<R>(x, y)) {
      const id = `${hit.name}\u0000${hit.index}`;
      if (seen.has(id)) continue;
      seen.add(id);
      if (hit.range.blockId !== caret.blockId) continue;
      if (caret.offset < hit.start || caret.offset > hit.start + hit.length) continue;
      hits.push(hit);
    }
  }
  return hits;
}

/**
 * The highlights that the open menu is about, or an empty array while it is closed. A pointer
 * open reads under the pointer. A keyboard open reads at the caret.
 */
export function readContextMenuHighlights<R extends HighlightRange>(
  editor: Editor | null,
  anchor: ContextMenuAnchor | null,
  name: string | undefined
): readonly HighlightHit<R>[] {
  if (!editor || !anchor) return NO_HITS;
  if (anchor.trigger === 'keyboard') return ofSet(highlightsAtCaret<R>(editor), name);
  return ofSet(editor.getHighlightsAt<R>(anchor.x, anchor.y), name);
}
