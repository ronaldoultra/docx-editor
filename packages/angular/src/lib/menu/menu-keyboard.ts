// Keyboard navigation for the menu bar, as DOM helpers and not as an item registry.
//
// The bar declares `menubar`, `menu`, and `menuitem`. A screen reader then gives the arrow keys
// to the widget, so the keys must work. The rows are found in the DOM, because a panel mixes
// registry rows, host overrides, and host content. Disabled rows are not skipped: they carry
// `aria-disabled` and stay focusable, so the engine's reason stays reachable.

const ITEM_SELECTOR = '[role="menuitem"], [role="menuitemcheckbox"], [role="menuitemradio"]';

/** The rows of one panel, in visual order, without the rows of nested panels. */
export function panelItems(panel: HTMLElement): HTMLElement[] {
  return [...panel.querySelectorAll<HTMLElement>(ITEM_SELECTOR)].filter(
    (item) => item.closest('[role="menu"]') === panel
  );
}

/** The triggers of a menu bar, in bar order. */
export function barTriggers(bar: HTMLElement): HTMLElement[] {
  return [...bar.querySelectorAll<HTMLElement>('.docx-menubar__trigger')];
}

/** Move the focus in a list by `step`, and wrap at both ends. */
export function focusBy(
  items: readonly HTMLElement[],
  from: Element | null,
  step: number
): boolean {
  if (items.length === 0) return false;
  const index = from ? items.indexOf(from as HTMLElement) : -1;
  const next =
    index === -1 ? (step > 0 ? 0 : items.length - 1) : (index + step + items.length) % items.length;
  items[next]?.focus();
  return true;
}

/** Focus the first or the last row of a list. */
export function focusEdge(items: readonly HTMLElement[], edge: 'first' | 'last'): boolean {
  const target = edge === 'first' ? items[0] : items[items.length - 1];
  if (!target) return false;
  target.focus();
  return true;
}

/**
 * True when a shortcut target is this editor's pages in another window than the bar. A host
 * can keep the bar in its main window and put the pages in a pop-out window. `editorElement` is
 * the mount element of this editor.
 */
export function isPoppedOutEditorTarget(
  target: Node,
  barDocument: Document,
  editorElement: Element | null
): boolean {
  if (!editorElement || editorElement.ownerDocument === barDocument) return false;
  return editorElement.contains(target);
}
