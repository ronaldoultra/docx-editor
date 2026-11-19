// Helpers for the context menu specs. Not part of the package build.
import { settle } from './harness';

/** The open context menu panel, or `null`. */
export function contextMenuPanel(root: Element): HTMLElement | null {
  return root.querySelector<HTMLElement>('[role="menu"].docx-contextmenu');
}

/** The rows of the open panel, nested submenu rows excluded. */
export function contextMenuRows(root: Element): HTMLElement[] {
  const panel = contextMenuPanel(root);
  if (!panel) return [];
  return [
    ...panel.querySelectorAll<HTMLElement>(
      '[role="menuitem"], [role="menuitemcheckbox"], [role="menuitemradio"]'
    ),
  ].filter((row) => row.closest('[role="menu"]') === panel);
}

/** The `data-slot` of each row of the open panel. */
export function contextMenuSlots(root: Element): (string | undefined)[] {
  return contextMenuRows(root).map((row) => row.dataset['slot']);
}

/** The rows and separators of the open panel, as text and `---`. */
export function contextMenuOutline(root: Element): string[] {
  const panel = contextMenuPanel(root);
  if (!panel) return [];
  return [...panel.querySelectorAll('[role^="menuitem"], [role="separator"]')]
    .filter((item) => item.closest('[role="menu"]') === panel)
    .map((item) =>
      item.getAttribute('role') === 'separator' ? '---' : (item.textContent ?? '').trim()
    );
}

/**
 * Dispatch a right-click on `target`, or the keyboard's context menu key when `keyboard` is
 * set. Returns the event, so a spec can read `defaultPrevented`.
 */
export async function rightClick(
  target: Element,
  point: { x: number; y: number } = { x: 120, y: 140 },
  keyboard = false
): Promise<MouseEvent> {
  const event = new MouseEvent('contextmenu', {
    bubbles: true,
    cancelable: true,
    button: keyboard ? -1 : 2,
    clientX: keyboard ? 0 : point.x,
    clientY: keyboard ? 0 : point.y,
  });
  target.dispatchEvent(event);
  await settle(undefined, 2);
  return event;
}

/** The scroller of the first editor under `root`. */
export function scrollerOf(root: Element): HTMLElement {
  const scroller = root.querySelector<HTMLElement>('.docx-editor__scroll-container');
  if (!scroller) throw new Error('no scroll container');
  return scroller;
}

/** The painted surface of the first editor, or its scroller before the surface mounts. */
export function surfaceOf(root: Element): HTMLElement {
  return root.querySelector<HTMLElement>('.docx-paginated-surface') ?? scrollerOf(root);
}

/** Press a key on `target`. */
export async function pressKey(target: Element, init: KeyboardEventInit): Promise<KeyboardEvent> {
  const event = new KeyboardEvent('keydown', { bubbles: true, cancelable: true, ...init });
  target.dispatchEvent(event);
  await settle(undefined, 2);
  return event;
}
