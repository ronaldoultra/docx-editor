// How a toolbar popup closes: on Escape, when the focus leaves it, and on a press outside it.
// Escape goes through core's `listenForPopupEscape`, the one rule for every toolbar popup. An
// Escape from this editor closes the popup and stops there, so the same key does not also leave
// a header, a note, or the format painter. An Escape from host content or from another editor
// closes the popup and keeps going.

import { NgZone, afterRenderEffect, inject, untracked, type Signal } from '@angular/core';
import {
  isDomNode,
  isElementNode,
  listenForOutsidePress,
  listenForPopupEscape,
  listenForPopupFocusLeave,
} from '@docx-editor.dev/core/editor';
import { injectDocxEditorContext } from '../context';

/** What {@link injectPopupDismiss} binds. */
export interface PopupDismissOptions {
  /** True while the popup is open. */
  readonly open: Signal<boolean>;
  /** The popup's root: the trigger and the list. */
  readonly root: () => HTMLElement | null;
  /** Close the popup. `fromInside` is true when the key came from inside the popup. */
  readonly close: (fromInside: boolean) => void;
  /**
   * `'picker'` also moves between `role="option"` rows with the arrow keys, Home, and End, and
   * gives the focus back to the trigger after an Escape from inside.
   */
  readonly kind?: 'picker' | 'popup';
  /** Return true to leave an Escape to a handler inside the popup. */
  readonly skip?: (event: KeyboardEvent) => boolean;
}

/** Bind the dismissal of one popup while it is open. Call it in an injection context. */
export function injectPopupDismiss(options: PopupDismissOptions): void {
  const zone = inject(NgZone);
  const context = injectDocxEditorContext();
  afterRenderEffect((onCleanup) => {
    const open = options.open();
    const root = options.root();
    if (!open || !root) return;
    untracked(() =>
      zone.runOutsideAngular(() => {
        const close = (fromInside: boolean) => zone.run(() => options.close(fromInside));
        const contains = (node: Node) => root.contains(node);
        const picker = options.kind === 'picker';
        const stopEscape = listenForPopupEscape({
          popup: root,
          contains,
          chromeRoot: () => root.closest('.docx-editor'),
          editorElements: () => [context?.contentElement()],
          skip: (event) => options.skip?.(event) === true,
          close: (fromInside) => {
            if (picker && fromInside) root.querySelector<HTMLElement>('[aria-haspopup]')?.focus();
            close(fromInside);
          },
        });
        const stopFocus = listenForPopupFocusLeave({
          popup: root,
          contains,
          close: () => close(false),
        });
        const stopPress = listenForOutsidePress({
          popup: root,
          contains,
          close: () => close(false),
        });
        const stopKeys = picker ? bindPickerKeys(root, () => close(false)) : () => {};
        onCleanup(() => {
          stopEscape();
          stopFocus();
          stopPress();
          stopKeys();
        });
      })
    );
  });
}

/** Arrow keys, Home, and End between the options of an open picker. */
function bindPickerKeys(root: HTMLElement, close: () => void): () => void {
  const keydown = (event: KeyboardEvent) => {
    if (event.defaultPrevented || event.altKey || event.ctrlKey || event.metaKey) return;
    const target = event.target;
    if (!isElementNode(target)) return;
    // The editable size input owns its draft and its step keys.
    if (target.matches('input[role="combobox"]')) return;
    const typing = target.matches('input, textarea, select');
    if (typing && event.key !== 'ArrowDown' && event.key !== 'ArrowUp') return;
    const options = Array.from(root.querySelectorAll<HTMLElement>('[role="option"]')).filter(
      (option) => !option.matches(':disabled, [aria-disabled="true"]')
    );
    if (options.length === 0) return;
    const current = options.indexOf(target as HTMLElement);
    let next: number;
    if (event.key === 'Home') next = 0;
    else if (event.key === 'End') next = options.length - 1;
    else if (event.key === 'ArrowDown') next = (current + 1) % options.length;
    else if (event.key === 'ArrowUp') {
      next = current < 0 ? options.length - 1 : (current - 1 + options.length) % options.length;
    } else return;
    event.preventDefault();
    event.stopPropagation();
    options[next]?.focus();
  };
  const focusout = (event: FocusEvent) => {
    if (isDomNode(event.relatedTarget) && !root.contains(event.relatedTarget)) close();
  };
  root.addEventListener('keydown', keydown);
  root.addEventListener('focusout', focusout);
  return () => {
    root.removeEventListener('keydown', keydown);
    root.removeEventListener('focusout', focusout);
  };
}

let nextPopupId = 0;

/** A unique element id for a popup list. */
export function popupId(prefix: string): string {
  nextPopupId += 1;
  return `docx-${prefix}-${nextPopupId}`;
}
