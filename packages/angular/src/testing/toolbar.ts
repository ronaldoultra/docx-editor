// Helpers for the toolbar specs. Not part of the package build.
import { chromeSlotId, defaultChromeGroups } from '@docx-editor.dev/core/editor';
import { settle } from './harness';

/** A ResizeObserver that reports only when a spec flushes it. */
export class MockResizeObserver {
  static readonly instances: MockResizeObserver[] = [];
  readonly observed: Element[] = [];
  private readonly callback: ResizeObserverCallback;

  constructor(callback: ResizeObserverCallback) {
    this.callback = callback;
    MockResizeObserver.instances.push(this);
  }

  observe(element: Element): void {
    this.observed.push(element);
  }

  unobserve(): void {}

  disconnect(): void {
    const index = MockResizeObserver.instances.indexOf(this);
    if (index >= 0) MockResizeObserver.instances.splice(index, 1);
  }

  flush(): void {
    this.callback([], this as unknown as ResizeObserver);
  }
}

type ObserverWindow = { ResizeObserver?: unknown };

const windows = (): ObserverWindow[] =>
  [globalThis as ObserverWindow, document.defaultView as unknown as ObserverWindow].filter(
    (view, index, all) => view && all.indexOf(view) === index
  );

const saved = new Map<ObserverWindow, unknown>();

/** Install the mock observer in the spec's window. */
export function useMockResizeObserver(): void {
  for (const view of windows()) {
    if (!saved.has(view)) saved.set(view, view.ResizeObserver);
    view.ResizeObserver = MockResizeObserver;
  }
}

/** Put the real observer back and forget the mock instances. */
export function restoreResizeObserver(): void {
  for (const [view, value] of saved) view.ResizeObserver = value;
  saved.clear();
  MockResizeObserver.instances.length = 0;
}

/** Run the pending observer callbacks, then wait for a frame and for Angular. */
export async function flushObservers(): Promise<void> {
  for (const observer of [...MockResizeObserver.instances]) observer.flush();
  await new Promise<void>((resolve) => requestAnimationFrame(() => resolve()));
  await settle(undefined, 2);
}

/** Give an element a fixed layout width. */
export function setWidth(
  element: Element,
  property: 'clientWidth' | 'offsetWidth',
  width: number | (() => number)
): void {
  const get = typeof width === 'function' ? width : () => width;
  Object.defineProperty(element, property, { configurable: true, get });
}

/** Give the bar a width and each group a width, then measure. */
export async function measure(bar: HTMLElement, barWidth: number, groupWidth = 90): Promise<void> {
  setWidth(bar, 'clientWidth', barWidth);
  for (const element of bar.querySelectorAll('[data-toolbar-group], [data-toolbar-fixed]')) {
    setWidth(element, 'offsetWidth', groupWidth);
  }
  await flushObservers();
}

/** The toolbar of a container. */
export const toolbarOf = (container: Element): HTMLElement =>
  container.querySelector<HTMLElement>('[data-testid="docx-toolbar"]')!;

/** The open "⋯" panel, or `null`. */
export const panelOf = (container: Element): HTMLElement | null =>
  container.querySelector<HTMLElement>('[data-testid="toolbar-overflow-panel"]');

/** The bar's children that render: separators as `|`, groups by id, fixed blocks as `fixed`. */
export function barShape(toolbar: Element): string[] {
  return [...toolbar.children]
    .filter((child) => !child.hasAttribute('hidden'))
    .map((child) =>
      child.getAttribute('role') === 'separator'
        ? '|'
        : (child.getAttribute('data-toolbar-group') ??
          (child.hasAttribute('data-toolbar-fixed') ? 'fixed' : child.className))
    );
}

/** The bar's controls by slot, with `separator` between groups. */
export function toolbarArrangement(toolbar: Element): string[] {
  return [...toolbar.children]
    .filter((child) => !child.hasAttribute('hidden'))
    .flatMap((child) => {
      if (child.getAttribute('role') === 'separator') return 'separator';
      if (
        child.classList.contains('docx-toolbar__group') ||
        child.classList.contains('docx-toolbar__contextual')
      ) {
        // Each control sits in its own `docx-toolbar__item` collapse wrapper.
        return [...child.children]
          .flatMap((entry) =>
            entry.classList.contains('docx-toolbar__item') ? [...entry.children] : [entry]
          )
          .map(
            (entry) =>
              entry.getAttribute('data-slot') ?? entry.getAttribute('aria-label') ?? entry.className
          );
      }
      return child.getAttribute('data-slot') ?? child.getAttribute('aria-label') ?? child.className;
    });
}

/** The default arrangement, as `toolbarArrangement` reads it. */
export function expectedArrangement(alignment: string): string[] {
  return defaultChromeGroups().flatMap((group, index) => [
    ...(index > 0 ? ['separator'] : []),
    ...(group.id === 'alignment'
      ? [alignment]
      : group.controls.map((control) => chromeSlotId(group, control) as string)),
  ]);
}

/** Dispatch a cancelable Escape keydown on `target`, then wait for Angular. */
export async function pressEscape(target: EventTarget): Promise<KeyboardEvent> {
  const event = new KeyboardEvent('keydown', { key: 'Escape', bubbles: true, cancelable: true });
  target.dispatchEvent(event);
  await settle(undefined, 2);
  return event;
}
