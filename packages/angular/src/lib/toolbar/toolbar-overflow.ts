// The toolbar's own measurement: how wide each control is, and how much room the bar has. The
// reading (`readToolbarRow`) and the fit (`toolbarOverflowItems`) live in core, so every adapter
// collapses the same controls at the same widths.

import {
  DestroyRef,
  NgZone,
  afterRenderEffect,
  inject,
  signal,
  untracked,
  type Signal,
} from '@angular/core';
import {
  hostWindowOf,
  readToolbarRow,
  sameItemOverflow,
  TOOLBAR_OVERFLOW_HYSTERESIS,
  toolbarGroupsKey,
  toolbarItemKey,
  toolbarOverflowItems,
  type HostWindow,
  type ToolbarFitGroup,
  type ToolbarItemOverflow,
  type ToolbarRowMemory,
} from '@docx-editor.dev/core/editor';
import { FIXED_ATTRIBUTE, GROUP_ATTRIBUTE } from './toolbar-registry';

const ASSUMED_MORE_WIDTH = 34;
const NONE: ToolbarItemOverflow = new Map<string, number>();

/** What {@link injectToolbarOverflow} reads. */
export interface ToolbarOverflowSource {
  /** The bar element, once it renders. */
  readonly bar: () => HTMLElement | null;
  /** True when the bar measures itself. */
  readonly enabled: Signal<boolean>;
  /** The collapsible groups and their controls, in bar order. */
  readonly groups: Signal<readonly ToolbarFitGroup[]>;
  /** The collapse order of the group ids. */
  readonly order: Signal<readonly string[]>;
}

/**
 * Measure the bar after each render that changes its groups, and on each resize of the bar,
 * its parent, or a group. The signal holds, for each group id, how many controls from the end
 * of the group render in the "⋯" panel. Call it in an injection context.
 */
export function injectToolbarOverflow(source: ToolbarOverflowSource): Signal<ToolbarItemOverflow> {
  const zone = inject(NgZone);
  const overflow = signal<ToolbarItemOverflow>(NONE);
  // Widths SURVIVE a control leaving the bar. See `ToolbarRowMemory`.
  const memory: ToolbarRowMemory = {
    itemWidths: new Map(),
    groupBases: new Map(),
    groupGaps: new Map(),
    more: ASSUMED_MORE_WIDTH,
  };
  let frame = 0;
  let frameWindow: HostWindow | null = null;
  let observer: ResizeObserver | undefined;
  // The controls that the previous answer counted from. A count from the end names other
  // controls once a group gains or loses one, so the slack applies only to the same controls.
  let answeredKey = '';

  const measure = () => {
    const bar = source.bar();
    if (!bar || !untracked(source.enabled)) return;
    const groups = untracked(source.groups);
    const current = untracked(overflow);
    const reading = readToolbarRow(bar, memory, toolbarItemKey);
    // The arithmetic charges each group a few px of gap and separator that it may not use, so
    // on a bar sized to its content it can call a bar that fits too narrow. So the first
    // collapse waits for a control that really runs past the bar's box. Without layout to
    // read (null), the arithmetic decides alone.
    const fits = current.size === 0 && reading.overflowing === false;
    const key = toolbarGroupsKey(groups);
    const same = answeredKey === key;
    answeredKey = key;
    const next = fits
      ? NONE
      : toolbarOverflowItems({
          ...reading,
          groups,
          order: untracked(source.order),
          itemWidths: memory.itemWidths,
          groupBases: memory.groupBases,
          groupGaps: memory.groupGaps,
          ...(same ? { previous: current } : {}),
          hysteresis: TOOLBAR_OVERFLOW_HYSTERESIS,
        });
    if (!sameItemOverflow(next, current)) overflow.set(next);
  };

  const stopObserver = () => {
    observer?.disconnect();
    observer = undefined;
    if (frame !== 0) {
      frameWindow?.cancelAnimationFrame(frame);
      frame = 0;
    }
  };

  const observe = (bar: HTMLElement) => {
    stopObserver();
    // The bar's own window: a toolbar in a pop-out window is observed and measured there.
    const win = hostWindowOf(bar);
    if (typeof win.ResizeObserver === 'undefined') return;
    observer = new win.ResizeObserver(() => {
      if (frame !== 0) return;
      frameWindow = win;
      frame = win.requestAnimationFrame(() => {
        frame = 0;
        measure();
      });
    });
    observer.observe(bar);
    // The parent too: a bar as wide as its content does not resize when the room around it
    // grows, so only the parent reports the space that a collapsed control can come back into.
    if (bar.parentElement) observer.observe(bar.parentElement);
    for (const element of bar.querySelectorAll<HTMLElement>(
      `[${GROUP_ATTRIBUTE}], [${FIXED_ATTRIBUTE}]`
    )) {
      observer.observe(element);
    }
  };

  afterRenderEffect(() => {
    const enabled = source.enabled();
    toolbarGroupsKey(source.groups());
    source.order();
    overflow();
    const bar = source.bar();
    untracked(() =>
      zone.runOutsideAngular(() => {
        if (!enabled || !bar) {
          stopObserver();
          if (enabled === false && overflow().size > 0) overflow.set(NONE);
          return;
        }
        measure();
        observe(bar);
      })
    );
  });

  inject(DestroyRef).onDestroy(stopObserver);
  return overflow.asReadonly();
}
