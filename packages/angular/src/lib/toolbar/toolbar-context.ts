import { Injectable, InjectionToken, computed, inject, signal, type Signal } from '@angular/core';
import type { TranslationKey } from '@docx-editor.dev/core/i18n';
import { injectTranslation } from '../i18n';

/**
 * A translation function for the toolbar labels. Return a string to replace the label of a
 * key. The catalog of the editor root gives every other label.
 *
 * @public
 */
export type DocxToolbarTranslate = (key: string) => string | undefined;

/**
 * Where a toolbar control renders: in the bar, or as a row of the "⋯" panel.
 *
 * @public
 */
export type DocxToolbarPlacement = 'bar' | 'panel';

/**
 * The placement of the controls below this injector. The toolbar provides `'panel'` for the
 * rows of its "⋯" panel. Without a provider, the placement is `'bar'`.
 *
 * @public
 */
export const DOCX_TOOLBAR_PLACEMENT = new InjectionToken<DocxToolbarPlacement>(
  'DOCX_TOOLBAR_PLACEMENT',
  { factory: () => 'bar' }
);

/**
 * The state that one toolbar shares with its controls: its own labels and its "⋯" panel.
 *
 * @internal
 */
@Injectable()
export class DocxToolbarState {
  private readonly translateSource = signal<() => DocxToolbarTranslate | undefined>(
    () => undefined
  );

  /** The toolbar's own translation function, from its `t` input. */
  readonly translate: Signal<DocxToolbarTranslate | undefined> = computed(() =>
    this.translateSource()()
  );

  /** True while the "⋯" panel is open. */
  readonly moreOpen = signal(false);

  /** The "⋯" trigger, while it is in the bar. */
  moreTrigger: HTMLElement | null = null;

  /** Read the toolbar's translation function from `source`, which can read signals. */
  useTranslate(source: () => DocxToolbarTranslate | undefined): void {
    this.translateSource.set(source);
  }

  /**
   * Close the "⋯" panel. `focusTrigger` moves the focus to the trigger, which a keyboard
   * activation needs, because the focused row leaves the page.
   */
  closeMore(focusTrigger: boolean): void {
    this.moreOpen.set(false);
    if (focusTrigger) this.moreTrigger?.focus();
  }
}

/**
 * A function that gives the toolbar label of a key: the toolbar's `t` first, then the active
 * catalog. Call it in an injection context.
 *
 * @public
 */
export function injectToolbarLabel(): Signal<(key: string) => string> {
  const state = inject(DocxToolbarState, { optional: true });
  const { t } = injectTranslation();
  return computed(() => {
    const own = state?.translate();
    const translate = t();
    return (key: string) => own?.(key) ?? translate(key as TranslationKey);
  });
}
