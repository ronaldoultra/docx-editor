import { computed, inject, signal, type Signal } from '@angular/core';
import { DOCX_EDITOR_SCOPED, injectDocxEditorContext } from './context';

/** The color theme of the editor chrome. `'system'` follows the operating system. @public */
export type DocxEditorColorMode = 'light' | 'dark' | 'system';

/** The theme that a root resolved from its `colorMode`, never `'system'`. @public */
export type ResolvedColorMode = 'light' | 'dark';

const DARK_QUERY = '(prefers-color-scheme: dark)';

/**
 * Follow the operating system's dark setting. Calls `onChange` at once with the current value,
 * then on each change. Returns the function that stops it. Does nothing without `matchMedia`.
 *
 * @internal
 */
export function subscribeSystemDark(
  view: (Window & typeof globalThis) | null,
  onChange: (dark: boolean) => void
): () => void {
  if (typeof view?.matchMedia !== 'function') return () => {};
  const media = view.matchMedia(DARK_QUERY);
  onChange(media.matches);
  const listener = (event: MediaQueryListEvent) => onChange(event.matches);
  media.addEventListener('change', listener);
  return () => media.removeEventListener('change', listener);
}

/** True when `colorMode` is dark now. @internal */
export function resolveIsDark(colorMode: DocxEditorColorMode, systemDark: boolean): boolean {
  return colorMode === 'dark' || (colorMode === 'system' && systemDark);
}

/** What {@link injectEditorColorMode} returns. @public */
export interface EditorColorModeState {
  /** The resolved theme. It changes with the root's `colorMode` and the system setting. */
  readonly colorMode: Signal<ResolvedColorMode>;
}

const NOT_DARK: Signal<boolean> = signal(false).asReadonly();

/**
 * The theme that the enclosing editor root resolved, so that parts you build match the chrome.
 * Outside a root, or without `colorMode`, the theme is `'light'`. Call it in an injection
 * context.
 *
 * @public
 */
export function injectEditorColorMode(): EditorColorModeState {
  const dark = injectDocxEditorContext()?.dark ?? NOT_DARK;
  return { colorMode: computed(() => (dark() ? 'dark' : 'light')) };
}

/**
 * The scope classes for the root element of a chrome part: `'docx-editor '`, with `dark` under
 * a dark root, or `''` when an ancestor part already has the scope. The value ends with a space,
 * so that you can put your own classes after it. Call it in an injection context.
 *
 * @public
 */
export function injectScopeClassName(): Signal<'' | 'docx-editor ' | 'docx-editor dark '> {
  return scopeClassName(inject(DOCX_EDITOR_SCOPED));
}

/**
 * The scope classes of a part whose own providers set `DOCX_EDITOR_SCOPED` for its content:
 * it reads the token from its parent injector.
 *
 * @internal
 */
export function injectOwnScopeClassName(): Signal<'' | 'docx-editor ' | 'docx-editor dark '> {
  return scopeClassName(inject(DOCX_EDITOR_SCOPED, { skipSelf: true, optional: true }) ?? false);
}

function scopeClassName(scoped: boolean): Signal<'' | 'docx-editor ' | 'docx-editor dark '> {
  const dark = injectDocxEditorContext()?.dark ?? NOT_DARK;
  return computed(() => {
    if (scoped) return '';
    return dark() ? 'docx-editor dark ' : 'docx-editor ';
  });
}
