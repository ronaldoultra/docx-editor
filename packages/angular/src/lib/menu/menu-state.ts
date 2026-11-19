// The state that one menu bar shares with its menus and rows: which menu is open, the host's
// handlers, its own labels, and which menus moved into the "⋯" menu.

import {
  Injectable,
  InjectionToken,
  computed,
  inject,
  signal,
  type Signal,
  type TemplateRef,
} from '@angular/core';
import type { TranslationKey } from '@docx-editor.dev/core/i18n';
import { injectTranslation } from '../i18n';

/**
 * A translation function for the menu labels. Return a string to replace the label of a key.
 * The catalog of the editor root gives every other label.
 *
 * @public
 */
export type DocxMenuTranslate = (key: string) => string | undefined;

/** The id of the "⋯" menu that holds the menus that do not fit. */
export const MENU_OVERFLOW_ID = 'docx-menubar-more';

/**
 * The tab index of the menu rows below this injector. The menu bar provides `0`, so a row
 * directly in the bar is a tab stop. A menu panel provides `-1`, where the arrow keys move.
 *
 * @internal
 */
export const DOCX_MENU_ROW_TABINDEX = new InjectionToken<number>('DOCX_MENU_ROW_TABINDEX', {
  factory: () => -1,
});

/** The handlers that the rows of the bar run. `undefined` disables a row. @internal */
export interface MenuHandlers {
  readonly open?: () => void;
  readonly save?: () => void;
  readonly pageSetup?: (trigger: HTMLElement | null) => void;
  readonly paragraphDialog?: (trigger: HTMLElement | null) => void;
  readonly watermarkDialog?: (trigger: HTMLElement | null) => void;
  readonly reportIssue?: () => void;
  /** `false` removes Help › Report issue, and the Help menu when it has no other content. */
  readonly reportIssueRow: boolean;
}

/** One menu of the bar, for the "⋯" menu: its label and its rows. @internal */
export interface MenuRegistration {
  readonly id: string;
  readonly text: Signal<string>;
  readonly rows: Signal<TemplateRef<unknown> | undefined>;
}

const NO_HANDLERS: MenuHandlers = { reportIssueRow: true };

/** @internal */
@Injectable()
export class DocxMenuBarState {
  private readonly translateSource = signal<() => DocxMenuTranslate | undefined>(() => undefined);
  private readonly handlersSource = signal<() => MenuHandlers>(() => NO_HANDLERS);
  private readonly openingSource = signal<() => boolean>(() => false);
  private readonly registered = signal<readonly MenuRegistration[]>([]);

  /** The menu bar's own translation function. */
  readonly translate: Signal<DocxMenuTranslate | undefined> = computed(() =>
    this.translateSource()()
  );
  /** The handlers that the rows run. */
  readonly handlers: Signal<MenuHandlers> = computed(() => this.handlersSource()());
  /** True while the next document opens: the triggers are disabled. */
  readonly opening: Signal<boolean> = computed(() => this.openingSource()());
  /** The id of the open menu, or `null`. */
  readonly openMenu = signal<string | null>(null);
  /** The id of the menu whose trigger holds the bar's tab stop. */
  readonly activeMenu = signal<string | null>(null);
  /** True when the bar measures its menus. */
  readonly measuring = signal(false);
  /** The ids of the menus that render inside the "⋯" menu. */
  readonly overflow = signal<ReadonlySet<string>>(new Set());
  /** The menus of the bar, in registration order. */
  readonly menus: Signal<readonly MenuRegistration[]> = this.registered.asReadonly();

  /** Open a menu, or close the open menu with `null`. */
  setOpenMenu(id: string | null): void {
    this.openMenu.set(id);
    if (id !== null) this.activeMenu.set(id);
  }

  /** @internal */
  useTranslate(source: () => DocxMenuTranslate | undefined): void {
    this.translateSource.set(source);
  }

  /** @internal */
  useHandlers(source: () => MenuHandlers): void {
    this.handlersSource.set(source);
  }

  /** @internal */
  useOpening(source: () => boolean): void {
    this.openingSource.set(source);
  }

  /** Add a menu to the bar. Returns a function that removes it. @internal */
  register(menu: MenuRegistration): () => void {
    this.registered.update((menus) => [...menus, menu]);
    return () => this.registered.update((menus) => menus.filter((entry) => entry !== menu));
  }
}

/**
 * The menu bar state of the enclosing bar, or a new state for a menu outside a bar.
 *
 * @internal
 */
export function provideMenuBarState(): DocxMenuBarState {
  return inject(DocxMenuBarState, { optional: true, skipSelf: true }) ?? new DocxMenuBarState();
}

/**
 * A function that gives the menu label of a key: the bar's `t` first, then the active catalog.
 * Call it in an injection context.
 *
 * @internal
 */
export function injectMenuLabel(): Signal<(key: string) => string> {
  const state = inject(DocxMenuBarState, { optional: true });
  const { t } = injectTranslation();
  return computed(() => {
    const own = state?.translate();
    const translate = t();
    return (key: string) => own?.(key) ?? translate(key as TranslationKey);
  });
}

const ISSUE_URL = 'https://github.com/eigenpal/docx-editor/issues/new';

/**
 * The prefilled issue form of Help › Report issue. The page URL has no query string and no
 * fragment, because hosts keep tokens and ids there, and the issue form is public.
 */
export function reportIssueUrl(view: Window): string {
  let page = '';
  try {
    const url = new URL(view.location.href);
    page = `${url.origin}${url.pathname}`;
  } catch {
    page = '';
  }
  const body = [
    '### What happened',
    '',
    '### Steps to reproduce',
    '1. ',
    '2. ',
    '3. ',
    '',
    '### Expected',
    '',
    '### Actual',
    '',
    '### Attach the DOCX',
    'Please drag the `.docx` file that reproduces this onto the issue (below this text box). Bugs are much faster to fix with a real file.',
    '',
    '### Environment',
    `- URL: ${page}`,
    `- Viewport: ${view.innerWidth} x ${view.innerHeight}`,
    `- User agent: ${view.navigator.userAgent}`,
    '',
  ].join('\n');
  return `${ISSUE_URL}?${new URLSearchParams({ title: '[Bug] ', body }).toString()}`;
}
