import {
  Injectable,
  InjectionToken,
  computed,
  inject,
  signal,
  type Provider,
  type Signal,
  type Type,
} from '@angular/core';
import type {
  ContentControlWidgetSession,
  InvalidTextFormFieldSession,
} from '@docx-editor.dev/core/editor';

/**
 * A popup component with static inputs. The host renders `component` and sets each entry of
 * `inputs` on it, beside the inputs that the popup gets from the editor, such as `session`.
 * The editor's own inputs win a collision.
 *
 * @public
 */
export interface DocxEditorPopupEntry {
  /** The standalone component to render. */
  readonly component: Type<unknown>;
  /** Static inputs for the component. */
  readonly inputs?: Readonly<Record<string, unknown>>;
}

/**
 * One entry of the popup registry: a standalone component, a component with static inputs, or
 * `false` to turn the popup off.
 *
 * The editor renders the component with `NgComponentOutlet` and sets the inputs in `Inputs`
 * on it. The component reads the editor with `injectDocxEditor()`. A new registry object with the
 * same component keeps the rendered instance and its state. A new session renders a new
 * instance, so no draft of the earlier session stays.
 *
 * @public
 */
// `Inputs` documents the inputs that the editor sets; a component class cannot carry them.
export type DocxEditorPopup<Inputs extends object = Record<string, never>> =
  false | Type<unknown> | DocxEditorPopupEntry;

/**
 * The popups that the editor hosts, by id. An omitted entry keeps the default behavior: the
 * packaged default popup, or the engine's own UI. `false` turns a popup off.
 *
 * @public
 */
export interface DocxEditorPopups {
  /**
   * The value picker of a drop-down list, combo box, date, or building block gallery content
   * control. Inputs: `session`.
   */
  contentControlWidget?: DocxEditorPopup<{ session: ContentControlWidgetSession }>;
  /**
   * The renderer for checkbox presses. Omitted, the engine toggles the box itself and
   * `contentControlWidget` never sees a checkbox session. Inputs: `session`.
   */
  contentControlCheckbox?: DocxEditorPopup<{ session: ContentControlWidgetSession }>;
  /**
   * The renderer for picture presses. Omitted, the engine opens its own file picker and
   * `contentControlWidget` never sees a picture session. Inputs: `session`.
   */
  contentControlPicture?: DocxEditorPopup<{ session: ContentControlWidgetSession }>;
  /** The message for a value that a protected text form field refuses. Inputs: `session`. */
  invalidTextFormField?: DocxEditorPopup<{ session: InvalidTextFormFieldSession }>;
  /** The hyperlink popup. It registers its own link UI with the engine. */
  hyperlink?: DocxEditorPopup;
  /** The content control popup. */
  contentControl?: DocxEditorPopup;
  /** The equation popup. It registers its own equation UI with the engine. */
  equation?: DocxEditorPopup;
  /** A context menu that the editor hosts beside the pages. */
  contextMenu?: DocxEditorPopup;
  /**
   * The page setup dialog that File › Page setup opens. Inputs: `closeHandler`. Omitted, the
   * editor renders `<docx-editor-page-setup-dialog>`. `false` opens nothing.
   */
  pageSetup?: DocxEditorPopup<{ closeHandler: () => void }>;
  /**
   * The paragraph dialog that Format › Paragraph and the line spacing menu open. Inputs:
   * `closeHandler`. Omitted, the editor renders `<docx-editor-paragraph-dialog>`. `false` opens
   * nothing.
   */
  paragraph?: DocxEditorPopup<{ closeHandler: () => void }>;
  /**
   * The watermark dialog that Insert › Watermark opens. Inputs: `closeHandler`. Omitted, the
   * editor renders `<docx-editor-watermark-dialog>`. `false` opens nothing.
   */
  watermark?: DocxEditorPopup<{ closeHandler: () => void }>;
}

/**
 * The popup registry for every editor root below the provider. A root's `popups` input replaces
 * single entries of it.
 *
 * @example
 * ```ts
 * providers: [provideDocxEditorPopups({ hyperlink: MyLinkPopupComponent })]
 * ```
 *
 * @public
 */
export const DOCX_EDITOR_POPUPS = new InjectionToken<DocxEditorPopups>('DOCX_EDITOR_POPUPS');

/** Provide the popup registry for the editor roots below. @public */
export function provideDocxEditorPopups(popups: DocxEditorPopups): Provider {
  return { provide: DOCX_EDITOR_POPUPS, useValue: popups };
}

/**
 * The registry of one editor root: the provided `DOCX_EDITOR_POPUPS`, with the root's `popups`
 * input over it.
 *
 * @internal
 */
@Injectable()
export class DocxEditorPopupRegistry {
  private readonly provided = inject(DOCX_EDITOR_POPUPS, { optional: true });
  private readonly source = signal<() => DocxEditorPopups | undefined>(() => undefined);

  /** The resolved registry. */
  readonly popups: Signal<DocxEditorPopups> = computed(() => {
    const own = this.source()();
    if (!own) return this.provided ?? {};
    return this.provided ? { ...this.provided, ...own } : own;
  });

  /** Read the root's own entries from `source`, which can read signals. */
  use(source: () => DocxEditorPopups | undefined): void {
    this.source.set(source);
  }
}

const NO_POPUPS: Signal<DocxEditorPopups> = signal<DocxEditorPopups>({}).asReadonly();

/**
 * The popup registry of the enclosing editor root, as a signal. Outside a root, it holds the
 * provided `DOCX_EDITOR_POPUPS`, or an empty registry. Call it in an injection context.
 *
 * @public
 */
export function injectDocxEditorPopups(): Signal<DocxEditorPopups> {
  const registry = inject(DocxEditorPopupRegistry, { optional: true });
  if (registry) return registry.popups;
  const provided = inject(DOCX_EDITOR_POPUPS, { optional: true });
  return provided ? signal(provided).asReadonly() : NO_POPUPS;
}

/** The component of an entry, or `null` for an omitted or turned-off entry. @internal */
export function popupComponent(entry: DocxEditorPopup<object> | undefined): Type<unknown> | null {
  if (!entry) return null;
  return typeof entry === 'function' ? entry : entry.component;
}

/** The static inputs of an entry with the editor's inputs over them. @internal */
export function popupInputs(
  entry: DocxEditorPopup<object> | undefined,
  inputs: Readonly<Record<string, unknown>>
): Record<string, unknown> {
  const own = entry && typeof entry !== 'function' ? entry.inputs : undefined;
  return own ? { ...own, ...inputs } : { ...inputs };
}
