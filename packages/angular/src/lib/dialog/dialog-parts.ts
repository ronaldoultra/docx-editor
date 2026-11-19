import {
  Directive,
  EmbeddedViewRef,
  Injectable,
  Renderer2,
  TemplateRef,
  ViewContainerRef,
  booleanAttribute,
  computed,
  effect,
  inject,
  input,
  signal,
  untracked,
  type Signal,
} from '@angular/core';

/**
 * The draft state and the actions of a packaged dialog. Inject it with `injectDialog()` in a
 * part template or in a component inside the dialog.
 *
 * @public
 */
export interface DocxDialogState<T extends object> {
  /** The draft values. */
  readonly values: Signal<T>;
  /** The errors by field, and `form` for an error of the whole dialog. */
  readonly errors: Signal<Readonly<Partial<Record<keyof T | 'form', string>>>>;
  /** True when the engine can apply the draft. */
  readonly isEnabled: Signal<boolean>;
  /** Change one draft value. */
  setValue<K extends keyof T>(name: K, value: T[K]): void;
  /** Apply the draft. The dialog closes when the engine accepts it. */
  apply(): void;
  /** Close the dialog without a write. */
  cancel(): void;
  /** Set the draft back to the values that the dialog opened with. */
  reset?(): void;
}

/**
 * The presentation inputs of a dialog part template (`docxDialogPart`).
 *
 * @public
 */
export interface DialogPartProps {
  /** Classes to add to the part's element. The packaged classes stay. */
  className?: string;
  /** Inline styles to add to the part's element. */
  styles?: Readonly<Record<string, string>>;
  /**
   * Attributes for the part's element, such as `id`, `aria-label`, or `data-testid`. The part's
   * own wiring wins a collision: `type`, `disabled`, `class`, `style`, `data-docx-part`,
   * `data-docx-field`, and event handler attributes do not apply.
   */
  attributes?: Readonly<Record<string, string>>;
  /** Remove the part. */
  hidden?: boolean;
  /**
   * Render the template in the place of the part's element. Put `docxDialogApply`,
   * `docxDialogCancel`, `docxDialogReset`, or `docxDialogField` on your own element to wire it.
   */
  asChild?: boolean;
  /** Keep the packaged content and apply only the presentation inputs. */
  keepContent?: boolean;
}

/**
 * The layout inputs of a packaged dialog.
 *
 * @public
 */
export interface DialogCustomizationProps {
  /** Classes for the dialog element. */
  className?: string;
  /** Inline styles for the dialog element. */
  styles?: Readonly<Record<string, string>>;
  /**
   * With `true` (the default), the dialog renders its packaged parts, the part templates in
   * their place, and the other content after them. With `false`, it renders only its content.
   */
  preset?: boolean;
}

/**
 * Replace or restyle one part of a packaged dialog. Put the template inside the dialog. The part
 * names are `header`, `title`, `body`, `footer`, `apply`, `cancel`, `reset`, `error`,
 * `field:<name>` for a field, and the names of a dialog's own parts.
 *
 * The template replaces the content of the part's element, and the element keeps its wiring.
 * With `asChild`, the template replaces the element.
 *
 * @example
 * ```html
 * <docx-editor-text-form-field-dialog [session]="session">
 *   <ng-template docxDialogPart="apply" className="app-action">Save <strong>field</strong></ng-template>
 *   <ng-template docxDialogPart="cancel" asChild>
 *     <button docxDialogCancel class="app-link">Close</button>
 *   </ng-template>
 * </docx-editor-text-form-field-dialog>
 * ```
 *
 * @public
 */
@Directive({ selector: 'ng-template[docxDialogPart]' })
export class DocxDialogPartDirective {
  /** The name of the part. */
  readonly docxDialogPart = input.required<string>();
  /** See {@link DialogPartProps.className}. */
  readonly className = input<string | undefined>(undefined);
  /** See {@link DialogPartProps.styles}. */
  readonly styles = input<Readonly<Record<string, string>> | undefined>(undefined);
  /** See {@link DialogPartProps.attributes}. */
  readonly attributes = input<Readonly<Record<string, string>> | undefined>(undefined);
  /** See {@link DialogPartProps.hidden}. */
  readonly hidden = input(false, { transform: booleanAttribute });
  /** See {@link DialogPartProps.asChild}. */
  readonly asChild = input(false, { transform: booleanAttribute });
  /** See {@link DialogPartProps.keepContent}. */
  readonly keepContent = input(false, { transform: booleanAttribute });

  /** The template. */
  readonly template: TemplateRef<unknown> = inject(TemplateRef);
}

/**
 * The state and the part templates of one packaged dialog. Each packaged dialog provides one
 * and connects it in its constructor.
 *
 * @internal
 */
@Injectable()
export class DocxDialogParts {
  private current: DocxDialogState<object> | null = null;
  private overrides: Signal<readonly DocxDialogPartDirective[]> = signal([]);

  /** The dialog state. */
  get state(): DocxDialogState<object> | null {
    return this.current;
  }

  /** Connect the dialog's state and its `contentChildren(DocxDialogPartDirective)` query. */
  connect(
    state: DocxDialogState<object>,
    overrides: Signal<readonly DocxDialogPartDirective[]>
  ): void {
    this.current = state;
    this.overrides = overrides;
  }

  /** The last template for the part `name`. */
  override(name: string): DocxDialogPartDirective | undefined {
    const all = this.overrides();
    for (let index = all.length - 1; index >= 0; index -= 1) {
      if (all[index]!.docxDialogPart() === name) return all[index];
    }
    return undefined;
  }
}

/**
 * The state of the enclosing packaged dialog. Call it in an injection context inside the dialog,
 * such as in a component in a part template.
 *
 * @public
 */
export function injectDialog<T extends object>(): DocxDialogState<T> {
  const state = inject(DocxDialogParts, { optional: true })?.state;
  if (!state) throw new Error('injectDialog() must run inside a packaged dialog.');
  return state as unknown as DocxDialogState<T>;
}

function dialogState(): DocxDialogState<object> | null {
  return inject(DocxDialogParts, { optional: true })?.state ?? null;
}

/**
 * Wire an element as the Apply action of the enclosing dialog: a press applies the draft, and
 * the element is disabled while the engine cannot apply it.
 *
 * @public
 */
@Directive({
  selector: '[docxDialogApply]',
  host: {
    type: 'button',
    'data-docx-part': 'apply',
    '[attr.disabled]': 'enabled() ? null : ""',
    '(click)': 'press()',
  },
})
export class DocxDialogApplyDirective {
  private readonly state = dialogState();
  /** @internal */
  protected enabled(): boolean {
    return this.state?.isEnabled() ?? false;
  }
  /** @internal */
  protected press(): void {
    if (this.enabled()) this.state?.apply();
  }
}

/** Wire an element as the Cancel action of the enclosing dialog. @public */
@Directive({
  selector: '[docxDialogCancel]',
  host: { type: 'button', 'data-docx-part': 'cancel', '(click)': 'press()' },
})
export class DocxDialogCancelDirective {
  private readonly state = dialogState();
  /** @internal */
  protected press(): void {
    this.state?.cancel();
  }
}

/** Wire an element as the Reset action of the enclosing dialog. @public */
@Directive({
  selector: '[docxDialogReset]',
  host: {
    type: 'button',
    'data-docx-part': 'reset',
    '[attr.disabled]': 'enabled() ? null : ""',
    '(click)': 'press()',
  },
})
export class DocxDialogResetDirective {
  private readonly state = dialogState();
  /** @internal */
  protected enabled(): boolean {
    return (this.state?.isEnabled() ?? false) && !!this.state?.reset;
  }
  /** @internal */
  protected press(): void {
    if (this.enabled()) this.state?.reset?.();
  }
}

/** Mark an element as the field `name` of the enclosing dialog. @public */
@Directive({
  selector: '[docxDialogField]',
  host: { 'data-docx-part': 'field', '[attr.data-docx-field]': 'docxDialogField()' },
})
export class DocxDialogFieldDirective {
  /** The field name. */
  readonly docxDialogField = input.required<string>();
}

/** The context of a part's default template: the host's content template, if any. @internal */
export interface DocxDialogSlotContext {
  $implicit: TemplateRef<unknown> | null;
}

const PROTECTED_ATTRIBUTES = new Set([
  'type',
  'disabled',
  'class',
  'style',
  'data-docx-part',
  'data-docx-field',
]);

/** True for an attribute that a host may set on a part's element. */
function allowedAttribute(name: string): boolean {
  const lower = name.toLowerCase();
  return !PROTECTED_ATTRIBUTES.has(lower) && !lower.startsWith('on');
}

/**
 * Render a dialog part: the host's template for it, or the packaged default with the host's
 * presentation inputs. Packaged dialogs only.
 *
 * ```html
 * <button *docxDialogSlot="'apply'; let content" docxDialogApply class="docx-dialog__button">
 *   @if (content) { <ng-container [ngTemplateOutlet]="content" /> } @else { Apply }
 * </button>
 * ```
 *
 * @internal
 */
@Directive({ selector: '[docxDialogSlot]' })
export class DocxDialogSlotDirective {
  readonly docxDialogSlot = input.required<string>();

  private readonly parts = inject(DocxDialogParts);
  private readonly template = inject<TemplateRef<DocxDialogSlotContext>>(TemplateRef);
  private readonly container = inject(ViewContainerRef);
  private readonly renderer = inject(Renderer2);
  private readonly element = signal<Element | null>(null);
  private rendered: readonly unknown[] = [];
  private applied: { classes: string[]; styles: string[]; attributes: string[] } = {
    classes: [],
    styles: [],
    attributes: [],
  };

  constructor() {
    const part = computed(() => this.parts.override(this.docxDialogSlot()));
    effect(() => {
      const override = part();
      const key = [
        override,
        override?.hidden() ?? false,
        override?.asChild() ?? false,
        override?.keepContent() ?? false,
      ];
      untracked(() => this.render(key));
    });
    effect(() => {
      const element = this.element();
      const override = part();
      const presentation = {
        classes: (override?.className() ?? '').split(/\s+/).filter(Boolean),
        styles: override?.styles() ?? {},
        attributes: override?.attributes() ?? {},
      };
      untracked(() => this.present(element, presentation));
    });
  }

  private render(key: readonly unknown[]): void {
    if (key.every((value, index) => value === this.rendered[index])) return;
    this.rendered = key;
    this.container.clear();
    this.applied = { classes: [], styles: [], attributes: [] };
    const [override, hidden, asChild, keepContent] = key as [
      DocxDialogPartDirective | undefined,
      boolean,
      boolean,
      boolean,
    ];
    if (hidden) {
      this.element.set(null);
      return;
    }
    if (override && asChild) {
      this.container.createEmbeddedView(override.template);
      this.element.set(null);
      return;
    }
    const view: EmbeddedViewRef<DocxDialogSlotContext> = this.container.createEmbeddedView(
      this.template,
      { $implicit: override && !keepContent ? override.template : null }
    );
    const element = view.rootNodes.find((node): node is Element => node.nodeType === 1) ?? null;
    this.element.set(element);
  }

  private present(
    element: Element | null,
    next: { classes: string[]; styles: Readonly<Record<string, string>>; attributes: object }
  ): void {
    if (!element) return;
    for (const name of this.applied.classes) this.renderer.removeClass(element, name);
    for (const name of this.applied.styles) this.renderer.removeStyle(element, name);
    for (const name of this.applied.attributes) this.renderer.removeAttribute(element, name);
    const attributes = Object.entries(next.attributes as Record<string, string>).filter(([name]) =>
      allowedAttribute(name)
    );
    for (const name of next.classes) this.renderer.addClass(element, name);
    for (const [name, value] of Object.entries(next.styles)) {
      this.renderer.setStyle(element, name, value);
    }
    for (const [name, value] of attributes) this.renderer.setAttribute(element, name, value);
    this.applied = {
      classes: next.classes,
      styles: Object.keys(next.styles),
      attributes: attributes.map(([name]) => name),
    };
  }
}
