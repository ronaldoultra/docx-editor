import {
  Directive,
  afterNextRender,
  EmbeddedViewRef,
  Injectable,
  Injector,
  Renderer2,
  TemplateRef,
  ViewContainerRef,
  computed,
  effect,
  inject,
  input,
  signal,
  untracked,
  type Signal,
} from '@angular/core';

/**
 * What a part template gives: the presentation inputs and the template. The dialog parts
 * (`docxDialogPart`) and the hyperlink parts (`docxHyperlinkPart`) have this shape.
 *
 * @internal
 */
export interface DocxPartOverride {
  readonly className: Signal<string | undefined>;
  readonly styles: Signal<Readonly<Record<string, string>> | undefined>;
  readonly attributes: Signal<Readonly<Record<string, string>> | undefined>;
  readonly hidden: Signal<boolean>;
  readonly asChild: Signal<boolean>;
  readonly keepContent: Signal<boolean>;
  readonly template: TemplateRef<unknown>;
}

/**
 * The part templates of one compound, such as a packaged dialog. The compound provides it, and
 * connects its `contentChildren` query in its constructor.
 *
 * @internal
 */
@Injectable()
export class DocxPartRegistry {
  private overrides: Signal<readonly DocxPartOverride[]> = signal([]);
  private nameOf: (part: DocxPartOverride) => string = () => '';

  /** Connect the compound's part templates and the function that reads a template's name. */
  connectParts<T extends DocxPartOverride>(
    overrides: Signal<readonly T[]>,
    nameOf: (part: T) => string
  ): void {
    this.overrides = overrides;
    this.nameOf = nameOf as (part: DocxPartOverride) => string;
  }

  /** The last template for the part `name`. */
  override(name: string): DocxPartOverride | undefined {
    const all = this.overrides();
    for (let index = all.length - 1; index >= 0; index -= 1) {
      if (this.nameOf(all[index]!) === name) return all[index];
    }
    return undefined;
  }
}

/** The context of a part's default template: the host's content template, if any. @internal */
export interface DocxPartSlotContext {
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
 * Render a part of a compound: the packaged default with the host's presentation inputs, the
 * default with the host's content, or, with `asChild`, the host's template in its place. The
 * presentation inputs apply to the first element in each case, so `className` reaches a host's
 * own `asChild` element too.
 *
 * ```html
 * <button *docxPartSlot="'apply'; let content" docxDialogApply class="docx-dialog__button">
 *   @if (content) { <ng-container [ngTemplateOutlet]="content" /> } @else { Apply }
 * </button>
 * ```
 *
 * @internal
 */
@Directive({ selector: '[docxPartSlot]' })
export class DocxPartSlotDirective {
  readonly docxPartSlot = input.required<string>();

  private readonly parts = inject(DocxPartRegistry);
  private readonly template = inject<TemplateRef<DocxPartSlotContext>>(TemplateRef);
  private readonly container = inject(ViewContainerRef);
  private readonly renderer = inject(Renderer2);
  private readonly injector = inject(Injector);
  private view: EmbeddedViewRef<unknown> | null = null;
  private readonly element = signal<Element | null>(null);
  private rendered: readonly unknown[] = [];
  private applied: { classes: string[]; styles: string[]; attributes: string[] } = {
    classes: [],
    styles: [],
    attributes: [],
  };

  constructor() {
    const part = computed(() => this.parts.override(this.docxPartSlot()));
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
      DocxPartOverride | undefined,
      boolean,
      boolean,
      boolean,
    ];
    if (hidden) {
      this.view = null;
      this.element.set(null);
      return;
    }
    const view: EmbeddedViewRef<unknown> =
      override && asChild
        ? this.container.createEmbeddedView(override.template)
        : this.container.createEmbeddedView(this.template, {
            $implicit: override && !keepContent ? override.template : null,
          });
    this.view = view;
    const first = () => view.rootNodes.find((node): node is Element => node.nodeType === 1) ?? null;
    this.element.set(first());
    // An element inside a control flow block of the template exists only after the render.
    afterNextRender(
      () => {
        if (this.view === view && !view.destroyed) this.element.set(first());
      },
      { injector: this.injector }
    );
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
