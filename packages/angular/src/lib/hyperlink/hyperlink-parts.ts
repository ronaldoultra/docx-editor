import {
  Directive,
  TemplateRef,
  booleanAttribute,
  computed,
  inject,
  input,
  type Signal,
} from '@angular/core';
import { injectDocxEditor } from '../context';
import { injectTranslation } from '../i18n';
import type { HyperlinkPopup } from './hyperlink-state';

/**
 * The popup that the parts inside `<docx-editor-hyperlink>` act on. The component provides it.
 *
 * @internal
 */
export abstract class DocxHyperlinkHost {
  abstract readonly popup: HyperlinkPopup;
}

/** Keep the caret: a mousedown that reaches the pages moves it. Form fields keep theirs. */
export function guardMousedown(event: MouseEvent): void {
  const tag = (event.target as HTMLElement | null)?.tagName;
  if (tag === 'INPUT' || tag === 'SELECT' || tag === 'TEXTAREA') return;
  event.preventDefault();
}

/**
 * Replace or restyle one part of `<docx-editor-hyperlink>`. Put the template inside the popup.
 * The part names are `url`, `copy`, `edit`, `unlink`, `fields`, `error`, `apply`, and `cancel`.
 *
 * The template replaces the content of the part's element (for an action, its icon), and the
 * element keeps its wiring. With `asChild`, the template replaces the element. Put
 * `docxHyperlinkUrl`, `docxHyperlinkCopy`, `docxHyperlinkEdit`, `docxHyperlinkUnlink`,
 * `docxHyperlinkApply`, or `docxHyperlinkCancel` on your element to wire it. The packaged
 * `docx-hyperlink-popup__*` class stays off your element; `className` still reaches it.
 *
 * @example
 * ```html
 * <docx-editor-hyperlink>
 *   <ng-template docxHyperlinkPart="unlink" asChild>
 *     <button docxHyperlinkUnlink class="brand-button">Remove</button>
 *   </ng-template>
 * </docx-editor-hyperlink>
 * ```
 *
 * @public
 */
@Directive({ selector: 'ng-template[docxHyperlinkPart]' })
export class DocxHyperlinkPartDirective {
  /** The name of the part. */
  readonly docxHyperlinkPart = input.required<string>();
  /** Classes to add to the part's element. */
  readonly className = input<string | undefined>(undefined);
  /** Inline styles to add to the part's element. */
  readonly styles = input<Readonly<Record<string, string>> | undefined>(undefined);
  /** Attributes for the part's element, such as `id`. The part's own wiring wins a collision. */
  readonly attributes = input<Readonly<Record<string, string>> | undefined>(undefined);
  /** Remove the part. */
  readonly hidden = input(false, { transform: booleanAttribute });
  /** Render the template in the place of the part's element. */
  readonly asChild = input(false, { transform: booleanAttribute });
  /** Keep the packaged content and apply only the presentation inputs. */
  readonly keepContent = input(false, { transform: booleanAttribute });

  /** The template. */
  readonly template: TemplateRef<unknown> = inject(TemplateRef);
}

/** The label and focus helpers of one wiring directive. Call it in an injection context. */
function actionTools(): { label(key: string): string; returnFocus(): void } {
  const editor = injectDocxEditor();
  const translation = injectTranslation();
  return {
    label: (key) => {
      const t = translation.t();
      return t(key as Parameters<typeof t>[0]);
    },
    // Give the keyboard back to the document after the popup closed itself.
    returnFocus: () => editor()?.focus(),
  };
}

/**
 * Wire an element as the address of the link: a press follows the link. A link whose target the
 * editor refuses, such as a `javascript:` address, is inert: it has no action.
 *
 * @public
 */
@Directive({
  selector: '[docxHyperlinkUrl]',
  host: {
    'data-testid': 'hyperlink-popup-url',
    '[attr.data-inert]': 'inert() ? "" : null',
    '[attr.type]': 'inert() ? null : "button"',
    '[attr.title]': 'hint()',
    '(mousedown)': 'guard($event)',
    '(click)': 'press()',
  },
})
export class DocxHyperlinkUrlDirective {
  /** @internal */
  protected readonly popup: HyperlinkPopup = inject(DocxHyperlinkHost).popup;
  private readonly tools = actionTools();
  /** @internal */
  protected guard(event: MouseEvent): void {
    guardMousedown(event);
  }
  /** True when the link has no target that the editor opens. */
  readonly inert: Signal<boolean> = computed(() => !this.popup.state().link?.href);
  /** The hint for the address. */
  readonly hint: Signal<string> = computed(() => {
    const link = this.popup.state().link;
    if (!link?.href) return this.tools.label('hyperlinkPopup.inertTarget');
    return this.tools.label(
      link.kind === 'internal' ? 'hyperlinkPopup.bookmarkTarget' : 'hyperlinkPopup.openLink'
    );
  });
  /** @internal */
  protected press(): void {
    if (!this.inert()) this.popup.openTarget();
  }
}

/** Wire an element as the Copy link action. @public */
@Directive({
  selector: '[docxHyperlinkCopy]',
  host: {
    type: 'button',
    'data-testid': 'hyperlink-popup-copy',
    '[attr.data-copied]': 'popup.state().copied ? "" : null',
    '[attr.aria-label]': 'text()',
    '[attr.title]': 'text()',
    '(mousedown)': 'guard($event)',
    '(click)': 'press()',
  },
})
export class DocxHyperlinkCopyDirective {
  /** @internal */
  protected readonly popup: HyperlinkPopup = inject(DocxHyperlinkHost).popup;
  private readonly tools = actionTools();
  /** @internal */
  protected guard(event: MouseEvent): void {
    guardMousedown(event);
  }
  /** @internal */
  protected readonly text: Signal<string> = computed(() =>
    this.tools.label(this.popup.state().copied ? 'editor.linkCopied' : 'hyperlinkPopup.copyLink')
  );
  /** @internal */
  protected press(): void {
    void this.popup.copy();
  }
}

/** Wire an element as the Edit link action. @public */
@Directive({
  selector: '[docxHyperlinkEdit]',
  host: {
    type: 'button',
    'data-testid': 'hyperlink-popup-edit',
    '[attr.aria-label]': 'text()',
    '[attr.title]': 'text()',
    '(mousedown)': 'guard($event)',
    '(click)': 'popup.beginEdit()',
  },
})
export class DocxHyperlinkEditDirective {
  /** @internal */
  protected readonly popup: HyperlinkPopup = inject(DocxHyperlinkHost).popup;
  private readonly tools = actionTools();
  /** @internal */
  protected guard(event: MouseEvent): void {
    guardMousedown(event);
  }
  /** @internal */
  protected readonly text: Signal<string> = computed(() =>
    this.tools.label('hyperlinkPopup.editLink')
  );
}

/** Wire an element as the Remove link action. @public */
@Directive({
  selector: '[docxHyperlinkUnlink]',
  host: {
    type: 'button',
    'data-testid': 'hyperlink-popup-unlink',
    '[attr.aria-label]': 'text()',
    '[attr.title]': 'text()',
    '(mousedown)': 'guard($event)',
    '(click)': 'press()',
  },
})
export class DocxHyperlinkUnlinkDirective {
  /** @internal */
  protected readonly popup: HyperlinkPopup = inject(DocxHyperlinkHost).popup;
  private readonly tools = actionTools();
  /** @internal */
  protected guard(event: MouseEvent): void {
    guardMousedown(event);
  }
  /** @internal */
  protected readonly text: Signal<string> = computed(() =>
    this.tools.label('hyperlinkPopup.removeLink')
  );
  /** @internal */
  protected press(): void {
    if (this.popup.unlink()) this.tools.returnFocus();
  }
}

/** Wire an element as the Apply action. It is disabled while the address is empty. @public */
@Directive({
  selector: '[docxHyperlinkApply]',
  host: {
    type: 'button',
    'data-testid': 'hyperlink-popup-apply',
    '[attr.disabled]': 'popup.state().url.trim().length === 0 ? "" : null',
    '(mousedown)': 'guard($event)',
    '(click)': 'press()',
  },
})
export class DocxHyperlinkApplyDirective {
  /** @internal */
  protected readonly popup: HyperlinkPopup = inject(DocxHyperlinkHost).popup;
  private readonly tools = actionTools();
  /** @internal */
  protected guard(event: MouseEvent): void {
    guardMousedown(event);
  }
  /** @internal */
  protected press(): void {
    if (this.popup.state().url.trim().length === 0) return;
    if (this.popup.commitEdit()) this.tools.returnFocus();
  }
}

/** Wire an element as the Cancel action. @public */
@Directive({
  selector: '[docxHyperlinkCancel]',
  host: {
    type: 'button',
    'data-testid': 'hyperlink-popup-cancel',
    '(mousedown)': 'guard($event)',
    '(click)': 'press()',
  },
})
export class DocxHyperlinkCancelDirective {
  /** @internal */
  protected readonly popup: HyperlinkPopup = inject(DocxHyperlinkHost).popup;
  private readonly tools = actionTools();
  /** @internal */
  protected guard(event: MouseEvent): void {
    guardMousedown(event);
  }
  /** @internal */
  protected press(): void {
    this.popup.close();
    this.tools.returnFocus();
  }
}
