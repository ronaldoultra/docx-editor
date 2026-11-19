import { NgTemplateOutlet } from '@angular/common';
import {
  ChangeDetectionStrategy,
  Component,
  ElementRef,
  Injector,
  afterNextRender,
  booleanAttribute,
  computed,
  contentChildren,
  effect,
  inject,
  input,
  signal,
  untracked,
  viewChild,
  type Signal,
} from '@angular/core';
import { hostDocumentOf, listenForOutsidePress } from '@docx-editor.dev/core/editor';
import { injectDocxEditor } from '../context';
import { injectTranslation } from '../i18n';
import { DocxPartRegistry, DocxPartSlotDirective } from '../parts/part-slot';
import {
  DocxHyperlinkApplyDirective,
  DocxHyperlinkCancelDirective,
  DocxHyperlinkCopyDirective,
  DocxHyperlinkEditDirective,
  DocxHyperlinkHost,
  DocxHyperlinkPartDirective,
  DocxHyperlinkUnlinkDirective,
  DocxHyperlinkUrlDirective,
  guardMousedown,
} from './hyperlink-parts';
import {
  injectHyperlinkPopup,
  isFieldLink,
  type HyperlinkPopup,
  type HyperlinkPopupState,
} from './hyperlink-state';

const ICONS = {
  copy: 'M360-240q-33 0-56.5-23.5T280-320v-480q0-33 23.5-56.5T360-880h360q33 0 56.5 23.5T800-800v480q0 33-23.5 56.5T720-240H360Zm0-80h360v-480H360v480ZM200-80q-33 0-56.5-23.5T120-160v-560h80v560h440v80H200Zm160-240v-480 480Z',
  edit: 'M200-200h57l391-391-57-57-391 391v57Zm-80 80v-170l528-527q12-11 26.5-17t30.5-6q16 0 31 6t26 18l55 56q12 11 17.5 26t5.5 30q0 16-5.5 30.5T817-647L290-120H120Zm640-584-56-56 56 56Zm-141 85-28-29 57 57-29-28Z',
  unlink:
    'M770-302 656-416l57-57 114 114-57 57ZM603-469 469-603l57-57 134 134-57 57ZM280-280h133v80H280q-83 0-141.5-58.5T80-400q0-83 58.5-141.5T280-600h133v80H280q-50 0-85 35t-35 85q0 50 35 85t85 35Zm40-80v-80h101l80 80H320Zm493-15-57-57q23-11 33.5-32t10.5-36q0-50-35-85t-85-35H547v-80h133q83 0 141.5 58.5T880-400q0 32-11 61.5T813-375ZM792-56 56-792l56-56 736 736-56 56Z',
  open: 'M200-120q-33 0-56.5-23.5T120-200v-560q0-33 23.5-56.5T200-840h280v80H200v560h560v-280h80v280q0 33-23.5 56.5T760-120H200Zm188-212-56-56 372-372H560v-80h280v280h-80v-144L388-332Z',
} as const;

let nextId = 0;

/**
 * The hyperlink popup: it shows a link that the reader presses, with Copy, Edit, and Remove, and
 * it edits or inserts a link with Ctrl+K, the `text.link` control, and the context menu row.
 *
 * Put it inside `<docx-editor-viewport>`, or name it as the `hyperlink` entry of the popup
 * registry. While it exists, the editor's link controls are enabled. Replace or restyle a part
 * with `docxHyperlinkPart` templates; other content renders after the parts.
 *
 * @example
 * ```html
 * <docx-editor-viewport>
 *   <docx-editor-content />
 *   <docx-editor-hyperlink />
 * </docx-editor-viewport>
 * ```
 *
 * @public
 */
@Component({
  selector: 'docx-editor-hyperlink',
  imports: [
    NgTemplateOutlet,
    DocxPartSlotDirective,
    DocxHyperlinkUrlDirective,
    DocxHyperlinkCopyDirective,
    DocxHyperlinkEditDirective,
    DocxHyperlinkUnlinkDirective,
    DocxHyperlinkApplyDirective,
    DocxHyperlinkCancelDirective,
  ],
  providers: [
    DocxPartRegistry,
    { provide: DocxHyperlinkHost, useExisting: DocxEditorHyperlinkComponent },
  ],
  changeDetection: ChangeDetectionStrategy.OnPush,
  host: { style: 'display: contents' },
  templateUrl: './hyperlink.component.html',
})
export class DocxEditorHyperlinkComponent {
  /** Classes for the panel. */
  readonly className = input<string | undefined>(undefined);
  /**
   * With `true` (the default), the panel renders its packaged parts, the part templates in their
   * place, and the other content after them. With `false`, it renders only its content.
   */
  readonly preset = input(true, { transform: booleanAttribute });
  /** Render no panel: only the content, while the popup is open. */
  readonly hidden = input(false, { transform: booleanAttribute });

  /** The popup state. Inside a root, it is the state that the root shares. */
  readonly popup: HyperlinkPopup = injectHyperlinkPopup();

  /** @internal */
  protected readonly state: Signal<HyperlinkPopupState> = this.popup.state;
  /** @internal */
  protected readonly icons: Readonly<Record<'copy' | 'edit' | 'unlink' | 'open', string>> = ICONS;
  /** @internal */
  protected readonly textId: string = `docx-hyperlink-${++nextId}-text`;
  /** @internal */
  protected readonly urlId: string = `docx-hyperlink-${nextId}-url`;
  /** @internal */
  protected readonly placement = signal<{ left: string; top: string } | null>(null);
  /** @internal */
  protected readonly title: Signal<string> = computed(() => {
    const state = this.state();
    if (state.mode !== 'editing') return this.label('hyperlinkPopup.editLink');
    return this.label(state.link ? 'hyperlinkPopup.editTitle' : 'hyperlinkPopup.insertTitle');
  });
  /** @internal */
  protected readonly editable: Signal<boolean> = computed(() => {
    const state = this.state();
    return state.canEdit && !(state.link && isFieldLink(state.link));
  });

  private readonly translation = injectTranslation();
  private readonly editor = injectDocxEditor();
  private readonly injector = inject(Injector);
  private readonly host = inject(ElementRef<HTMLElement>);
  private readonly panel = viewChild<ElementRef<HTMLElement>>('panel');
  private readonly urlInput = viewChild<ElementRef<HTMLInputElement>>('url');
  private readonly parts = contentChildren(DocxHyperlinkPartDirective);

  constructor() {
    inject(DocxPartRegistry, { self: true }).connectParts(this.parts, (part) =>
      part.docxHyperlinkPart()
    );
    const mode = computed(() => this.state().mode);
    // Escape and a press outside close the open popup.
    effect((onCleanup) => {
      if (mode() === 'closed') return;
      const doc = hostDocumentOf(this.host.nativeElement);
      const onKeydown = (event: KeyboardEvent): void => {
        if (event.key !== 'Escape') return;
        event.preventDefault();
        this.popup.close();
        this.returnFocus();
      };
      doc.addEventListener('keydown', onKeydown);
      const stopPress = listenForOutsidePress({
        popup: doc,
        contains: (node) => this.panel()?.nativeElement.contains(node) === true,
        close: () => this.popup.close(),
      });
      onCleanup(() => {
        doc.removeEventListener('keydown', onKeydown);
        stopPress();
      });
    });
    // The address field takes the focus when an edit starts.
    effect(() => {
      if (mode() !== 'editing') return;
      afterNextRender(
        () => {
          const input = this.urlInput()?.nativeElement;
          input?.focus({ preventScroll: true });
          input?.select();
        },
        { injector: this.injector }
      );
    });
    effect(() => {
      const anchor = this.state().anchor;
      untracked(() => afterNextRender(() => this.place(anchor), { injector: this.injector }));
    });
  }

  /** @internal */
  protected label(key: string): string {
    const t = this.translation.t();
    return t(key as Parameters<typeof t>[0]);
  }

  /** @internal */
  protected guard(event: MouseEvent): void {
    guardMousedown(event);
  }

  /** @internal */
  protected onFieldKeydown(event: KeyboardEvent): void {
    if (event.key === 'Enter') {
      event.preventDefault();
      // A refused draft keeps the popup open, and the focus with it, so the value can be fixed.
      if (this.popup.commitEdit()) this.returnFocus();
      return;
    }
    if (event.key === 'Escape') {
      event.preventDefault();
      event.stopPropagation();
      this.popup.close();
      this.returnFocus();
    }
  }

  /** @internal */
  protected inputValue(event: Event): string {
    return (event.target as HTMLInputElement).value;
  }

  private returnFocus(): void {
    this.editor()?.focus();
  }

  /** Put the panel at the anchor, inside its positioned container. */
  private place(anchor: HyperlinkPopupState['anchor']): void {
    const panel = this.panel()?.nativeElement;
    const container = panel?.offsetParent as HTMLElement | null | undefined;
    if (!anchor || !panel || !container) {
      this.placement.set(null);
      return;
    }
    const box = container.getBoundingClientRect();
    const left = anchor.left - box.left - container.clientLeft + container.scrollLeft;
    const top = anchor.top - box.top - container.clientTop + container.scrollTop;
    const maxLeft = Math.max(0, container.scrollWidth - panel.offsetWidth);
    const px = (value: number) => `${Math.round(value * 100) / 100}px`;
    this.placement.set({ left: px(Math.max(0, Math.min(left, maxLeft))), top: px(top) });
  }
}
