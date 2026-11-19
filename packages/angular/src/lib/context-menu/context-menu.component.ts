import { NgTemplateOutlet } from '@angular/common';
import {
  ChangeDetectionStrategy,
  Component,
  ElementRef,
  Injector,
  afterNextRender,
  afterRenderEffect,
  booleanAttribute,
  computed,
  contentChildren,
  effect,
  inject,
  input,
  output,
  signal,
  untracked,
  viewChild,
  type Signal,
  type TemplateRef,
} from '@angular/core';
import {
  hostDocumentOf,
  hostWindowOf,
  isDomNode,
  isHTMLElementNode,
  placeContextMenu,
  revealMenuItem,
  showInTopLayer,
} from '@docx-editor.dev/core/editor';
import { injectScopeClassName } from '../color-mode';
import { injectDocxEditor, injectDocxEditorContext } from '../context';
import { injectEditorState } from '../editor-state';
import { injectChromeDir, type DocxEditorDirection } from '../i18n';
import { DocxMenuItemComponent } from '../menu/menu-item.component';
import { focusBy, focusEdge, panelItems } from '../menu/menu-keyboard';
import { DocxMenuSeparatorComponent } from '../menu/menu-row.component';
import {
  DOCX_MENU_ROW_TABINDEX,
  DocxMenuBarState,
  injectMenuLabel,
  type DocxMenuTranslate,
} from '../menu/menu-state';
import { DocxSlotDirective } from '../slot.directive';
import { toolbarDevWarning } from '../toolbar/toolbar-registry';
import {
  DocxContextMenuRowDirective,
  DocxContextMenuStartComponent,
} from './context-menu-item.component';
import {
  DocxContextMenuCellAlignmentComponent,
  DocxContextMenuRowComponent,
  isPackagedContextMenuRow,
  type ContextMenuPackagedRow,
} from './context-menu-row.component';
import { DocxContextMenuBarState, DocxContextMenuState } from './context-menu-state';
import {
  contextMenuAnchorOf,
  type ContextMenuAnchor,
  isDocumentContextMenuTarget,
  isKeyboardContextMenu,
} from './context-menu-target';

/** One entry of the packaged rows: a row or a separator. @internal */
export interface ContextMenuEntry {
  readonly kind: 'row' | 'separator';
  readonly id: string;
}

const row = (id: string): ContextMenuEntry => ({ kind: 'row', id });
const separator = (id: string): ContextMenuEntry => ({ kind: 'separator', id });

const BASE_ROWS: readonly ContextMenuEntry[] = [
  row('field.edit'),
  row('image.editChart'),
  row('image.formatShape'),
  row('edit.cut'),
  row('edit.copy'),
  row('edit.paste'),
  row('edit.pasteWithoutFormatting'),
  // The two halves of the format painter, in the clipboard group.
  row('format.copyFormatting'),
  row('format.pasteFormatting'),
  separator('sep.clipboard'),
  row('edit.delete'),
  row('edit.selectAll'),
  separator('sep.selection'),
  row('text.link'),
];

const TABLE_ROWS: readonly ContextMenuEntry[] = [
  separator('sep.table'),
  row('table.selectTable'),
  row('table.insertRowAbove'),
  row('table.insertRowBelow'),
  separator('sep.table.columns'),
  row('table.insertColumnLeft'),
  row('table.insertColumnRight'),
  separator('sep.table.cells'),
  row('table.mergeCells'),
  row('table.splitCell'),
  separator('sep.table.destructive'),
  row('table.deleteRow'),
  row('table.deleteColumn'),
  row('table.deleteTable'),
  separator('sep.table.alignment'),
  row('table.cellVerticalAlignment'),
];

const TOC_ROWS: readonly ContextMenuEntry[] = [
  separator('sep.toc'),
  row('toc.refresh'),
  row('toc.refreshPageNumbers'),
  row('toc.remove'),
];

/** The packaged rows for a menu in a table or on a table of contents. @internal */
export function contextMenuDefaultSet(
  inTable: boolean,
  onToc: boolean
): readonly ContextMenuEntry[] {
  return [...BASE_ROWS, ...(inTable ? TABLE_ROWS : []), ...(onToc ? TOC_ROWS : [])];
}

/** One override template: a row id, its template, and whether it removes the row. @internal */
export interface ContextMenuRowOverride {
  readonly id: string;
  readonly hidden: boolean;
  readonly template: TemplateRef<unknown>;
}

/** Where the measured panel sits, and the size past which it scrolls. */
type Placement = ReturnType<typeof placeContextMenu>;

const SCROLLER = '.docx-editor__scroll-container';

/**
 * The document context menu. A right-click on the pages, Shift+F10, or the Menu key opens it.
 * Put it in the viewport, after the content.
 *
 * The panel holds the packaged rows: clipboard, format painter, delete, select all, and link.
 * In a table, it adds the table rows, and on a table of contents, the table of contents rows.
 * On a form field, a chart, or a shape, it adds the row that edits the object. Each row asks
 * the engine whether it is enabled.
 *
 * Change the rows with content:
 *
 * - `<ng-template docxContextMenuRow="edit.cut">` replaces one row, and
 *   `docxContextMenuRowHidden` removes it. `ng-template[docxSlot]` does the same for a chrome
 *   slot row, such as `text.link`.
 * - `docx-context-menu-start` holds rows that go above the packaged rows.
 * - Other content, such as `docx-context-menu-item`, comes after the packaged rows.
 *
 * With `[preset]="false"`, the panel holds only your content.
 *
 * @example
 * ```html
 * <docx-editor-viewport>
 *   <docx-editor-content />
 *   <div docxEditorContextMenu (openChange)="menuOpen = $event">
 *     <docx-context-menu-item label="Look up" (activate)="lookUp()" />
 *   </div>
 * </docx-editor-viewport>
 * ```
 *
 * @public
 */
@Component({
  selector: 'div[docxEditorContextMenu]',
  changeDetection: ChangeDetectionStrategy.OnPush,
  imports: [
    NgTemplateOutlet,
    DocxMenuItemComponent,
    DocxMenuSeparatorComponent,
    DocxContextMenuRowComponent,
    DocxContextMenuCellAlignmentComponent,
  ],
  providers: [
    DocxContextMenuState,
    { provide: DocxMenuBarState, useClass: DocxContextMenuBarState },
    { provide: DOCX_MENU_ROW_TABINDEX, useValue: -1 },
  ],
  host: {
    class: 'docx-contextmenu-host',
    style: 'position: absolute; width: 0; height: 0; overflow: visible; pointer-events: none',
  },
  template: `
    @if (anchor(); as opened) {
      <div
        #panel
        role="menu"
        tabindex="-1"
        [class]="panelClass()"
        [attr.aria-label]="label()('contextMenu.ariaLabel')"
        [attr.dir]="chromeDir() ?? null"
        [style.position]="'fixed'"
        [style.left.px]="placement()?.left ?? opened.x"
        [style.top.px]="placement()?.top ?? opened.y"
        [style.max-width.px]="placement()?.maxWidth ?? null"
        [style.max-height.px]="placement()?.maxHeight ?? null"
        [style.visibility]="placement() ? 'visible' : 'hidden'"
        [style.pointer-events]="'auto'"
        (keydown)="onPanelKey($event)"
      >
        <ng-content select="docx-context-menu-start" />
        @if (preset() && startRows()) {
          <div docxMenuSeparator></div>
        }
        @if (preset()) {
          @for (entry of entries(); track entry.id) {
            @if (entry.kind === 'separator') {
              <div docxMenuSeparator></div>
            } @else if (overrides().get(entry.id); as override) {
              @if (!override.hidden) {
                <ng-container [ngTemplateOutlet]="override.template" />
              }
            } @else if (entry.id === 'text.link') {
              <docx-menu-item slotId="text.link" shortcutKey="contextMenu.insertLinkShortcut" />
            } @else if (entry.id === 'table.cellVerticalAlignment') {
              <docx-context-menu-cell-alignment />
            } @else if (isRow(entry.id)) {
              <docx-context-menu-row [row]="entry.id" />
            }
          }
          @for (override of appended(); track override.id) {
            <ng-container [ngTemplateOutlet]="override.template" />
          }
        }
        <ng-content />
      </div>
    }
  `,
})
export class DocxEditorContextMenuComponent {
  /** Render the packaged rows. With `false`, the panel holds only your content. */
  readonly preset = input(true, { transform: booleanAttribute });
  /** Do not open, and close an open menu. A right-click then shows the browser's menu. */
  readonly disabled = input(false, { transform: booleanAttribute });
  /** Your own labels. A key that this function does not answer uses the active catalog. */
  readonly t = input<DocxMenuTranslate | undefined>(undefined);
  /** More classes for the panel. */
  readonly className = input<string | undefined>(undefined);
  /** Opens the split cells dialog. Without it, the menu has no Split cells row. */
  readonly splitCellHandler = input<(() => void) | undefined>(undefined);
  /** Emits `true` when the menu opens and `false` when it closes. */
  readonly openChange = output<boolean>();

  private readonly state = inject(DocxContextMenuState);
  private readonly bar = inject(DocxMenuBarState);
  private readonly context = injectDocxEditorContext();
  private readonly editor = injectDocxEditor();
  private readonly host: HTMLElement = inject(ElementRef).nativeElement;
  private readonly injector = inject(Injector);
  private readonly panel = viewChild<ElementRef<HTMLElement>>('panel');
  private readonly rowTemplates = contentChildren(DocxContextMenuRowDirective);
  private readonly slotTemplates = contentChildren(DocxSlotDirective);
  private readonly starts = contentChildren(DocxContextMenuStartComponent);
  private readonly inTable = injectEditorState((snapshot) => snapshot.table != null);
  /** The scroller that the open menu came from: its document and window own the menu. */
  private menuScroller: HTMLElement | null = null;

  /** @internal */
  protected readonly anchor: Signal<ContextMenuAnchor | null> = this.state.anchor.asReadonly();
  /** @internal */
  protected readonly placement = signal<Placement | null>(null);
  /** @internal */
  protected readonly label = injectMenuLabel();
  /** @internal */
  protected readonly chromeDir: Signal<DocxEditorDirection | undefined> = injectChromeDir();
  /** @internal */
  protected readonly startRows = signal(false);
  /** @internal */
  protected isRow(id: string): id is ContextMenuPackagedRow {
    return isPackagedContextMenuRow(id);
  }

  private readonly scope = injectScopeClassName();

  /** @internal */
  protected readonly panelClass = computed(() => {
    const extra = this.className();
    return `${this.scope()}docx-toolbar__menu docx-contextmenu${extra ? ` ${extra}` : ''}`;
  });

  /** @internal */
  protected readonly entries = computed(() =>
    contextMenuDefaultSet(this.inTable(), this.state.tocId() !== null)
  );

  /** @internal */
  protected readonly overrides = computed(() => {
    const map = new Map<string, ContextMenuRowOverride>();
    for (const slot of this.slotTemplates()) {
      map.set(slot.docxSlot(), {
        id: slot.docxSlot(),
        hidden: slot.docxSlotHidden(),
        template: slot.template,
      });
    }
    for (const own of this.rowTemplates()) {
      map.set(own.docxContextMenuRow(), {
        id: own.docxContextMenuRow(),
        hidden: own.docxContextMenuRowHidden(),
        template: own.template,
      });
    }
    return map;
  });

  /** The templates that name no packaged row: they come after the rows. @internal */
  protected readonly appended = computed(() => {
    const placed = new Set(contextMenuDefaultSet(true, true).map((entry) => entry.id));
    return [...this.overrides().values()].filter(
      (override) => !override.hidden && !placed.has(override.id)
    );
  });

  /**
   * The scroller to listen on. It is resolved again after each engine change, so a scroller
   * that mounts later is found. A change that keeps the same scroller binds nothing again.
   */
  private readonly scroller = computed(() => {
    this.context?.revision();
    if (this.disabled()) return null;
    return (
      this.host.closest<HTMLElement>(SCROLLER) ??
      hostDocumentOf(this.host).querySelector<HTMLElement>(SCROLLER)
    );
  });

  constructor() {
    const state = this.state;
    state.useClose((restoreFocus) => this.close(restoreFocus));
    this.bar.useTranslate(() => this.t());
    effect(() => state.preset.set(this.preset()));
    effect(() => state.splitCell.set(this.splitCellHandler()));
    effect(() => {
      if (this.disabled()) untracked(() => this.close(false));
    });
    effect(() => {
      const placed = new Set(contextMenuDefaultSet(true, true).map((entry) => entry.id));
      for (const override of this.overrides().values()) {
        if (override.hidden && !placed.has(override.id)) {
          toolbarDevWarning(
            `The context menu has no row "${override.id}" to hide. Check the row id.`
          );
        }
      }
    });
    effect((onCleanup) => {
      const scroller = this.scroller();
      if (!scroller) return;
      const onContextMenu = (event: MouseEvent) => this.onContextMenu(scroller, event);
      scroller.addEventListener('contextmenu', onContextMenu);
      onCleanup(() => scroller.removeEventListener('contextmenu', onContextMenu));
    });
    effect((onCleanup) => {
      if (!state.anchor()) return;
      onCleanup(untracked(() => this.listenWhileOpen()));
    });
    let wasOpen = false;
    effect(() => {
      const open = state.anchor() !== null;
      if (open === wasOpen) return;
      wasOpen = open;
      untracked(() => this.openChange.emit(open));
    });
    afterRenderEffect(() => {
      this.startRows.set(this.starts().some((start) => start.hasRows()));
    });
    // Measure the panel once it renders, then place it inside the window.
    afterRenderEffect(() => {
      const panel = this.panel()?.nativeElement;
      const anchor = this.anchor();
      if (!panel || !anchor || untracked(this.placement)) return;
      // The top layer first: a host stacking context otherwise paints over the panel, whatever
      // its z-index. The panel must also be open there to be measured.
      showInTopLayer(panel);
      const { width, height } = panel.getBoundingClientRect();
      const view = hostWindowOf(panel);
      // Right-to-left chrome opens toward the left of the pointer.
      const rtl = view.getComputedStyle(panel).direction === 'rtl';
      this.placement.set(
        placeContextMenu(
          anchor,
          { width, height },
          { width: view.innerWidth, height: view.innerHeight },
          rtl
        )
      );
    });
    // After the render that makes the panel visible: a hidden panel cannot take the focus, and
    // the arrow keys work on the panel.
    effect(() => {
      if (!this.placement()) return;
      afterNextRender(() => untracked(this.panel)?.nativeElement.focus({ preventScroll: true }), {
        injector: this.injector,
      });
    });
  }

  private onContextMenu(scroller: HTMLElement, event: MouseEvent): void {
    const panel = untracked(this.panel)?.nativeElement;
    // A second right-click on the open panel does not move it, and shows no browser menu.
    if (panel && isDomNode(event.target) && panel.contains(event.target)) {
      event.preventDefault();
      return;
    }
    // Another menu answered this press already. One press opens one menu.
    if (event.defaultPrevented) return;
    // Chrome in the scroller, such as the link popup, keeps the browser's own menu.
    if (!isDocumentContextMenuTarget(scroller, event.target)) return;
    event.preventDefault();
    const keyboard = isKeyboardContextMenu(event);
    const editor = untracked(this.editor);
    this.menuScroller = scroller;
    this.placement.set(null);
    this.state.tocId.set(editor?.snapshot().tocContext?.id ?? null);
    this.state.target.set(keyboard || !isHTMLElementNode(event.target) ? null : event.target);
    this.state.anchor.set(contextMenuAnchorOf(event, keyboard, editor, scroller));
  }

  /** The listeners that close the open menu. Returns the function that removes them. */
  private listenWhileOpen(): () => void {
    const inPanel = (target: EventTarget | null) => {
      const panel = untracked(this.panel)?.nativeElement;
      return !!panel && isDomNode(target) && panel.contains(target);
    };
    const onPointerDown = (event: Event) => {
      if (!inPanel(event.target)) this.close(false);
    };
    const onKeyDown = (event: KeyboardEvent) => {
      if (event.key === 'Escape') {
        event.preventDefault();
        this.close(true);
      } else if (event.key === 'Tab') {
        this.close(false);
      }
    };
    const onScroll = (event: Event) => {
      // The panel's own scroll, when a long menu scrolls inside, is not the user moving on.
      if (event.type === 'scroll' && inPanel(event.target)) return;
      this.close(false);
    };
    const origin = this.menuScroller ?? this.host;
    const doc = hostDocumentOf(origin);
    const view = hostWindowOf(origin);
    doc.addEventListener('pointerdown', onPointerDown, true);
    doc.addEventListener('keydown', onKeyDown);
    doc.addEventListener('scroll', onScroll, { capture: true, passive: true });
    view.addEventListener('blur', onScroll);
    view.addEventListener('resize', onScroll);
    return () => {
      doc.removeEventListener('pointerdown', onPointerDown, true);
      doc.removeEventListener('keydown', onKeyDown);
      doc.removeEventListener('scroll', onScroll, true);
      view.removeEventListener('blur', onScroll);
      view.removeEventListener('resize', onScroll);
    };
  }

  private close(restoreFocus: boolean): void {
    if (untracked(this.state.anchor) === null) return;
    this.state.anchor.set(null);
    this.state.tocId.set(null);
    this.state.target.set(null);
    this.placement.set(null);
    if (restoreFocus) untracked(this.editor)?.focus();
  }

  /** @internal */
  protected onPanelKey(event: KeyboardEvent): void {
    const panel = untracked(this.panel)?.nativeElement;
    if (!panel) return;
    const items = panelItems(panel);
    const doc = hostDocumentOf(panel);
    if (event.key === 'ArrowDown') focusBy(items, doc.activeElement, 1);
    else if (event.key === 'ArrowUp') focusBy(items, doc.activeElement, -1);
    else if (event.key === 'Home') focusEdge(items, 'first');
    else if (event.key === 'End') focusEdge(items, 'last');
    else return;
    event.preventDefault();
    // A menu taller than the window scrolls: keep the focused row in view.
    revealMenuItem(panel, doc.activeElement);
  }
}
