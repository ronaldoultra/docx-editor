import {
  ChangeDetectionStrategy,
  Component,
  ElementRef,
  afterRenderEffect,
  computed,
  inject,
  input,
  signal,
  untracked,
  viewChild,
  type OnInit,
} from '@angular/core';
import {
  hostDocumentOf,
  hostWindowOf,
  isDomNode,
  placeSubmenu,
  submenuKeys,
} from '@docx-editor.dev/core/editor';
import { DocxChromeIconComponent } from '../toolbar/chrome-icon.component';
import { popupId } from '../toolbar/toolbar-popup';
import { guardToolbarMousedown, toolbarDevWarning } from '../toolbar/toolbar-registry';
import { focusBy, focusEdge, panelItems } from './menu-keyboard';
import { DOCX_MENU_ROW_TABINDEX, injectMenuLabel } from './menu-state';

/**
 * A menu row that opens a nested panel at its inline end, such as Insert › Break. The row runs
 * nothing, so it stays enabled, and each row inside answers for itself. The panel opens on
 * hover, on click, and on focus. The arrow key toward the panel moves the focus into it, and
 * the opposite arrow key or Escape gives the focus back to the row.
 *
 * @example
 * ```html
 * <docx-menu-submenu labelKey="toolbar.break">
 *   <docx-menu-item slotId="insert.pageBreak" />
 * </docx-menu-submenu>
 * ```
 *
 * @public
 */
@Component({
  selector: 'docx-menu-submenu',
  changeDetection: ChangeDetectionStrategy.OnPush,
  imports: [DocxChromeIconComponent],
  providers: [{ provide: DOCX_MENU_ROW_TABINDEX, useValue: -1 }],
  host: {
    role: 'none',
    '[class]': 'hostClass()',
    '(mouseenter)': 'open.set(true)',
    '(mouseleave)': 'open.set(false)',
    '(keydown)': 'onKey($event)',
    '(focusout)': 'onFocusOut($event)',
  },
  template: `<button
      #parent
      type="button"
      role="menuitem"
      aria-haspopup="menu"
      class="docx-toolbar__menu-item docx-menubar__item"
      tabindex="-1"
      [attr.aria-expanded]="open()"
      [attr.aria-controls]="open() ? panelId : null"
      [attr.data-open]="open() ? '' : null"
      (mousedown)="guard($event)"
      (focus)="open.set(true)"
      (click)="open.set(true)"
    >
      <span class="docx-menubar__item-icon" aria-hidden="true">
        @if (paths(); as icon) {
          <svg [docxChromeIcon]="icon"></svg>
        }
      </span>
      <span class="docx-menubar__item-label">{{ text() }}</span>
      <span class="docx-menubar__item-caret" aria-hidden="true">›</span>
    </button>
    @if (open()) {
      <div
        #panel
        role="menu"
        class="docx-toolbar__menu docx-menubar__menu docx-menubar__submenu-panel"
        style="position: fixed; visibility: hidden"
        [id]="panelId"
        [attr.aria-label]="text()"
        (keydown)="onPanelKey($event)"
      >
        <ng-content />
      </div>
    }`,
})
export class DocxMenuSubmenuComponent implements OnInit {
  /** The row text. It wins over `labelKey`. */
  readonly label = input<string | undefined>(undefined);
  /** The i18n key of the row text. */
  readonly labelKey = input<string | undefined>(undefined);
  /** Material Symbols paths for the row's icon. */
  readonly paths = input<readonly string[] | null | undefined>(undefined);
  /** More classes for the submenu root. */
  readonly className = input<string | undefined>(undefined);

  private readonly menuLabel = injectMenuLabel();
  private readonly host: HTMLElement = inject(ElementRef).nativeElement;
  private readonly parent = viewChild.required<ElementRef<HTMLElement>>('parent');
  private readonly panel = viewChild<ElementRef<HTMLElement>>('panel');
  private focusFirst = false;

  /** @internal */
  protected readonly open = signal(false);
  /** @internal */
  protected readonly panelId = popupId('menu-flyout');
  /** @internal */
  protected readonly guard: (event: Event) => void = guardToolbarMousedown;
  /** @internal */
  protected readonly hostClass = computed(() => {
    const extra = this.className();
    return `docx-menubar__submenu${extra ? ` ${extra}` : ''}`;
  });
  /** @internal */
  protected readonly text = computed(() => this.label() ?? this.menuLabel()(this.labelKey() ?? ''));

  /** @internal */
  ngOnInit(): void {
    if (this.label() === undefined && this.labelKey() === undefined) {
      toolbarDevWarning(
        'A docx-menu-submenu has neither label nor labelKey, so its row has no name.'
      );
    }
  }

  constructor() {
    afterRenderEffect((onCleanup) => {
      const panel = this.panel()?.nativeElement;
      if (!panel) return;
      const row = this.parent().nativeElement;
      const view = hostWindowOf(row);
      const place = () => {
        // The row's inline-end side first, and the other side when that side has no room.
        const box = placeSubmenu(
          row.getBoundingClientRect(),
          { width: panel.offsetWidth, height: panel.offsetHeight },
          { width: view.innerWidth, height: view.innerHeight },
          view.getComputedStyle(row).direction === 'rtl'
        );
        panel.style.left = `${box.left}px`;
        panel.style.top = `${box.top}px`;
        panel.style.visibility = '';
      };
      place();
      if (this.focusFirst) {
        this.focusFirst = false;
        focusEdge(panelItems(panel), 'first');
      }
      // A long menu scrolls inside itself, and the panel follows its row.
      const parentPanel = row.closest('[role="menu"]');
      parentPanel?.addEventListener('scroll', place, { passive: true });
      onCleanup(() => parentPanel?.removeEventListener('scroll', place));
    });
  }

  /** @internal */
  protected onKey(event: KeyboardEvent): void {
    const keys = submenuKeys(
      hostWindowOf(this.host).getComputedStyle(this.host).direction === 'rtl'
    );
    const row = this.parent().nativeElement;
    if (event.key === keys.open && hostDocumentOf(row).activeElement === row) {
      event.preventDefault();
      const panel = untracked(this.panel)?.nativeElement;
      if (untracked(this.open) && panel) focusEdge(panelItems(panel), 'first');
      else this.focusFirst = true;
      this.open.set(true);
    } else if ((event.key === keys.close || event.key === 'Escape') && untracked(this.open)) {
      event.preventDefault();
      event.stopPropagation();
      this.open.set(false);
      row.focus();
    }
  }

  /** @internal */
  protected onFocusOut(event: FocusEvent): void {
    const next = event.relatedTarget;
    if (!isDomNode(next) || !this.host.contains(next)) this.open.set(false);
  }

  /** @internal */
  protected onPanelKey(event: KeyboardEvent): void {
    const panel = untracked(this.panel)?.nativeElement;
    if (!panel) return;
    const step = event.key === 'ArrowDown' ? 1 : event.key === 'ArrowUp' ? -1 : 0;
    if (step === 0) return;
    event.preventDefault();
    event.stopPropagation();
    focusBy(panelItems(panel), hostDocumentOf(panel).activeElement, step);
  }
}
