import { NgTemplateOutlet } from '@angular/common';
import {
  ChangeDetectionStrategy,
  Component,
  DestroyRef,
  ElementRef,
  afterRenderEffect,
  booleanAttribute,
  computed,
  contentChildren,
  inject,
  input,
  untracked,
  viewChild,
  type OnInit,
  type Signal,
  type TemplateRef,
} from '@angular/core';
import {
  CHROME_MENUS,
  hostDocumentOf,
  toolbarPanelPlacement,
  type ChromeMenuEntry,
  type ChromeSlotId,
} from '@docx-editor.dev/core/editor';
import { DocxSlotDirective } from '../slot.directive';
import { DocxChromeIconComponent } from '../toolbar/chrome-icon.component';
import { popupId } from '../toolbar/toolbar-popup';
import { guardToolbarMousedown } from '../toolbar/toolbar-registry';
import { DocxMenuItemComponent } from './menu-item.component';
import { focusBy, focusEdge, panelItems } from './menu-keyboard';
import {
  DocxMenuChartPickerComponent,
  DocxMenuReportIssueComponent,
  DocxMenuReviewersComponent,
  DocxMenuTablePickerComponent,
} from './menu-pickers.component';
import {
  DocxMenuGroupComponent,
  DocxMenuRowComponent,
  DocxMenuSeparatorComponent,
} from './menu-row.component';
import {
  DOCX_MENU_ROW_TABINDEX,
  DocxMenuBarState,
  MENU_OVERFLOW_ID,
  injectMenuLabel,
  provideMenuBarState,
} from './menu-state';
import { DocxMenuSubmenuComponent } from './menu-submenu.component';

/** Every slot that a list of registry entries places, submenus included. */
function entrySlots(entries: readonly ChromeMenuEntry[], into = new Set<string>()): Set<string> {
  for (const entry of entries) {
    if (entry.kind === 'item') into.add(entry.slot);
    else if (entry.kind === 'submenu') entrySlots(entry.items, into);
  }
  return into;
}

/**
 * One menu of the menu bar: a trigger and the panel that it opens.
 *
 * With a registry id (`file`, `format`, `insert`, `review`, `help`), the panel holds the
 * engine's rows for that menu (`CHROME_MENUS`). Inside the menu, put `ng-template` elements to
 * change those rows in place:
 *
 * - `<ng-template docxSlot="file.save">` replaces one row, also inside a submenu.
 * - `<ng-template docxSlot="file.exportPdf" docxSlotHidden />` removes one row.
 *
 * Other content, such as `docx-menu-item` or `button[docxMenuRow]`, comes after the registry
 * rows. With `[preset]="false"`, the panel holds only your content.
 *
 * A click opens the panel and a second click closes it. While a menu is open, the pointer over
 * another trigger opens that menu. Enter, Space, and the arrow keys open the panel from the
 * keyboard, and Escape closes it.
 *
 * @example
 * ```html
 * <div docxEditorMenu>
 *   <docx-menu menu="file">
 *     <ng-template docxSlot="file.exportPdf" docxSlotHidden />
 *   </docx-menu>
 *   <docx-menu menu="tools" label="Tools">
 *     <docx-menu-item slotId="format.clear" />
 *   </docx-menu>
 * </div>
 * ```
 *
 * @public
 */
@Component({
  selector: 'docx-menu',
  changeDetection: ChangeDetectionStrategy.OnPush,
  imports: [
    NgTemplateOutlet,
    DocxChromeIconComponent,
    DocxMenuItemComponent,
    DocxMenuSeparatorComponent,
    DocxMenuSubmenuComponent,
    DocxMenuTablePickerComponent,
    DocxMenuChartPickerComponent,
    DocxMenuReviewersComponent,
    DocxMenuReportIssueComponent,
  ],
  providers: [
    { provide: DOCX_MENU_ROW_TABINDEX, useValue: -1 },
    { provide: DocxMenuBarState, useFactory: provideMenuBarState },
  ],
  host: { style: 'display: contents' },
  template: `
    <ng-template #entriesTpl let-list>
      @for (entry of list; track $index) {
        @switch (entry.kind) {
          @case ('separator') {
            <div docxMenuSeparator></div>
          }
          @case ('submenu') {
            <docx-menu-submenu [labelKey]="entry.labelKey" [paths]="entry.paths">
              <ng-container
                [ngTemplateOutlet]="entriesTpl"
                [ngTemplateOutletContext]="{ $implicit: entry.items }"
              />
            </docx-menu-submenu>
          }
          @default {
            @if (overrides().get(entry.slot); as override) {
              @if (!override.docxSlotHidden()) {
                <ng-container [ngTemplateOutlet]="override.template" />
              }
            } @else if (entry.picker === 'tableGrid') {
              <docx-menu-table-picker [labelKey]="entry.labelKey" />
            } @else if (entry.picker === 'chartType') {
              <docx-menu-chart-picker [labelKey]="entry.labelKey" />
            } @else {
              <docx-menu-item
                [slotId]="entry.slot"
                [labelKey]="entry.labelKey"
                [shortcutKey]="entry.shortcutKey"
              />
            }
          }
        }
      }
    </ng-template>
    <ng-template #rowsTpl>
      @if (preset()) {
        <ng-container
          [ngTemplateOutlet]="entriesTpl"
          [ngTemplateOutletContext]="{ $implicit: entries() }"
        />
        @if (menu() === 'review') {
          <docx-menu-submenu
            labelKey="reviewers.markupOptions"
            className="docx-menubar__markup-options"
          >
            <docx-menu-reviewers className="docx-menubar__reviewers-submenu" />
          </docx-menu-submenu>
        } @else if (menu() === 'help') {
          <docx-menu-report-issue />
        }
        @for (slot of appended(); track slot) {
          <ng-container [ngTemplateOutlet]="slot.template" />
        }
      }
      <ng-content />
    </ng-template>
    @if (shown()) {
      <div
        role="none"
        [class]="rootClass()"
        [attr.data-menu]="menu()"
        [attr.data-toolbar-group]="groupAttribute()"
        [attr.data-toolbar-more]="isMore() ? '' : null"
      >
        <button
          #trigger
          type="button"
          role="menuitem"
          aria-haspopup="menu"
          class="docx-menubar__trigger"
          [disabled]="bar.opening()"
          [attr.tabindex]="tabStop() ? 0 : -1"
          [attr.aria-expanded]="open()"
          [attr.aria-controls]="open() ? panelId : null"
          [attr.data-open]="open() ? '' : null"
          [attr.aria-label]="isMore() ? text() : null"
          [attr.title]="isMore() ? text() : null"
          (mousedown)="guard($event)"
          (keydown)="onTriggerKey($event)"
          (click)="onTriggerClick()"
          (mouseenter)="onTriggerEnter()"
          (mouseleave)="switchedByHover = false"
        >
          @if (icon(); as paths) {
            <span class="docx-menubar__trigger-icon" aria-hidden="true">
              <svg [docxChromeIcon]="paths"></svg>
            </span>
          }
          @if (!isMore()) {
            {{ text() }}
          }
        </button>
        @if (open()) {
          <div
            #panel
            role="menu"
            class="docx-toolbar__menu docx-menubar__menu"
            [id]="panelId"
            [attr.aria-label]="text()"
            (keydown)="onPanelKey($event)"
          >
            <ng-container [ngTemplateOutlet]="rowsTpl" />
          </div>
        }
      </div>
    }
  `,
})
export class DocxMenuComponent implements OnInit {
  /** The menu id: a registry id such as `'file'`, or an id of your own. */
  readonly menu = input.required<string>();
  /** The i18n key of the trigger label. Defaults to the registry's. */
  readonly labelKey = input<string | undefined>(undefined);
  /** The trigger label. It wins over `labelKey`, and names a menu of your own. */
  readonly label = input<string | undefined>(undefined);
  /** Material Symbols paths for an icon before the label. Menu bars have no icons by default. */
  readonly icon = input<readonly string[] | null | undefined>(undefined);
  /** More classes for the menu root. */
  readonly className = input<string | undefined>(undefined);
  /** Render nothing. In the menu bar, this removes the registry menu with the same id. */
  readonly hidden = input(false, { transform: booleanAttribute });
  /** Render the registry rows. With `false`, the panel holds only your content. */
  readonly preset = input(true, { transform: booleanAttribute });

  /** @internal */
  protected readonly bar: DocxMenuBarState = inject(DocxMenuBarState);
  private readonly slotTemplates = contentChildren(DocxSlotDirective);
  private readonly ownRows = contentChildren(DocxMenuItemComponent);
  private readonly ownButtons = contentChildren(DocxMenuRowComponent);
  private readonly ownSubmenus = contentChildren(DocxMenuSubmenuComponent);
  private readonly ownGroups = contentChildren(DocxMenuGroupComponent);
  private readonly menuLabel = injectMenuLabel();
  private readonly destroyRef = inject(DestroyRef);
  private readonly triggerRef = viewChild<ElementRef<HTMLElement>>('trigger');
  private readonly panelRef = viewChild<ElementRef<HTMLElement>>('panel');
  private readonly rowsTemplate: Signal<TemplateRef<unknown> | undefined> = viewChild('rowsTpl');
  private openedByKey = false;

  /** @internal */
  protected switchedByHover = false;
  /** @internal */
  protected readonly panelId = popupId('menu-panel');
  /** @internal */
  protected readonly guard: (event: Event) => void = guardToolbarMousedown;
  /** @internal */
  protected readonly isMore = computed(() => this.menu() === MENU_OVERFLOW_ID);
  /** @internal */
  protected readonly open = computed(() => this.bar.openMenu() === this.menu());

  private readonly registry = computed(() => CHROME_MENUS.find((menu) => menu.id === this.menu()));

  /** @internal */
  protected readonly entries = computed(() => this.registry()?.entries ?? []);

  /** @internal */
  protected readonly overrides = computed(() => {
    const map = new Map<ChromeSlotId, DocxSlotDirective>();
    for (const slot of this.slotTemplates()) map.set(slot.docxSlot(), slot);
    return map;
  });

  /** The slot templates that name no registry row of this menu: they come after the rows. @internal */
  protected readonly appended = computed(() => {
    const placed = entrySlots(this.entries());
    return this.slotTemplates().filter(
      (slot) => !slot.docxSlotHidden() && !placed.has(slot.docxSlot())
    );
  });

  /** @internal */
  protected readonly text = computed(
    () =>
      this.label() ?? this.menuLabel()(this.labelKey() ?? this.registry()?.labelKey ?? this.menu())
  );

  /**
   * Help with no rows renders nothing: without Report issue, the packaged Help menu is empty,
   * and a trigger that opens an empty panel helps no one.
   */
  private readonly empty = computed(
    () =>
      this.menu() === 'help' &&
      this.bar.handlers().reportIssueRow === false &&
      this.slotTemplates().length === 0 &&
      this.ownRows().length + this.ownButtons().length === 0 &&
      this.ownSubmenus().length + this.ownGroups().length === 0
  );

  /** @internal */
  protected readonly shown = computed(
    () => !this.hidden() && !this.empty() && !this.bar.overflow().has(this.menu())
  );

  /** @internal */
  protected readonly groupAttribute = computed(() =>
    this.bar.measuring() && !this.isMore() ? this.menu() : null
  );

  /** @internal */
  protected readonly rootClass = computed(() => {
    const extra = this.className();
    return `docx-menubar__menu-root${extra ? ` ${extra}` : ''}`;
  });

  /** When the active menu moved into "⋯", the "⋯" trigger holds the tab stop. @internal */
  protected readonly tabStop = computed(() => {
    const active = this.bar.activeMenu();
    return (
      active === this.menu() ||
      (this.isMore() && active !== null && this.bar.overflow().has(active))
    );
  });

  constructor() {
    afterRenderEffect(() => {
      const panel = this.panelRef()?.nativeElement;
      if (!panel) return;
      if (untracked(this.isMore)) this.clampPanel(panel);
      if (this.openedByKey) {
        this.openedByKey = false;
        focusEdge(panelItems(panel), 'first');
      }
    });
  }

  /** @internal */
  ngOnInit(): void {
    if (this.isMore()) return;
    const stop = this.bar.register({ id: this.menu(), text: this.text, rows: this.rowsTemplate });
    this.destroyRef.onDestroy(stop);
  }

  /** @internal */
  protected onTriggerKey(event: KeyboardEvent): void {
    if (['ArrowDown', 'ArrowUp', 'Enter', ' '].includes(event.key)) {
      event.preventDefault();
      this.openedByKey = true;
      this.bar.setOpenMenu(this.menu());
    } else if (event.key === 'Escape' && untracked(this.open)) {
      event.preventDefault();
      this.closeToTrigger();
    }
  }

  /** @internal */
  protected onTriggerClick(): void {
    if (this.switchedByHover) {
      this.switchedByHover = false;
      return;
    }
    this.bar.setOpenMenu(untracked(this.open) ? null : this.menu());
  }

  /** @internal */
  protected onTriggerEnter(): void {
    const current = untracked(this.bar.openMenu);
    if (current !== null && current !== this.menu()) {
      this.switchedByHover = true;
      this.bar.setOpenMenu(this.menu());
    }
  }

  /** @internal */
  protected onPanelKey(event: KeyboardEvent): void {
    const panel = untracked(this.panelRef)?.nativeElement;
    if (!panel) return;
    const items = panelItems(panel);
    const focused = hostDocumentOf(panel).activeElement;
    if (event.key === 'ArrowDown') focusBy(items, focused, 1);
    else if (event.key === 'ArrowUp') focusBy(items, focused, -1);
    else if (event.key === 'Home') focusEdge(items, 'first');
    else if (event.key === 'End') focusEdge(items, 'last');
    else if (event.key === 'Escape') {
      event.stopPropagation();
      this.closeToTrigger();
    } else {
      if (event.key === 'Tab') this.bar.setOpenMenu(null);
      return;
    }
    event.preventDefault();
  }

  private closeToTrigger(): void {
    this.bar.setOpenMenu(null);
    untracked(this.triggerRef)?.nativeElement.focus();
  }

  /** Keep the "⋯" panel inside the viewport, with the toolbar's placement rule. */
  private clampPanel(panel: HTMLElement): void {
    const trigger = untracked(this.triggerRef)?.nativeElement;
    const root = trigger?.parentElement;
    const view = trigger?.ownerDocument.defaultView;
    if (!trigger || !root || !view) return;
    const rect = trigger.getBoundingClientRect();
    const next = toolbarPanelPlacement({
      triggerLeft: rect.left,
      triggerRight: rect.right,
      panelWidth: panel.offsetWidth,
      viewportWidth: view.innerWidth,
    });
    panel.style.left = `${next.left - root.getBoundingClientRect().left}px`;
    panel.style.right = 'auto';
    panel.style.maxInlineSize = `${next.maxWidth}px`;
  }
}
