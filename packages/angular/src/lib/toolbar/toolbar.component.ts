import { NgComponentOutlet, NgTemplateOutlet } from '@angular/common';
import {
  ChangeDetectionStrategy,
  Component,
  ElementRef,
  Injector,
  NgZone,
  ViewEncapsulation,
  afterRenderEffect,
  booleanAttribute,
  computed,
  contentChildren,
  inject,
  input,
  signal,
  untracked,
  viewChild,
  type Signal,
  type TemplateRef,
} from '@angular/core';
import type { EditorSnapshot } from '@docx-editor.dev/core/contracts/editor';
import {
  TOOLBAR_COLLAPSE_ORDER,
  chartPaneAvailable,
  collapseOrder,
  isElementNode,
  shapePaneAvailable,
  toolbarGroupsKey,
  type ChromeSlotId,
  type ToolbarFitGroup,
} from '@docx-editor.dev/core/editor';
import { injectOwnScopeClassName } from '../color-mode';
import { DOCX_EDITOR_SCOPED, injectDocxEditor } from '../context';
import { injectEditorState } from '../editor-state';
import { injectChromeDir, type DocxEditorDirection } from '../i18n';
import { DocxSlotDirective } from '../slot.directive';
import {
  arrangeGroups,
  buildDefaultGroups,
  layoutToolbar,
  readSlotOverrides,
  type BarGroup,
  type PanelSection,
  type SlotOverrides,
  type ToolbarGroupSpec,
} from './toolbar-arrangement';
import {
  DocxToolbarButtonComponent,
  DocxToolbarMoreItemComponent,
} from './toolbar-button.component';
import {
  DOCX_TOOLBAR_PLACEMENT,
  DocxToolbarState,
  injectToolbarLabel,
  type DocxToolbarTranslate,
} from './toolbar-context';
import { DocxToolbarGroupDirective, hostGroupText } from './toolbar-group.directive';
import { DocxToolbarMoreComponent } from './toolbar-more.component';
import { injectToolbarOverflow } from './toolbar-overflow';
import {
  TABLE_CONTEXTUAL_GROUP_ID,
  chromeControlForSlot,
  guardToolbarMousedown,
} from './toolbar-registry';
import { DocxTableChromeState } from './table-chrome-state';
import { DocxToolbarTableChromeComponent, TABLE_GROUP_SLOTS } from './toolbar-table.component';

const selectImage = (snapshot: EditorSnapshot) => snapshot.image;
const selectChart = (snapshot: EditorSnapshot) => snapshot.chart !== null;
const selectDisabled = (snapshot: EditorSnapshot) =>
  snapshot.isLoading || snapshot.isOpening === true;
const selectTable = (snapshot: EditorSnapshot) => snapshot.table !== null;

const sameGroups = (a: readonly ToolbarFitGroup[], b: readonly ToolbarFitGroup[]) =>
  toolbarGroupsKey(a) === toolbarGroupsKey(b);
const sameOrder = (a: readonly string[], b: readonly string[]) =>
  a.join('\u0000') === b.join('\u0000');

/**
 * The formatting toolbar. Put it inside an editor root.
 *
 * The toolbar renders the engine's default groups and controls (`CHROME_GROUPS`) in registry
 * order. Each control carries its slot id in `data-slot`. When the bar has no room, controls
 * leave from the end of the groups in collapse order, into a "⋯" panel.
 *
 * Inside the toolbar, put `ng-template` elements to change the default set:
 *
 * - `<ng-template docxSlot="text.bold">` replaces one control in its place.
 * - `<ng-template docxSlot="text.strike" docxSlotHidden />` removes one control.
 * - `<ng-template docxToolbarGroup="mine" label="Mine">` adds a group of your own, and a
 *   built-in id adds content to that group or, with `hidden`, removes it.
 *
 * Other content renders after the groups, in a block that does not collapse. With
 * `[preset]="false"`, the toolbar renders only your content.
 *
 * @example
 * ```html
 * <ng-container docxEditorRoot [document]="bytes">
 *   <fieldset docxEditorToolbar>
 *     <ng-template docxSlot="text.strike" docxSlotHidden />
 *   </fieldset>
 *   <docx-editor-viewport><docx-editor-content /></docx-editor-viewport>
 * </ng-container>
 * ```
 *
 * @public
 */
@Component({
  selector: 'fieldset[docxEditorToolbar]',
  changeDetection: ChangeDetectionStrategy.OnPush,
  encapsulation: ViewEncapsulation.None,
  imports: [
    NgComponentOutlet,
    NgTemplateOutlet,
    DocxToolbarButtonComponent,
    DocxToolbarMoreItemComponent,
    DocxToolbarMoreComponent,
    DocxToolbarTableChromeComponent,
  ],
  providers: [
    DocxToolbarState,
    DocxTableChromeState,
    { provide: DOCX_EDITOR_SCOPED, useValue: true },
  ],
  host: {
    class: 'docx-toolbar',
    '[class.docx-editor]': 'scope() !== ""',
    '[class.dark]': 'scope().includes("dark")',
    'data-testid': 'docx-toolbar',
    '[disabled]': 'disabled()',
    '[attr.aria-label]': 'label()("formattingBar.label")',
    '[attr.dir]': 'chromeDir() ?? null',
    '[attr.data-overflow]': "measuring() ? '' : null",
    '(mousedown)': 'guard($event)',
  },
  template: `
    <ng-template #projected><ng-content /></ng-template>
    <ng-template #entryTpl let-entry>
      @if (overrides().get(entry.slot); as override) {
        <ng-container [ngTemplateOutlet]="override.template" />
      } @else if (entry.part) {
        <ng-container [ngComponentOutlet]="entry.part" />
      } @else {
        <button [docxToolbarButton]="entry.button"></button>
      }
    </ng-template>
    @if (preset()) {
      @for (group of bar(); track group.id) {
        @if (!$first) {
          <div role="separator" aria-orientation="vertical" class="docx-toolbar__separator"></div>
        }
        <div
          [class]="group.className"
          [attr.role]="group.heading === null ? null : 'group'"
          [attr.aria-label]="group.heading"
          [attr.data-toolbar-group]="group.pinned ? null : group.id"
          [attr.data-toolbar-fixed]="group.pinned ? '' : null"
        >
          @for (item of group.items; track item.id) {
            <div class="docx-toolbar__item" [attr.data-toolbar-item]="item.id">
              @if (item.entry) {
                <ng-container
                  [ngTemplateOutlet]="entryTpl"
                  [ngTemplateOutletContext]="{ $implicit: item.entry }"
                />
              } @else if (item.extras) {
                <ng-container [ngTemplateOutlet]="item.extras" />
              }
            </div>
          }
        </div>
      }
      @if (tableVisible() && !tableOverflowed()) {
        <div role="separator" aria-orientation="vertical" class="docx-toolbar__separator"></div>
      }
      @if (!tableOverflowed()) {
        <div class="docx-toolbar__contextual" [attr.data-toolbar-group]="tableGroupId">
          @if (tableVisible()) {
            <docx-toolbar-table-chrome [overrides]="tableTemplates()" [slots]="tableSlots()" />
          }
        </div>
      }
      <div
        #appended
        class="docx-toolbar__group"
        [attr.data-toolbar-fixed]="hasAppended() ? '' : null"
        [attr.hidden]="hasAppended() ? null : ''"
        [style.display]="hasAppended() ? null : 'none'"
      >
        <ng-container [ngTemplateOutlet]="projected" />
        @for (slot of slotOverrides().appended; track slot) {
          <ng-container [ngTemplateOutlet]="slot.template" />
        }
      </div>
      @if (sections().length > 0 || (tableVisible() && tableOverflowed())) {
        <div docxToolbarMore>
          @if (tableVisible() && tableOverflowed() && tableSlots().length > 0) {
            <div class="docx-toolbar__more-section" role="group" [attr.aria-label]="tableLabel()">
              <span class="docx-toolbar__more-heading" aria-hidden="true">{{ tableLabel() }}</span>
              @for (slot of tableSlots(); track slot) {
                <div class="docx-toolbar__more-control">
                  <span class="docx-toolbar__more-control-label">{{ tableRowLabel(slot) }}</span>
                  <span class="docx-toolbar__more-control-body">
                    <docx-toolbar-table-chrome [overrides]="tableTemplates()" [slots]="[slot]" />
                  </span>
                </div>
              }
            </div>
          }
          @for (section of sections(); track section.id) {
            <div class="docx-toolbar__more-section" role="group" [attr.aria-label]="section.label">
              <span class="docx-toolbar__more-heading" aria-hidden="true">{{ section.label }}</span>
              @for (row of section.rows; track row.key) {
                @switch (row.kind) {
                  @case ('command') {
                    <button [docxToolbarMoreItem]="$any(row).slot"></button>
                  }
                  @case ('control') {
                    <div class="docx-toolbar__more-control">
                      <span class="docx-toolbar__more-control-label">{{ $any(row).label }}</span>
                      <span class="docx-toolbar__more-control-body">
                        @if ($any(row).template) {
                          <ng-container [ngTemplateOutlet]="$any(row).template" />
                        } @else {
                          <ng-container
                            [ngTemplateOutlet]="entryTpl"
                            [ngTemplateOutletContext]="{ $implicit: $any(row).entry }"
                          />
                        }
                      </span>
                    </div>
                  }
                  @case ('row') {
                    <ng-container
                      [ngTemplateOutlet]="entryTpl"
                      [ngTemplateOutletContext]="{ $implicit: $any(row).entry }"
                    />
                  }
                  @case ('host') {
                    @if ($any(row).own) {
                      <ng-container
                        [ngTemplateOutlet]="$any(row).template"
                        [ngTemplateOutletInjector]="panelInjector"
                      />
                    } @else {
                      <div class="docx-toolbar__more-host">
                        <ng-container
                          [ngTemplateOutlet]="$any(row).template"
                          [ngTemplateOutletInjector]="panelInjector"
                        />
                      </div>
                    }
                  }
                }
              }
            </div>
          }
        </div>
      }
    } @else {
      <ng-container [ngTemplateOutlet]="projected" />
      @for (group of standaloneGroups(); track group.id) {
        <div
          role="group"
          [class]="group.className"
          [attr.aria-label]="group.heading"
          [attr.data-toolbar-host-group]="group.id"
        >
          <ng-container [ngTemplateOutlet]="group.template" />
        </div>
      }
    }
  `,
})
export class DocxEditorToolbarComponent {
  /**
   * Render the default arrangement. With `false`, the toolbar renders only its content and
   * the host groups, in a bar that does not measure itself.
   */
  readonly preset = input(true, { transform: booleanAttribute });
  /** Collapse the controls that do not fit into the "⋯" panel. With `false`, the bar wraps. */
  readonly overflow = input(true, { transform: booleanAttribute });
  /** Your own labels. A key that this function does not answer uses the active catalog. */
  readonly t = input<DocxToolbarTranslate | undefined>(undefined);

  private readonly slotTemplates = contentChildren(DocxSlotDirective);
  private readonly groupTemplates = contentChildren(DocxToolbarGroupDirective);
  private readonly appendedRef = viewChild<ElementRef<HTMLElement>>('appended');
  private readonly editor = injectDocxEditor();
  private readonly image = injectEditorState(selectImage);
  private readonly chart = injectEditorState(selectChart);

  /** @internal */
  protected readonly scope = injectOwnScopeClassName();
  /** @internal */
  protected readonly chromeDir: Signal<DocxEditorDirection | undefined> = injectChromeDir();
  /** @internal */
  protected readonly label = injectToolbarLabel();
  /** @internal */
  protected readonly disabled = injectEditorState(selectDisabled);
  /** @internal */
  protected readonly tableVisible = injectEditorState(selectTable);
  /** @internal */
  protected readonly measuring = computed(() => this.preset() && this.overflow());
  /** @internal */
  protected readonly tableGroupId = TABLE_CONTEXTUAL_GROUP_ID;
  /** @internal */
  protected readonly hasAppended = signal(false);
  /** @internal */
  protected readonly guard: (event: Event) => void = guardToolbarMousedown;
  /**
   * The injector of host content in the "⋯" panel: its `docxToolbarAction` buttons render as
   * panel rows.
   *
   * @internal
   */
  protected readonly panelInjector = Injector.create({
    providers: [{ provide: DOCX_TOOLBAR_PLACEMENT, useValue: 'panel' }],
    parent: inject(Injector),
  });

  private readonly defaultGroups = computed(() => {
    const editor = this.editor();
    // A right-to-left interface language puts the paragraph direction buttons in the bar.
    return buildDefaultGroups(
      this.image(),
      this.chromeDir() === 'rtl',
      { chart: chartPaneAvailable(editor), shape: shapePaneAvailable(editor) },
      this.chart()
    );
  });

  /** @internal */
  protected readonly slotOverrides: Signal<SlotOverrides> = computed(() =>
    readSlotOverrides(this.slotTemplates(), this.defaultGroups())
  );
  /** @internal */
  protected readonly overrides: Signal<SlotOverrides['overrides']> = computed(
    () => this.slotOverrides().overrides
  );

  private readonly groupSpecs = computed<readonly ToolbarGroupSpec[]>(() =>
    this.groupTemplates().map((group) => ({
      id: group.docxToolbarGroup(),
      label: group.label(),
      labelKey: group.labelKey(),
      priority: group.priority(),
      pinned: group.pinned(),
      after: group.after(),
      hidden: group.hidden(),
      overflowContent: group.overflowContent(),
      className: group.className(),
      template: group.template,
    }))
  );

  private readonly arranged = computed(() =>
    arrangeGroups(this.defaultGroups(), this.groupSpecs(), this.overrides())
  );

  private readonly fitGroups = computed<readonly ToolbarFitGroup[]>(
    () => {
      const overrides = this.overrides();
      return [
        ...this.arranged()
          .filter((group) => !group.pinned)
          .map((group) => ({
            id: group.id,
            items: group.entries
              .filter((entry) => overrides.get(entry.slot)?.docxSlotHidden() !== true)
              .map((entry) => entry.slot as string)
              .concat(group.extras ? ['host-content'] : []),
          })),
        // Table chrome is one unit: its border controls act together on one target.
        ...(this.tableVisible()
          ? [{ id: TABLE_CONTEXTUAL_GROUP_ID, items: [TABLE_CONTEXTUAL_GROUP_ID] }]
          : []),
      ];
    },
    { equal: sameGroups }
  );

  private readonly collapseOrder = computed(
    () => {
      const priorities = new Map<string, number>();
      for (const group of this.arranged()) {
        if (group.priority !== undefined) priorities.set(group.id, group.priority);
      }
      return collapseOrder(
        this.fitGroups().map((group) => group.id),
        TOOLBAR_COLLAPSE_ORDER,
        priorities
      );
    },
    { equal: sameOrder }
  );

  private readonly overflowAnswer = injectToolbarOverflow({
    bar: () => this.host,
    enabled: this.measuring,
    groups: this.fitGroups,
    order: this.collapseOrder,
  });

  private readonly layout = computed(() =>
    layoutToolbar(this.arranged(), this.overrides(), this.overflowAnswer(), this.label())
  );

  /** @internal */
  protected readonly bar: Signal<readonly BarGroup[]> = computed(() => this.layout().bar);
  /** @internal */
  protected readonly tableOverflowed = computed(() =>
    this.overflowAnswer().has(TABLE_CONTEXTUAL_GROUP_ID)
  );
  /** @internal */
  protected readonly sections: Signal<readonly PanelSection[]> = computed(
    () => this.layout().sections
  );
  /** The table slots that the host did not hide, in bar order. @internal */
  protected readonly tableSlots = computed(() => {
    const overrides = this.slotOverrides().tableOverrides;
    return TABLE_GROUP_SLOTS.filter((slot) => overrides.get(slot)?.docxSlotHidden() !== true);
  });
  /** The host's templates for table slots. @internal */
  protected readonly tableTemplates = computed(() => {
    const templates = new Map<string, TemplateRef<unknown>>();
    for (const [slot, override] of this.slotOverrides().tableOverrides) {
      if (!override.docxSlotHidden()) templates.set(slot, override.template);
    }
    return templates;
  });
  /** @internal */
  protected readonly tableLabel = computed(() => this.label()('formattingBar.groups.table'));

  /** The label of a table row in the "⋯" panel. @internal */
  protected tableRowLabel(slot: ChromeSlotId): string {
    return this.label()(chromeControlForSlot(slot)?.labelKey ?? 'formattingBar.groups.table');
  }

  /** @internal */
  protected readonly standaloneGroups = computed(() => {
    const label = this.label();
    return this.groupSpecs()
      .filter((spec) => !spec.hidden)
      .map((spec) => ({
        id: spec.id,
        heading: hostGroupText(spec, label),
        className: `docx-toolbar__group${spec.className ? ` ${spec.className}` : ''}`,
        template: spec.template,
      }));
  });

  private readonly host: HTMLElement = inject(ElementRef).nativeElement;

  constructor() {
    inject(DocxToolbarState).useTranslate(() => this.t());
    const zone = inject(NgZone);
    // The block after the groups shows only when your content renders something in it.
    afterRenderEffect((onCleanup) => {
      const element = this.appendedRef()?.nativeElement;
      if (!element) return;
      const check = () => {
        const filled = [...element.childNodes].some(
          (node) => isElementNode(node) || (node.nodeType === 3 && node.textContent?.trim())
        );
        if (filled !== untracked(this.hasAppended)) zone.run(() => this.hasAppended.set(filled));
      };
      check();
      const view = element.ownerDocument.defaultView;
      if (!view || typeof view.MutationObserver === 'undefined') return;
      const observer = zone.runOutsideAngular(() => new view.MutationObserver(check));
      observer.observe(element, { childList: true, characterData: true, subtree: true });
      onCleanup(() => observer.disconnect());
    });
  }
}
