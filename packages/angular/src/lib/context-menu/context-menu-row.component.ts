import {
  ChangeDetectionStrategy,
  Component,
  NgZone,
  booleanAttribute,
  computed,
  inject,
  input,
  untracked,
} from '@angular/core';
import type { EditorCommand } from '@docx-editor.dev/core/contracts/editor';
import {
  chartPaneAvailable,
  hostWindowOf,
  openChartPaneFor,
  shapePaneAvailable,
  tableChromeIconPaths,
  type HostWindow,
} from '@docx-editor.dev/core/editor';
import { injectDocxEditor, injectDocxEditorContext } from '../context';
import {
  injectEditorCommand,
  type EditorCommandState,
  type EditorCommandTarget,
} from '../editor-command';
import { injectEditorState } from '../editor-state';
import { DocxMenuRowComponent } from '../menu/menu-row.component';
import { injectMenuLabel } from '../menu/menu-state';
import { DocxChromeIconComponent } from '../toolbar/chrome-icon.component';
import { chromeControlForSlot } from '../toolbar/toolbar-registry';
import {
  COPY_PATHS,
  CUT_PATHS,
  DELETE_PATHS,
  PASTE_PATHS,
  REFRESH_TOC_PAGE_NUMBERS_PATHS,
  REFRESH_TOC_PATHS,
  SELECT_ALL_PATHS,
} from './context-menu-icons';
import { DocxContextMenuState, type ContextMenuRowId } from './context-menu-state';

/**
 * A packaged row of the context menu: a row id that is not a chrome slot, or one of the two
 * object rows that are.
 *
 * @public
 */
export type ContextMenuPackagedRow =
  | Exclude<ContextMenuRowId, 'table.cellVerticalAlignment'>
  | 'image.editChart'
  | 'image.formatShape';

/** How a row runs, and when it shows. */
type RowKind =
  'command' | 'paste' | 'pastePlain' | 'table' | 'splitCell' | 'toc' | 'field' | 'chart' | 'shape';

interface RowDefinition {
  readonly kind: RowKind;
  readonly labelKey: string;
  readonly shortcutKey?: string;
  readonly paths: readonly string[];
  /** The engine command, or the command for a table of contents id. */
  readonly command?: EditorCommand | ((tocId: string | null) => EditorCommand);
  readonly destructive?: boolean;
}

/** The paintbrush of the registry, so the rows and the toolbar button draw one glyph. */
const PAINTER_PATHS: readonly string[] = chromeControlForSlot('format.painter')?.paths ?? [];

/** The table of contents that the menu opened on. Without an id, the engine picks one. */
const withToc = (tocId: string | null) => (tocId ? { tocId } : {});

const ROWS: Readonly<Record<ContextMenuPackagedRow, RowDefinition>> = {
  'field.edit': {
    kind: 'field',
    labelKey: 'textFormField.edit',
    paths: [],
    command: { type: 'editTextFormField' },
  },
  'image.editChart': { kind: 'chart', labelKey: 'chart.edit', paths: [] },
  'image.formatShape': {
    kind: 'shape',
    labelKey: 'shape.formatShape',
    paths: chromeControlForSlot('image.formatShape')?.paths ?? [],
  },
  'edit.cut': {
    kind: 'command',
    labelKey: 'contextMenu.cut',
    shortcutKey: 'contextMenu.cutShortcut',
    paths: CUT_PATHS,
    command: { type: 'cut' },
  },
  'edit.copy': {
    kind: 'command',
    labelKey: 'contextMenu.copy',
    shortcutKey: 'contextMenu.copyShortcut',
    paths: COPY_PATHS,
    command: { type: 'copy' },
  },
  'edit.paste': {
    kind: 'paste',
    labelKey: 'contextMenu.paste',
    shortcutKey: 'contextMenu.pasteShortcut',
    paths: PASTE_PATHS,
    command: { type: 'paste', text: ' ' },
  },
  'edit.pasteWithoutFormatting': {
    kind: 'pastePlain',
    labelKey: 'contextMenu.pasteWithoutFormatting',
    shortcutKey: 'contextMenu.pasteWithoutFormattingShortcut',
    paths: PASTE_PATHS,
    command: { type: 'pasteWithoutFormatting', text: ' ' },
  },
  'format.copyFormatting': {
    kind: 'command',
    labelKey: 'contextMenu.copyFormatting',
    shortcutKey: 'contextMenu.copyFormattingShortcut',
    paths: PAINTER_PATHS,
    command: { type: 'copyFormatting' },
  },
  'format.pasteFormatting': {
    kind: 'command',
    labelKey: 'contextMenu.pasteFormatting',
    shortcutKey: 'contextMenu.pasteFormattingShortcut',
    paths: PAINTER_PATHS,
    command: { type: 'pasteFormatting' },
  },
  'edit.delete': {
    kind: 'command',
    labelKey: 'contextMenu.delete',
    shortcutKey: 'contextMenu.deleteShortcut',
    paths: DELETE_PATHS,
    command: { type: 'deleteText' },
  },
  'edit.selectAll': {
    kind: 'command',
    labelKey: 'contextMenu.selectAll',
    shortcutKey: 'contextMenu.selectAllShortcut',
    paths: SELECT_ALL_PATHS,
    command: { type: 'selectAll' },
  },
  'table.selectTable': {
    kind: 'table',
    labelKey: 'table.selectTable',
    paths: tableChromeIconPaths('select_all'),
    command: { type: 'selectTableRegion', region: 'table' },
  },
  'table.insertRowAbove': {
    kind: 'table',
    labelKey: 'table.insertRowAbove',
    paths: tableChromeIconPaths('table_rows'),
    command: { type: 'insertRow', where: 'above' },
  },
  'table.insertRowBelow': {
    kind: 'table',
    labelKey: 'table.insertRowBelow',
    paths: tableChromeIconPaths('table_rows'),
    command: { type: 'insertRow', where: 'below' },
  },
  'table.insertColumnLeft': {
    kind: 'table',
    labelKey: 'table.insertColumnLeft',
    paths: tableChromeIconPaths('view_column'),
    command: { type: 'insertColumn', where: 'left' },
  },
  'table.insertColumnRight': {
    kind: 'table',
    labelKey: 'table.insertColumnRight',
    paths: tableChromeIconPaths('view_column'),
    command: { type: 'insertColumn', where: 'right' },
  },
  'table.mergeCells': {
    kind: 'table',
    labelKey: 'table.mergeCells',
    paths: tableChromeIconPaths('call_merge'),
    command: { type: 'mergeCells' },
  },
  'table.splitCell': {
    kind: 'splitCell',
    labelKey: 'table.splitCell',
    paths: tableChromeIconPaths('call_split'),
  },
  'table.deleteRow': {
    kind: 'table',
    labelKey: 'table.deleteRow',
    paths: tableChromeIconPaths('delete_sweep'),
    command: { type: 'deleteRow' },
    destructive: true,
  },
  'table.deleteColumn': {
    kind: 'table',
    labelKey: 'table.deleteColumn',
    paths: tableChromeIconPaths('view_column'),
    command: { type: 'deleteColumn' },
    destructive: true,
  },
  'table.deleteTable': {
    kind: 'table',
    labelKey: 'table.deleteTable',
    paths: tableChromeIconPaths('delete'),
    command: { type: 'deleteTable' },
    destructive: true,
  },
  'toc.refresh': {
    kind: 'toc',
    labelKey: 'toc.refresh',
    paths: REFRESH_TOC_PATHS,
    command: (tocId) => ({ type: 'refreshToc', mode: 'entire', ...withToc(tocId) }),
  },
  'toc.refreshPageNumbers': {
    kind: 'toc',
    labelKey: 'toc.refreshPageNumbers',
    paths: REFRESH_TOC_PAGE_NUMBERS_PATHS,
    command: (tocId) => ({ type: 'refreshToc', mode: 'pageNumbers', ...withToc(tocId) }),
  },
  'toc.remove': {
    kind: 'toc',
    labelKey: 'toc.remove',
    paths: DELETE_PATHS,
    command: (tocId) => ({ type: 'removeToc', ...withToc(tocId) }),
  },
};

/** True for a row id that `docx-context-menu-row` renders. @internal */
export function isPackagedContextMenuRow(id: string): id is ContextMenuPackagedRow {
  return Object.hasOwn(ROWS, id);
}

/**
 * Both clipboard flavors where the async read allows it, plain text where it does not. The
 * engine's paste router uses the flavors that survive.
 */
async function readClipboardPayload(
  view: HostWindow
): Promise<{ text: string; html: string | null }> {
  // The editor's own window: a pop-out window has the focus, and the opener's clipboard
  // refuses a read from a window without it.
  const clipboard = view.navigator.clipboard;
  if (typeof clipboard.read === 'function' && typeof view.ClipboardItem !== 'undefined') {
    try {
      const items = await clipboard.read();
      let text = '';
      let html: string | null = null;
      for (const item of items) {
        if (item.types.includes('text/plain')) {
          text = await (await item.getType('text/plain')).text();
        }
        if (item.types.includes('text/html')) {
          html = await (await item.getType('text/html')).text();
        }
      }
      if (text || html) return { text, html };
    } catch {
      // Use the plain read. A refusal there is the one that the row reports.
    }
  }
  return { text: await clipboard.readText(), html: null };
}

const refusalOf = (error: unknown): string =>
  error instanceof Error ? error.message : 'the clipboard is not readable';

/**
 * One packaged row of the context menu, such as Cut, Insert row above, or Refresh table of
 * contents. The engine decides whether the row is enabled. Choosing the row runs its command,
 * closes the menu, and gives the focus back to the document.
 *
 * Table rows show only in a table, table of contents rows only on a table of contents, and the
 * object rows only on their object. Use the row in an override template to change its label,
 * icon, or class.
 *
 * @example
 * ```html
 * <div docxEditorContextMenu>
 *   <ng-template docxContextMenuRow="edit.cut">
 *     <docx-context-menu-row row="edit.cut" labelKey="myApp.cut" className="my-cut" />
 *   </ng-template>
 * </div>
 * ```
 *
 * @public
 */
@Component({
  selector: 'docx-context-menu-row',
  changeDetection: ChangeDetectionStrategy.OnPush,
  imports: [DocxMenuRowComponent],
  host: { style: 'display: contents' },
  template: `@if (shown()) {
    <button
      [docxMenuRow]="row()"
      [icon]="paths()"
      [shortcut]="shortcut()"
      [disabled]="!enabled()"
      [reason]="reason()"
      [rowClass]="rowClass()"
      (activate)="run()"
    >
      {{ text() }}
    </button>
  }`,
})
export class DocxContextMenuRowComponent {
  /** The row id, such as `'edit.cut'` or `'table.insertRowAbove'`. */
  readonly row = input.required<ContextMenuPackagedRow>();
  /** The i18n key of the label, in place of the default. */
  readonly labelKey = input<string | undefined>(undefined);
  /** The i18n key of the shortcut text, in place of the default. */
  readonly shortcutKey = input<string | undefined>(undefined);
  /** Material Symbols paths for the icon, in place of the default. */
  readonly icon = input<readonly string[] | undefined>(undefined);
  /** More classes for the row. */
  readonly className = input<string | undefined>(undefined);
  /** Mark a table row as destructive. Delete row, column, and table are destructive. */
  readonly destructive = input<boolean | undefined>(undefined);
  /** Render nothing. */
  readonly hidden = input(false, { transform: booleanAttribute });

  private readonly menu = inject(DocxContextMenuState);
  private readonly context = injectDocxEditorContext();
  private readonly editor = injectDocxEditor();
  private readonly zone = inject(NgZone);
  private readonly label = injectMenuLabel();
  private readonly definition = computed(() => ROWS[this.row()]);
  private readonly kind = computed(() => this.definition().kind);
  private readonly inTable = injectEditorState((snapshot) => snapshot.table != null);
  private readonly chartId = injectEditorState((snapshot) => snapshot.chart?.drawingNodeId ?? null);
  private readonly shapeId = injectEditorState((snapshot) =>
    snapshot.image?.shapeFormat ? snapshot.image.id : null
  );
  private readonly command = injectEditorCommand((): EditorCommandTarget => {
    const definition = this.definition();
    if (definition.kind === 'splitCell') return 'table.splitCell';
    if (definition.kind === 'shape') return 'image.formatShape';
    const command = definition.command;
    if (typeof command === 'function') return command(this.menu.tocId());
    return command ?? { type: 'selectAll' };
  });

  /** @internal */
  protected readonly shown = computed(() => {
    if (this.hidden()) return false;
    switch (this.kind()) {
      case 'table':
        return this.inTable();
      case 'splitCell':
        return this.inTable() && this.menu.splitCell() !== undefined;
      case 'toc':
        return this.menu.tocId() !== null;
      case 'field':
        return this.command.isEnabled();
      case 'chart': {
        const editor = this.editor();
        return !!editor && this.chartId() !== null && untracked(() => chartPaneAvailable(editor));
      }
      case 'shape': {
        const editor = this.editor();
        return !!editor && this.shapeId() !== null && untracked(() => shapePaneAvailable(editor));
      }
      default:
        return true;
    }
  });

  private readonly blocked = computed(() => {
    const kind = this.kind();
    return (kind === 'paste' || kind === 'pastePlain') && this.menu.clipboardRefusal() !== null;
  });

  /** @internal */
  protected readonly enabled = computed(
    () => this.kind() === 'chart' || (this.command.isEnabled() && !this.blocked())
  );

  /** @internal */
  protected readonly reason = computed(() =>
    this.blocked() ? this.menu.clipboardRefusal() : this.command.disabledReason()
  );

  /** @internal */
  protected readonly paths = computed(() => {
    const own = this.icon() ?? this.definition().paths;
    return own.length > 0 ? own : null;
  });

  /** @internal */
  protected readonly text = computed(() =>
    this.label()(this.labelKey() ?? this.definition().labelKey)
  );

  /** @internal */
  protected readonly shortcut = computed(() => {
    const key = this.shortcutKey() ?? this.definition().shortcutKey;
    return key ? this.label()(key) : undefined;
  });

  /** @internal */
  protected readonly rowClass = computed(() => {
    const classes: string[] = [];
    if (this.kind() === 'table' && (this.destructive() ?? this.definition().destructive)) {
      classes.push('docx-table-chrome__destructive-row');
    }
    const extra = this.className();
    if (extra) classes.push(extra);
    return classes.length > 0 ? classes.join(' ') : undefined;
  });

  /** @internal */
  protected run(): void {
    const kind = untracked(this.kind);
    switch (kind) {
      case 'paste':
      case 'pastePlain':
        this.paste(kind === 'pastePlain');
        return;
      case 'table':
        if (this.command.execute()) this.menu.close(true);
        return;
      case 'splitCell': {
        const open = untracked(this.menu.splitCell);
        // The focus goes back to the document first, so that closing the dialog returns it there.
        this.menu.close(true);
        open?.();
        return;
      }
      case 'chart': {
        const editor = untracked(this.editor);
        const id = untracked(this.chartId);
        this.menu.close(true);
        if (editor && id) this.zone.runOutsideAngular(() => openChartPaneFor(editor, id));
        return;
      }
      case 'field':
      case 'shape':
        // Close first: the command opens a dialog or a pane that takes the focus.
        this.menu.close(true);
        this.command.execute();
        return;
      default:
        this.command.execute();
        this.menu.close(true);
    }
  }

  /**
   * Close first: the read can wait on a browser permission prompt, and a menu that stays open
   * over it looks stuck. The read still starts inside this click.
   */
  private paste(plain: boolean): void {
    const editor = untracked(this.editor);
    const view = hostWindowOf(untracked(this.context?.contentElement ?? (() => null)));
    const menu = this.menu;
    menu.close(true);
    const read = plain
      ? view.navigator.clipboard.readText().then((text) => ({ text, html: null }))
      : readClipboardPayload(view);
    read.then(
      ({ text, html }) => {
        if (!editor || !(text || html)) return;
        this.zone.runOutsideAngular(() =>
          editor.exec(
            plain
              ? { type: 'pasteWithoutFormatting', text }
              : { type: 'paste', text, ...(html ? { html } : {}) }
          )
        );
      },
      (error: unknown) => menu.clipboardRefusal.set(refusalOf(error))
    );
  }
}

const VERTICAL_ALIGNMENTS = [
  { alignment: 'top', labelKey: 'tableAdvanced.top', icon: 'vertical_align_top' },
  { alignment: 'center', labelKey: 'tableAdvanced.middle', icon: 'vertical_align_center' },
  { alignment: 'bottom', labelKey: 'tableAdvanced.bottom', icon: 'vertical_align_bottom' },
] as const;

/**
 * The vertical alignment of the table cells at the selection: top, middle, and bottom buttons
 * in one row of the context menu. It shows only in a table.
 *
 * @public
 */
@Component({
  selector: 'docx-context-menu-cell-alignment',
  changeDetection: ChangeDetectionStrategy.OnPush,
  imports: [DocxChromeIconComponent],
  host: { style: 'display: contents' },
  template: `@if (!hidden() && inTable()) {
    <div class="docx-contextmenu__table-align" data-slot="table.cellVerticalAlignment">
      <span class="docx-contextmenu__table-align-label">{{ heading() }}</span>
      <div class="docx-contextmenu__table-align-buttons" role="group" [attr.aria-label]="heading()">
        @for (item of items; track item.alignment; let index = $index) {
          <button
            type="button"
            role="menuitemradio"
            aria-checked="false"
            tabindex="-1"
            class="docx-contextmenu__table-align-button"
            [attr.aria-label]="label()(item.labelKey)"
            [attr.title]="commands[index]!.disabledReason() ?? label()(item.labelKey)"
            [attr.aria-disabled]="!commands[index]!.isEnabled()"
            (mousedown)="$event.preventDefault(); $event.stopPropagation()"
            (click)="align(index)"
          >
            <svg [docxChromeIcon]="item.paths"></svg>
          </button>
        }
      </div>
    </div>
  }`,
})
export class DocxContextMenuCellAlignmentComponent {
  /** Render nothing. */
  readonly hidden = input(false, { transform: booleanAttribute });

  private readonly menu = inject(DocxContextMenuState);

  /** @internal */
  protected readonly label = injectMenuLabel();
  /** @internal */
  protected readonly inTable = injectEditorState((snapshot) => snapshot.table != null);
  /** @internal */
  protected readonly heading = computed(() => this.label()('tableAdvanced.verticalAlignment'));
  /** @internal */
  protected readonly items = VERTICAL_ALIGNMENTS.map((item) => ({
    ...item,
    paths: tableChromeIconPaths(item.icon),
  }));
  /** @internal */
  protected readonly commands: EditorCommandState[] = VERTICAL_ALIGNMENTS.map(({ alignment }) =>
    injectEditorCommand({ type: 'setTableCellVerticalAlignment', alignment })
  );

  /** @internal */
  protected align(index: number): void {
    if (this.commands[index]!.execute()) this.menu.close(true);
  }
}
