import { NgTemplateOutlet } from '@angular/common';
import {
  ChangeDetectionStrategy,
  Component,
  ElementRef,
  NgZone,
  ViewEncapsulation,
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
} from '@angular/core';
import type { EditorSnapshot } from '@docx-editor.dev/core/contracts/editor';
import {
  CHROME_MENUS,
  editorHostElement,
  editorInstanceScope,
  hostDocumentOf,
  hostWindowOf,
  isDomNode,
  isElementNode,
  type ToolbarFitGroup,
} from '@docx-editor.dev/core/editor';
import { injectOwnScopeClassName } from '../color-mode';
import { DOCX_EDITOR_SCOPED, injectDocxEditor } from '../context';
import { injectEditorHostDocument, injectEditorState } from '../editor-state';
import { injectChromeDir, type DocxEditorDirection } from '../i18n';
import { DOCX_PARAGRAPH_DIALOG } from '../toolbar/toolbar-paragraph.component';
import { injectToolbarOverflow } from '../toolbar/toolbar-overflow';
import { injectPopupDismiss } from '../toolbar/toolbar-popup';
import { MORE_PATHS, guardToolbarMousedown } from '../toolbar/toolbar-registry';
import { download, downloadName } from './download';
import { barTriggers, isPoppedOutEditorTarget } from './menu-keyboard';
import { DocxMenuComponent } from './menu.component';
import {
  DOCX_MENU_ROW_TABINDEX,
  DocxMenuBarState,
  MENU_OVERFLOW_ID,
  injectMenuLabel,
  type DocxMenuTranslate,
  type MenuHandlers,
  type MenuRegistration,
} from './menu-state';
import { DocxMenuSubmenuComponent } from './menu-submenu.component';

const REGISTRY_IDS = CHROME_MENUS.map((menu) => menu.id as string);
const DOCX_MIME = 'application/vnd.openxmlformats-officedocument.wordprocessingml.document';

const selectOpening = (snapshot: EditorSnapshot): boolean => snapshot.isOpening === true;
const sameIds = (a: readonly string[], b: readonly string[]) =>
  a.join('\u0000') === b.join('\u0000');

/**
 * The menu bar. Put it inside an editor root.
 *
 * The bar renders the engine's menus (`CHROME_MENUS`): File, Format, Insert, Review, and Help.
 * Each menu carries its id in `data-menu`, and each row carries its slot id in `data-slot`.
 * When the menus do not fit on one line, the menus at the end move into one "⋯" menu, where
 * each one becomes a submenu with the same rows.
 *
 * Inside the bar, put `docx-menu` elements to change the default menus:
 *
 * - `<docx-menu menu="file">` replaces the File menu in its place. Write the id as a plain
 *   attribute, so that the bar can put the menu in its place.
 * - `<docx-menu menu="review" hidden />` removes the Review menu.
 * - `<docx-menu menu="tools" label="Tools">` adds a menu after the registry menus.
 *
 * Other content renders after the menus, and never moves into "⋯". With `[preset]="false"`,
 * the bar renders only your content.
 *
 * File › Open and Ctrl+O open a file picker, and File › Save and Ctrl+S download the document.
 * Give `openHandler` or `saveHandler` to replace them.
 *
 * @example
 * ```html
 * <ng-container docxEditorRoot [document]="bytes">
 *   <div docxEditorMenu fileName="Report"></div>
 *   <docx-editor-viewport><docx-editor-content /></docx-editor-viewport>
 * </ng-container>
 * ```
 *
 * @public
 */
@Component({
  selector: 'div[docxEditorMenu]',
  changeDetection: ChangeDetectionStrategy.OnPush,
  encapsulation: ViewEncapsulation.None,
  imports: [NgTemplateOutlet, DocxMenuComponent, DocxMenuSubmenuComponent],
  providers: [
    DocxMenuBarState,
    { provide: DOCX_MENU_ROW_TABINDEX, useValue: 0 },
    { provide: DOCX_EDITOR_SCOPED, useValue: true },
  ],
  host: {
    role: 'menubar',
    class: 'docx-menubar',
    'data-testid': 'docx-menubar',
    '[class.docx-editor]': 'scope() !== ""',
    '[class.dark]': 'scope().includes("dark")',
    '[attr.aria-label]': 'label()("titleBar.menuBarAriaLabel")',
    '[attr.dir]': 'chromeDir() ?? null',
    '[attr.data-overflow]': "measuring() ? '' : null",
    '(mousedown)': 'guard($event)',
    '(keydown)': 'onKey($event)',
  },
  template: `
    @if (preset() && !hostMenus().has('file')) {
      <docx-menu menu="file" />
    }
    <ng-content select="docx-menu[menu=file]" />
    @if (preset() && !hostMenus().has('format')) {
      <docx-menu menu="format" />
    }
    <ng-content select="docx-menu[menu=format]" />
    @if (preset() && !hostMenus().has('insert')) {
      <docx-menu menu="insert" />
    }
    <ng-content select="docx-menu[menu=insert]" />
    @if (preset() && !hostMenus().has('review')) {
      <docx-menu menu="review" />
    }
    <ng-content select="docx-menu[menu=review]" />
    @if (preset() && !hostMenus().has('help')) {
      <docx-menu menu="help" />
    }
    <ng-content select="docx-menu[menu=help]" />
    <ng-content select="docx-menu" />
    <div
      #hostBlock
      role="none"
      class="docx-menubar__host"
      [attr.data-toolbar-fixed]="measuring() && hasHostContent() ? '' : null"
      [style.display]="hasHostContent() ? null : 'none'"
    >
      <ng-content />
    </div>
    @if (collapsed().length > 0) {
      <docx-menu [menu]="moreId" [icon]="morePaths" [preset]="false" [label]="moreLabel()">
        @for (entry of collapsed(); track entry.id) {
          <docx-menu-submenu [label]="entry.text()">
            @if (entry.rows(); as rows) {
              <ng-container [ngTemplateOutlet]="rows" />
            }
          </docx-menu-submenu>
        }
      </docx-menu>
    }
    <input
      #fileInput
      type="file"
      accept=".docx,application/vnd.openxmlformats-officedocument.wordprocessingml.document"
      style="display: none"
      (change)="onFileChange($event)"
    />
  `,
})
export class DocxEditorMenuComponent {
  /** Render the registry menus. With `false`, the bar renders only your content. */
  readonly preset = input(true, { transform: booleanAttribute });
  /**
   * Move the menus that do not fit into one "⋯" menu. With `false`, the bar wraps onto more
   * lines. The bar must be able to shrink, for example with `min-width: 0` in a flex row.
   */
  readonly overflow = input(true, { transform: booleanAttribute });
  /** Your own labels. A key that this function does not answer uses the active catalog. */
  readonly t = input<DocxMenuTranslate | undefined>(undefined);
  /** The name of the saved file. `.docx` is added. Defaults to the name of the opened file. */
  readonly fileName = input<string | undefined>(undefined);
  /** Replaces File › Open and Ctrl+O. */
  readonly openHandler = input<(() => void) | undefined>(undefined);
  /** Replaces File › Save and Ctrl+S. */
  readonly saveHandler = input<(() => void) | undefined>(undefined);
  /** Opens page setup from File › Page setup. Without it, the row is disabled. */
  readonly pageSetupHandler = input<((trigger: HTMLElement | null) => void) | undefined>(undefined);
  /**
   * Opens the paragraph dialog from Format › Paragraph. Defaults to the dialog that the
   * `DOCX_PARAGRAPH_DIALOG` token gives. Without either, the row is disabled.
   */
  readonly paragraphDialogHandler = input<((trigger: HTMLElement | null) => void) | undefined>(
    undefined
  );
  /** Opens the watermark dialog from Insert › Watermark. Without it, the row is disabled. */
  readonly watermarkDialogHandler = input<((trigger: HTMLElement | null) => void) | undefined>(
    undefined
  );
  /** Replaces Help › Report issue, which opens this project's issue form. */
  readonly reportIssueHandler = input<(() => void) | undefined>(undefined);
  /** Show Help › Report issue. With `false`, Help has no packaged row, and an empty Help leaves. */
  readonly reportIssue = input(true, { transform: booleanAttribute });
  /** Emits the file that the user picks in File › Open, before the editor loads it. */
  readonly openFile = output<File>();

  private readonly state = inject(DocxMenuBarState);
  private readonly editor = injectDocxEditor();
  private readonly zone = inject(NgZone);
  private readonly paragraphDialog = inject(DOCX_PARAGRAPH_DIALOG, { optional: true });
  private readonly host: HTMLElement = inject(ElementRef).nativeElement;
  private readonly hostDocument = injectEditorHostDocument();
  private readonly ownMenus = contentChildren(DocxMenuComponent);
  private readonly hostBlock = viewChild.required<ElementRef<HTMLElement>>('hostBlock');
  private readonly fileInput = viewChild.required<ElementRef<HTMLInputElement>>('fileInput');
  private readonly openedName = signal<string | null>(null);

  /** @internal */
  protected readonly scope = injectOwnScopeClassName();
  /** @internal */
  protected readonly chromeDir: Signal<DocxEditorDirection | undefined> = injectChromeDir();
  /** @internal */
  protected readonly label = injectMenuLabel();
  /** @internal */
  protected readonly guard: (event: Event) => void = guardToolbarMousedown;
  /** @internal */
  protected readonly moreId = MENU_OVERFLOW_ID;
  /** @internal */
  protected readonly morePaths = MORE_PATHS;
  /** @internal */
  protected readonly measuring = computed(() => this.overflow());
  /** @internal */
  protected readonly hasHostContent = signal(false);
  /** @internal */
  protected readonly moreLabel = computed(() => this.label()('formattingBar.more'));
  /** The ids of the registry menus that your content replaces. @internal */
  protected readonly hostMenus = computed(
    () => new Set(this.ownMenus().map((menu) => menu.menu()))
  );

  /** The menus in bar order: the registry menus first, then the others. */
  private readonly menuIds = computed(
    () => {
      const ids = this.state.menus().map((menu) => menu.id);
      const known = REGISTRY_IDS.filter((id) => ids.includes(id));
      return [...known, ...ids.filter((id) => !REGISTRY_IDS.includes(id))];
    },
    { equal: sameIds }
  );

  private readonly answer = injectToolbarOverflow({
    bar: () => this.host,
    enabled: this.measuring,
    groups: computed<readonly ToolbarFitGroup[]>(() =>
      this.menuIds().map((id) => ({ id, items: [id] }))
    ),
    // The end of the bar leaves first: Help, then Review, and so on toward File.
    order: computed(() => [...this.menuIds()].reverse(), { equal: sameIds }),
  });

  /** The menus in the "⋯" menu, in bar order. @internal */
  protected readonly collapsed: Signal<MenuRegistration[]> = computed(() => {
    const overflow = this.state.overflow();
    const menus = this.state.menus();
    return this.menuIds()
      .filter((id) => overflow.has(id))
      .map((id) => menus.find((menu) => menu.id === id)!)
      .filter(Boolean);
  });

  constructor() {
    const state = this.state;
    state.useTranslate(() => this.t());
    state.useOpening(injectEditorState(selectOpening));
    state.useHandlers(() => this.handlers());
    effect(() => state.measuring.set(this.measuring()));
    effect(() => {
      const keys = new Set(this.answer().keys());
      untracked(() => state.overflow.set(keys));
    });
    // The previous document stays mounted while the next one opens: no menu acts on it.
    effect(() => {
      if (state.opening()) untracked(() => state.setOpenMenu(null));
    });
    const open = computed(() => state.openMenu() !== null);
    injectPopupDismiss({
      open,
      root: () => this.host,
      // Keys inside the bar stay with the menus' own handlers.
      skip: (event) => isDomNode(event.target) && this.host.contains(event.target),
      close: () => state.setOpenMenu(null),
    });
    this.bindShortcuts();
    this.watchHostContent();
    afterRenderEffect(() => {
      if (untracked(state.activeMenu) !== null) return;
      const first = barTriggers(this.host)[0]?.closest('[data-menu]')?.getAttribute('data-menu');
      if (first) untracked(() => state.activeMenu.set(first));
    });
  }

  private readonly handlers = computed<MenuHandlers>(() => {
    const ready = this.editor() !== null;
    const paragraph = this.paragraphDialogHandler();
    const dialog = this.paragraphDialog;
    return {
      open: ready ? (this.openHandler() ?? (() => this.packagedOpen())) : undefined,
      save: ready ? (this.saveHandler() ?? (() => this.packagedSave())) : undefined,
      pageSetup: ready ? this.pageSetupHandler() : undefined,
      paragraphDialog: ready
        ? (paragraph ?? (dialog ? (trigger) => dialog.open(trigger) : undefined))
        : undefined,
      watermarkDialog: ready ? this.watermarkDialogHandler() : undefined,
      reportIssue: this.reportIssueHandler(),
      reportIssueRow: this.reportIssue(),
    };
  });

  /** @internal */
  protected onKey(event: KeyboardEvent): void {
    if (event.key !== 'ArrowLeft' && event.key !== 'ArrowRight') return;
    const triggers = barTriggers(this.host);
    const index = triggers.indexOf(hostDocumentOf(this.host).activeElement as HTMLElement);
    if (index === -1) return;
    event.preventDefault();
    const rtl = hostWindowOf(this.host).getComputedStyle(this.host).direction === 'rtl';
    const step = (event.key === 'ArrowRight' ? 1 : -1) * (rtl ? -1 : 1);
    const next = triggers[(index + step + triggers.length) % triggers.length];
    const id = next?.closest('[data-menu]')?.getAttribute('data-menu');
    if (!next || !id) return;
    this.state.activeMenu.set(id);
    next.focus();
    if (untracked(this.state.openMenu) !== null) this.state.setOpenMenu(id);
  }

  /** @internal */
  protected onFileChange(event: Event): void {
    const element = event.target as HTMLInputElement;
    const file = element.files?.[0];
    element.value = '';
    const editor = untracked(this.editor);
    if (!file || !editor) return;
    this.openedName.set(file.name);
    this.openFile.emit(file);
    this.zone.runOutsideAngular(() => {
      void file
        .arrayBuffer()
        .then((buffer) => editor.load(new Uint8Array(buffer)))
        .catch((error: unknown) => console.error('[docx-editor] could not open the file', error));
    });
  }

  private packagedOpen(): void {
    this.fileInput().nativeElement.click();
  }

  private packagedSave(): void {
    const editor = untracked(this.editor);
    if (!editor) return;
    const name = downloadName(untracked(this.fileName) ?? untracked(this.openedName) ?? undefined);
    this.zone.runOutsideAngular(() => {
      void editor
        .save()
        .then((buffer) => download(buffer, name, DOCX_MIME, this.host))
        .catch((error: unknown) => console.error('[docx-editor] save failed', error));
    });
  }

  /**
   * Ctrl+O and Ctrl+S, or Cmd on macOS, from this editor: its pages, its chrome, and its pages
   * in a pop-out window. The keys listen in the bar's document and in the editor's document.
   */
  private bindShortcuts(): void {
    afterRenderEffect((onCleanup) => {
      const editorDocument = this.hostDocument();
      const handlers = this.handlers();
      const barDocument = hostDocumentOf(this.host);
      const keydown = (event: KeyboardEvent) => {
        if (!(event.metaKey || event.ctrlKey) || event.altKey || event.shiftKey) return;
        const key = event.key.toLowerCase();
        const run = key === 's' ? handlers.save : key === 'o' ? handlers.open : undefined;
        const target = event.target;
        if (!run || !isDomNode(target)) return;
        const editor = untracked(this.editor);
        const scope = editorInstanceScope(this.host) ?? this.host;
        const inScope =
          scope.contains(target) ||
          this.host.contains(target) ||
          isPoppedOutEditorTarget(target, barDocument, editor ? editorHostElement(editor) : null);
        if (!inScope) return;
        event.preventDefault();
        this.zone.run(run);
      };
      const docs =
        editorDocument && editorDocument !== barDocument
          ? [barDocument, editorDocument]
          : [barDocument];
      this.zone.runOutsideAngular(() => {
        for (const doc of docs) doc.addEventListener('keydown', keydown);
      });
      onCleanup(() => {
        for (const doc of docs) doc.removeEventListener('keydown', keydown);
      });
    });
  }

  /** The block after the menus shows only when your content renders something in it. */
  private watchHostContent(): void {
    afterRenderEffect((onCleanup) => {
      const element = this.hostBlock().nativeElement;
      const check = () => {
        const filled = [...element.childNodes].some(
          (node) => isElementNode(node) || (node.nodeType === 3 && node.textContent?.trim())
        );
        if (filled !== untracked(this.hasHostContent)) {
          this.zone.run(() => this.hasHostContent.set(filled));
        }
      };
      check();
      const view = element.ownerDocument.defaultView;
      if (!view || typeof view.MutationObserver === 'undefined') return;
      const observer = this.zone.runOutsideAngular(() => new view.MutationObserver(check));
      observer.observe(element, { childList: true, characterData: true, subtree: true });
      onCleanup(() => observer.disconnect());
    });
  }
}
