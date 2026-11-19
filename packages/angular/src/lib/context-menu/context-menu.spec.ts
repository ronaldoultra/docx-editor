// Task 5.2: the claims of packages/vue/test/context-menu.test.ts and context-menu-chart.test.ts.
import { ChangeDetectionStrategy, Component, Directive, signal, viewChild } from '@angular/core';
import {
  listenPaneRequests,
  takePaneRequest,
  type DocxEditorInstance,
} from '@docx-editor.dev/core/editor';
import { afterEach, beforeEach, describe, expect, test } from 'vitest';
import {
  contextMenuPanel,
  contextMenuSlots,
  pressKey,
  rightClick,
  scrollerOf,
} from '../../testing/context-menu';
import { docx, paragraph } from '../../testing/docx';
import { configureZoneless, engineOf, render, settle } from '../../testing/harness';
import { DocxEditorContentComponent } from '../content.component';
import { DocxEditorRootDirective } from '../root.directive';
import { DocxEditorViewportComponent } from '../viewport.component';
import { DocxEditorContextMenuComponent } from './context-menu.component';

const t = (key: string): string => key;
const SOURCE = docx(paragraph('hello world'));
const TABLE_SOURCE = docx(
  paragraph('outside table') +
    '<w:tbl><w:tblGrid><w:gridCol w:w="3600"/></w:tblGrid>' +
    `<w:tr><w:tc>${paragraph('inside table')}</w:tc></w:tr></w:tbl>`
);
const TOC_SOURCE = docx(
  '<w:p><w:r><w:t>Before</w:t></w:r></w:p>' +
    '<w:p><w:r><w:fldChar w:fldCharType="begin"/><w:instrText> TOC \\o "1-1" \\h </w:instrText><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>Old entry</w:t></w:r></w:p>' +
    '<w:p><w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>' +
    '<w:p><w:r><w:t>After</w:t></w:r></w:p>'
);
const FIELD_SOURCE = docx(
  '<w:p><w:r><w:fldChar w:fldCharType="begin"><w:ffData><w:textInput><w:default w:val="Sample"/></w:textInput></w:ffData></w:fldChar></w:r><w:r><w:instrText> FORMTEXT </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>Sample</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>'
);

@Directive()
abstract class Host {
  readonly bytes = signal<Uint8Array>(SOURCE);
  readonly t = t;
  readonly opens: boolean[] = [];
  readonly root = viewChild.required(DocxEditorRootDirective);

  editor(): DocxEditorInstance {
    return engineOf(this.root().context.editor());
  }
}

@Component({
  imports: [
    DocxEditorRootDirective,
    DocxEditorViewportComponent,
    DocxEditorContentComponent,
    DocxEditorContextMenuComponent,
  ],
  changeDetection: ChangeDetectionStrategy.OnPush,
  template: `<ng-container docxEditorRoot [document]="bytes()">
    <docx-editor-viewport>
      <docx-editor-content />
      <div docxEditorContextMenu [t]="t" (openChange)="opens.push($event)"></div>
    </docx-editor-viewport>
  </ng-container>`,
})
class MenuHostComponent extends Host {}

@Component({
  imports: [
    DocxEditorRootDirective,
    DocxEditorViewportComponent,
    DocxEditorContentComponent,
    DocxEditorContextMenuComponent,
  ],
  changeDetection: ChangeDetectionStrategy.OnPush,
  template: `<ng-container docxEditorRoot [document]="bytes()">
    <docx-editor-viewport>
      <docx-editor-content />
      <div docxEditorContextMenu [t]="t" [splitCellHandler]="splitCell"></div>
    </docx-editor-viewport>
  </ng-container>`,
})
class SplitCellHostComponent extends Host {
  splits = 0;
  readonly splitCell = () => (this.splits += 1);
}

async function mount<T extends Host>(
  component: new () => T,
  bytes: Uint8Array = SOURCE
): Promise<{ root: HTMLElement; host: T; destroy: () => void }> {
  const fixture = await render(component);
  const root: HTMLElement = fixture.nativeElement;
  document.body.append(root);
  if (bytes !== SOURCE) {
    fixture.componentInstance.bytes.set(bytes);
    await settle(fixture);
  }
  return {
    root,
    host: fixture.componentInstance,
    destroy: () => {
      fixture.destroy();
      root.remove();
    },
  };
}

const openMenu = (root: Element) => rightClick(scrollerOf(root));

/** Put the caret in paragraph `index` of the body: 1 is the table cell of `TABLE_SOURCE`. */
function caretAt(editor: DocxEditorInstance, index: number, offset = 1): void {
  const surface = editor.surface!;
  const paragraphId = surface.session.paragraphIds()[index]!;
  surface.setSelection({ anchor: { paragraphId, offset }, head: { paragraphId, offset } });
}

function clipboard(value: Partial<Clipboard>): void {
  Object.defineProperty(globalThis.navigator, 'clipboard', { value, configurable: true });
}

describe('docxEditorContextMenu', () => {
  beforeEach(() => configureZoneless());
  afterEach(() => {
    document.body.replaceChildren();
  });

  test('a right-click opens the panel at the pointer on the first open', async () => {
    const view = await mount(MenuHostComponent);
    expect(view.host.editor().surface).toBeTruthy();
    expect(contextMenuPanel(view.root)).toBeNull();
    const event = await openMenu(view.root);
    expect(event.defaultPrevented).toBe(true);
    const panel = contextMenuPanel(view.root)!;
    expect(panel.style.left).toBe('120px');
    expect(panel.style.top).toBe('140px');
    expect(panel.getAttribute('aria-label')).toBe('contextMenu.ariaLabel');
    expect(view.host.opens).toEqual([true]);
    view.destroy();
  });

  test('engine state changes keep the one contextmenu listener', async () => {
    const view = await mount(MenuHostComponent);
    const scroller = scrollerOf(view.root);
    const added: string[] = [];
    const removed: string[] = [];
    const add = scroller.addEventListener.bind(scroller);
    const remove = scroller.removeEventListener.bind(scroller);
    scroller.addEventListener = ((type: string, ...rest: [never, never?]) => {
      added.push(type);
      return add(type, ...rest);
    }) as typeof scroller.addEventListener;
    scroller.removeEventListener = ((type: string, ...rest: [never, never?]) => {
      removed.push(type);
      return remove(type, ...rest);
    }) as typeof scroller.removeEventListener;
    const editor = view.host.editor();
    editor.setZoom(1.25);
    await settle();
    editor.setZoom(1.5);
    await settle();
    expect(added.filter((type) => type === 'contextmenu')).toEqual([]);
    expect(removed.filter((type) => type === 'contextmenu')).toEqual([]);
    await openMenu(view.root);
    expect(contextMenuPanel(view.root)).not.toBeNull();
    view.destroy();
  });

  test('the default set is the packaged rows, without Add comment', async () => {
    const view = await mount(MenuHostComponent);
    await openMenu(view.root);
    expect(contextMenuSlots(view.root)).toEqual([
      'edit.cut',
      'edit.copy',
      'edit.paste',
      'edit.pasteWithoutFormatting',
      'format.copyFormatting',
      'format.pasteFormatting',
      'edit.delete',
      'edit.selectAll',
      'text.link',
    ]);
    expect(view.root.querySelector('[data-slot="review.addComment"]')).toBeNull();
    view.destroy();
  });

  test('Escape closes the panel and gives the focus back to the document', async () => {
    const view = await mount(MenuHostComponent);
    const editor = view.host.editor();
    let focused = 0;
    const focus = editor.focus.bind(editor);
    editor.focus = (scope) => {
      focused += 1;
      return focus(scope);
    };
    await openMenu(view.root);
    expect(document.activeElement).toBe(contextMenuPanel(view.root));
    await pressKey(document.body, { key: 'Escape' });
    expect(contextMenuPanel(view.root)).toBeNull();
    // jsdom cannot focus the editable pages, so the spec counts the engine's focus call.
    expect(focused).toBe(1);
    expect(view.host.opens).toEqual([true, false]);
    view.destroy();
  });

  test('Paste closes the panel before the clipboard read settles', async () => {
    let resolveRead: (text: string) => void = () => {};
    clipboard({ readText: () => new Promise<string>((resolve) => (resolveRead = resolve)) });
    const view = await mount(MenuHostComponent);
    const surface = view.host.editor().surface!;
    const paragraphId = surface.session.paragraphIds()[0]!;
    surface.setSelection({ anchor: { paragraphId, offset: 0 }, head: { paragraphId, offset: 0 } });
    await openMenu(view.root);
    view.root.querySelector<HTMLElement>('[data-slot="edit.paste"]')!.click();
    await settle();
    // The read waits on the permission prompt. The menu does not wait with it.
    expect(contextMenuPanel(view.root)).toBeNull();
    resolveRead('pasted ');
    await settle();
    expect(surface.session.bodyText()).toBe('pasted hello world');
    view.destroy();
  });

  test('a refused read disables Paste with the reason on the next open', async () => {
    clipboard({ readText: () => Promise.reject(new Error('Read permission denied.')) });
    const view = await mount(MenuHostComponent);
    await openMenu(view.root);
    view.root.querySelector<HTMLElement>('[data-slot="edit.paste"]')!.click();
    await settle();
    expect(contextMenuPanel(view.root)).toBeNull();
    await openMenu(view.root);
    const row = view.root.querySelector<HTMLElement>('[data-slot="edit.paste"]')!;
    expect(row.getAttribute('aria-disabled')).toBe('true');
    expect(row.getAttribute('title')).toBe('Read permission denied.');
    view.destroy();
  });

  test('the field row shows on a form field and opens the field options dialog', async () => {
    // jsdom has no modal dialogs.
    HTMLDialogElement.prototype.showModal ??= function showModal(this: HTMLDialogElement) {
      this.setAttribute('open', '');
    };
    const view = await mount(MenuHostComponent, FIELD_SOURCE);
    const field = view.root.querySelector<HTMLElement>('[data-field-atom="form"]')!;
    expect(field).not.toBeNull();
    await rightClick(field, { x: 30, y: 30 });
    const row = view.root.querySelector<HTMLElement>('[data-slot="field.edit"]')!;
    expect(row).not.toBeNull();
    expect(row.textContent).toContain('textFormField.edit');
    row.click();
    await settle();
    expect(contextMenuPanel(view.root)).toBeNull();
    expect(view.root.querySelector('dialog')).not.toBeNull();
    view.destroy();
  });

  test('the table of contents rows include Remove, and Remove deletes the table', async () => {
    const view = await mount(MenuHostComponent, TOC_SOURCE);
    const toc = view.root.querySelector<HTMLElement>('[data-docx-read-only]')!;
    await rightClick(toc, { x: 30, y: 30 });
    expect(contextMenuSlots(view.root).slice(-3)).toEqual([
      'toc.refresh',
      'toc.refreshPageNumbers',
      'toc.remove',
    ]);
    const remove = view.root.querySelector<HTMLElement>('[data-slot="toc.remove"]')!;
    expect(remove.textContent).toContain('toc.remove');
    expect(remove.querySelector('svg')).not.toBeNull();
    remove.click();
    await settle();
    expect(contextMenuPanel(view.root)).toBeNull();
    expect(view.root.querySelector('[data-docx-read-only]')).toBeNull();
    expect(view.host.editor().surface!.session.bodyText()).toBe('Before\nAfter');
    view.destroy();
  });

  test('in a table the table rows follow, and Insert row below adds a row', async () => {
    const view = await mount(MenuHostComponent, TABLE_SOURCE);
    const editor = view.host.editor();
    await openMenu(view.root);
    expect(contextMenuSlots(view.root)).not.toContain('table.insertRowBelow');
    await pressKey(document.body, { key: 'Escape' });

    caretAt(editor, 1);
    await settle();
    await openMenu(view.root);
    expect(contextMenuSlots(view.root).slice(9).filter(Boolean)).toEqual([
      'table.selectTable',
      'table.insertRowAbove',
      'table.insertRowBelow',
      'table.insertColumnLeft',
      'table.insertColumnRight',
      'table.mergeCells',
      'table.deleteRow',
      'table.deleteColumn',
      'table.deleteTable',
    ]);
    const panel = contextMenuPanel(view.root)!;
    expect(panel.querySelector('[data-slot="table.splitCell"]')).toBeNull();
    expect(
      panel
        .querySelector('[data-slot="table.deleteTable"]')!
        .classList.contains('docx-table-chrome__destructive-row')
    ).toBe(true);
    const alignment = panel.querySelector('[data-slot="table.cellVerticalAlignment"]')!;
    expect(alignment.querySelectorAll('[role="menuitemradio"]')).toHaveLength(3);

    const before = editor.surface!.session.paragraphIds().length;
    panel.querySelector<HTMLElement>('[data-slot="table.insertRowBelow"]')!.click();
    await settle();
    expect(contextMenuPanel(view.root)).toBeNull();
    // The new row holds one empty paragraph.
    expect(editor.surface!.session.paragraphIds()).toHaveLength(before + 1);
    view.destroy();
  });

  test('Split cells shows with a handler and runs it after the menu closes', async () => {
    const view = await mount(SplitCellHostComponent, TABLE_SOURCE);
    caretAt(view.host.editor(), 1);
    await settle();
    await openMenu(view.root);
    const row = view.root.querySelector<HTMLElement>('[data-slot="table.splitCell"]')!;
    expect(row).not.toBeNull();
    row.click();
    await settle();
    expect(contextMenuPanel(view.root)).toBeNull();
    expect(view.host.splits).toBe(1);
    view.destroy();
  });
});

describe('the Edit chart row (context-menu-chart.test.ts)', () => {
  beforeEach(() => configureZoneless());
  afterEach(() => {
    document.body.replaceChildren();
  });

  async function withChart(view: { host: Host }): Promise<DocxEditorInstance> {
    const editor = view.host.editor();
    caretAt(editor, 0);
    expect(editor.exec({ type: 'insertChart', chartType: 'pie' }).ok).toBe(true);
    await settle();
    expect(editor.snapshot().chart).not.toBeNull();
    return editor;
  }

  test('shows only for a selected chart and asks for the Chart pane', async () => {
    const view = await mount(MenuHostComponent);
    const stop = listenPaneRequests(view.host.editor(), () => {});
    await openMenu(view.root);
    expect(view.root.querySelector('[data-slot="image.editChart"]')).toBeNull();
    await pressKey(document.body, { key: 'Escape' });

    const editor = await withChart(view);
    takePaneRequest(editor);
    await openMenu(view.root);
    const row = view.root.querySelector<HTMLElement>('[data-slot="image.editChart"]');
    expect(row?.textContent).toContain('chart.edit');
    row!.click();
    await settle();
    expect(takePaneRequest(editor)?.pane).toBe('chart');
    stop();
    view.destroy();
  });

  test('stays out of the menu when the host shows no chart pane', async () => {
    const view = await mount(MenuHostComponent);
    await withChart(view);
    await openMenu(view.root);
    expect(contextMenuPanel(view.root)).not.toBeNull();
    expect(view.root.querySelector('[data-slot="image.editChart"]')).toBeNull();
    view.destroy();
  });
});
