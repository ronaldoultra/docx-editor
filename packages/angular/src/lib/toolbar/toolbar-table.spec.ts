// Task 4.4: the table, table border, and table cell fill controls, and Insert Table. The table
// chrome claim of packages/vue/test/chrome-composition.test.ts and the table grid claims of
// toolbar-host-groups.test.ts.
import { ChangeDetectionStrategy, Component, Directive, signal, viewChild } from '@angular/core';
import type { DocxEditorInstance } from '@docx-editor.dev/core/editor';
import { createT, en, type TranslationKey } from '@docx-editor.dev/core/i18n';
import { afterEach, beforeEach, describe, expect, test, vi } from 'vitest';
import { docx, paragraph } from '../../testing/docx';
import { configureZoneless, engineOf, render, settle } from '../../testing/harness';
import {
  MockResizeObserver,
  flushObservers,
  panelOf,
  restoreResizeObserver,
  setWidth,
  toolbarOf,
  useMockResizeObserver,
} from '../../testing/toolbar';
import { DocxEditorContentComponent } from '../content.component';
import { DocxEditorRootDirective } from '../root.directive';
import { DocxEditorViewportComponent } from '../viewport.component';
import { DocxToolbarActionComponent } from './toolbar-action.component';
import { DocxToolbarButtonComponent } from './toolbar-button.component';
import { DocxToolbarGroupDirective } from './toolbar-group.directive';
import { DocxToolbarTableInsertComponent } from './toolbar-table-insert.component';
import {
  DocxToolbarTableBorderTargetComponent,
  injectTableBorderTargetLabel,
} from './toolbar-table.component';
import { DocxEditorToolbarComponent } from './toolbar.component';

const label = createT(en);
const TABLE_SOURCE = docx(
  paragraph('outside table') +
    '<w:tbl><w:tblGrid><w:gridCol w:w="3600"/></w:tblGrid>' +
    `<w:tr><w:tc>${paragraph('inside table')}</w:tc></w:tr></w:tbl>`
);
const t = (key: string) => (key === 'formattingBar.more' ? 'More' : undefined);

@Component({ selector: 'border-target-label', template: '{{ text() }}' })
class BorderTargetLabelComponent {
  readonly text = injectTableBorderTargetLabel();
}

const PARTS = [
  DocxEditorRootDirective,
  DocxEditorViewportComponent,
  DocxEditorContentComponent,
  DocxEditorToolbarComponent,
  DocxToolbarButtonComponent,
  DocxToolbarActionComponent,
  DocxToolbarGroupDirective,
  DocxToolbarTableInsertComponent,
  DocxToolbarTableBorderTargetComponent,
  BorderTargetLabelComponent,
];

@Directive()
abstract class Host {
  readonly bytes = signal<Uint8Array>(TABLE_SOURCE);
  readonly t = t;
  readonly root = viewChild.required(DocxEditorRootDirective);

  editor(): DocxEditorInstance {
    return engineOf(this.root().context.editor());
  }
}

@Component({
  imports: PARTS,
  changeDetection: ChangeDetectionStrategy.OnPush,
  template: `<ng-container docxEditorRoot [document]="bytes()">
    <fieldset docxEditorToolbar [t]="t"></fieldset>
    <docx-editor-viewport><docx-editor-content /></docx-editor-viewport>
  </ng-container>`,
})
class DefaultHostComponent extends Host {}

@Component({
  imports: PARTS,
  changeDetection: ChangeDetectionStrategy.OnPush,
  template: `<ng-container docxEditorRoot [document]="bytes()">
    <fieldset docxEditorToolbar [t]="t"><border-target-label /></fieldset>
    <docx-editor-viewport><docx-editor-content /></docx-editor-viewport>
  </ng-container>`,
})
class LabelHostComponent extends Host {}

@Component({
  imports: PARTS,
  changeDetection: ChangeDetectionStrategy.OnPush,
  template: `<ng-container docxEditorRoot [document]="bytes()">
    <fieldset docxEditorToolbar [preset]="false" [overflow]="false">
      <span docxToolbarTableInsert></span>
      <button docxToolbarButton="table.mergeCells"></button>
      <div docxToolbarTableBorderTarget></div>
    </fieldset>
    <button type="button" data-testid="elsewhere">elsewhere</button>
    <docx-editor-viewport><docx-editor-content /></docx-editor-viewport>
  </ng-container>`,
})
class BareHostComponent extends Host {}

@Component({
  imports: PARTS,
  changeDetection: ChangeDetectionStrategy.OnPush,
  template: `<ng-container docxEditorRoot [document]="bytes()">
    <fieldset docxEditorToolbar [t]="t">
      <ng-template docxToolbarGroup="ins" label="Insert">
        <span docxToolbarTableInsert></span>
      </ng-template>
    </fieldset>
    <docx-editor-viewport><docx-editor-content /></docx-editor-viewport>
  </ng-container>`,
})
class InsertGroupHostComponent extends Host {}

async function mount<T extends Host>(component: new () => T) {
  const fixture = await render(component);
  const element: HTMLElement = fixture.nativeElement;
  document.body.append(element);
  const update = async (action: () => unknown = () => {}) => {
    action();
    await settle(fixture, 2);
  };
  const editor = () => fixture.componentInstance.editor();
  /** Put the caret in paragraph `index`: 0 is outside the table, 1 is in its cell. */
  const caretIn = (index: number) =>
    update(() => {
      const paragraphId = editor().surface!.session.paragraphIds()[index]!;
      editor().surface!.setSelection({
        anchor: { paragraphId, offset: 1 },
        head: { paragraphId, offset: 1 },
      });
    });
  return { fixture, element, update, editor, caretIn };
}

beforeEach(() => configureZoneless());
afterEach(() => {
  restoreResizeObserver();
  document.body.replaceChildren();
});

describe('contextual table chrome', () => {
  test('shows in a table only, and collapses before the ordinary formatting groups', async () => {
    useMockResizeObserver();
    const { element, update, caretIn } = await mount(DefaultHostComponent);
    const toolbar = toolbarOf(element);
    const contextual = () => toolbar.querySelector<HTMLElement>('.docx-toolbar__contextual');
    expect(contextual()!.getAttribute('data-toolbar-group')).toBe('contextual-table');
    expect(contextual()!.hasAttribute('data-toolbar-fixed')).toBe(false);
    expect(contextual()!.querySelector('[data-slot]')).toBeNull();
    expect(
      MockResizeObserver.instances.some((observer) => observer.observed.includes(contextual()!))
    ).toBe(true);

    const sizeGroups = () => {
      setWidth(toolbar, 'clientWidth', () => width);
      for (const group of toolbar.querySelectorAll<HTMLElement>('[data-toolbar-group]')) {
        setWidth(group, 'offsetWidth', () =>
          group === contextual() && group.querySelector('[data-slot]') ? 220 : 40
        );
      }
      for (const fixed of toolbar.querySelectorAll('[data-toolbar-fixed]')) {
        setWidth(fixed, 'offsetWidth', 40);
      }
      const separator = toolbar.querySelector('.docx-toolbar__separator');
      if (separator) setWidth(separator, 'offsetWidth', 1);
    };
    // Wide enough for every group and the 220 px table group.
    let width = 900;
    sizeGroups();
    await flushObservers();
    expect(toolbar.querySelector('[data-slot="toolbar.more"]')).toBeNull();

    await caretIn(1);
    expect(contextual()!.querySelector('[data-slot="table.borderTarget"]')).not.toBeNull();
    width = 320;
    sizeGroups();
    await flushObservers();
    const trigger = toolbar.querySelector<HTMLButtonElement>('[data-slot="toolbar.more"]');
    expect(trigger).not.toBeNull();
    expect(toolbar.querySelector('[data-slot="text.color"]')).not.toBeNull();
    expect(contextual()).toBeNull();
    await update(() => trigger!.click());
    const panel = panelOf(element)!;
    const sections = panel.querySelectorAll<HTMLElement>('.docx-toolbar__more-section');
    expect(sections.length).toBeGreaterThan(1);
    expect(sections[0]!.getAttribute('aria-label')).toBe(
      label('formattingBar.groups.table' as TranslationKey)
    );
    expect(panel.querySelector('[data-slot="table.borderTarget"]')).not.toBeNull();
    expect(panel.querySelector('[data-slot="text.color"]')).toBeNull();
    await update(() =>
      panel.querySelector<HTMLButtonElement>('[data-slot="table.borderTarget"] button')!.click()
    );
    expect(panel.querySelector('.docx-table-chrome__panel')).not.toBeNull();
    // The table menu's own Escape closes the menu, and only the menu.
    const row = panel.querySelector<HTMLElement>(
      '.docx-table-chrome__panel [role="menuitemradio"]'
    )!;
    await update(() =>
      row.dispatchEvent(
        new KeyboardEvent('keydown', { key: 'Escape', bubbles: true, cancelable: true })
      )
    );
    expect(panel.querySelector('.docx-table-chrome__panel')).toBeNull();
    expect(panelOf(element)).toBe(panel);
    // Escape from outside the editor closes the panel and leaves the key alone.
    const outside = new KeyboardEvent('keydown', {
      key: 'Escape',
      bubbles: true,
      cancelable: true,
    });
    await update(() => document.body.dispatchEvent(outside));
    expect(outside.defaultPrevented).toBe(false);
    expect(panelOf(element)).toBeNull();

    width = 900;
    await caretIn(0);
    sizeGroups();
    await flushObservers();
    expect(toolbar.querySelector('[data-slot="toolbar.more"]')).toBeNull();
    expect(contextual()?.querySelector('[data-slot]')).toBeNull();
  });

  test('a border target applies to the draft, and the target label follows it', async () => {
    const { element, update, caretIn } = await mount(LabelHostComponent);
    await caretIn(1);
    const root = element.querySelector<HTMLElement>('[data-slot="table.borderTarget"]')!;
    const trigger = root.querySelector<HTMLButtonElement>('[aria-haspopup="menu"]')!;
    expect(trigger.getAttribute('aria-label')).toBe(
      label('table.borders.tooltip' as TranslationKey)
    );
    expect(element.querySelector('border-target-label')!.textContent).toBe(
      label('table.borders.all' as TranslationKey)
    );
    await update(() => trigger.click());
    const all = root.querySelector<HTMLElement>('[data-value="all"]')!;
    // The open menu takes the focus on the target in force.
    expect(document.activeElement).toBe(all);
    expect(all.getAttribute('aria-checked')).toBe('true');
    await update(() => root.querySelector<HTMLButtonElement>('[data-value="outside"]')!.click());
    expect(root.querySelector('[role="menu"]')).toBeNull();
    expect(element.querySelector('border-target-label')!.textContent).toBe(
      label('table.borders.outside' as TranslationKey)
    );
    await update(() => trigger.click());
    expect(root.querySelector('[data-value="outside"]')!.getAttribute('aria-checked')).toBe('true');
  });

  test('a cell fill applies, and the clear row removes it', async () => {
    const { element, update, editor, caretIn } = await mount(DefaultHostComponent);
    await caretIn(1);
    const root = element.querySelector<HTMLElement>('[data-slot="table.cellFill"]')!;
    const exec = vi.spyOn(editor(), 'exec');
    await update(() => root.querySelector<HTMLButtonElement>('[aria-haspopup="dialog"]')!.click());
    await update(() => root.querySelector<HTMLButtonElement>('[data-value="00B050"]')!.click());
    expect(root.querySelector('[role="dialog"]')).toBeNull();
    expect(exec.mock.calls.some(([command]) => JSON.stringify(command).includes('00B050'))).toBe(
      true
    );
    exec.mockClear();
    await update(() => root.querySelector<HTMLButtonElement>('[aria-haspopup="dialog"]')!.click());
    await update(() =>
      root.querySelector<HTMLButtonElement>('.docx-toolbar__swatch-clear')!.click()
    );
    expect(exec).toHaveBeenCalled();
    expect(editor().snapshot().canUndo).toBe(true);
  });

  test('in viewing mode, the table triggers stay focusable and announce the engine reason', async () => {
    const { element, update, editor, caretIn } = await mount(DefaultHostComponent);
    await caretIn(1);
    await update(() => editor().exec({ type: 'setEditingMode', mode: 'viewing' }));
    const trigger = element.querySelector<HTMLButtonElement>(
      '[data-slot="table.borderTarget"] [aria-haspopup="menu"]'
    )!;
    expect(trigger.disabled).toBe(false);
    expect(trigger.getAttribute('aria-disabled')).toBe('true');
    const reasonId = trigger.getAttribute('aria-describedby')!;
    const reason = element.querySelector(`[id="${reasonId}"]`)!;
    expect(reason.textContent!.length).toBeGreaterThan(0);
    expect(trigger.title).toBe(reason.textContent);
    await update(() => trigger.click());
    expect(element.querySelector('[data-slot="table.borderTarget"] [role="menu"]')).toBeNull();
  });

  test('outside a table, the table parts render nothing, and merge is disabled with a reason', async () => {
    const { element, caretIn } = await mount(BareHostComponent);
    await caretIn(0);
    const target = element.querySelector<HTMLElement>('[data-slot="table.borderTarget"]')!;
    expect(target.querySelector('button')).toBeNull();
    const merge = element.querySelector<HTMLButtonElement>('[data-slot="table.mergeCells"]')!;
    expect(merge.disabled).toBe(true);
    expect(merge.title.length).toBeGreaterThan(0);
    expect(merge.title).not.toBe(merge.getAttribute('aria-label'));
  });
});

describe('Insert Table', () => {
  test('opens a size grid and inserts the picked size', async () => {
    const { element, update, editor, caretIn } = await mount(BareHostComponent);
    await caretIn(0);
    const trigger = element.querySelector<HTMLButtonElement>('[data-slot="table.insert"]')!;
    expect(trigger.disabled).toBe(false);
    expect(trigger.getAttribute('aria-haspopup')).toBe('grid');
    await update(() => trigger.click());
    const grid = element.querySelector('[role="grid"]')!;
    expect(grid.querySelectorAll('[role="gridcell"]').length).toBe(36);
    const exec = vi.spyOn(editor(), 'exec');
    await update(() => grid.querySelector<HTMLButtonElement>('[data-cell="2x3"]')!.click());
    expect(exec).toHaveBeenCalledWith({ type: 'insertTable', rows: 2, cols: 3 });
    expect(element.querySelector('[role="grid"]')).toBeNull();
  });

  test('a second click on the trigger closes the grid and focuses the document', async () => {
    const { element, update, editor } = await mount(BareHostComponent);
    const trigger = element.querySelector<HTMLButtonElement>('[data-slot="table.insert"]')!;
    const focus = vi.spyOn(editor(), 'focus').mockImplementation(() => ({ ok: true }) as never);
    await update(() => trigger.click());
    expect(element.querySelector('[role="grid"]')).not.toBeNull();
    await update(() => trigger.click());
    expect(element.querySelector('[role="grid"]')).toBeNull();
    expect(focus).toHaveBeenCalledTimes(1);
  });

  test('focus that leaves the grid closes it', async () => {
    const { element, update } = await mount(BareHostComponent);
    const trigger = element.querySelector<HTMLButtonElement>('[data-slot="table.insert"]')!;
    await update(() => trigger.click());
    const cell = element.querySelector<HTMLElement>('[role="gridcell"]')!;
    await update(() =>
      cell.dispatchEvent(new FocusEvent('focusout', { bubbles: true, relatedTarget: trigger }))
    );
    expect(element.querySelector('[role="grid"]')).not.toBeNull();
    const elsewhere = element.querySelector('[data-testid="elsewhere"]')!;
    await update(() =>
      cell.dispatchEvent(new FocusEvent('focusout', { bubbles: true, relatedTarget: elsewhere }))
    );
    expect(element.querySelector('[role="grid"]')).toBeNull();
  });

  test('the grid moves left to stay inside the viewport', async () => {
    const { element, update } = await mount(BareHostComponent);
    const trigger = element.querySelector<HTMLButtonElement>('[data-slot="table.insert"]')!;
    const view = element.ownerDocument.defaultView!;
    trigger.parentElement!.getBoundingClientRect = () =>
      ({ left: view.innerWidth - 40, right: view.innerWidth - 6, width: 34 }) as DOMRect;
    const offsetWidth = Object.getOwnPropertyDescriptor(HTMLElement.prototype, 'offsetWidth');
    Object.defineProperty(HTMLElement.prototype, 'offsetWidth', {
      configurable: true,
      get(this: HTMLElement) {
        return this.classList.contains('docx-toolbar__table-insert-menu') ? 200 : 0;
      },
    });
    try {
      await update(() => trigger.click());
    } finally {
      if (offsetWidth) Object.defineProperty(HTMLElement.prototype, 'offsetWidth', offsetWidth);
    }
    const popup = element.querySelector<HTMLElement>('.docx-toolbar__table-insert-menu')!;
    expect(popup.style.left).toBe('-168px');
  });

  test('in the "⋯" panel, Escape closes the grid, and an insert closes the panel', async () => {
    useMockResizeObserver();
    const { element, update } = await mount(InsertGroupHostComponent);
    const toolbar = toolbarOf(element);
    setWidth(toolbar, 'clientWidth', 420);
    for (const group of toolbar.querySelectorAll('[data-toolbar-group], [data-toolbar-fixed]')) {
      setWidth(group, 'offsetWidth', 90);
    }
    await flushObservers();
    const openPanel = async () => {
      await update(() => toolbar.querySelector<HTMLButtonElement>('[aria-label="More"]')!.click());
      return panelOf(element)!;
    };
    let panel = await openPanel();
    await update(() =>
      panel.querySelector<HTMLButtonElement>('[data-slot="table.insert"]')!.click()
    );
    expect(panel.querySelector('[role="grid"]')).not.toBeNull();
    // An open moves the focus into the grid, so its own handler sees Escape.
    expect(document.activeElement?.getAttribute('role')).toBe('gridcell');
    await update(() =>
      document.activeElement!.dispatchEvent(
        new KeyboardEvent('keydown', { key: 'Escape', bubbles: true })
      )
    );
    expect(panel.querySelector('[role="grid"]')).toBeNull();
    panel = panelOf(element) ?? (await openPanel());
    await update(() =>
      panel.querySelector<HTMLButtonElement>('[data-slot="table.insert"]')!.click()
    );
    await update(() => panel.querySelector<HTMLButtonElement>('[data-cell="2x2"]')!.click());
    expect(panelOf(element)).toBeNull();
  });
});
