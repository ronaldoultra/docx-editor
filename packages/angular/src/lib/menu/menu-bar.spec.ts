// Task 5.1: the claims of packages/vue/test/menu-composition.test.ts and menu-opening.test.ts.
import { ChangeDetectionStrategy, Component, Directive, signal, viewChild } from '@angular/core';
import { CHROME_MENUS, type DocxEditorInstance } from '@docx-editor.dev/core/editor';
import { afterEach, beforeEach, describe, expect, test } from 'vitest';
import { docx, largeDocx, paragraph } from '../../testing/docx';
import { configureZoneless, engineOf, render, settle } from '../../testing/harness';
import { measure, restoreResizeObserver, useMockResizeObserver } from '../../testing/toolbar';
import { DocxEditorContentComponent } from '../content.component';
import { DocxEditorRootDirective } from '../root.directive';
import { DocxSlotDirective } from '../slot.directive';
import { DocxEditorViewportComponent } from '../viewport.component';
import { DocxEditorMenuComponent } from './menu-bar.component';
import { DocxMenuItemComponent } from './menu-item.component';
import { DocxMenuComponent } from './menu.component';

const PARTS = [
  DocxEditorRootDirective,
  DocxEditorViewportComponent,
  DocxEditorContentComponent,
  DocxEditorMenuComponent,
  DocxMenuComponent,
  DocxMenuItemComponent,
  DocxSlotDirective,
];

@Directive()
abstract class Host {
  readonly bytes = signal<Uint8Array>(docx(paragraph('hello world')));
  readonly root = viewChild.required(DocxEditorRootDirective);

  editor(): DocxEditorInstance {
    return engineOf(this.root().context.editor());
  }
}

@Component({
  imports: PARTS,
  changeDetection: ChangeDetectionStrategy.OnPush,
  template: `<ng-container docxEditorRoot [document]="bytes()">
    <div
      docxEditorMenu
      [overflow]="false"
      [saveHandler]="save"
      [pageSetupHandler]="pageSetup"
    ></div>
    <docx-editor-viewport><docx-editor-content /></docx-editor-viewport>
  </ng-container>`,
})
class DefaultHostComponent extends Host {
  saves = 0;
  pageSetups: (HTMLElement | null)[] = [];
  readonly save = () => (this.saves += 1);
  readonly pageSetup = (trigger: HTMLElement | null) => this.pageSetups.push(trigger);
}

@Component({
  imports: PARTS,
  changeDetection: ChangeDetectionStrategy.OnPush,
  template: `<ng-container docxEditorRoot [document]="bytes()">
    <div docxEditorMenu [overflow]="false">
      <docx-menu menu="file">
        <ng-template docxSlot="file.exportPdf" docxSlotHidden />
        <ng-template docxSlot="file.exportMarkdown">
          <docx-menu-item slotId="file.exportMarkdown" className="custom-export" />
        </ng-template>
      </docx-menu>
      <docx-menu menu="review" hidden />
      <docx-menu menu="tools" label="Tools">
        <docx-menu-item slotId="format.clear" />
      </docx-menu>
      <span class="host-badge">Draft</span>
    </div>
    <docx-editor-viewport><docx-editor-content /></docx-editor-viewport>
  </ng-container>`,
})
class OverrideHostComponent extends Host {}

@Component({
  imports: PARTS,
  changeDetection: ChangeDetectionStrategy.OnPush,
  template: `<ng-container docxEditorRoot [document]="bytes()">
    <div
      docxEditorMenu
      [preset]="false"
      [overflow]="false"
      [pageSetupHandler]="noop"
      [paragraphDialogHandler]="noop"
    >
      <docx-menu-item slotId="file.pageSetup" />
      <docx-menu-item slotId="paragraph.dialog" />
      <docx-menu menu="file" />
    </div>
    <docx-editor-viewport><docx-editor-content /></docx-editor-viewport>
  </ng-container>`,
})
class DirectHostComponent extends Host {
  readonly noop = () => {};
}

@Component({
  imports: PARTS,
  changeDetection: ChangeDetectionStrategy.OnPush,
  template: `<ng-container docxEditorRoot [document]="bytes()">
    <div docxEditorMenu [reportIssue]="false"></div>
    <docx-editor-viewport><docx-editor-content /></docx-editor-viewport>
  </ng-container>`,
})
class MeasuredHostComponent extends Host {}

const barOf = (root: Element) => root.querySelector<HTMLElement>('[data-testid="docx-menubar"]')!;
const triggerOf = (root: Element, id: string) =>
  root.querySelector<HTMLButtonElement>(`[data-menu="${id}"] .docx-menubar__trigger`)!;
const triggers = (root: Element) => [
  ...root.querySelectorAll<HTMLButtonElement>('.docx-menubar__trigger'),
];

async function key(target: Element, name: string): Promise<KeyboardEvent> {
  const event = new KeyboardEvent('keydown', { key: name, bubbles: true, cancelable: true });
  target.dispatchEvent(event);
  await settle(undefined, 2);
  return event;
}

/** A pointerdown on the pages that the pages cancel, as the surface does. */
function pressCancelledPages(container: Element): MouseEvent {
  const target = container.ownerDocument.createElement('span');
  container.querySelector('.docx-pages')!.append(target);
  target.addEventListener('pointerdown', (event) => {
    event.preventDefault();
    event.stopPropagation();
  });
  // jsdom has no PointerEvent, so a MouseEvent carries the pointerdown type.
  const event = new MouseEvent('pointerdown', { bubbles: true, cancelable: true });
  target.dispatchEvent(event);
  target.remove();
  return event;
}

async function openSubmenu(root: Element): Promise<void> {
  root.querySelector('.docx-menubar__submenu')!.dispatchEvent(new MouseEvent('mouseenter'));
  await settle(undefined, 2);
}

describe('menu bar composition', () => {
  beforeEach(() => configureZoneless());
  afterEach(() => restoreResizeObserver());

  test('renders the registry menus on the default bar, in registry order', async () => {
    const fixture = await render(DefaultHostComponent);
    const bar = barOf(fixture.nativeElement);
    expect(bar.getAttribute('role')).toBe('menubar');
    expect(bar.getAttribute('aria-label')).toBeTruthy();
    const ids = [...bar.querySelectorAll('[data-menu]')].map((menu) =>
      menu.getAttribute('data-menu')
    );
    expect(ids).toEqual(CHROME_MENUS.map((menu) => menu.id));
    // One tab stop for the bar: the first trigger.
    expect(triggers(bar).map((trigger) => trigger.tabIndex)).toEqual([0, -1, -1, -1, -1]);
    fixture.destroy();
  });

  test('a press on the pages closes the open menu, although the pages cancel it', async () => {
    const fixture = await render(DefaultHostComponent);
    const root: HTMLElement = fixture.nativeElement;
    document.body.append(root);
    triggers(root)[0]!.click();
    await settle(fixture, 2);
    expect(root.querySelectorAll('[role="menu"]')).toHaveLength(1);
    const event = pressCancelledPages(root);
    await settle(fixture, 2);
    expect(event.defaultPrevented).toBe(true);
    expect(root.querySelectorAll('[role="menu"]')).toHaveLength(0);
    fixture.destroy();
    root.remove();
  });

  test('saveHandler runs when File › Save is chosen, and the menu closes', async () => {
    const fixture = await render(DefaultHostComponent);
    const root: HTMLElement = fixture.nativeElement;
    const trigger = triggerOf(root, 'file');
    trigger.focus();
    await key(trigger, 'Enter');
    const save = root.querySelector<HTMLButtonElement>('[data-slot="file.save"]')!;
    expect(save.getAttribute('aria-disabled')).toBeNull();
    save.dispatchEvent(new MouseEvent('click', { bubbles: true }));
    await settle(fixture, 2);
    expect(fixture.componentInstance.saves).toBe(1);
    expect(root.querySelector('[role="menu"]')).toBeNull();
    fixture.destroy();
  });

  test('File › Page setup runs pageSetupHandler with the File trigger', async () => {
    const fixture = await render(DefaultHostComponent);
    const root: HTMLElement = fixture.nativeElement;
    const trigger = triggerOf(root, 'file');
    await key(trigger, 'Enter');
    root.querySelector<HTMLButtonElement>('[data-slot="file.pageSetup"]')!.click();
    await settle(fixture, 2);
    expect(fixture.componentInstance.pageSetups).toEqual([trigger]);
    expect(root.querySelector('[role="menu"]')).toBeNull();
    fixture.destroy();
  });

  test('Enter opens a menu with the focus on its first row, and Escape gives it back', async () => {
    const fixture = await render(DefaultHostComponent);
    const root: HTMLElement = fixture.nativeElement;
    document.body.append(root);
    const trigger = triggerOf(root, 'format');
    trigger.focus();
    await key(trigger, 'Enter');
    const panel = root.querySelector<HTMLElement>('[role="menu"]')!;
    expect(trigger.getAttribute('aria-expanded')).toBe('true');
    expect(document.activeElement?.getAttribute('data-slot')).toBe('text.bold');
    await key(document.activeElement!, 'ArrowDown');
    expect(document.activeElement?.getAttribute('data-slot')).toBe('text.italic');
    await key(panel, 'End');
    expect(document.activeElement?.getAttribute('data-slot')).toBe('format.clear');
    await key(document.activeElement!, 'Escape');
    expect(root.querySelector('[role="menu"]')).toBeNull();
    expect(document.activeElement).toBe(trigger);
    fixture.destroy();
    root.remove();
  });

  test('rows follow the engine: toggles carry aria-checked and refusals carry the reason', async () => {
    const fixture = await render(DefaultHostComponent);
    const root: HTMLElement = fixture.nativeElement;
    await key(triggerOf(root, 'format'), 'Enter');
    const bold = root.querySelector<HTMLElement>('[data-slot="text.bold"]')!;
    expect(bold.getAttribute('role')).toBe('menuitemcheckbox');
    expect(bold.getAttribute('aria-checked')).toBe('false');
    const left = root.querySelector<HTMLElement>('[data-slot="alignment.left"]')!;
    expect(left.getAttribute('role')).toBe('menuitemradio');
    fixture.componentInstance.editor().exec({ type: 'setEditingMode', mode: 'viewing' });
    await settle(fixture);
    const refused = root.querySelector<HTMLElement>('[data-slot="text.bold"]')!;
    expect(refused.getAttribute('aria-disabled')).toBe('true');
    expect(refused.tabIndex).toBe(-1);
    const reason = refused.getAttribute('title');
    expect(reason).toBeTruthy();
    const described = refused.getAttribute('aria-describedby')!;
    expect(root.querySelector(`#${described}`)?.textContent).toBe(reason);
    fixture.destroy();
  });

  test('Insert › Table opens the size grid, and a pick inserts the table and closes the bar', async () => {
    const fixture = await render(DefaultHostComponent);
    const root: HTMLElement = fixture.nativeElement;
    triggerOf(root, 'insert').click();
    await settle(fixture, 2);
    const table = root.querySelector<HTMLElement>('docx-menu-table-picker .docx-menubar__submenu')!;
    expect(table).not.toBeNull();
    table.dispatchEvent(new MouseEvent('mouseenter'));
    await settle(fixture, 2);
    const cell = root.querySelector<HTMLButtonElement>('[data-cell="2x3"]')!;
    expect(cell).not.toBeNull();
    cell.click();
    await settle(fixture);
    expect(root.querySelector('[role="menu"]')).toBeNull();
    expect(fixture.componentInstance.editor().snapshot().table).not.toBeNull();
    fixture.destroy();
  });

  test('the Review menu ends with Markup options and its Reviewers row', async () => {
    const fixture = await render(DefaultHostComponent);
    const root: HTMLElement = fixture.nativeElement;
    triggerOf(root, 'review').click();
    await settle(fixture, 2);
    const markup = root.querySelector<HTMLElement>('.docx-menubar__markup-options')!;
    expect(markup).not.toBeNull();
    markup.dispatchEvent(new MouseEvent('mouseenter'));
    await settle(fixture, 2);
    // The document has no review authors, so the row is disabled with the engine's reason.
    const reviewers = markup.querySelector<HTMLElement>('[data-slot="review.authors"]')!;
    expect(reviewers.getAttribute('aria-disabled')).toBe('true');
    fixture.destroy();
  });

  test('a row override reaches the export submenu without a duplicate row', async () => {
    const fixture = await render(OverrideHostComponent);
    const root: HTMLElement = fixture.nativeElement;
    triggerOf(root, 'file').click();
    await settle(fixture, 2);
    await openSubmenu(root);
    expect(root.querySelectorAll('[data-slot="file.exportPdf"]')).toHaveLength(0);
    const rows = root.querySelectorAll('[data-slot="file.exportMarkdown"]');
    expect(rows).toHaveLength(1);
    expect(rows[0]!.classList.contains('custom-export')).toBe(true);
    expect(rows[0]!.closest('.docx-menubar__submenu')).not.toBeNull();
    fixture.destroy();
  });

  test('a host menu replaces, removes, or adds a menu, and host content stays at the end', async () => {
    const fixture = await render(OverrideHostComponent);
    const bar = barOf(fixture.nativeElement);
    const ids = [...bar.querySelectorAll('[data-menu]')].map((menu) =>
      menu.getAttribute('data-menu')
    );
    expect(ids).toEqual(['file', 'format', 'insert', 'help', 'tools']);
    const host = bar.querySelector('.docx-menubar__host')!;
    expect(host.querySelector('.host-badge')).not.toBeNull();
    expect(bar.lastElementChild?.tagName).toBe('INPUT');
    triggerOf(bar, 'tools').click();
    await settle(fixture, 2);
    const panel = bar.querySelector('[data-menu="tools"] [role="menu"]')!;
    expect(panel.querySelector('[data-slot="format.clear"]')).not.toBeNull();
    fixture.destroy();
  });

  test('direct bar actions stay tabbable and follow command state', async () => {
    const fixture = await render(DirectHostComponent);
    const root: HTMLElement = fixture.nativeElement;
    const page = root.querySelector<HTMLButtonElement>('[data-slot="file.pageSetup"]')!;
    const paragraphRow = root.querySelector<HTMLButtonElement>('[data-slot="paragraph.dialog"]')!;
    for (const action of [page, paragraphRow]) {
      expect(action.tabIndex).toBe(0);
      expect(action.getAttribute('aria-disabled')).not.toBe('true');
    }
    fixture.componentInstance.editor().exec({ type: 'setEditingMode', mode: 'viewing' });
    await settle(fixture);
    for (const action of [page, paragraphRow]) {
      expect(action.getAttribute('aria-disabled')).toBe('true');
    }
    fixture.componentInstance.editor().exec({ type: 'setEditingMode', mode: 'editing' });
    await settle(fixture);
    for (const action of [page, paragraphRow]) {
      expect(action.getAttribute('aria-disabled')).not.toBe('true');
    }
    triggerOf(root, 'file').click();
    await settle(fixture, 2);
    const popup = root.querySelector('[role="menu"]')!;
    expect(popup.querySelector<HTMLButtonElement>('[data-slot="file.pageSetup"]')!.tabIndex).toBe(
      -1
    );
    fixture.destroy();
  });

  test('right-to-left bar arrows follow the visual order and react to direction changes', async () => {
    const fixture = await render(DefaultHostComponent);
    const root: HTMLElement = fixture.nativeElement;
    document.body.append(root);
    const bar = barOf(root);
    const all = triggers(bar);
    bar.style.direction = 'rtl';
    all[0]!.focus();
    await key(all[0]!, 'ArrowLeft');
    expect(document.activeElement).toBe(all[1]);
    expect(all[1]!.tabIndex).toBe(0);
    await key(all[1]!, 'ArrowRight');
    await key(all[0]!, 'ArrowRight');
    expect(document.activeElement).toBe(all.at(-1));
    bar.style.direction = 'ltr';
    await key(all.at(-1)!, 'ArrowRight');
    expect(document.activeElement).toBe(all[0]);
    fixture.destroy();
    root.remove();
  });

  test('menus that do not fit move into "⋯" as submenus, from the end of the bar', async () => {
    useMockResizeObserver();
    const fixture = await render(MeasuredHostComponent);
    const root: HTMLElement = fixture.nativeElement;
    const bar = barOf(root);
    // Help has no row without Report issue, so it leaves the bar.
    expect(bar.querySelector('[data-menu="help"]')).toBeNull();
    expect(bar.hasAttribute('data-overflow')).toBe(true);
    await measure(bar, 200, 60);
    await settle(fixture, 2);
    const shown = [...bar.querySelectorAll('[data-menu]')].map((menu) =>
      menu.getAttribute('data-menu')
    );
    expect(shown.at(-1)).toBe('docx-menubar-more');
    expect(shown).not.toContain('review');
    const more = triggerOf(bar, 'docx-menubar-more');
    expect(more.getAttribute('aria-label')).toBeTruthy();
    more.click();
    await settle(fixture, 2);
    const labels = [
      ...bar.querySelectorAll('[data-menu="docx-menubar-more"] .docx-menubar__submenu'),
    ].map((submenu) => submenu.querySelector('.docx-menubar__item-label')?.textContent?.trim());
    expect(labels.at(-1)).toBe('Review');
    fixture.destroy();
  });
});

describe('menu bar while a document opens', () => {
  beforeEach(() => configureZoneless());

  test('the menu bar is disabled and closed while the next document opens', async () => {
    const fixture = await render(DefaultHostComponent);
    const root: HTMLElement = fixture.nativeElement;
    expect(triggers(root).length).toBeGreaterThan(0);
    triggers(root)[0]!.click();
    await settle(fixture, 1);
    expect(root.querySelector('[role="menu"]')).not.toBeNull();
    const editor = fixture.componentInstance.editor();
    void editor.load(largeDocx());
    expect(editor.snapshot().isOpening).toBe(true);
    await settle(fixture, 1);
    expect(triggers(root).every((trigger) => trigger.disabled)).toBe(true);
    expect(root.querySelector('[role="menu"]')).toBeNull();
    const deadline = Date.now() + 10_000;
    while (editor.snapshot().isOpening && Date.now() < deadline) await settle(fixture, 1);
    await settle(fixture);
    expect(triggers(root).every((trigger) => !trigger.disabled)).toBe(true);
    fixture.destroy();
  });
});
