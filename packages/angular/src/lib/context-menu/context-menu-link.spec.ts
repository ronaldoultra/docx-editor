// Task 5.2: the menu claims of packages/vue/test/context-menu-link.test.ts. The claims about
// the link popup itself port with the popup in task 6.2.
import { ChangeDetectionStrategy, Component, Directive, signal, viewChild } from '@angular/core';
import type { DocxEditorInstance } from '@docx-editor.dev/core/editor';
import { afterEach, beforeEach, expect, test } from 'vitest';
import { contextMenuPanel, rightClick, surfaceOf } from '../../testing/context-menu';
import { docx } from '../../testing/docx';
import { configureZoneless, engineOf, render, settle } from '../../testing/harness';
import { DocxEditorContentComponent } from '../content.component';
import { DocxMenuItemComponent } from '../menu/menu-item.component';
import { DocxEditorRootDirective } from '../root.directive';
import { DocxSlotDirective } from '../slot.directive';
import { DocxEditorViewportComponent } from '../viewport.component';
import { DocxEditorContextMenuComponent } from './context-menu.component';

const SOURCE = docx('<w:p><w:r><w:t>Visit example today</w:t></w:r></w:p>');
const LINKED = docx(
  '<w:p><w:r><w:t xml:space="preserve">Visit </w:t></w:r>' +
    '<w:hyperlink w:anchor="target"><w:r><w:t>example</w:t></w:r></w:hyperlink>' +
    '<w:r><w:t xml:space="preserve"> today</w:t></w:r></w:p>'
);
const PID = '/word/document.xml#0.0.0';
const t = (key: string) => key;
const PARTS = [
  DocxEditorRootDirective,
  DocxEditorViewportComponent,
  DocxEditorContentComponent,
  DocxEditorContextMenuComponent,
  DocxMenuItemComponent,
  DocxSlotDirective,
];

@Directive()
abstract class Host {
  readonly bytes = signal<Uint8Array>(SOURCE);
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
    <docx-editor-viewport>
      <docx-editor-content />
      <div docxEditorContextMenu [t]="t"></div>
    </docx-editor-viewport>
  </ng-container>`,
})
class DefaultHostComponent extends Host {}

@Component({
  imports: PARTS,
  changeDetection: ChangeDetectionStrategy.OnPush,
  template: `<ng-container docxEditorRoot [document]="bytes()">
    <docx-editor-viewport>
      <docx-editor-content />
      <div docxEditorContextMenu [t]="t">
        <ng-template docxSlot="text.link">
          <docx-menu-item slotId="text.link" className="host-link-row" />
        </ng-template>
      </div>
    </docx-editor-viewport>
  </ng-container>`,
})
class SlotHostComponent extends Host {}

async function mount<T extends Host>(component: new () => T, bytes = SOURCE) {
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

/** Register a link UI, as the link popup does. Returns the number of open requests. */
function linkUi(view: { host: Host }): { requests: number } {
  const counter = { requests: 0 };
  view.host.editor().setHyperlinkChrome({ onRequest: () => (counter.requests += 1) });
  return counter;
}

async function selectAndRightClick(
  view: { root: HTMLElement; host: Host },
  anchor = 6,
  head = 13
): Promise<void> {
  view.host.editor().surface!.setSelection({
    anchor: { paragraphId: PID, offset: anchor },
    head: { paragraphId: PID, offset: head },
  });
  await settle();
  await rightClick(surfaceOf(view.root));
}

const linkRows = (root: Element) =>
  root.querySelectorAll<HTMLElement>('[role="menu"] [data-slot="text.link"]');

beforeEach(() => configureZoneless());
afterEach(() => {
  document.body.replaceChildren();
});

test('the link row is enabled for a selection, and opens the link UI after the menu closes', async () => {
  const view = await mount(DefaultHostComponent);
  const ui = linkUi(view);
  await selectAndRightClick(view);
  const row = linkRows(view.root)[0]!;
  expect(row.getAttribute('aria-disabled')).toBeNull();
  expect(row.textContent).toContain('formattingBar.insertLink');
  expect(row.textContent).toContain('contextMenu.insertLinkShortcut');
  row.click();
  await settle();
  expect(contextMenuPanel(view.root)).toBeNull();
  expect(ui.requests).toBe(1);
  view.destroy();
});

test('without a link UI the link row is disabled with the engine reason', async () => {
  const view = await mount(DefaultHostComponent);
  await selectAndRightClick(view);
  const row = linkRows(view.root)[0]!;
  expect(row.getAttribute('aria-disabled')).toBe('true');
  expect(row.getAttribute('title')).toBeTruthy();
  view.destroy();
});

test('in viewing mode the link row is disabled, with the engine reason', async () => {
  const view = await mount(DefaultHostComponent);
  const ui = linkUi(view);
  view.host.editor().exec({ type: 'setEditingMode', mode: 'viewing' });
  await settle();
  await selectAndRightClick(view);
  const row = linkRows(view.root)[0]!;
  expect(row.getAttribute('aria-disabled')).toBe('true');
  expect(row.getAttribute('title')).toBeTruthy();
  row.click();
  await settle();
  expect(contextMenuPanel(view.root)).not.toBeNull();
  expect(ui.requests).toBe(0);
  view.destroy();
});

test('a text.link slot template replaces the packaged row with the same working row', async () => {
  const view = await mount(SlotHostComponent);
  const ui = linkUi(view);
  await selectAndRightClick(view);
  const rows = linkRows(view.root);
  expect(rows).toHaveLength(1);
  const row = rows[0]!;
  expect(row.classList.contains('host-link-row')).toBe(true);
  expect(row.getAttribute('aria-disabled')).toBeNull();
  expect(row.textContent).toContain('formattingBar.insertLink');
  row.click();
  await settle();
  expect(contextMenuPanel(view.root)).toBeNull();
  expect(ui.requests).toBe(1);
  view.destroy();
});

test('on an existing link the row says Edit link', async () => {
  const view = await mount(DefaultHostComponent, LINKED);
  await selectAndRightClick(view, 8, 8);
  const row = linkRows(view.root)[0]!;
  expect(row.textContent).toContain('hyperlinkPopup.editLink');
  expect(row.textContent).not.toContain('formattingBar.insertLink');
  view.destroy();
});
