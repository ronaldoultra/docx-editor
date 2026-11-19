// Task 5.4: the `menu` and `contextMenu` inputs of <docx-editor>.
import { ChangeDetectionStrategy, Component, signal } from '@angular/core';
import { afterEach, beforeEach, describe, expect, test } from 'vitest';
import {
  contextMenuPanel,
  contextMenuSlots,
  rightClick,
  scrollerOf,
} from '../testing/context-menu';
import { docx, paragraph } from '../testing/docx';
import { configureZoneless, render, settle } from '../testing/harness';
import {
  DocxEditorComponent,
  type DocxEditorContextMenuOptions,
  type DocxEditorMenuOptions,
} from './docx-editor.component';

const SOURCE = docx(paragraph('hello world'));

@Component({
  selector: 'packaged-menus-host',
  imports: [DocxEditorComponent],
  changeDetection: ChangeDetectionStrategy.OnPush,
  template: `<docx-editor
    [document]="bytes"
    [chrome]="chrome()"
    [menu]="menu()"
    [contextMenu]="contextMenu()"
  />`,
})
class HostComponent {
  readonly bytes = SOURCE;
  readonly chrome = signal(true);
  readonly menu = signal<boolean | DocxEditorMenuOptions>(true);
  readonly contextMenu = signal<boolean | DocxEditorContextMenuOptions>(true);
}

@Component({
  selector: 'packaged-menus-attribute-host',
  imports: [DocxEditorComponent],
  changeDetection: ChangeDetectionStrategy.OnPush,
  template: `<docx-editor [document]="bytes" menu="false" contextMenu="false" />`,
})
class AttributeHostComponent {
  readonly bytes = SOURCE;
}

async function mount<T>(component: new () => T) {
  const fixture = await render(component);
  const root: HTMLElement = fixture.nativeElement;
  document.body.append(root);
  await settle(fixture);
  return {
    fixture,
    root,
    destroy: () => {
      fixture.destroy();
      root.remove();
    },
  };
}

const menuBar = (root: Element) => root.querySelector<HTMLElement>('[data-testid="docx-menubar"]');
const menuIds = (root: Element) =>
  [...(menuBar(root)?.querySelectorAll<HTMLElement>('[data-menu]') ?? [])].map(
    (menu) => menu.dataset['menu']
  );

describe('<docx-editor> menu and contextMenu inputs', () => {
  beforeEach(() => {
    configureZoneless();
  });
  afterEach(() => {
    document.body.replaceChildren();
  });

  test('shows the menu bar above the toolbar and opens the context menu by default', async () => {
    const { root, destroy } = await mount(HostComponent);
    const bar = menuBar(root);
    expect(bar).not.toBeNull();
    const toolbar = root.querySelector('[data-testid="docx-toolbar"]')!;
    expect(bar!.compareDocumentPosition(toolbar) & Node.DOCUMENT_POSITION_FOLLOWING).toBeTruthy();
    expect(menuIds(root)).toEqual(expect.arrayContaining(['file', 'format', 'insert', 'help']));

    const event = await rightClick(scrollerOf(root));
    expect(event.defaultPrevented).toBe(true);
    expect(contextMenuSlots(root)).toContain('edit.copy');
    destroy();
  });

  test('false removes each part, and true shows it again', async () => {
    const { fixture, root, destroy } = await mount(HostComponent);
    fixture.componentInstance.menu.set(false);
    fixture.componentInstance.contextMenu.set(false);
    await settle(fixture);
    expect(menuBar(root)).toBeNull();
    const event = await rightClick(scrollerOf(root));
    expect(event.defaultPrevented).toBe(false);
    expect(contextMenuPanel(root)).toBeNull();

    fixture.componentInstance.menu.set(true);
    fixture.componentInstance.contextMenu.set(true);
    await settle(fixture);
    expect(menuBar(root)).not.toBeNull();
    await rightClick(scrollerOf(root));
    expect(contextMenuPanel(root)).not.toBeNull();
    destroy();
  });

  test('a string attribute "false" removes each part', async () => {
    const { root, destroy } = await mount(AttributeHostComponent);
    expect(menuBar(root)).toBeNull();
    const event = await rightClick(scrollerOf(root));
    expect(event.defaultPrevented).toBe(false);
    destroy();
  });

  test('chrome false hides the menu bar and keeps the context menu', async () => {
    const { fixture, root, destroy } = await mount(HostComponent);
    fixture.componentInstance.chrome.set(false);
    await settle(fixture);
    expect(menuBar(root)).toBeNull();
    await rightClick(scrollerOf(root));
    expect(contextMenuPanel(root)).not.toBeNull();
    destroy();
  });

  test('an options object reaches the menu bar and the context menu', async () => {
    const { fixture, root, destroy } = await mount(HostComponent);
    const opens: boolean[] = [];
    fixture.componentInstance.menu.set({ preset: false });
    fixture.componentInstance.contextMenu.set({
      className: 'host-menu',
      onOpenChange: (open) => opens.push(open),
    });
    await settle(fixture);
    expect(menuBar(root)).not.toBeNull();
    expect(menuIds(root)).toEqual([]);

    await rightClick(scrollerOf(root));
    expect(contextMenuPanel(root)?.classList.contains('host-menu')).toBe(true);
    expect(opens).toEqual([true]);

    fixture.componentInstance.contextMenu.set({ disabled: true });
    await settle(fixture);
    expect(contextMenuPanel(root)).toBeNull();
    const event = await rightClick(scrollerOf(root));
    expect(event.defaultPrevented).toBe(false);
    destroy();
  });
});
