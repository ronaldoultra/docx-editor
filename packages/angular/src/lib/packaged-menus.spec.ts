// Task 5.4: the `menu`, `contextMenu`, and `title` inputs of <docx-editor>, with the menu,
// title, context menu, and chrome band claims of packages/vue/test/sugar-host.test.ts and the
// title claim of live-props.test.ts.
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

@Component({
  selector: 'packaged-title-host',
  imports: [DocxEditorComponent],
  changeDetection: ChangeDetectionStrategy.OnPush,
  template: `<docx-editor
    [document]="bytes"
    [(title)]="title"
    [menu]="{ saveHandler: save, openHandler: open }"
  />`,
})
class TitleHostComponent {
  readonly bytes = SOURCE;
  readonly title = signal<string | undefined>('First title');
  readonly calls: string[] = [];
  readonly save = () => this.calls.push('save');
  readonly open = () => this.calls.push('open');
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

const titleInput = (root: Element) =>
  root.querySelector<HTMLInputElement>('[data-testid="docx-editor-title"]');

async function chooseFileRow(root: Element, slot: string): Promise<void> {
  root.querySelector<HTMLButtonElement>('[data-menu="file"] .docx-menubar__trigger')!.click();
  await settle(undefined, 2);
  const row = root.querySelector<HTMLElement>(`[data-slot="${slot}"]`)!;
  row.dispatchEvent(new MouseEvent('click', { bubbles: true }));
  await settle(undefined, 2);
}

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

  test('the title bar, the menu bar, and the toolbar share one chrome band', async () => {
    const { root, destroy } = await mount(HostComponent);
    const band = root.querySelector<HTMLElement>('[data-testid="docx-editor-chrome-band"]')!;
    expect(band.querySelector('[data-testid="docx-editor-title"]')).not.toBeNull();
    expect(band.querySelector('[data-testid="docx-menubar"]')).not.toBeNull();
    expect(band.querySelector('[data-testid="docx-toolbar"]')).not.toBeNull();
    // The band is a stacking context, so it must stack in the overlay band to keep its popups
    // above other chrome.
    expect(band.style.zIndex).toBe('var(--doc-z-overlay)');
    destroy();
  });

  test('chrome false renders no band and no title', async () => {
    const { fixture, root, destroy } = await mount(HostComponent);
    fixture.componentInstance.chrome.set(false);
    await settle(fixture);
    expect(root.querySelector('[data-testid="docx-editor-chrome-band"]')).toBeNull();
    expect(titleInput(root)).toBeNull();
    destroy();
  });

  test('without a title, the title bar shows the untitled placeholder', async () => {
    const { root, destroy } = await mount(HostComponent);
    const input = titleInput(root)!;
    expect(input.value).toBe('');
    expect(input.placeholder).toBe('Untitled');
    expect(input.getAttribute('aria-label')).toBe('Document name');
    destroy();
  });

  test('a new title replaces the old one, and an edit writes back through [(title)]', async () => {
    const { fixture, root, destroy } = await mount(TitleHostComponent);
    const input = titleInput(root)!;
    expect(input.value).toBe('First title');
    fixture.componentInstance.title.set('Second title');
    await settle(fixture);
    expect(input.value).toBe('Second title');

    input.value = 'Renamed';
    input.dispatchEvent(new Event('input', { bubbles: true }));
    await settle(fixture);
    expect(fixture.componentInstance.title()).toBe('Renamed');
    destroy();
  });

  test('the menu handlers run for File › Save and File › Open', async () => {
    const { fixture, root, destroy } = await mount(TitleHostComponent);
    await chooseFileRow(root, 'file.save');
    await chooseFileRow(root, 'file.open');
    expect(fixture.componentInstance.calls).toEqual(['save', 'open']);
    destroy();
  });
});
