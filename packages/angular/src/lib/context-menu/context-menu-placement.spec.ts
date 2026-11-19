// Task 5.2: the claims of packages/vue/test/context-menu-placement.test.ts. The panel gets a
// size from a stubbed `getBoundingClientRect`, and the window gets a small size.
import { ChangeDetectionStrategy, Component, signal } from '@angular/core';
import { afterEach, beforeEach, describe, expect, test } from 'vitest';
import { contextMenuPanel, rightClick, surfaceOf } from '../../testing/context-menu';
import { docx, paragraph } from '../../testing/docx';
import { configureZoneless, render, settle } from '../../testing/harness';
import { DocxEditorContentComponent } from '../content.component';
import { DocxMenuSubmenuComponent } from '../menu/menu-submenu.component';
import { DocxEditorRootDirective } from '../root.directive';
import { DocxEditorViewportComponent } from '../viewport.component';
import { DocxContextMenuItemComponent } from './context-menu-item.component';
import { DocxEditorContextMenuComponent } from './context-menu.component';

const VIEW = { width: 1000, height: 600 };
let panelSize = { width: 240, height: 300 };

const originalRect = HTMLElement.prototype.getBoundingClientRect;
const originalSize = { width: window.innerWidth, height: window.innerHeight };

function rectOf(left: number, top: number, width: number, height: number): DOMRect {
  return {
    x: left,
    y: top,
    left,
    top,
    width,
    height,
    right: left + width,
    bottom: top + height,
    toJSON: () => ({}),
  } as DOMRect;
}

const PARTS = [
  DocxEditorRootDirective,
  DocxEditorViewportComponent,
  DocxEditorContentComponent,
  DocxEditorContextMenuComponent,
  DocxContextMenuItemComponent,
  DocxMenuSubmenuComponent,
];

@Component({
  imports: PARTS,
  changeDetection: ChangeDetectionStrategy.OnPush,
  template: `<ng-container docxEditorRoot [document]="bytes()">
    <docx-editor-viewport>
      <docx-editor-content />
      <div docxEditorContextMenu></div>
    </docx-editor-viewport>
  </ng-container>`,
})
class PlainHostComponent {
  readonly bytes = signal(docx(paragraph('hello world')));
}

@Component({
  imports: PARTS,
  changeDetection: ChangeDetectionStrategy.OnPush,
  template: `<ng-container docxEditorRoot [document]="bytes()">
    <docx-editor-viewport>
      <docx-editor-content />
      <div docxEditorContextMenu>
        <docx-menu-submenu label="More"><docx-context-menu-item label="Inner" /></docx-menu-submenu>
        <docx-context-menu-item label="Edit" (activate)="edit()" />
      </div>
    </docx-editor-viewport>
  </ng-container>`,
})
class RowsHostComponent {
  readonly bytes = signal(docx(paragraph('hello world')));
  readonly field = document.createElement('input');
  edit(): void {
    this.field.focus();
  }
}

async function open(
  component: typeof PlainHostComponent | typeof RowsHostComponent,
  x: number,
  y: number
): Promise<{ root: HTMLElement; panel: HTMLElement; host: unknown }> {
  const fixture = await render<unknown>(component);
  const root: HTMLElement = fixture.nativeElement;
  document.body.append(root);
  await rightClick(surfaceOf(root), { x, y });
  await settle(fixture, 2);
  const panel = contextMenuPanel(root);
  if (!panel) throw new Error('the menu did not open');
  return { root, panel, host: fixture.componentInstance };
}

function placed(panel: HTMLElement): { top: number; bottom: number; maxHeight: number } {
  const top = Number.parseFloat(panel.style.top);
  const maxHeight = Number.parseFloat(panel.style.maxHeight);
  return { top, bottom: top + Math.min(panelSize.height, maxHeight), maxHeight };
}

beforeEach(() => {
  configureZoneless();
  panelSize = { width: 240, height: 300 };
  Object.defineProperty(window, 'innerWidth', { value: VIEW.width, configurable: true });
  Object.defineProperty(window, 'innerHeight', { value: VIEW.height, configurable: true });
  HTMLElement.prototype.getBoundingClientRect = function (this: HTMLElement): DOMRect {
    if (this.classList.contains('docx-contextmenu')) {
      const left = Number.parseFloat(this.style.left) || 0;
      const top = Number.parseFloat(this.style.top) || 0;
      const cap = Number.parseFloat(this.style.maxHeight) || Infinity;
      return rectOf(left, top, panelSize.width, Math.min(panelSize.height, cap));
    }
    return originalRect.call(this);
  };
});

afterEach(() => {
  document.body.replaceChildren();
  HTMLElement.prototype.getBoundingClientRect = originalRect;
  Object.defineProperty(window, 'innerWidth', { value: originalSize.width, configurable: true });
  Object.defineProperty(window, 'innerHeight', { value: originalSize.height, configurable: true });
  delete (HTMLElement.prototype as Partial<HTMLElement>).showPopover;
});

describe('the menu stays inside the window', () => {
  test('a press near the bottom keeps a tall menu on screen, capped and scrolling', async () => {
    panelSize = { width: 240, height: 900 };
    const box = placed((await open(PlainHostComponent, 120, 560)).panel);
    expect(box.maxHeight).toBe(VIEW.height - 16);
    expect(box.top).toBeGreaterThanOrEqual(8);
    expect(box.bottom).toBeLessThanOrEqual(VIEW.height - 8);
  });

  test('a press near the top keeps a tall menu on screen, capped and scrolling', async () => {
    panelSize = { width: 240, height: 900 };
    const box = placed((await open(PlainHostComponent, 120, 40)).panel);
    expect(box.maxHeight).toBe(VIEW.height - 16);
    expect(box.top).toBeGreaterThanOrEqual(8);
    expect(box.bottom).toBeLessThanOrEqual(VIEW.height - 8);
  });

  test('a menu that fits on neither side of the press moves up only as far as it must', async () => {
    panelSize = { width: 240, height: 400 };
    const box = placed((await open(PlainHostComponent, 120, 300)).panel);
    expect(box.top).toBe(VIEW.height - 8 - 400);
    expect(box.bottom).toBe(VIEW.height - 8);
  });
});

describe('a menu taller than the window scrolls', () => {
  test('the panel takes the focus once it is visible, so the arrow keys work', async () => {
    // A hidden element cannot take the focus in a browser, so the focus moves only after the
    // render that makes the panel visible.
    const seen: string[] = [];
    const onFocus = (event: FocusEvent) => {
      const target = event.target as HTMLElement;
      if (target.classList?.contains('docx-contextmenu')) seen.push(target.style.visibility);
    };
    document.addEventListener('focusin', onFocus, true);
    const { panel } = await open(PlainHostComponent, 120, 100);
    document.removeEventListener('focusin', onFocus, true);
    expect(document.activeElement).toBe(panel);
    expect(seen).toEqual(['visible']);
  });

  test('the panel scrolling inside itself does not close it', async () => {
    panelSize = { width: 240, height: 900 };
    const { root, panel } = await open(PlainHostComponent, 120, 300);
    panel.dispatchEvent(new Event('scroll'));
    await settle(undefined, 2);
    expect(contextMenuPanel(root)).not.toBeNull();
  });

  test('a scroll of the document closes it', async () => {
    const { root } = await open(PlainHostComponent, 120, 300);
    document.dispatchEvent(new Event('scroll'));
    await settle(undefined, 2);
    expect(contextMenuPanel(root)).toBeNull();
  });

  test('the arrow keys scroll the focused row into view', async () => {
    panelSize = { width: 240, height: 900 };
    const { panel } = await open(PlainHostComponent, 120, 300);
    Object.defineProperty(panel, 'clientTop', { value: 0, configurable: true });
    Object.defineProperty(panel, 'clientHeight', { value: VIEW.height - 16, configurable: true });
    const items = [...panel.querySelectorAll<HTMLElement>('[role^="menuitem"]')].filter(
      (item) => item.closest('[role="menu"]') === panel
    );
    items.forEach((item, index) => {
      item.getBoundingClientRect = () => rectOf(0, 8 + index * 100 - panel.scrollTop, 240, 100);
    });
    const last = items.at(-1)!;
    expect(last.getBoundingClientRect().bottom).toBeGreaterThan(VIEW.height);

    panel.dispatchEvent(new KeyboardEvent('keydown', { key: 'End', bubbles: true }));
    await settle(undefined, 1);
    expect(document.activeElement).toBe(last);
    const lastBox = last.getBoundingClientRect();
    expect(lastBox.bottom).toBeLessThanOrEqual(VIEW.height - 8);
    expect(lastBox.top).toBeGreaterThanOrEqual(8);

    panel.dispatchEvent(new KeyboardEvent('keydown', { key: 'ArrowDown', bubbles: true }));
    await settle(undefined, 1);
    expect(document.activeElement).toBe(items[0]!);
    expect(panel.scrollTop).toBe(0);
  });
});

describe('stacking', () => {
  test('the panel opens in the top layer, above any host stacking context', async () => {
    const shown: HTMLElement[] = [];
    (HTMLElement.prototype as { showPopover?: () => void }).showPopover = function (
      this: HTMLElement
    ) {
      shown.push(this);
    };
    const { panel } = await open(PlainHostComponent, 120, 100);
    expect(panel.getAttribute('popover')).toBe('manual');
    expect(shown).toEqual([panel]);
  });
});

describe('a narrow window', () => {
  test('caps the menu width so that its right edge stays in the window', async () => {
    Object.defineProperty(window, 'innerWidth', { value: 200, configurable: true });
    const { panel } = await open(PlainHostComponent, 150, 100);
    expect(Number.parseFloat(panel.style.left)).toBe(8);
    expect(Number.parseFloat(panel.style.maxWidth)).toBe(200 - 16);
  });
});

describe('a submenu in a scrolling menu', () => {
  test('follows its row when the menu scrolls', async () => {
    panelSize = { width: 240, height: 900 };
    const { panel } = await open(RowsHostComponent, 120, 300);
    const wrapper = panel.querySelector<HTMLElement>('.docx-menubar__submenu')!;
    const row = wrapper.querySelector<HTMLElement>('[aria-haspopup="menu"]')!;
    row.getBoundingClientRect = () => rectOf(0, 400 - panel.scrollTop, 240, 32);
    wrapper.dispatchEvent(new MouseEvent('mouseenter'));
    await settle(undefined, 2);
    const submenu = () => panel.querySelector<HTMLElement>('.docx-menubar__submenu-panel')!;
    const startTop = Number.parseFloat(submenu().style.top);

    panel.scrollTop = 100;
    panel.dispatchEvent(new Event('scroll'));
    await settle(undefined, 2);
    expect(Number.parseFloat(submenu().style.top)).toBe(startTop - 100);
  });
});

describe('a host row that opens a dialog', () => {
  test('keeps the focus that the dialog takes', async () => {
    const { root, panel, host } = await open(RowsHostComponent, 120, 100);
    const field = (host as RowsHostComponent).field;
    document.body.append(field);
    const row = [...panel.querySelectorAll<HTMLElement>('[role="menuitem"]')].find(
      (item) => item.textContent?.trim() === 'Edit'
    )!;
    row.click();
    await settle(undefined, 2);
    expect(contextMenuPanel(root)).toBeNull();
    expect(document.activeElement).toBe(field);
  });
});
