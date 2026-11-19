// Task 5.3: the claims of packages/vue/test/context-menu-keyboard.test.ts. A context menu that
// the keyboard opens (Shift+F10 or the Menu key) opens below the caret. The keyboard event
// reports no pointer position, so the menu reads the caret rectangle from the editor. Rows in
// `docx-context-menu-start` learn how the menu was opened.
import {
  ChangeDetectionStrategy,
  Component,
  computed,
  effect,
  signal,
  viewChild,
  type Signal,
} from '@angular/core';
import type { DocxEditorInstance } from '@docx-editor.dev/core/editor';
import { afterEach, beforeEach, expect, test } from 'vitest';
import { contextMenuPanel, rightClick, surfaceOf } from '../../testing/context-menu';
import { docx, paragraph } from '../../testing/docx';
import { configureZoneless, engineOf, render, settle } from '../../testing/harness';
import { DocxEditorContentComponent } from '../content.component';
import { DocxEditorRootDirective } from '../root.directive';
import { DocxEditorViewportComponent } from '../viewport.component';
import { DocxContextMenuStartComponent } from './context-menu-item.component';
import { injectContextMenuAnchor, injectContextMenuTarget } from './context-menu-state';
import type { ContextMenuAnchor } from './context-menu-target';
import { DocxEditorContextMenuComponent } from './context-menu.component';

const t = (key: string): string => key;
const SOURCE = docx(paragraph('hello world, typed'));

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

/** What the probe row saw at each open: how the menu opened, and on what. */
const seen: { trigger: string | null; target: HTMLElement | null }[] = [];
/** How many times the probe row was created. */
let mounts = 0;

@Component({
  selector: 'docx-spec-opened-by',
  changeDetection: ChangeDetectionStrategy.OnPush,
  // Rendered, so a test can see the row follow a reopen without a new mount.
  template: `<span data-opened-by>{{ text() }}</span>`,
})
class OpenedByComponent {
  private readonly anchor: Signal<ContextMenuAnchor | null> = injectContextMenuAnchor();
  private readonly target: Signal<HTMLElement | null> = injectContextMenuTarget();
  readonly text = computed(() => this.anchor()?.trigger ?? 'closed');

  constructor() {
    mounts += 1;
    // Angular creates projected rows once, with the menu. The row reads each open from its
    // signals, as a Vue row reads it when it mounts.
    effect(() => {
      const anchor = this.anchor();
      if (anchor) seen.push({ trigger: anchor.trigger, target: this.target() });
    });
  }
}

@Component({
  selector: 'docx-spec-keyboard-host',
  imports: [
    DocxEditorRootDirective,
    DocxEditorViewportComponent,
    DocxEditorContentComponent,
    DocxEditorContextMenuComponent,
    DocxContextMenuStartComponent,
    OpenedByComponent,
  ],
  changeDetection: ChangeDetectionStrategy.OnPush,
  template: `<ng-container docxEditorRoot [document]="bytes()">
    <docx-editor-viewport>
      <docx-editor-content />
      <div docxEditorContextMenu [t]="t">
        <docx-context-menu-start><docx-spec-opened-by /></docx-context-menu-start>
      </div>
    </docx-editor-viewport>
  </ng-container>`,
})
class KeyboardHostComponent {
  readonly bytes = signal(SOURCE);
  readonly t = t;
  readonly root = viewChild.required(DocxEditorRootDirective);

  editor(): DocxEditorInstance {
    return engineOf(this.root().context.editor());
  }
}

const KEYBOARD = true;

async function mount(): Promise<{ root: HTMLElement; editor: () => DocxEditorInstance }> {
  seen.length = 0;
  mounts = 0;
  const fixture = await render(KeyboardHostComponent);
  const root: HTMLElement = fixture.nativeElement;
  document.body.append(root);
  await settle(fixture);
  return { root, editor: () => fixture.componentInstance.editor() };
}

async function open(
  root: HTMLElement,
  keyboard: boolean,
  point = { x: 420, y: 260 }
): Promise<HTMLElement> {
  await rightClick(surfaceOf(root), point, keyboard);
  const panel = contextMenuPanel(root);
  if (!panel) throw new Error('the menu did not open');
  return panel;
}

async function placeCaret(editor: DocxEditorInstance, offset: number): Promise<void> {
  const paragraphId = editor.surface!.session.paragraphIds()[0]!;
  const at = { blockId: paragraphId, offset };
  expect(editor.exec({ type: 'setSelection', range: { anchor: at, head: at } }).ok).toBe(true);
  await settle(undefined, 2);
}

const left = (panel: HTMLElement) => Number.parseFloat(panel.style.left);
const top = (panel: HTMLElement) => Number.parseFloat(panel.style.top);

beforeEach(() => {
  configureZoneless();
  Object.defineProperty(window, 'innerWidth', { value: 1200, configurable: true });
  Object.defineProperty(window, 'innerHeight', { value: 900, configurable: true });
  HTMLElement.prototype.getBoundingClientRect = function (this: HTMLElement): DOMRect {
    if (this.classList.contains('docx-contextmenu')) {
      return rectOf(left(this) || 0, top(this) || 0, 200, 240);
    }
    // The page content box that holds the caret, as if the page were scrolled into view.
    if (this.classList.contains('docx-page-content')) return rectOf(300, 120, 600, 800);
    if (this.classList.contains('docx-editor__scroll-container')) return rectOf(0, 0, 1200, 900);
    return originalRect.call(this);
  };
});

afterEach(() => {
  document.body.replaceChildren();
  HTMLElement.prototype.getBoundingClientRect = originalRect;
  Object.defineProperty(window, 'innerWidth', { value: originalSize.width, configurable: true });
  Object.defineProperty(window, 'innerHeight', { value: originalSize.height, configurable: true });
});

test('a keyboard-opened menu opens below the caret', async () => {
  const { root, editor } = await mount();
  await placeCaret(editor(), 5);
  const caret = editor().getCaretRect();
  expect(caret).not.toBeNull();
  expect(caret!.left).toBeGreaterThan(300);
  const panel = await open(root, KEYBOARD);
  expect(left(panel)).toBeCloseTo(caret!.left, 5);
  expect(top(panel)).toBeCloseTo(caret!.bottom, 5);
  expect(seen.at(-1)).toEqual({ trigger: 'keyboard', target: null });
});

test('a keyboard-opened menu follows the caret to another position', async () => {
  const { root, editor } = await mount();
  await placeCaret(editor(), 0);
  const start = editor().getCaretRect()!;
  await placeCaret(editor(), 11);
  const end = editor().getCaretRect()!;
  expect(end.left).toBeGreaterThan(start.left);
  const panel = await open(root, KEYBOARD);
  expect(left(panel)).toBeCloseTo(end.left, 5);
});

test('a right-click opens at the pointer and names its target', async () => {
  const { root } = await mount();
  const panel = await open(root, !KEYBOARD, { x: 420, y: 260 });
  expect(left(panel)).toBe(420);
  expect(top(panel)).toBe(260);
  expect(seen.at(-1)?.trigger).toBe('pointer');
  expect(seen.at(-1)?.target).not.toBeNull();
});

test('a row that stays mounted follows a keyboard reopen of an open menu', async () => {
  const { root, editor } = await mount();
  await placeCaret(editor(), 4);
  await open(root, !KEYBOARD, { x: 420, y: 260 });
  const row = () => root.querySelector('[data-opened-by]')!.textContent;
  expect(row()).toBe('pointer');
  const created = mounts;
  const panel = await open(root, KEYBOARD);
  expect(mounts).toBe(created);
  expect(seen.at(-1)).toEqual({ trigger: 'keyboard', target: null });
  expect(row()).toBe('keyboard');
  expect(left(panel)).toBeCloseTo(editor().getCaretRect()!.left, 5);
});

test('a keyboard-opened menu with no painted caret opens at the document corner', async () => {
  const { root, editor } = await mount();
  await placeCaret(editor(), 3);
  for (const page of root.querySelectorAll('.docx-page-content')) {
    page.classList.remove('docx-page-content');
  }
  expect(editor().getCaretRect()).toBeNull();
  const panel = await open(root, KEYBOARD);
  expect(left(panel)).toBe(16);
  expect(top(panel)).toBe(16);
  expect(seen.at(-1)?.trigger).toBe('keyboard');
});
