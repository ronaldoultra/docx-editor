// Task 4.3: the color, highlight, alignment, list, indent, and line spacing controls. The
// line-spacing and alignment-in-More claims of packages/vue/test/chrome-composition.test.ts, and
// the color claims of toolbar-picker-accessibility.test.ts.
import { ChangeDetectionStrategy, Component, Directive, signal, viewChild } from '@angular/core';
import type { DocxEditorInstance } from '@docx-editor.dev/core/editor';
import { en } from '@docx-editor.dev/core/i18n';
import { afterEach, beforeEach, describe, expect, test } from 'vitest';
import { docx, headerDocx, paragraph } from '../../testing/docx';
import { configureZoneless, engineOf, render, settle } from '../../testing/harness';
import {
  measure,
  panelOf,
  restoreResizeObserver,
  toolbarOf,
  useMockResizeObserver,
} from '../../testing/toolbar';
import { DocxEditorContentComponent } from '../content.component';
import { DocxEditorRootDirective } from '../root.directive';
import { DocxEditorViewportComponent } from '../viewport.component';
import { DocxEditorToolbarComponent } from './toolbar.component';

const SOURCE = docx(paragraph('hello world'));

@Directive()
abstract class Host {
  readonly bytes = signal<Uint8Array>(SOURCE);
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
    DocxEditorToolbarComponent,
  ],
  changeDetection: ChangeDetectionStrategy.OnPush,
  template: `<ng-container docxEditorRoot [document]="bytes()">
    <fieldset docxEditorToolbar></fieldset>
    <docx-editor-viewport><docx-editor-content /></docx-editor-viewport>
  </ng-container>`,
})
class ToolbarHostComponent extends Host {}

async function mount(bytes?: Uint8Array) {
  const fixture = await render(ToolbarHostComponent);
  if (bytes) {
    fixture.componentInstance.bytes.set(bytes);
    await settle(fixture);
  }
  const element: HTMLElement = fixture.nativeElement;
  document.body.append(element);
  // jsdom has no `contentEditable` property, so the pages cannot take the focus without a tab
  // stop.
  const pages = () => element.querySelector<HTMLElement>('.docx-pages')!;
  pages().tabIndex = -1;
  const update = async (action: () => unknown = () => {}) => {
    action();
    await settle(fixture, 2);
  };
  const editor = () => fixture.componentInstance.editor();
  await update(() => editor().surface!.selectAll());
  return { fixture, element, update, editor, pages };
}

const keydown = (target: EventTarget, key: string) => {
  const event = new KeyboardEvent('keydown', { key, bubbles: true, cancelable: true });
  target.dispatchEvent(event);
  return event;
};

const pointerDown = (target: EventTarget, cancelable = false): Event => {
  const event = new MouseEvent('pointerdown', { bubbles: true, cancelable });
  target.dispatchEvent(event);
  return event;
};

beforeEach(() => configureZoneless());
afterEach(() => {
  restoreResizeObserver();
  document.body.replaceChildren();
});

describe('line spacing', () => {
  test('the menu ticks the spacing in force, even one that no preset matches', async () => {
    const { element, update, editor } = await mount();
    const root = element.querySelector<HTMLElement>('[data-slot="list.lineSpacing"]')!;
    const rowsWhenOpen = async () => {
      await update(() => root.querySelector<HTMLButtonElement>('[aria-haspopup]')!.click());
      const rows = [...root.querySelectorAll<HTMLButtonElement>('[role="menuitemradio"]')];
      const texts = rows.map((row) => row.textContent!.replace(/\s+/g, ''));
      await update(() => rows.find((row) => row.hasAttribute('data-selected'))?.click());
      return texts;
    };
    // The fixture has no styles part, so the application defaults' multiple is in force; it
    // is no preset, so it joins the list in order.
    expect(await rowsWhenOpen()).toEqual(['1.0', '1.15', '1.16✓', '1.5', '2.0', '2.5', '3.0']);
    await update(() => editor().exec({ type: 'setLineSpacing', rule: 'atLeast', value: 14 }));
    expect(await rowsWhenOpen()).toEqual([
      'Atleast14pt✓',
      '1.0',
      '1.15',
      '1.5',
      '2.0',
      '2.5',
      '3.0',
    ]);
    // Picking the ticked row only closed the menu.
    expect(root.querySelector('[role="menu"]')).toBeNull();
    expect(editor().snapshot().formatting?.lineSpacing).toEqual({ rule: 'atLeast', value: 14 });
  });

  test('a preset applies, and the space rows add and remove the space', async () => {
    const { element, update, editor } = await mount();
    const root = element.querySelector<HTMLElement>('[data-slot="list.lineSpacing"]')!;
    const open = () =>
      update(() => root.querySelector<HTMLButtonElement>('[aria-haspopup]')!.click());
    await open();
    const double = [...root.querySelectorAll<HTMLButtonElement>('[role="menuitemradio"]')].find(
      (row) => row.textContent!.trim() === '2.0'
    )!;
    await update(() => double.click());
    expect(editor().snapshot().formatting?.lineSpacing).toEqual({ rule: 'multiple', value: 2 });
    await open();
    const addBefore = [...root.querySelectorAll<HTMLButtonElement>('[role="menuitem"]')].find(
      (row) => row.textContent!.trim() === en.lineSpacing.addSpaceBefore
    )!;
    await update(() => addBefore.click());
    expect(editor().snapshot().formatting?.spaceBeforePt).toBe(10);
    await open();
    const removeBefore = [...root.querySelectorAll<HTMLButtonElement>('[role="menuitem"]')].find(
      (row) => row.textContent!.trim() === en.lineSpacing.removeSpaceBefore
    )!;
    await update(() => removeBefore.click());
    expect(editor().snapshot().formatting?.spaceBeforePt ?? 0).toBe(0);
    // Without a paragraph dialog provider, the menu has no options row.
    await open();
    expect(root.querySelector('[data-slot="paragraph.dialog"]')).toBeNull();
  });
});

describe('alignment', () => {
  test('the trigger shows the alignment in force, and an option applies its alignment', async () => {
    const { element, update, editor } = await mount();
    const root = element.querySelector<HTMLElement>('[data-slot="alignment"]')!;
    const trigger = root.querySelector<HTMLButtonElement>('[aria-haspopup]')!;
    expect(trigger.disabled).toBe(false);
    expect(trigger.getAttribute('aria-label')).toBe(en.alignment.alignLeft);
    await update(() => trigger.click());
    const center = root.querySelector<HTMLButtonElement>('[data-slot="alignment.center"]')!;
    expect(center.getAttribute('aria-pressed')).toBe('false');
    await update(() => center.click());
    expect(root.querySelector('.docx-toolbar__alignment-popup')).toBeNull();
    expect(editor().snapshot().formatting?.alignment).toBe('center');
    expect(trigger.getAttribute('aria-label')).toBe(en.alignment.center);
  });

  test('a dropdown inside the "⋯" panel takes the first Escape, and the panel the second', async () => {
    useMockResizeObserver();
    const { element, update, pages } = await mount(headerDocx());
    const toolbar = toolbarOf(element);
    await measure(toolbar, 280);
    await update(() =>
      toolbar.querySelector<HTMLButtonElement>('[data-slot="toolbar.more"]')!.click()
    );
    const dropdown = panelOf(element)!.querySelector<HTMLButtonElement>(
      '[data-slot="alignment"] [aria-haspopup]'
    )!;
    expect(dropdown).not.toBeNull();
    await update(() => dropdown.click());
    expect(dropdown.getAttribute('aria-expanded')).toBe('true');
    await update(() => keydown(pages(), 'Escape'));
    expect(dropdown.getAttribute('aria-expanded')).toBe('false');
    expect(panelOf(element)).not.toBeNull();
    await update(() => keydown(pages(), 'Escape'));
    expect(panelOf(element)).toBeNull();
  });
});

describe('colors', () => {
  test('the color popups are dialogs that give the focus back on Escape and after a pick', async () => {
    const { element, update, pages } = await mount();
    for (const slot of ['text.color', 'text.highlight']) {
      const root = element.querySelector<HTMLElement>(`[data-slot="${slot}"]`)!;
      const trigger = root.querySelector<HTMLButtonElement>('[aria-haspopup]')!;
      await update(() => {
        trigger.focus();
        trigger.click();
      });
      const dialog = root.querySelector<HTMLElement>('[role="dialog"]')!;
      expect(dialog).not.toBeNull();
      expect(trigger.getAttribute('aria-haspopup')).toBe('dialog');
      const option = dialog.querySelector<HTMLButtonElement>('button')!;
      await update(() => {
        option.focus();
        keydown(option, 'Escape');
      });
      expect(root.querySelector('[role="dialog"]')).toBeNull();
      expect(document.activeElement).toBe(trigger);
      await update(() => trigger.click());
      const swatch = root.querySelector<HTMLButtonElement>('[role="dialog"] button')!;
      await update(() => {
        swatch.focus();
        swatch.click();
      });
      expect(root.querySelector('[role="dialog"]')).toBeNull();
      expect(document.activeElement).toBe(pages());
    }
  });

  for (const slot of ['text.color', 'text.highlight']) {
    test(`${slot}: a canceled press on the pages closes the popup, and a press inside keeps it`, async () => {
      const { element, update } = await mount();
      const root = element.querySelector<HTMLElement>(`[data-slot="${slot}"]`)!;
      await update(() => root.querySelector<HTMLElement>('[aria-haspopup]')!.click());
      const popup = () => root.querySelector<HTMLElement>('[role="dialog"]');
      await update(() => pointerDown(popup()!.querySelector('button')!));
      expect(popup()).not.toBeNull();
      const target = document.createElement('span');
      element.querySelector('.docx-pages')!.append(target);
      target.addEventListener('pointerdown', (event) => {
        event.preventDefault();
        event.stopPropagation();
      });
      let event!: Event;
      await update(() => (event = pointerDown(target, true)));
      expect(event.defaultPrevented).toBe(true);
      expect(popup()).toBeNull();
    });
  }

  test('a font color swatch applies its color, and the main button applies it again', async () => {
    const { element, update, editor } = await mount();
    const root = element.querySelector<HTMLElement>('[data-slot="text.color"]')!;
    await update(() => root.querySelector<HTMLButtonElement>('[aria-haspopup]')!.click());
    const blue = root.querySelector<HTMLButtonElement>('[data-value="0070C0"]')!;
    expect(blue.getAttribute('aria-label')).toBe(en.colorPicker.colors.blue);
    await update(() => blue.click());
    expect(editor().snapshot().formatting?.color).toEqual({ kind: 'hex', value: '0070C0' });
    await update(() => root.querySelector<HTMLButtonElement>('[aria-haspopup]')!.click());
    expect(root.querySelector('[data-value="0070C0"]')!.getAttribute('aria-pressed')).toBe('true');
    // Automatic clears the color.
    await update(() =>
      root.querySelector<HTMLButtonElement>('.docx-toolbar__swatch-clear')!.click()
    );
    expect(editor().snapshot().formatting?.color?.kind).not.toBe('hex');
    // The main button applies the last color that a swatch applied.
    await update(() =>
      root.querySelector<HTMLButtonElement>('.docx-toolbar__colorsplit-main')!.click()
    );
    expect(editor().snapshot().formatting?.color).toEqual({ kind: 'hex', value: '0070C0' });
  });

  test('a custom hex color applies, and only six hex digits enable Apply', async () => {
    const { element, update, editor } = await mount();
    const root = element.querySelector<HTMLElement>('[data-slot="text.color"]')!;
    await update(() => root.querySelector<HTMLButtonElement>('[aria-haspopup]')!.click());
    const field = root.querySelector<HTMLInputElement>('.docx-toolbar__swatch-hex')!;
    const apply = root.querySelector<HTMLButtonElement>('.docx-toolbar__swatch-apply')!;
    field.value = '12zz34';
    await update(() => field.dispatchEvent(new Event('input')));
    expect(field.value).toBe('1234');
    expect(apply.disabled).toBe(true);
    field.value = '1234ab';
    await update(() => field.dispatchEvent(new Event('input')));
    expect(apply.disabled).toBe(false);
    await update(() => apply.click());
    expect(editor().snapshot().formatting?.color).toEqual({ kind: 'hex', value: '1234AB' });
  });

  test('a highlight swatch applies, and No color removes it', async () => {
    const { element, update, editor } = await mount();
    const root = element.querySelector<HTMLElement>('[data-slot="text.highlight"]')!;
    await update(() => root.querySelector<HTMLButtonElement>('[aria-haspopup]')!.click());
    await update(() => root.querySelector<HTMLButtonElement>('[data-value="green"]')!.click());
    expect(editor().snapshot().formatting?.highlight).toBe('green');
    await update(() => root.querySelector<HTMLButtonElement>('[aria-haspopup]')!.click());
    await update(() =>
      root.querySelector<HTMLButtonElement>('.docx-toolbar__swatch-clear')!.click()
    );
    expect(editor().snapshot().formatting?.highlight ?? null).toBeNull();
  });
});

describe('lists and indent', () => {
  test('the list buttons toggle and report their state, and indent follows the list', async () => {
    const { element, update, editor } = await mount();
    const bullet = element.querySelector<HTMLButtonElement>('[data-slot="list.bullet"]')!;
    const numbered = element.querySelector<HTMLButtonElement>('[data-slot="list.numbered"]')!;
    const indent = element.querySelector<HTMLButtonElement>('[data-slot="list.indent"]')!;
    expect(bullet.disabled).toBe(false);
    expect(bullet.hasAttribute('data-active')).toBe(false);
    await update(() => bullet.click());
    expect(bullet.hasAttribute('data-active')).toBe(true);
    expect(numbered.hasAttribute('data-active')).toBe(false);
    expect(indent.disabled).toBe(false);
    await update(() => indent.click());
    expect(bullet.hasAttribute('data-active')).toBe(true);
    expect(editor().snapshot().canUndo).toBe(true);
    await update(() => numbered.click());
    expect(numbered.hasAttribute('data-active')).toBe(true);
    expect(bullet.hasAttribute('data-active')).toBe(false);
  });

  test('in viewing mode, the controls of the task are disabled with the engine reason', async () => {
    const { element, update, editor } = await mount();
    await update(() => editor().exec({ type: 'setEditingMode', mode: 'viewing' }));
    for (const selector of [
      '[data-slot="text.color"] .docx-toolbar__colorsplit-main',
      '[data-slot="text.highlight"] [aria-haspopup]',
      '[data-slot="list.lineSpacing"] [aria-haspopup]',
      '[data-slot="alignment"] [aria-haspopup]',
      '[data-slot="list.bullet"]',
      '[data-slot="list.indent"]',
    ]) {
      const control = element.querySelector<HTMLButtonElement>(selector)!;
      expect(control.disabled, selector).toBe(true);
      expect(control.hasAttribute('data-disabled'), selector).toBe(true);
      expect(control.title.length, selector).toBeGreaterThan(0);
    }
  });
});
