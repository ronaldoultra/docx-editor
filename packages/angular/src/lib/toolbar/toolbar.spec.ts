// The toolbar claims of packages/vue/test/chrome-composition.test.ts and
// packages/vue/test/toolbar-host-groups.test.ts, for `fieldset[docxEditorToolbar]`.
import {
  ChangeDetectionStrategy,
  Component,
  Directive,
  signal,
  viewChild,
  type Type,
} from '@angular/core';
import type { ComponentFixture } from '@angular/core/testing';
import type { DocxEditorInstance } from '@docx-editor.dev/core/editor';
import { createT, en, type TranslationKey } from '@docx-editor.dev/core/i18n';
import { afterEach, beforeEach, describe, expect, test, vi } from 'vitest';
import { docx, headerDocx, paragraph } from '../../testing/docx';
import { configureZoneless, engineOf, render, settle } from '../../testing/harness';
import {
  barShape,
  expectedArrangement,
  flushObservers,
  measure,
  MockResizeObserver,
  panelOf,
  pressEscape,
  restoreResizeObserver,
  setWidth,
  toolbarArrangement,
  toolbarOf,
  useMockResizeObserver,
} from '../../testing/toolbar';
import { DocxEditorContentComponent } from '../content.component';
import { DocxEditorComponent } from '../docx-editor.component';
import { DocxEditorRootDirective } from '../root.directive';
import { DocxSlotDirective } from '../slot.directive';
import { DocxEditorViewportComponent } from '../viewport.component';
import { DocxToolbarActionComponent } from './toolbar-action.component';
import {
  DocxToolbarButtonComponent,
  DocxToolbarSeparatorComponent,
} from './toolbar-button.component';
import { DocxToolbarGroupDirective } from './toolbar-group.directive';
import { DocxEditorToolbarComponent } from './toolbar.component';

const SOURCE = docx(paragraph('hello world'));
const label = createT(en);
const byLabel = (key: string) => `[aria-label=${JSON.stringify(label(key as TranslationKey))}]`;
const t = (key: string) => (key === 'formattingBar.more' ? 'More' : undefined);
/** The alignment dropdown takes one place in the default arrangement. */
const ALIGNMENT = 'alignment';

const PARTS = [
  DocxEditorRootDirective,
  DocxEditorViewportComponent,
  DocxEditorContentComponent,
  DocxEditorToolbarComponent,
  DocxToolbarButtonComponent,
  DocxToolbarSeparatorComponent,
  DocxToolbarActionComponent,
  DocxToolbarGroupDirective,
  DocxSlotDirective,
];

/** The base of each host: the engine of its root, and the document it opens. */
@Directive()
abstract class Host {
  readonly bytes = signal<Uint8Array>(SOURCE);
  readonly t = t;
  readonly root = viewChild.required(DocxEditorRootDirective);

  editor(): DocxEditorInstance {
    return engineOf(this.root().context.editor());
  }
}

const OPEN = '<ng-container docxEditorRoot [document]="bytes()">';
const CLOSE = '<docx-editor-viewport><docx-editor-content /></docx-editor-viewport></ng-container>';

@Component({
  imports: PARTS,
  changeDetection: ChangeDetectionStrategy.OnPush,
  template: OPEN + '<fieldset docxEditorToolbar [t]="t"></fieldset>' + CLOSE,
})
class DefaultHostComponent extends Host {}

@Component({
  imports: PARTS,
  changeDetection: ChangeDetectionStrategy.OnPush,
  template:
    OPEN +
    `<fieldset docxEditorToolbar>
        <ng-template docxSlot="text.strike" docxSlotHidden />
      </fieldset>` +
    CLOSE,
})
class Host1Component extends Host {}

@Component({
  imports: PARTS,
  changeDetection: ChangeDetectionStrategy.OnPush,
  template:
    OPEN +
    `<fieldset docxEditorToolbar>
        <ng-template docxSlot="text.bold">
          <button docxToolbarButton="text.bold" class="custom-bold"></button>
        </ng-template>
      </fieldset>` +
    CLOSE,
})
class Host2Component extends Host {}

@Component({
  imports: PARTS,
  changeDetection: ChangeDetectionStrategy.OnPush,
  template:
    OPEN +
    `<fieldset docxEditorToolbar [preset]="false">
        <button docxToolbarButton="text.bold"></button>
        <div docxToolbarSeparator></div>
        <button docxToolbarButton="history.undo"></button>
      </fieldset>` +
    CLOSE,
})
class Host3Component extends Host {}

@Component({
  imports: PARTS,
  changeDetection: ChangeDetectionStrategy.OnPush,
  template:
    OPEN +
    `<fieldset docxEditorToolbar>
        <ng-template docxToolbarGroup="mine" label="Mine" after="history">
          <button docxToolbarAction label="Go"></button>
        </ng-template>
        <ng-template docxToolbarGroup="font">
          <button docxToolbarAction label="Extra"></button>
        </ng-template>
        <ng-template docxToolbarGroup="zoom" hidden />
        <ng-template docxSlot="review.editingMode" docxSlotHidden />
        <ng-template docxSlot="review.comments" docxSlotHidden />
        <ng-template docxSlot="review.revisionMarkup" docxSlotHidden />
      </fieldset>` +
    CLOSE,
})
class Host4Component extends Host {}

@Component({
  imports: PARTS,
  changeDetection: ChangeDetectionStrategy.OnPush,
  template:
    OPEN +
    `<fieldset docxEditorToolbar>
        <ng-template docxToolbarGroup="a" label="formattingBar.groups.font">
          <button docxToolbarAction label="A"></button>
        </ng-template>
        <ng-template docxToolbarGroup="b" labelKey="formattingBar.groups.font">
          <button docxToolbarAction label="B"></button>
        </ng-template>
        <ng-template docxToolbarGroup="c" label="Mine" labelKey="formattingBar.groups.font">
          <button docxToolbarAction label="C"></button>
        </ng-template>
        <ng-template docxToolbarGroup="d"><button docxToolbarAction label="D"></button></ng-template>
      </fieldset>` +
    CLOSE,
})
class Host5Component extends Host {}

@Component({
  imports: PARTS,
  changeDetection: ChangeDetectionStrategy.OnPush,
  template:
    OPEN +
    `<fieldset docxEditorToolbar>
        <ng-template docxToolbarGroup="mine" className="from-name">
          <button docxToolbarAction label="A"></button>
        </ng-template>
      </fieldset>` +
    CLOSE,
})
class Host6Component extends Host {}

@Component({
  imports: PARTS,
  changeDetection: ChangeDetectionStrategy.OnPush,
  template:
    OPEN +
    `<fieldset docxEditorToolbar [preset]="false">
        <ng-template docxToolbarGroup="mine" className="from-name">
          <button docxToolbarAction label="A"></button>
        </ng-template>
      </fieldset>` +
    CLOSE,
})
class Host7Component extends Host {}

@Component({
  imports: PARTS,
  changeDetection: ChangeDetectionStrategy.OnPush,
  template:
    OPEN +
    `<fieldset docxEditorToolbar [t]="t">
        <ng-template docxToolbarGroup="zoom" [overflowContent]="extra" />
        <ng-template #extra><span data-testid="zoom-extra"></span></ng-template>
      </fieldset>` +
    CLOSE,
})
class Host8Component extends Host {}

@Component({
  imports: PARTS,
  changeDetection: ChangeDetectionStrategy.OnPush,
  template:
    OPEN +
    `<fieldset docxEditorToolbar [t]="t">
        <ng-template docxSlot="zoom.level" [docxSlotOverflow]="panelZoom">
          <button type="button" data-testid="my-zoom">Z</button>
        </ng-template>
        <ng-template #panelZoom><span data-testid="panel-zoom"></span></ng-template>
      </fieldset>` +
    CLOSE,
})
class Host9Component extends Host {}

@Component({
  imports: PARTS,
  changeDetection: ChangeDetectionStrategy.OnPush,
  template:
    OPEN +
    `<fieldset docxEditorToolbar>
        <button docxToolbarButton="format.painter"></button>
      </fieldset>` +
    CLOSE,
})
class Host10Component extends Host {}

@Component({
  imports: PARTS,
  changeDetection: ChangeDetectionStrategy.OnPush,
  template: '<fieldset docxEditorToolbar></fieldset>',
})
class LooseHostComponent {}

@Component({
  imports: PARTS,
  changeDetection: ChangeDetectionStrategy.OnPush,
  template: `<ng-container docxEditorRoot [document]="bytes" colorMode="dark">
    <fieldset docxEditorToolbar></fieldset>
    <docx-editor-viewport><docx-editor-content /></docx-editor-viewport>
  </ng-container>`,
})
class DarkHostComponent {
  readonly bytes = SOURCE;
}

@Component({
  imports: PARTS,
  changeDetection: ChangeDetectionStrategy.OnPush,
  template:
    OPEN +
    `<fieldset docxEditorToolbar>
  <ng-template docxToolbarGroup="ng-mine" after="ng-missing">
    <button docxToolbarAction label="Go"></button>
  </ng-template>
  <ng-template [docxSlot]="nope"><span data-testid="nope"></span></ng-template>
  <ng-template docxToolbarGroup="font" label="Fonts" labelKey="x.y" after="history" [priority]="95">
    <button docxToolbarAction label="F"></button>
  </ng-template>
  <ng-template docxToolbarGroup="contextual-table" label="Mine">
    <button docxToolbarAction label="G"></button>
  </ng-template>
    </fieldset>` +
    CLOSE,
})
class WarningHostComponent extends Host {
  readonly nope = 'ng.nope' as never;
}

async function mount<T extends Host>(
  component: Type<T>,
  bytes?: Uint8Array
): Promise<{
  fixture: ComponentFixture<T>;
  element: HTMLElement;
  editor: () => DocxEditorInstance;
}> {
  const fixture = await render(component);
  if (bytes) {
    fixture.componentInstance.bytes.set(bytes);
    await settle(fixture);
  }
  const element: HTMLElement = fixture.nativeElement;
  document.body.append(element);
  return { fixture, element, editor: () => fixture.componentInstance.editor() };
}

beforeEach(() => configureZoneless());

afterEach(() => {
  restoreResizeObserver();
  document.body.replaceChildren();
});

describe('default arrangement', () => {
  test('the groups and controls follow the engine registry, with their slot ids', async () => {
    const { element } = await mount(DefaultHostComponent);
    expect(toolbarArrangement(toolbarOf(element))).toEqual(expectedArrangement(ALIGNMENT));
  });

  test('a docxSlotHidden template removes a slot from the default bar', async () => {
    const { element } = await mount(Host1Component);
    expect(element.querySelector(byLabel('formattingBar.strikethrough'))).toBeNull();
    expect(toolbarArrangement(toolbarOf(element)).length).toBe(
      expectedArrangement(ALIGNMENT).length - 1
    );
  });

  test('a slot template keeps registry order and its own class', async () => {
    const { element } = await mount(Host2Component);
    const bold = element.querySelector(byLabel('formattingBar.boldShortcut'))!;
    expect(bold.className).toContain('custom-bold');
    expect(bold.className).toContain('docx-toolbar__button');
    expect(toolbarArrangement(toolbarOf(element))).toEqual(expectedArrangement(ALIGNMENT));
  });

  test('[preset]="false" renders only the composed content', async () => {
    const { element } = await mount(Host3Component);
    expect(toolbarArrangement(toolbarOf(element))).toEqual([
      'text.bold',
      'separator',
      'history.undo',
    ]);
  });

  test('Bold reflects live engine state after a click', async () => {
    const { fixture, element, editor } = await mount(DefaultHostComponent);
    editor().surface!.selectAll();
    await settle(fixture);
    const bold = element.querySelector<HTMLButtonElement>(byLabel('formattingBar.boldShortcut'))!;
    expect(bold.hasAttribute('data-active')).toBe(false);
    expect(bold.getAttribute('aria-pressed')).toBe('false');
    bold.click();
    await settle(fixture);
    expect(bold.hasAttribute('data-active')).toBe(true);
    expect(bold.getAttribute('aria-pressed')).toBe('true');
    expect(editor().snapshot().formatting?.bold).toBe(true);
  });

  test('outside a root, the toolbar is disabled and its controls show the engine reason', async () => {
    const fixture = await render(LooseHostComponent);
    const toolbar = toolbarOf(fixture.nativeElement) as HTMLFieldSetElement;
    expect(toolbar.disabled).toBe(true);
    expect(toolbar.classList.contains('docx-editor')).toBe(true);
    const undo = toolbar.querySelector<HTMLButtonElement>('[data-slot="history.undo"]')!;
    expect(undo.disabled).toBe(true);
    expect(undo.hasAttribute('data-disabled')).toBe(true);
  });

  test('the toolbar carries the scope class, and the dark class under a dark root', async () => {
    const fixture = await render(DarkHostComponent);
    const toolbar = toolbarOf(fixture.nativeElement) as HTMLFieldSetElement;
    expect(toolbar.classList.contains('docx-editor')).toBe(true);
    expect(toolbar.classList.contains('dark')).toBe(true);
    expect(toolbar.getAttribute('aria-label')).toBe(en.formattingBar.label);
    expect(toolbar.disabled).toBe(false);
  });
});

/** A bar of 280 px with 90 px groups, and the "⋯" panel open. */
async function collapsedMore(bytes = headerDocx()) {
  useMockResizeObserver();
  const view = await mount(DefaultHostComponent, bytes);
  const toolbar = toolbarOf(view.element);
  await measure(toolbar, 280);
  const trigger = toolbar.querySelector<HTMLButtonElement>('[data-slot="toolbar.more"]')!;
  expect(trigger).not.toBeNull();
  return { ...view, toolbar, trigger };
}

describe('the "⋯" panel', () => {
  test('Escape closes the panel before the header scope or the format painter', async () => {
    const { fixture, element, editor, trigger } = await collapsedMore();
    const pages = element.querySelector<HTMLElement>('.docx-pages')!;
    expect(editor().exec({ type: 'editHeaderFooter', position: 'header' }).ok).toBe(true);
    pages.focus();
    trigger.click();
    await settle(fixture);
    expect(panelOf(element)).not.toBeNull();
    expect((await pressEscape(pages)).defaultPrevented).toBe(true);
    expect(panelOf(element)).toBeNull();
    expect(editor().surface!.activeScope().kind).toBe('headerFooter');
    // With nothing open, the next Escape reaches the surface and leaves the header.
    await pressEscape(pages);
    expect(editor().surface!.activeScope().kind).toBe('body');

    expect(editor().surface!.formatPainter.press()).toBe(true);
    trigger.click();
    await settle(fixture);
    expect(panelOf(element)).not.toBeNull();
    await pressEscape(pages);
    expect(panelOf(element)).toBeNull();
    expect(editor().surface!.formatPainter.state().mode).not.toBe('off');
    await pressEscape(pages);
    expect(editor().surface!.formatPainter.state().mode).toBe('off');
  });

  test('the panel closes when the focus moves to a control outside it', async () => {
    const { fixture, element, trigger } = await collapsedMore();
    trigger.click();
    await settle(fixture);
    expect(panelOf(element)).not.toBeNull();
    const outside = document.createElement('input');
    document.body.append(outside);
    outside.focus();
    await settle(fixture);
    expect(panelOf(element)).toBeNull();
  });

  test('ArrowDown on the trigger opens the panel with the focus on its first control', async () => {
    const { fixture, element, trigger } = await collapsedMore();
    trigger.dispatchEvent(new KeyboardEvent('keydown', { key: 'ArrowDown', bubbles: true }));
    await settle(fixture);
    const panel = panelOf(element)!;
    expect(trigger.getAttribute('aria-expanded')).toBe('true');
    expect(trigger.getAttribute('aria-controls')).toBe(panel.id);
    expect(panel.contains(document.activeElement)).toBe(true);
    // Escape inside the panel closes it and gives the focus back to the trigger.
    await pressEscape(document.activeElement!);
    expect(panelOf(element)).toBeNull();
    expect(document.activeElement).toBe(trigger);
  });

  test('stays inside a narrow viewport, also in a right-to-left bar', async () => {
    for (const dir of ['ltr', 'rtl']) {
      useMockResizeObserver();
      const { fixture, element } = await mount(DefaultHostComponent);
      const toolbar = toolbarOf(element);
      toolbar.setAttribute('dir', dir);
      await measure(toolbar, 300);
      const trigger = toolbar.querySelector<HTMLButtonElement>('[aria-label="More"]')!;
      const rect = { left: 178, right: 212, top: 0, bottom: 34, width: 34, height: 34 } as DOMRect;
      trigger.getBoundingClientRect = () => rect;
      trigger.parentElement!.getBoundingClientRect = () => rect;
      const view = element.ownerDocument.defaultView!;
      const innerWidth = view.innerWidth;
      Object.defineProperty(view, 'innerWidth', { configurable: true, value: 390 });
      const offsetWidth = Object.getOwnPropertyDescriptor(HTMLElement.prototype, 'offsetWidth');
      Object.defineProperty(HTMLElement.prototype, 'offsetWidth', {
        configurable: true,
        get(this: HTMLElement) {
          return this.classList.contains('docx-toolbar__more-panel') ? 340 : 0;
        },
      });
      try {
        trigger.click();
        await settle(fixture);
        const panel = panelOf(element)!;
        expect(panel.getAttribute('data-anchor')).toBe('clamped');
        expect(panel.style.left).toBe('-170px');
        expect(panel.style.right).toBe('auto');
        expect(panel.style.maxInlineSize).toBe('374px');
      } finally {
        if (offsetWidth) Object.defineProperty(HTMLElement.prototype, 'offsetWidth', offsetWidth);
        Object.defineProperty(view, 'innerWidth', { configurable: true, value: innerWidth });
      }
      fixture.destroy();
      restoreResizeObserver();
    }
  });
});

describe('per-control collapse', () => {
  test('controls leave one at a time from the end of a group, and "⋯" holds exactly those', async () => {
    useMockResizeObserver();
    const { fixture, element } = await mount(DefaultHostComponent);
    const toolbar = toolbarOf(element);
    const items = (group: Element) =>
      [...group.querySelectorAll<HTMLElement>(':scope > [data-toolbar-item]')].map((item) =>
        item.getAttribute('data-toolbar-item')!
      );
    const full = new Map<string, string[]>();
    for (const group of toolbar.querySelectorAll<HTMLElement>('[data-toolbar-group]')) {
      if (group.children.length > 0)
        full.set(group.getAttribute('data-toolbar-group')!, items(group));
    }
    // 30 px for each control, no gaps (no stylesheet here), a 1 px separator, 100 px fixed.
    for (const item of toolbar.querySelectorAll('[data-toolbar-item]'))
      setWidth(item, 'offsetWidth', 30);
    for (const group of toolbar.querySelectorAll('[data-toolbar-group]')) {
      // As wide as the controls that it still draws, as in a browser.
      setWidth(
        group,
        'offsetWidth',
        () => group.querySelectorAll(':scope > [data-toolbar-item]').length * 30
      );
    }
    for (const fixed of toolbar.querySelectorAll('[data-toolbar-fixed]'))
      setWidth(fixed, 'offsetWidth', 100);
    setWidth(toolbar.querySelector('.docx-toolbar__separator')!, 'offsetWidth', 1);
    let total = 100;
    for (const list of full.values()) total += list.length * 30 + 1;
    setWidth(toolbar, 'clientWidth', total - 31 + 34 - 10);
    await flushObservers();

    const hidden: string[] = [];
    const partial: string[] = [];
    for (const [id, list] of full) {
      const group = toolbar.querySelector<HTMLElement>(`[data-toolbar-group="${id}"]`);
      const shown = group ? items(group) : [];
      expect(shown).toEqual(list.slice(0, shown.length));
      hidden.push(...list.slice(shown.length));
      if (shown.length > 0 && shown.length < list.length) partial.push(id);
    }
    expect(hidden[0]).toBe('zoom.level');
    expect(hidden.length).toBe(2);
    expect(partial.length).toBe(1);

    toolbar.querySelector<HTMLButtonElement>('[aria-label="More"]')!.click();
    await settle(fixture);
    const panel = panelOf(element)!;
    expect(
      [...panel.querySelectorAll<HTMLElement>('[data-slot]')].map((e) =>
        e.getAttribute('data-slot')
      )
    ).toEqual(hidden);
    expect(
      [...panel.querySelectorAll('.docx-toolbar__more-section')].map((s) =>
        s.getAttribute('aria-label')
      )
    ).toEqual([
      label('formattingBar.groups.zoom' as TranslationKey),
      label(
        `formattingBar.groups.${partial[0] === 'text' ? 'textFormatting' : partial[0]}` as TranslationKey
      ),
    ]);
  });

  test('a content-sized bar gets its groups back when the room grows again', async () => {
    useMockResizeObserver();
    const { element } = await mount(DefaultHostComponent);
    const toolbar = toolbarOf(element);
    const parent = toolbar.parentElement!;
    parent.style.display = 'flex';
    parent.style.justifyContent = 'center';
    toolbar.style.width = 'max-content';
    toolbar.style.maxWidth = '100%';
    toolbar.style.boxSizing = 'border-box';
    (toolbar as unknown as { computedStyleMap: () => unknown }).computedStyleMap = () => ({
      get: (property: string) => (property === 'width' ? { value: 'max-content' } : undefined),
    });
    let room = 2000;
    setWidth(parent, 'clientWidth', () => room);
    const contentWidth = () =>
      toolbar.querySelectorAll('[data-toolbar-group]').length * 90 +
      toolbar.querySelectorAll('[data-toolbar-fixed]').length * 90 +
      (toolbar.querySelector('[data-toolbar-more]') ? 34 : 0);
    setWidth(toolbar, 'clientWidth', () => Math.min(contentWidth(), room));
    setWidth(toolbar, 'offsetWidth', () => Math.min(contentWidth(), room));
    const flushGroups = async () => {
      for (const group of toolbar.querySelectorAll('[data-toolbar-group], [data-toolbar-fixed]')) {
        setWidth(group, 'offsetWidth', 90);
      }
      await flushObservers();
    };
    await flushGroups();
    const all = toolbar.querySelectorAll('[data-toolbar-group]').length;
    expect(toolbar.querySelector('[aria-label="More"]')).toBeNull();

    room = 400;
    await flushGroups();
    expect(toolbar.querySelectorAll('[data-toolbar-group]').length).toBeLessThan(all);
    expect(toolbar.querySelector('[aria-label="More"]')).not.toBeNull();

    room = 2000;
    await flushGroups();
    expect(toolbar.querySelectorAll('[data-toolbar-group]').length).toBe(all);
    expect(toolbar.querySelector('[aria-label="More"]')).toBeNull();
  });

  test('the bar observes itself, its parent, and each group', async () => {
    useMockResizeObserver();
    const { element } = await mount(DefaultHostComponent);
    const toolbar = toolbarOf(element);
    const observed = MockResizeObserver.instances.flatMap((observer) => observer.observed);
    expect(observed).toContain(toolbar);
    expect(observed).toContain(toolbar.parentElement);
    const contextual = toolbar.querySelector('.docx-toolbar__contextual')!;
    expect(contextual.getAttribute('data-toolbar-group')).toBe('contextual-table');
    expect(observed).toContain(contextual);
    expect(toolbar.hasAttribute('data-overflow')).toBe(true);
  });
});

@Component({
  imports: PARTS,
  changeDetection: ChangeDetectionStrategy.OnPush,
  template:
    OPEN +
    `<fieldset docxEditorToolbar [t]="t">
    <ng-template docxToolbarGroup="review-nav" label="Review navigation">
      <button docxToolbarAction label="Next change" (select)="selected = selected + 1">N</button>
    </ng-template>
  </fieldset>` +
    CLOSE,
})
class HostGroupComponent extends Host {
  selected = 0;
}

describe('host groups', () => {
  test('a host group is measured, collapses, and shows its actions as panel rows', async () => {
    useMockResizeObserver();
    const { fixture, element } = await mount(HostGroupComponent);
    const toolbar = toolbarOf(element);
    const group = toolbar.querySelector('[data-toolbar-group="review-nav"]');
    expect(group?.getAttribute('role')).toBe('group');
    expect(group?.getAttribute('aria-label')).toBe('Review navigation');
    const shape = barShape(toolbar);
    expect(shape.indexOf('review-nav')).toBe(shape.lastIndexOf('|') + 1);
    expect(shape.filter((entry) => entry === 'fixed').length).toBe(1);

    await measure(toolbar, 420);
    expect(toolbar.querySelector('[data-toolbar-group="review-nav"]')).toBeNull();
    toolbar.querySelector<HTMLButtonElement>('[aria-label="More"]')!.click();
    await settle(fixture);
    const row = panelOf(element)!.querySelector<HTMLButtonElement>(
      '[aria-label="Review navigation"] .docx-toolbar__more-command'
    );
    expect(row?.textContent).toContain('Next change');
    row!.click();
    await settle(fixture);
    expect(fixture.componentInstance.selected).toBe(1);
    expect(panelOf(element)).toBeNull();
  });

  test('after, built-in ids, hidden groups, and groups whose slots are all hidden', async () => {
    const { element } = await mount(Host4Component);
    const toolbar = toolbarOf(element);
    const shape = barShape(toolbar);
    expect(shape.slice(0, 3)).toEqual(['history', '|', 'mine']);
    expect(
      toolbar.querySelector('[data-toolbar-group="font"] [aria-label="Extra"]')
    ).not.toBeNull();
    expect(toolbar.querySelector('[data-toolbar-group="zoom"]')).toBeNull();
    // The review group lost every slot: no empty group, no trailing or doubled separator.
    expect(shape).not.toContain('fixed');
    expect(shape[shape.length - 1]).not.toBe('|');
    expect(shape.join(' ')).not.toContain('| |');
    const groups = [...toolbar.querySelectorAll('.docx-toolbar__group:not([hidden])')];
    expect(groups.every((group) => group.children.length > 0)).toBe(true);
  });

  test('label renders as written, labelKey is translated, and the id is the fallback', async () => {
    const { element } = await mount(Host5Component);
    const toolbar = toolbarOf(element);
    const name = (id: string) =>
      toolbar.querySelector(`[data-toolbar-group="${id}"]`)!.getAttribute('aria-label');
    // A literal label is never looked up, even when it reads like a key.
    expect(name('a')).toBe('formattingBar.groups.font');
    expect(name('b')).toBe(en.formattingBar.groups.font);
    expect(name('c')).toBe('Mine');
    expect(name('d')).toBe('d');
  });

  test('className styles the group, in the default bar and in a bare bar', async () => {
    const preset = await mount(Host6Component);
    const group = preset.element.querySelector('[data-toolbar-group="mine"]')!;
    expect(group.classList.contains('from-name')).toBe(true);
    preset.fixture.destroy();
    const bare = await mount(Host7Component);
    const own = bare.element.querySelector('[data-toolbar-host-group="mine"]')!;
    expect(own.classList.contains('from-name')).toBe(true);
    expect(own.classList.contains('docx-toolbar__group')).toBe(true);
    expect(own.querySelector('[aria-label="A"]')).not.toBeNull();
  });

  test('overflowContent on a built-in group renders in the panel', async () => {
    useMockResizeObserver();
    const { fixture, element } = await mount(Host8Component);
    const toolbar = toolbarOf(element);
    await measure(toolbar, 760);
    toolbar.querySelector<HTMLButtonElement>('[aria-label="More"]')!.click();
    await settle(fixture);
    expect(panelOf(element)!.querySelector('[data-testid="zoom-extra"]')).not.toBeNull();
  });

  test('unknown anchors, unknown slot ids, built-in settings, and the reserved id warn', async () => {
    const warn = vi.spyOn(console, 'warn').mockImplementation(() => {});
    let messages: string[] = [];
    try {
      const { element } = await mount(WarningHostComponent);
      expect(element.querySelector('[aria-label="G"]')).toBeNull();
      expect(element.querySelector('[aria-label="F"]')).not.toBeNull();
      // The content of an unknown slot renders after the groups.
      expect(element.querySelector('[data-toolbar-fixed] [data-testid="nope"]')).not.toBeNull();
      messages = warn.mock.calls.map((call) => String(call[0]));
    } finally {
      warn.mockRestore();
    }
    expect(messages.some((text) => text.includes('ng-missing'))).toBe(true);
    expect(messages.some((text) => text.includes('ng.nope'))).toBe(true);
    for (const setting of ['label', 'labelKey', 'after']) {
      expect(
        messages.some((text) => text.includes('"font"') && text.includes(` ${setting} `))
      ).toBe(true);
    }
    expect(messages.some((text) => text.includes('priority'))).toBe(false);
    expect(messages.some((text) => text.includes('"contextual-table"'))).toBe(true);
  });
});

describe('slot templates', () => {
  test('a slot template replaces a built-in slot in place and collapses with its group', async () => {
    useMockResizeObserver();
    const { fixture, element } = await mount(Host9Component);
    const toolbar = toolbarOf(element);
    expect(
      toolbar.querySelector('[data-toolbar-group="zoom"] [data-testid="my-zoom"]')
    ).not.toBeNull();
    expect(toolbar.querySelector('[data-slot="zoom.level"]')).toBeNull();
    await measure(toolbar, 760);
    toolbar.querySelector<HTMLButtonElement>('[aria-label="More"]')!.click();
    await settle(fixture);
    expect(panelOf(element)!.querySelector('[data-testid="panel-zoom"]')).not.toBeNull();
  });

  test('a docxToolbarButton outside the arrangement renders after the groups', async () => {
    const { element } = await mount(Host10Component);
    const toolbar = toolbarOf(element);
    const fixed = [...toolbar.querySelectorAll('[data-toolbar-fixed]')].at(-1)!;
    expect(fixed.querySelector('button[data-slot="format.painter"]')).not.toBeNull();
    expect(barShape(toolbar).filter((entry) => entry === 'fixed').length).toBe(2);
  });
});

@Component({
  imports: [DocxEditorComponent],
  changeDetection: ChangeDetectionStrategy.OnPush,
  template: `<docx-editor [document]="bytes" [chrome]="chrome()" />`,
})
class PackagedHostComponent {
  readonly bytes = SOURCE;
  readonly chrome = signal(true);
}

describe('the packaged editor', () => {
  test('shows the toolbar above the pages, and chrome=false hides it', async () => {
    const fixture = await render(PackagedHostComponent);
    const element: HTMLElement = fixture.nativeElement;
    const wrapper = element.querySelector('.docx-editor')!;
    const toolbar = toolbarOf(element);
    expect(toolbar.parentElement).toBe(wrapper.firstElementChild?.parentElement);
    // The wrapper carries the scope class, so the toolbar does not repeat it.
    expect(toolbar.classList.contains('docx-editor')).toBe(false);
    const viewport = element.querySelector('docx-editor-viewport')!;
    expect(
      toolbar.compareDocumentPosition(viewport) & Node.DOCUMENT_POSITION_FOLLOWING
    ).toBeTruthy();
    expect(toolbar.querySelector('[data-slot="text.bold"]')).not.toBeNull();
    fixture.componentInstance.chrome.set(false);
    await settle(fixture);
    expect(element.querySelector('[data-testid="docx-toolbar"]')).toBeNull();
    expect(element.querySelector('.docx-page')).not.toBeNull();
  });
});
