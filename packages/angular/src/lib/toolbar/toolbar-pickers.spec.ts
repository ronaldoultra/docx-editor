// The claims of packages/vue/test/toolbar-picker-accessibility.test.ts (the font family, style,
// zoom, and font size parts; the color parts port with task 4.3), the claim of
// paragraph-style-labels.test.ts, the picker claim of host-window-chrome.test.ts, the
// editing-mode claim of chrome-composition.test.ts, and toolbar-dev-warnings.test.ts.
import { ChangeDetectionStrategy, Component, Directive, signal, viewChild } from '@angular/core';
import { TestBed } from '@angular/core/testing';
import type { DocxEditorInstance, EditorModule } from '@docx-editor.dev/core/editor';
import { en, type Translations } from '@docx-editor.dev/core/i18n';
import { afterEach, beforeEach, describe, expect, test, vi } from 'vitest';
import { headerDocx, paragraph, paragraphStyle, stylesDocx } from '../../testing/docx';
import { configureZoneless, engineOf, render, settle } from '../../testing/harness';
import { DocxEditorContentComponent } from '../content.component';
import { DocxEditorI18nDirective } from '../i18n';
import { DocxEditorRootDirective } from '../root.directive';
import { DocxEditorViewportComponent } from '../viewport.component';
import { DocxEditorWindowComponent } from '../window.component';
import { DocxToolbarEditingModeComponent } from './toolbar-editing-mode.component';
import {
  DocxToolbarFontFamilyComponent,
  DocxToolbarParagraphStyleComponent,
} from './toolbar-pickers.component';
import { toolbarDevWarning } from './toolbar-registry';
import {
  DocxToolbarFontSizeComponent,
  DocxToolbarZoomComponent,
} from './toolbar-steppers.component';
import { DocxEditorToolbarComponent } from './toolbar.component';

const PICKER_SOURCE = stylesDocx(
  paragraphStyle('Normal', 'Normal', ' w:default="1"') + paragraphStyle('Heading1', 'Heading 1'),
  paragraph('hello world')
);

const PARTS = [
  DocxEditorRootDirective,
  DocxEditorViewportComponent,
  DocxEditorContentComponent,
  DocxEditorToolbarComponent,
  DocxToolbarFontFamilyComponent,
  DocxToolbarParagraphStyleComponent,
  DocxToolbarFontSizeComponent,
  DocxToolbarZoomComponent,
  DocxToolbarEditingModeComponent,
];

@Directive()
abstract class Host {
  readonly bytes = signal<Uint8Array>(PICKER_SOURCE);
  readonly root = viewChild.required(DocxEditorRootDirective);

  editor(): DocxEditorInstance {
    return engineOf(this.root().context.editor());
  }
}

@Component({
  imports: PARTS,
  changeDetection: ChangeDetectionStrategy.OnPush,
  template: `<ng-container docxEditorRoot [document]="bytes()">
    <fieldset docxEditorToolbar [preset]="false" [overflow]="false">
      <div docxToolbarFontFamily></div>
      <div docxToolbarParagraphStyle></div>
      <span docxToolbarZoom></span>
      <span docxToolbarFontSize></span>
      <div docxToolbarEditingMode></div>
    </fieldset>
    <docx-editor-viewport><docx-editor-content /></docx-editor-viewport>
  </ng-container>`,
})
class PickersHostComponent extends Host {}

async function mount(bytes?: Uint8Array) {
  const fixture = await render(PickersHostComponent);
  if (bytes) {
    fixture.componentInstance.bytes.set(bytes);
    await settle(fixture);
  }
  const element: HTMLElement = fixture.nativeElement;
  document.body.append(element);
  // jsdom has no `contentEditable` property, so the engine's editable pages get no attribute
  // there and cannot take the focus. A tab stop stands in for it.
  const pages = () => element.querySelector<HTMLElement>('.docx-pages')!;
  pages().tabIndex = -1;
  const update = async (action: () => unknown = () => {}) => {
    action();
    await settle(fixture, 2);
  };
  return {
    fixture,
    element,
    update,
    editor: () => fixture.componentInstance.editor(),
    pages,
  };
}

/** A pointerdown from the target's own window. jsdom has no PointerEvent. */
const pointerDown = (target: EventTarget, cancelable = false): Event => {
  const view = ((target as Node).ownerDocument ?? (target as Document)).defaultView!;
  const event = new view.MouseEvent('pointerdown', { bubbles: true, cancelable });
  target.dispatchEvent(event);
  return event;
};

const keydown = (target: EventTarget, key: string) => {
  const event = new KeyboardEvent('keydown', { key, bubbles: true, cancelable: true });
  target.dispatchEvent(event);
  return event;
};

beforeEach(() => configureZoneless());
afterEach(() => document.body.replaceChildren());

describe('named pickers', () => {
  for (const slot of ['font.family', 'styles.style', 'zoom.level']) {
    test(`${slot}: options support arrows, Escape, and focus return`, async () => {
      const { element, update } = await mount();
      const root = element.querySelector<HTMLElement>(`[data-slot="${slot}"]`)!;
      const trigger = root.querySelector<HTMLButtonElement>('[aria-haspopup="listbox"]')!;
      expect(trigger.disabled).toBe(false);
      const visible = trigger.textContent!.replace('▾', '').trim();
      const name = trigger.getAttribute('aria-label') ?? trigger.textContent!;
      expect(name).toContain(visible);
      await update(() => {
        trigger.focus();
        trigger.click();
      });
      const list = root.querySelector<HTMLElement>('[role="listbox"]')!;
      expect(list).not.toBeNull();
      expect(list.getAttribute('aria-label')?.length).toBeGreaterThan(0);
      const options = [...list.querySelectorAll<HTMLElement>('[role="option"]')];
      expect(options.length).toBeGreaterThan(1);
      expect(options.every((option) => option.tabIndex === -1)).toBe(true);
      await update(() => keydown(trigger, 'ArrowDown'));
      expect(document.activeElement).toBe(options[0]);
      await update(() => keydown(options[0]!, 'End'));
      expect(document.activeElement).toBe(options.at(-1));
      await update(() => keydown(options.at(-1)!, 'Escape'));
      expect(root.querySelector('[role="listbox"]')).toBeNull();
      expect(document.activeElement).toBe(trigger);
    });
  }

  test('the font size input names its open list', async () => {
    const { element, update } = await mount();
    const size = element.querySelector<HTMLInputElement>('[data-slot="font.size"] input')!;
    await update(() => size.focus());
    const list = element.querySelector('[data-slot="font.size"] [role="listbox"]')!;
    expect(list.id.length).toBeGreaterThan(0);
    expect(size.getAttribute('aria-controls')).toBe(list.id);
  });

  test('a typed font size applies on Enter, and the steppers follow the presets', async () => {
    const { element, update, editor } = await mount();
    await update(() => editor().surface!.selectAll());
    const size = element.querySelector<HTMLInputElement>('[data-slot="font.size"] input')!;
    await update(() => size.focus());
    size.value = '13';
    await update(() => size.dispatchEvent(new Event('input', { bubbles: true })));
    await update(() => keydown(size, 'Enter'));
    expect(editor().snapshot().formatting?.fontSizePt).toBe(13);
    const increase = element.querySelector<HTMLButtonElement>(
      `[data-slot="font.size"] [aria-label="${en.fontSize.increase}"]`
    )!;
    await update(() => increase.click());
    expect(editor().snapshot().formatting?.fontSizePt).toBe(14);
  });

  test('Escape closes a picker opened by a click while the focus stays in the pages', async () => {
    const { element, update, pages } = await mount();
    for (const slot of ['font.family', 'styles.style', 'zoom.level']) {
      const root = element.querySelector<HTMLElement>(`[data-slot="${slot}"]`)!;
      const trigger = root.querySelector<HTMLButtonElement>('[aria-haspopup="listbox"]')!;
      await update(() => {
        pages().focus();
        trigger.click();
      });
      expect(root.querySelector('[role="listbox"]')).not.toBeNull();
      let escape!: KeyboardEvent;
      await update(() => (escape = keydown(pages(), 'Escape')));
      expect(root.querySelector('[role="listbox"]')).toBeNull();
      // The key came from this editor, so the surface does not also act on it.
      expect(escape.defaultPrevented).toBe(true);
      expect(document.activeElement).toBe(pages());
    }
    // Escape in another field, such as a dialog or the find bar, still reaches that field.
    const field = document.createElement('input');
    element.append(field);
    let fieldSawEscape = false;
    field.addEventListener('keydown', () => (fieldSawEscape = true));
    const root = element.querySelector<HTMLElement>('[data-slot="font.family"]')!;
    await update(() => root.querySelector<HTMLButtonElement>('[aria-haspopup="listbox"]')!.click());
    expect(root.querySelector('[role="listbox"]')).not.toBeNull();
    let fieldEscape!: KeyboardEvent;
    await update(() => {
      field.focus();
      fieldEscape = keydown(field, 'Escape');
    });
    expect(fieldSawEscape).toBe(true);
    // A key from outside the editor is left alone: the picker closes, the key goes on.
    expect(fieldEscape.defaultPrevented).toBe(false);
    expect(root.querySelector('[role="listbox"]')).toBeNull();
  });

  test('Escape closes a picker before the header scope or the format painter', async () => {
    const { element, update, editor, pages } = await mount(headerDocx());
    const root = element.querySelector<HTMLElement>('[data-slot="font.family"]')!;
    const trigger = root.querySelector<HTMLButtonElement>('[aria-haspopup="listbox"]')!;
    const escape = async () => {
      let event!: KeyboardEvent;
      await update(() => (event = keydown(pages(), 'Escape')));
      return event;
    };
    expect(editor().exec({ type: 'editHeaderFooter', position: 'header' }).ok).toBe(true);
    await update(() => {
      pages().focus();
      trigger.click();
    });
    expect(root.querySelector('[role="listbox"]')).not.toBeNull();
    expect((await escape()).defaultPrevented).toBe(true);
    expect(root.querySelector('[role="listbox"]')).toBeNull();
    expect(editor().surface!.activeScope().kind).toBe('headerFooter');
    // With no picker open, the next Escape reaches the surface and leaves the header.
    await escape();
    expect(editor().surface!.activeScope().kind).toBe('body');

    expect(editor().surface!.formatPainter.press()).toBe(true);
    await update(() => trigger.click());
    expect(root.querySelector('[role="listbox"]')).not.toBeNull();
    await escape();
    expect(root.querySelector('[role="listbox"]')).toBeNull();
    expect(editor().surface!.formatPainter.state().mode).not.toBe('off');
    await escape();
    expect(editor().surface!.formatPainter.state().mode).toBe('off');
  });

  for (const slot of ['font.family', 'styles.style', 'zoom.level', 'font.size']) {
    test(`${slot}: a canceled press on the pages closes the popup, and a press inside keeps it`, async () => {
      const { element, update } = await mount();
      const root = element.querySelector<HTMLElement>(`[data-slot="${slot}"]`)!;
      const trigger = root.querySelector<HTMLElement>('[aria-haspopup]')!;
      await update(() => {
        trigger.focus();
        if (slot !== 'font.size') trigger.click();
      });
      const popup = () => root.querySelector<HTMLElement>('[role="listbox"], [role="dialog"]');
      expect(popup()).not.toBeNull();
      const item = popup()!.querySelector<HTMLElement>('button, [role="option"]')!;
      await update(() => pointerDown(item));
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

  test('picking a style applies it, and the trigger shows it', async () => {
    const { element, update, editor } = await mount();
    const root = element.querySelector<HTMLElement>('[data-slot="styles.style"]')!;
    const trigger = root.querySelector<HTMLButtonElement>('[aria-haspopup="listbox"]')!;
    expect(trigger.textContent!.trim()).toBe(en.styles.normal);
    await update(() => trigger.click());
    const heading = [...root.querySelectorAll<HTMLElement>('[role="option"]')].find((option) =>
      option.textContent!.includes('1')
    )!;
    await update(() => heading.click());
    expect(editor().snapshot().formatting?.styleId).toBe('Heading1');
    expect(trigger.textContent!.trim()).toBe(heading.querySelector('span')!.textContent);
  });
});

describe('editing mode', () => {
  test('Escape from a host input closes the menu and keeps its default', async () => {
    const { element, update } = await mount(headerDocx());
    const outside = document.createElement('input');
    document.body.append(outside);
    const trigger = element.querySelector<HTMLButtonElement>(
      '[data-testid="editing-mode-trigger"]'
    )!;
    await update(() => trigger.click());
    expect(trigger.getAttribute('aria-expanded')).toBe('true');
    let escape!: KeyboardEvent;
    await update(() => (escape = keydown(outside, 'Escape')));
    expect(trigger.getAttribute('aria-expanded')).toBe('false');
    expect(escape.defaultPrevented).toBe(false);
  });

  test('the open menu focuses the checked mode, and a refused mode is disabled', async () => {
    const { element, update } = await mount();
    const trigger = element.querySelector<HTMLButtonElement>(
      '[data-testid="editing-mode-trigger"]'
    )!;
    expect(trigger.getAttribute('data-mode')).toBe('editing');
    await update(() => trigger.click());
    expect(document.activeElement?.getAttribute('data-testid')).toBe('editing-mode-editing');
    // Without the review module, the engine refuses Suggesting.
    const suggesting = element.querySelector<HTMLButtonElement>(
      '[data-testid="editing-mode-suggesting"]'
    )!;
    expect(suggesting.disabled).toBe(true);
    expect(suggesting.title.length).toBeGreaterThan(0);
    await update(() =>
      element.querySelector<HTMLButtonElement>('[data-testid="editing-mode-viewing"]')!.click()
    );
    expect(trigger.getAttribute('data-mode')).toBe('viewing');
  });
});

const REVIEW: EditorModule = {
  id: 'review',
  review: {
    displayModes: ['all-markup', 'proposed', 'original'],
    collectReviewItems: () => [],
    revisionItemsOfParagraph: () => [],
  },
};

@Component({
  imports: PARTS,
  changeDetection: ChangeDetectionStrategy.OnPush,
  template: `<ng-container docxEditorRoot [document]="bytes()" [modules]="modules">
    <fieldset docxEditorToolbar></fieldset>
    <docx-editor-viewport><docx-editor-content /></docx-editor-viewport>
  </ng-container>`,
})
class ReviewHostComponent extends Host {
  readonly modules = [REVIEW];
}

describe('editing mode with the review module', () => {
  async function mountReview() {
    const fixture = await render(ReviewHostComponent);
    const element: HTMLElement = fixture.nativeElement;
    document.body.append(element);
    const button = (name: string) =>
      element.querySelector<HTMLButtonElement>(`[data-testid="editing-mode-${name}"]`)!;
    return { fixture, button, editor: () => fixture.componentInstance.editor() };
  }

  test('the open menu tracks author arrival and removal without changing the mode', async () => {
    const { fixture, button, editor } = await mountReview();
    expect(button('trigger').disabled).toBe(false);
    button('trigger').click();
    await settle(fixture);
    expect(button('suggesting').disabled).toBe(true);
    expect(button('suggesting').title).toBe(en.disabledReason.suggestingNeedsAuthor);
    editor().setAuthor('Ada');
    await settle(fixture);
    expect(button('suggesting').disabled).toBe(false);
    expect(button('suggesting').title).toBe('');
    editor().setAuthor(undefined);
    await settle(fixture);
    expect(button('suggesting').disabled).toBe(true);
    expect(button('viewing').disabled).toBe(false);
    button('viewing').click();
    await settle(fixture);
    expect(editor().getEditingMode()).toBe('viewing');
  });

  test('keyboard travel skips Suggesting when no author exists', async () => {
    const { fixture, button } = await mountReview();
    button('trigger').click();
    await settle(fixture);
    expect(document.activeElement).toBe(button('editing'));
    keydown(button('editing'), 'ArrowDown');
    expect(document.activeElement).toBe(button('viewing'));
  });
});

const LABEL_SOURCE = stylesDocx(
  paragraphStyle('Normal', 'Normal', ' w:default="1"') +
    paragraphStyle('Title', 'Title') +
    paragraphStyle('Ttulo1', 'heading 1') +
    paragraphStyle('Heading2', 'Heading 2') +
    paragraphStyle('NoSpacing', 'No Spacing') +
    paragraphStyle('ListParagraph', 'List Paragraph') +
    paragraphStyle('Callout', 'Callout'),
  paragraph('hello')
);

const CATALOGUE = {
  styles: {
    normal: 'Padrão',
    title: 'Título',
    headingLevel: 'Título {level}',
    noSpacing: 'Sem Espaçamento',
    listParagraph: 'Parágrafo da Lista',
  },
} as Translations;

@Component({
  imports: [...PARTS, DocxEditorI18nDirective],
  changeDetection: ChangeDetectionStrategy.OnPush,
  template: `<section [docxEditorI18n]="catalogue">
    <ng-container docxEditorRoot [document]="bytes()">
      <fieldset docxEditorToolbar></fieldset>
      <docx-editor-viewport><docx-editor-content /></docx-editor-viewport>
    </ng-container>
  </section>`,
})
class LabelsHostComponent extends Host {
  readonly catalogue = CATALOGUE;
}

describe('paragraph style labels', () => {
  test('built-in styles take the catalog label, and custom styles keep their name', async () => {
    const fixture = await render(LabelsHostComponent);
    fixture.componentInstance.bytes.set(LABEL_SOURCE);
    await settle(fixture);
    const element: HTMLElement = fixture.nativeElement;
    const trigger = element.querySelector<HTMLButtonElement>(
      '[data-slot="styles.style"] .docx-toolbar__style-trigger'
    )!;
    expect(trigger.textContent!.trim()).toBe('Padrão');
    trigger.click();
    await settle(fixture);
    const rows = [...element.querySelectorAll('[data-slot="styles.style"] [role="option"]')];
    expect(rows.map((row) => row.querySelector('span')?.textContent)).toEqual([
      'Padrão',
      'Título',
      'Título 1',
      'Título 2',
      'Sem Espaçamento',
      'Parágrafo da Lista',
      'Callout',
    ]);
  });
});

@Component({
  imports: [...PARTS, DocxEditorWindowComponent],
  changeDetection: ChangeDetectionStrategy.OnPush,
  template: `<ng-container docxEditorRoot [document]="bytes()">
    <docx-editor-window [open]="true">
      <ng-template>
        <fieldset docxEditorToolbar [preset]="false" [overflow]="false">
          <div docxToolbarFontFamily></div>
        </fieldset>
      </ng-template>
    </docx-editor-window>
    <docx-editor-viewport><docx-editor-content /></docx-editor-viewport>
  </ng-container>`,
})
class PopoutHostComponent extends Host {}

describe('a picker in another window', () => {
  test('closes on a press in the window it renders in, not in the opener', async () => {
    const frame = document.createElement('iframe');
    document.body.append(frame);
    const other = frame.contentWindow as Window & typeof globalThis;
    const original = window.open;
    window.open = (() => other) as typeof window.open;
    other.close = () => {};
    try {
      const fixture = TestBed.createComponent(PopoutHostComponent);
      await settle(fixture);
      const trigger = other.document.querySelector<HTMLElement>(
        '[data-slot="font.family"] [aria-haspopup]'
      )!;
      expect(trigger).not.toBeNull();
      trigger.click();
      await settle(fixture);
      expect(trigger.getAttribute('aria-expanded')).toBe('true');
      const press = async (target: EventTarget) => {
        pointerDown(target, true);
        await settle(fixture);
      };
      // A press in the opener's document is outside the editor's window: it says nothing.
      await press(document.body);
      expect(trigger.getAttribute('aria-expanded')).toBe('true');
      const outside = other.document.createElement('button');
      other.document.body.append(outside);
      await press(outside);
      expect(trigger.getAttribute('aria-expanded')).toBe('false');
      fixture.destroy();
    } finally {
      window.open = original;
      frame.remove();
    }
  });
});

describe('development warnings', () => {
  test('a warning prints once in development mode', () => {
    const warn = vi.spyOn(console, 'warn').mockImplementation(() => {});
    try {
      toolbarDevWarning('development build');
      toolbarDevWarning('development build');
      expect(warn).toHaveBeenCalledTimes(1);
      expect(String(warn.mock.calls[0]![0])).toBe('[docx-editor] development build');
    } finally {
      warn.mockRestore();
    }
  });
});
