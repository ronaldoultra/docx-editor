// Task 4.5: the claims of packages/vue/test/toolbar-link-platform.test.ts, and the direction
// button claims of toolbar-direction.test.ts.
import { ChangeDetectionStrategy, Component, Directive, signal, viewChild } from '@angular/core';
import type { DocxEditorInstance } from '@docx-editor.dev/core/editor';
import { serializeOoxmlPart } from '@docx-editor.dev/core/store';
import { en, type Translations } from '@docx-editor.dev/core/i18n';
import { afterEach, beforeEach, describe, expect, test } from 'vitest';
import { docx } from '../../testing/docx';
import { configureZoneless, engineOf, render, settle } from '../../testing/harness';
import { DocxEditorContentComponent } from '../content.component';
import { DocxEditorI18nDirective } from '../i18n';
import { DocxEditorRootDirective } from '../root.directive';
import { DocxEditorViewportComponent } from '../viewport.component';
import { DocxToolbarButtonComponent } from './toolbar-button.component';
import { DocxEditorToolbarComponent } from './toolbar.component';

const p = (text: string, pPr = ''): string =>
  `<w:p>${pPr ? `<w:pPr>${pPr}</w:pPr>` : ''}<w:r><w:t>${text}</w:t></w:r></w:p>`;

const PARTS = [
  DocxEditorRootDirective,
  DocxEditorViewportComponent,
  DocxEditorContentComponent,
  DocxEditorToolbarComponent,
  DocxToolbarButtonComponent,
  DocxEditorI18nDirective,
];

@Directive()
abstract class Host {
  readonly bytes = signal<Uint8Array>(docx(p('alpha')));
  readonly t = signal<((key: string) => string | undefined) | undefined>(undefined);
  readonly root = viewChild.required(DocxEditorRootDirective);

  editor(): DocxEditorInstance {
    return engineOf(this.root().context.editor());
  }
}

@Component({
  imports: PARTS,
  changeDetection: ChangeDetectionStrategy.OnPush,
  template: `<ng-container docxEditorRoot [document]="bytes()">
    <fieldset docxEditorToolbar [preset]="false" [overflow]="false" [t]="t()">
      <button docxToolbarButton="text.bold"></button>
      <button docxToolbarButton="text.link"></button>
      <button docxToolbarButton="direction.ltr"></button>
      <button docxToolbarButton="direction.rtl"></button>
    </fieldset>
    <docx-editor-viewport><docx-editor-content /></docx-editor-viewport>
  </ng-container>`,
})
class ExplicitHostComponent extends Host {}

@Component({
  imports: PARTS,
  changeDetection: ChangeDetectionStrategy.OnPush,
  template: `<section [docxEditorI18n]="catalog()">
    <ng-container docxEditorRoot [document]="bytes()">
      <fieldset docxEditorToolbar></fieldset>
      <docx-editor-viewport><docx-editor-content /></docx-editor-viewport>
    </ng-container>
  </section>`,
})
class DefaultHostComponent extends Host {
  readonly catalog = signal<Translations | undefined>(undefined);
}

const platformDescriptor = Object.getOwnPropertyDescriptor(navigator, 'platform');
const agentDataDescriptor = Object.getOwnPropertyDescriptor(navigator, 'userAgentData');

function platform(value: string): void {
  Object.defineProperty(navigator, 'platform', { value, configurable: true });
  Object.defineProperty(navigator, 'userAgentData', { value: undefined, configurable: true });
}

async function mount<T extends Host>(component: new () => T, setup?: (host: T) => void) {
  const fixture = await render(component);
  setup?.(fixture.componentInstance);
  await settle(fixture);
  const element: HTMLElement = fixture.nativeElement;
  const update = async (action: () => unknown = () => {}) => {
    action();
    await settle(fixture, 2);
  };
  const editor = () => fixture.componentInstance.editor();
  const caretIn = (index: number) =>
    update(() => {
      const paragraphId = editor().surface!.session.paragraphIds()[index]!;
      editor().surface!.setSelection({
        anchor: { paragraphId, offset: 1 },
        head: { paragraphId, offset: 1 },
      });
    });
  const button = (label: string) => {
    const found = element.querySelector<HTMLButtonElement>(
      `button[aria-label=${JSON.stringify(label)}]`
    );
    if (!found) throw new Error(`no toolbar button labelled ${label}`);
    return found;
  };
  return { fixture, element, update, editor, caretIn, button };
}

beforeEach(() => configureZoneless());
afterEach(() => {
  if (platformDescriptor) Object.defineProperty(navigator, 'platform', platformDescriptor);
  else Reflect.deleteProperty(navigator, 'platform');
  if (agentDataDescriptor) Object.defineProperty(navigator, 'userAgentData', agentDataDescriptor);
  else Reflect.deleteProperty(navigator, 'userAgentData');
});

describe('the link control', () => {
  const link = (element: HTMLElement) =>
    element.querySelector<HTMLButtonElement>('[data-slot="text.link"]')!;

  /** Mount the explicit toolbar with a link UI registered, as the hyperlink popup does. */
  async function mountWithLinkUi(setup?: (host: ExplicitHostComponent) => void) {
    const view = await mount(ExplicitHostComponent, setup);
    let requests = 0;
    await view.update(() => view.editor().setHyperlinkChrome({ onRequest: () => (requests += 1) }));
    return { ...view, requests: () => requests };
  }

  test('uses the same Mac modifier as the other buttons', async () => {
    platform('MacIntel');
    const { element, update, requests } = await mountWithLinkUi();
    expect(element.querySelector('[data-slot="text.bold"]')!.getAttribute('title')).toBe(
      'Bold (⌘+B)'
    );
    expect(link(element).getAttribute('title')).toBe('Insert link (⌘+K)');
    expect(link(element).getAttribute('aria-label')).toBe('Insert link (⌘+K)');
    // The button sends the same request as Ctrl/Cmd+K to the registered link UI.
    await update(() => link(element).click());
    expect(requests()).toBe(1);
  });

  test('keeps Ctrl on Windows', async () => {
    platform('Win32');
    const { element } = await mountWithLinkUi();
    expect(link(element).getAttribute('title')).toBe('Insert link (Ctrl+K)');
    expect(link(element).getAttribute('aria-label')).toBe('Insert link (Ctrl+K)');
  });

  test('normalizes the host translation', async () => {
    platform('MacIntel');
    const { element } = await mountWithLinkUi((host) =>
      host.t.set((key) =>
        key === 'formattingBar.insertLinkShortcut' ? 'Lien (Ctrl+K)' : undefined
      )
    );
    expect(link(element).getAttribute('title')).toBe('Lien (⌘+K)');
    expect(link(element).getAttribute('aria-label')).toBe('Lien (⌘+K)');
  });

  test('without a link UI, the control is disabled with the engine reason', async () => {
    const { element } = await mount(ExplicitHostComponent);
    expect(link(element).disabled).toBe(true);
    expect(link(element).title).toBe(en.disabledReason.noLinkView);
  });
});

describe('the paragraph direction buttons', () => {
  test('Right-to-left writes w:bidi and reads back pressed', async () => {
    const { update, editor, caretIn, button } = await mount(ExplicitHostComponent);
    await caretIn(0);
    const rtl = button(en.toolbar.rightToLeft);
    const ltr = button(en.toolbar.leftToRight);
    expect(rtl.disabled).toBe(false);
    expect(rtl.getAttribute('aria-pressed')).toBe('false');
    expect(ltr.getAttribute('aria-pressed')).toBe('true');
    await update(() => rtl.click());
    expect(editor().snapshot().formatting?.direction).toBe('rtl');
    expect(serializeOoxmlPart(editor().surface!.session.part())).toContain('<w:bidi');
    expect(rtl.getAttribute('aria-pressed')).toBe('true');
    expect(ltr.getAttribute('aria-pressed')).toBe('false');
  });

  test('the pressed state follows the caret between paragraphs of each direction', async () => {
    const { caretIn, button } = await mount(ExplicitHostComponent, (host) =>
      host.bytes.set(docx(p('one', '<w:bidi/>') + p('two')))
    );
    const rtl = button(en.toolbar.rightToLeft);
    const ltr = button(en.toolbar.leftToRight);
    await caretIn(0);
    expect(rtl.getAttribute('aria-pressed')).toBe('true');
    expect(ltr.getAttribute('aria-pressed')).toBe('false');
    await caretIn(1);
    expect(rtl.getAttribute('aria-pressed')).toBe('false');
    expect(ltr.getAttribute('aria-pressed')).toBe('true');
  });

  test('the default toolbar omits them for a left-to-right interface language', async () => {
    const { element } = await mount(DefaultHostComponent);
    expect(element.querySelector(`button[aria-label="${en.toolbar.leftToRight}"]`)).toBeNull();
    expect(element.querySelector(`button[aria-label="${en.toolbar.rightToLeft}"]`)).toBeNull();
  });

  test('a right-to-left interface language shows them after alignment', async () => {
    const { element, button } = await mount(DefaultHostComponent, (host) =>
      host.catalog.set({ _lang: 'ar', _dir: 'rtl' })
    );
    const rtl = button(en.toolbar.rightToLeft);
    const ltr = button(en.toolbar.leftToRight);
    const alignment = element.querySelector('[data-slot="alignment"]')!;
    expect(alignment.compareDocumentPosition(ltr) & Node.DOCUMENT_POSITION_FOLLOWING).toBeTruthy();
    expect(ltr.compareDocumentPosition(rtl) & Node.DOCUMENT_POSITION_FOLLOWING).toBeTruthy();
    expect(element.querySelector('[data-testid="docx-toolbar"]')!.getAttribute('dir')).toBe('rtl');
  });
});
