// The base-layer claims of packages/vue/test/toolbar-direction.test.ts and
// packages/vue/test/form-field-i18n.test.ts. The toolbar claims port with the toolbar (phase 4).
import { ChangeDetectionStrategy, Component, signal, viewChild, viewChildren } from '@angular/core';
import { blankDocumentBytes } from '@docx-editor.dev/core/editor';
import { en, ptBR, type Translations } from '@docx-editor.dev/core/i18n';
import { configureZoneless, engineOf, render, settle } from '../testing/harness';
import { DocxEditorContentComponent } from './content.component';
import {
  DocxEditorI18nDirective,
  injectChromeDir,
  injectTranslation,
  provideDocxEditorI18n,
  type DocxEditorDirection,
} from './i18n';
import { DocxEditorRootDirective, type DocxEditorTranslate } from './root.directive';
import { DocxEditorViewportComponent } from './viewport.component';

/** Right-to-left catalogs as a host declares them. The package ships none. */
const ar: Translations = { _lang: 'ar', _dir: 'rtl' };
const he: Translations = { _lang: 'he' };

@Component({ selector: 'i18n-probe', template: '' })
class ProbeComponent {
  readonly chromeDir = injectChromeDir();
  readonly translation = injectTranslation();
}

@Component({
  imports: [
    DocxEditorI18nDirective,
    DocxEditorRootDirective,
    DocxEditorViewportComponent,
    DocxEditorContentComponent,
    ProbeComponent,
  ],
  changeDetection: ChangeDetectionStrategy.OnPush,
  template: `
    <section [docxEditorI18n]="catalog()" [docxEditorDir]="dir()">
      <ng-container docxEditorRoot [document]="bytes" [locale]="locale()" [translate]="translate()">
        <i18n-probe />
        <docx-editor-viewport><docx-editor-content /></docx-editor-viewport>
      </ng-container>
    </section>
  `,
})
class HostComponent {
  readonly bytes = blankDocumentBytes();
  readonly catalog = signal<Translations | undefined>(undefined);
  readonly dir = signal<DocxEditorDirection | undefined>(undefined);
  readonly locale = signal<string | undefined>(undefined);
  readonly translate = signal<DocxEditorTranslate | undefined>(undefined);
  readonly roots = viewChildren(DocxEditorRootDirective);
  readonly probe = viewChild.required(ProbeComponent);
}

async function mount() {
  const fixture = await render(HostComponent);
  const view = fixture.componentInstance;
  const host = fixture.nativeElement as HTMLElement;
  const dialogMount = () => host.querySelector('.docx-dialog-mount')!;
  const viewport = () => host.querySelector('docx-editor-viewport')!;
  return { fixture, view, dialogMount, viewport };
}

describe('chrome direction', () => {
  beforeEach(() => configureZoneless());

  it('mirrors the chrome for a catalog with _dir, and for a right-to-left _lang', async () => {
    const { fixture, view, dialogMount } = await mount();
    view.catalog.set(ar);
    await settle(fixture);
    expect(view.probe().chromeDir()).toBe('rtl');
    expect(dialogMount().getAttribute('dir')).toBe('rtl');
    view.catalog.set(he);
    await settle(fixture);
    expect(view.probe().chromeDir()).toBe('rtl');
    fixture.destroy();
  });

  it('lets the dir input decide over the catalog', async () => {
    const { fixture, view, dialogMount } = await mount();
    view.catalog.set(ptBR);
    view.dir.set('rtl');
    await settle(fixture);
    expect(view.probe().chromeDir()).toBe('rtl');
    expect(dialogMount().getAttribute('dir')).toBe('rtl');
    fixture.destroy();
  });

  it('leaves the host direction in charge for a left-to-right catalog', async () => {
    const { fixture, view, dialogMount } = await mount();
    view.catalog.set(ptBR);
    await settle(fixture);
    expect(view.probe().chromeDir()).toBeUndefined();
    expect(view.probe().translation.dir()).toBe('ltr');
    expect(dialogMount().hasAttribute('dir')).toBe(false);
    fixture.destroy();
  });
});

describe('chrome translations', () => {
  beforeEach(() => configureZoneless());

  it('translates the chrome and follows a catalog change without a new engine', async () => {
    const { fixture, view, viewport } = await mount();
    const editor = engineOf(view.roots()[0]!.context.editor());
    expect(viewport().getAttribute('aria-label')).toBe(en.editor.documentViewport);
    view.catalog.set(ptBR);
    await settle(fixture);
    expect(viewport().getAttribute('aria-label')).toBe(ptBR.editor!.documentViewport);
    expect(view.probe().translation.lang()).toBe(ptBR._lang);
    expect(view.roots()[0]!.context.editor()).toBe(editor);
    fixture.destroy();
  });

  it('merges a partial catalog over the outer provider, then over English', async () => {
    @Component({
      selector: 'nested-i18n',
      imports: [DocxEditorI18nDirective, ProbeComponent],
      providers: [provideDocxEditorI18n(ptBR)],
      template: `<div [docxEditorI18n]="partial"><i18n-probe /></div>`,
    })
    class NestedComponent {
      readonly partial: Translations = { editor: { documentViewport: 'Inner label' } };
      readonly probe = viewChild.required(ProbeComponent);
    }
    const fixture = await render(NestedComponent);
    const t = fixture.componentInstance.probe().translation.t();
    expect(t('editor.documentViewport')).toBe('Inner label');
    expect(t('toolbar.rightToLeft')).toBe(ptBR.toolbar!.rightToLeft);
    fixture.destroy();
  });
});

describe('engine labels and locale', () => {
  beforeEach(() => configureZoneless());

  it('keeps engine labels live with the catalog, and an explicit translate wins', async () => {
    const { fixture, view } = await mount();
    const editor = engineOf(view.roots()[0]!.context.editor());
    const setTranslate = vi.spyOn(editor, 'setTranslate');
    const catalog: Translations = { textFormField: { maxLength: 'Comprimento máximo' } };
    view.catalog.set(catalog);
    await settle(fixture);
    const latest = () => setTranslate.mock.calls.at(-1)![0]!;
    expect(latest()('textFormField.maxLength')).toBe('Comprimento máximo');

    view.translate.set((key) => (key === 'textFormField.maxLength' ? 'Explicit' : key));
    await settle(fixture);
    expect(latest()('textFormField.maxLength')).toBe('Explicit');
    expect(view.roots()[0]!.context.editor()).toBe(editor);
    fixture.destroy();
  });

  it('applies a locale change to the same engine', async () => {
    const { fixture, view } = await mount();
    const editor = engineOf(view.roots()[0]!.context.editor());
    const setLocale = vi.spyOn(editor, 'setLocale');
    view.locale.set('pl-PL');
    await settle(fixture);
    expect(setLocale).toHaveBeenLastCalledWith('pl-PL');
    expect(view.roots()[0]!.context.editor()).toBe(editor);
    fixture.destroy();
  });
});
