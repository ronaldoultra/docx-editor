// The claims of packages/vue/test/live-props.test.ts, field-results-prop.test.ts,
// content-lock-policy-prop.test.ts, and locked-control-navigation-prop.test.ts, and the
// packaged-editor claims of form-field-i18n.test.ts and toolbar-direction.test.ts.
import { ChangeDetectionStrategy, Component, signal, viewChild, type Type } from '@angular/core';
import { TestBed, type ComponentFixture } from '@angular/core/testing';
import type {
  ContentLockPolicy,
  DocxEditorInstance,
  FieldResultsMode,
  LockedControlNavigation,
} from '@docx-editor.dev/core/editor';
import { paragraphTextOf, serializeOoxmlPart } from '@docx-editor.dev/core/store';
import { ptBR, type Translations } from '@docx-editor.dev/core/i18n';
import { docx, paragraph, run } from '../testing/docx';
import { configureZoneless, engineOf, settle } from '../testing/harness';
import type { DocxEditorColorMode } from './color-mode';
import { DocxEditorContentComponent } from './content.component';
import { DocxEditorComponent, type DocxEditorMode } from './docx-editor.component';
import { DocxEditorRootDirective, type DocxEditorTranslate } from './root.directive';
import { DocxEditorViewportComponent } from './viewport.component';

const SOURCE = docx(paragraph('hello world'));

/** "ab " then a MERGEFIELD whose saved result is "Name", then " cd". */
const FIELD_SOURCE = docx(
  `<w:p>${run('ab ')}<w:fldSimple w:instr=" MERGEFIELD Name "><w:r><w:t>Name</w:t></w:r>` +
    `</w:fldSimple>${run(' cd')}</w:p>`
);

/** "ABC", a control with locked content "TOKEN", then "DEF". */
const LOCKED_SOURCE = docx(
  '<w:p><w:r><w:t>ABC</w:t></w:r><w:sdt><w:sdtPr><w:tag w:val="t"/>' +
    '<w:lock w:val="sdtContentLocked"/></w:sdtPr><w:sdtContent><w:r><w:t>TOKEN</w:t></w:r>' +
    '</w:sdtContent></w:sdt><w:r><w:t>DEF</w:t></w:r></w:p>'
);

/** "aa", then a locked inline control "TOKEN" over offsets 2..7, then "zz". */
const INLINE_LOCKED_SOURCE = docx(
  `<w:p>${run('aa')}<w:sdt><w:sdtPr><w:alias w:val="T"/><w:lock w:val="sdtContentLocked"/>` +
    `<w:text/></w:sdtPr><w:sdtContent>${run('TOKEN')}</w:sdtContent></w:sdt>${run('zz')}</w:p>`
);

/** The signals that both hosts bind to their inputs. */
class Inputs {
  readonly bytes = signal<Uint8Array>(SOURCE);
  readonly author = signal<string | undefined>(undefined);
  readonly locale = signal<string | undefined>(undefined);
  readonly translate = signal<DocxEditorTranslate | undefined>(undefined);
  readonly zoom = signal<number | undefined>(undefined);
  readonly fieldResults = signal<FieldResultsMode | undefined>(undefined);
  readonly contentLockPolicy = signal<ContentLockPolicy | undefined>(undefined);
  readonly lockedControlNavigation = signal<LockedControlNavigation | undefined>(undefined);
}

interface EditorHost extends Inputs {
  editor(): DocxEditorInstance;
}

@Component({
  imports: [DocxEditorRootDirective, DocxEditorViewportComponent, DocxEditorContentComponent],
  changeDetection: ChangeDetectionStrategy.OnPush,
  template: `
    <ng-container
      docxEditorRoot
      [document]="bytes()"
      [mode]="mode()"
      [author]="author()"
      [locale]="locale()"
      [translate]="translate()"
      [zoom]="zoom()"
      [fieldResults]="fieldResults()"
      [contentLockPolicy]="contentLockPolicy()"
      [lockedControlNavigation]="lockedControlNavigation()"
    >
      <docx-editor-viewport><docx-editor-content /></docx-editor-viewport>
    </ng-container>
  `,
})
class RootHostComponent extends Inputs implements EditorHost {
  readonly mode = signal<DocxEditorMode | undefined>(undefined);
  readonly root = viewChild.required(DocxEditorRootDirective);
  editor(): DocxEditorInstance {
    return engineOf(this.root().context.editor());
  }
}

@Component({
  imports: [DocxEditorComponent],
  changeDetection: ChangeDetectionStrategy.OnPush,
  template: `
    <docx-editor
      [document]="bytes()"
      [mode]="mode()"
      [author]="author()"
      [locale]="locale()"
      [i18n]="i18n()"
      [translate]="translate()"
      [colorMode]="colorMode()"
      [zoom]="zoom()"
      [fieldResults]="fieldResults()"
      [contentLockPolicy]="contentLockPolicy()"
      [lockedControlNavigation]="lockedControlNavigation()"
    />
  `,
})
class PackagedHostComponent extends Inputs implements EditorHost {
  readonly mode = signal<DocxEditorMode>('edit');
  readonly i18n = signal<Translations | undefined>(undefined);
  readonly colorMode = signal<DocxEditorColorMode>('light');
  readonly packaged = viewChild.required(DocxEditorComponent);
  editor(): DocxEditorInstance {
    return engineOf(this.packaged().getEditor());
  }
}

/** Create a host, set its inputs before the first render, and wait for its engine. */
async function mount<T>(type: Type<T>, setup: (host: T) => void = () => {}) {
  const fixture: ComponentFixture<T> = TestBed.createComponent(type);
  setup(fixture.componentInstance);
  await settle(fixture);
  return { fixture, host: fixture.componentInstance, element: fixture.nativeElement as Element };
}

function select(editor: DocxEditorInstance, start: number, end: number): void {
  const paragraphId = editor.surface!.session.paragraphIds()[0]!;
  editor.surface!.setSelection({
    anchor: { paragraphId, offset: start },
    head: { paragraphId, offset: end },
  });
}

/** Collect the console warnings that mention `text` while `body` runs. */
async function warningsAbout(text: string, body: () => Promise<void>): Promise<string[]> {
  const warn = vi.spyOn(console, 'warn').mockImplementation(() => {});
  try {
    await body();
    return warn.mock.calls.map((call) => String(call[0])).filter((m) => m.includes(text));
  } finally {
    warn.mockRestore();
  }
}

const HOSTS: ReadonlyArray<readonly [string, Type<EditorHost>]> = [
  ['root', RootHostComponent],
  ['packaged', PackagedHostComponent],
];

describe('live inputs on the root', () => {
  beforeEach(() => configureZoneless());

  it('applies a changed author without a new engine', async () => {
    const { fixture, host } = await mount(RootHostComponent, (h) => h.author.set('Initial'));
    const editor = host.editor();
    select(editor, 0, 1);
    expect(editor.exec({ type: 'proposeReplacement', replaceWith: 'X' })).toMatchObject({
      ok: true,
      changed: true,
    });
    host.author.set('Updated');
    await settle(fixture);
    expect(host.editor()).toBe(editor);
    expect(editor.getConfiguredAuthor()).toBe('Updated');
    select(editor, 1, 2);
    expect(editor.exec({ type: 'proposeReplacement', replaceWith: 'Y' })).toMatchObject({
      ok: true,
      changed: true,
    });
    const xml = serializeOoxmlPart(editor.surface!.session.part());
    expect(xml.match(/w:author="Initial"/g)).toHaveLength(2);
    expect(xml.match(/w:author="Updated"/g)).toHaveLength(2);
    fixture.destroy();
  });

  it('applies mode, translate, and locale changes in place', async () => {
    const { fixture, host } = await mount(RootHostComponent, (h) => {
      h.mode.set('edit');
      h.translate.set((key) => `initial:${key}`);
      h.locale.set('en');
    });
    const editor = host.editor();
    const setTranslate = vi.spyOn(editor, 'setTranslate');
    const setLocale = vi.spyOn(editor, 'setLocale');
    host.mode.set('view');
    host.translate.set((key) => `updated:${key}`);
    host.locale.set('de-DE');
    await settle(fixture);
    expect(host.editor()).toBe(editor);
    expect(editor.getEditingMode()).toBe('viewing');
    expect(setTranslate.mock.calls.at(-1)![0]!('probe')).toBe('updated:probe');
    expect(setLocale).toHaveBeenLastCalledWith('de-DE');
    fixture.destroy();
  });

  it('does not restore the host mode after a reader change when another input changes', async () => {
    const { fixture, host } = await mount(RootHostComponent, (h) => {
      h.mode.set('edit');
      h.author.set('First');
    });
    const editor = host.editor();
    expect(editor.exec({ type: 'setEditingMode', mode: 'viewing' })).toMatchObject({ ok: true });
    host.author.set('Second');
    await settle(fixture);
    expect(host.editor()).toBe(editor);
    expect(editor.getEditingMode()).toBe('viewing');
    fixture.destroy();
  });
});

describe('live inputs on <docx-editor>', () => {
  beforeEach(() => configureZoneless());

  it('follows a catalog change, then an explicit translate, with the same engine', async () => {
    const catalog = (label: string): Translations => ({ image: { pendingResource: label } });
    const { fixture, host } = await mount(PackagedHostComponent, (h) =>
      h.i18n.set(catalog('First'))
    );
    const editor = host.editor();
    const setTranslate = vi.spyOn(editor, 'setTranslate');
    const latest = () => setTranslate.mock.calls.at(-1)![0]!;
    host.i18n.set(catalog('Second'));
    await settle(fixture);
    expect(host.editor()).toBe(editor);
    expect(latest()('image.pendingResource')).toBe('Second');

    const catalogResolver = latest();
    host.translate.set((key) => `custom:${key}`);
    await settle(fixture);
    expect(host.editor()).toBe(editor);
    expect(latest()).not.toBe(catalogResolver);
    expect(latest()('image.pendingResource')).toBe('custom:image.pendingResource');
    fixture.destroy();
  });

  it('keeps form labels live with the catalog without a new engine', async () => {
    const { fixture, host } = await mount(PackagedHostComponent, (h) => h.locale.set('pl-PL'));
    const editor = host.editor();
    const setTranslate = vi.spyOn(editor, 'setTranslate');
    host.i18n.set({ textFormField: { maxLength: 'Maksymalna długość' } });
    await settle(fixture);
    expect(setTranslate.mock.calls.at(-1)![0]!('textFormField.maxLength')).toBe(
      'Maksymalna długość'
    );
    expect(host.editor()).toBe(editor);
    fixture.destroy();
  });

  it('does not publish a selection change for a new but equal translate function', async () => {
    const { fixture, host } = await mount(PackagedHostComponent, (h) =>
      h.translate.set((key) => `same:${key}`)
    );
    const editor = host.editor();
    let changes = 0;
    editor.on('selectionChange', () => changes++);
    host.translate.set((key) => `same:${key}`);
    await settle(fixture);
    expect(host.editor()).toBe(editor);
    expect(changes).toBe(0);
    fixture.destroy();
  });

  it('applies colorMode and zoom changes without a new engine', async () => {
    const { fixture, host, element } = await mount(PackagedHostComponent);
    const editor = host.editor();
    expect(element.querySelector('.docx-editor.dark')).toBeNull();
    host.colorMode.set('dark');
    host.zoom.set(1.5);
    await settle(fixture);
    expect(element.querySelector('.docx-editor.dark')).not.toBeNull();
    expect(editor.getZoom()).toBe(1.5);
    expect(host.editor()).toBe(editor);
    fixture.destroy();
  });

  it('mirrors its chrome for a right-to-left catalog, and leaves it for a left-to-right one', async () => {
    const direction = async (i18n: Translations) => {
      const { fixture, element } = await mount(PackagedHostComponent, (h) => h.i18n.set(i18n));
      const dir = element.querySelector('.docx-editor')!.getAttribute('dir');
      fixture.destroy();
      return dir;
    };
    expect(await direction({ _lang: 'ar', _dir: 'rtl' })).toBe('rtl');
    expect(await direction({ _lang: 'he' })).toBe('rtl');
    expect(await direction(ptBR)).toBeNull();
  });

  it('destroys the engine for a null document, and creates one for new bytes', async () => {
    const { fixture, host } = await mount(PackagedHostComponent);
    const editor = host.editor();
    const destroy = vi.spyOn(editor, 'destroy');
    host.bytes.set(null as unknown as Uint8Array);
    await settle(fixture);
    expect(destroy).toHaveBeenCalled();
    expect(host.packaged().getEditor()).toBeNull();
    host.bytes.set(SOURCE);
    await settle(fixture);
    expect(host.editor()).not.toBe(editor);
    fixture.destroy();
  });
});

describe.each(HOSTS)('creation-time inputs on the %s', (_name, type) => {
  beforeEach(() => configureZoneless());

  it.each(['editable', undefined] as const)('fieldResults %s', async (mode) => {
    const { fixture, host } = await mount(type, (h) => {
      h.bytes.set(FIELD_SOURCE);
      h.fieldResults.set(mode);
    });
    const surface = host.editor().surface!;
    const paragraphId = surface.session.paragraphIds()[0]!;
    surface.setSelection({
      anchor: { paragraphId, offset: 5 },
      head: { paragraphId, offset: 5 },
    });
    surface.type('X');
    expect(paragraphTextOf(surface.session.part(), paragraphId, { fieldResults: 'editable' })).toBe(
      mode === 'editable' ? 'ab NaXme cd' : 'ab Name Xcd'
    );
    fixture.destroy();
  });

  it.each([{ allowFormattingInLockedContent: true }, undefined])(
    'contentLockPolicy %o',
    async (policy) => {
      const { fixture, host } = await mount(type, (h) => {
        h.bytes.set(LOCKED_SOURCE);
        h.contentLockPolicy.set(policy);
      });
      const editor = host.editor();
      editor.surface!.selectAll();
      const result = editor.exec({ type: 'toggleMark', mark: 'bold' });
      expect(result.ok && 'changed' in result && result.changed === true).toBe(
        policy !== undefined
      );
      expect(editor.surface!.session.bodyText()).toBe('ABCTOKENDEF');
      fixture.destroy();
    }
  );

  it.each(['atomic', undefined] as const)('lockedControlNavigation %s', async (navigation) => {
    const { fixture, host } = await mount(type, (h) => {
      h.bytes.set(INLINE_LOCKED_SOURCE);
      h.lockedControlNavigation.set(navigation);
    });
    expect(stepRightFromLeftEdge(host.editor())).toBe(navigation === 'atomic' ? 7 : 3);
    fixture.destroy();
  });
});

/** Put the caret on the left edge of the locked control, press ArrowRight, read the head. */
function stepRightFromLeftEdge(editor: DocxEditorInstance): number {
  const surface = editor.surface!;
  const paragraphId = surface.session.paragraphIds()[0]!;
  surface.setSelection({ anchor: { paragraphId, offset: 2 }, head: { paragraphId, offset: 2 } });
  surface.navigate('right');
  return surface.state().selection.head.offset;
}

// The warnings print once per option, so each option has one warning case.
describe('a creation-time input that changes later', () => {
  beforeEach(() => configureZoneless());

  it('compares contentLockPolicy by value, ignores a change, and warns once', async () => {
    let first: DocxEditorInstance | undefined;
    const warnings = await warningsAbout('contentLockPolicy changed', async () => {
      const { fixture, host } = await mount(RootHostComponent, (h) => {
        h.bytes.set(LOCKED_SOURCE);
        h.contentLockPolicy.set({ allowFormattingInLockedContent: true });
      });
      first = host.editor();
      host.contentLockPolicy.set({ allowFormattingInLockedContent: true });
      await settle(fixture);
      expect(console.warn).not.toHaveBeenCalled();
      host.contentLockPolicy.set({ allowFormattingInLockedContent: false });
      await settle(fixture);
      host.contentLockPolicy.set(undefined);
      await settle(fixture);
      expect(host.editor()).toBe(first);
      first.surface!.selectAll();
      expect(first.exec({ type: 'toggleMark', mark: 'bold' })).toMatchObject({ changed: true });
      fixture.destroy();
    });
    expect(warnings).toHaveLength(1);
  });

  it('ignores a lockedControlNavigation change, and warns once for a real change', async () => {
    const warnings = await warningsAbout('lockedControlNavigation changed', async () => {
      const { fixture, host } = await mount(RootHostComponent, (h) =>
        h.bytes.set(INLINE_LOCKED_SOURCE)
      );
      const first = host.editor();
      for (const navigation of [undefined, 'characters', 'characters', undefined] as const) {
        host.lockedControlNavigation.set(navigation);
        await settle(fixture);
      }
      expect(console.warn).not.toHaveBeenCalled();
      for (const navigation of ['atomic', 'characters', 'atomic'] as const) {
        host.lockedControlNavigation.set(navigation);
        await settle(fixture);
      }
      expect(host.editor()).toBe(first);
      expect(stepRightFromLeftEdge(first)).toBe(3);
      fixture.destroy();
    });
    expect(warnings).toHaveLength(1);
  });

  it('ignores a fieldResults change, and warns once', async () => {
    const warnings = await warningsAbout('fieldResults changed', async () => {
      const { fixture, host } = await mount(RootHostComponent, (h) => h.bytes.set(FIELD_SOURCE));
      const first = host.editor();
      host.fieldResults.set('atomic');
      await settle(fixture);
      expect(console.warn).not.toHaveBeenCalled();
      host.fieldResults.set('editable');
      await settle(fixture);
      host.fieldResults.set(undefined);
      await settle(fixture);
      expect(host.editor()).toBe(first);
      fixture.destroy();
    });
    expect(warnings).toHaveLength(1);
  });
});
