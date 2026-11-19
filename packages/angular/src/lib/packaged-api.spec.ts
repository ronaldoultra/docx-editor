// The outputs and methods of <docx-editor>, with the claims of
// packages/vue/test/lifecycle-regression.test.ts, sugar-host.test.ts, and
// paginated-form-save.test.ts that need no chrome.
import {
  ChangeDetectionStrategy,
  Component,
  ElementRef,
  inject,
  signal,
  viewChild,
  type Type,
} from '@angular/core';
import { TestBed } from '@angular/core/testing';
import type {
  DocumentChange,
  DocumentOpenFailure,
  Editor,
  EditorError,
  EditorFontError,
  EditorSnapshot,
} from '@docx-editor.dev/core/contracts/editor';
import { LOADING_SNAPSHOT, type DocxEditorInstance } from '@docx-editor.dev/core/editor';
import { strFromU8, unzipSync } from 'fflate';
import { docx, formFieldDocx, paragraph } from '../testing/docx';
import { configureZoneless, engineOf, settle } from '../testing/harness';
import { DocxEditorContentComponent } from './content.component';
import { DocxEditorComponent } from './docx-editor.component';
import {
  DocxEditorRootDirective,
  type DocxEditorFonts,
  type DocxEditorTranslate,
} from './root.directive';
import { DocxEditorViewportComponent } from './viewport.component';

const SOURCE = docx(paragraph('regression'));

@Component({
  imports: [DocxEditorComponent],
  changeDetection: ChangeDetectionStrategy.OnPush,
  template: `
    @if (shown()) {
      <docx-editor
        [document]="bytes()"
        [fonts]="fonts()"
        (ready)="ready.push($event)"
        (change)="changes.push($event)"
        (selectionChange)="selections.push($event)"
        (error)="errors.push($event)"
        (openFailure)="failures.push($event)"
        (fontError)="fontErrors.push($event)"
      />
    }
  `,
})
class PackagedHostComponent {
  readonly shown = signal(true);
  readonly bytes = signal<Uint8Array | undefined>(SOURCE);
  readonly fonts = signal<DocxEditorFonts | undefined>(undefined);
  readonly ready: Editor[] = [];
  readonly changes: DocumentChange[] = [];
  readonly selections: EditorSnapshot[] = [];
  readonly errors: EditorError[] = [];
  readonly failures: DocumentOpenFailure[] = [];
  readonly fontErrors: EditorFontError[] = [];
  readonly packaged = viewChild(DocxEditorComponent);
  editor(): DocxEditorInstance {
    return engineOf(this.packaged()?.getEditor());
  }
}

@Component({
  imports: [DocxEditorRootDirective, DocxEditorViewportComponent, DocxEditorContentComponent],
  changeDetection: ChangeDetectionStrategy.OnPush,
  template: `
    <span>{{ unrelated() }}</span>
    <ng-container
      docxEditorRoot
      [document]="document()"
      [translate]="translate()"
      (ready)="onReady($event)"
    >
      <docx-editor-viewport><docx-editor-content /></docx-editor-viewport>
    </ng-container>
  `,
})
class RootHostComponent {
  readonly document = signal<Uint8Array | 'blank'>(SOURCE);
  readonly translate = signal<DocxEditorTranslate | undefined>(undefined);
  readonly unrelated = signal(0);
  readonly ready: Editor[] = [];
  /** The number of painted pages when `ready` fired. */
  pagesAtReady = -1;
  private readonly element: Element = inject(ElementRef).nativeElement;

  onReady(editor: Editor): void {
    this.ready.push(editor);
    this.pagesAtReady = this.element.querySelectorAll('.docx-page').length;
  }
}

async function mount<T>(type: Type<T>, setup: (host: T) => void = () => {}) {
  const fixture = TestBed.createComponent(type);
  setup(fixture.componentInstance);
  await settle(fixture);
  return { fixture, host: fixture.componentInstance };
}

describe('<docx-editor> outputs', () => {
  beforeEach(() => configureZoneless());

  it('emits ready once, and change once for each edit', async () => {
    const { fixture, host } = await mount(PackagedHostComponent);
    expect(host.ready).toHaveLength(1);
    expect(host.ready[0]).toBe(host.editor());
    const before = host.changes.length;
    host.editor().exec({ type: 'insertText', text: '!' });
    await settle(fixture);
    expect(host.changes).toHaveLength(before + 1);
    fixture.destroy();
  });

  it('emits selectionChange with the snapshot', async () => {
    const { fixture, host } = await mount(PackagedHostComponent);
    const editor = host.editor();
    const paragraphId = editor.surface!.session.paragraphIds()[0]!;
    editor.surface!.setSelection({
      anchor: { paragraphId, offset: 0 },
      head: { paragraphId, offset: 3 },
    });
    await settle(fixture);
    expect(host.selections.length).toBeGreaterThan(0);
    expect(host.selections.at(-1)!.selection).toBeDefined();
    fixture.destroy();
  });

  it('emits fontError when the font resolver throws', async () => {
    const { fixture, host } = await mount(PackagedHostComponent, (h) =>
      h.fonts.set(() => {
        throw new Error('resolver exploded');
      })
    );
    await settle(fixture);
    expect(host.fontErrors).toHaveLength(1);
    expect(host.fontErrors[0]!.message).toContain('resolver exploded');
    fixture.destroy();
  });

  it('emits openFailure for bytes that are not a DOCX file', async () => {
    const consoleError = vi.spyOn(console, 'error').mockImplementation(() => {});
    try {
      const { fixture, host } = await mount(PackagedHostComponent, (h) =>
        h.bytes.set(new Uint8Array([1, 2, 3, 4]))
      );
      expect(host.failures).toHaveLength(1);
      expect(host.failures[0]!.reason).toBeDefined();
      expect(host.editor().snapshot().openFailure).toBe(host.failures[0]);
      fixture.destroy();
    } finally {
      consoleError.mockRestore();
    }
  });

  it('emits error when the engine refuses a document source', async () => {
    const consoleError = vi.spyOn(console, 'error').mockImplementation(() => {});
    try {
      const { fixture, host } = await mount(PackagedHostComponent);
      host.packaged()!.load('not a source' as never);
      await settle(fixture);
      expect(host.errors.length).toBeGreaterThan(0);
      fixture.destroy();
    } finally {
      consoleError.mockRestore();
    }
  });
});

describe('<docx-editor> methods', () => {
  beforeEach(() => configureZoneless());

  it('answers without an engine: exec refuses, snapshot is empty, save rejects', async () => {
    const { fixture, host } = await mount(PackagedHostComponent, (h) =>
      h.bytes.set(null as unknown as Uint8Array)
    );
    const packaged = host.packaged()!;
    expect(packaged.getEditor()).toBeNull();
    expect(packaged.getDocumentHandle()).toBeNull();
    expect(packaged.exec({ type: 'insertText', text: 'x' })).toEqual({
      ok: false,
      code: 'notFound',
      reason: 'no editor is mounted',
    });
    expect(packaged.snapshot()).toBe(LOADING_SNAPSHOT);
    await expect(packaged.save()).rejects.toThrow('no document is loaded');
    expect(() => packaged.focus()).not.toThrow();
    fixture.destroy();
  });

  it('runs exec, snapshot, load, and getDocumentHandle on the engine', async () => {
    const { fixture, host } = await mount(PackagedHostComponent);
    const packaged = host.packaged()!;
    const editor = host.editor();
    expect(packaged.exec({ type: 'insertText', text: 'X' }).ok).toBe(true);
    expect(packaged.snapshot().canUndo).toBe(true);
    expect(packaged.getDocumentHandle()).not.toBeNull();
    packaged.load(docx(paragraph('loaded')));
    await settle(fixture);
    expect(packaged.getEditor()).toBe(editor);
    expect(editor.surface!.session.bodyText()).toBe('loaded');
    fixture.destroy();
  });

  it('refuses to save a form field value that its format does not accept', async () => {
    const { fixture, host } = await mount(PackagedHostComponent, (h) =>
      h.bytes.set(formFieldDocx(true))
    );
    const surface = host.editor().surface!;
    surface.navigate('documentStart');
    for (let index = 0; index < 10; index += 1) surface.navigate('right', true);
    surface.type('invalid');
    await expect(host.packaged()!.save()).rejects.toMatchObject({ code: 'invalidArgs' });
    surface.undo();
    const saved = unzipSync(new Uint8Array(await host.packaged()!.save()));
    const xml = strFromU8(saved['word/document.xml']!);
    expect(xml).toContain('01/02/2030');
    expect(xml).not.toContain('invalid');
    fixture.destroy();
  });
});

describe('docxEditorRoot lifecycle', () => {
  beforeEach(() => configureZoneless());

  it('emits ready after the pages are painted', async () => {
    const { fixture, host } = await mount(RootHostComponent);
    expect(host.ready).toHaveLength(1);
    expect(host.pagesAtReady).toBeGreaterThan(0);
    fixture.destroy();
  });

  it("opens the 'blank' document source without warnings", async () => {
    const warn = vi.spyOn(console, 'warn');
    try {
      const { fixture } = await mount(RootHostComponent, (h) => h.document.set('blank'));
      expect((fixture.nativeElement as Element).querySelector('.docx-page')).not.toBeNull();
      expect(warn).not.toHaveBeenCalled();
      fixture.destroy();
    } finally {
      warn.mockRestore();
    }
  });

  it('keeps the engine and its undo history for a new inline translate function', async () => {
    const { fixture, host } = await mount(RootHostComponent, (h) =>
      h.translate.set((key) => `0:${key}`)
    );
    const editor = host.ready[0]!;
    expect(editor.exec({ type: 'insertText', text: '!' }).ok).toBe(true);
    host.translate.set((key) => `1:${key}`);
    await settle(fixture);
    expect(host.ready).toHaveLength(1);
    expect(editor.snapshot().canUndo).toBe(true);
    fixture.destroy();
  });

  it('does not create a new engine when the host renders again for another reason', async () => {
    const { fixture, host } = await mount(RootHostComponent);
    host.unrelated.set(1);
    await settle(fixture);
    expect(host.ready).toHaveLength(1);
    fixture.destroy();
  });
});
