// The claims of packages/vue/test/fonts-prop.test.ts, font-shaping-lifecycle.test.ts, and
// image-decode-port-prop.test.ts, the engine claim of editing-mode-author.test.ts, and the
// `modules` scenario of the packaging spec.
import { ChangeDetectionStrategy, Component, signal, viewChild, type Type } from '@angular/core';
import { TestBed } from '@angular/core/testing';
import type { Editor } from '@docx-editor.dev/core/contracts/editor';
import {
  validateRasterHeader,
  type DocxEditorInstance,
  type EditorModule,
  type FontResolver,
  type FontShapingOptions,
  type ImageDecodePort,
} from '@docx-editor.dev/core/editor';
import { serializeOoxmlPart } from '@docx-editor.dev/core/store';
import { strToU8, unzipSync, zipSync } from 'fflate';
import { docx, paragraph } from '../testing/docx';
import { configureZoneless, engineOf, settle } from '../testing/harness';
import { DocxEditorContentComponent } from './content.component';
import { DocxEditorComponent, type DocxEditorMode } from './docx-editor.component';
import { DocxEditorRootDirective, type DocxEditorFonts } from './root.directive';
import { DocxEditorViewportComponent } from './viewport.component';

const SOURCE = docx(paragraph('Fonts stay in place'));

const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const PNG = Uint8Array.from(
  atob(
    'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNkYAAAAAYAAjCB0C8AAAAASUVORK5CYII='
  ),
  (c) => c.charCodeAt(0)
);
/** One paragraph with an inline 1x1 PNG picture. */
const PICTURE_SOURCE = zipSync({
  '[Content_Types].xml': strToU8(
    '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="png" ContentType="image/png"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>'
  ),
  '_rels/.rels': strToU8(
    `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="office" Type="${R}/officeDocument" Target="word/document.xml"/></Relationships>`
  ),
  'word/_rels/document.xml.rels': strToU8(
    `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="picture" Type="${R}/image" Target="media/pixel.png"/></Relationships>`
  ),
  'word/document.xml': strToU8(
    `<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="${R}" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"><w:body><w:p><w:r><w:drawing><wp:inline><wp:extent cx="914400" cy="914400"/><wp:docPr id="1" name="pixel"/><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic><pic:nvPicPr><pic:cNvPr id="1" name="pixel"/><pic:cNvPicPr/></pic:nvPicPr><pic:blipFill><a:blip r:embed="picture"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill><pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="914400" cy="914400"/></a:xfrm><a:prstGeom prst="rect"/></pic:spPr></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p></w:body></w:document>`
  ),
  'word/media/pixel.png': PNG,
});

const REVIEW: EditorModule = {
  id: 'review',
  review: {
    displayModes: ['all-markup', 'proposed', 'original'],
    collectReviewItems: () => [],
    revisionItemsOfParagraph: () => [],
  },
};

/** A resolver that records each request under `name` and supplies nothing. */
function resolver(name: string, calls: string[]): FontResolver {
  return async () => {
    calls.push(name);
    return undefined;
  };
}

/** A decoder that records each call, and can hold its answer until `release`. */
function observedPort(hold = false) {
  const calls: { signal: AbortSignal | undefined; mime: string }[] = [];
  let release!: () => void;
  const gate = new Promise<void>((resolve) => {
    release = resolve;
  });
  const decode: ImageDecodePort['decode'] = async (bytes, mime, _limits, abort) => {
    const header = validateRasterHeader(bytes, mime);
    if (!header) throw new Error('Invalid synthetic PNG');
    calls.push({ signal: abort, mime });
    if (hold) await gate;
    return Object.freeze({ ...header, dpiX: 96, dpiY: 96 });
  };
  return { port: Object.freeze({ decode }) as ImageDecodePort, calls, release };
}

class Inputs {
  readonly bytes = signal<Uint8Array>(SOURCE);
  readonly fonts = signal<DocxEditorFonts | undefined>(undefined);
  readonly fontShaping = signal<FontShapingOptions | undefined>(undefined);
  readonly modules = signal<readonly EditorModule[] | undefined>(undefined);
  readonly imageDecodePort = signal<ImageDecodePort | undefined>(undefined);
  readonly author = signal<string | undefined>(undefined);
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
      [fonts]="fonts()"
      [fontShaping]="fontShaping()"
      [modules]="modules()"
      [imageDecodePort]="imageDecodePort()"
      [author]="author()"
      (ready)="ready.push($event)"
    >
      <docx-editor-viewport><docx-editor-content /></docx-editor-viewport>
    </ng-container>
  `,
})
class RootHostComponent extends Inputs implements EditorHost {
  readonly ready: Editor[] = [];
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
      [fonts]="fonts()"
      [fontShaping]="fontShaping()"
      [modules]="modules()"
      [imageDecodePort]="imageDecodePort()"
      [author]="author()"
    />
  `,
})
class PackagedHostComponent extends Inputs implements EditorHost {
  readonly mode = signal<DocxEditorMode>('edit');
  readonly packaged = viewChild.required(DocxEditorComponent);
  editor(): DocxEditorInstance {
    return engineOf(this.packaged().getEditor());
  }
}

async function mount<T>(type: Type<T>, setup: (host: T) => void = () => {}) {
  const fixture = TestBed.createComponent(type);
  setup(fixture.componentInstance);
  await settle(fixture);
  return { fixture, host: fixture.componentInstance };
}

const HOSTS: ReadonlyArray<readonly [string, Type<EditorHost>]> = [
  ['root', RootHostComponent],
  ['packaged', PackagedHostComponent],
];

describe.each(HOSTS)('construction-time inputs on the %s', (_name, type) => {
  beforeEach(() => configureZoneless());

  it('replaces the fonts of the same engine for a new fonts value', async () => {
    const calls: string[] = [];
    const { fixture, host } = await mount(type, (h) => h.fonts.set(resolver('first', calls)));
    await settle(fixture);
    const editor = host.editor();
    const session = editor.surface?.session;
    host.fonts.set(resolver('second', calls));
    await settle(fixture, 8);
    expect(host.editor()).toBe(editor);
    expect(editor.surface?.session).toBe(session);
    expect(calls).toEqual(['first', 'second']);
    fixture.destroy();
  });

  it('decodes pictures with the port, keeps them on save, and a new port makes a new engine', async () => {
    const observed = observedPort();
    const { fixture, host } = await mount(type, (h) => {
      h.bytes.set(PICTURE_SOURCE);
      h.imageDecodePort.set(observed.port);
    });
    expect(observed.calls.length).toBeGreaterThan(0);
    expect(observed.calls.every((call) => call.mime === 'image/png')).toBe(true);
    const first = host.editor();
    const saved = await first.save();
    expect(unzipSync(new Uint8Array(saved))['word/media/pixel.png']).toEqual(PNG);
    const replacement = observedPort();
    host.imageDecodePort.set(replacement.port);
    await settle(fixture);
    expect(host.editor()).not.toBe(first);
    expect(replacement.calls.length).toBeGreaterThan(0);
    fixture.destroy();
  });

  it('aborts a running decode when Angular destroys the editor', async () => {
    const observed = observedPort(true);
    const { fixture } = await mount(type, (h) => {
      h.bytes.set(PICTURE_SOURCE);
      h.imageDecodePort.set(observed.port);
    });
    expect(observed.calls.length).toBeGreaterThan(0);
    expect(observed.calls[0]!.signal?.aborted).toBe(false);
    fixture.destroy();
    observed.release();
    expect(observed.calls[0]!.signal?.aborted).toBe(true);
  });

  it('creates a new engine for new modules, and opens the document again', async () => {
    const { fixture, host } = await mount(type);
    const first = host.editor();
    const text = first.surface!.session.bodyText();
    host.modules.set([REVIEW]);
    await settle(fixture);
    const second = host.editor();
    expect(second).not.toBe(first);
    expect(second.surface!.session.bodyText()).toBe(text);
    fixture.destroy();
  });
});

describe('fontShaping on the root', () => {
  beforeEach(() => configureZoneless());

  it('keeps edits and undo for an equal new value, and makes a new engine for a changed one', async () => {
    const { fixture, host } = await mount(RootHostComponent, (h) =>
      h.fontShaping.set({ maxFontBytes: undefined })
    );
    expect(host.ready).toHaveLength(1);
    const first = host.editor();
    expect(first.exec({ type: 'insertText', text: 'kept' }).ok).toBe(true);
    const edited = first.surface!.session.bodyText();
    const selection = first.surface!.state().selection;
    host.fontShaping.set({ maxFontBytes: undefined });
    await settle(fixture);
    expect(host.ready).toHaveLength(1);
    expect(first.surface!.session.bodyText()).toBe(edited);
    expect(first.surface!.state().selection).toEqual(selection);
    expect(first.snapshot().canUndo).toBe(true);

    host.fontShaping.set({ maxFontBytes: 20 * 1024 * 1024 });
    await settle(fixture);
    expect(host.ready).toHaveLength(2);
    const second = host.editor();
    expect(second).not.toBe(first);
    expect(second.exec({ type: 'insertText', text: 'second' }).ok).toBe(true);
    const secondText = second.surface!.session.bodyText();
    host.fontShaping.set({ maxFontBytes: 20 * 1024 * 1024 });
    await settle(fixture);
    expect(host.ready).toHaveLength(2);
    expect(second.surface!.session.bodyText()).toBe(secondText);
    expect(second.snapshot().canUndo).toBe(true);
    fixture.destroy();
  });
});

describe("<docx-editor mode='suggesting'>", () => {
  beforeEach(() => configureZoneless());

  it('waits for an author, then records typing as tracked changes by that author', async () => {
    const { fixture, host } = await mount(PackagedHostComponent, (h) => {
      h.modules.set([REVIEW]);
      h.mode.set('suggesting');
    });
    const editor = host.editor();
    expect(editor.getEditingMode()).toBe('editing');
    host.author.set('Ada');
    await settle(fixture);
    expect(host.editor()).toBe(editor);
    expect(editor.getEditingMode()).toBe('suggesting');
    expect(editor.exec({ type: 'insertText', text: 'X' }).ok).toBe(true);
    const xml = serializeOoxmlPart(editor.surface!.session.part());
    expect(xml).toMatch(/<w:ins [^>]*w:author="Ada"/);
    fixture.destroy();
  });
});
