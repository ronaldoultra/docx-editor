// Task 6.3: the claims of packages/vue/test/watermark-dialog.test.ts.
import {
  ChangeDetectionStrategy,
  Component,
  Directive,
  signal,
  viewChild,
  type Type,
} from '@angular/core';
import { TestBed, type ComponentFixture } from '@angular/core/testing';
import type { Watermark } from '@docx-editor.dev/core/contracts/editor';
import type { DocxEditorInstance } from '@docx-editor.dev/core/editor';
import { afterAll, afterEach, beforeAll, beforeEach, describe, expect, test, vi } from 'vitest';
import { docx, paragraph } from '../../testing/docx';
import { configureZoneless, engineOf, settle } from '../../testing/harness';
import { DocxEditorContentComponent } from '../content.component';
import { DocxDialogOutletDirective, DocxDialogPartDirective } from '../dialog/dialog-parts';
import { DocxEditorMenuComponent } from '../menu/menu-bar.component';
import { DocxEditorRootDirective } from '../root.directive';
import { DocxEditorViewportComponent } from '../viewport.component';
import { DocxEditorWatermarkDialogComponent } from './watermark-dialog.component';

const PNG_1X1 = Uint8Array.from(
  atob(
    'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNkYAAAAAYAAjCB0C8AAAAASUVORK5CYII='
  ),
  (c) => c.charCodeAt(0)
);

const PARTS = [
  DocxEditorRootDirective,
  DocxEditorViewportComponent,
  DocxEditorContentComponent,
  DocxEditorWatermarkDialogComponent,
  DocxDialogPartDirective,
  DocxDialogOutletDirective,
];

@Directive()
abstract class Host {
  readonly bytes = signal<Uint8Array>(docx(paragraph('hello world')));
  readonly show = signal(false);
  readonly presets = signal<readonly string[] | undefined>(undefined);
  closed = 0;
  readonly root = viewChild.required(DocxEditorRootDirective);

  editor(): DocxEditorInstance {
    return engineOf(this.root().context.editor());
  }
}

@Component({
  selector: 'spec-watermark-menu',
  imports: [...PARTS, DocxEditorMenuComponent],
  changeDetection: ChangeDetectionStrategy.OnPush,
  template: `<ng-container docxEditorRoot [document]="bytes()">
    <div docxEditorMenu [overflow]="false"></div>
    <docx-editor-viewport><docx-editor-content /></docx-editor-viewport>
  </ng-container>`,
})
class MenuHostComponent extends Host {}

@Component({
  selector: 'spec-watermark-dialog',
  imports: PARTS,
  changeDetection: ChangeDetectionStrategy.OnPush,
  template: `<ng-container docxEditorRoot [document]="bytes()">
    <docx-editor-viewport>
      <docx-editor-content />
      @if (show()) {
        <docx-editor-watermark-dialog [presets]="presets()" (closed)="closed = closed + 1" />
      }
    </docx-editor-viewport>
  </ng-container>`,
})
class DialogHostComponent extends Host {}

@Component({
  selector: 'spec-watermark-composed',
  imports: PARTS,
  changeDetection: ChangeDetectionStrategy.OnPush,
  template: `<ng-container docxEditorRoot [document]="bytes()">
    <docx-editor-viewport>
      <docx-editor-content />
      @if (show()) {
        <docx-editor-watermark-dialog [preset]="false">
          <ng-template docxDialogPart="preview" keepContent className="host-preview" />
          <ng-container docxDialogOutlet="field:textWatermark" />
          <ng-container docxDialogOutlet="preview" />
        </docx-editor-watermark-dialog>
      }
    </docx-editor-viewport>
  </ng-container>`,
})
class ComposedHostComponent extends Host {}

interface Mounted<T extends Host> {
  fixture: ComponentFixture<T>;
  host: T;
  editor: () => DocxEditorInstance;
  update: (action?: () => unknown) => Promise<void>;
}

/** Mount the host. With `watermark`, the document has it before the dialog opens. */
async function mount<T extends Host>(
  component: Type<T>,
  options: { watermark?: Watermark; presets?: readonly string[]; show?: boolean } = {}
): Promise<Mounted<T>> {
  const fixture = TestBed.createComponent(component);
  const host = fixture.componentInstance;
  host.presets.set(options.presets);
  document.body.append(fixture.nativeElement);
  await settle(fixture);
  const update = async (action: () => unknown = () => {}) => {
    action();
    await settle(fixture, 2);
  };
  if (options.watermark) host.editor().exec({ type: 'setWatermark', watermark: options.watermark });
  if (options.show !== false) await update(() => host.show.set(true));
  return { fixture, host, editor: () => host.editor(), update };
}

const dialog = () => document.querySelector('[data-docx-dialog="watermark"]');
const preview = () =>
  document.querySelector<HTMLElement>('[data-docx-dialog="watermark"] [data-docx-part="preview"]')!;

/** One control, found by the visible text of its label. */
function control(label: string): HTMLInputElement & HTMLSelectElement {
  const owner = [
    ...document.querySelectorAll<HTMLLabelElement>('[data-docx-dialog="watermark"] label'),
  ].find((node) => node.textContent?.trim() === label);
  if (!owner) throw new Error(`no label ${label}`);
  const found = owner.htmlFor
    ? document.getElementById(owner.htmlFor)
    : owner.querySelector('input, select');
  if (!found) throw new Error(`no control labelled ${label}`);
  return found as HTMLInputElement & HTMLSelectElement;
}

function button(text: string): HTMLButtonElement {
  const found = [
    ...document.querySelectorAll<HTMLButtonElement>('[data-docx-dialog="watermark"] button'),
  ].find((node) => node.textContent?.trim() === text);
  if (!found) throw new Error(`no button ${text}`);
  return found;
}

function type(element: HTMLInputElement & HTMLSelectElement, value: string): void {
  element.value = value;
  element.dispatchEvent(new Event(element.localName === 'select' ? 'change' : 'input'));
}

/** The `matrix(a b c d e f)` fit of a preview text transform. */
function fitOf(text: Element): number[] {
  const match = /matrix\(([^)]*)\)/.exec(text.getAttribute('transform') ?? '');
  return match ? match[1]!.split(' ').map(Number) : [];
}

async function choosePicture(update: Mounted<Host>['update']): Promise<void> {
  const input = document.querySelector<HTMLInputElement>(
    '[data-docx-dialog="watermark"] input[type="file"]'
  )!;
  const file = new File([PNG_1X1], 'logo.png', { type: 'image/png' });
  // jsdom has no `Blob.arrayBuffer`.
  Object.defineProperty(file, 'arrayBuffer', { value: async () => PNG_1X1.slice().buffer });
  Object.defineProperty(input, 'files', { value: [file], configurable: true });
  await update(() => input.dispatchEvent(new Event('change')));
}

// jsdom has no object URLs.
const realCreate = URL.createObjectURL;
const realRevoke = URL.revokeObjectURL;
beforeAll(() => {
  let next = 0;
  URL.createObjectURL = vi.fn(() => `blob:spec-${++next}`);
  URL.revokeObjectURL = vi.fn();
});
afterAll(() => {
  URL.createObjectURL = realCreate;
  URL.revokeObjectURL = realRevoke;
});
beforeEach(() => configureZoneless());
afterEach(() => document.body.replaceChildren());

describe('the watermark dialog', () => {
  test('Insert › Watermark opens it, and Apply sets a text watermark in one undo step', async () => {
    const { fixture, editor, update } = await mount(MenuHostComponent, { show: false });
    const root: HTMLElement = fixture.nativeElement;
    const trigger = root.querySelector<HTMLButtonElement>(
      '[data-menu="insert"] .docx-menubar__trigger'
    )!;
    trigger.focus();
    await update(() =>
      trigger.dispatchEvent(new KeyboardEvent('keydown', { key: 'Enter', bubbles: true }))
    );
    const row = root.querySelector<HTMLElement>('[data-slot="insert.watermark"]')!;
    expect(row.getAttribute('aria-disabled')).not.toBe('true');
    expect(row.hasAttribute('disabled')).toBe(false);
    await update(() => row.click());
    expect(dialog()).not.toBeNull();
    expect(control('No watermark').checked).toBe(true);
    expect(control('Text').disabled).toBe(true);

    await update(() => control('Text watermark').click());
    expect(control('Text').disabled).toBe(false);
    await update(() => type(control('Text'), 'CONFIDENTIAL'));
    await update(() => type(control('Size'), '72'));
    await update(() => control('Bold').click());
    await update(() => type(control('Transparency'), '25'));
    await update(() => type(control('Color'), '#ff0000'));
    await update(() => control('Horizontal').click());
    await update(() => button('Apply').click());

    expect(dialog()).toBeNull();
    expect(editor().getWatermark()).toMatchObject({
      kind: 'text',
      text: 'CONFIDENTIAL',
      fontSizePt: 72,
      bold: true,
      color: 'FF0000',
      layout: 'horizontal',
    });
    expect(editor().getWatermark()!.transparency).toBeCloseTo(0.25, 2);
    editor().exec({ type: 'undo' });
    expect(editor().getWatermark()).toBeNull();
    fixture.destroy();
  });

  test('it seeds from the existing watermark, and No watermark removes it', async () => {
    const { fixture, host, editor, update } = await mount(DialogHostComponent, {
      watermark: { text: 'SAMPLE', italic: true, layout: 'horizontal', fontSizePt: 54 },
    });
    expect(control('Text watermark').checked).toBe(true);
    expect(control('Text').value).toBe('SAMPLE');
    expect(control('Italic').checked).toBe(true);
    expect(control('Size').value).toBe('54');
    expect(control('Horizontal').checked).toBe(true);
    await update(() => control('No watermark').click());
    await update(() => button('Apply').click());
    expect(host.closed).toBe(1);
    expect(editor().getWatermark()).toBeNull();
    fixture.destroy();
  });

  test('Cancel writes nothing', async () => {
    const { fixture, host, editor, update } = await mount(DialogHostComponent);
    const before = editor().getDocumentHandle().revision;
    await update(() => control('Text watermark').click());
    await update(() => type(control('Text'), 'URGENT'));
    await update(() => button('Cancel').click());
    expect(host.closed).toBe(1);
    expect(editor().getDocumentHandle().revision).toBe(before);
    expect(editor().getWatermark()).toBeNull();
    fixture.destroy();
  });

  test('a chosen picture applies as a picture watermark', async () => {
    const { fixture, editor, update } = await mount(DialogHostComponent, {
      presets: ['INTERNAL'],
    });
    await update(() => control('Picture watermark').click());
    expect(button('Apply').disabled).toBe(true);
    await choosePicture(update);
    expect(dialog()!.textContent).toContain('logo.png');
    // Auto fits the picture to the text area; a scale multiplies its natural size.
    const picture = () => preview().querySelector('[data-docx-watermark-preview="picture"]')!;
    expect(picture().getAttribute('width')).toBe('468');
    await update(() => type(control('Scale'), '2'));
    expect(picture().getAttribute('width')).toBe('1.5');
    await update(() => control('Washout').click());
    await update(() => button('Apply').click());
    expect(editor().getWatermark()).toMatchObject({ kind: 'image', washout: false });
    fixture.destroy();
  });

  test('custom presets fill the preset select, and typed text shows Custom', async () => {
    const { fixture, update } = await mount(DialogHostComponent, {
      presets: ['INTERNAL', 'PUBLIC'],
    });
    const select = control('Preset');
    expect([...select.options].map((option) => option.textContent?.trim())).toEqual([
      'Custom',
      'INTERNAL',
      'PUBLIC',
    ]);
    expect(select.disabled).toBe(true);
    await update(() => control('Text watermark').click());
    expect(select.value).toBe('');
    await update(() => type(select, 'PUBLIC'));
    expect(control('Text').value).toBe('PUBLIC');
    expect(select.value).toBe('PUBLIC');
    await update(() => type(control('Text'), 'PUBLIC v2'));
    expect(select.value).toBe('');
    expect(select.selectedOptions[0]!.textContent?.trim()).toBe('Custom');
    fixture.destroy();
  });

  test('the preview draws the draft and follows every field', async () => {
    const { fixture, update } = await mount(DialogHostComponent);
    expect(preview().getAttribute('aria-hidden')).toBe('true');
    expect(preview().querySelector('[data-docx-watermark-preview]')).toBeNull();
    const text = () => preview().querySelector('[data-docx-watermark-preview="text"]')!;

    await update(() => control('Text watermark').click());
    expect(text().textContent).toBe('DRAFT');
    expect(text().getAttribute('fill')).toBe('#C0C0C0');
    expect(text().getAttribute('fill-opacity')).toBe('0.5');
    expect(text().getAttribute('transform')).toContain('rotate(-45');
    expect(fitOf(text())).toHaveLength(6);

    await update(() => type(control('Text'), 'SECRET'));
    await update(() => type(control('Color'), '#ff0000'));
    await update(() => type(control('Transparency'), '25'));
    await update(() => control('Bold').click());
    await update(() => control('Horizontal').click());
    expect(text().textContent).toBe('SECRET');
    expect(text().getAttribute('fill')).toBe('#FF0000');
    expect(text().getAttribute('fill-opacity')).toBe('0.75');
    expect(text().getAttribute('font-weight')).toBe('bold');
    expect(text().getAttribute('transform')).not.toContain('rotate');

    // Auto fills the 6.5 inch text area of a Letter page, so the ink starts at the 72 point
    // margin plus the 7.2 point inset. A fixed 72 point size is an estimated advance centered
    // on the page.
    await update(() => type(control('Text'), 'xd'));
    expect(text().textContent).toBe('xd');
    expect(fitOf(text())[4]).toBeCloseTo(79.2, 3);
    await update(() => type(control('Size'), '72'));
    expect(fitOf(text())[4]).toBeCloseTo(273.6, 3);

    await update(() => control('No watermark').click());
    expect(preview().querySelector('[data-docx-watermark-preview]')).toBeNull();
    fixture.destroy();
  });

  test('Select picture is enabled for any kind, and a chosen file switches to Picture', async () => {
    const { fixture, update } = await mount(DialogHostComponent);
    await update(() => control('Text watermark').click());
    expect(button('Select picture…').disabled).toBe(false);
    expect(control('Scale').disabled).toBe(true);
    expect(control('Washout').disabled).toBe(true);
    await choosePicture(update);
    expect(control('Picture watermark').checked).toBe(true);
    expect(control('Scale').disabled).toBe(false);
    const picture = () =>
      preview().querySelector<SVGImageElement>('[data-docx-watermark-preview="picture"]')!;
    expect(picture().getAttribute('href')).toMatch(/^blob:/);
    expect(picture().style.filter).toContain('brightness');
    await update(() => control('Washout').click());
    expect(picture().style.filter).toBe('');
    // The button alone switches the kind too.
    await update(() => control('Text watermark').click());
    await update(() => button('Select picture…').click());
    expect(control('Picture watermark').checked).toBe(true);
    fixture.destroy();
  });

  test('the first focus lands on the checked kind', async () => {
    const { fixture } = await mount(DialogHostComponent, { watermark: { text: 'SAMPLE' } });
    expect(document.activeElement).toBe(control('Text watermark'));
    fixture.destroy();
  });

  test('an unchanged Apply closes without a write', async () => {
    const { fixture, host, editor, update } = await mount(DialogHostComponent, {
      watermark: { text: 'SAMPLE', italic: true },
    });
    const before = editor().getDocumentHandle().revision;
    await update(() => button('Apply').click());
    expect(host.closed).toBe(1);
    expect(editor().getDocumentHandle().revision).toBe(before);
    // One undo removes the watermark: the Apply added no empty step on top of it.
    editor().exec({ type: 'undo' });
    expect(editor().getWatermark()).toBeNull();
    fixture.destroy();
  });

  test('No watermark on a document without one applies as no change', async () => {
    const { fixture, host, editor, update } = await mount(DialogHostComponent);
    const before = editor().getDocumentHandle().revision;
    await update(() => button('Apply').click());
    expect(host.closed).toBe(1);
    expect(editor().getDocumentHandle().revision).toBe(before);
    expect(editor().snapshot().canUndo ?? false).toBe(false);
    fixture.destroy();
  });

  test('a custom layout places the preview part, which follows the draft', async () => {
    const { fixture, update } = await mount(ComposedHostComponent);
    const part = preview();
    expect(part.classList.contains('host-preview')).toBe(true);
    expect(part.classList.contains('docx-dialog__preview')).toBe(true);
    expect(part.getAttribute('aria-hidden')).toBe('true');
    expect(part.querySelector('[data-docx-watermark-preview]')).toBeNull();
    await update(() => control('Text watermark').click());
    expect(preview().querySelector('[data-docx-watermark-preview="text"]')?.textContent).toBe(
      'DRAFT'
    );
    fixture.destroy();
  });
});
