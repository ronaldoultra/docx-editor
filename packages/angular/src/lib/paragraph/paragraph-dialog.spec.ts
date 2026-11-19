// Task 6.3: the claims of packages/vue/test/paragraph-dialog.test.ts.
import {
  ChangeDetectionStrategy,
  Component,
  Directive,
  signal,
  viewChild,
  type Type,
} from '@angular/core';
import type { ComponentFixture } from '@angular/core/testing';
import { TestBed } from '@angular/core/testing';
import type { DocxEditorInstance } from '@docx-editor.dev/core/editor';
import { serializeOoxmlPart } from '@docx-editor.dev/core/store';
import { afterEach, beforeEach, describe, expect, test } from 'vitest';
import { stylesDocx } from '../../testing/docx';
import { configureZoneless, engineOf, settle } from '../../testing/harness';
import { DocxEditorContentComponent } from '../content.component';
import {
  DocxDialogApplyDirective,
  DocxDialogCancelDirective,
  DocxDialogOutletDirective,
  DocxDialogPartDirective,
} from '../dialog/dialog-parts';
import { DocxEditorRootDirective } from '../root.directive';
import { DocxEditorViewportComponent } from '../viewport.component';
import {
  DocxEditorParagraphDialogComponent,
  injectParagraphDialog,
} from './paragraph-dialog.component';

// Pinned format defaults, so the fixture does not take the application defaults.
const DEFAULTS =
  '<w:docDefaults><w:rPrDefault><w:rPr><w:kern w:val="2"/></w:rPr></w:rPrDefault><w:pPrDefault/></w:docDefaults>';
const source = (body: string) => stylesDocx(DEFAULTS, body);
const p = (text: string, pPr = ''): string =>
  `<w:p>${pPr ? `<w:pPr>${pPr}</w:pPr>` : ''}<w:r><w:t>${text}</w:t></w:r></w:p>`;

/** A custom control that sets the alignment through the dialog state. */
@Component({
  selector: 'spec-custom-alignment',
  template: `<button
    type="button"
    data-custom-control
    (click)="dialog.setValue('alignment', 'right')"
  >
    {{ dialog.values().alignment }}
  </button>`,
})
class CustomAlignmentComponent {
  readonly dialog = injectParagraphDialog();
}

/** A custom control that reads and sets the direction through the dialog state. */
@Component({
  selector: 'spec-custom-direction',
  template: `<button
    type="button"
    data-custom-direction
    (click)="dialog.setValue('direction', 'rtl')"
  >
    {{ dialog.values().direction }}
  </button>`,
})
class CustomDirectionComponent {
  readonly dialog = injectParagraphDialog();
}

const PARTS = [
  DocxEditorRootDirective,
  DocxEditorViewportComponent,
  DocxEditorContentComponent,
  DocxEditorParagraphDialogComponent,
  DocxDialogPartDirective,
  DocxDialogApplyDirective,
  DocxDialogCancelDirective,
  DocxDialogOutletDirective,
  CustomAlignmentComponent,
  CustomDirectionComponent,
];

@Directive()
abstract class Host {
  readonly bytes = signal<Uint8Array>(source(p('alpha')));
  readonly open = signal(false);
  readonly root = viewChild.required(DocxEditorRootDirective);

  editor(): DocxEditorInstance {
    return engineOf(this.root().context.editor());
  }
}

@Component({
  selector: 'spec-paragraph-default',
  imports: PARTS,
  changeDetection: ChangeDetectionStrategy.OnPush,
  template: `<ng-container docxEditorRoot [document]="bytes()">
    <docx-editor-viewport>
      <docx-editor-content />
      <docx-editor-paragraph-dialog [open]="open()" (closed)="open.set(false)" />
    </docx-editor-viewport>
  </ng-container>`,
})
class DefaultHostComponent extends Host {}

@Component({
  selector: 'spec-paragraph-as-child',
  imports: PARTS,
  changeDetection: ChangeDetectionStrategy.OnPush,
  template: `<ng-container docxEditorRoot [document]="bytes()">
    <docx-editor-viewport>
      <docx-editor-content />
      <docx-editor-paragraph-dialog [open]="open()" (closed)="open.set(false)">
        <ng-template docxDialogPart="apply" asChild>
          <button docxDialogApply data-custom-apply>Save paragraph</button>
        </ng-template>
      </docx-editor-paragraph-dialog>
    </docx-editor-viewport>
  </ng-container>`,
})
class AsChildHostComponent extends Host {}

@Component({
  selector: 'spec-paragraph-layout',
  imports: PARTS,
  changeDetection: ChangeDetectionStrategy.OnPush,
  template: `<ng-container docxEditorRoot [document]="bytes()">
    <docx-editor-viewport>
      <docx-editor-content />
      <docx-editor-paragraph-dialog [open]="open()" [preset]="false" (closed)="open.set(false)">
        <button docxDialogCancel>Cancel</button>
        <spec-custom-alignment />
        <ng-container docxDialogOutlet="field:spaceBefore" />
        <button docxDialogApply>OK</button>
      </docx-editor-paragraph-dialog>
    </docx-editor-viewport>
  </ng-container>`,
})
class LayoutHostComponent extends Host {}

@Component({
  selector: 'spec-paragraph-direction',
  imports: PARTS,
  changeDetection: ChangeDetectionStrategy.OnPush,
  template: `<ng-container docxEditorRoot [document]="bytes()">
    <docx-editor-viewport>
      <docx-editor-content />
      <docx-editor-paragraph-dialog [open]="open()" [preset]="false" (closed)="open.set(false)">
        <spec-custom-direction />
        <button docxDialogApply>OK</button>
      </docx-editor-paragraph-dialog>
    </docx-editor-viewport>
  </ng-container>`,
})
class DirectionHostComponent extends Host {}

interface Mounted<T extends Host> {
  fixture: ComponentFixture<T>;
  host: T;
  editor: () => DocxEditorInstance;
  update: (action?: () => unknown) => Promise<void>;
}

/** Mount the host on `body`, select all, then open the dialog on a later tick, as a user does. */
async function mount<T extends Host>(component: Type<T>, body?: string): Promise<Mounted<T>> {
  const fixture = TestBed.createComponent(component);
  const host = fixture.componentInstance;
  if (body !== undefined) host.bytes.set(source(body));
  document.body.append(fixture.nativeElement);
  await settle(fixture);
  const update = async (action: () => unknown = () => {}) => {
    action();
    await settle(fixture, 2);
  };
  const editor = () => host.editor();
  await update(() => editor().surface!.selectAll());
  await update(() => host.open.set(true));
  return { fixture, host, editor, update };
}

/** One control, found by its accessible name. */
function field(label: string): HTMLInputElement & HTMLSelectElement {
  const found = [...document.querySelectorAll('input, select')].find(
    (node) => node.getAttribute('aria-label') === label
  );
  if (!found) throw new Error(`no control labelled ${label}`);
  return found as HTMLInputElement & HTMLSelectElement;
}

/** A checkbox, found by the text of the label that wraps it. */
function checkboxFor(text: string): HTMLInputElement {
  const label = [...document.querySelectorAll('label')].find((node) =>
    node.textContent?.includes(text)
  );
  const box = label?.querySelector<HTMLInputElement>('input[type="checkbox"]');
  if (!box) throw new Error(`no checkbox labelled ${text}`);
  return box;
}

function okButton(): HTMLButtonElement {
  const found = [...document.querySelectorAll('button')].find((node) =>
    /^(OK|Apply)$/i.test(node.textContent?.trim() ?? '')
  );
  if (!found) throw new Error('no OK button');
  return found;
}

function pick(node: HTMLSelectElement, value: string): void {
  node.value = value;
  node.dispatchEvent(new Event('change', { bubbles: true }));
}

function type(node: HTMLInputElement, value: string): void {
  node.value = value;
  node.dispatchEvent(new Event('input', { bubbles: true }));
}

beforeEach(() => configureZoneless());
afterEach(() => document.body.replaceChildren());

describe('the paragraph dialog', () => {
  test('seeds every field from the live selection when it opens', async () => {
    const { fixture } = await mount(
      DefaultHostComponent,
      p('alpha', '<w:spacing w:before="240" w:after="120"/><w:ind w:left="720"/>')
    );
    expect(field('Before').value).toBe('12');
    expect(field('After').value).toBe('6');
    expect(field('Before text').value).toBe('0.5');
    fixture.destroy();
  });

  test('a setting that the selection disagrees about is indeterminate, and survives OK', async () => {
    const { fixture, editor, update } = await mount(
      DefaultHostComponent,
      p('one', '<w:keepNext/>') + p('two')
    );
    expect(editor().snapshot().formatting?.paragraphFlags?.keepNext).toBeNull();
    // Indeterminate, not unchecked: unchecked would claim the paragraphs agree it is off.
    expect(checkboxFor('Keep with next').indeterminate).toBe(true);
    await update(() => okButton().click());
    // Untouched, so not written.
    expect(editor().snapshot().formatting?.paragraphFlags?.keepNext).toBeNull();
    fixture.destroy();
  });

  test('it opens focused, so Escape closes it without a Tab first', async () => {
    const { fixture, host, update } = await mount(DefaultHostComponent);
    const panel = document.querySelector('[role="dialog"]')!;
    expect(panel).not.toBeNull();
    expect(panel.contains(document.activeElement)).toBe(true);
    await update(() =>
      panel.dispatchEvent(
        new KeyboardEvent('keydown', { key: 'Escape', bubbles: true, cancelable: true })
      )
    );
    expect(host.open()).toBe(false);
    fixture.destroy();
  });

  test('every mixed field can be resolved, not only shown as mixed', async () => {
    const { fixture, editor, update } = await mount(
      DefaultHostComponent,
      p('one', '<w:jc w:val="right"/><w:spacing w:line="400" w:lineRule="exact"/>') +
        p('two', '<w:jc w:val="left"/><w:spacing w:line="280" w:lineRule="exact"/>')
    );
    const alignment = field('Alignment');
    const rule = field('Line spacing');
    const at = field('At');
    expect(alignment.value).toBe('');
    expect(rule.value).toBe('');
    // Locked until a rule says what a number in it means.
    expect(at.disabled).toBe(true);
    await update(() => {
      pick(alignment, 'center');
      pick(rule, 'exact');
    });
    expect(rule.value).toBe('exact');
    expect(at.disabled).toBe(false);
    await update(() => type(at, '16'));
    await update(() => okButton().click());
    const formatting = editor().snapshot().formatting;
    expect(formatting?.alignment).toBe('center');
    // Points, not sixteen line heights.
    expect(formatting?.lineSpacing).toEqual({ rule: 'exact', value: 16 });
    fixture.destroy();
  });

  test('a changed field is written, and only that field', async () => {
    const { fixture, editor, update } = await mount(DefaultHostComponent);
    await update(() => type(field('Before'), '18'));
    await update(() => okButton().click());
    const formatting = editor().snapshot().formatting;
    expect(formatting?.spaceBeforePt).toBe(18);
    // The line spacing that the box only showed was never written.
    expect(formatting?.lineSpacing).toEqual({ rule: 'multiple', value: 1 });
    expect(serializeOoxmlPart(editor().surface!.session.part())).not.toContain('w:line=');
    fixture.destroy();
  });

  test('a draft cannot apply to a replacement document before the next render', async () => {
    const { fixture, host, editor, update } = await mount(DefaultHostComponent, p('first'));
    await update(() => type(field('Before'), '18'));
    const apply = document.querySelector<HTMLButtonElement>('[data-docx-part="apply"]')!;
    editor().load(source(p('replacement')));
    await update(() => apply.click());
    expect(host.open()).toBe(false);
    await update(() => editor().surface!.selectAll());
    expect(editor().snapshot().formatting?.spaceBeforePt).not.toBe(18);
    fixture.destroy();
  });
});

describe('the paragraph dialog direction field', () => {
  test('it seeds the current direction, and OK writes the one that you pick', async () => {
    const { fixture, editor, update } = await mount(DefaultHostComponent, p('alpha', '<w:bidi/>'));
    const direction = field('Direction');
    expect(direction.value).toBe('rtl');
    await update(() => pick(direction, 'ltr'));
    await update(() => okButton().click());
    expect(editor().snapshot().formatting?.direction).toBe('ltr');
    fixture.destroy();
  });

  test('a mixed selection opens blank, stays mixed if untouched, and resolves when set', async () => {
    const { fixture, host, editor, update } = await mount(
      DefaultHostComponent,
      p('one', '<w:bidi/>') + p('two')
    );
    expect(editor().snapshot().formatting?.disagrees?.direction).toBe(true);
    expect(field('Direction').value).toBe('');
    // Untouched: OK writes nothing.
    await update(() => okButton().click());
    expect(host.open()).toBe(false);
    expect(editor().snapshot().canUndo).toBe(false);
    await update(() => host.open.set(true));
    const direction = field('Direction');
    expect(direction.value).toBe('');
    await update(() => pick(direction, 'rtl'));
    await update(() => okButton().click());
    expect(editor().snapshot().formatting?.direction).toBe('rtl');
    expect(editor().snapshot().formatting?.disagrees?.direction).toBe(false);
    fixture.destroy();
  });
});

describe('paragraph dialog composition', () => {
  test('an asChild Apply keeps the command and the packaged fields', async () => {
    const { fixture, host, editor, update } = await mount(AsChildHostComponent);
    await update(() => type(field('Before'), '18'));
    expect(document.querySelectorAll('[data-docx-part="apply"]')).toHaveLength(1);
    await update(() => document.querySelector<HTMLButtonElement>('[data-custom-apply]')!.click());
    expect(host.open()).toBe(false);
    expect(editor().snapshot().formatting?.spaceBeforePt).toBe(18);
    fixture.destroy();
  });

  test('a custom layout keeps Apply, places a packaged field, and sets a value', async () => {
    const { fixture, host, editor, update } = await mount(LayoutHostComponent);
    expect(document.querySelector('[data-docx-part="header"]')).toBeNull();
    const panel = document.querySelector('[data-docx-dialog="paragraph"]')!;
    expect(panel.firstElementChild?.getAttribute('data-docx-part')).toBe('cancel');
    // The outlet places the packaged Before field, with its wiring.
    expect(panel.querySelector('[data-docx-field="spaceBefore"]')).not.toBeNull();
    await update(() => type(field('Before'), '6'));
    await update(() => document.querySelector<HTMLButtonElement>('[data-custom-control]')!.click());
    await update(() =>
      document.querySelector<HTMLButtonElement>('[data-docx-part="apply"]')!.click()
    );
    expect(editor().snapshot().formatting?.alignment).toBe('right');
    expect(editor().snapshot().formatting?.spaceBeforePt).toBe(6);
    expect(host.open()).toBe(false);
    fixture.destroy();
  });

  test('a custom layout reads and sets the direction through injectParagraphDialog', async () => {
    const { fixture, host, editor, update } = await mount(DirectionHostComponent);
    const control = document.querySelector<HTMLButtonElement>('[data-custom-direction]')!;
    expect(control.textContent?.trim()).toBe('ltr');
    await update(() => control.click());
    expect(control.textContent?.trim()).toBe('rtl');
    await update(() =>
      document.querySelector<HTMLButtonElement>('[data-docx-part="apply"]')!.click()
    );
    expect(editor().snapshot().formatting?.direction).toBe('rtl');
    expect(host.open()).toBe(false);
    fixture.destroy();
  });
});
