// Task 6.3: the page setup dialog, and the `pageSetup`, `paragraph`, and `watermark` entries of
// the popup registry. The Page Setup popup claims of packages/vue/test/menu-composition.test.ts.
import {
  ChangeDetectionStrategy,
  Component,
  Directive,
  input,
  signal,
  viewChild,
  type Type,
} from '@angular/core';
import { TestBed, type ComponentFixture } from '@angular/core/testing';
import type { DocxEditorInstance } from '@docx-editor.dev/core/editor';
import { afterEach, beforeEach, describe, expect, test } from 'vitest';
import { docx, paragraph } from '../../testing/docx';
import { configureZoneless, engineOf, settle } from '../../testing/harness';
import { DocxEditorContentComponent } from '../content.component';
import { DocxEditorMenuComponent } from '../menu/menu-bar.component';
import type { DocxEditorPopups } from '../popups/popup-config';
import { DocxEditorRootDirective } from '../root.directive';
import { DocxEditorViewportComponent } from '../viewport.component';

/** A registry dialog: it shows a button that closes it through `closeHandler`. */
@Component({
  selector: 'spec-own-page-setup',
  template: `<button type="button" data-own-page-setup (click)="closeHandler()()">Done</button>`,
})
class OwnPageSetupComponent {
  readonly closeHandler = input.required<() => void>();
}

@Directive()
abstract class Host {
  readonly bytes = signal<Uint8Array>(docx(paragraph('hello world')));
  readonly popups = signal<DocxEditorPopups | undefined>(undefined);
  readonly root = viewChild.required(DocxEditorRootDirective);

  editor(): DocxEditorInstance {
    return engineOf(this.root().context.editor());
  }
}

@Component({
  selector: 'spec-page-setup-menu',
  imports: [
    DocxEditorRootDirective,
    DocxEditorViewportComponent,
    DocxEditorContentComponent,
    DocxEditorMenuComponent,
  ],
  changeDetection: ChangeDetectionStrategy.OnPush,
  template: `<ng-container docxEditorRoot [document]="bytes()" [popups]="popups()">
    <div docxEditorMenu [overflow]="false"></div>
    <docx-editor-viewport><docx-editor-content /></docx-editor-viewport>
  </ng-container>`,
})
class MenuHostComponent extends Host {}

interface Mounted {
  fixture: ComponentFixture<MenuHostComponent>;
  root: HTMLElement;
  editor: () => DocxEditorInstance;
  update: (action?: () => unknown) => Promise<void>;
  openRow: (menu: string, slot: string) => Promise<HTMLButtonElement>;
}

async function mount(
  component: Type<MenuHostComponent> = MenuHostComponent,
  popups?: DocxEditorPopups
): Promise<Mounted> {
  const fixture = TestBed.createComponent(component);
  fixture.componentInstance.popups.set(popups);
  document.body.append(fixture.nativeElement);
  await settle(fixture);
  const root: HTMLElement = fixture.nativeElement;
  const update = async (action: () => unknown = () => {}) => {
    action();
    await settle(fixture, 2);
  };
  /** Open `menu` from its trigger and press the row `slot`. Returns the trigger. */
  const openRow = async (menu: string, slot: string) => {
    const trigger = root.querySelector<HTMLButtonElement>(
      `[data-menu="${menu}"] .docx-menubar__trigger`
    )!;
    trigger.focus();
    await update(() =>
      trigger.dispatchEvent(new KeyboardEvent('keydown', { key: 'Enter', bubbles: true }))
    );
    await update(() => root.querySelector<HTMLButtonElement>(`[data-slot="${slot}"]`)!.click());
    return trigger;
  };
  return { fixture, root, editor: () => fixture.componentInstance.editor(), update, openRow };
}

function field(label: string): HTMLInputElement & HTMLSelectElement {
  const found = [
    ...document.querySelectorAll('[data-docx-dialog] input, [data-docx-dialog] select'),
  ].find((node) => node.getAttribute('aria-label') === label);
  if (!found) throw new Error(`no control labelled ${label}`);
  return found as HTMLInputElement & HTMLSelectElement;
}

function type(node: HTMLInputElement, value: string): void {
  node.value = value;
  node.dispatchEvent(new Event('input', { bubbles: true }));
}

const dialog = (kind: string) => document.querySelector(`[data-docx-dialog="${kind}"]`);

beforeEach(() => configureZoneless());
afterEach(() => document.body.replaceChildren());

describe('the page setup dialog', () => {
  test('File › Page setup opens it seeded, and Apply writes the margins in one step', async () => {
    const { fixture, editor, update, openRow } = await mount();
    const trigger = await openRow('file', 'file.pageSetup');
    expect(dialog('pageSetup')).not.toBeNull();
    expect(field('Top').value).toBe('1');
    expect(field('Orientation').value).toBe('portrait');
    await update(() => type(field('Top'), '2'));
    const apply = document.querySelector<HTMLButtonElement>(
      '[data-docx-dialog="pageSetup"] [data-docx-part="apply"]'
    )!;
    await update(() => apply.click());
    expect(dialog('pageSetup')).toBeNull();
    expect(editor().snapshot().pageSetup?.marginsTwips.top).toBe(2880);
    // The focus goes back to the menu trigger, which outlives the menu row.
    expect(document.activeElement).toBe(trigger);
    editor().exec({ type: 'undo' });
    expect(editor().snapshot().pageSetup?.marginsTwips.top).toBe(1440);
    fixture.destroy();
  });

  test('a new orientation swaps the page width and height', async () => {
    const { fixture, editor, update, openRow } = await mount();
    await openRow('file', 'file.pageSetup');
    const orientation = field('Orientation');
    await update(() => {
      orientation.value = 'landscape';
      orientation.dispatchEvent(new Event('change', { bubbles: true }));
    });
    await update(() =>
      document.querySelector<HTMLButtonElement>('[data-docx-part="apply"]')!.click()
    );
    const setup = editor().snapshot().pageSetup!;
    expect(setup.orientation).toBe('landscape');
    expect(setup.pageWidthTwips).toBe(15840);
    expect(setup.pageHeightTwips).toBe(12240);
    fixture.destroy();
  });

  test('Escape closes it without a write, and a new document closes it', async () => {
    const { fixture, editor, update, openRow } = await mount();
    await openRow('file', 'file.pageSetup');
    const before = editor().getDocumentHandle().revision;
    await update(() =>
      dialog('pageSetup')!.dispatchEvent(
        new KeyboardEvent('keydown', { key: 'Escape', bubbles: true, cancelable: true })
      )
    );
    expect(dialog('pageSetup')).toBeNull();
    expect(editor().getDocumentHandle().revision).toBe(before);
    await openRow('file', 'file.pageSetup');
    expect(dialog('pageSetup')).not.toBeNull();
    await update(() => fixture.componentInstance.bytes.set(docx(paragraph('next'))));
    await settle(fixture);
    expect(dialog('pageSetup')).toBeNull();
    fixture.destroy();
  });
});

describe('the dialog entries of the popup registry', () => {
  test('a registry component replaces the packaged dialog and closes with closeHandler', async () => {
    const { fixture, update, openRow } = await mount(MenuHostComponent, {
      pageSetup: OwnPageSetupComponent,
    });
    await openRow('file', 'file.pageSetup');
    expect(dialog('pageSetup')).toBeNull();
    const own = document.querySelector<HTMLButtonElement>('[data-own-page-setup]')!;
    expect(own).not.toBeNull();
    await update(() => own.click());
    expect(document.querySelector('[data-own-page-setup]')).toBeNull();
    fixture.destroy();
  });

  test('false opens no dialog, and the other entries keep their packaged dialog', async () => {
    const { fixture, openRow } = await mount(MenuHostComponent, {
      pageSetup: false,
      watermark: false,
    });
    await openRow('file', 'file.pageSetup');
    expect(dialog('pageSetup')).toBeNull();
    await openRow('insert', 'insert.watermark');
    expect(dialog('watermark')).toBeNull();
    await openRow('format', 'paragraph.dialog');
    expect(dialog('paragraph')).not.toBeNull();
    fixture.destroy();
  });

  test('an entry that turns to false closes its open dialog', async () => {
    const { fixture, update, openRow } = await mount();
    await openRow('insert', 'insert.watermark');
    expect(dialog('watermark')).not.toBeNull();
    await update(() => fixture.componentInstance.popups.set({ watermark: false }));
    expect(dialog('watermark')).toBeNull();
    fixture.destroy();
  });
});
