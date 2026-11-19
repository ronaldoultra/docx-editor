// Task 5.3: the claims of packages/vue/test/chrome-popup-dismissal.test.ts. Every toolbar and
// menu popup follows one dismissal rule. An Escape from the pages closes the popup and stops
// there. An Escape from host content closes it and keeps its default. Focus that moves to
// another control, or into a host's modal dialog, closes it.
import { ChangeDetectionStrategy, Component, type Type } from '@angular/core';
import type { EditorModule } from '@docx-editor.dev/core/editor';
import { en } from '@docx-editor.dev/core/i18n';
import { afterEach, beforeEach, describe, expect, test } from 'vitest';
import { docx, paragraph } from '../testing/docx';
import { configureZoneless, render, settle } from '../testing/harness';
import { pressEscape } from '../testing/toolbar';
import { DocxEditorContentComponent } from './content.component';
import { DocxEditorMenuComponent } from './menu/menu-bar.component';
import { DocxEditorRootDirective } from './root.directive';
import { DocxToolbarFontFamilyComponent } from './toolbar/toolbar-pickers.component';
import { DocxToolbarReviewersComponent } from './toolbar/toolbar-reviewers.component';
import { DocxEditorToolbarComponent } from './toolbar/toolbar.component';
import { DocxEditorViewportComponent } from './viewport.component';

const SOURCE = docx(paragraph('Body text'));
/** One paragraph with a tracked insertion, so the reviewers menu has an author to list. */
const TRACKED_SOURCE = docx(
  '<w:p><w:r><w:t xml:space="preserve">Body </w:t></w:r>' +
    '<w:ins w:id="1" w:author="Ada" w:date="2026-01-02T03:04:05Z"><w:r><w:t>text</w:t></w:r></w:ins>' +
    '</w:p>'
);

const REVIEW_MODULE: EditorModule = {
  id: 'review',
  review: {
    displayModes: ['all-markup', 'proposed', 'original'],
    collectReviewItems: () => [],
    revisionItemsOfParagraph: () => [],
  },
};

const PARTS = [
  DocxEditorRootDirective,
  DocxEditorViewportComponent,
  DocxEditorContentComponent,
  DocxEditorMenuComponent,
  DocxEditorToolbarComponent,
  DocxToolbarFontFamilyComponent,
  DocxToolbarReviewersComponent,
];

abstract class Host {
  readonly modules = [REVIEW_MODULE];
}

@Component({
  selector: 'docx-spec-dismiss-menu-host',
  imports: PARTS,
  changeDetection: ChangeDetectionStrategy.OnPush,
  template: `<ng-container docxEditorRoot [document]="bytes" [modules]="modules">
    <div docxEditorMenu [overflow]="false"></div>
    <docx-editor-viewport><docx-editor-content /></docx-editor-viewport>
  </ng-container>`,
})
class MenuHostComponent extends Host {
  readonly bytes = SOURCE;
}

@Component({
  selector: 'docx-spec-dismiss-reviewers-host',
  imports: PARTS,
  changeDetection: ChangeDetectionStrategy.OnPush,
  template: `<ng-container docxEditorRoot [document]="bytes" [modules]="modules">
    <fieldset docxEditorToolbar [preset]="false" [overflow]="false">
      <div docxToolbarReviewers></div>
    </fieldset>
    <docx-editor-viewport><docx-editor-content /></docx-editor-viewport>
  </ng-container>`,
})
class ReviewersHostComponent extends Host {
  readonly bytes = TRACKED_SOURCE;
}

@Component({
  selector: 'docx-spec-dismiss-picker-host',
  imports: PARTS,
  changeDetection: ChangeDetectionStrategy.OnPush,
  template: `<ng-container docxEditorRoot [document]="bytes" [modules]="modules">
    <fieldset docxEditorToolbar [preset]="false" [overflow]="false">
      <div docxToolbarFontFamily></div>
    </fieldset>
    <docx-editor-viewport><docx-editor-content /></docx-editor-viewport>
  </ng-container>`,
})
class PickerHostComponent extends Host {
  readonly bytes = SOURCE;
}

async function mount(component: Type<unknown>) {
  const fixture = await render(component);
  const container: HTMLElement = fixture.nativeElement;
  document.body.append(container);
  await settle(fixture);
  const host = document.createElement('input');
  document.body.append(host);
  return {
    container,
    host,
    pages: () => container.querySelector<HTMLElement>('.docx-pages')!,
  };
}

async function press(trigger: HTMLElement): Promise<void> {
  trigger.dispatchEvent(new MouseEvent('mousedown', { bubbles: true, cancelable: true }));
  trigger.click();
  await settle(undefined, 2);
}

beforeEach(() => configureZoneless());
afterEach(() => document.body.replaceChildren());

describe('the menu bar', () => {
  const trigger = (container: HTMLElement) =>
    container.querySelector<HTMLElement>('.docx-menubar__trigger')!;
  const menuOpen = (container: HTMLElement) =>
    container.querySelector('.docx-menubar__trigger[aria-expanded="true"]') !== null;

  test('Escape from the pages closes it and stops there', async () => {
    const view = await mount(MenuHostComponent);
    await press(trigger(view.container));
    expect(menuOpen(view.container)).toBe(true);
    expect((await pressEscape(view.pages())).defaultPrevented).toBe(true);
    expect(menuOpen(view.container)).toBe(false);
  });

  test('Escape from a host input closes it and keeps its default', async () => {
    const view = await mount(MenuHostComponent);
    await press(trigger(view.container));
    expect(menuOpen(view.container)).toBe(true);
    expect((await pressEscape(view.host)).defaultPrevented).toBe(false);
    expect(menuOpen(view.container)).toBe(false);
  });

  test('focus that moves to another control closes it', async () => {
    const view = await mount(MenuHostComponent);
    await press(trigger(view.container));
    expect(menuOpen(view.container)).toBe(true);
    view.host.focus();
    await settle(undefined, 2);
    expect(menuOpen(view.container)).toBe(false);
  });
});

describe('the reviewers menu', () => {
  const trigger = (container: HTMLElement) =>
    container.querySelector<HTMLElement>('[data-slot="review.authors"] [aria-haspopup]')!;

  test('Escape from the pages closes it and stops there', async () => {
    const view = await mount(ReviewersHostComponent);
    await press(trigger(view.container));
    expect(trigger(view.container).getAttribute('aria-expanded')).toBe('true');
    expect((await pressEscape(view.pages())).defaultPrevented).toBe(true);
    expect(trigger(view.container).getAttribute('aria-expanded')).toBe('false');
  });

  test('Escape from a host input closes it and keeps its default', async () => {
    const view = await mount(ReviewersHostComponent);
    await press(trigger(view.container));
    expect(trigger(view.container).getAttribute('aria-expanded')).toBe('true');
    expect((await pressEscape(view.host)).defaultPrevented).toBe(false);
    expect(trigger(view.container).getAttribute('aria-expanded')).toBe('false');
  });

  test('the open menu lists the authors, and the arrow keys move between its rows', async () => {
    const view = await mount(ReviewersHostComponent);
    await press(trigger(view.container));
    const rows = [...view.container.querySelectorAll<HTMLElement>('[role="menuitemcheckbox"]')];
    expect(rows.map((row) => row.textContent?.replace('✓', '').trim())).toEqual([
      en.reviewers.all,
      'Ada',
    ]);
    expect(document.activeElement).toBe(rows[0]);
    rows[0]!.dispatchEvent(
      new KeyboardEvent('keydown', { key: 'ArrowDown', bubbles: true, cancelable: true })
    );
    expect(document.activeElement).toBe(rows[1]);
    expect(rows[1]!.getAttribute('aria-checked')).toBe('true');
    rows[1]!.dispatchEvent(
      new KeyboardEvent('keydown', { key: 'Home', bubbles: true, cancelable: true })
    );
    expect(document.activeElement).toBe(rows[0]);
  });
});

describe('a picker', () => {
  const trigger = (container: HTMLElement) =>
    container.querySelector<HTMLElement>('[data-slot="font.family"] [aria-haspopup]')!;

  test('focus that moves to another control closes it, then Escape reaches that control', async () => {
    const view = await mount(PickerHostComponent);
    await press(trigger(view.container));
    expect(trigger(view.container).getAttribute('aria-expanded')).toBe('true');
    view.host.focus();
    await settle(undefined, 2);
    expect(trigger(view.container).getAttribute('aria-expanded')).toBe('false');
    expect((await pressEscape(view.host)).defaultPrevented).toBe(false);
  });

  test("focus that moves into a host's modal dialog closes the picker", async () => {
    const view = await mount(PickerHostComponent);
    await press(trigger(view.container));
    expect(trigger(view.container).getAttribute('aria-expanded')).toBe('true');
    const dialog = document.createElement('div');
    dialog.setAttribute('role', 'dialog');
    dialog.setAttribute('aria-modal', 'true');
    const field = document.createElement('input');
    dialog.append(field);
    document.body.append(dialog);
    field.focus();
    await settle(undefined, 2);
    expect(trigger(view.container).getAttribute('aria-expanded')).toBe('false');
  });
});
