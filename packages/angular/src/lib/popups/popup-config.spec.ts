// Task 6.1: the registry claims of packages/vue/test/popup-config.test.ts. The packaged-editor
// claims (legacy flags, one hyperlink inside the content, context menu translation) port with
// task 6.7, and the note popup claim with task 9.2.
import {
  ChangeDetectionStrategy,
  Component,
  DestroyRef,
  inject,
  input,
  signal,
  viewChildren,
  type OnInit,
} from '@angular/core';
import { TestBed } from '@angular/core/testing';
import type { DocxEditorInstance } from '@docx-editor.dev/core/editor';
import { afterEach, beforeEach, describe, expect, test } from 'vitest';
import { docx } from '../../testing/docx';
import { configureZoneless, render, settle } from '../../testing/harness';
import { DocxEditorContentComponent } from '../content.component';
import { injectDocxEditor } from '../context';
import { DocxEditorRootDirective } from '../root.directive';
import { DocxEditorViewportComponent } from '../viewport.component';
import {
  injectDocxEditorPopups,
  provideDocxEditorPopups,
  type DocxEditorPopups,
} from './popup-config';
import { DocxPopupOutletComponent } from './popup-outlet.component';

const SOURCE = docx('<w:p><w:r><w:t>Popups</w:t></w:r></w:p>');
const PARTS = [DocxEditorRootDirective, DocxEditorViewportComponent, DocxEditorContentComponent];

@Component({
  selector: 'spec-link-popup',
  template: `<span data-custom-hyperlink>{{ label() }}</span>`,
})
class LinkPopupComponent {
  readonly label = input('Link');
}

/** A popup with its own state, which records its engine and its lifetime. */
@Component({
  selector: 'spec-stateful-popup',
  template: `<button data-component-popup (click)="press()">{{ count() }}</button>`,
})
class StatefulPopupComponent {
  static setups = 0;
  static destroys = 0;
  static owners: (DocxEditorInstance | null)[] = [];
  private readonly editor = injectDocxEditor();
  readonly count = signal(0);
  constructor() {
    StatefulPopupComponent.setups += 1;
    inject(DestroyRef).onDestroy(() => (StatefulPopupComponent.destroys += 1));
  }
  press(): void {
    StatefulPopupComponent.owners.push(this.editor());
    this.count.update((n) => n + 1);
  }
}

@Component({
  selector: 'popups-root-host',
  imports: PARTS,
  changeDetection: ChangeDetectionStrategy.OnPush,
  template: `<ng-container docxEditorRoot [document]="bytes" [popups]="popups()">
    <docx-editor-viewport><docx-editor-content /></docx-editor-viewport>
  </ng-container>`,
})
class RootHostComponent {
  readonly bytes = SOURCE;
  readonly popups = signal<DocxEditorPopups | undefined>({});
}

@Component({
  selector: 'popups-two-roots-host',
  imports: PARTS,
  changeDetection: ChangeDetectionStrategy.OnPush,
  template: `@for (key of [0, 1]; track key) {
    <ng-container docxEditorRoot [document]="bytes" [popups]="popups()">
      <docx-editor-viewport><docx-editor-content /></docx-editor-viewport>
    </ng-container>
  }`,
})
class TwoRootsHostComponent {
  readonly bytes = SOURCE;
  readonly popups = signal<DocxEditorPopups>({ hyperlink: StatefulPopupComponent });
  readonly roots = viewChildren(DocxEditorRootDirective);
}

/** A form popup with a draft that starts from its session. */
@Component({
  selector: 'spec-draft-popup',
  template: `<input [value]="draft()" (input)="draft.set($any($event.target).value)" />`,
})
class DraftPopupComponent implements OnInit {
  static setups = 0;
  readonly session = input.required<{ value: string }>();
  readonly draft = signal('');
  ngOnInit(): void {
    DraftPopupComponent.setups += 1;
    this.draft.set(this.session().value);
  }
}

@Component({
  selector: 'popups-outlet-host',
  imports: [DocxPopupOutletComponent],
  changeDetection: ChangeDetectionStrategy.OnPush,
  template: `<docx-popup-outlet
    [entry]="entry()"
    [inputs]="{ session: session() }"
    [session]="session()"
  />`,
})
class OutletHostComponent {
  readonly session = signal({ value: 'First' });
  readonly entry = signal<typeof DraftPopupComponent | { component: typeof DraftPopupComponent }>(
    DraftPopupComponent
  );
}

describe('popup registry', () => {
  afterEach(() => document.body.replaceChildren());

  describe('in a root', () => {
    beforeEach(() => configureZoneless());

    test('the root hosts only configured popups, inside the content, and honors live false entries', async () => {
      const fixture = await render(RootHostComponent);
      const element = fixture.nativeElement as HTMLElement;
      expect(element.querySelector('[data-custom-hyperlink]')).toBeNull();
      fixture.componentInstance.popups.set({ hyperlink: LinkPopupComponent });
      await settle(fixture);
      const matches = element.querySelectorAll('[data-custom-hyperlink]');
      expect(matches.length).toBe(1);
      expect(matches[0]!.closest('.docx-content-mount')).not.toBeNull();
      expect(matches[0]!.closest('.docx-paginated-surface')).toBeNull();
      fixture.componentInstance.popups.set({ hyperlink: false });
      await settle(fixture);
      expect(element.querySelector('[data-custom-hyperlink]')).toBeNull();
    });

    test('an entry with static inputs sets them on the component', async () => {
      const fixture = await render(RootHostComponent);
      fixture.componentInstance.popups.set({
        hyperlink: { component: LinkPopupComponent, inputs: { label: 'Custom' } },
      });
      await settle(fixture);
      const element = fixture.nativeElement as HTMLElement;
      expect(element.querySelector('[data-custom-hyperlink]')?.textContent).toBe('Custom');
    });

    test('components keep their state across new registry objects and stay with their editor', async () => {
      StatefulPopupComponent.setups = 0;
      StatefulPopupComponent.destroys = 0;
      StatefulPopupComponent.owners = [];
      const fixture = await render(TwoRootsHostComponent);
      const element = fixture.nativeElement as HTMLElement;
      const buttons = element.querySelectorAll<HTMLButtonElement>('[data-component-popup]');
      expect(buttons.length).toBe(2);
      expect(StatefulPopupComponent.setups).toBe(2);
      buttons[0]!.click();
      await settle(fixture);
      // A new registry object with the same component keeps each instance.
      fixture.componentInstance.popups.set({ hyperlink: StatefulPopupComponent });
      await settle(fixture);
      expect(element.querySelector('[data-component-popup]')).toBe(buttons[0]);
      expect(buttons[0]!.textContent).toBe('1');
      expect(buttons[1]!.textContent).toBe('0');
      expect(StatefulPopupComponent.setups).toBe(2);
      buttons[1]!.click();
      const [first, second] = fixture.componentInstance.roots();
      expect(StatefulPopupComponent.owners).toEqual([
        first!.context.editor(),
        second!.context.editor(),
      ]);
      expect(StatefulPopupComponent.owners[0]).not.toBe(StatefulPopupComponent.owners[1]);
      fixture.componentInstance.popups.set({ hyperlink: false });
      await settle(fixture);
      expect(element.querySelector('[data-component-popup]')).toBeNull();
      expect(StatefulPopupComponent.destroys).toBe(2);
    });
  });

  describe('provided with DOCX_EDITOR_POPUPS', () => {
    beforeEach(() =>
      configureZoneless([provideDocxEditorPopups({ hyperlink: LinkPopupComponent })])
    );

    test('a root hosts the provided entries, and its own entries replace them', async () => {
      const fixture = await render(RootHostComponent);
      const element = fixture.nativeElement as HTMLElement;
      expect(element.querySelectorAll('[data-custom-hyperlink]').length).toBe(1);
      fixture.componentInstance.popups.set({ hyperlink: false });
      await settle(fixture);
      expect(element.querySelector('[data-custom-hyperlink]')).toBeNull();
    });

    test('outside a root, injectDocxEditorPopups reads the provided registry', () => {
      const popups = TestBed.runInInjectionContext(() => injectDocxEditorPopups());
      expect(popups().hyperlink).toBe(LinkPopupComponent);
    });
  });

  describe('outlet', () => {
    beforeEach(() => configureZoneless());

    test('a new session resets the draft, and a new entry object with the same component keeps it', async () => {
      DraftPopupComponent.setups = 0;
      const fixture = TestBed.createComponent(OutletHostComponent);
      await settle(fixture, 1);
      const element = fixture.nativeElement as HTMLElement;
      const input = element.querySelector('input')!;
      expect(input.value).toBe('First');
      input.value = 'Draft';
      input.dispatchEvent(new Event('input', { bubbles: true }));
      fixture.componentInstance.entry.set({ component: DraftPopupComponent });
      await settle(fixture, 1);
      expect(element.querySelector('input')).toBe(input);
      expect(input.value).toBe('Draft');
      expect(DraftPopupComponent.setups).toBe(1);
      fixture.componentInstance.session.set({ value: 'Second' });
      await settle(fixture, 1);
      expect(element.querySelector('input')).not.toBe(input);
      expect(element.querySelector('input')!.value).toBe('Second');
      expect(DraftPopupComponent.setups).toBe(2);
    });
  });
});
