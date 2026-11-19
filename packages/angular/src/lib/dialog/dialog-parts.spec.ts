// Task 6.1: the claims of packages/vue/test/dialog-customization.test.ts. The spec dialog below
// has the session contract and the layout of the Field Options dialog (task 6.4), so the claims
// test the parts, not one dialog.
import { NgTemplateOutlet } from '@angular/common';
import {
  ChangeDetectionStrategy,
  Component,
  computed,
  contentChildren,
  effect,
  inject,
  input,
  signal,
  type Type,
} from '@angular/core';
import { TestBed, type ComponentFixture } from '@angular/core/testing';
import { afterEach, beforeEach, describe, expect, test } from 'vitest';
import { configureZoneless, settle } from '../../testing/harness';
import {
  DocxDialogApplyDirective,
  DocxDialogCancelDirective,
  DocxDialogFieldDirective,
  DocxDialogPartDirective,
  DocxDialogParts,
  DocxDialogSlotDirective,
  injectDialog,
  type DocxDialogState,
} from './dialog-parts';
import { DocxNativeDialogComponent } from './native-dialog.component';

interface FieldSession {
  readonly defaultText: string;
  readonly signal: AbortSignal;
  canApply(): boolean;
  apply(text: string): boolean;
  cancel(): void;
}

interface Fields {
  defaultText: string;
}

/** A packaged dialog in miniature: one field, an error, Cancel, and Apply. */
@Component({
  selector: 'spec-field-dialog',
  imports: [
    NgTemplateOutlet,
    DocxNativeDialogComponent,
    DocxDialogSlotDirective,
    DocxDialogApplyDirective,
    DocxDialogCancelDirective,
    DocxDialogFieldDirective,
  ],
  providers: [DocxDialogParts],
  changeDetection: ChangeDetectionStrategy.OnPush,
  template: `
    @if (open()) {
      <docx-native-dialog
        kind="textFormField"
        label="Field options"
        [sessionSignal]="session().signal"
        [dismissOutside]="false"
        (dismiss)="state.cancel()"
      >
        @if (preset()) {
          <div
            *docxDialogSlot="'header'; let content"
            class="docx-dialog__header"
            data-docx-part="header"
          >
            @if (content) {
              <ng-container [ngTemplateOutlet]="content" />
            } @else {
              Field options
            }
          </div>
          <label
            *docxDialogSlot="'field:defaultText'"
            docxDialogField="defaultText"
            class="docx-dialog__row"
          >
            <input
              class="docx-dialog__input"
              [value]="state.values().defaultText"
              (input)="state.setValue('defaultText', $any($event.target).value)"
            />
          </label>
          <span
            *docxDialogSlot="'error'; let content"
            role="alert"
            class="docx-dialog__error"
            data-docx-part="error"
          >
            @if (content) {
              <ng-container [ngTemplateOutlet]="content" />
            } @else {
              {{ state.errors().form ?? '' }}
            }
          </span>
          <button
            *docxDialogSlot="'cancel'; let content"
            docxDialogCancel
            class="docx-dialog__button docx-dialog__cancel"
          >
            @if (content) {
              <ng-container [ngTemplateOutlet]="content" />
            } @else {
              Cancel
            }
          </button>
          <button
            *docxDialogSlot="'apply'; let content"
            docxDialogApply
            class="docx-dialog__button docx-dialog__apply"
          >
            @if (content) {
              <ng-container [ngTemplateOutlet]="content" />
            } @else {
              Apply
            }
          </button>
        }
        <ng-content />
      </docx-native-dialog>
    }
  `,
})
class SpecFieldDialogComponent {
  readonly session = input.required<FieldSession>();
  readonly preset = input(true);

  private readonly values = signal<Fields>({ defaultText: '' });
  private readonly refused = signal(false);
  private readonly closed = signal(false);
  private readonly parts = contentChildren(DocxDialogPartDirective);

  readonly state: DocxDialogState<Fields> = {
    values: this.values.asReadonly(),
    errors: computed(() => (this.refused() ? { form: 'The options are not valid.' } : {})),
    isEnabled: computed(() => !this.closed() && this.session().canApply()),
    setValue: (name, value) => this.values.update((current) => ({ ...current, [name]: value })),
    apply: () => this.refused.set(!this.session().apply(this.values().defaultText)),
    cancel: () => this.session().cancel(),
  };

  readonly open = computed(() => !this.closed());

  constructor() {
    inject(DocxDialogParts, { self: true }).connect(
      this.state as unknown as DocxDialogState<object>,
      this.parts
    );
    effect((onCleanup) => {
      const session = this.session();
      this.values.set({ defaultText: session.defaultText });
      this.refused.set(false);
      this.closed.set(session.signal.aborted);
      const abort = () => this.closed.set(true);
      session.signal.addEventListener('abort', abort, { once: true });
      onCleanup(() => session.signal.removeEventListener('abort', abort));
    });
  }
}

/** A host component in a part template that reads the dialog state. */
@Component({
  selector: 'spec-draft-echo',
  template: `<output data-echo>{{ dialog.values().defaultText }}</output>`,
})
class DraftEchoComponent {
  readonly dialog = injectDialog<Fields>();
}

function session(accept: boolean) {
  const controller = new AbortController();
  const saved: string[] = [];
  const value: FieldSession = {
    defaultText: 'Initial',
    signal: controller.signal,
    canApply: () => !controller.signal.aborted,
    cancel: () => controller.abort(),
    apply: (text) => {
      saved.push(text);
      if (accept) controller.abort();
      return accept;
    },
  };
  return { controller, saved, value };
}

const IMPORTS = [
  SpecFieldDialogComponent,
  DocxDialogPartDirective,
  DocxDialogApplyDirective,
  DocxDialogCancelDirective,
  DraftEchoComponent,
];

@Component({
  selector: 'parts-as-child-host',
  imports: IMPORTS,
  template: `<spec-field-dialog [session]="session">
    <ng-template docxDialogPart="apply" asChild>
      <button docxDialogApply data-save>Save field</button>
    </ng-template>
  </spec-field-dialog>`,
})
class AsChildHostComponent {
  session!: FieldSession;
}

@Component({
  selector: 'parts-class-host',
  imports: IMPORTS,
  template: `<spec-field-dialog [session]="session">
    <ng-template docxDialogPart="apply" [asChild]="asChild" className="paired-class">
      @if (asChild) {
        <button docxDialogApply class="app-action app-active">Save field</button>
      } @else {
        Save field
      }
    </ng-template>
  </spec-field-dialog>`,
})
class ClassHostComponent {
  session!: FieldSession;
  asChild = false;
}

@Component({
  selector: 'parts-attributes-host',
  imports: IMPORTS,
  template: `<spec-field-dialog [session]="session">
    <ng-template
      docxDialogPart="apply"
      keepContent
      className="app-action"
      [styles]="{ color: 'red' }"
      [attributes]="{
        id: 'app-apply',
        'aria-label': 'Save the field',
        'data-testid': 'app-save',
        'data-docx-part': 'other',
        type: 'submit',
        onclick: 'alert(1)',
      }"
    />
  </spec-field-dialog>`,
})
class AttributesHostComponent {
  session!: FieldSession;
}

@Component({
  selector: 'parts-content-host',
  imports: IMPORTS,
  template: `<spec-field-dialog [session]="session">
    <ng-template docxDialogPart="apply"
      ><strong>{{ label() }}</strong> field</ng-template
    >
    <ng-template docxDialogPart="header" hidden />
    <ng-template docxDialogPart="cancel"><spec-draft-echo /></ng-template>
  </spec-field-dialog>`,
})
class ContentHostComponent {
  session!: FieldSession;
  readonly label = signal('Save');
}

@Component({
  selector: 'parts-verbatim-host',
  imports: IMPORTS,
  template: `<spec-field-dialog [session]="session" [preset]="false">
    <p data-own>Own layout</p>
    <button docxDialogCancel data-own-cancel>Close</button>
  </spec-field-dialog>`,
})
class VerbatimHostComponent {
  session!: FieldSession;
}

async function mount<T extends { session: FieldSession }>(
  component: Type<T>,
  accept: boolean,
  setup: (host: T) => void = () => {}
): Promise<{ fixture: ComponentFixture<T>; element: HTMLElement } & ReturnType<typeof session>> {
  const created = session(accept);
  const fixture = TestBed.createComponent(component);
  fixture.componentInstance.session = created.value;
  setup(fixture.componentInstance);
  document.body.append(fixture.nativeElement);
  await settle(fixture, 1);
  return { fixture, element: fixture.nativeElement as HTMLElement, ...created };
}

function type(element: HTMLElement, value: string): void {
  const input = element.querySelector<HTMLInputElement>('[data-docx-field="defaultText"] input')!;
  input.value = value;
  input.dispatchEvent(new Event('input', { bubbles: true }));
}

describe('dialog parts', () => {
  beforeEach(() => configureZoneless());
  afterEach(() => document.body.replaceChildren());

  test('a custom Apply keeps the field wiring and closes only after an accepted write', async () => {
    const { fixture, element, saved } = await mount(AsChildHostComponent, true);
    type(element, 'Updated');
    await settle(fixture, 1);
    element.querySelector<HTMLButtonElement>('[data-save]')!.click();
    await settle(fixture, 1);
    expect(saved).toEqual(['Updated']);
    expect(element.querySelector('dialog')).toBeNull();
  });

  test('a refused write keeps the dialog open with an error, and an abort closes it', async () => {
    const { fixture, element, controller } = await mount(AsChildHostComponent, false);
    element.querySelector<HTMLButtonElement>('[data-save]')!.click();
    await settle(fixture, 1);
    expect(element.querySelector('[role="alert"]')?.textContent?.trim()).not.toBe('');
    expect(element.querySelector('dialog')).not.toBeNull();
    controller.abort();
    await settle(fixture, 1);
    expect(element.querySelector('dialog')).toBeNull();
  });

  for (const asChild of [false, true]) {
    test(`host classes merge with the packaged classes (asChild=${asChild})`, async () => {
      const { fixture, element, saved } = await mount(ClassHostComponent, true, (host) => {
        host.asChild = asChild;
      });
      const button = element.querySelector<HTMLButtonElement>('[data-docx-part="apply"]')!;
      expect(button.textContent?.trim()).toBe('Save field');
      const expected = asChild ? ['app-action', 'app-active'] : ['paired-class'];
      for (const name of expected) expect(button.classList.contains(name)).toBe(true);
      // `asChild` hands the look to the host: the packaged classes stay off its element.
      expect(button.classList.contains('docx-dialog__button')).toBe(!asChild);
      button.click();
      await settle(fixture, 1);
      expect(saved).toHaveLength(1);
    });
  }

  test('a part forwards host attributes but keeps its own wiring', async () => {
    const { fixture, element, saved } = await mount(AttributesHostComponent, true);
    const button = element.querySelector<HTMLButtonElement>('[data-docx-part="apply"]')!;
    expect(button.id).toBe('app-apply');
    expect(button.getAttribute('aria-label')).toBe('Save the field');
    expect(button.getAttribute('data-testid')).toBe('app-save');
    expect(button.getAttribute('type')).toBe('button');
    expect(button.hasAttribute('onclick')).toBe(false);
    expect(button.classList.contains('app-action')).toBe(true);
    expect(button.classList.contains('docx-dialog__button')).toBe(true);
    expect(button.style.color).toBe('red');
    // `keepContent` keeps the packaged label.
    expect(button.textContent?.trim()).toBe('Apply');
    button.click();
    await settle(fixture, 1);
    expect(saved).toHaveLength(1);
  });

  test('replacement content keeps its markup, a live label, the element, and the wiring', async () => {
    const { fixture, element, saved } = await mount(ContentHostComponent, false);
    const button = element.querySelector<HTMLButtonElement>('[data-docx-part="apply"]')!;
    expect(button.textContent?.trim()).toBe('Save field');
    expect(button.querySelector('strong')?.textContent).toBe('Save');
    fixture.componentInstance.label.set('Update');
    await settle(fixture, 1);
    expect(element.querySelector('[data-docx-part="apply"]')).toBe(button);
    expect(button.textContent?.trim()).toBe('Update field');
    button.click();
    await settle(fixture, 1);
    expect(saved).toHaveLength(1);
    expect(element.querySelector('[role="alert"]')?.textContent?.trim()).not.toBe('');
  });

  test('a hidden part leaves, and a component in a part template reads the dialog state', async () => {
    const { fixture, element } = await mount(ContentHostComponent, false);
    expect(element.querySelector('[data-docx-part="header"]')).toBeNull();
    type(element, 'Echoed');
    await settle(fixture, 1);
    expect(element.querySelector('[data-echo]')?.textContent).toBe('Echoed');
  });

  test('without the preset, the dialog renders only its content', async () => {
    const { fixture, element, controller } = await mount(VerbatimHostComponent, true);
    expect(element.querySelector('[data-own]')).not.toBeNull();
    expect(element.querySelector('[data-docx-part="apply"]')).toBeNull();
    element.querySelector<HTMLButtonElement>('[data-own-cancel]')!.click();
    await settle(fixture, 1);
    expect(controller.signal.aborted).toBe(true);
    expect(element.querySelector('dialog')).toBeNull();
  });
});

describe('native dialog', () => {
  beforeEach(() => configureZoneless());
  afterEach(() => document.body.replaceChildren());

  test('it focuses the first control, closes on Escape, and gives the focus back', async () => {
    const opener = document.createElement('button');
    document.body.append(opener);
    opener.focus();
    const { fixture, element, controller } = await mount(AsChildHostComponent, true);
    const dialog = element.querySelector('dialog')!;
    expect(dialog.hasAttribute('open')).toBe(true);
    expect(dialog.getAttribute('aria-modal')).toBe('true');
    expect(dialog.getAttribute('data-docx-dialog')).toBe('textFormField');
    expect(document.activeElement).toBe(element.querySelector('input'));
    dialog.dispatchEvent(
      new KeyboardEvent('keydown', { bubbles: true, cancelable: true, key: 'Escape' })
    );
    await settle(fixture, 1);
    expect(controller.signal.aborted).toBe(true);
    expect(element.querySelector('dialog')).toBeNull();
    expect(document.activeElement).toBe(opener);
  });

  test('outside an editor scope it adds the scope class itself', async () => {
    const { element } = await mount(AsChildHostComponent, true);
    const dialog = element.querySelector('dialog')!;
    expect(dialog.classList.contains('docx-dialog')).toBe(true);
    expect(dialog.classList.contains('docx-editor')).toBe(true);
  });
});
