// The claims of packages/vue/test/use-editor-command.test.ts, for the Angular functions.
import {
  ChangeDetectionStrategy,
  Component,
  ElementRef,
  effect,
  inject,
  signal,
  viewChild,
  viewChildren,
} from '@angular/core';
import type { ExecResult } from '@docx-editor.dev/core/contracts/editor';
import { blankDocumentBytes } from '@docx-editor.dev/core/editor';
import { configureZoneless, engineOf, render, settle, typeText } from '../testing/harness';
import { DocxEditorContentComponent } from './content.component';
import { DocxEditorContext } from './context';
import {
  injectEditorCommand,
  injectEditorValueCommand,
  injectHistoryGroup,
} from './editor-command';
import { injectEditorState } from './editor-state';
import { DocxEditorRootDirective } from './root.directive';
import { DocxEditorViewportComponent } from './viewport.component';

@Component({ selector: 'command-probe', template: '' })
class CommandProbeComponent {
  readonly context = inject(DocxEditorContext);
  readonly mark = signal('bold');
  readonly toggle = injectEditorCommand(() => ({ type: 'toggleMark', mark: this.mark() }));
  readonly cut = injectEditorCommand({ type: 'cut' });
  readonly bold = injectEditorCommand('text.bold');
  readonly inline = Array.from({ length: 10 }, () => injectEditorCommand({ type: 'selectAll' }));
  readonly page = injectEditorState(
    (snapshot) => snapshot.page,
    (a, b) => a.current === b.current && a.total === b.total
  );
  readonly formattingBold = injectEditorState((snapshot) => snapshot.formatting?.bold ?? false);
  pageNotifications = 0;
  revisions = 0;

  constructor() {
    effect(() => {
      this.page();
      this.pageNotifications += 1;
    });
    effect(() => {
      this.context.revision();
      this.revisions += 1;
    });
  }
}

@Component({
  selector: 'range-probe',
  template: `<input #range data-live-size type="range" min="16" max="100" (input)="onInput()" />`,
})
class RangeProbeComponent {
  readonly size = injectEditorValueCommand('font.size');
  readonly range = viewChild.required<ElementRef<HTMLInputElement>>('range');
  readonly gesture = injectHistoryGroup({ kind: 'range', element: this.range });
  readonly results: ExecResult[] = [];

  onInput(): void {
    const value = Number(this.range().nativeElement.value);
    this.results.push(this.size.execute(value, this.gesture.options()));
  }
}

@Component({
  selector: 'stepper-probe',
  template: `<button #stepper data-stepper (keydown)="step()">Increase</button>`,
})
class StepperProbeComponent {
  readonly size = injectEditorValueCommand('font.size');
  readonly stepper = viewChild.required<ElementRef<HTMLElement>>('stepper');
  readonly gesture = injectHistoryGroup({ kind: 'repeat', element: this.stepper });
  readonly results: ExecResult[] = [];
  private value = 22;

  step(): void {
    this.value += 2;
    this.results.push(this.size.execute(this.value, this.gesture.options()));
  }
}

@Component({
  imports: [
    DocxEditorRootDirective,
    DocxEditorViewportComponent,
    DocxEditorContentComponent,
    CommandProbeComponent,
    RangeProbeComponent,
    StepperProbeComponent,
  ],
  changeDetection: ChangeDetectionStrategy.OnPush,
  template: `
    <ng-container docxEditorRoot [document]="bytes">
      <command-probe />
      @if (gestures()) {
        <range-probe />
        <stepper-probe />
      }
      <docx-editor-viewport><docx-editor-content /></docx-editor-viewport>
    </ng-container>
  `,
})
class HostComponent {
  readonly bytes = blankDocumentBytes();
  readonly gestures = signal(true);
  readonly roots = viewChildren(DocxEditorRootDirective);
  readonly probe = viewChild.required(CommandProbeComponent);
  readonly range = viewChild(RangeProbeComponent);
  readonly stepper = viewChild(StepperProbeComponent);
}

async function boldDocument() {
  const fixture = await render(HostComponent);
  const view = fixture.componentInstance;
  const editor = engineOf(view.probe().context.editor());
  typeText(editor, 'bold text');
  editor.surface!.selectAll();
  editor.exec({ type: 'toggleMark', mark: 'bold' });
  await settle(fixture);
  return { fixture, view, editor };
}

describe('injectEditorCommand with an engine command', () => {
  beforeEach(() => configureZoneless());

  it('derives the state again when the payload changes but the type does not', async () => {
    const { fixture, view } = await boldDocument();
    const probe = view.probe();
    expect(probe.toggle.isActive()).toBe(true);
    probe.mark.set('italic');
    await settle(fixture);
    expect(probe.toggle.isActive()).toBe(false);
    fixture.destroy();
  });

  it('lets an engine event that a state read publishes write a signal', async () => {
    // A state read can flush pending input, and the engine then publishes selectionChange at
    // once. A listener that writes a signal must not see a reactive context (NG0600).
    const fixture = await render(HostComponent);
    const probe = fixture.componentInstance.probe();
    const editor = engineOf(probe.context.editor());
    const written = signal(0);
    const isActive = editor.isActive.bind(editor);
    editor.isActive = (command) => {
      written.update((count) => count + 1);
      return isActive(command);
    };
    probe.mark.set('italic');
    expect(() => probe.toggle.isActive()).not.toThrow();
    expect(written()).toBeGreaterThan(0);
    fixture.destroy();
  });

  it('returns false from execute when the engine refuses, and changes nothing', async () => {
    const fixture = await render(HostComponent);
    const probe = fixture.componentInstance.probe();
    const editor = engineOf(probe.context.editor());
    const before = editor.surface!.session.revision();
    expect(probe.cut.execute()).toBe(false);
    expect(editor.surface!.session.revision()).toBe(before);
    fixture.destroy();
  });
});

describe('injectEditorCommand and injectEditorState subscriptions', () => {
  beforeEach(() => configureZoneless());

  it('serves many commands with the root listeners only', async () => {
    const fixture = await render(HostComponent);
    expect(fixture.componentInstance.probe().context.engineListenerCount).toBe(4);
    fixture.destroy();
  });

  it('does not notify when the custom equality says the slice is the same', async () => {
    const fixture = await render(HostComponent);
    const probe = fixture.componentInstance.probe();
    const before = probe.pageNotifications;
    engineOf(probe.context.editor()).exec({ type: 'insertText', text: 'z' });
    await settle(fixture);
    expect(probe.pageNotifications).toBe(before);
    fixture.destroy();
  });

  it('collapses a burst of changes into at most two notifications', async () => {
    const fixture = await render(HostComponent);
    const probe = fixture.componentInstance.probe();
    const editor = engineOf(probe.context.editor());
    const before = probe.revisions;
    for (let index = 0; index < 5; index += 1) editor.exec({ type: 'insertText', text: 'a' });
    await settle(fixture);
    expect(probe.revisions - before).toBeLessThanOrEqual(2);
    fixture.destroy();
  });
});

describe('injectEditorCommand outside a root', () => {
  beforeEach(() => configureZoneless());

  it('is disabled, and execute returns false', async () => {
    @Component({ selector: 'lonely-command', template: '' })
    class LonelyComponent {
      readonly bold = injectEditorCommand('text.bold');
      readonly size = injectEditorValueCommand('font.size');
    }
    const fixture = await render(LonelyComponent);
    const lonely = fixture.componentInstance;
    expect(lonely.bold.execute()).toBe(false);
    expect(lonely.bold.isEnabled()).toBe(false);
    expect(lonely.size.isEnabled()).toBe(false);
    expect(lonely.size.value()).toBeNull();
    fixture.destroy();
  });
});

describe('injectHistoryGroup', () => {
  beforeEach(() => configureZoneless());

  it('groups keyboard repeats on a range into one undo step each, and cleans up', async () => {
    const { fixture, view, editor } = await boldDocument();
    const range = view.range()!;
    const input = range.range().nativeElement;
    for (const values of [
      [24, 28, 32],
      [40, 44],
    ]) {
      for (const [index, value] of values.entries()) {
        input.dispatchEvent(
          new KeyboardEvent('keydown', { bubbles: true, key: 'ArrowUp', repeat: index > 0 })
        );
        input.value = String(value);
        input.dispatchEvent(new Event('input', { bubbles: true }));
        await settle(fixture, 1);
      }
      input.dispatchEvent(new KeyboardEvent('keyup', { bubbles: true, key: 'ArrowUp' }));
      await settle(fixture, 1);
    }
    expect(range.results.map((r) => (r.ok ? r.history?.kind : r.reason))).toEqual([
      'started',
      'extended',
      'extended',
      'started',
      'extended',
    ]);
    editor.exec({ type: 'undo' });
    await settle(fixture);
    expect(range.size.value()).toBe(32);
    const handle = range.gesture.options().historyGroup!;
    view.gestures.set(false);
    await settle(fixture);
    expect(handle.state).toBe('closed');
    expect(() => range.gesture.options()).toThrow('not mounted');
    fixture.destroy();
  });

  it('starts a held stepper before its own keydown handler writes', async () => {
    const { fixture, view, editor } = await boldDocument();
    const stepper = view.stepper()!;
    const button = stepper.stepper().nativeElement;
    for (const repeat of [false, true, true]) {
      button.dispatchEvent(new KeyboardEvent('keydown', { bubbles: true, key: 'ArrowUp', repeat }));
      await settle(fixture, 1);
    }
    button.dispatchEvent(new KeyboardEvent('keyup', { bubbles: true, key: 'ArrowUp' }));
    await settle(fixture, 1);
    expect(stepper.results.map((r) => (r.ok ? r.history?.kind : r.reason))).toEqual([
      'started',
      'extended',
      'extended',
    ]);
    const undo = editor.exec({ type: 'undo' });
    expect(undo.ok).toBe(true);
    fixture.destroy();
  });
});
