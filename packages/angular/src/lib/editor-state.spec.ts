import {
  ChangeDetectionStrategy,
  Component,
  effect,
  inject,
  viewChildren,
  type Signal,
} from '@angular/core';
import { blankDocumentBytes } from '@docx-editor.dev/core/editor';
import { configureZoneless, engineOf, render, settle, typeText } from '../testing/harness';
import { DocxEditorContentComponent } from './content.component';
import { DocxEditorContext } from './context';
import { injectEditorEvent, injectEditorState } from './editor-state';
import { DocxEditorRootDirective } from './root.directive';
import { DocxEditorViewportComponent } from './viewport.component';

/** A control that reads the page number, the page count, and bold, and counts notifications. */
@Component({
  selector: 'state-probe',
  template: '',
})
class StateProbeComponent {
  readonly context = inject(DocxEditorContext);
  readonly page = injectEditorState((snapshot) => snapshot.page.current);
  readonly total = injectEditorState((snapshot) => snapshot.page.total);
  readonly bold = injectEditorState((snapshot) => snapshot.formatting?.bold ?? false);
  readonly many: Signal<number>[] = Array.from({ length: 50 }, () =>
    injectEditorState((snapshot) => snapshot.page.current)
  );
  pageNotifications = 0;
  changes = 0;

  constructor() {
    effect(() => {
      this.page();
      this.pageNotifications += 1;
    });
    injectEditorEvent('change', () => (this.changes += 1));
  }
}

@Component({
  imports: [
    DocxEditorRootDirective,
    DocxEditorViewportComponent,
    DocxEditorContentComponent,
    StateProbeComponent,
  ],
  changeDetection: ChangeDetectionStrategy.OnPush,
  template: `
    <ng-container docxEditorRoot [document]="bytes">
      <state-probe />
      <docx-editor-viewport><docx-editor-content /></docx-editor-viewport>
    </ng-container>
  `,
})
class HostComponent {
  readonly bytes = blankDocumentBytes();
  readonly roots = viewChildren(DocxEditorRootDirective);
  readonly probes = viewChildren(StateProbeComponent);
}

describe('injectEditorState', () => {
  beforeEach(() => configureZoneless());

  it('does not notify a page-number reader when the user toggles bold', async () => {
    const fixture = await render(HostComponent);
    const probe = fixture.componentInstance.probes()[0]!;
    const editor = engineOf(probe.context.editor());
    typeText(editor, 'some text');
    await settle(fixture);
    editor.surface!.selectAll();
    await settle(fixture);
    const before = probe.pageNotifications;
    expect(probe.bold()).toBe(false);

    expect(editor.exec({ type: 'toggleMark', mark: 'bold' }).ok).toBe(true);
    await settle(fixture);
    expect(probe.bold()).toBe(true);
    expect(probe.pageNotifications).toBe(before);
    fixture.destroy();
  });

  it('reads the page count after the engine publishes its layout', async () => {
    const fixture = await render(HostComponent);
    const probe = fixture.componentInstance.probes()[0]!;
    const editor = engineOf(probe.context.editor());
    typeText(editor, 'first page');
    await settle(fixture);
    expect(probe.total()).toBe(1);
    expect(editor.exec({ type: 'insertBreak', kind: 'page' }).ok).toBe(true);
    // A break starts a page only for the content after it.
    expect(editor.exec({ type: 'insertText', text: 'second page' }).ok).toBe(true);
    await settle(fixture);
    expect(probe.total()).toBe(editor.snapshot().page.total);
    expect(probe.total()).toBe(2);
    fixture.destroy();
  });

  it('serves fifty readers with one set of engine listeners', async () => {
    const fixture = await render(HostComponent);
    const probe = fixture.componentInstance.probes()[0]!;
    expect(probe.many.every((reader) => reader() === 1)).toBe(true);
    expect(probe.context.engineListenerCount).toBe(4);
    fixture.destroy();
  });

  it('removes an event listener with its component', async () => {
    const fixture = await render(HostComponent);
    const probe = fixture.componentInstance.probes()[0]!;
    const editor = engineOf(probe.context.editor());
    typeText(editor, 'a');
    await settle(fixture);
    expect(probe.changes).toBeGreaterThan(0);
    const counted = probe.changes;
    fixture.destroy();
    editor.exec({ type: 'insertText', text: 'b' });
    expect(probe.changes).toBe(counted);
  });
});

describe('injectEditorState outside a root', () => {
  beforeEach(() => configureZoneless());

  it('reads the empty loading snapshot and does not throw', async () => {
    @Component({ selector: 'lonely-probe', template: '' })
    class LonelyComponent {
      readonly opening = injectEditorState((snapshot) => snapshot.isOpening);
      readonly total = injectEditorState((snapshot) => snapshot.page.total);
    }
    const fixture = await render(LonelyComponent);
    expect(typeof fixture.componentInstance.opening()).toBe('boolean');
    expect(fixture.componentInstance.total()).toBeGreaterThanOrEqual(0);
    fixture.destroy();
  });
});
