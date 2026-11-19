import { ChangeDetectionStrategy, Component, PLATFORM_ID, viewChildren } from '@angular/core';

import { blankDocumentBytes, type DocxEditorInstance } from '@docx-editor.dev/core/editor';
import { configureZoneless, engineOf, render, settle, typeText } from '../testing/harness';
import { DocxEditorContentComponent } from './content.component';
import { DocxEditorRootDirective } from './root.directive';
import { DocxEditorViewportComponent } from './viewport.component';

const PARTS = [DocxEditorRootDirective, DocxEditorViewportComponent, DocxEditorContentComponent];

@Component({
  imports: PARTS,
  changeDetection: ChangeDetectionStrategy.OnPush,
  template: `
    @for (bytes of documents; track $index) {
      <section docxEditorRoot [document]="bytes">
        <docx-editor-viewport><docx-editor-content /></docx-editor-viewport>
      </section>
    }
  `,
})
class TwoRootsComponent {
  readonly documents = [blankDocumentBytes(), blankDocumentBytes()];
  readonly roots = viewChildren(DocxEditorRootDirective);
}

@Component({
  imports: PARTS,
  changeDetection: ChangeDetectionStrategy.OnPush,
  template: `
    <ng-container docxEditorRoot [document]="bytes">
      <docx-editor-viewport><docx-editor-content /></docx-editor-viewport>
    </ng-container>
  `,
})
class OneRootComponent {
  readonly bytes = blankDocumentBytes();
  readonly roots = viewChildren(DocxEditorRootDirective);
}

describe('docxEditorRoot', () => {
  beforeEach(() => configureZoneless());

  it('gives two roots two independent engines', async () => {
    const fixture = await render(TwoRootsComponent);
    const [first, second] = fixture.componentInstance.roots().map((root) => root.context.editor());
    expect(first).toBeTruthy();
    expect(second).toBeTruthy();
    expect(first).not.toBe(second);
    const before = engineOf(second).snapshot();
    typeText(engineOf(first), 'typed in the first');
    await settle(fixture);
    expect(engineOf(first).snapshot()).not.toBe(before);
    expect(engineOf(second).snapshot()).toBe(before);
    fixture.destroy();
  });

  it('paints the pages inside the content element, on an ng-container root', async () => {
    const fixture = await render(OneRootComponent);
    const host = fixture.nativeElement as HTMLElement;
    const content = host.querySelector('docx-editor-content');
    await vi.waitFor(() => expect(content?.querySelector('.docx-page')).not.toBeNull());
    fixture.destroy();
  });

  it('destroys its engine and removes its listeners when Angular destroys it', async () => {
    const fixture = await render(OneRootComponent);
    const root = fixture.componentInstance.roots()[0]!;
    const editor = engineOf(root.context.editor());
    const destroy = vi.spyOn(editor as DocxEditorInstance, 'destroy');
    expect(root.context.engineListenerCount).toBeGreaterThan(0);
    fixture.destroy();
    expect(destroy).toHaveBeenCalledTimes(1);
    expect(root.context.engineListenerCount).toBe(0);
    expect(root.context.editor()).toBeNull();
  });

  it('emits ready once and change for each committed edit', async () => {
    @Component({
      imports: PARTS,
      template: `
        <ng-container
          docxEditorRoot
          [document]="bytes"
          (ready)="ready = ready + 1"
          (change)="changes = changes + 1"
        >
          <docx-editor-viewport><docx-editor-content /></docx-editor-viewport>
        </ng-container>
      `,
    })
    class EventsComponent {
      readonly bytes = blankDocumentBytes();
      readonly roots = viewChildren(DocxEditorRootDirective);
      ready = 0;
      changes = 0;
    }
    const fixture = await render(EventsComponent);
    const view = fixture.componentInstance;
    expect(view.ready).toBe(1);
    const before = view.changes;
    typeText(engineOf(view.roots()[0]!.context.editor()), 'a');
    await settle(fixture);
    expect(view.changes).toBeGreaterThan(before);
    expect(view.ready).toBe(1);
    fixture.destroy();
  });
});

describe('docxEditorRoot on the server', () => {
  beforeEach(() => configureZoneless([{ provide: PLATFORM_ID, useValue: 'server' }]));

  it('renders the viewport structure and creates no engine', async () => {
    const fixture = await render(OneRootComponent);
    const host = fixture.nativeElement as HTMLElement;
    expect(host.querySelector('.docx-editor__scroll-container')).not.toBeNull();
    expect(host.querySelector('.docx-paginated-surface')).not.toBeNull();
    expect(fixture.componentInstance.roots()[0]!.context.editor()).toBeNull();
    expect(host.querySelector('.docx-page')).toBeNull();
    fixture.destroy();
  });
});
