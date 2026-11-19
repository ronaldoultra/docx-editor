import { ChangeDetectionStrategy, Component, signal, viewChildren } from '@angular/core';
import { blankDocumentBytes } from '@docx-editor.dev/core/editor';
import { en } from '@docx-editor.dev/i18n';
import { configureZoneless, engineOf, render, settle, typeText } from '../testing/harness';
import { DocxEditorContentComponent } from './content.component';
import { DocxEditorRootDirective } from './root.directive';
import { DocxEditorViewportComponent } from './viewport.component';

@Component({
  imports: [DocxEditorRootDirective, DocxEditorViewportComponent, DocxEditorContentComponent],
  changeDetection: ChangeDetectionStrategy.OnPush,
  template: `
    <ng-container docxEditorRoot [document]="bytes">
      <docx-editor-viewport class="host-class">
        @if (showContent()) {
          <docx-editor-content />
        }
      </docx-editor-viewport>
    </ng-container>
  `,
})
class HostComponent {
  readonly bytes = blankDocumentBytes();
  readonly showContent = signal(true);
  readonly roots = viewChildren(DocxEditorRootDirective);
}

describe('docx-editor-viewport and docx-editor-content', () => {
  beforeEach(() => configureZoneless());

  it('renders the scroller with the classes and the region role that the engine expects', async () => {
    const fixture = await render(HostComponent);
    const scroller = (fixture.nativeElement as HTMLElement).querySelector('docx-editor-viewport')!;
    expect(scroller.classList.contains('docx-editor__scroll-container')).toBe(true);
    expect(scroller.classList.contains('docx-editor')).toBe(true);
    expect(scroller.classList.contains('host-class')).toBe(true);
    expect(scroller.getAttribute('role')).toBe('region');
    expect(scroller.getAttribute('aria-label')).toBe(en.editor.documentViewport);
    expect(scroller.querySelector('.docx-paginated-surface')).not.toBeNull();
    fixture.destroy();
  });

  it('keeps undo history, selection, and the document when the content leaves and comes back', async () => {
    const fixture = await render(HostComponent);
    const view = fixture.componentInstance;
    const editor = engineOf(view.roots()[0]!.context.editor());
    typeText(editor, 'kept');
    await settle(fixture);
    const selection = editor.snapshot().selection;
    expect(editor.snapshot().canUndo).toBe(true);

    view.showContent.set(false);
    await settle(fixture);
    expect(view.roots()[0]!.context.editor()).toBe(editor);
    expect(view.roots()[0]!.context.contentElement()).toBeNull();
    expect(editor.snapshot().canUndo).toBe(true);

    view.showContent.set(true);
    await settle(fixture);
    const host = fixture.nativeElement as HTMLElement;
    await vi.waitFor(() => expect(host.querySelector('.docx-page')).not.toBeNull());
    expect(editor.snapshot().canUndo).toBe(true);
    expect(editor.snapshot().selection).toEqual(selection);
    expect(editor.exec({ type: 'undo' }).ok).toBe(true);
    fixture.destroy();
  });

  it('zooms the pages with Ctrl and the plus key', async () => {
    const fixture = await render(HostComponent);
    const editor = engineOf(fixture.componentInstance.roots()[0]!.context.editor());
    const scroller = (fixture.nativeElement as HTMLElement).querySelector('docx-editor-viewport')!;
    const zoom = editor.getZoom();
    scroller.dispatchEvent(
      new KeyboardEvent('keydown', { key: '+', ctrlKey: true, bubbles: true, cancelable: true })
    );
    expect(editor.getZoom()).toBeGreaterThan(zoom);
    fixture.destroy();
  });
});
