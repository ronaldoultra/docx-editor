import { ChangeDetectionStrategy, Component, provideZonelessChangeDetection } from '@angular/core';
import { TestBed } from '@angular/core/testing';
import { blankDocumentBytes } from '@docx-editor.dev/core/editor';
import { DocxEditorComponent } from './docx-editor.component';

@Component({
  imports: [DocxEditorComponent],
  changeDetection: ChangeDetectionStrategy.OnPush,
  template: `<docx-editor [document]="bytes" />`,
})
class HostComponent {
  readonly bytes = blankDocumentBytes();
}

describe('DocxEditorComponent', () => {
  beforeEach(() => {
    TestBed.configureTestingModule({ providers: [provideZonelessChangeDetection()] });
  });

  it('mounts a document and paints its first page', async () => {
    const fixture = TestBed.createComponent(HostComponent);
    await fixture.whenStable();
    const host = fixture.nativeElement as HTMLElement;
    const region = host.querySelector('.docx-editor__scroll-container');
    expect(region?.getAttribute('role')).toBe('region');
    await vi.waitFor(() => expect(host.querySelector('.docx-page')).not.toBeNull());
    fixture.destroy();
  });
});
