// Zone handling: engine work runs outside the Angular zone, with zone.js and without it.
// This file loads zone.js. Vitest gives each spec file its own environment, so the other spec
// files stay zoneless.
import 'zone.js';
import 'zone.js/testing';
import {
  ApplicationRef,
  ChangeDetectionStrategy,
  Component,
  inject,
  provideZoneChangeDetection,
  provideZonelessChangeDetection,
  viewChild,
} from '@angular/core';
import { TestBed } from '@angular/core/testing';
import { blankDocumentBytes } from '@docx-editor.dev/core/editor';
import { engineOf, settle, typeText } from '../testing/harness';
import { DocxEditorContentComponent } from './content.component';
import { DocxEditorContext } from './context';
import { injectEditorCommand } from './editor-command';
import { DocxEditorRootDirective } from './root.directive';
import { DocxEditorViewportComponent } from './viewport.component';

@Component({
  selector: 'bold-label',
  template: `<output data-testid="bold">{{ bold.isActive() ? 'bold' : 'plain' }}</output>`,
  changeDetection: ChangeDetectionStrategy.OnPush,
})
class BoldLabelComponent {
  readonly context = inject(DocxEditorContext);
  readonly bold = injectEditorCommand('text.bold');
}

@Component({
  imports: [
    DocxEditorRootDirective,
    DocxEditorViewportComponent,
    DocxEditorContentComponent,
    BoldLabelComponent,
  ],
  changeDetection: ChangeDetectionStrategy.OnPush,
  template: `
    <ng-container docxEditorRoot [document]="bytes">
      <bold-label />
      <docx-editor-viewport><docx-editor-content /></docx-editor-viewport>
    </ng-container>
  `,
})
class HostComponent {
  readonly bytes = blankDocumentBytes();
  readonly label = viewChild.required(BoldLabelComponent);
}

async function mountAndBold() {
  const fixture = TestBed.createComponent(HostComponent);
  fixture.autoDetectChanges();
  await settle(fixture);
  const host = fixture.nativeElement as HTMLElement;
  const label = () => host.querySelector('[data-testid="bold"]')!.textContent;
  const editor = engineOf(fixture.componentInstance.label().context.editor());
  typeText(editor, 'some text');
  editor.surface!.selectAll();
  await settle(fixture);
  expect(label()).toBe('plain');
  // The toggle runs where a host's own code runs: outside any Angular task.
  editor.exec({ type: 'toggleMark', mark: 'bold' });
  await settle(fixture);
  return { fixture, host, label, editor };
}

describe('with zone.js', () => {
  beforeEach(() =>
    TestBed.configureTestingModule({
      providers: [provideZoneChangeDetection({ eventCoalescing: true })],
    })
  );

  it('updates the controls after a committed change', async () => {
    const { fixture, label } = await mountAndBold();
    expect(label()).toBe('bold');
    fixture.destroy();
  });

  it('does not run change detection for each DOM event on the pages', async () => {
    const { fixture, host } = await mountAndBold();
    const ticks = vi.spyOn(TestBed.inject(ApplicationRef), 'tick');
    const surface = host.querySelector<HTMLElement>('.docx-paginated-surface')!;
    for (let index = 0; index < 20; index += 1) {
      surface.dispatchEvent(new MouseEvent('mousemove', { bubbles: true, clientX: index }));
      surface.dispatchEvent(new KeyboardEvent('keyup', { bubbles: true, key: 'Shift' }));
    }
    await settle(fixture);
    // Only the state that the events publish may run change detection, not the 40 events.
    expect(ticks.mock.calls.length).toBeLessThanOrEqual(2);
    fixture.destroy();
  });
});

describe('without zone.js', () => {
  beforeEach(() =>
    TestBed.configureTestingModule({ providers: [provideZonelessChangeDetection()] })
  );

  it('updates the controls after a committed change', async () => {
    const { fixture, label } = await mountAndBold();
    expect(label()).toBe('bold');
    fixture.destroy();
  });
});
