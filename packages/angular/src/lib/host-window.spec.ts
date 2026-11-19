// The `DocxEditorWindow` claims of packages/vue/test/host-window-chrome.test.ts, the pop-out claim
// of root-color-mode.test.ts, and the host-window resolution of the editor. The picker claim of
// host-window-chrome.test.ts needs the font family control, so it ports with the toolbar (4.2).
import { ChangeDetectionStrategy, Component, signal, viewChild } from '@angular/core';
import { TestBed } from '@angular/core/testing';
import { blankDocumentBytes } from '@docx-editor.dev/core/editor';
import { configureZoneless, engineOf, settle, typeText } from '../testing/harness';
import type { DocxEditorColorMode } from './color-mode';
import { DocxEditorContentComponent } from './content.component';
import { injectEditorHostDocument } from './editor-state';
import { DocxEditorRootDirective } from './root.directive';
import { DocxEditorViewportComponent } from './viewport.component';
import { DocxEditorWindowComponent } from './window.component';

@Component({ selector: 'host-document-probe', template: '' })
class HostDocumentProbeComponent {
  readonly hostDocument = injectEditorHostDocument();
}

@Component({
  imports: [
    DocxEditorRootDirective,
    DocxEditorViewportComponent,
    DocxEditorContentComponent,
    DocxEditorWindowComponent,
    HostDocumentProbeComponent,
  ],
  changeDetection: ChangeDetectionStrategy.OnPush,
  template: `
    <ng-container docxEditorRoot [document]="bytes" [colorMode]="colorMode()">
      <host-document-probe />
      <docx-editor-window [(open)]="open" (close)="closed = closed + 1">
        <ng-template>
          <p data-testid="panel">panel</p>
          @if (withPages()) {
            <docx-editor-viewport><docx-editor-content /></docx-editor-viewport>
          }
        </ng-template>
      </docx-editor-window>
    </ng-container>
  `,
})
class HostComponent {
  readonly bytes = blankDocumentBytes();
  readonly open = signal(true);
  readonly withPages = signal(false);
  readonly colorMode = signal<DocxEditorColorMode | undefined>(undefined);
  readonly root = viewChild.required(DocxEditorRootDirective);
  readonly probe = viewChild.required(HostDocumentProbeComponent);
  closed = 0;
}

/** A same-origin frame that stands in for the pop-out window. */
function stubPopout() {
  const frame = document.createElement('iframe');
  document.body.append(frame);
  const view = frame.contentWindow as Window & typeof globalThis;
  // The frame has its own canvas prototype. Give it the same stub as src/test-setup.ts.
  view.HTMLCanvasElement.prototype.getContext = (() =>
    null) as typeof HTMLCanvasElement.prototype.getContext;
  const closes = { count: 0, onClose: () => {} };
  view.close = () => {
    closes.count += 1;
    closes.onClose();
  };
  const original = window.open;
  window.open = (() => view) as typeof window.open;
  const restore = () => {
    window.open = original;
    frame.remove();
  };
  return { view, closes, restore };
}

async function mount(setup?: (view: HostComponent) => void) {
  const popout = stubPopout();
  const fixture = TestBed.createComponent(HostComponent);
  setup?.(fixture.componentInstance);
  await settle(fixture);
  const host = fixture.nativeElement as HTMLElement;
  return {
    fixture,
    popout,
    view: fixture.componentInstance,
    panelHere: () => host.querySelector('[data-testid="panel"]'),
    panelThere: () => popout.view.document.querySelector('[data-testid="panel"]'),
  };
}

describe('docx-editor-window', () => {
  beforeEach(() => configureZoneless());
  const cleanups: Array<() => void> = [];
  afterEach(() => {
    for (const cleanup of cleanups.splice(0)) cleanup();
  });

  it('renders in the pop-out while open, and back in place when the window closes', async () => {
    const { fixture, popout, view, panelHere, panelThere } = await mount();
    cleanups.push(() => fixture.destroy(), popout.restore);
    expect(panelThere()).not.toBeNull();
    expect(panelHere()).toBeNull();

    popout.view.dispatchEvent(new popout.view.Event('pagehide'));
    await settle(fixture);
    expect(view.closed).toBe(1);
    expect(view.open()).toBe(false);
    expect(panelHere()).not.toBeNull();
  });

  it('moves the panel back before the window closes when open turns off', async () => {
    const { fixture, popout, view, panelHere, panelThere } = await mount();
    cleanups.push(() => fixture.destroy(), popout.restore);
    expect(panelThere()).not.toBeNull();
    let backWhenClosed: boolean | null = null;
    popout.closes.onClose = () => (backWhenClosed = panelHere() !== null);
    view.open.set(false);
    await settle(fixture);
    expect(popout.closes.count).toBe(1);
    expect(backWhenClosed).toBe(true);
  });

  it("follows the root's colorMode in the pop-out", async () => {
    const { fixture, popout, panelThere } = await mount((view) => view.colorMode.set('dark'));
    cleanups.push(() => fixture.destroy(), popout.restore);
    expect(panelThere()?.closest('.docx-editor')?.classList.contains('dark')).toBe(true);
  });
});

describe('an editor in a pop-out window', () => {
  beforeEach(() => configureZoneless());

  it('resolves its host document from the pop-out, and keeps undo across the move', async () => {
    const { fixture, popout, view } = await mount((host) => {
      host.open.set(false);
      host.withPages.set(true);
    });
    try {
      const editor = engineOf(view.root().context.editor());
      expect(view.probe().hostDocument()).toBe(document);
      typeText(editor, 'before the move');
      await settle(fixture);
      expect(editor.snapshot().canUndo).toBe(true);

      view.open.set(true);
      await settle(fixture);
      const pages = popout.view.document.querySelector('docx-editor-content .docx-page');
      expect(pages).not.toBeNull();
      expect(view.probe().hostDocument()).toBe(popout.view.document);

      // The edit from before the move is still on the undo stack, and the editor still types.
      expect(editor.snapshot().canUndo).toBe(true);
      typeText(editor, 'after ');
      await settle(fixture);
      expect(editor.exec({ type: 'undo' }).ok).toBe(true);

      view.open.set(false);
      await settle(fixture);
      expect(view.probe().hostDocument()).toBe(document);
    } finally {
      fixture.destroy();
      popout.restore();
    }
  });
});
