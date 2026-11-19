// The claims of packages/vue/test/root-color-mode.test.ts and scope-dispose.test.ts. The
// toolbar and menu bar take the same scope classes when they port (phases 4 and 5).
import { ChangeDetectionStrategy, Component, signal, viewChild } from '@angular/core';
import { blankDocumentBytes } from '@docx-editor.dev/core/editor';
import { configureZoneless, render, settle } from '../testing/harness';
import {
  injectEditorColorMode,
  injectScopeClassName,
  type DocxEditorColorMode,
} from './color-mode';
import { DocxEditorContentComponent } from './content.component';
import { DocxEditorRootDirective } from './root.directive';
import { DocxEditorViewportComponent } from './viewport.component';

type Listener = (event: { matches: boolean }) => void;

/** A controllable `(prefers-color-scheme: dark)` query. */
const media = {
  matches: false,
  listeners: new Set<Listener>(),
  set(matches: boolean) {
    media.matches = matches;
    for (const listener of media.listeners) listener({ matches });
  },
};

const realMatchMedia = window.matchMedia;

@Component({
  selector: 'theme-probe',
  template: `<output data-testid="theme">{{ colorMode() }}</output>
    <div data-testid="part" [class]="scope() + 'my-part'"></div>`,
})
class ProbeComponent {
  readonly colorMode = injectEditorColorMode().colorMode;
  readonly scope = injectScopeClassName();
}

@Component({
  imports: [
    DocxEditorRootDirective,
    DocxEditorViewportComponent,
    DocxEditorContentComponent,
    ProbeComponent,
  ],
  changeDetection: ChangeDetectionStrategy.OnPush,
  template: `
    <ng-container docxEditorRoot [document]="bytes" [colorMode]="colorMode()">
      <theme-probe />
      <docx-editor-viewport><docx-editor-content /></docx-editor-viewport>
    </ng-container>
  `,
})
class HostComponent {
  readonly bytes = blankDocumentBytes();
  readonly colorMode = signal<DocxEditorColorMode | undefined>(undefined);
  readonly probe = viewChild.required(ProbeComponent);
}

async function mount(initial?: DocxEditorColorMode) {
  const fixture = await render(HostComponent);
  const view = fixture.componentInstance;
  if (initial) {
    view.colorMode.set(initial);
    await settle(fixture);
  }
  const host = fixture.nativeElement as HTMLElement;
  return {
    fixture,
    view,
    viewport: () => host.querySelector<HTMLElement>('docx-editor-viewport')!,
    theme: () => host.querySelector('[data-testid="theme"]')!.textContent,
    probe: () => view.probe(),
    part: () => host.querySelector<HTMLElement>('[data-testid="part"]')!,
  };
}

describe('docxEditorRoot colorMode', () => {
  beforeEach(() => {
    configureZoneless();
    media.matches = false;
    media.listeners.clear();
    (window as unknown as { matchMedia: unknown }).matchMedia = () => ({
      get matches() {
        return media.matches;
      },
      addEventListener: (_type: string, listener: Listener) => media.listeners.add(listener),
      removeEventListener: (_type: string, listener: Listener) => media.listeners.delete(listener),
    });
  });

  afterEach(() => {
    (window as unknown as { matchMedia: unknown }).matchMedia = realMatchMedia;
  });

  it("puts .dark beside .docx-editor on the viewport for 'dark'", async () => {
    const view = await mount('dark');
    expect(view.viewport().classList.contains('docx-editor')).toBe(true);
    expect(view.viewport().classList.contains('dark')).toBe(true);
    view.fixture.destroy();
  });

  it("keeps the light chrome when colorMode is omitted or 'light'", async () => {
    const view = await mount();
    expect(view.viewport().classList.contains('dark')).toBe(false);
    view.view.colorMode.set('light');
    await settle(view.fixture);
    expect(view.viewport().classList.contains('dark')).toBe(false);
    view.fixture.destroy();
  });

  it('applies a changed colorMode without new elements', async () => {
    const view = await mount('light');
    const viewport = view.viewport();
    view.view.colorMode.set('dark');
    await settle(view.fixture);
    expect(view.viewport()).toBe(viewport);
    expect(viewport.classList.contains('dark')).toBe(true);
    view.fixture.destroy();
  });

  it("follows the system setting for 'system', and stops when it leaves 'system'", async () => {
    media.matches = true;
    const view = await mount('system');
    expect(view.viewport().classList.contains('dark')).toBe(true);
    media.set(false);
    await settle(view.fixture);
    expect(view.viewport().classList.contains('dark')).toBe(false);
    media.set(true);
    await settle(view.fixture);
    expect(view.viewport().classList.contains('dark')).toBe(true);
    expect(media.listeners.size).toBeGreaterThan(0);
    view.view.colorMode.set('light');
    await settle(view.fixture);
    expect(media.listeners.size).toBe(0);
    view.fixture.destroy();
  });

  it('reports the resolved theme to parts that the host builds', async () => {
    const view = await mount();
    expect(view.theme()).toBe('light');
    view.view.colorMode.set('dark');
    await settle(view.fixture);
    expect(view.theme()).toBe('dark');
    view.view.colorMode.set('system');
    await settle(view.fixture);
    expect(view.theme()).toBe('light');
    media.set(true);
    await settle(view.fixture);
    expect(view.theme()).toBe('dark');
    view.fixture.destroy();
  });

  it('adds dark to the scope classes under a dark root', async () => {
    const view = await mount('dark');
    expect(view.probe().scope()).toBe('docx-editor dark ');
    expect([...view.part().classList].sort()).toEqual(['dark', 'docx-editor', 'my-part']);
    view.fixture.destroy();
  });
});

describe('mounting the composed editor', () => {
  beforeEach(() => configureZoneless());

  it('logs no warning and no error', async () => {
    const warn = vi.spyOn(console, 'warn');
    const error = vi.spyOn(console, 'error');
    try {
      const view = await mount();
      view.fixture.destroy();
      expect(warn).not.toHaveBeenCalled();
      expect(error).not.toHaveBeenCalled();
    } finally {
      warn.mockRestore();
      error.mockRestore();
    }
  });
});
