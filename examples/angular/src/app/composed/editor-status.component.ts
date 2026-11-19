import { DOCUMENT } from '@angular/common';
import { ChangeDetectionStrategy, Component, computed, inject, signal } from '@angular/core';
import {
  injectDocxEditor,
  injectEditorColorMode,
  injectEditorEvent,
  injectEditorHostDocument,
  injectEditorState,
  injectChromeDir,
  injectScopeClassName,
  injectTranslation,
} from '@docx-editor.dev/angular';

const ZOOM_LEVELS = [0.5, 0.75, 1, 1.25, 1.5, 2];

/** A status bar: page number, zoom, and values that other `inject…` functions give. */
@Component({
  selector: 'app-editor-status',
  changeDetection: ChangeDetectionStrategy.OnPush,
  host: { '[class]': "scope() + 'editor-status'", '[attr.dir]': 'chromeDir()' },
  template: `
    <span data-testid="composed-page">Página {{ page().current }} de {{ page().total }}</span>
    <span class="zoom">
      <button
        type="button"
        [attr.aria-label]="t()('zoom.zoomOut')"
        [disabled]="!editor() || zoom() <= minZoom"
        (click)="stepZoom(-1)"
      >
        −
      </button>
      <output [attr.aria-label]="t()('zoom.zoomLevel')">{{ zoomPercent() }}%</output>
      <button
        type="button"
        [attr.aria-label]="t()('zoom.zoomIn')"
        [disabled]="!editor() || zoom() >= maxZoom"
        (click)="stepZoom(1)"
      >
        +
      </button>
    </span>
    <span>Seleções: {{ selections() }}</span>
    <span>Tema: {{ colorMode() === 'dark' ? 'escuro' : 'claro' }}</span>
    <span>Páginas em: {{ inPopout() ? 'outra janela' : 'esta janela' }}</span>
  `,
})
export class EditorStatusComponent {
  protected readonly scope = injectScopeClassName();
  protected readonly chromeDir = injectChromeDir();
  protected readonly editor = injectDocxEditor();
  protected readonly colorMode = injectEditorColorMode().colorMode;
  private readonly translation = injectTranslation();
  private readonly hostDocument = injectEditorHostDocument();
  private readonly appDocument = inject(DOCUMENT);

  protected readonly minZoom = ZOOM_LEVELS[0]!;
  protected readonly maxZoom = ZOOM_LEVELS[ZOOM_LEVELS.length - 1]!;

  // This signal notifies only when the page numbers change, not on each edit.
  protected readonly page = injectEditorState(
    (snapshot) => snapshot.page,
    (a, b) => a.current === b.current && a.total === b.total
  );
  protected readonly zoom = injectEditorState((snapshot) => snapshot.zoom);
  protected readonly zoomPercent = computed(() => Math.round(this.zoom() * 100));
  protected readonly selections = signal(0);
  protected readonly inPopout = computed(() => {
    const current = this.hostDocument();
    return current !== null && current !== this.appDocument;
  });
  protected readonly t = computed(() => this.translation.t() as (key: string) => string);

  constructor() {
    injectEditorEvent('selectionChange', () => this.selections.update((count) => count + 1));
  }

  protected stepZoom(direction: 1 | -1) {
    const editor = this.editor();
    if (!editor) return;
    const zoom = this.zoom();
    const next =
      direction > 0
        ? ZOOM_LEVELS.find((level) => level > zoom + 0.001)
        : [...ZOOM_LEVELS].reverse().find((level) => level < zoom - 0.001);
    if (next !== undefined) editor.setZoom(next);
  }
}
