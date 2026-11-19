import {
  ChangeDetectionStrategy,
  Component,
  DestroyRef,
  ElementRef,
  NgZone,
  ViewEncapsulation,
  computed,
  inject,
} from '@angular/core';
import { isElementNode } from '@docx-editor.dev/core/editor';
import { injectOwnScopeClassName } from './color-mode';
import { DOCX_EDITOR_SCOPED, injectDocxEditor } from './context';
import { injectEditorState } from './editor-state';
import { injectTranslation } from './i18n';
import { zoomLevelForShortcut } from './zoom-levels';

/**
 * The scroll container of the pages. Put `<docx-editor-content>` inside it.
 *
 * The element carries the `docx-editor__scroll-container` class that the engine stylesheet and
 * geometry expect, and it is a named `region`. With Ctrl or Cmd, the `+`, `-`, and `0` keys
 * zoom the pages.
 *
 * @public
 */
@Component({
  selector: 'docx-editor-viewport',
  changeDetection: ChangeDetectionStrategy.OnPush,
  encapsulation: ViewEncapsulation.None,
  providers: [{ provide: DOCX_EDITOR_SCOPED, useValue: true }],
  host: {
    class:
      'docx-editor-one-surface docx-editor-one-surface__viewport docx-editor__scroll-container',
    '[class.docx-editor]': 'scoped()',
    '[class.dark]': 'dark()',
    'data-testid': 'docx-editor-scroll',
    tabindex: '0',
    role: 'region',
    '[attr.aria-label]': 'label()',
    '[attr.data-zoom-fit]': "fitting() ? '' : null",
  },
  template: '<ng-content />',
})
export class DocxEditorViewportComponent {
  private readonly editor = injectDocxEditor();
  private readonly scope = injectOwnScopeClassName();
  private readonly translation = injectTranslation();

  /** @internal */
  protected readonly scoped = computed(() => this.scope() !== '');
  /** @internal */
  protected readonly dark = computed(() => this.scope().includes('dark'));
  /** @internal */
  protected readonly label = computed(() => this.translation.t()('editor.documentViewport'));
  /** @internal */
  protected readonly fitting = injectEditorState((snapshot) => snapshot.zoomMode?.type === 'fit');

  constructor() {
    const element: HTMLElement = inject(ElementRef).nativeElement;
    const onKeyDown = (event: KeyboardEvent) => {
      const editor = this.editor();
      if (!editor || !(event.metaKey || event.ctrlKey) || event.altKey) return;
      const target = event.target;
      if (isElementNode(target) && target.closest('input, textarea, select, [role="dialog"]')) {
        return;
      }
      const next = zoomLevelForShortcut(event.key, editor.getZoom());
      if (next === null) return;
      event.preventDefault();
      editor.setZoom(next);
    };
    // A capture listener, so the shortcut wins over the page's own key handling.
    inject(NgZone).runOutsideAngular(() => element.addEventListener('keydown', onKeyDown, true));
    inject(DestroyRef).onDestroy(() => element.removeEventListener('keydown', onKeyDown, true));
  }
}
