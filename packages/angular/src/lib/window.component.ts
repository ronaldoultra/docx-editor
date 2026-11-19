import {
  ChangeDetectionStrategy,
  Component,
  DestroyRef,
  NgZone,
  TemplateRef,
  ViewContainerRef,
  contentChild,
  effect,
  inject,
  input,
  output,
  untracked,
  type EmbeddedViewRef,
} from '@angular/core';
import { openEditorWindow, type EditorWindow } from '@docx-editor.dev/core/editor';
import { injectEditorColorMode } from './color-mode';
import { injectTranslation } from './i18n';

/** Copies the stylesheets of one document into another. Returns the function that stops it. @public */
export type DocxEditorWindowStyles = boolean | ((from: Document, to: Document) => () => void);

/**
 * Show the editor panel in its own browser window while `open` is `true`, and in place
 * otherwise.
 *
 * Put it inside the element with `docxEditorRoot`, around the parts to move, and give the parts
 * as an `ng-template`. The root stays in this window, so the editor moves to the pop-out and back
 * with its document, undo history, selection, and collaboration session. The component creates
 * the panel again on each move, which is what moves the editor. When the user closes the window,
 * the panel comes back, and the component emits `openChange` with `false` and `close`.
 *
 * @example
 * ```html
 * <ng-container docxEditorRoot [document]="bytes()">
 *   <button (click)="popped.set(!popped())">Pop out</button>
 *   <docx-editor-window [(open)]="popped">
 *     <ng-template>
 *       <docx-editor-viewport><docx-editor-content /></docx-editor-viewport>
 *     </ng-template>
 *   </docx-editor-window>
 * </ng-container>
 * ```
 *
 * @public
 */
@Component({
  selector: 'docx-editor-window',
  changeDetection: ChangeDetectionStrategy.OnPush,
  template: '',
})
export class DocxEditorWindowComponent {
  /** Show the panel in a pop-out window. While `false`, the panel shows in place. */
  readonly open = input.required<boolean>();
  /** A same-origin page to open. Defaults to `about:blank`. */
  readonly url = input<string | undefined>(undefined);
  /**
   * The window name. Defaults to `docx-editor`. Two windows with the same name share one
   * browser window, so give each editor its own name.
   */
  readonly name = input<string | undefined>(undefined);
  /** The `window.open` features. Defaults to `popup,width=1000,height=900`. */
  readonly features = input<string | undefined>(undefined);
  /** The title of the pop-out window. */
  readonly title = input<string | undefined>(undefined);
  /** Copy this page's stylesheets into the window, and keep them in step. Defaults to `true`. */
  readonly styles = input<DocxEditorWindowStyles | undefined>(undefined);

  /** Emits `false` when the pop-out went away. Use it with `[(open)]`. */
  readonly openChange = output<boolean>();
  /** Emits when the pop-out went away: the user closed it, or the browser blocked it. */
  readonly close = output<void>();

  private readonly panel = contentChild.required(TemplateRef);
  private readonly container = inject(ViewContainerRef);
  private readonly zone = inject(NgZone);
  private readonly colorMode = injectEditorColorMode().colorMode;
  private readonly dir = injectTranslation().dir;
  private popup: EditorWindow | null = null;
  private wrapper: HTMLElement | null = null;
  private view: EmbeddedViewRef<unknown> | null = null;
  private generation = 0;

  constructor() {
    effect(() => {
      const open = this.open();
      this.panel();
      untracked(() => (open ? this.openWindow() : this.closeWindow()));
    });
    // The window's document has no host direction or theme to inherit, so set them on the
    // wrapper of the panel.
    effect(() => {
      const dark = this.colorMode() === 'dark';
      const dir = this.dir();
      if (!this.wrapper) return;
      this.wrapper.className = dark ? 'docx-editor dark' : 'docx-editor';
      this.wrapper.setAttribute('dir', dir);
    });
    inject(DestroyRef).onDestroy(() => {
      this.generation += 1;
      const current = this.popup;
      this.popup = null;
      this.view?.destroy();
      this.view = null;
      current?.close();
    });
  }

  private openWindow(): void {
    if (this.popup) return;
    if (!this.view) this.render(null);
    const mine = ++this.generation;
    void this.zone
      .runOutsideAngular(() =>
        openEditorWindow({
          ...(this.url() !== undefined ? { url: this.url() } : {}),
          ...(this.name() !== undefined ? { name: this.name() } : {}),
          ...(this.features() !== undefined ? { features: this.features() } : {}),
          ...(this.title() !== undefined ? { title: this.title() } : {}),
          ...(this.styles() !== undefined ? { styles: this.styles() } : {}),
        })
      )
      .then((opened) =>
        this.zone.run(() => {
          if (mine !== this.generation) {
            opened?.close();
            return;
          }
          if (!opened) {
            this.openChange.emit(false);
            this.close.emit();
            return;
          }
          opened.onClose(() =>
            this.zone.run(() => {
              if (mine !== this.generation || this.popup !== opened) return;
              // Inside the window's `pagehide`: the panel comes back before its document goes.
              this.popup = null;
              this.render(null);
              this.openChange.emit(false);
              this.close.emit();
            })
          );
          this.popup = opened;
          this.render(opened);
        })
      );
  }

  private closeWindow(): void {
    this.generation += 1;
    const current = this.popup;
    this.popup = null;
    if (!current) {
      if (!this.view && untracked(this.open) === false) this.render(null);
      return;
    }
    // The panel moves home first, then the window closes.
    this.render(null);
    current.close();
  }

  /** Create the panel again, in the pop-out's element or in place. */
  private render(popup: EditorWindow | null): void {
    this.view?.destroy();
    this.view = null;
    this.wrapper = null;
    const view = this.container.createEmbeddedView(this.panel());
    this.view = view;
    if (!popup) return;
    const wrapper = popup.element.ownerDocument.createElement('div');
    wrapper.className = this.colorMode() === 'dark' ? 'docx-editor dark' : 'docx-editor';
    wrapper.style.height = '100%';
    wrapper.setAttribute('dir', this.dir());
    for (const node of view.rootNodes as Node[]) wrapper.append(node);
    popup.element.append(wrapper);
    this.wrapper = wrapper;
  }
}
