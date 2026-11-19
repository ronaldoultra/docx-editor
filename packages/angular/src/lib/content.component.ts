import { isPlatformBrowser } from '@angular/common';
import {
  ChangeDetectionStrategy,
  Component,
  DestroyRef,
  ElementRef,
  NgZone,
  PLATFORM_ID,
  ViewEncapsulation,
  effect,
  inject,
  untracked,
  viewChild,
  type Signal,
} from '@angular/core';
import {
  editorDropHandled,
  openChartPaneAt,
  type DocxEditorInstance,
} from '@docx-editor.dev/core/editor';
import { injectDocxEditorContext } from './context';
import { injectChromeDir, type DocxEditorDirection } from './i18n';
import { DocxConfiguredPopupsComponent } from './popups/configured-popups.component';

/**
 * The element that the engine paints the pages into. The painted pages are the editable
 * surface. Put it inside `<docx-editor-viewport>`, under an element with `docxEditorRoot`.
 *
 * When Angular destroys the content component and the root stays, the engine detaches. It keeps
 * its document, undo history, and selection, and a new content component attaches it again.
 *
 * @public
 */
@Component({
  selector: 'docx-editor-content',
  imports: [DocxConfiguredPopupsComponent],
  changeDetection: ChangeDetectionStrategy.OnPush,
  encapsulation: ViewEncapsulation.None,
  host: { class: 'docx-content-mount' },
  template: `
    <div
      #surface
      class="docx-paginated-surface"
      (drop)="onDrop($event)"
      (dblclick)="onDoubleClick($event)"
    ></div>
    <div class="docx-dialog-mount" [attr.dir]="chromeDir()"></div>
    <docx-configured-popups />
  `,
})
export class DocxEditorContentComponent {
  private readonly context = injectDocxEditorContext();
  private readonly zone = inject(NgZone);
  private readonly isBrowser = isPlatformBrowser(inject(PLATFORM_ID));
  private readonly surface = viewChild.required<ElementRef<HTMLDivElement>>('surface');
  private attached: DocxEditorInstance | null = null;

  /** @internal */
  protected readonly chromeDir: Signal<DocxEditorDirection | undefined> = injectChromeDir();

  constructor() {
    effect(() => {
      const editor = this.context?.editor() ?? null;
      const element = this.surface().nativeElement;
      if (!this.isBrowser) return;
      untracked(() => this.zone.runOutsideAngular(() => this.attach(editor, element)));
    });
    inject(DestroyRef).onDestroy(() => this.zone.runOutsideAngular(() => this.detach()));
  }

  /** @internal */
  protected onDrop(event: DragEvent): void {
    // A file dropped on the pages must not make the browser open it in place of the app.
    event.preventDefault();
    editorDropHandled(event);
  }

  /** @internal */
  protected onDoubleClick(event: MouseEvent): void {
    // A double press on a chart opens the chart pane, as the Edit chart command does.
    const editor = this.context?.editor();
    if (editor && openChartPaneAt(editor, event)) event.preventDefault();
  }

  private attach(editor: DocxEditorInstance | null, element: HTMLDivElement): void {
    if (editor === this.attached) return;
    this.detach();
    if (!editor) return;
    this.attached = editor;
    try {
      editor.attach(element);
    } catch (error) {
      // A document that fails to open is a refusal on the snapshot, not a crash of the host.
      // Other errors are defects, so they continue to the caller.
      if (editor.snapshot().openFailure?.code !== 'open.failed') throw error;
      console.error('[docx-editor] The document failed to open.', error);
    }
    this.context?.setContentElement(element);
  }

  private detach(): void {
    const editor = this.attached;
    if (!editor) return;
    this.attached = null;
    // A destroyed engine has nothing to detach.
    if (untracked(() => this.context?.editor()) === editor) editor.detach();
    if (untracked(() => this.context?.contentElement()) === this.surface().nativeElement) {
      this.context?.setContentElement(null);
    }
  }
}
