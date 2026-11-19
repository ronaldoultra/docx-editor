import {
  ChangeDetectionStrategy,
  Component,
  computed,
  effect,
  signal,
  type Signal,
} from '@angular/core';
import type { EditorSnapshot } from '@docx-editor.dev/core/contracts/editor';
import { editorHostDocument } from '@docx-editor.dev/core/editor';
import { injectDialog } from '../dialog/dialog-parts';
import { injectDocxEditor } from '../context';
import { injectEditorState } from '../editor-state';
import {
  WATERMARK_PREVIEW_FONT_SIZE,
  WATERMARK_WASHOUT_FILTER,
  watermarkPreviewInk,
  watermarkPreviewLines,
  watermarkPreviewPage,
  watermarkPreviewPicture,
  watermarkPreviewText,
  watermarkPreviewTextTransform,
  type WatermarkDialogFields,
  type WatermarkPreviewPage,
  type WatermarkPreviewRect,
  type WatermarkPreviewText,
} from './watermark-dialog-model';

const selectPageSetup = (snapshot: EditorSnapshot) => snapshot.pageSetup ?? null;

/**
 * The page thumbnail of the enclosing watermark dialog, with the draft drawn over it as the
 * page draws it. It is decorative: the fields say the same. The packaged dialog renders it in
 * its `preview` part. Put it in your own layout inside the dialog.
 *
 * @public
 */
@Component({
  selector: 'docx-editor-watermark-preview',
  changeDetection: ChangeDetectionStrategy.OnPush,
  host: { style: 'display: contents' },
  template: `
    <svg
      class="docx-dialog__preview-page"
      [attr.viewBox]="'0 0 ' + page().width + ' ' + page().height"
      focusable="false"
      aria-hidden="true"
    >
      <rect
        class="docx-dialog__preview-sheet"
        x="0"
        y="0"
        [attr.width]="page().width"
        [attr.height]="page().height"
      />
      <!-- The watermark paints behind the text lines, as it does on the page. -->
      @if (text(); as text) {
        <text
          data-docx-watermark-preview="text"
          x="0"
          y="0"
          [attr.font-family]="text.fontFamily"
          [attr.font-size]="fontSize"
          [attr.font-weight]="text.fontWeight"
          [attr.font-style]="text.fontStyle"
          [attr.fill]="text.fill"
          [attr.fill-opacity]="text.opacity"
          [attr.transform]="transform(text)"
          [textContent]="text.text"
        ></text>
      }
      @if (picture(); as picture) {
        <image
          data-docx-watermark-preview="picture"
          class="docx-dialog__preview-picture"
          preserveAspectRatio="none"
          [attr.href]="pictureUrl()"
          [attr.x]="picture.x"
          [attr.y]="picture.y"
          [attr.width]="picture.width"
          [attr.height]="picture.height"
          [style.filter]="values().washout ? washout : null"
        />
      }
      @for (line of lines(); track $index) {
        <rect
          class="docx-dialog__preview-line"
          [attr.x]="line.x"
          [attr.y]="line.y"
          [attr.width]="line.width"
          [attr.height]="line.height"
        />
      }
    </svg>
  `,
})
export class DocxEditorWatermarkPreviewComponent {
  /** @internal */
  protected readonly fontSize: number = WATERMARK_PREVIEW_FONT_SIZE;
  /** @internal */
  protected readonly washout: string = WATERMARK_WASHOUT_FILTER;
  /** @internal */
  protected readonly values: Signal<WatermarkDialogFields> =
    injectDialog<WatermarkDialogFields>().values;
  /** @internal */
  protected readonly page: Signal<WatermarkPreviewPage> = computed(() =>
    watermarkPreviewPage(this.pageSetup())
  );
  /** @internal */
  protected readonly text: Signal<WatermarkPreviewText | null> = computed(() =>
    watermarkPreviewText(this.values(), this.page())
  );
  /** @internal */
  protected readonly pictureUrl = signal<string | null>(null);
  /** @internal */
  protected readonly picture: Signal<WatermarkPreviewRect | null> = computed(() =>
    this.pictureUrl() ? watermarkPreviewPicture(this.values(), this.page()) : null
  );
  /** @internal */
  protected readonly lines: Signal<readonly WatermarkPreviewRect[]> = computed(() =>
    watermarkPreviewLines(this.page())
  );

  private readonly editor = injectDocxEditor();
  private readonly pageSetup = injectEditorState(selectPageSetup);

  constructor() {
    const bytes = computed(() => {
      const values = this.values();
      return values.kind === 'picture' ? values.imageData : null;
    });
    // An object URL for the picture bytes, revoked when the bytes change and on destroy.
    effect((onCleanup) => {
      const data = bytes();
      if (!data || typeof URL.createObjectURL !== 'function') {
        this.pictureUrl.set(null);
        return;
      }
      const url = URL.createObjectURL(new Blob([data as BlobPart]));
      this.pictureUrl.set(url);
      onCleanup(() => URL.revokeObjectURL(url));
    });
  }

  /** @internal */
  protected transform(text: WatermarkPreviewText): string {
    const editor = this.editor();
    return watermarkPreviewTextTransform(
      text,
      watermarkPreviewInk(text, editor ? editorHostDocument(editor) : null)
    );
  }
}
