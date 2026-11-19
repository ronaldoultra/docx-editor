import {
  ChangeDetectionStrategy,
  Component,
  computed,
  effect,
  untracked,
  type Signal,
} from '@angular/core';
import { injectDocxEditor } from '../context';
import { injectEditorState } from '../editor-state';
import { DocxEditorPageSetupDialogComponent } from '../page-setup/page-setup-dialog.component';
import { DocxEditorParagraphDialogComponent } from '../paragraph/paragraph-dialog.component';
import {
  injectDocxEditorPopups,
  type DocxEditorPopup,
  type DocxEditorPopups,
} from '../popups/popup-config';
import { DocxPopupOutletComponent } from '../popups/popup-outlet.component';
import { DocxEditorWatermarkDialogComponent } from '../watermark/watermark-dialog.component';
import { injectDialogHost, type DocxEditorDialogKind } from './dialog-host';

/**
 * Render the dialog that the root's dialog host has open: the popup registry entry, or the
 * packaged dialog when the registry has no entry. The dialog closes when the engine mounts
 * another document, and when its entry turns to `false`.
 *
 * @internal
 */
@Component({
  selector: 'docx-configured-dialogs',
  imports: [
    DocxPopupOutletComponent,
    DocxEditorPageSetupDialogComponent,
    DocxEditorParagraphDialogComponent,
    DocxEditorWatermarkDialogComponent,
  ],
  changeDetection: ChangeDetectionStrategy.OnPush,
  host: { style: 'display: contents' },
  template: `
    @switch (active()) {
      @case ('pageSetup') {
        @if (entry(); as popup) {
          <docx-popup-outlet [entry]="popup" [inputs]="inputs" />
        } @else {
          <docx-editor-page-setup-dialog (closed)="close()" />
        }
      }
      @case ('paragraph') {
        @if (entry(); as popup) {
          <docx-popup-outlet [entry]="popup" [inputs]="inputs" />
        } @else {
          <docx-editor-paragraph-dialog (closed)="close()" />
        }
      }
      @case ('watermark') {
        @if (entry(); as popup) {
          <docx-popup-outlet [entry]="popup" [inputs]="inputs" />
        } @else {
          <docx-editor-watermark-dialog (closed)="close()" />
        }
      }
    }
  `,
})
export class DocxConfiguredDialogsComponent {
  private readonly host = injectDialogHost();
  private readonly popups: Signal<DocxEditorPopups> = injectDocxEditorPopups();

  /** @internal */
  protected readonly active: Signal<DocxEditorDialogKind | null> = computed(
    () => this.host?.active() ?? null
  );
  /** The registry entry of the open dialog, or `undefined` for the packaged dialog. @internal */
  protected readonly entry: Signal<DocxEditorPopup<object> | undefined> = computed(() => {
    const kind = this.active();
    const entry = kind ? this.popups()[kind] : undefined;
    return entry === false ? undefined : entry;
  });
  /** @internal */
  protected readonly inputs: Readonly<Record<string, unknown>> = {
    closeHandler: () => this.close(),
  };

  constructor() {
    const editor = injectDocxEditor();
    const generation = injectEditorState(() => editor()?.mountGeneration ?? 0);
    let previous: number | null = null;
    effect(() => {
      const current = generation();
      untracked(() => {
        if (previous !== null && previous !== current) this.close();
        previous = current;
      });
    });
    effect(() => {
      const kind = this.active();
      if (kind && this.popups()[kind] === false) untracked(() => this.close());
    });
  }

  /** @internal */
  protected close(): void {
    this.host?.close();
  }
}
