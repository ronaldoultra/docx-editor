import {
  Injectable,
  Injector,
  afterNextRender,
  inject,
  signal,
  untracked,
  type Signal,
} from '@angular/core';
import {
  editorHostDocument,
  hostDocumentOf,
  isHTMLElementNode,
} from '@docx-editor.dev/core/editor';
import { DocxEditorContext } from '../context';
import { DocxEditorPopupRegistry } from '../popups/popup-config';

/**
 * The dialogs that an editor root opens from its chrome.
 *
 * @public
 */
export type DocxEditorDialogKind = 'pageSetup' | 'paragraph' | 'watermark';

/**
 * The dialogs of one editor root. The menu bar, the toolbar, and host code open a dialog with
 * `open()`. The content component renders the open dialog: the popup registry entry, or the
 * packaged dialog. One dialog is open at a time.
 *
 * @internal
 */
@Injectable()
export class DocxDialogHost {
  private readonly registry = inject(DocxEditorPopupRegistry);
  private readonly context = inject(DocxEditorContext);
  private readonly injector = inject(Injector);
  private readonly activeSignal = signal<DocxEditorDialogKind | null>(null);
  private opener: HTMLElement | null = null;

  /** The open dialog, or `null`. */
  readonly active: Signal<DocxEditorDialogKind | null> = this.activeSignal.asReadonly();

  /**
   * Open the dialog `kind`. The focus goes back to `trigger` when the dialog closes, or to the
   * element that has the focus now. A `false` registry entry opens nothing.
   */
  open(kind: DocxEditorDialogKind, trigger?: HTMLElement | null): void {
    this.close();
    if (untracked(this.registry.popups)[kind] === false) return;
    const editor = untracked(this.context.editor);
    const doc = trigger ? hostDocumentOf(trigger) : editor ? editorHostDocument(editor) : null;
    const focused = doc?.activeElement;
    this.opener =
      trigger ??
      (isHTMLElementNode(focused) && focused !== focused.ownerDocument.body ? focused : null);
    this.activeSignal.set(kind);
  }

  /** Close the open dialog. The focus goes back to its opener after the render. */
  close(): void {
    if (untracked(this.activeSignal) === null) return;
    this.activeSignal.set(null);
    const opener = this.opener;
    this.opener = null;
    if (!opener) return;
    afterNextRender(
      () => {
        if (untracked(this.activeSignal) === null && opener.isConnected) {
          opener.focus({ preventScroll: true });
        }
      },
      { injector: this.injector }
    );
  }
}

/** The dialogs of the enclosing root, or `null` outside a root. @internal */
export function injectDialogHost(): DocxDialogHost | null {
  return inject(DocxDialogHost, { optional: true });
}
