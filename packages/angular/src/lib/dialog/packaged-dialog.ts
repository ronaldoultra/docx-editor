import { Directive, booleanAttribute, input, output } from '@angular/core';
import { injectTranslation } from '../i18n';

/**
 * The inputs and the output that every packaged open-and-close dialog has, such as the page
 * setup, paragraph, and watermark dialogs. The host owns the visibility: it shows the dialog
 * with `open` or with an `@if` block, and hides it on `closed`.
 *
 * @public
 */
@Directive()
export abstract class DocxEditorDialogBase {
  /** Show the dialog. Defaults to `true`, so an `@if` block around the dialog is enough. */
  readonly open = input(true, { transform: booleanAttribute });
  /**
   * Called on Cancel, Escape, a press outside the dialog, and after an accepted Apply, beside
   * the `closed` output. The editor's popup registry gives this input to its dialogs.
   */
  readonly closeHandler = input<(() => void) | undefined>(undefined);
  /** Classes for the dialog element. */
  readonly className = input<string | undefined>(undefined);
  /** Inline styles for the dialog element. */
  readonly styles = input<Readonly<Record<string, string>> | undefined>(undefined);
  /**
   * With `true` (the default), the dialog renders its packaged parts, the part templates in
   * their place, and the other content after them. With `false`, it renders only its content.
   */
  readonly preset = input(true, { transform: booleanAttribute });

  /** Emits when the dialog asks to close. */
  readonly closed = output<void>();

  private readonly translation = injectTranslation();

  /** Ask the host to close the dialog. */
  protected close(): void {
    this.closed.emit();
    this.closeHandler()?.();
  }

  /** @internal */
  protected label(key: string, params?: Record<string, string | number>): string {
    const t = this.translation.t();
    return t(key as Parameters<typeof t>[0], params);
  }
}
