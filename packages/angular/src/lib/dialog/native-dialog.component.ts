import {
  ChangeDetectionStrategy,
  Component,
  DestroyRef,
  ElementRef,
  afterNextRender,
  booleanAttribute,
  computed,
  effect,
  inject,
  input,
  output,
  signal,
  viewChild,
  type Signal,
} from '@angular/core';
import { hostDocumentOf, isHTMLElementNode, trapTabWithin } from '@docx-editor.dev/core/editor';
import { injectDocxEditorContext } from '../context';

/**
 * A modal `<dialog>` for the packaged dialogs. It opens with `showModal()`, moves the focus to
 * its first control, keeps Tab inside, and closes on Escape, on a press outside its box, and when
 * `sessionSignal` aborts. When Angular destroys it, the focus goes back to the element that had
 * it before, if that element is not the page body.
 *
 * Outside an element with the `docx-editor` scope class, it adds that class itself, and the
 * `dark` class when the editor's color mode is dark.
 *
 * @public
 */
@Component({
  selector: 'docx-native-dialog',
  changeDetection: ChangeDetectionStrategy.OnPush,
  host: { style: 'display: contents' },
  template: `
    <dialog
      #panel
      [attr.role]="role() ?? 'dialog'"
      aria-modal="true"
      [attr.aria-label]="label()"
      [attr.aria-describedby]="describedBy() ?? null"
      [attr.data-docx-busy]="busy() ? '' : null"
      [attr.data-docx-dialog]="kind()"
      [class]="classes()"
      [style]="styles() ?? null"
      (keydown)="onKeydown($event)"
      (cancel)="onCancel($event)"
      (click)="onClick($event)"
      (mousedown)="$event.stopPropagation()"
    >
      <ng-content />
    </dialog>
  `,
})
export class DocxNativeDialogComponent {
  /** The accessible name of the dialog. */
  readonly label = input.required<string>();
  /** The dialog kind, written to `data-docx-dialog`, such as `'pageSetup'`. */
  readonly kind = input.required<string>();
  /** The ARIA role. Defaults to `dialog`. Use `alertdialog` for a message. */
  readonly role = input<string | undefined>(undefined);
  /** Close the dialog on a press outside its box. Defaults to `true`. */
  readonly dismissOutside = input(true, { transform: booleanAttribute });
  /** Give the focus back to the earlier element when the dialog closes. Defaults to `true`. */
  readonly restoreFocus = input(true, { transform: booleanAttribute });
  /** The signal of the engine session. Its abort closes the dialog. */
  readonly sessionSignal = input<AbortSignal | undefined>(undefined);
  /** The id of the element that describes the dialog, such as an alert's message. */
  readonly describedBy = input<string | undefined>(undefined);
  /** True while the dialog works, such as during a compare. */
  readonly busy = input(false, { transform: booleanAttribute });
  /** Classes for the dialog element. */
  readonly className = input<string | undefined>(undefined);
  /** Inline styles for the dialog element. */
  readonly styles = input<Readonly<Record<string, string>> | undefined>(undefined);

  /** Emits when the user asks to close: Escape, the browser's cancel, or a press outside. */
  readonly dismiss = output<void>();
  /** Emits the other key presses inside the dialog, such as Enter. */
  readonly dialogKeydown = output<KeyboardEvent>();

  private readonly panel = viewChild.required<ElementRef<HTMLDialogElement>>('panel');
  private readonly context = injectDocxEditorContext();
  private readonly defaultScope = signal(false);
  private opener: HTMLElement | null = null;

  /** @internal */
  protected readonly classes: Signal<string> = computed(() => {
    const scoped = this.defaultScope();
    return [
      'docx-dialog',
      scoped ? 'docx-editor' : '',
      scoped && this.context?.dark() ? 'dark' : '',
      this.className() ?? '',
    ]
      .filter(Boolean)
      .join(' ');
  });

  constructor() {
    afterNextRender(() => this.open());
    effect((onCleanup) => {
      const session = this.sessionSignal();
      const panel = this.panel().nativeElement;
      if (!session) return;
      const abort = () => {
        if (panel.open) panel.close?.();
      };
      if (session.aborted) abort();
      session.addEventListener('abort', abort, { once: true });
      onCleanup(() => session.removeEventListener('abort', abort));
    });
    inject(DestroyRef).onDestroy(() => {
      const panel = this.panel().nativeElement;
      if (panel.open) panel.close?.();
      if (this.restoreFocus() && this.opener?.isConnected) {
        this.opener.focus({ preventScroll: true });
      }
    });
  }

  private open(): void {
    const panel = this.panel().nativeElement;
    this.defaultScope.set(!panel.parentElement?.closest('.docx-editor'));
    const active = hostDocumentOf(panel).activeElement;
    // The page itself is no opener: restoring it would take the focus from the editor.
    this.opener = isHTMLElementNode(active) && active !== active.ownerDocument.body ? active : null;
    if (this.sessionSignal()?.aborted) return;
    if (!panel.open && typeof panel.showModal === 'function') panel.showModal();
    else panel.setAttribute('open', '');
    (
      panel.querySelector<HTMLElement>(
        'input:not([disabled]),select:not([disabled]),button:not([disabled])'
      ) ??
      // A link takes the focus only when the dialog has no control, such as an alert whose one
      // action is a support link.
      panel.querySelector<HTMLElement>('a[href]') ??
      panel
    ).focus({ preventScroll: true });
  }

  /** @internal */
  protected onKeydown(event: KeyboardEvent): void {
    event.stopPropagation();
    if (event.isComposing || event.defaultPrevented) return;
    if (trapTabWithin(this.panel().nativeElement, event)) {
      event.preventDefault();
      return;
    }
    if (event.key === 'Escape') {
      event.preventDefault();
      this.dismiss.emit();
    } else this.dialogKeydown.emit(event);
  }

  /** @internal */
  protected onCancel(event: Event): void {
    event.preventDefault();
    this.dismiss.emit();
  }

  /** @internal */
  protected onClick(event: MouseEvent): void {
    event.stopPropagation();
    const panel = this.panel().nativeElement;
    if (!this.dismissOutside() || event.target !== panel) return;
    const box = panel.getBoundingClientRect();
    if (
      event.clientX < box.left ||
      event.clientX > box.right ||
      event.clientY < box.top ||
      event.clientY > box.bottom
    ) {
      this.dismiss.emit();
    }
  }
}
