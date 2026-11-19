import {
  ChangeDetectionStrategy,
  Component,
  NgZone,
  computed,
  effect,
  inject,
  signal,
  untracked,
  type Signal,
  type WritableSignal,
} from '@angular/core';
import type {
  ContentControlWidgetSession,
  InvalidTextFormFieldSession,
} from '@docx-editor.dev/core/editor';
import { injectDocxEditor } from '../context';
import {
  injectDocxEditorPopups,
  type DocxEditorPopup,
  type DocxEditorPopups,
} from './popup-config';
import { DocxPopupOutletComponent } from './popup-outlet.component';

type WidgetEntry = 'contentControlWidget' | 'contentControlCheckbox' | 'contentControlPicture';

/** Which registry entry renders a widget session: checkbox and picture presses have their own. */
function widgetEntry(kind: ContentControlWidgetSession['kind']): WidgetEntry {
  if (kind === 'checkbox') return 'contentControlCheckbox';
  if (kind === 'picture') return 'contentControlPicture';
  return 'contentControlWidget';
}

/** The widget kinds that the configured entries can render. */
function widgetKinds(popups: DocxEditorPopups): ContentControlWidgetSession['kind'][] {
  return [
    ...(popups.contentControlWidget !== undefined
      ? (['dropdown', 'comboBox', 'date', 'buildingBlockGallery'] as const)
      : []),
    ...(popups.contentControlCheckbox !== undefined ? (['checkbox'] as const) : []),
    ...(popups.contentControlPicture !== undefined ? (['picture'] as const) : []),
  ];
}

/**
 * Host the popups of the root's registry inside the content: the configured hyperlink, content
 * control, equation, and context menu entries, and the sessions that the engine opens for the
 * content-control widget and the invalid text form field entries.
 *
 * An omitted widget entry leaves those presses to the engine's own UI. A `false` entry cancels
 * the session that the engine opens.
 *
 * @internal
 */
@Component({
  selector: 'docx-configured-popups',
  imports: [DocxPopupOutletComponent],
  changeDetection: ChangeDetectionStrategy.OnPush,
  host: { style: 'display: contents' },
  template: `
    <docx-popup-outlet [entry]="popups().hyperlink" />
    <docx-popup-outlet [entry]="popups().contentControl" />
    <docx-popup-outlet [entry]="popups().equation" />
    <docx-popup-outlet [entry]="popups().contextMenu" />
    @if (widget(); as session) {
      <docx-popup-outlet [entry]="widgetPopup()" [inputs]="{ session }" [session]="session" />
    }
    @if (invalid(); as session) {
      <docx-popup-outlet
        [entry]="popups().invalidTextFormField"
        [inputs]="{ session }"
        [session]="session"
      />
    }
  `,
})
export class DocxConfiguredPopupsComponent {
  /** @internal */
  protected readonly popups: Signal<DocxEditorPopups> = injectDocxEditorPopups();
  /** @internal */
  protected readonly widget = signal<ContentControlWidgetSession | null>(null);
  /** @internal */
  protected readonly invalid = signal<InvalidTextFormFieldSession | null>(null);
  /** @internal */
  protected readonly widgetPopup: Signal<DocxEditorPopup<object> | undefined> = computed(() => {
    const session = this.widget();
    return session ? this.popups()[widgetEntry(session.kind)] : undefined;
  });

  private readonly zone = inject(NgZone);
  private readonly editor = injectDocxEditor();
  /** The kinds as a key, so a new registry object with the same entries keeps the registration. */
  private readonly kinds = computed(() => widgetKinds(this.popups()).join(' '));
  private readonly invalidConfigured = computed(
    () => this.popups().invalidTextFormField !== undefined
  );

  constructor() {
    effect((onCleanup) => {
      const editor = this.editor();
      const kinds = this.kinds();
      if (!editor || !kinds) return;
      const dispose = untracked(() =>
        editor.setContentControlWidgetChrome(
          {
            kinds: kinds.split(' ') as ContentControlWidgetSession['kind'][],
            onRequest: (session) => this.open(this.widget, session, widgetEntry(session.kind)),
          },
          { fallback: true }
        )
      );
      onCleanup(() => {
        dispose();
        this.widget.set(null);
      });
    });
    effect((onCleanup) => {
      const editor = this.editor();
      if (!editor || !this.invalidConfigured()) return;
      const dispose = untracked(() =>
        editor.setInvalidTextFormFieldChrome(
          {
            onRequest: (session) => this.open(this.invalid, session, 'invalidTextFormField'),
          },
          { fallback: true }
        )
      );
      onCleanup(() => {
        dispose();
        this.invalid.set(null);
      });
    });
  }

  /** Show `session` until it ends, or cancel it when its entry is turned off. */
  private open<S extends { readonly signal: AbortSignal; cancel(): void }>(
    state: WritableSignal<S | null>,
    session: S,
    entry: keyof DocxEditorPopups
  ): void {
    if (untracked(this.popups)[entry] === false) {
      session.cancel();
      return;
    }
    this.zone.run(() => state.set(session));
    session.signal.addEventListener(
      'abort',
      () => {
        if (untracked(state) === session) this.zone.run(() => state.set(null));
      },
      { once: true }
    );
  }
}
