import {
  ChangeDetectionStrategy,
  Component,
  ElementRef,
  NgZone,
  afterRenderEffect,
  computed,
  inject,
  signal,
  untracked,
  viewChild,
} from '@angular/core';
import {
  hasOpenNestedPopup,
  listenForOutsidePress,
  listenForPopupEscape,
  listenForPopupFocusLeave,
  toolbarPanelPlacement,
  type ToolbarPanelPlacement,
} from '@docx-editor.dev/core/editor';
import { injectDocxEditorContext } from '../context';
import { DocxChromeIconComponent } from './chrome-icon.component';
import { DOCX_TOOLBAR_PLACEMENT, DocxToolbarState, injectToolbarLabel } from './toolbar-context';
import { MORE_PATHS, guardToolbarMousedown } from './toolbar-registry';

let nextPanelId = 0;

const FOCUSABLE =
  'button:not([disabled]), input:not([disabled]), select:not([disabled]), textarea:not([disabled]), [tabindex]:not([tabindex="-1"])';

/**
 * The "⋯" trigger and its panel. The toolbar puts the collapsed controls in it as content,
 * one labeled section for each group.
 *
 * @internal
 */
@Component({
  selector: 'div[docxToolbarMore]',
  changeDetection: ChangeDetectionStrategy.OnPush,
  imports: [DocxChromeIconComponent],
  providers: [{ provide: DOCX_TOOLBAR_PLACEMENT, useValue: 'panel' }],
  host: {
    class: 'docx-toolbar__more',
    'data-toolbar-more': '',
  },
  template: `<button
      #trigger
      type="button"
      class="docx-toolbar__button docx-toolbar__more-trigger"
      data-slot="toolbar.more"
      aria-haspopup="dialog"
      [attr.aria-expanded]="open()"
      [attr.aria-controls]="open() ? panelId : null"
      [attr.aria-label]="text()"
      [attr.title]="text()"
      [attr.data-active]="open() ? '' : null"
      (mousedown)="guard($event)"
      (click)="toggle()"
      (keydown)="onTriggerKey($event)"
    >
      <svg [docxChromeIcon]="paths"></svg>
    </button>
    @if (open()) {
      <div
        #panel
        [id]="panelId"
        role="dialog"
        [attr.aria-label]="text()"
        class="docx-toolbar__more-panel"
        data-testid="toolbar-overflow-panel"
        [attr.data-anchor]="placement()?.anchor ?? null"
        [style.left.px]="placement()?.left ?? null"
        [style.right]="placement() ? 'auto' : null"
        [style.max-inline-size.px]="placement()?.maxWidth ?? null"
        (keydown)="onPanelKey($event)"
      >
        <ng-content />
      </div>
    }`,
})
export class DocxToolbarMoreComponent {
  private readonly state = inject(DocxToolbarState);
  private readonly context = injectDocxEditorContext();
  private readonly host: HTMLElement = inject(ElementRef).nativeElement;
  private readonly label = injectToolbarLabel();
  private readonly trigger = viewChild.required<ElementRef<HTMLButtonElement>>('trigger');
  private readonly panel = viewChild<ElementRef<HTMLElement>>('panel');
  private focusOnOpen = false;

  /** @internal */
  protected readonly paths = MORE_PATHS;
  /** @internal */
  protected readonly panelId = `docx-toolbar-overflow-${(nextPanelId += 1)}`;
  /** @internal */
  protected readonly open = this.state.moreOpen;
  /** @internal */
  protected readonly text = computed(() => this.label()('formattingBar.more'));
  /** @internal */
  protected readonly placement = signal<ToolbarPanelPlacement | null>(null);
  /** @internal */
  protected readonly guard: (event: Event) => void = guardToolbarMousedown;

  constructor() {
    const zone = inject(NgZone);
    afterRenderEffect((onCleanup) => {
      this.state.moreTrigger = this.trigger().nativeElement;
      onCleanup(() => {
        if (this.state.moreTrigger === this.trigger().nativeElement) this.state.moreTrigger = null;
      });
    });
    afterRenderEffect((onCleanup) => {
      const panel = this.panel()?.nativeElement;
      if (!this.open() || !panel) {
        untracked(() => this.placement.set(null));
        return;
      }
      untracked(() =>
        zone.runOutsideAngular(() => {
          const view = this.host.ownerDocument.defaultView;
          // Clamped into the viewport. The stylesheet lines the panel up with the trigger's
          // end, which runs off the left edge when the bar is centered or narrow.
          const place = () => {
            const rect = this.trigger().nativeElement.getBoundingClientRect();
            if (!view) return;
            const next = toolbarPanelPlacement({
              triggerLeft: rect.left,
              triggerRight: rect.right,
              panelWidth: panel.offsetWidth,
              viewportWidth: view.innerWidth,
            });
            // In the root's own coordinates, because the panel is positioned against it.
            const left = next.left - this.host.getBoundingClientRect().left;
            zone.run(() => this.placement.set({ ...next, left }));
          };
          place();
          view?.addEventListener('resize', place);
          const contains = (node: Node) => this.host.contains(node);
          const close = () => zone.run(() => this.state.moreOpen.set(false));
          // Escape in the capture phase, ahead of the surface. An open nested popup (a table
          // menu, a picker) takes this Escape, and the panel stays open.
          const stopEscape = listenForPopupEscape({
            popup: this.host,
            contains,
            chromeRoot: () => this.host.closest('.docx-editor'),
            editorElements: () => [this.context?.contentElement()],
            skip: () => hasOpenNestedPopup(panel),
            close: (fromInside) => zone.run(() => this.state.closeMore(fromInside)),
          });
          // Focus that leaves the panel, such as Ctrl+F into the find field, closes it.
          const stopFocus = listenForPopupFocusLeave({ popup: this.host, contains, close });
          const stopPress = listenForOutsidePress({ popup: this.host, close });
          if (this.focusOnOpen) {
            this.focusOnOpen = false;
            panel.querySelector<HTMLElement>(FOCUSABLE)?.focus();
          }
          onCleanup(() => {
            view?.removeEventListener('resize', place);
            stopEscape();
            stopFocus();
            stopPress();
          });
        })
      );
    });
  }

  /** @internal */
  protected toggle(): void {
    this.open.update((open) => !open);
  }

  /** @internal */
  protected onTriggerKey(event: KeyboardEvent): void {
    if (event.key !== 'ArrowDown') return;
    event.preventDefault();
    this.focusOnOpen = true;
    this.open.set(true);
  }

  /** @internal */
  protected onPanelKey(event: KeyboardEvent): void {
    if (event.key !== 'Escape' || event.defaultPrevented) return;
    event.preventDefault();
    this.state.closeMore(true);
  }
}
