import {
  ChangeDetectionStrategy,
  Component,
  booleanAttribute,
  computed,
  inject,
  input,
  output,
} from '@angular/core';
import { platformShortcut } from '@docx-editor.dev/core/i18n';
import { DocxChromeIconComponent } from '../toolbar/chrome-icon.component';
import { popupId } from '../toolbar/toolbar-popup';
import { guardToolbarMousedown } from '../toolbar/toolbar-registry';
import { DOCX_MENU_ROW_TABINDEX, injectMenuLabel } from './menu-state';

/**
 * One presentational menu row: an icon column, the label, and a shortcut column.
 *
 * The icon column stays when a row has no icon, so the labels line up. A disabled row keeps
 * the focus (`aria-disabled`, not the native attribute), and its `reason` is the accessible
 * description, so a keyboard user hears why. Set `active` only on a row that toggles, and
 * `radio` on a row of a one-of-many set, such as the alignments.
 *
 * @example
 * ```html
 * <button docxMenuRow="host.share" [icon]="sharePaths" (activate)="share()">Share</button>
 * ```
 *
 * @public
 */
@Component({
  selector: 'button[docxMenuRow]',
  changeDetection: ChangeDetectionStrategy.OnPush,
  imports: [DocxChromeIconComponent],
  host: {
    type: 'button',
    '[attr.role]': 'role()',
    '[class]': 'className()',
    '[attr.tabindex]': 'tabIndex',
    '[attr.data-slot]': 'docxMenuRow() || null',
    '[attr.data-active]': "active() ? '' : null",
    '[attr.data-disabled]': "disabled() ? '' : null",
    '[attr.aria-disabled]': "disabled() ? 'true' : null",
    '[attr.aria-checked]': 'active() === undefined ? null : active()',
    '[attr.aria-describedby]': 'described() ? reasonId : null',
    '[attr.title]': 'reason() || null',
    '(mousedown)': 'guard($event)',
    '(click)': 'onClick()',
  },
  template: `<span class="docx-menubar__item-icon" aria-hidden="true">
      @if (icon(); as paths) {
        <svg [docxChromeIcon]="paths"></svg>
      }
      <ng-content select="[docxMenuRowIcon]" />
    </span>
    <span class="docx-menubar__item-label"><ng-content /></span>
    @if (shortcutText(); as text) {
      <span class="docx-menubar__item-shortcut">{{ text }}</span>
    }
    @if (described()) {
      <span class="docx-editor-sr-only" [id]="reasonId">{{ reason() }}</span>
    }`,
})
export class DocxMenuRowComponent {
  /** The row's `data-slot` marker for hosts and tests. */
  readonly docxMenuRow = input<string>('');
  /** Material Symbols paths for the icon column. */
  readonly icon = input<readonly string[] | null | undefined>(undefined);
  /** The shortcut text. Ctrl names the reader's own modifier key. */
  readonly shortcut = input<string | undefined>(undefined);
  /** Disable the row. It stays focusable. */
  readonly disabled = input(false, { transform: booleanAttribute });
  /** Why the row is disabled. Set it only for the engine's reason. */
  readonly reason = input<string | null | undefined>(undefined);
  /** The checked state of a row that toggles. Leave it unset on a row that only acts. */
  readonly active = input<boolean | undefined>(undefined);
  /** Mark the row as one of a set where only one is on. */
  readonly radio = input(false, { transform: booleanAttribute });
  /** More classes for the row. */
  readonly rowClass = input<string | undefined>(undefined);
  /** Emits when the user chooses an enabled row. */
  readonly activate = output<void>();

  /** @internal */
  protected readonly tabIndex = inject(DOCX_MENU_ROW_TABINDEX);
  /** @internal */
  protected readonly reasonId = popupId('menu-reason');
  /** @internal */
  protected readonly guard: (event: Event) => void = guardToolbarMousedown;
  /** @internal */
  protected readonly role = computed(() =>
    this.active() === undefined ? 'menuitem' : this.radio() ? 'menuitemradio' : 'menuitemcheckbox'
  );
  /** @internal */
  protected readonly className = computed(() => {
    const extra = this.rowClass();
    return `docx-toolbar__menu-item docx-menubar__item${extra ? ` ${extra}` : ''}`;
  });
  /** @internal */
  protected readonly described = computed(() => this.disabled() && !!this.reason());
  /** @internal */
  protected readonly shortcutText = computed(() => {
    const shortcut = this.shortcut();
    return shortcut ? platformShortcut(shortcut) : null;
  });

  /** @internal */
  protected onClick(): void {
    if (!this.disabled()) this.activate.emit();
  }
}

/**
 * A horizontal rule between groups of menu rows.
 *
 * @example
 * ```html
 * <div docxMenuSeparator></div>
 * ```
 *
 * @public
 */
@Component({
  selector: 'div[docxMenuSeparator]',
  changeDetection: ChangeDetectionStrategy.OnPush,
  host: { role: 'separator', class: 'docx-toolbar__menu-separator' },
  template: '',
})
export class DocxMenuSeparatorComponent {}

/**
 * A named section of a menu panel: a visible heading and the rows under it. The heading is
 * the group's accessible name.
 *
 * @example
 * ```html
 * <div docxMenuGroup label="Sharing">
 *   <button docxMenuRow (activate)="share()">Share</button>
 * </div>
 * ```
 *
 * @public
 */
@Component({
  selector: 'div[docxMenuGroup]',
  changeDetection: ChangeDetectionStrategy.OnPush,
  host: {
    role: 'group',
    class: 'docx-menubar__group',
    '[attr.aria-label]': 'text()',
  },
  template: `@if (text(); as heading) {
      <div class="docx-menubar__group-label" aria-hidden="true">{{ heading }}</div>
    }
    <ng-content />`,
})
export class DocxMenuGroupComponent {
  /** The heading text. It wins over `labelKey`. */
  readonly label = input<string | undefined>(undefined);
  /** The i18n key of the heading. */
  readonly labelKey = input<string | undefined>(undefined);

  private readonly menuLabel = injectMenuLabel();

  /** @internal */
  protected readonly text = computed(() => {
    const key = this.labelKey();
    return this.label() ?? (key === undefined ? null : this.menuLabel()(key));
  });
}
