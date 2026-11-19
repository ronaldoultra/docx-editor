import {
  ChangeDetectionStrategy,
  Component,
  ElementRef,
  booleanAttribute,
  computed,
  inject,
  input,
  untracked,
  type OnInit,
} from '@angular/core';
import type { EditorCommand } from '@docx-editor.dev/core/contracts/editor';
import {
  chromeProbeForSlot,
  chromeSlotIsToggle,
  commandForSlot,
  type ChromeSlotId,
} from '@docx-editor.dev/core/editor';
import { isApplePlatform } from '@docx-editor.dev/i18n';
import { injectDocxEditor } from '../context';
import { injectEditorCommand } from '../editor-command';
import { injectEditorState } from '../editor-state';
import { chromeControlForSlot, toolbarDevWarning } from '../toolbar/toolbar-registry';
import { DocxMenuRowComponent } from './menu-row.component';
import { DocxMenuBarState, injectMenuLabel, type MenuHandlers } from './menu-state';

/** The Go To chord's request: refused while the Go To UI turned its shortcut off. */
const GO_TO_CHORD: EditorCommand = { type: 'openGoTo', fromShortcut: true };

/** The radio sets: one alignment, one direction, one review display. */
const RADIO_COMMANDS = new Set(['setAlignment', 'setParagraphDirection', 'setReviewDisplayMode']);

/**
 * The rows that run a host handler and not an engine command. The engine says, correctly, that
 * opening or saving a file is not a command, so a command row would stay disabled.
 */
const ACTION_ROWS: Partial<
  Record<ChromeSlotId, { labelKey?: string; shortcutKey?: string; handler: keyof MenuHandlers }>
> = {
  'file.open': { labelKey: 'toolbar.open', shortcutKey: 'toolbar.openShortcut', handler: 'open' },
  'file.save': { labelKey: 'toolbar.save', shortcutKey: 'toolbar.saveShortcut', handler: 'save' },
};

/**
 * The rows that open a dialog. The slot's probe command asks the engine, and the row is also
 * disabled while no dialog handler exists.
 */
const DIALOG_ROWS: Partial<
  Record<ChromeSlotId, { labelKey: string; handler: keyof MenuHandlers }>
> = {
  'file.pageSetup': { labelKey: 'file.pageSetup', handler: 'pageSetup' },
  'paragraph.dialog': { labelKey: 'lineSpacing.options', handler: 'paragraphDialog' },
  'insert.watermark': { labelKey: 'toolbar.watermark', handler: 'watermarkDialog' },
};

/**
 * The rows that the Angular adapter cannot run yet: export and print come with the export and
 * print dialogs, and the image row with the image insert control. They render disabled.
 */
const PENDING_ROWS: ReadonlySet<ChromeSlotId> = new Set<ChromeSlotId>([
  'file.exportMarkdown',
  'file.exportPdf',
  'file.print',
  'image.insert',
]);

/** The default shortcut keys of the rows that have one. */
const DEFAULT_SHORTCUTS: Partial<Record<ChromeSlotId, string>> = {
  'insert.bookmark': 'toolbar.bookmarkShortcut',
};

/**
 * One chrome slot as a menu row. The engine decides whether the row is enabled, and the
 * registry gives its label and icon. Choosing the row runs the slot's command and closes the
 * menu bar.
 *
 * Some slots run something else: `file.open` and `file.save` run the bar's handlers,
 * `file.pageSetup`, `paragraph.dialog`, and `insert.watermark` open their dialog, and
 * `text.link` opens the link popup.
 *
 * @example
 * ```html
 * <docx-menu-item slotId="insert.pageBreak" />
 * <docx-menu-item slotId="text.bold" labelKey="formattingBar.bold" className="my-row" />
 * ```
 *
 * @public
 */
@Component({
  selector: 'docx-menu-item',
  changeDetection: ChangeDetectionStrategy.OnPush,
  imports: [DocxMenuRowComponent],
  host: { style: 'display: contents' },
  template: `@if (!hidden() && slotId()) {
    <button
      [docxMenuRow]="slotId()!"
      [icon]="paths()"
      [shortcut]="shortcut()"
      [disabled]="!enabled()"
      [reason]="reason()"
      [active]="active()"
      [radio]="radio()"
      [rowClass]="className()"
      [attr.data-value]="value()"
      (activate)="run()"
    >
      {{ text() }}
    </button>
  }`,
})
export class DocxMenuItemComponent implements OnInit {
  /** The chrome slot of the row, such as `'insert.pageBreak'`. */
  readonly slotId = input<ChromeSlotId | undefined>(undefined);
  /** A plain-label i18n key, in place of the slot's tooltip label. */
  readonly labelKey = input<string | undefined>(undefined);
  /** The i18n key of the shortcut in the right column. */
  readonly shortcutKey = input<string | undefined>(undefined);
  /** More classes for the row. */
  readonly className = input<string | undefined>(undefined);
  /** Render nothing. */
  readonly hidden = input(false, { transform: booleanAttribute });

  private readonly editor = injectDocxEditor();
  private readonly host: HTMLElement = inject(ElementRef).nativeElement;
  private readonly bar = inject(DocxMenuBarState, { optional: true });
  private readonly label = injectMenuLabel();
  private readonly handlers = computed(() => this.bar?.handlers() ?? { reportIssueRow: true });
  private readonly command = injectEditorCommand(() => {
    const slot = this.slotId() ?? ('' as ChromeSlotId);
    return DIALOG_ROWS[slot] ? (chromeProbeForSlot(slot) ?? slot) : slot;
  });
  private readonly chord = injectEditorCommand(GO_TO_CHORD);
  // The link row says Edit link when the caret is in a link.
  private readonly onLink = injectEditorState(() => {
    const editor = this.editor();
    // Outside the reactive context: a read can flush pending input.
    return untracked(() => editor?.surface?.hyperlinks.linkAtCaret() != null);
  });

  /** @internal */
  ngOnInit(): void {
    if (!this.slotId()) toolbarDevWarning('A docx-menu-item has no slotId, so it renders nothing.');
  }

  private readonly kind = computed(() => {
    const slot = this.slotId();
    if (!slot) return 'none';
    if (ACTION_ROWS[slot]) return 'action';
    if (DIALOG_ROWS[slot]) return 'dialog';
    if (PENDING_ROWS.has(slot)) return 'pending';
    return 'command';
  });

  /** @internal */
  protected readonly paths = computed(() => {
    const slot = this.slotId();
    return slot ? (chromeControlForSlot(slot)?.paths ?? null) : null;
  });

  /** @internal */
  protected readonly text = computed(() => {
    const slot = this.slotId();
    if (!slot) return '';
    const label = this.label();
    const own = this.labelKey();
    if (own) return label(own);
    if (slot === 'text.link') {
      return label(this.onLink() ? 'hyperlinkPopup.editLink' : 'formattingBar.insertLink');
    }
    const key =
      ACTION_ROWS[slot]?.labelKey ??
      chromeControlForSlot(slot)?.labelKey ??
      DIALOG_ROWS[slot]?.labelKey ??
      slot;
    return label(key);
  });

  /** @internal */
  protected readonly shortcut = computed(() => {
    const slot = this.slotId();
    if (!slot) return undefined;
    let key = this.shortcutKey() ?? ACTION_ROWS[slot]?.shortcutKey ?? DEFAULT_SHORTCUTS[slot];
    // Go To shows its chord only while the chord opens Go To.
    if (!key && slot === 'insert.goToBookmark' && this.chord.isEnabled()) {
      key = isApplePlatform()
        ? 'toolbar.goToBookmarkShortcutApple'
        : 'toolbar.goToBookmarkShortcut';
    }
    return key ? this.label()(key) : undefined;
  });

  /** @internal */
  protected readonly enabled = computed(() => {
    const slot = this.slotId();
    if (!slot) return false;
    switch (this.kind()) {
      case 'action':
        return !!this.handlers()[ACTION_ROWS[slot]!.handler];
      case 'dialog':
        return this.command.isEnabled() && !!this.handlers()[DIALOG_ROWS[slot]!.handler];
      case 'pending':
        return false;
      default:
        return this.command.isEnabled();
    }
  });

  /** @internal */
  protected readonly reason = computed(() => {
    const kind = this.kind();
    return kind === 'command' || kind === 'dialog' ? this.command.disabledReason() : null;
  });

  /** @internal */
  protected readonly active = computed(() => {
    const slot = this.slotId();
    if (!slot || this.kind() !== 'command' || !chromeSlotIsToggle(slot)) return undefined;
    return this.command.isActive();
  });

  /** @internal */
  protected readonly radio = computed(() => {
    const slot = this.slotId();
    return !!slot && RADIO_COMMANDS.has(commandForSlot(slot)?.type ?? '');
  });

  /** @internal */
  protected readonly value = computed(() =>
    this.kind() === 'command' ? this.command.value() : null
  );

  /** @internal */
  protected run(): void {
    const slot = untracked(this.slotId);
    if (!slot) return;
    const handlers = untracked(this.handlers);
    const kind = untracked(this.kind);
    if (kind === 'action') {
      const handler = handlers[ACTION_ROWS[slot]!.handler] as (() => void) | undefined;
      this.close();
      handler?.();
      return;
    }
    if (kind === 'dialog') {
      const handler = handlers[DIALOG_ROWS[slot]!.handler] as
        ((trigger: HTMLElement | null) => void) | undefined;
      const trigger = this.menuTrigger();
      this.close();
      handler?.(trigger);
      return;
    }
    if (slot === 'text.link') {
      // Close first: closing gives the focus back to the document, and the link popup then
      // moves it to its address field.
      this.close();
      this.command.execute();
      return;
    }
    this.command.execute();
    this.close();
  }

  private menuTrigger(): HTMLElement | null {
    return (
      this.host.closest('[data-menu]')?.querySelector<HTMLElement>('.docx-menubar__trigger') ?? null
    );
  }

  private close(): void {
    this.bar?.setOpenMenu(null);
  }
}
