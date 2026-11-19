import { NgTemplateOutlet } from '@angular/common';
import {
  ChangeDetectionStrategy,
  Component,
  computed,
  contentChildren,
  effect,
  inject,
  signal,
  untracked,
  viewChildren,
  type Signal,
} from '@angular/core';
import {
  NO_MIXED_FIELDS,
  TAB_ALIGNMENT_LABELS,
  alignmentAfterDirectionChange,
  changedFields,
  formatInches,
  inchesToTwips,
  isHTMLElementNode,
  mixedFieldsOf,
  seedFields,
  twipsToInches,
  withTabStop,
  type ParagraphDialogFields,
  type ParagraphDialogMixed,
  type ParagraphFlagKey,
  type ParagraphTabStop,
  type TabAlignment,
  type TabLeaderName,
} from '@docx-editor.dev/core/editor';
import {
  DocxDialogApplyDirective,
  DocxDialogCancelDirective,
  DocxDialogDefaultDirective,
  DocxDialogFieldDirective,
  DocxDialogOutletDirective,
  DocxDialogPartDirective,
  DocxDialogParts,
  injectDialog,
  provideDialogParts,
  type DocxDialogState,
} from '../dialog/dialog-parts';
import { injectDialogGeneration } from '../dialog/dialog-generation';
import { DocxNativeDialogComponent } from '../dialog/native-dialog.component';
import { DocxEditorDialogBase } from '../dialog/packaged-dialog';
import { DocxPartSlotDirective } from '../parts/part-slot';
import { injectParagraphFormat, type ParagraphFormatState } from './paragraph-format';

/**
 * The draft state, the mixed fields, and the actions of the enclosing paragraph dialog.
 *
 * @public
 */
export interface ParagraphDialogState extends DocxDialogState<ParagraphDialogFields> {
  /**
   * The fields that the selection's paragraphs disagree about and that the user did not set
   * yet. Apply leaves such a field as it is in each paragraph.
   */
  readonly mixed: Signal<ParagraphDialogMixed>;
}

/**
 * The state of the enclosing paragraph dialog. Call it in an injection context inside the
 * dialog, such as in a component in a part template.
 *
 * @public
 */
export function injectParagraphDialog(): ParagraphDialogState {
  return injectDialog<ParagraphDialogFields>() as ParagraphDialogState;
}

/** A field that the selection can disagree about. @internal */
export type ParagraphMixedKey = keyof ParagraphDialogMixed;
/** An indent field of the paragraph draft, in inches on screen. @internal */
export type ParagraphInchField = 'indentLeft' | 'indentRight';
/** A spacing field of the paragraph draft, in points. @internal */
export type ParagraphPointField = 'spaceBefore' | 'spaceAfter';

/** The tab stop entry row of the paragraph dialog. @internal */
export interface ParagraphTabEntry {
  readonly position: number;
  readonly alignment: TabAlignment;
  readonly leader: TabLeaderName;
}

const EMPTY_TAB_ENTRY: ParagraphTabEntry = { position: 0, alignment: 'left', leader: 'none' };

const DEFAULT_FIELDS: ParagraphDialogFields = {
  alignment: 'left',
  direction: 'ltr',
  indentLeft: 0,
  indentRight: 0,
  special: 'none',
  specialBy: 0,
  spaceBefore: 0,
  spaceAfter: 0,
  lineRule: 'multiple',
  lineValue: 1.08,
  contextualSpacing: false,
  keepNext: false,
  keepLines: false,
  widowControl: true,
  pageBreakBefore: false,
  tabStops: [],
  clearedAllTabStops: false,
};

/** The line value a rule opens with: the seeded one for the seeded rule. */
function lineValueFor(rule: ParagraphDialogFields['lineRule'], seed: ParagraphDialogFields | null) {
  if (rule === seed?.lineRule) return seed.lineValue;
  return rule === 'multiple' ? 1.08 : 12;
}

let nextId = 0;

/**
 * The paragraph dialog: alignment, direction, indentation, tab stops, spacing, line spacing,
 * and the pagination flags, applied as one undo step. Apply writes only the fields that the
 * user changed, and a field that the selection disagrees about stays mixed until the user sets
 * it. The editor opens it from Format › Paragraph and from the line spacing menu.
 *
 * Replace or restyle a part with `docxDialogPart` templates. The part names are `header`,
 * `title`, `body`, `footer`, `error`, `cancel`, `apply`, and `field:<name>` for `alignment`,
 * `direction`, `indentLeft`, `indentRight`, `special`, `specialBy`, `tabStops`, `spaceBefore`,
 * `spaceAfter`, `lineRule`, `lineValue`, `contextualSpacing`, `keepNext`, `widowControl`,
 * `keepLines`, and `pageBreakBefore`. In your own layout, `docxDialogOutlet` renders one of them.
 *
 * @public
 */
@Component({
  selector: 'docx-editor-paragraph-dialog',
  imports: [
    NgTemplateOutlet,
    DocxNativeDialogComponent,
    DocxPartSlotDirective,
    DocxDialogDefaultDirective,
    DocxDialogOutletDirective,
    DocxDialogApplyDirective,
    DocxDialogCancelDirective,
    DocxDialogFieldDirective,
  ],
  providers: provideDialogParts(),
  changeDetection: ChangeDetectionStrategy.OnPush,
  host: { style: 'display: contents' },
  templateUrl: './paragraph-dialog.component.html',
})
export class DocxEditorParagraphDialogComponent extends DocxEditorDialogBase {
  /** @internal */
  protected readonly id: string = `docx-paragraph-${++nextId}`;
  /** @internal */
  protected readonly inchFields: readonly { field: ParagraphInchField; key: string }[] = [
    { field: 'indentLeft', key: 'beforeText' },
    { field: 'indentRight', key: 'afterText' },
  ];
  /** @internal */
  protected readonly pointFields: readonly ParagraphPointField[] = ['spaceBefore', 'spaceAfter'];
  /** @internal */
  protected readonly flags: readonly ParagraphFlagKey[] = [
    'contextualSpacing',
    'keepNext',
    'widowControl',
    'keepLines',
    'pageBreakBefore',
  ];
  /** @internal */
  protected readonly tabLabels: typeof TAB_ALIGNMENT_LABELS = TAB_ALIGNMENT_LABELS;
  /** @internal */
  protected readonly format: ParagraphFormatState = injectParagraphFormat();
  /** @internal */
  protected readonly values = signal<ParagraphDialogFields>({ ...DEFAULT_FIELDS });
  /** @internal */
  protected readonly mixed = signal<ParagraphDialogMixed>(NO_MIXED_FIELDS);
  /** @internal */
  protected readonly indentUnknown = signal(false);
  /** @internal */
  protected readonly refused = signal(false);
  /** @internal */
  protected readonly tabEntry = signal<ParagraphTabEntry>(EMPTY_TAB_ENTRY);
  /** @internal */
  protected readonly state: ParagraphDialogState = {
    values: this.values.asReadonly(),
    mixed: this.mixed.asReadonly(),
    errors: computed(() =>
      this.refused() ? { form: this.label('dialogs.paragraph.refused') } : {}
    ),
    isEnabled: this.format.isEnabled,
    setValue: (name, value) => this.setValue(name, value),
    apply: () => this.applyDraft(),
    cancel: () => this.close(),
  };

  /** What the selection disagreed about when the dialog opened. */
  private seedMixed: ParagraphDialogMixed = NO_MIXED_FIELDS;
  private seed: ParagraphDialogFields | null = null;
  private readonly parts = contentChildren(DocxDialogPartDirective);
  private readonly defaults = viewChildren(DocxDialogDefaultDirective);
  private readonly current = injectDialogGeneration(this.open, () => this.close());

  constructor() {
    super();
    const parts = inject(DocxDialogParts, { self: true });
    parts.connect(this.state as unknown as DocxDialogState<object>, this.parts);
    parts.connectDefaults(this.defaults);
    // Seed when the dialog opens, never on later ticks, or a concurrent edit would fight the
    // user's typing.
    let seeded = false;
    effect(() => {
      const open = this.open();
      const format = this.format.format();
      untracked(() => {
        if (!open) {
          seeded = false;
          return;
        }
        if (seeded || format === null) return;
        const seed = seedFields(format);
        const mixed = mixedFieldsOf(format);
        this.values.set({ ...seed, clearedAllTabStops: false });
        this.mixed.set(mixed);
        this.indentUnknown.set(format.indentUnknown);
        // The entry row is a scratch pad, not a setting.
        this.tabEntry.set(EMPTY_TAB_ENTRY);
        this.seedMixed = mixed;
        this.seed = seed;
        this.refused.set(false);
        seeded = true;
      });
    });
  }

  /** @internal */
  protected blankLabel(key: ParagraphMixedKey): string {
    // An indent that the engine cannot place, such as in a table cell, is no disagreement.
    return this.indentUnknown() &&
      (key === 'indentLeft' || key === 'indentRight' || key === 'special')
      ? this.label('dialogs.paragraph.notShown')
      : this.label('dialogs.paragraph.mixed');
  }

  /** @internal */
  protected inches(twips: number): number {
    return twipsToInches(twips);
  }

  /** @internal */
  protected tabPositionLabel(stop: ParagraphTabStop): string {
    return formatInches(stop.positionTwips);
  }

  /** @internal */
  protected target(event: Event): HTMLInputElement {
    return event.target as HTMLInputElement;
  }

  /** @internal */
  protected setInches(
    field: ParagraphInchField | 'specialBy',
    mixedKey: ParagraphMixedKey,
    event: Event
  ): void {
    this.resolve(mixedKey);
    const twips = inchesToTwips(Number(this.target(event).value) || 0);
    this.patch({ [field]: field === 'specialBy' ? Math.max(0, twips) : twips });
  }

  /** @internal */
  protected setPoints(field: ParagraphPointField, event: Event): void {
    this.resolve(field);
    this.patch({ [field]: Math.max(0, Number(this.target(event).value) || 0) });
  }

  /** @internal */
  protected setFlag(field: ParagraphFlagKey, event: Event): void {
    // Touching the box resolves the disagreement: from here it means what it shows.
    this.resolve(field);
    this.patch({ [field]: this.target(event).checked });
  }

  /** @internal */
  protected setChoice(field: 'alignment' | 'special', event: Event): void {
    this.resolve(field);
    this.patch({ [field]: this.target(event).value });
  }

  /** @internal */
  protected setDirection(event: Event): void {
    this.resolve('direction');
    this.changeDirection(this.target(event).value as 'ltr' | 'rtl');
  }

  /** @internal */
  protected setLineRule(event: Event): void {
    // Picking a rule resolves the disagreement, and the value follows the rule: lines under
    // Multiple, points otherwise.
    const rule = this.target(event).value as ParagraphDialogFields['lineRule'];
    this.resolve('lineSpacing');
    this.patch({ lineRule: rule, lineValue: lineValueFor(rule, this.seed) });
  }

  /** @internal */
  protected setLineValue(event: Event): void {
    const next = this.target(event).value.trim();
    const rule = untracked(this.values).lineRule;
    this.patch({ lineValue: next === '' ? (rule === 'multiple' ? 1.08 : 12) : Number(next) || 0 });
  }

  /** @internal */
  protected setTabEntry(change: Partial<ParagraphTabEntry>): void {
    this.tabEntry.update((entry) => ({ ...entry, ...change }));
  }

  /** @internal */
  protected setTabPosition(event: Event): void {
    this.setTabEntry({
      position: Math.max(0, inchesToTwips(Number(this.target(event).value) || 0)),
    });
  }

  /** @internal */
  protected removeTabStop(stop: ParagraphTabStop): void {
    this.patch({ tabStops: untracked(this.values).tabStops.filter((entry) => entry !== stop) });
  }

  /** @internal */
  protected clearTabStops(): void {
    // Over a mixed list this is still a decision, even though the list already shows empty.
    this.resolve('tabStops');
    this.patch({ tabStops: [], clearedAllTabStops: true });
  }

  /**
   * Add or replace the stop at the entry's position. A position of zero is no stop: the field
   * opens at 0, so Set without a typed position adds nothing.
   *
   * @internal
   */
  protected setTabStop(): boolean {
    const entry = untracked(this.tabEntry);
    const positionTwips = Math.round(entry.position);
    if (positionTwips <= 0) return false;
    this.resolve('tabStops');
    this.patch({
      tabStops: withTabStop(untracked(this.values).tabStops, {
        positionTwips,
        alignment: entry.alignment,
        ...(entry.leader !== 'none' ? { leader: entry.leader } : {}),
      }),
    });
    this.tabEntry.set(EMPTY_TAB_ENTRY);
    return true;
  }

  /** @internal */
  protected onKeydown(event: KeyboardEvent): void {
    const target = event.target;
    if (
      event.key !== 'Enter' ||
      event.isComposing ||
      (isHTMLElementNode(target) &&
        (target.localName === 'button' || target.localName === 'select'))
    ) {
      return;
    }
    event.preventDefault();
    if (isHTMLElementNode(target) && target.dataset['docxTabEntry'] !== undefined) {
      this.setTabStop();
      return;
    }
    if (this.format.isEnabled()) this.applyDraft();
  }

  private setValue<K extends keyof ParagraphDialogFields>(
    name: K,
    value: ParagraphDialogFields[K]
  ): void {
    if (name === 'direction') this.changeDirection(value as 'ltr' | 'rtl');
    else this.patch({ [name]: value } as Partial<ParagraphDialogFields>);
    if (name === 'lineRule') {
      this.patch({
        lineValue: lineValueFor(value as ParagraphDialogFields['lineRule'], this.seed),
      });
    }
    if (name === 'lineValue') return;
    const mixedName = name === 'lineRule' ? 'lineSpacing' : name === 'specialBy' ? 'special' : name;
    if (mixedName in NO_MIXED_FIELDS) this.resolve(mixedName as ParagraphMixedKey);
  }

  /**
   * A new direction trades the leading and trailing edges, so the alignment follows it, as the
   * toolbar's direction buttons do.
   */
  private changeDirection(next: 'ltr' | 'rtl'): void {
    const current = untracked(this.values);
    const follow =
      next !== current.direction && !this.seedMixed.direction && !this.seedMixed.alignment;
    this.patch({
      direction: next,
      ...(follow ? { alignment: alignmentAfterDirectionChange(current.alignment) } : {}),
    });
  }

  private resolve(key: ParagraphMixedKey): void {
    this.mixed.update((mixed) => (mixed[key] ? { ...mixed, [key]: false } : mixed));
  }

  private patch(change: Partial<ParagraphDialogFields>): void {
    this.values.update((values) => ({ ...values, ...change }));
  }

  private applyDraft(): void {
    if (!this.current()) return;
    const seed = this.seed;
    const update =
      seed === null
        ? null
        : changedFields(seed, untracked(this.values), this.seedMixed, untracked(this.mixed));
    // Nothing moved, so there is nothing to write. Closing is the whole job.
    if (update === null) {
      this.close();
      return;
    }
    // A refused write keeps the dialog open and says so.
    if (this.format.apply(update)) {
      this.refused.set(false);
      this.close();
      return;
    }
    this.refused.set(true);
  }
}
