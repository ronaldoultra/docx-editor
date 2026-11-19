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
import { isHTMLElementNode } from '@docx-editor.dev/core/editor';
import { injectDialogGeneration } from '../dialog/dialog-generation';
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
import { DocxNativeDialogComponent } from '../dialog/native-dialog.component';
import { DocxEditorDialogBase } from '../dialog/packaged-dialog';
import { DocxPartSlotDirective } from '../parts/part-slot';
import { injectPageSetup, type PageSetupState } from './page-setup';

const PAGE_SIZES = [
  { labelKey: 'dialogs.pageSetup.pageSizes.letter', width: 12240, height: 15840 },
  { labelKey: 'dialogs.pageSetup.pageSizes.a4', width: 11906, height: 16838 },
  { labelKey: 'dialogs.pageSetup.pageSizes.legal', width: 12240, height: 20160 },
  { labelKey: 'dialogs.pageSetup.pageSizes.a3', width: 16838, height: 23811 },
  { labelKey: 'dialogs.pageSetup.pageSizes.a5', width: 8391, height: 11906 },
  { labelKey: 'dialogs.pageSetup.pageSizes.b5', width: 9979, height: 14175 },
  { labelKey: 'dialogs.pageSetup.pageSizes.executive', width: 10440, height: 15120 },
] as const;

const TWIPS_PER_INCH = 1440;
const DEFAULT_WIDTH = 12240;
const DEFAULT_HEIGHT = 15840;
const DEFAULT_MARGIN = 1440;

const twipsToInches = (twips: number): number => Math.round((twips / TWIPS_PER_INCH) * 100) / 100;
const inchesToTwips = (inches: number): number => Math.round(inches * TWIPS_PER_INCH);

function pageSizeIndex(width: number, height: number): number {
  const short = Math.min(width, height);
  const long = Math.max(width, height);
  return PAGE_SIZES.findIndex(
    (size) => Math.abs(size.width - short) < 20 && Math.abs(size.height - long) < 20
  );
}

/**
 * The page setup draft, in twips.
 *
 * @public
 */
export interface PageSetupDialogFields {
  pageWidth: number;
  pageHeight: number;
  orientation: 'portrait' | 'landscape';
  marginTop: number;
  marginBottom: number;
  marginLeft: number;
  marginRight: number;
  scope: 'document' | 'section';
}

/** A margin field of the page setup draft. @internal */
export type PageSetupMarginField = 'marginTop' | 'marginBottom' | 'marginLeft' | 'marginRight';

const DEFAULT_FIELDS: PageSetupDialogFields = {
  pageWidth: DEFAULT_WIDTH,
  pageHeight: DEFAULT_HEIGHT,
  orientation: 'portrait',
  marginTop: DEFAULT_MARGIN,
  marginBottom: DEFAULT_MARGIN,
  marginLeft: DEFAULT_MARGIN,
  marginRight: DEFAULT_MARGIN,
  scope: 'document',
};

/**
 * The draft state and the actions of the enclosing page setup dialog. Call it in an injection
 * context inside the dialog, such as in a component in a part template.
 *
 * @public
 */
export function injectPageSetupDialog(): DocxDialogState<PageSetupDialogFields> {
  return injectDialog<PageSetupDialogFields>();
}

/**
 * The page setup dialog: page size, orientation, margins, and the scope, applied as one undo
 * step. The editor opens it from File › Page setup. Show it yourself with an `@if` block or the
 * `open` input, and hide it on `closed`.
 *
 * Replace or restyle a part with `docxDialogPart` templates. The part names are `header`,
 * `title`, `body`, `footer`, `error`, `cancel`, `apply`, and `field:pageSize`,
 * `field:orientation`, `field:marginTop`, `field:marginBottom`, `field:marginLeft`,
 * `field:marginRight`, and `field:scope`. In your own layout, `docxDialogOutlet` renders one of
 * them.
 *
 * @example
 * ```html
 * @if (pageSetupOpen()) {
 *   <docx-editor-page-setup-dialog (closed)="pageSetupOpen.set(false)" />
 * }
 * ```
 *
 * @public
 */
@Component({
  selector: 'docx-editor-page-setup-dialog',
  imports: [
    NgTemplateOutlet,
    DocxNativeDialogComponent,
    DocxPartSlotDirective,
    DocxDialogDefaultDirective,
    DocxDialogApplyDirective,
    DocxDialogCancelDirective,
    DocxDialogFieldDirective,
    DocxDialogOutletDirective,
  ],
  providers: provideDialogParts(),
  changeDetection: ChangeDetectionStrategy.OnPush,
  host: { style: 'display: contents' },
  template: `
    <ng-template docxDialogDefault="header" #header>
      <div
        *docxPartSlot="'header'; let content"
        data-docx-part="header"
        class="docx-dialog__header"
      >
        @if (content) {
          <ng-container [ngTemplateOutlet]="content" />
        } @else {
          <span
            *docxPartSlot="'title'; let title"
            data-docx-part="title"
            class="docx-dialog__title"
          >
            @if (title) {
              <ng-container [ngTemplateOutlet]="title" />
            } @else {
              {{ label('dialogs.pageSetup.title') }}
            }
          </span>
        }
      </div>
    </ng-template>

    <ng-template docxDialogDefault="field:pageSize" #pageSize>
      <label
        *docxPartSlot="'field:pageSize'; let content"
        docxDialogField="pageSize"
        class="docx-dialog__row"
      >
        @if (content) {
          <ng-container [ngTemplateOutlet]="content" />
        } @else {
          <span class="docx-dialog__label">{{ label('dialogs.pageSetup.sizeLabel') }}</span>
          <select
            class="docx-dialog__input"
            [attr.aria-label]="label('dialogs.pageSetup.sizeLabel')"
            (change)="setSize($event)"
          >
            @for (size of sizes; track size.labelKey; let index = $index) {
              <option [value]="index" [selected]="index === sizeIndex()">
                {{ label(size.labelKey) }}
              </option>
            }
            @if (sizeIndex() < 0) {
              <option value="-1" selected>{{ label('dialogs.pageSetup.custom') }}</option>
            }
          </select>
        }
      </label>
    </ng-template>

    <ng-template docxDialogDefault="field:orientation" #orientation>
      <label
        *docxPartSlot="'field:orientation'; let content"
        docxDialogField="orientation"
        class="docx-dialog__row"
      >
        @if (content) {
          <ng-container [ngTemplateOutlet]="content" />
        } @else {
          <span class="docx-dialog__label">{{ label('dialogs.pageSetup.orientation') }}</span>
          <select
            class="docx-dialog__input"
            [attr.aria-label]="label('dialogs.pageSetup.orientation')"
            (change)="setOrientation($any(selectValue($event)))"
          >
            <option value="portrait" [selected]="values().orientation === 'portrait'">
              {{ label('dialogs.pageSetup.portrait') }}
            </option>
            <option value="landscape" [selected]="values().orientation === 'landscape'">
              {{ label('dialogs.pageSetup.landscape') }}
            </option>
          </select>
        }
      </label>
    </ng-template>

    @for (margin of margins; track margin.field) {
      <ng-template [docxDialogDefault]="'field:' + margin.field">
        <label
          *docxPartSlot="'field:' + margin.field; let content"
          [docxDialogField]="margin.field"
          class="docx-dialog__row"
        >
          @if (content) {
            <ng-container [ngTemplateOutlet]="content" />
          } @else {
            <span class="docx-dialog__label">{{ label('dialogs.pageSetup.' + margin.side) }}</span>
            <input
              class="docx-dialog__input"
              type="number"
              min="0"
              max="22"
              step="0.1"
              [attr.aria-label]="label('dialogs.pageSetup.' + margin.side)"
              [value]="inches(margin.field)"
              (input)="setMargin(margin.field, $event)"
            />
            <span class="docx-dialog__unit">{{ label('dialogs.paragraph.unitInches') }}</span>
          }
        </label>
      </ng-template>
    }

    <ng-template docxDialogDefault="field:scope" #scope>
      <label
        *docxPartSlot="'field:scope'; let content"
        docxDialogField="scope"
        class="docx-dialog__row"
      >
        @if (content) {
          <ng-container [ngTemplateOutlet]="content" />
        } @else {
          <span class="docx-dialog__label">{{ label('dialogs.pageSetup.applyTo') }}</span>
          <select
            class="docx-dialog__input"
            [attr.aria-label]="label('dialogs.pageSetup.applyTo')"
            (change)="setScope($event)"
          >
            <option value="document" [selected]="values().scope === 'document'">
              {{ label('dialogs.pageSetup.applyToDocument') }}
            </option>
            <option value="section" [selected]="values().scope === 'section'">
              {{ label('dialogs.pageSetup.applyToSection') }}
            </option>
          </select>
        }
      </label>
    </ng-template>

    <ng-template docxDialogDefault="body" #body>
      <div *docxPartSlot="'body'; let content" data-docx-part="body" class="docx-dialog__body">
        @if (content) {
          <ng-container [ngTemplateOutlet]="content" />
        } @else {
          <div class="docx-dialog__section-label">{{ label('dialogs.pageSetup.pageSize') }}</div>
          <ng-container [ngTemplateOutlet]="pageSize" />
          <ng-container [ngTemplateOutlet]="orientation" />
          <div class="docx-dialog__section-label docx-dialog__section-label--spaced">
            {{ label('dialogs.pageSetup.margins') }}
          </div>
          @for (margin of margins; track margin.field) {
            <ng-container [docxDialogOutlet]="'field:' + margin.field" />
          }
          <ng-container [ngTemplateOutlet]="scope" />
        }
      </div>
    </ng-template>

    <ng-template docxDialogDefault="footer" #footer>
      <div
        *docxPartSlot="'footer'; let content"
        data-docx-part="footer"
        class="docx-dialog__footer"
      >
        @if (content) {
          <ng-container [ngTemplateOutlet]="content" />
        } @else {
          <span
            *docxPartSlot="'error'; let error"
            data-docx-part="error"
            class="docx-dialog__error"
            role="alert"
          >
            @if (error) {
              <ng-container [ngTemplateOutlet]="error" />
            } @else {
              {{ state.errors().form ?? '' }}
            }
          </span>
          <button
            *docxPartSlot="'cancel'; let cancel"
            docxDialogCancel
            class="docx-dialog__button docx-dialog__cancel"
          >
            @if (cancel) {
              <ng-container [ngTemplateOutlet]="cancel" />
            } @else {
              {{ label('common.cancel') }}
            }
          </button>
          <button
            *docxPartSlot="'apply'; let apply"
            docxDialogApply
            class="docx-dialog__button docx-dialog__apply"
          >
            @if (apply) {
              <ng-container [ngTemplateOutlet]="apply" />
            } @else {
              {{ label('common.apply') }}
            }
          </button>
        }
      </div>
    </ng-template>

    @if (open()) {
      <docx-native-dialog
        kind="pageSetup"
        [label]="label('dialogs.pageSetup.title')"
        [className]="className()"
        [styles]="styles()"
        (dismiss)="close()"
        (dialogKeydown)="onKeydown($event)"
      >
        @if (preset()) {
          <ng-container [ngTemplateOutlet]="header" />
          <ng-container [ngTemplateOutlet]="body" />
          <ng-container [ngTemplateOutlet]="footer" />
        }
        <ng-content />
      </docx-native-dialog>
    }
  `,
})
export class DocxEditorPageSetupDialogComponent extends DocxEditorDialogBase {
  /** @internal */
  protected readonly sizes: readonly {
    readonly labelKey: string;
    readonly width: number;
    readonly height: number;
  }[] = PAGE_SIZES;
  /** @internal */
  protected readonly margins: readonly { field: PageSetupMarginField; side: string }[] = [
    { field: 'marginTop', side: 'top' },
    { field: 'marginBottom', side: 'bottom' },
    { field: 'marginLeft', side: 'left' },
    { field: 'marginRight', side: 'right' },
  ];
  /** @internal */
  protected readonly values = signal<PageSetupDialogFields>({ ...DEFAULT_FIELDS });
  /** @internal */
  protected readonly refused = signal(false);
  /** @internal */
  protected readonly sizeIndex: Signal<number> = computed(() =>
    pageSizeIndex(this.values().pageWidth, this.values().pageHeight)
  );
  /** @internal */
  protected readonly setup: PageSetupState = injectPageSetup();
  /** @internal */
  protected readonly state: DocxDialogState<PageSetupDialogFields> = {
    values: this.values.asReadonly(),
    errors: computed(() =>
      this.refused() ? { form: this.label('dialogs.paragraph.refused') } : {}
    ),
    isEnabled: this.setup.isEnabled,
    setValue: (name, value) => {
      if (name === 'orientation') this.setOrientation(value as 'portrait' | 'landscape');
      else this.values.update((current) => ({ ...current, [name]: value }));
    },
    apply: () => this.applyDraft(),
    cancel: () => this.close(),
  };

  private readonly parts = contentChildren(DocxDialogPartDirective);
  private readonly defaults = viewChildren(DocxDialogDefaultDirective);
  private readonly current = injectDialogGeneration(this.open, () => this.close());

  constructor() {
    super();
    const parts = inject(DocxDialogParts, { self: true });
    parts.connect(this.state as unknown as DocxDialogState<object>, this.parts);
    parts.connectDefaults(this.defaults);
    // Seed when the dialog opens on a page setup, never on later ticks, or a concurrent edit
    // would fight the user's typing. A document that loads while the dialog is open seeds it.
    let seeded: 'no' | 'loading' | 'yes' = 'no';
    effect(() => {
      const open = this.open();
      const setup = this.setup.pageSetup();
      untracked(() => {
        if (!open) {
          this.refused.set(false);
          seeded = 'no';
          return;
        }
        if (seeded === 'yes' && setup === null) {
          seeded = 'no';
          return;
        }
        if (seeded === 'yes' || (seeded === 'loading' && setup === null)) return;
        this.values.set({
          pageWidth: setup?.pageWidthTwips ?? DEFAULT_WIDTH,
          pageHeight: setup?.pageHeightTwips ?? DEFAULT_HEIGHT,
          orientation: setup?.orientation ?? 'portrait',
          marginTop: setup?.marginsTwips.top ?? DEFAULT_MARGIN,
          marginBottom: setup?.marginsTwips.bottom ?? DEFAULT_MARGIN,
          marginLeft: setup?.marginsTwips.left ?? DEFAULT_MARGIN,
          marginRight: setup?.marginsTwips.right ?? DEFAULT_MARGIN,
          scope: 'document',
        });
        seeded = setup === null ? 'loading' : 'yes';
      });
    });
  }

  /** @internal */
  protected inches(field: PageSetupMarginField): number {
    return twipsToInches(this.values()[field]);
  }

  /** @internal */
  protected setMargin(field: PageSetupMarginField, event: Event): void {
    const value = Number((event.target as HTMLInputElement).value) || 0;
    this.values.update((current) => ({ ...current, [field]: Math.max(0, inchesToTwips(value)) }));
  }

  /** @internal */
  protected setSize(event: Event): void {
    const size = PAGE_SIZES[Number((event.target as HTMLSelectElement).value)];
    if (!size) return;
    this.values.update((current) => {
      const landscape = current.orientation === 'landscape';
      return {
        ...current,
        pageWidth: landscape ? size.height : size.width,
        pageHeight: landscape ? size.width : size.height,
      };
    });
  }

  /** @internal */
  protected selectValue(event: Event): string {
    return (event.target as HTMLSelectElement).value;
  }

  /** @internal */
  protected setOrientation(next: 'portrait' | 'landscape'): void {
    this.values.update((current) =>
      next === current.orientation
        ? current
        : {
            ...current,
            orientation: next,
            pageWidth: current.pageHeight,
            pageHeight: current.pageWidth,
          }
    );
  }

  /** @internal */
  protected setScope(event: Event): void {
    const scope = this.selectValue(event) as 'document' | 'section';
    this.values.update((current) => ({ ...current, scope }));
  }

  /** @internal */
  protected onKeydown(event: KeyboardEvent): void {
    if (event.key !== 'Enter' || event.isComposing) return;
    const target = event.target;
    if (
      isHTMLElementNode(target) &&
      (target.localName === 'button' || target.localName === 'select')
    ) {
      return;
    }
    if (!this.setup.isEnabled()) return;
    event.preventDefault();
    this.applyDraft();
  }

  private applyDraft(): void {
    if (!this.current()) return;
    const draft = untracked(this.values);
    const accepted = this.setup.apply({
      pageWidthTwips: draft.pageWidth,
      pageHeightTwips: draft.pageHeight,
      orientation: draft.orientation,
      marginTopTwips: draft.marginTop,
      marginRightTwips: draft.marginRight,
      marginBottomTwips: draft.marginBottom,
      marginLeftTwips: draft.marginLeft,
      scope: draft.scope,
    });
    this.refused.set(!accepted);
    if (accepted) this.close();
  }
}
