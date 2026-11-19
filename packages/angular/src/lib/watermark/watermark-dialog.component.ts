import { NgTemplateOutlet } from '@angular/common';
import {
  ChangeDetectionStrategy,
  Component,
  ElementRef,
  Injector,
  afterNextRender,
  computed,
  contentChildren,
  effect,
  inject,
  input,
  signal,
  untracked,
  viewChildren,
  type Signal,
} from '@angular/core';
import type { EditorSnapshot } from '@docx-editor.dev/core/contracts/editor';
import { isHTMLElementNode } from '@docx-editor.dev/core/editor';
import { injectDocxEditor } from '../context';
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
import { injectEditorState } from '../editor-state';
import { DocxPartSlotDirective } from '../parts/part-slot';
import { injectWatermark, type WatermarkCommandState } from './watermark';
import {
  DEFAULT_WATERMARK_FIELDS,
  DEFAULT_WATERMARK_PRESETS,
  WATERMARK_IMAGE_ACCEPT,
  WATERMARK_SCALES,
  WATERMARK_SIZES,
  colorFromInput,
  colorInputValue,
  seedWatermarkFields,
  transparencyFromPercent,
  transparencyPercent,
  watermarkBaseline,
  watermarkFieldsEqual,
  watermarkFromFields,
  watermarkPresetValue,
  type WatermarkDialogFields,
} from './watermark-dialog-model';
import { DocxEditorWatermarkPreviewComponent } from './watermark-preview.component';

/**
 * The named controls of the watermark dialog. A `docxDialogPart` template replaces one as
 * `field:<name>`.
 *
 * @public
 */
export type WatermarkDialogFieldName =
  | 'noWatermark'
  | 'pictureWatermark'
  | 'picture'
  | 'scale'
  | 'washout'
  | 'textWatermark'
  | 'preset'
  | 'text'
  | 'fontFamily'
  | 'fontSizePt'
  | 'bold'
  | 'italic'
  | 'transparency'
  | 'color'
  | 'layout';

/**
 * The draft state and the actions of the enclosing watermark dialog. Call it in an injection
 * context inside the dialog, such as in a component in a part template.
 *
 * @public
 */
export function injectWatermarkDialog(): DocxDialogState<WatermarkDialogFields> {
  return injectDialog<WatermarkDialogFields>();
}

const selectDocumentAbsent = (snapshot: EditorSnapshot): boolean =>
  snapshot.isLoading || snapshot.parseError !== null || snapshot.pageSetup == null;

let nextId = 0;

/**
 * The watermark dialog: no watermark, a picture, or text, with a live page preview. Apply sets
 * the whole form as one undo step, and an unchanged form writes nothing. The editor opens it from
 * Insert › Watermark.
 *
 * Replace or restyle a part with `docxDialogPart` templates. The part names are `header`,
 * `title`, `body`, `preview`, `footer`, `error`, `cancel`, `apply`, and `field:<name>` for each
 * {@link WatermarkDialogFieldName}. In your own layout, `docxDialogOutlet` renders one of them.
 *
 * @public
 */
@Component({
  selector: 'docx-editor-watermark-dialog',
  imports: [
    NgTemplateOutlet,
    DocxNativeDialogComponent,
    DocxPartSlotDirective,
    DocxDialogDefaultDirective,
    DocxDialogOutletDirective,
    DocxDialogApplyDirective,
    DocxDialogCancelDirective,
    DocxDialogFieldDirective,
    DocxEditorWatermarkPreviewComponent,
  ],
  providers: provideDialogParts(),
  changeDetection: ChangeDetectionStrategy.OnPush,
  host: { style: 'display: contents' },
  templateUrl: './watermark-dialog.component.html',
})
export class DocxEditorWatermarkDialogComponent extends DocxEditorDialogBase {
  /** The text choices for a text watermark. Defaults to `DEFAULT_WATERMARK_PRESETS`. */
  readonly presets = input<readonly string[] | undefined>(undefined);

  /** @internal */
  protected readonly id: string = `docx-watermark-${++nextId}`;
  /** @internal */
  protected readonly accept: string = WATERMARK_IMAGE_ACCEPT;
  /** @internal */
  protected readonly scales: readonly (number | null)[] = WATERMARK_SCALES;
  /** @internal */
  protected readonly sizes: readonly (number | null)[] = WATERMARK_SIZES;
  /** @internal */
  protected readonly kinds: readonly {
    field: WatermarkDialogFieldName;
    kind: WatermarkDialogFields['kind'];
    key: string;
  }[] = [
    { field: 'noWatermark', kind: 'none', key: 'noWatermark' },
    { field: 'pictureWatermark', kind: 'picture', key: 'picture' },
    { field: 'textWatermark', kind: 'text', key: 'text' },
  ];
  /** @internal */
  protected readonly boxes: readonly {
    field: 'washout' | 'bold' | 'italic';
    kind: WatermarkDialogFields['kind'];
  }[] = [
    { field: 'washout', kind: 'picture' },
    { field: 'bold', kind: 'text' },
    { field: 'italic', kind: 'text' },
  ];
  /** @internal */
  protected readonly layouts: readonly ('diagonal' | 'horizontal')[] = ['diagonal', 'horizontal'];
  /** @internal */
  protected readonly watermark: WatermarkCommandState = injectWatermark();
  /** @internal */
  protected readonly values = signal<WatermarkDialogFields>({ ...DEFAULT_WATERMARK_FIELDS });
  /** @internal */
  protected readonly fonts = signal<readonly string[]>([]);
  /** @internal */
  protected readonly refusal = signal<string | null>(null);
  /** @internal */
  protected readonly pictureFailed = signal(false);
  /** @internal */
  protected readonly presetChoices: Signal<readonly string[]> = computed(
    () => this.presets() ?? DEFAULT_WATERMARK_PRESETS
  );
  /** @internal */
  protected readonly fontChoices: Signal<readonly string[]> = computed(() => {
    const fonts = this.fonts();
    const family = this.values().fontFamily;
    return fonts.includes(family) ? fonts : [family, ...fonts];
  });
  /** @internal */
  protected readonly canApply: Signal<boolean> = computed(
    () => this.watermark.isEnabled() && watermarkFromFields(this.values()) !== undefined
  );
  /** @internal */
  protected readonly errorText: Signal<string | null> = computed(() =>
    this.refusal() !== null ? this.label('dialogs.watermark.refused') : null
  );
  /** @internal */
  protected readonly state: DocxDialogState<WatermarkDialogFields> = {
    values: this.values.asReadonly(),
    errors: computed(() => {
      const error = this.errorText();
      return error ? { form: error } : {};
    }),
    isEnabled: this.canApply,
    setValue: (name, value) => this.setValue(name, value),
    apply: () => this.applyDraft(),
    cancel: () => this.close(),
  };

  private baseline: WatermarkDialogFields | null = null;
  private readonly editor = injectDocxEditor();
  private readonly injector = inject(Injector);
  private readonly host: HTMLElement = inject(ElementRef).nativeElement;
  private readonly absent = injectEditorState(selectDocumentAbsent);
  private readonly parts = contentChildren(DocxDialogPartDirective);
  private readonly defaults = viewChildren(DocxDialogDefaultDirective);
  private readonly current = injectDialogGeneration(this.open, () => this.close());

  constructor() {
    super();
    const parts = inject(DocxDialogParts, { self: true });
    parts.connect(this.state as unknown as DocxDialogState<object>, this.parts);
    parts.connectDefaults(this.defaults);
    // Seed when the dialog opens on a loaded document, never on later ticks, or a concurrent
    // edit would fight the user's choices. An unload while open seeds from the next document.
    let seeded = false;
    effect(() => {
      const open = this.open();
      const absent = this.absent();
      const editor = this.editor();
      untracked(() => {
        if (!open || absent || !editor) {
          seeded = false;
          return;
        }
        if (seeded) return;
        seeded = true;
        const state = editor.getWatermark();
        const seed = seedWatermarkFields(state, this.watermark.read());
        this.baseline = watermarkBaseline(state, seed);
        this.values.set(seed);
        this.fonts.set(editor.getAvailableFonts());
        this.refusal.set(null);
        this.pictureFailed.set(false);
        // The focus follows the seed: the checked kind, not the first choice. The read phase
        // runs after the dialog's own first focus.
        afterNextRender(
          {
            read: () => {
              this.host
                .querySelector<HTMLInputElement>(`input[name="${this.id}-kind"]:checked`)
                ?.focus({ preventScroll: true });
            },
          },
          { injector: this.injector }
        );
      });
    });
  }

  /** @internal */
  protected percent(transparency: number): number {
    return transparencyPercent(transparency);
  }

  /** @internal */
  protected percentLabel(factor: number): string {
    return `${Math.round(factor * 100)}%`;
  }

  /** @internal */
  protected presetValue(): string {
    return watermarkPresetValue(this.values().text, this.presetChoices());
  }

  /** @internal */
  protected colorValue(): string {
    return colorInputValue(this.values().color);
  }

  /** @internal */
  protected target(event: Event): HTMLInputElement {
    return event.target as HTMLInputElement;
  }

  /** @internal */
  protected setValue<K extends keyof WatermarkDialogFields>(
    name: K,
    value: WatermarkDialogFields[K]
  ): void {
    this.refusal.set(null);
    this.values.update((current) => ({ ...current, [name]: value }));
  }

  /** @internal */
  protected setChoice(name: 'scale' | 'fontSizePt', event: Event): void {
    const value = this.target(event).value;
    this.setValue(name, value === 'auto' ? null : Number(value));
  }

  /** @internal */
  protected setPreset(event: Event): void {
    const value = this.target(event).value;
    if (value !== '') this.setValue('text', value);
  }

  /** @internal */
  protected setTransparency(event: Event): void {
    const previous = untracked(this.values).transparency;
    this.setValue('transparency', transparencyFromPercent(this.target(event).value, previous));
  }

  /** @internal */
  protected setColor(event: Event): void {
    const color = colorFromInput(this.target(event).value);
    if (color) this.setValue('color', color);
  }

  /** @internal */
  protected selectPicture(input: HTMLInputElement): void {
    this.setValue('kind', 'picture');
    input.click();
  }

  /** @internal */
  protected onFile(event: Event): void {
    const target = this.target(event);
    const file = target.files?.[0];
    // Cleared, so the same file chosen again fires a change event again.
    target.value = '';
    if (!file) return;
    this.pictureFailed.set(false);
    void file
      .arrayBuffer()
      .then((buffer) => {
        this.values.update((current) => ({
          ...current,
          kind: 'picture',
          imageData: new Uint8Array(buffer),
          imageName: file.name,
        }));
        this.refusal.set(null);
      })
      .catch(() => this.pictureFailed.set(true));
  }

  /** @internal */
  protected onKeydown(event: KeyboardEvent): void {
    const target = event.target;
    if (
      event.key === 'Enter' &&
      !event.isComposing &&
      isHTMLElementNode(target) &&
      target.localName === 'input' &&
      ((target as HTMLInputElement).type === 'text' ||
        (target as HTMLInputElement).type === 'number') &&
      this.canApply()
    ) {
      event.preventDefault();
      this.applyDraft();
    }
  }

  private applyDraft(): void {
    const draft = untracked(this.values);
    const next = watermarkFromFields(draft);
    if (!untracked(this.canApply) || next === undefined) return;
    if (!this.current()) return;
    // An unchanged draft writes nothing, so it leaves no empty undo step.
    if (this.baseline && watermarkFieldsEqual(this.baseline, draft)) {
      this.close();
      return;
    }
    // A refused write keeps the dialog open: closing would claim a success.
    const result = this.watermark.apply(next);
    if (result.ok) this.close();
    else this.refusal.set(result.reason);
  }
}
