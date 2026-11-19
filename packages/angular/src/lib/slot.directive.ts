import { Directive, TemplateRef, booleanAttribute, inject, input } from '@angular/core';
import type { ChromeSlotId } from '@docx-editor.dev/core/editor';

/**
 * Replace one control of a chrome part's default set. Put the template inside the part, such as
 * the toolbar. The part renders the template in the place of the default control, and keeps
 * every other default control.
 *
 * With `docxSlotHidden`, the part removes the control. With `docxSlotOverflow`, the toolbar
 * renders that template in the "⋯" panel when the control leaves the bar.
 *
 * @example
 * ```html
 * <fieldset docxEditorToolbar>
 *   <ng-template docxSlot="text.bold">
 *     <button docxToolbarButton="text.bold">B</button>
 *   </ng-template>
 *   <ng-template docxSlot="text.strike" docxSlotHidden />
 * </fieldset>
 * ```
 *
 * @public
 */
@Directive({ selector: 'ng-template[docxSlot]' })
export class DocxSlotDirective {
  /** The chrome slot that the template replaces, such as `'text.bold'`. */
  readonly docxSlot = input.required<ChromeSlotId>();
  /** Remove the slot from the default set. */
  readonly docxSlotHidden = input(false, { transform: booleanAttribute });
  /** The content of the slot's "⋯" panel row. Without it, the row shows the template. */
  readonly docxSlotOverflow = input<TemplateRef<unknown> | undefined>(undefined);

  /** The replacement content. */
  readonly template: TemplateRef<unknown> = inject(TemplateRef);
}
