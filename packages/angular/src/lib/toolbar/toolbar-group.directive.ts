import {
  Directive,
  TemplateRef,
  booleanAttribute,
  inject,
  input,
  numberAttribute,
} from '@angular/core';
import type { ChromeGroupId } from '@docx-editor.dev/core/editor';

const optionalBoolean = (value: unknown): boolean | undefined =>
  value === undefined || value === null ? undefined : booleanAttribute(value);
const optionalNumber = (value: unknown): number | undefined =>
  value === undefined || value === null ? undefined : numberAttribute(value);

/**
 * A group of the toolbar, as a template inside the toolbar.
 *
 * A host group (any id that is not built in) is measured with the other groups. It collapses
 * in priority order, and it becomes a labeled section of the "⋯" panel when the bar has no
 * room for it. In the panel, `docxToolbarAction` buttons render as panel rows.
 *
 * A built-in id (`'history'`, `'font'`, `'text'`, `'review'`, …) adds the template content
 * to that group, after its own controls, or removes the group with `hidden`.
 *
 * @example
 * ```html
 * <fieldset docxEditorToolbar>
 *   <ng-template docxToolbarGroup="review-nav" label="Review" [priority]="15">
 *     <button docxToolbarAction label="Next change" (select)="next()">…</button>
 *   </ng-template>
 *   <ng-template docxToolbarGroup="zoom" hidden />
 * </fieldset>
 * ```
 *
 * @public
 */
@Directive({ selector: 'ng-template[docxToolbarGroup]' })
export class DocxToolbarGroupDirective {
  /** The group id. A built-in id adds the content to that group. */
  readonly docxToolbarGroup = input.required<ChromeGroupId | (string & {})>();
  /**
   * Catalog key for the host group's panel heading and accessible name. Built-in groups keep
   * their own label.
   */
  readonly labelKey = input<string | undefined>(undefined);
  /** Literal heading and accessible name for a host group. It wins over `labelKey`. */
  readonly label = input<string | undefined>(undefined);
  /**
   * Collapse priority. A lower number leaves the bar first. Built-in groups use 10 (zoom)
   * through 90 (history) in steps of 10. A host group without a priority collapses before
   * every built-in group.
   */
  readonly priority = input<number | undefined, unknown>(undefined, { transform: optionalNumber });
  /** `true` keeps the group in the bar at every width, like the review group. */
  readonly pinned = input<boolean | undefined, unknown>(undefined, { transform: optionalBoolean });
  /**
   * The group that this one follows in the bar. Without it, a host group follows every
   * built-in group, in order of appearance. Ignored for built-in ids.
   */
  readonly after = input<ChromeGroupId | (string & {}) | undefined>(undefined);
  /** Remove the group, built-in or host, from the bar and from the "⋯" panel. */
  readonly hidden = input(false, { transform: booleanAttribute });
  /**
   * The content of the group's "⋯" panel section, after a built-in group's own rows. Without
   * it, the panel renders the template content as panel rows.
   */
  readonly overflowContent = input<TemplateRef<unknown> | undefined>(undefined);
  /** Classes for the group element. */
  readonly className = input<string | undefined>(undefined);

  /** The group's content. */
  readonly template: TemplateRef<unknown> = inject(TemplateRef);
}

/** A host group's heading: the literal label, else the translated key, else the id. */
export function hostGroupText(
  spec: { readonly id: string; readonly label?: string; readonly labelKey?: string },
  translate: (key: string) => string
): string {
  if (spec.label !== undefined) return spec.label;
  return spec.labelKey !== undefined ? translate(spec.labelKey) : spec.id;
}
