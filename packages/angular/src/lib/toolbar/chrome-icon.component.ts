import { ChangeDetectionStrategy, Component, input } from '@angular/core';

/**
 * A registry icon as inline SVG: Material Symbols paths in the `0 -960 960 960` view box. The
 * icon is decoration, so it is hidden from assistive technology.
 *
 * @example
 * ```html
 * <svg [docxChromeIcon]="paths"></svg>
 * ```
 *
 * @public
 */
@Component({
  selector: 'svg[docxChromeIcon]',
  changeDetection: ChangeDetectionStrategy.OnPush,
  host: {
    viewBox: '0 -960 960 960',
    width: '18',
    height: '18',
    'aria-hidden': 'true',
    focusable: 'false',
  },
  template: `@for (d of docxChromeIcon(); track $index) {
    <svg:path [attr.d]="d" fill="currentColor" />
  }`,
})
export class DocxChromeIconComponent {
  /** The SVG path data, one entry for each path. */
  readonly docxChromeIcon = input.required<readonly string[]>();
}
