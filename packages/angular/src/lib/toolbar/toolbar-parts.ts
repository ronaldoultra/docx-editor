// The default parts of the toolbar arrangement. A slot without an entry here renders as a
// `docxToolbarButton` with the slot's registry icon.

import type { Type } from '@angular/core';
import type { ChromeSlotId } from '@docx-editor.dev/core/editor';
import {
  DocxToolbarFontColorComponent,
  DocxToolbarHighlightComponent,
} from './toolbar-colors.component';
import { DocxToolbarEditingModeComponent } from './toolbar-editing-mode.component';
import { DocxToolbarReviewersComponent } from './toolbar-reviewers.component';
import { DocxToolbarTableInsertComponent } from './toolbar-table-insert.component';
import {
  DocxToolbarAlignmentComponent,
  DocxToolbarLineSpacingComponent,
} from './toolbar-paragraph.component';
import {
  DocxToolbarFontFamilyComponent,
  DocxToolbarParagraphStyleComponent,
} from './toolbar-pickers.component';
import {
  DocxToolbarFontSizeComponent,
  DocxToolbarZoomComponent,
} from './toolbar-steppers.component';

/** A key of the default arrangement: a slot, or the one alignment dropdown. @internal */
export type ToolbarArrangementKey = ChromeSlotId | 'alignment';

/** The parts that are more than a command button: pickers, steppers, and dropdowns. */
export const SHAPED_PARTS: Partial<Record<ToolbarArrangementKey, Type<unknown>>> = {
  'zoom.level': DocxToolbarZoomComponent,
  'styles.style': DocxToolbarParagraphStyleComponent,
  'font.family': DocxToolbarFontFamilyComponent,
  'font.size': DocxToolbarFontSizeComponent,
  'review.editingMode': DocxToolbarEditingModeComponent,
  'review.authors': DocxToolbarReviewersComponent,
  'text.color': DocxToolbarFontColorComponent,
  'text.highlight': DocxToolbarHighlightComponent,
  'list.lineSpacing': DocxToolbarLineSpacingComponent,
  alignment: DocxToolbarAlignmentComponent,
  'table.insert': DocxToolbarTableInsertComponent,
};

/**
 * The slots whose state is a value, not on or off. In the "⋯" panel, they render as a labeled
 * control row and not as a command row.
 */
export const VALUE_SLOTS: ReadonlySet<ToolbarArrangementKey> = new Set<ToolbarArrangementKey>([
  'alignment',
  'zoom.level',
  'styles.style',
  'font.family',
  'font.size',
  'text.color',
  'text.highlight',
  'text.link',
  'list.lineSpacing',
  'review.editingMode',
  'review.authors',
  'file.save',
  'table.borderTarget',
  'table.borderColor',
  'table.borderStyle',
  'table.borderWidth',
  'table.cellFill',
  'table.splitCell',
  'image.insert',
  'image.wrap',
  'insert.chart',
  'image.altText',
]);

/**
 * Shaped parts that render as a "⋯" panel row themselves. Their popups stay inside the part.
 */
export const PANEL_ROW_SLOTS: ReadonlySet<ToolbarArrangementKey> = new Set<ToolbarArrangementKey>([
  'image.insert',
  'image.wrap',
  'image.altText',
]);

/** The part of a default entry, or `null` for a command button. */
export function defaultPart(slot: ToolbarArrangementKey): Type<unknown> | null {
  return SHAPED_PARTS[slot] ?? null;
}
