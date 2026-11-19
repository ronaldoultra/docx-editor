// What every toolbar part reads from the engine's chrome registry, and the development warnings
// for toolbar compositions that cannot do what they say.

import { isDevMode } from '@angular/core';
import {
  CHROME_GROUPS,
  TOOLBAR_ITEM_ATTRIBUTE,
  chromeSlotId,
  isElementNode,
  type ChromeControl,
  type ChromeSlotId,
} from '@docx-editor.dev/core/editor';

/** Marks a collapsible group of the bar. The value is the group id. */
export const GROUP_ATTRIBUTE = 'data-toolbar-group';
/** Marks one control of a collapsible group. The value is the control id. */
export const ITEM_ATTRIBUTE = TOOLBAR_ITEM_ATTRIBUTE;
/** Marks a block of the bar that never collapses. */
export const FIXED_ATTRIBUTE = 'data-toolbar-fixed';
/** Marks the "⋯" trigger's root. */
export const MORE_ATTRIBUTE = 'data-toolbar-more';

/** The id of the contextual table group. A host group cannot use it. */
export const TABLE_CONTEXTUAL_GROUP_ID = 'contextual-table';

/** The slots of the contextual table group, in bar order. */
export const TABLE_CHROME_SLOTS: readonly ChromeSlotId[] = [
  'table.borderTarget',
  'table.borderColor',
  'table.borderStyle',
  'table.borderWidth',
  'table.cellFill',
  'table.mergeCells',
  'table.splitCell',
];

/** The "⋯" icon (Material Symbols `more_horiz`). */
export const MORE_PATHS: readonly string[] = [
  'M240-400q-33 0-56.5-23.5T160-480q0-33 23.5-56.5T240-560q33 0 56.5 23.5T320-480q0 33-23.5 56.5T240-400Zm240 0q-33 0-56.5-23.5T400-480q0-33 23.5-56.5T480-560q33 0 56.5 23.5T560-480q0 33-23.5 56.5T480-400Zm240 0q-33 0-56.5-23.5T640-480q0-33 23.5-56.5T720-560q33 0 56.5 23.5T800-480q0 33-23.5 56.5T720-400Z',
];

const controlBySlot = new Map<string, ChromeControl>();
for (const group of CHROME_GROUPS) {
  for (const control of group.controls) {
    const slot = chromeSlotId(group, control);
    if (!controlBySlot.has(slot)) controlBySlot.set(slot, control);
  }
}

/** The registry control behind one slot id, for its label key and icon paths. */
export function chromeControlForSlot(slotId: string): ChromeControl | null {
  return controlBySlot.get(slotId) ?? null;
}

/**
 * Keep the caret: a mousedown on a toolbar control must not move the focus out of the pages.
 * Text fields keep their own mousedown, so that a click puts the caret in them.
 */
export function guardToolbarMousedown(event: Event): void {
  const target = event.target;
  const tag = isElementNode(target) ? target.tagName : '';
  if (tag === 'INPUT' || tag === 'SELECT' || tag === 'TEXTAREA') return;
  event.preventDefault();
}

const printed = new Set<string>();

/** Print `message` once per page, in development mode only. */
export function toolbarDevWarning(message: string): void {
  if (!isDevMode() || printed.has(message)) return;
  printed.add(message);
  console.warn(`[docx-editor] ${message}`);
}

/** A group template names an `after` anchor that is not a group of this toolbar. */
export function warnUnknownGroupAnchor(id: string, after: string, known: readonly string[]): void {
  toolbarDevWarning(
    `docxToolbarGroup "${id}" has after="${after}", but no group has that id. ` +
      `The group is placed after every other group. Known ids: ${known.join(', ')}.`
  );
}

/** A group template with a built-in id sets an input that only a host group uses. */
export function warnBuiltInGroupSetting(id: string, setting: string): void {
  toolbarDevWarning(
    `docxToolbarGroup "${id}" is a built-in group, so its ${setting} is ignored. ` +
      `A built-in group keeps its own label and place in the bar.`
  );
}

/** A group template uses the id that the toolbar keeps for its contextual table group. */
export function warnReservedGroupId(id: string): void {
  toolbarDevWarning(
    `docxToolbarGroup "${id}" uses the id of the contextual table group, so it is ignored. ` +
      `Give the group another id.`
  );
}

/** A slot template names a slot that the default arrangement does not draw. */
export function warnSlotOutsideArrangement(slot: string, known: boolean): void {
  toolbarDevWarning(
    known
      ? `Toolbar slot "${slot}" is not part of the default arrangement, so its content is ` +
          `appended after the groups. Put it in a docxToolbarGroup template to give it a place.`
      : `"${slot}" is not a toolbar slot id. Its content is appended after the groups. ` +
          `Use a ChromeSlotId such as "text.bold".`
  );
}
