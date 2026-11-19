// The toolbar's arrangement as data: the default groups from the engine registry, the host's
// slot templates and group templates applied to them, and what the bar and the "⋯" panel show
// under one overflow answer. The component renders this data and holds no arrangement logic.

import type { TemplateRef, Type } from '@angular/core';
import type { EditorSnapshot } from '@docx-editor.dev/core/contracts/editor';
import {
  TOOLBAR_PINNED_GROUPS,
  arrangeToolbarGroups,
  chromeSlotId,
  formattingBarChromeGroups,
  orderToolbarOverflowSections,
  toolbarGroupItems,
  toolbarHostPanelShown,
  toolbarShownCount,
  type ChromeSlotId,
  type ToolbarItemOverflow,
} from '@docx-editor.dev/core/editor';
import type { DocxSlotDirective } from '../slot.directive';
import { hostGroupText } from './toolbar-group.directive';
import {
  PANEL_ROW_SLOTS,
  VALUE_SLOTS,
  defaultPart,
  type ToolbarArrangementKey,
} from './toolbar-parts';
import {
  TABLE_CHROME_SLOTS,
  TABLE_CONTEXTUAL_GROUP_ID,
  chromeControlForSlot,
  warnBuiltInGroupSetting,
  warnReservedGroupId,
  warnSlotOutsideArrangement,
  warnUnknownGroupAnchor,
} from './toolbar-registry';

/** One control of a default group. @internal */
export interface DefaultEntry {
  readonly slot: ToolbarArrangementKey;
  /** The shaped part, or `null` for a command button. */
  readonly part: Type<unknown> | null;
  /** The slot that the fallback command button drives. */
  readonly button: ChromeSlotId;
}

/** One group of the default arrangement. */
export interface DefaultGroup {
  readonly id: string;
  readonly labelKey: string;
  readonly entries: readonly DefaultEntry[];
}

/** A group template, as the arrangement reads it. */
export interface ToolbarGroupSpec {
  readonly id: string;
  readonly label: string | undefined;
  readonly labelKey: string | undefined;
  readonly priority: number | undefined;
  readonly pinned: boolean | undefined;
  readonly after: string | undefined;
  readonly hidden: boolean;
  readonly overflowContent: TemplateRef<unknown> | undefined;
  readonly className: string | undefined;
  readonly template: TemplateRef<unknown>;
}

/** A group as the bar renders it: built-in or host-owned, in bar order. */
export interface ArrangedGroup {
  readonly id: string;
  /** Panel heading key of a built-in group. */
  readonly labelKey: string;
  /** A host group's heading inputs. */
  readonly heading: Pick<ToolbarGroupSpec, 'id' | 'label' | 'labelKey'> | undefined;
  readonly entries: readonly DefaultEntry[];
  /** Host content: a host group's template, or content added to a built-in group. */
  readonly extras: TemplateRef<unknown> | null;
  readonly pinned: boolean;
  readonly priority: number | undefined;
  readonly overflowContent: TemplateRef<unknown> | undefined;
  readonly className: string | undefined;
}

/** The host's slot templates, by role. @internal */
export interface SlotOverrides {
  /** Overrides of the default arrangement. The last template for a slot wins. */
  readonly overrides: ReadonlyMap<ToolbarArrangementKey, DocxSlotDirective>;
  /** Overrides of the contextual table chrome. */
  readonly tableOverrides: ReadonlyMap<ToolbarArrangementKey, DocxSlotDirective>;
  /** Templates for slots that the arrangement does not draw. They render after the groups. */
  readonly appended: readonly DocxSlotDirective[];
}

/** One control place of a bar group. @internal */
export interface BarItem {
  readonly id: string;
  readonly entry: DefaultEntry | null;
  readonly extras: TemplateRef<unknown> | null;
}

/** One group of the bar, with the controls that it still shows. @internal */
export interface BarGroup {
  readonly id: string;
  readonly className: string;
  /** The accessible name of a host group, or `null` for a built-in group. */
  readonly heading: string | null;
  readonly pinned: boolean;
  readonly items: readonly BarItem[];
}

/** One row of a "⋯" panel section. @internal */
export type PanelRow =
  | { readonly kind: 'command'; readonly key: string; readonly slot: ChromeSlotId }
  | {
      readonly kind: 'control';
      readonly key: string;
      readonly label: string;
      readonly entry: DefaultEntry;
      readonly template: TemplateRef<unknown> | null;
    }
  | { readonly kind: 'row'; readonly key: string; readonly entry: DefaultEntry }
  | {
      readonly kind: 'host';
      readonly key: string;
      readonly template: TemplateRef<unknown>;
      /** True for a group's `overflowContent`, which renders as is. */
      readonly own: boolean;
    };

/** One section of the "⋯" panel. @internal */
export interface PanelSection {
  readonly id: string;
  readonly label: string;
  readonly rows: readonly PanelRow[];
}

const entryOf = (slot: ToolbarArrangementKey): DefaultEntry => ({
  slot,
  part: defaultPart(slot),
  button: slot === 'alignment' ? 'alignment.left' : slot,
});

/** The default groups for the current selection, in bar order. */
export function buildDefaultGroups(
  image: EditorSnapshot['image'],
  direction: boolean,
  panes: { readonly chart: boolean; readonly shape: boolean },
  chart: boolean
): readonly DefaultGroup[] {
  // Each pane button shows only when the host can show its pane.
  return formattingBarChromeGroups(image, {
    direction,
    chart: panes.chart && chart,
    shape: panes.shape && !!image?.shapeFormat,
    shapeControls: !!image?.shapeFormat,
    chartControls: chart,
  }).map((group) => ({
    id: group.id,
    labelKey: group.labelKey,
    entries:
      group.id === 'alignment'
        ? [entryOf('alignment')]
        : group.controls.map((control) => entryOf(chromeSlotId(group, control))),
  }));
}

/** Sort the slot templates into overrides, table overrides, and appended templates. */
export function readSlotOverrides(
  slots: readonly DocxSlotDirective[],
  defaults: readonly DefaultGroup[]
): SlotOverrides {
  const drawn = new Set<string>(defaults.flatMap((group) => group.entries.map((e) => e.slot)));
  const overrides = new Map<ToolbarArrangementKey, DocxSlotDirective>();
  const tableOverrides = new Map<ToolbarArrangementKey, DocxSlotDirective>();
  const appended: DocxSlotDirective[] = [];
  for (const slot of slots) {
    const id = slot.docxSlot();
    if (drawn.has(id)) overrides.set(id, slot);
    else if ((TABLE_CHROME_SLOTS as readonly string[]).includes(id)) tableOverrides.set(id, slot);
    // Hiding a slot that the bar does not draw leaves nothing.
    else if (slot.docxSlotHidden()) continue;
    else {
      warnSlotOutsideArrangement(id, chromeControlForSlot(id) !== null);
      appended.push(slot);
    }
  }
  return { overrides, tableOverrides, appended };
}

const isHidden = (override: DocxSlotDirective | undefined): boolean =>
  override?.docxSlotHidden() === true;

/**
 * The bar's groups: the built-in ones with their group templates applied, and the host's own
 * groups placed by `arrangeToolbarGroups`. A group with nothing left to render is dropped, so
 * it leaves no empty box and no separator in the bar or the panel.
 */
export function arrangeGroups(
  defaults: readonly DefaultGroup[],
  specs: readonly ToolbarGroupSpec[],
  overrides: ReadonlyMap<ToolbarArrangementKey, DocxSlotDirective>
): readonly ArrangedGroup[] {
  // The last template for an id wins, and its position is its order of appearance.
  const byIdSpec = new Map<string, ToolbarGroupSpec>();
  for (const spec of specs) {
    byIdSpec.delete(spec.id);
    byIdSpec.set(spec.id, spec);
  }
  const unique = [...byIdSpec.values()];
  const builtInIds = new Set(defaults.map((group) => group.id));
  const byId = new Map<string, ArrangedGroup>();
  for (const group of defaults) {
    const spec = byIdSpec.get(group.id);
    if (spec) {
      for (const setting of ['label', 'labelKey', 'after'] as const) {
        if (spec[setting] !== undefined) warnBuiltInGroupSetting(spec.id, setting);
      }
    }
    if (spec?.hidden) continue;
    const extras = spec?.template ?? null;
    const visible = group.entries.some((entry) => !isHidden(overrides.get(entry.slot)));
    if (!visible && extras === null) continue;
    byId.set(group.id, {
      id: group.id,
      labelKey: group.labelKey,
      heading: undefined,
      entries: group.entries,
      extras,
      pinned: spec?.pinned ?? TOOLBAR_PINNED_GROUPS.has(group.id),
      priority: spec?.priority,
      overflowContent: spec?.overflowContent,
      className: spec?.className,
    });
  }
  for (const spec of unique) {
    if (spec.id === TABLE_CONTEXTUAL_GROUP_ID) warnReservedGroupId(spec.id);
  }
  const hosts = unique.filter(
    (spec) => !builtInIds.has(spec.id) && spec.id !== TABLE_CONTEXTUAL_GROUP_ID
  );
  for (const spec of hosts) {
    if (spec.hidden) continue;
    byId.set(spec.id, {
      id: spec.id,
      labelKey: spec.labelKey ?? spec.id,
      heading: { id: spec.id, label: spec.label, labelKey: spec.labelKey },
      entries: [],
      extras: spec.template,
      pinned: spec.pinned === true,
      priority: spec.priority,
      overflowContent: spec.overflowContent,
      className: spec.className,
    });
  }
  const known = [...builtInIds, ...hosts.map((spec) => spec.id)];
  for (const spec of hosts) {
    if (spec.after !== undefined && !known.includes(spec.after)) {
      warnUnknownGroupAnchor(spec.id, spec.after, known);
    }
  }
  // Placed over EVERY built-in id, hidden ones included, so `after` still finds its anchor.
  const order = arrangeToolbarGroups(
    defaults.map((group) => group.id),
    hosts.map((spec) => ({ id: spec.id, after: spec.after }))
  );
  return order.flatMap((id) => {
    const group = byId.get(id);
    return group ? [group] : [];
  });
}

/** A group's collapsible controls, in bar order. A hidden slot takes no place. */
export function groupItems(
  group: ArrangedGroup,
  overrides: ReadonlyMap<ToolbarArrangementKey, DocxSlotDirective>
): readonly string[] {
  const slots = group.entries
    .filter((entry) => !isHidden(overrides.get(entry.slot)))
    .map((entry) => entry.slot);
  return toolbarGroupItems(slots, group.extras !== null);
}

/** The label of a panel control row: the slot's own label, else the group's. */
function labelOf(label: (key: string) => string, entry: DefaultEntry, groupKey: string): string {
  const control = entry.slot === 'alignment' ? null : chromeControlForSlot(entry.slot);
  return label(control?.labelKey ?? groupKey);
}

/** The bar and the "⋯" panel under one overflow answer. */
export function layoutToolbar(
  arranged: readonly ArrangedGroup[],
  overrides: ReadonlyMap<ToolbarArrangementKey, DocxSlotDirective>,
  overflow: ToolbarItemOverflow,
  label: (key: string) => string
): { readonly bar: readonly BarGroup[]; readonly sections: readonly PanelSection[] } {
  const bar: BarGroup[] = [];
  const sections: PanelSection[] = [];
  for (const group of arranged) {
    // When the bar runs out of width, controls leave from the END of a group, one at a time,
    // so the bar keeps a prefix and "⋯" holds the rest.
    const items = groupItems(group, overrides);
    const entryBySlot = new Map<string, DefaultEntry>(group.entries.map((e) => [e.slot, e]));
    const shown = group.pinned ? items.length : toolbarShownCount(overflow, group.id, items.length);
    const hidden = items.slice(shown);
    const heading = group.heading ? hostGroupText(group.heading, label) : null;
    if (hidden.length > 0) {
      const rows: PanelRow[] = hidden.flatMap((item): PanelRow[] => {
        const entry = entryBySlot.get(item);
        if (!entry) return [];
        const override = overrides.get(entry.slot);
        if (!override && PANEL_ROW_SLOTS.has(entry.slot)) {
          return [{ kind: 'row', key: entry.slot, entry }];
        }
        if (override || VALUE_SLOTS.has(entry.slot)) {
          return [
            {
              kind: 'control',
              key: entry.slot,
              label: labelOf(label, entry, group.labelKey),
              entry,
              template: override ? (override.docxSlotOverflow() ?? override.template) : null,
            },
          ];
        }
        return [{ kind: 'command', key: entry.slot, slot: entry.slot as ChromeSlotId }];
      });
      if (toolbarHostPanelShown(group.extras !== null, hidden)) {
        // `overflowContent` renders even without bar content, so a built-in group can add a
        // panel row of its own.
        if (group.overflowContent) {
          rows.push({ kind: 'host', key: 'host', template: group.overflowContent, own: true });
        } else if (group.extras) {
          rows.push({ kind: 'host', key: 'host', template: group.extras, own: false });
        }
      }
      if (rows.length > 0) {
        sections.push({ id: group.id, label: heading ?? label(group.labelKey), rows });
      }
    }
    if (shown === 0) continue;
    bar.push({
      id: group.id,
      className: `docx-toolbar__group${group.className ? ` ${group.className}` : ''}`,
      heading,
      pinned: group.pinned,
      items: items.slice(0, shown).map((item) => {
        const entry = entryBySlot.get(item) ?? null;
        return { id: item, entry, extras: entry ? null : group.extras };
      }),
    });
  }
  // The selected object's own groups lead the panel.
  return { bar, sections: orderToolbarOverflowSections(sections) };
}
