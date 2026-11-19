import { computed, untracked, type Signal } from '@angular/core';
import type { EditorSnapshot, RunFormatting } from '@docx-editor.dev/core/contracts/editor';
import type { ParagraphFormatRead, ParagraphFormatUpdate } from '@docx-editor.dev/core/editor';
import { injectDocxEditor } from '../context';
import { injectEditorCommand } from '../editor-command';
import { injectEditorState } from '../editor-state';

/**
 * What {@link injectParagraphFormat} returns.
 *
 * @public
 */
export interface ParagraphFormatState {
  /** The paragraph formatting of the selection, or `null` without a selection. */
  readonly format: Signal<ParagraphFormatRead | null>;
  /** True when the engine can change the paragraph formatting of the selection now. */
  readonly isEnabled: Signal<boolean>;
  /** Change the paragraph formatting as one undo step. Returns `false` on a refusal. */
  apply(update: ParagraphFormatUpdate): boolean;
}

const EMPTY_FLAGS = {
  contextualSpacing: null,
  keepNext: null,
  keepLines: null,
  widowControl: null,
  pageBreakBefore: null,
} as const;

function selectFormat(snapshot: EditorSnapshot): ParagraphFormatRead | null {
  const formatting: RunFormatting | null = snapshot.formatting ?? null;
  if (!formatting) return null;
  const indent = formatting.indent;
  const flags = formatting.paragraphFlags ?? EMPTY_FLAGS;
  return {
    alignment: formatting.alignment === 'both' ? 'justify' : (formatting.alignment ?? null),
    direction: formatting.direction ?? null,
    spaceBeforePt: formatting.spaceBeforePt ?? null,
    spaceAfterPt: formatting.spaceAfterPt ?? null,
    lineSpacing: formatting.lineSpacing ?? null,
    // A mixed side reads `null`, so a control that can show "mixed" does.
    indentLeftTwips: indent && !indent.mixed.left ? indent.left : null,
    indentRightTwips: indent && !indent.mixed.right ? indent.right : null,
    indentFirstLineTwips: indent && !indent.mixed.firstLine ? indent.firstLine : null,
    contextualSpacing: flags.contextualSpacing,
    keepNext: flags.keepNext,
    keepLines: flags.keepLines,
    widowControl: flags.widowControl,
    pageBreakBefore: flags.pageBreakBefore,
    tabStops: formatting.tabStops ?? null,
    indentUnknown: !indent,
    disagrees: {
      alignment: formatting.disagrees?.alignment ?? false,
      direction: formatting.disagrees?.direction ?? false,
      spaceBeforePt: formatting.disagrees?.spaceBeforePt ?? false,
      spaceAfterPt: formatting.disagrees?.spaceAfterPt ?? false,
      lineSpacing: formatting.disagrees?.lineSpacing ?? false,
      tabStops: formatting.disagrees?.tabStops ?? false,
      indentLeft: indent?.mixed.left ?? false,
      indentRight: indent?.mixed.right ?? false,
      indentFirstLine: indent?.mixed.firstLine ?? false,
    },
  };
}

// Typed against `disagrees`, so a new member fails to compile until it is compared here.
const DISAGREEMENT_KEYS: readonly (keyof ParagraphFormatRead['disagrees'])[] = [
  'alignment',
  'direction',
  'spaceBeforePt',
  'spaceAfterPt',
  'lineSpacing',
  'tabStops',
  'indentLeft',
  'indentRight',
  'indentFirstLine',
];

function sameFormat(a: ParagraphFormatRead | null, b: ParagraphFormatRead | null): boolean {
  if (a === null || b === null) return a === b;
  return (
    a.alignment === b.alignment &&
    a.direction === b.direction &&
    a.spaceBeforePt === b.spaceBeforePt &&
    a.spaceAfterPt === b.spaceAfterPt &&
    a.lineSpacing?.rule === b.lineSpacing?.rule &&
    a.lineSpacing?.value === b.lineSpacing?.value &&
    a.indentLeftTwips === b.indentLeftTwips &&
    a.indentRightTwips === b.indentRightTwips &&
    a.indentFirstLineTwips === b.indentFirstLineTwips &&
    a.contextualSpacing === b.contextualSpacing &&
    a.keepNext === b.keepNext &&
    a.keepLines === b.keepLines &&
    a.widowControl === b.widowControl &&
    a.pageBreakBefore === b.pageBreakBefore &&
    // Equal values with other disagreements are another selection: a stale slice would show a
    // mixed selection as settled.
    DISAGREEMENT_KEYS.every((key) => a.disagrees[key] === b.disagrees[key]) &&
    a.indentUnknown === b.indentUnknown &&
    a.tabStops?.length === b.tabStops?.length &&
    (a.tabStops ?? []).every(
      (stop, index) =>
        stop.positionTwips === b.tabStops?.[index]?.positionTwips &&
        stop.alignment === b.tabStops?.[index]?.alignment &&
        (stop.leader ?? 'none') === (b.tabStops?.[index]?.leader ?? 'none')
    )
  );
}

const selectEditable = (snapshot: EditorSnapshot): boolean => snapshot.editable;
const SET_FORMAT_PROBE = { type: 'setParagraphFormat', alignment: 'left' } as const;

/**
 * The paragraph formatting of the selection in the enclosing editor root, and the command that
 * changes it. The enabled state follows the selection: the document open for viewing refuses
 * it, and a content lock or a data binding does not. Call it in an injection context.
 *
 * @public
 */
export function injectParagraphFormat(): ParagraphFormatState {
  const editor = injectDocxEditor();
  const format = injectEditorState(selectFormat, sameFormat);
  const editable = injectEditorState(selectEditable);
  const command = injectEditorCommand(SET_FORMAT_PROBE);
  return {
    format,
    isEnabled: computed(() => editable() && command.isEnabled()),
    apply: (update) => {
      const instance = untracked(editor);
      if (!instance) return false;
      const next = { type: 'setParagraphFormat' as const, ...update };
      if (!instance.can(next).ok) return false;
      return instance.exec(next).ok;
    },
  };
}
