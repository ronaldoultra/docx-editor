// The state of the toolbar pickers, as functions that a host can call to build its own picker.

import { NgZone, computed, inject, untracked, type Signal } from '@angular/core';
import type { EditorSnapshot, ZoomMode } from '@docx-editor.dev/core/contracts/editor';
import { FIT_WIDTH_ZOOM_MODE, commandForSlotValue } from '@docx-editor.dev/core/editor';
import type { TFunction, TranslationKey } from '@docx-editor.dev/core/i18n';
import { injectDocxEditor } from '../context';
import { injectEditorCommand } from '../editor-command';
import { injectEditorState } from '../editor-state';
import { injectTranslation } from '../i18n';
import { ZOOM_LEVELS, stepZoomLevel } from '../zoom-levels';

const EMPTY_FONTS: readonly string[] = Object.freeze([]);
const sameList = <T>(a: readonly T[], b: readonly T[], same: (x: T, y: T) => boolean): boolean =>
  a.length === b.length && a.every((item, index) => same(item, b[index]!));
const sameStyle = (a: ParagraphStyleOption, b: ParagraphStyleOption): boolean =>
  a.styleId === b.styleId &&
  a.name === b.name &&
  a.preview.fontFamily === b.preview.fontFamily &&
  a.preview.fontSizePt === b.preview.fontSizePt &&
  a.preview.bold === b.preview.bold &&
  a.preview.italic === b.preview.italic &&
  a.preview.color === b.preview.color;

/** What {@link injectFontFamily} returns. @public */
export interface FontFamilyState {
  /** The font family at the selection, or `null` when the selection has more than one. */
  readonly value: Signal<string | null>;
  /** Apply a font family to the selection. */
  readonly setValue: (family: string) => void;
  /** The font families that the engine offers. */
  readonly options: Signal<readonly string[]>;
  /** True when the selection can take a font family. */
  readonly isEnabled: Signal<boolean>;
}

/**
 * The font family of the selection, the families that the engine offers, and a function that
 * applies one. Call it in an injection context.
 *
 * @public
 */
export function injectFontFamily(): FontFamilyState {
  const editor = injectDocxEditor();
  const zone = inject(NgZone);
  const value = injectEditorState((snapshot) => snapshot.formatting?.fontFamily ?? null);
  const command = injectEditorCommand('font.family');
  // Read on each engine tick: a newly loaded document can leave the snapshot as it was.
  const options = injectEditorState(
    (snapshot) => {
      const instance = editor();
      return instance && !snapshot.isLoading
        ? untracked(() => instance.getAvailableFonts())
        : EMPTY_FONTS;
    },
    (a, b) => sameList(a, b, Object.is)
  );
  return {
    value,
    options,
    setValue: (family) => applySlotValue(zone, untracked(editor), 'font.family', family),
    isEnabled: command.isEnabled,
  };
}

/** One paragraph style of the document, as the style picker shows it. @public */
export interface ParagraphStyleOption {
  readonly styleId: string;
  /** The label: the catalog label of a built-in style, else the style's own name. */
  readonly name: string;
  readonly preview: {
    readonly fontFamily: string | null;
    readonly fontSizePt: number | null;
    readonly bold: boolean;
    readonly italic: boolean;
    readonly color: string | null;
  };
}

/** What {@link injectParagraphStyle} returns. @public */
export interface ParagraphStyleState {
  /** The style id at the selection, or `null` when the selection has more than one. */
  readonly value: Signal<string | null>;
  /** Apply a paragraph style to the selected paragraphs. */
  readonly setValue: (styleId: string) => void;
  /** The paragraph styles of the document, in the engine's gallery order. */
  readonly options: Signal<readonly ParagraphStyleOption[]>;
  /** True when the selection can take a paragraph style. */
  readonly isEnabled: Signal<boolean>;
}

const EMPTY_STYLES: readonly ParagraphStyleOption[] = Object.freeze([]);
const BUILT_IN_HEADING_NAME = /^heading ([1-9])$/;
const BUILT_IN_STYLE_KEYS: ReadonlyMap<string, TranslationKey> = new Map<string, TranslationKey>([
  ['normal', 'styles.normal'],
  ['title', 'styles.title'],
  ['subtitle', 'styles.subtitle'],
  ['no spacing', 'styles.noSpacing'],
  ['quote', 'styles.quote'],
  ['intense quote', 'styles.intenseQuote'],
  ['list paragraph', 'styles.listParagraph'],
]);

/**
 * The label of a paragraph style. A document stores the English `w:name` of a built-in style
 * whatever language wrote it, so a built-in style takes the catalog label. The match uses the
 * name, not the style id, because a style id can be localized (`Ttulo1`). A custom style keeps
 * its own name.
 */
export function paragraphStyleLabel(name: string, styleId: string, t: TFunction): string {
  const lower = name.toLowerCase();
  const key = BUILT_IN_STYLE_KEYS.get(lower);
  if (key) return t(key);
  const heading = BUILT_IN_HEADING_NAME.exec(lower);
  if (heading) return t('styles.headingLevel', { level: Number(heading[1]) });
  return name || styleId;
}

/**
 * The paragraph style of the selection, the document's paragraph styles with their labels in
 * the active language, and a function that applies one. Call it in an injection context.
 *
 * @public
 */
export function injectParagraphStyle(): ParagraphStyleState {
  const editor = injectDocxEditor();
  const zone = inject(NgZone);
  const { t } = injectTranslation();
  const value = injectEditorState((snapshot) => snapshot.formatting?.styleId ?? null);
  const command = injectEditorCommand('styles.style');
  // Read on each engine tick: a newly loaded document can leave the snapshot as it was.
  const options = injectEditorState(
    (snapshot): readonly ParagraphStyleOption[] => {
      const instance = editor();
      if (!instance || snapshot.isLoading) return EMPTY_STYLES;
      const translate = t();
      return untracked(() => instance.getDocumentStyles())
        .filter((style) => style.type === 'paragraph')
        .map((style) => ({
          styleId: style.styleId,
          name: paragraphStyleLabel(style.name, style.styleId, translate),
          preview: style.preview,
        }));
    },
    (a, b) => sameList(a, b, sameStyle)
  );
  return {
    value,
    options,
    setValue: (styleId) => applySlotValue(zone, untracked(editor), 'styles.style', styleId),
    isEnabled: command.isEnabled,
  };
}

function applySlotValue(
  zone: NgZone,
  editor: ReturnType<ReturnType<typeof injectDocxEditor>>,
  slot: 'font.family' | 'styles.style',
  value: string
): void {
  if (!editor) return;
  const command = commandForSlotValue(slot, value);
  if (!command) return;
  zone.runOutsideAngular(() => {
    if (editor.can(command).ok) editor.exec(command);
  });
}

/** What {@link injectZoom} returns. @public */
export interface ZoomState {
  /** The zoom factor, where `1` is 100 percent. */
  readonly zoom: Signal<number>;
  /** The zoom mode: fixed, or a fit to the page or its width. */
  readonly mode: Signal<ZoomMode>;
  /** True when the zoom fits the viewport. */
  readonly isFit: Signal<boolean>;
  readonly setZoom: (zoom: number) => void;
  readonly setMode: (mode: ZoomMode | 'auto') => void;
  readonly fitToWidth: () => void;
  readonly auto: () => void;
  /** Go back to a fixed 100 percent. */
  readonly reset: () => void;
  readonly zoomIn: () => void;
  readonly zoomOut: () => void;
  readonly canZoomIn: Signal<boolean>;
  readonly canZoomOut: Signal<boolean>;
  /** The zoom steps. */
  readonly levels: readonly number[];
}

const FIXED: ZoomMode = { type: 'fixed' };

/**
 * The zoom of the pages, and functions that change it. Call it in an injection context.
 *
 * @public
 */
export function injectZoom(): ZoomState {
  const editor = injectDocxEditor();
  const zone = inject(NgZone);
  const zoom = injectEditorState((snapshot) => snapshot.zoom);
  const rawMode = injectEditorState((snapshot) => snapshot.zoomMode);
  const mode = computed(() => rawMode() ?? FIXED);
  const previous = computed(() => stepZoomLevel(zoom(), 'out'));
  const next = computed(() => stepZoomLevel(zoom(), 'in'));
  const setZoom = (value: number) =>
    zone.runOutsideAngular(() => void untracked(editor)?.setZoom(value));
  const setMode = (value: ZoomMode | 'auto') =>
    zone.runOutsideAngular(() => void untracked(editor)?.setZoomMode(value));
  return {
    zoom,
    mode,
    isFit: computed(() => mode().type === 'fit'),
    setZoom,
    setMode,
    fitToWidth: () => setMode(FIT_WIDTH_ZOOM_MODE),
    auto: () => setMode('auto'),
    reset: () => {
      setMode(FIXED);
      setZoom(1);
    },
    zoomIn: () => {
      const value = untracked(next);
      if (value !== null) setZoom(value);
    },
    zoomOut: () => {
      const value = untracked(previous);
      if (value !== null) setZoom(value);
    },
    canZoomIn: computed(() => editor() !== null && next() !== null),
    canZoomOut: computed(() => editor() !== null && previous() !== null),
    levels: ZOOM_LEVELS,
  };
}
