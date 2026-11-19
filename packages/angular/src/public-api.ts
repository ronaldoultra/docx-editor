/*
 * Public API of @docx-editor.dev/angular.
 */
export {
  DocxEditorComponent,
  type DocxEditorDocument,
  type DocxEditorMode,
} from './lib/docx-editor.component';
export {
  DocxEditorRootDirective,
  type DocxEditorFonts,
  type DocxEditorRootMode,
  type DocxEditorTranslate,
} from './lib/root.directive';
export {
  injectDocxSource,
  injectFonts,
  type DocxFontOrigin,
  type DocxFontsInput,
  type DocxFontsSource,
  type DocxSource,
  type DocxSourceOptions,
  type DocxSourceState,
  type FontOriginInput,
} from './lib/fonts';
export { DocxEditorViewportComponent } from './lib/viewport.component';
export { DocxEditorContentComponent } from './lib/content.component';
export { DocxEditorWindowComponent, type DocxEditorWindowStyles } from './lib/window.component';
export { DocxEditorContext, injectDocxEditor } from './lib/context';
export { injectEditorEvent, injectEditorHostDocument, injectEditorState } from './lib/editor-state';
export {
  injectEditorCommand,
  injectEditorValueCommand,
  injectHistoryGroup,
  type EditorCommandState,
  type EditorCommandTarget,
  type EditorValueCommandState,
  type EditorValueSlot,
  type HistoryGroupOptions,
  type HistoryGroupState,
} from './lib/editor-command';
export {
  DocxEditorI18n,
  DocxEditorI18nDirective,
  injectChromeDir,
  injectTranslation,
  provideDocxEditorI18n,
  type DocxEditorDirection,
  type DocxEditorI18nOptions,
  type DocxEditorTranslation,
} from './lib/i18n';
export {
  injectEditorColorMode,
  injectScopeClassName,
  type DocxEditorColorMode,
  type EditorColorModeState,
  type ResolvedColorMode,
} from './lib/color-mode';
// The engine and i18n types that the signatures above use, so a host can name them without a
// second import path. These are type exports only: the engine stays one peer copy.
export type {
  DocumentChange,
  DocumentHandle,
  DocumentOpenFailure,
  DocumentSource,
  Editor,
  EditorCommand,
  EditorError,
  EditorEvents,
  EditorExecOptions,
  EditorFontError,
  EditorScope,
  EditorSnapshot,
  ExecResult,
  FontConfiguration,
  ZoomMode,
} from '@docx-editor.dev/core/contracts/editor';
export type {
  ChromeSlotId,
  ContentLockPolicy,
  DocxEditorConfig,
  DocxEditorInstance,
  EditorModule,
  FieldResultsMode,
  FontConfigurationFragment,
  FontOrigin,
  FontResolver,
  FontShapingOptions,
  HistoryGroupBindingOptions,
  ImageDecodePort,
  LockedControlNavigation,
  MarkedFontResolver,
  ObjectSlotValueMap,
  ToolbarValueMap,
} from '@docx-editor.dev/core/editor';
export type { LocaleStrings, TFunction, Translations } from '@docx-editor.dev/i18n';
