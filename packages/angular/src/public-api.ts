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
  type DocxEditorRootMode,
  type DocxEditorTranslate,
} from './lib/root.directive';
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
export type { DocumentChange } from '@docx-editor.dev/core/contracts/editor';
