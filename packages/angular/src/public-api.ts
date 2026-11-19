/*
 * Public API of @docx-editor.dev/angular.
 */
export {
  DocxEditorComponent,
  type DocxEditorContextMenuOptions,
  type DocxEditorDocument,
  type DocxEditorMenuOptions,
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
export { DocxSlotDirective } from './lib/slot.directive';
export { DocxEditorToolbarComponent } from './lib/toolbar/toolbar.component';
export {
  DocxToolbarButtonComponent,
  DocxToolbarSeparatorComponent,
} from './lib/toolbar/toolbar-button.component';
export { DocxToolbarActionComponent } from './lib/toolbar/toolbar-action.component';
export { DocxToolbarGroupDirective } from './lib/toolbar/toolbar-group.directive';
export { DocxChromeIconComponent } from './lib/toolbar/chrome-icon.component';
export {
  DocxToolbarFontFamilyComponent,
  DocxToolbarParagraphStyleComponent,
} from './lib/toolbar/toolbar-pickers.component';
export {
  DocxToolbarFontSizeComponent,
  DocxToolbarZoomComponent,
} from './lib/toolbar/toolbar-steppers.component';
export { DocxToolbarEditingModeComponent } from './lib/toolbar/toolbar-editing-mode.component';
export { DocxToolbarReviewersComponent } from './lib/toolbar/toolbar-reviewers.component';
export {
  DocxToolbarFontColorComponent,
  DocxToolbarHighlightComponent,
} from './lib/toolbar/toolbar-colors.component';
export {
  DOCX_PARAGRAPH_DIALOG,
  DocxToolbarAlignmentComponent,
  DocxToolbarLineSpacingComponent,
} from './lib/toolbar/toolbar-paragraph.component';
export {
  DOCX_TOOLBAR_DIALOGS,
  DocxToolbarTableBorderColorComponent,
  DocxToolbarTableBorderStyleComponent,
  DocxToolbarTableBorderTargetComponent,
  DocxToolbarTableBorderWidthComponent,
  DocxToolbarTableCellFillComponent,
  DocxToolbarTableSplitCellComponent,
  injectTableBorderTargetLabel,
} from './lib/toolbar/toolbar-table.component';
export { DocxToolbarTableInsertComponent } from './lib/toolbar/toolbar-table-insert.component';
export { DocxEditorMenuComponent } from './lib/menu/menu-bar.component';
export { DocxMenuComponent } from './lib/menu/menu.component';
export { DocxMenuItemComponent } from './lib/menu/menu-item.component';
export {
  DocxMenuGroupComponent,
  DocxMenuRowComponent,
  DocxMenuSeparatorComponent,
} from './lib/menu/menu-row.component';
export { DocxMenuSubmenuComponent } from './lib/menu/menu-submenu.component';
export { DocxMenuReviewersComponent } from './lib/menu/menu-pickers.component';
export type { DocxMenuTranslate } from './lib/menu/menu-state';
export { DocxEditorContextMenuComponent } from './lib/context-menu/context-menu.component';
export {
  DocxContextMenuCellAlignmentComponent,
  DocxContextMenuRowComponent,
  type ContextMenuPackagedRow,
} from './lib/context-menu/context-menu-row.component';
export {
  DocxContextMenuItemComponent,
  DocxContextMenuRowDirective,
  DocxContextMenuStartComponent,
} from './lib/context-menu/context-menu-item.component';
export {
  injectContextMenuAnchor,
  injectContextMenuHighlights,
  injectContextMenuTarget,
  type ContextMenuRowId,
} from './lib/context-menu/context-menu-state';
export type { ContextMenuAnchor } from './lib/context-menu/context-menu-target';
// Internal types that the protected members of the toolbar parts name. They are not API.
export type {
  BarGroup,
  BarItem,
  DefaultEntry,
  PanelRow,
  PanelSection,
  SlotOverrides,
} from './lib/toolbar/toolbar-arrangement';
export type { ToolbarArrangementKey } from './lib/toolbar/toolbar-parts';
export type { ModeOption } from './lib/toolbar/toolbar-editing-mode.component';
export type { DocxMenuBarState, MenuHandlers, MenuRegistration } from './lib/menu/menu-state';
export type {
  ContextMenuEntry,
  ContextMenuRowOverride,
} from './lib/context-menu/context-menu.component';
export { ColorSplitBase, type ColorSplitConfig } from './lib/toolbar/toolbar-colors.component';
export { TableColorBase, TableControlBase } from './lib/toolbar/toolbar-table.component';
export { DocxTableChromeState, type TableSlotState } from './lib/toolbar/table-chrome-state';
export {
  injectFontFamily,
  injectParagraphStyle,
  injectZoom,
  type FontFamilyState,
  type ParagraphStyleOption,
  type ParagraphStyleState,
  type ZoomState,
} from './lib/toolbar/picker-state';
export {
  DOCX_TOOLBAR_PLACEMENT,
  injectToolbarLabel,
  type DocxToolbarPlacement,
  type DocxToolbarTranslate,
} from './lib/toolbar/toolbar-context';
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
export type { LocaleStrings, TFunction, Translations } from '@docx-editor.dev/core/i18n';
export {
  DOCX_EDITOR_POPUPS,
  injectDocxEditorPopups,
  provideDocxEditorPopups,
  type DocxEditorPopup,
  type DocxEditorPopupEntry,
  type DocxEditorPopups,
} from './lib/popups/popup-config';
export {
  DocxDialogApplyDirective,
  DocxDialogCancelDirective,
  DocxDialogFieldDirective,
  DocxDialogPartDirective,
  DocxDialogResetDirective,
  injectDialog,
  type DialogCustomizationProps,
  type DialogPartProps,
  type DocxDialogState,
} from './lib/dialog/dialog-parts';
export { DocxNativeDialogComponent } from './lib/dialog/native-dialog.component';
