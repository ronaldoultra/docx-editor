# Angular adapter parity ledger

This ledger maps each `@docx-editor.dev/vue` export and each Vue test file to its Angular counterpart. `bun run check:angular-parity` (`scripts/check-angular-parity.mjs`) reads it. See design decision D6.

## Rules

- The export rows come from `docs/api/docx-editor-vue/index.api.md`. Exports with `@deprecated` or `@internal` have no row.
- The test rows come from the `*.test.ts` files under `packages/vue/test`.
- The status is `todo`, `ported`, or `dropped`.
- A `ported` row names its Angular counterpart: an export name, or a spec file path for a test.
- A `dropped` row gives its reason in the notes, for example "deprecated shell" or "Vue-only mechanism".
- For a test row, the notes name the Vue test claims that the Angular spec keeps.
- The Vue package can leave only when no row is `todo` (`bun run check:angular-parity --no-todo`).

## Exports

626 rows.

| Vue export | Kind | Angular counterpart | Status | Notes |
| --- | --- | --- | --- | --- |
| `acceptedOffsetFromText` | re-export from core | | todo | |
| `AcceptedTextRange` | re-export from core | | todo | |
| `acceptedTextRanges` | re-export from core | | todo | |
| `AnchorHighlightAnimation` | re-export from core | | todo | |
| `AnchorHighlightOptions` | re-export from core | | todo | |
| `BookmarkDialogFields` | interface | | todo | |
| `CHART_TYPE_CHOICES` | re-export from core | | todo | |
| `ChartColor` | re-export from core | | todo | |
| `ChartData` | re-export from core | | todo | |
| `ChartFormat` | re-export from core | | todo | |
| `ChartFormatChange` | re-export from core | | todo | |
| `ChartInsert` | const | | todo | |
| `ChartInsertPartComponent` | interface | | todo | |
| `ChartInsertProps` | interface | | todo | |
| `ChartPaneChange` | re-export from core | | todo | |
| `ChartPartProps` | type | | todo | |
| `ChartSeriesColorsProps` | interface | | todo | |
| `ChartTogglePartProps` | type | | todo | |
| `ChartType` | re-export from core | | todo | |
| `chartTypeLabelKey` | re-export from core | | todo | |
| `CHROME_GROUPS` | re-export from core | | todo | |
| `CHROME_MENUS` | re-export from core | | todo | |
| `ChromeExportFormat` | re-export from core | | todo | |
| `ChromeExportHandlers` | re-export from core | | todo | |
| `ChromeMenu` | re-export from core | | todo | |
| `ChromeMenuEntry` | re-export from core | | todo | |
| `ChromeMenuId` | re-export from core | | todo | |
| `ChromeMenuItemEntry` | re-export from core | | todo | |
| `ChromeMenuSeparatorEntry` | re-export from core | | todo | |
| `chromeMenuSlots` | re-export from core | | todo | |
| `ChromeMenuSubmenuEntry` | re-export from core | | todo | |
| `ChromeSlotId` | re-export from core | `ChromeSlotId` | ported |  |
| `ChromeTranslate` | type | | todo | |
| `ClearAnchorHighlightOptions` | re-export from core | | todo | |
| `ClearRefreshHighlightsOptions` | re-export from core | | todo | |
| `ColorPick` | re-export from core | | todo | |
| `ColorPicker` | const | | todo | |
| `ColorPickerCompound` | type | | todo | |
| `ColorPickerContentProps` | interface | | todo | |
| `ColorPickerHexInputProps` | interface | | todo | |
| `ColorPickerNoColorProps` | interface | | todo | |
| `ColorPickerRootProps` | interface | | todo | |
| `ColorPickerSwatchProps` | interface | | todo | |
| `ColorPickerTriggerProps` | interface | | todo | |
| `ColorPickerValue` | type | | todo | |
| `commandForSlot` | re-export from core | | todo | |
| `CompareDocumentsChromeHandlers` | re-export from core | | todo | |
| `CompareDocumentsDialogSession` | re-export from core | | todo | |
| `CompareDocumentsDialogValues` | interface | | todo | |
| `CompareDocumentsEngineErrorCode` | re-export from core | | todo | |
| `CompareDocumentsErrorCode` | re-export from core | | todo | |
| `CompareDocumentsOutcome` | re-export from core | | todo | |
| `CompareDocumentsRefusalReason` | re-export from core | | todo | |
| `CompareDocumentsRequest` | re-export from core | | todo | |
| `composeFontConfiguration` | re-export from core | | todo | |
| `composeFontOrigins` | re-export from core | | todo | |
| `CONTENT_CONTROL_SLOTS` | const | | todo | |
| `ContentControlActionProps` | interface | | todo | |
| `ContentControlInspectorState` | interface | | todo | |
| `ContentControlLock` | type | | todo | |
| `ContentControlPartProps` | interface | | todo | |
| `ContentControlProps` | interface | | todo | |
| `ContentControlSlotId` | type | | todo | |
| `ContentControlWidgetDayProps` | interface | | todo | |
| `ContentControlWidgetEntry` | interface | | todo | |
| `ContentControlWidgetItemProps` | interface | | todo | |
| `ContentControlWidgetPictureProps` | interface | | todo | |
| `ContextMenuAnchor` | interface | `ContextMenuAnchor` | ported |  |
| `ContextMenuCellVerticalAlignment` | const | `DocxContextMenuCellAlignmentComponent` | ported | `docx-context-menu-cell-alignment`. |
| `ContextMenuCommandProps` | interface | `DocxContextMenuRowComponent` | ported | The inputs `labelKey`, `shortcutKey`, `icon` (Material Symbols paths), `className`, and `hidden`. |
| `ContextMenuContextValue` | interface |  | dropped | Dropped: the Vue context value is internal state. Angular parts read it with `injectContextMenuAnchor` and `injectContextMenuTarget`. |
| `ContextMenuCopy` | const | `DocxContextMenuRowComponent` | ported | `<docx-context-menu-row row="edit.copy">`. One component with a `row` input replaces the Vue component of each row. |
| `ContextMenuCopyFormatting` | const | `DocxContextMenuRowComponent` | ported | `<docx-context-menu-row row="format.copyFormatting">`. One component with a `row` input replaces the Vue component of each row. |
| `ContextMenuCut` | const | `DocxContextMenuRowComponent` | ported | `<docx-context-menu-row row="edit.cut">`. One component with a `row` input replaces the Vue component of each row. |
| `ContextMenuDelete` | const | `DocxContextMenuRowComponent` | ported | `<docx-context-menu-row row="edit.delete">`. One component with a `row` input replaces the Vue component of each row. |
| `ContextMenuDeleteTable` | const | `DocxContextMenuRowComponent` | ported | `<docx-context-menu-row row="table.deleteTable">`. One component with a `row` input replaces the Vue component of each row. |
| `ContextMenuDeleteTableColumn` | const | `DocxContextMenuRowComponent` | ported | `<docx-context-menu-row row="table.deleteColumn">`. One component with a `row` input replaces the Vue component of each row. |
| `ContextMenuDeleteTableRow` | const | `DocxContextMenuRowComponent` | ported | `<docx-context-menu-row row="table.deleteRow">`. One component with a `row` input replaces the Vue component of each row. |
| `ContextMenuInsertColumnLeft` | const | `DocxContextMenuRowComponent` | ported | `<docx-context-menu-row row="table.insertColumnLeft">`. One component with a `row` input replaces the Vue component of each row. |
| `ContextMenuInsertColumnRight` | const | `DocxContextMenuRowComponent` | ported | `<docx-context-menu-row row="table.insertColumnRight">`. One component with a `row` input replaces the Vue component of each row. |
| `ContextMenuInsertRowAbove` | const | `DocxContextMenuRowComponent` | ported | `<docx-context-menu-row row="table.insertRowAbove">`. One component with a `row` input replaces the Vue component of each row. |
| `ContextMenuInsertRowBelow` | const | `DocxContextMenuRowComponent` | ported | `<docx-context-menu-row row="table.insertRowBelow">`. One component with a `row` input replaces the Vue component of each row. |
| `ContextMenuItem` | const | `DocxContextMenuItemComponent` | ported | `docx-context-menu-item`. The Vue `onSelect` is the `activate` output. |
| `ContextMenuItemProps` | interface | `DocxContextMenuItemComponent` | ported | Inputs `label`, `icon`, `shortcut`, `disabled`, `reason` (Vue `disabledReason`), `active`, `className`, and the `activate` output. |
| `ContextMenuMergeCells` | const | `DocxContextMenuRowComponent` | ported | `<docx-context-menu-row row="table.mergeCells">`. One component with a `row` input replaces the Vue component of each row. |
| `ContextMenuPaste` | const | `DocxContextMenuRowComponent` | ported | `<docx-context-menu-row row="edit.paste">`. One component with a `row` input replaces the Vue component of each row. |
| `ContextMenuPasteFormatting` | const | `DocxContextMenuRowComponent` | ported | `<docx-context-menu-row row="format.pasteFormatting">`. One component with a `row` input replaces the Vue component of each row. |
| `ContextMenuPasteWithoutFormatting` | const | `DocxContextMenuRowComponent` | ported | `<docx-context-menu-row row="edit.pasteWithoutFormatting">`. One component with a `row` input replaces the Vue component of each row. |
| `ContextMenuRefreshToc` | const | `DocxContextMenuRowComponent` | ported | `<docx-context-menu-row row="toc.refresh">`. One component with a `row` input replaces the Vue component of each row. |
| `ContextMenuRefreshTocPageNumbers` | const | `DocxContextMenuRowComponent` | ported | `<docx-context-menu-row row="toc.refreshPageNumbers">`. One component with a `row` input replaces the Vue component of each row. |
| `ContextMenuRemoveToc` | const | `DocxContextMenuRowComponent` | ported | `<docx-context-menu-row row="toc.remove">`. One component with a `row` input replaces the Vue component of each row. |
| `ContextMenuSelectAll` | const | `DocxContextMenuRowComponent` | ported | `<docx-context-menu-row row="edit.selectAll">`. One component with a `row` input replaces the Vue component of each row. |
| `ContextMenuSelectTable` | const | `DocxContextMenuRowComponent` | ported | `<docx-context-menu-row row="table.selectTable">`. One component with a `row` input replaces the Vue component of each row. |
| `ContextMenuSplitCell` | const | `DocxContextMenuRowComponent` | ported | `<docx-context-menu-row row="table.splitCell">`. The row shows in a table when the menu has a `splitCellHandler`. The split cell dialog lands in task 6.4. |
| `ContextMenuStart` | const | `DocxContextMenuStartComponent` | ported | `docx-context-menu-start`. The menu adds one separator below all starts, and only when a start has rows. |
| `ContextMenuStartProps` | interface | `DocxContextMenuStartComponent` | ported | The rows are projected content. |
| `ContextMenuTableRowProps` | interface | `DocxContextMenuRowComponent` | ported | The input `destructive` on a table row. |
| `createDocumentRefresh` | re-export from core | | todo | |
| `createFontSource` | re-export from core | | todo | |
| `DEFAULT_FONT` | re-export from core | | todo | |
| `DEFAULT_WATERMARK_PRESETS` | const | `DEFAULT_WATERMARK_PRESETS` | ported |  |
| `deferFonts` | re-export from core | | todo | |
| `defineFontResolver` | re-export from core | | todo | |
| `definePopup` | function | | dropped | Vue-only mechanism: an Angular registry entry is a component class, which needs no adapter. `DocxEditorPopup` takes the class. |
| `DialogCustomizationProps` | interface | `DialogCustomizationProps`, `DocxEditorDialogBase` | ported | `style` is `styles`. Vue `children` is the dialog's content. The open-and-close dialogs (page setup, paragraph, watermark) take these inputs from `DocxEditorDialogBase`. Added: `docxDialogOutlet` renders a packaged part, such as `field:spaceBefore` or `preview`, in a host's layout. |
| `DialogPartProps` | interface | `DialogPartProps`, `DocxDialogPartDirective` | ported | The props are the inputs of an `ng-template docxDialogPart` inside the dialog. `style` is `styles`, and `id`, `title`, `aria-*`, and `data-*` go in `attributes`. The template replaces the part's content; with `asChild` it replaces the element, and `docxDialogApply`, `docxDialogCancel`, `docxDialogReset`, and `docxDialogField` wire the host's element. Added: `keepContent`. |
| `DocumentOpenFailure` | re-export from core | `DocumentOpenFailure` | ported |  |
| `DocumentOpenFailureReason` | re-export from core | | todo | |
| `DocumentOutline` | const | | todo | |
| `DocumentRefresh` | re-export from core | | todo | |
| `DocumentRefreshError` | re-export from core | | todo | |
| `DocumentRefreshState` | re-export from core | | todo | |
| `DocumentSearchFindOptions` | re-export from core | | todo | |
| `DocumentSearchHighlight` | re-export from core | | todo | |
| `DocumentSearchNavigateOptions` | re-export from core | | todo | |
| `DocumentSearchOptions` | re-export from core | | todo | |
| `DocumentStyleSummary` | type | | todo | |
| `DocxDocument` | re-export from core | | todo | |
| `DocxEditor` | const | `DocxEditorComponent` | todo | `<docx-editor>` composes the root, the viewport, and the content. The chrome parts (toolbar, menus, rulers, navigation, dialogs) and the compound members (`DocxEditor.Toolbar` and the rest) port with phases 4 to 10. |
| `DocxEditorAuthorStyle` | const | | todo | |
| `DocxEditorAuthorStyleProps` | interface | | todo | |
| `DocxEditorBookmarkDialog` | const | | todo | |
| `DocxEditorBookmarkDialogProps` | interface | | todo | |
| `DocxEditorChartAxes` | const | | todo | |
| `DocxEditorChartCategoryAxisTitle` | const | | todo | |
| `DocxEditorChartData` | const | | todo | |
| `DocxEditorChartDataLabels` | const | | todo | |
| `DocxEditorChartGridlines` | const | | todo | |
| `DocxEditorChartLegend` | const | | todo | |
| `DocxEditorChartPane` | const | | todo | |
| `DocxEditorChartPaneProps` | interface | | todo | |
| `DocxEditorChartSeriesColors` | const | | todo | |
| `DocxEditorChartTitle` | const | | todo | |
| `DocxEditorChartType` | const | | todo | |
| `DocxEditorChartValueAxisTitle` | const | | todo | |
| `DocxEditorChildren` | type | | todo | |
| `DocxEditorColorByChangeType` | const | | todo | |
| `DocxEditorCompareDocumentsDialog` | const | | todo | |
| `DocxEditorCompareDocumentsDialogProps` | interface | | todo | |
| `DocxEditorContent` | const | `DocxEditorContentComponent` | ported |  |
| `DocxEditorContentControl` | const | | todo | |
| `DocxEditorContentControlNamespace` | interface | | todo | |
| `DocxEditorContentControlWidget` | const | | todo | |
| `DocxEditorContentControlWidgetNamespace` | interface | | todo | |
| `DocxEditorContentControlWidgetPartProps` | interface | | todo | |
| `DocxEditorContentControlWidgetProps` | interface | | todo | |
| `DocxEditorContentProps` | interface | `DocxEditorContentComponent` | ported | Angular hosts set `class` on the host element. |
| `DocxEditorContextMenu` | const | `DocxEditorContextMenuComponent` | ported | `div[docxEditorContextMenu]`. Overrides are `ng-template[docxContextMenuRow]` (`DocxContextMenuRowDirective`) and `ng-template[docxSlot]`. The Add comment row ports with the review rail in task 10.1. |
| `DocxEditorContextMenuNamespace` | interface |  | dropped | Dropped: Angular has no compound namespace object. Each part is a standalone component. The menu bar parts (`docx-menu-item`, `docx-menu-submenu`, `docxMenuRow`, `docxMenuSeparator`, `docxMenuGroup`) also work in the context menu. |
| `DocxEditorContextMenuProps` | interface | `DocxEditorContextMenuComponent` | ported | Inputs `preset`, `disabled`, `t`, `className`, `splitCellHandler`, and the `openChange` output (Vue `onOpenChange`). The `contextMenu` input of `<docx-editor>` takes the same options as `DocxEditorContextMenuOptions`, with `onOpenChange` for the output. |
| `DocxEditorDocumentOutline` | const | | todo | |
| `DocxEditorDocumentOutlineProps` | interface | | todo | |
| `DocxEditorEquation` | const | | todo | |
| `DocxEditorExportDialog` | const | | todo | |
| `DocxEditorExportDialogProps` | interface | | todo | |
| `DocxEditorFontNotice` | const | | todo | |
| `DocxEditorFontNoticeProps` | interface | | todo | |
| `DocxEditorHeaderFooterChrome` | const | | todo | |
| `DocxEditorHeaderFooterChromeProps` | interface | | todo | |
| `DocxEditorHorizontalRuler` | const | | todo | |
| `DocxEditorHyperLink` | const | `DocxEditorHyperlinkComponent` | ported | `<docx-editor-hyperlink>`. Put it inside the viewport, or name it as the `hyperlink` entry of the popup registry. The parts are `docxHyperlinkPart` templates. |
| `DocxEditorHyperLinkNamespace` | interface | `DocxHyperlinkPartDirective`, `DocxHyperlinkUrlDirective`, `DocxHyperlinkCopyDirective`, `DocxHyperlinkEditDirective`, `DocxHyperlinkUnlinkDirective`, `DocxHyperlinkApplyDirective`, `DocxHyperlinkCancelDirective` | ported | Vue `Url`, `Copy`, `Edit`, `Unlink`, `Fields`, `Error`, `Apply`, and `Cancel` are part names of a `docxHyperlinkPart` template. The wiring directives wire a host's `asChild` element. `Fields` and `Error` have no wiring directive. |
| `DocxEditorImageAltTextPopup` | const | | todo | |
| `DocxEditorImageAltTextPopupProps` | interface | | todo | |
| `DocxEditorImageInsertErrorDialog` | const | | todo | |
| `DocxEditorImageInsertErrorDialogProps` | interface | | todo | |
| `DocxEditorImagePropertiesDialog` | const | | todo | |
| `DocxEditorImagePropertiesDialogProps` | interface | | todo | |
| `DocxEditorInvalidTextFormFieldDialog` | const | | todo | |
| `DocxEditorInvalidTextFormFieldDialogProps` | interface | | todo | |
| `DocxEditorLoading` | const | | todo | |
| `DocxEditorLoadingComponent` | interface | | todo | |
| `DocxEditorLoadingProps` | interface | | todo | |
| `DocxEditorLoadingSpinner` | const | | todo | |
| `DocxEditorLoadingSpinnerProps` | interface | | todo | |
| `DocxEditorMenu` | const | `DocxEditorMenuComponent` | ported | `div[docxEditorMenu]`. The compound members are standalone parts (`docx-menu`, `docx-menu-item`, `docx-menu-submenu`, `docxMenuRow`, `docxMenuSeparator`, `docxMenuGroup`, `docx-menu-reviewers`) and `docxSlot` row overrides, see D3. A host `docx-menu` with a registry id in a plain `menu` attribute replaces that menu in place. |
| `DocxEditorMenuNamespace` | interface |  | dropped | Dropped: Angular has no compound namespace object. Each menu part is a standalone component with its own selector, and the pinned parts (`Menu.File`, `Menu.Save`, `Menu.PageSetup`, and the rest) are `<docx-menu menu="file">` and `<docx-menu-item slotId="file.save">`. |
| `DocxEditorMenuProps` | interface | `DocxEditorMenuComponent` | ported | Inputs `preset`, `overflow`, `t`, `fileName`, `openHandler`, `saveHandler`, `pageSetupHandler`, `paragraphDialogHandler`, `watermarkDialogHandler`, `reportIssueHandler`, `reportIssue`, and the `openFile` output. `className` is the host `class`. `exporters` and `printSettings` port with task 6.5. Without a handler, the page setup, paragraph, and watermark rows open the root's dialogs (task 6.3). The `menu` input of `<docx-editor>` takes the same options as `DocxEditorMenuOptions`, with `onOpenFile` for the output. |
| `DocxEditorNamespace` | interface | | todo | |
| `DocxEditorNavigation` | const | | todo | |
| `DocxEditorNavigationNamespace` | interface | | todo | |
| `DocxEditorNavigationProps` | interface | | todo | |
| `DocxEditorNotePreview` | const | | todo | |
| `DocxEditorNotePreviewProps` | interface | | todo | |
| `DocxEditorNotePropertiesDialog` | const | | todo | |
| `DocxEditorNotePropertiesDialogProps` | interface | | todo | |
| `DocxEditorNotesChrome` | const | | todo | |
| `DocxEditorNotesChromeProps` | interface | | todo | |
| `DocxEditorNotesContextMenu` | const | | todo | |
| `DocxEditorNotesContextMenuProps` | interface | | todo | |
| `DocxEditorOpenFailure` | const | | todo | |
| `DocxEditorOpenFailureAction` | const | | todo | |
| `DocxEditorOpenFailureActionProps` | interface | | todo | |
| `DocxEditorOpenFailureComponent` | interface | | todo | |
| `DocxEditorOpenFailureContext` | interface | | todo | |
| `DocxEditorOpenFailureMessage` | type | | todo | |
| `DocxEditorOpenFailureProps` | interface | | todo | |
| `DocxEditorPageNumber` | const | | todo | |
| `DocxEditorPageNumberProps` | interface | | todo | |
| `DocxEditorPageSetupDialog` | const | `DocxEditorPageSetupDialogComponent` | ported | `<docx-editor-page-setup-dialog>`. The Vue compound parts are `docxDialogPart` templates and `docxDialogOutlet`. |
| `DocxEditorPageSetupDialogProps` | interface | `DocxEditorDialogBase` | ported | Inputs `open` (default `true`), `closeHandler`, `className`, `styles`, and `preset`, and the `closed` output in the place of `onClose`. |
| `DocxEditorPaneSections` | const | | todo | |
| `DocxEditorPaneSectionsProps` | interface | | todo | |
| `DocxEditorParagraphDialog` | const | `DocxEditorParagraphDialogComponent` | ported | `<docx-editor-paragraph-dialog>`. The Vue compound parts are `docxDialogPart` templates and `docxDialogOutlet`. Changed: Clear all sets `clearedAllTabStops`, so a cleared list that equals the seed is still a decision. |
| `DocxEditorParagraphDialogProps` | interface | `DocxEditorDialogBase` | ported | Inputs `open` (default `true`), `closeHandler`, `className`, `styles`, and `preset`, and the `closed` output in the place of `onClose`. |
| `DocxEditorPopup` | type | `DocxEditorPopup` | ported | A component class, a `DocxEditorPopupEntry` (`{ component, inputs }`), or `false`, in the place of a render function. The editor renders the class and sets its inputs, such as `session`. |
| `DocxEditorPopups` | interface | `DocxEditorPopups` | todo | Has `contentControlWidget`, `contentControlCheckbox`, `contentControlPicture`, `invalidTextFormField`, `hyperlink`, `contentControl`, `equation`, `contextMenu`; `DOCX_EDITOR_POPUPS` and `provideDocxEditorPopups()` provide it, and the root's `popups` input replaces single entries. Since task 6.3: `pageSetup`, `paragraph`, and `watermark`, with the input `closeHandler`. Missing: `bookmark`, `splitCell`, `textFormField` (6.4), `export`, `print` (6.5), `imageInsertError`, `imageProperties`, `imageAltText` (8.1), `noteProperties`, `notesContextMenu`, `notePreview` (9.2), `revisionMarkup` (10.2), `compareDocuments` (10.3). |
| `DocxEditorPrintDialog` | const | | todo | |
| `DocxEditorPrintDialogProps` | interface | | todo | |
| `DocxEditorProps` | interface | `DocxEditorComponent` | todo | Has `document`, `chrome`, `title` (a `model`, so `[(title)]`; Vue `onTitleChange` is the `titleChange` output, and the title is always editable), `menu`, `contextMenu`, `mode`, `colorMode`, `i18n`, `translate` (Vue `t`), `locale`, `author`, `zoom`, `zoomMode`, `fieldResults`, `contentLockPolicy`, `lockedControlNavigation`, `fonts`, `fontShaping`, `modules`, `imageDecodePort`, and the outputs `ready`, `change`, `selectionChange`, `error`, `openFailure`, `fontError`. Vue `onSave` and `onOpen` are `saveHandler` and `openHandler` in `[menu]`. Missing: `popups`, `hyperlinkPopup` (6.7), `navigation` (7.3), `rulers`, `openFailure` dialog (9.6), `revisionMarkup`, `reviewDisplayMode` (10.4), the `titleBarLeft` and `titleBarRight` slots, `class` (hosts set it on the element). |
| `DocxEditorRef` | interface | `DocxEditorComponent` | ported | The methods `load`, `save`, `getDocumentHandle`, `getEditor`, `focus`, `exec`, and `snapshot` are on the component instance. `save` rejects without an engine. |
| `DocxEditorRefCallback` | type | | todo | |
| `DocxEditorRevisionMarkup` | const | | todo | |
| `DocxEditorRevisionMarkupDialog` | const | | todo | |
| `DocxEditorRevisionMarkupDialogProps` | interface | | todo | |
| `DocxEditorRevisionMarkupProps` | interface | | todo | |
| `DocxEditorRoot` | const | `DocxEditorRootDirective` | ported | Directive `docxEditorRoot`. The props and listeners rows track the inputs and outputs that are still missing. |
| `DocxEditorRootListeners` | interface | `DocxEditorRootDirective` | todo | Has the `ready`, `change`, `fontError`, `selectionChange`, `error`, and `openFailure` outputs. Missing: `revisionMarkupChange` (10.4). |
| `DocxEditorRootProps` | interface | `DocxEditorRootDirective` | todo | Has `document`, `mode`, `colorMode`, `locale`, `translate`, `author`, `fonts`, `fontShaping`, `imageDecodePort`, `modules`, `zoom`, `zoomMode`, `fieldResults`, `contentLockPolicy`, `lockedControlNavigation`, `popups`. Missing: `revisionMarkup`, `reviewDisplayMode` (10.4), `tableInteractionLabel` (4.4). |
| `DocxEditorRulerProps` | interface | | todo | |
| `DocxEditorShapeFill` | const | | todo | |
| `DocxEditorShapeFillTransparency` | const | | todo | |
| `DocxEditorShapeOutlineColor` | const | | todo | |
| `DocxEditorShapeOutlineDash` | const | | todo | |
| `DocxEditorShapeOutlineWidth` | const | | todo | |
| `DocxEditorShapePane` | const | | todo | |
| `DocxEditorShapePaneProps` | interface | | todo | |
| `DocxEditorSidePanel` | const | | todo | |
| `DocxEditorSidePanelProps` | interface | | todo | |
| `DocxEditorSplitCellDialog` | const | | todo | |
| `DocxEditorSplitCellDialogProps` | interface | | todo | |
| `DocxEditorTextFormFieldDialog` | const | | todo | |
| `DocxEditorTextFormFieldDialogProps` | interface | | todo | |
| `DocxEditorToolbar` | const | `DocxEditorToolbarComponent` | ported | `fieldset[docxEditorToolbar]`. The compound members are standalone parts (`docxToolbarButton`, `docxToolbarAction`, `docxToolbarSeparator`) and `ng-template` directives (`docxSlot`, `docxToolbarGroup`), see D3. |
| `DocxEditorToolbarGroupProps` | interface | `DocxToolbarGroupDirective` | ported | The props are the directive inputs. `class` is `className`, because an `ng-template` takes no class. |
| `DocxEditorToolbarNamespace` | interface |  | dropped | Dropped: Angular has no compound namespace object. Each toolbar part is a standalone component with its own selector (see the README table of toolbar parts), and `docxSlot` replaces a part in place. |
| `DocxEditorToolbarProps` | interface | `DocxEditorToolbarComponent` | ported | The props are the component inputs `preset`, `overflow`, and `t`. `onSave` ports with the save part. |
| `DocxEditorToolbarSlotProps` | interface | `DocxSlotDirective` | ported | `slotId`, `hidden`, and `overflowContent` are `docxSlot`, `docxSlotHidden`, and `docxSlotOverflow`. |
| `DocxEditorUnsupportedElement` | const | | todo | |
| `DocxEditorUnsupportedElementProps` | interface | | todo | |
| `DocxEditorVerticalRuler` | const | | todo | |
| `DocxEditorViewport` | const | `DocxEditorViewportComponent` | ported |  |
| `DocxEditorViewportProps` | interface | `DocxEditorViewportComponent` | ported | Angular hosts set `class` and `style` on the host element. Children are content projection. |
| `DocxEditorWatermarkDialog` | const | `DocxEditorWatermarkDialogComponent`, `DocxEditorWatermarkPreviewComponent` | ported | `<docx-editor-watermark-dialog>`. `Preview` is the `preview` part and the `<docx-editor-watermark-preview>` component. The other compound parts are `docxDialogPart` templates and `docxDialogOutlet`. |
| `DocxEditorWatermarkDialogProps` | interface | `DocxEditorWatermarkDialogComponent` | ported | Input `presets`. Inputs `open` (default `true`), `closeHandler`, `className`, `styles`, and `preset`, and the `closed` output in the place of `onClose`. |
| `DocxEditorWindow` | const | `DocxEditorWindowComponent` | ported |  |
| `DocxEditorWindowProps` | interface | `DocxEditorWindowComponent`, `DocxEditorWindowStyles` | ported | `onClose` is the `close` output. `open` has two-way binding through `openChange`. |
| `DocxFontOrigin` | type | `DocxFontOrigin` | ported |  |
| `DocxFontsInput` | type | `DocxFontsInput` | ported |  |
| `DocxFontsSource` | type | `DocxFontsSource` | ported |  |
| `DocxSource` | type | `DocxSource` | ported |  |
| `DrawingGraphicKind` | re-export from core | | todo | |
| `DrawingSupport` | re-export from core | | todo | |
| `DrawingSupportIssue` | re-export from core | | todo | |
| `Editor` | re-export from core | `Editor` | ported |  |
| `EditorCaret` | interface | | todo | |
| `EditorCommand` | re-export from core | `EditorCommand` | ported |  |
| `EditorCommandState` | interface | `EditorCommandState` | ported |  |
| `EditorFontError` | re-export from core | `EditorFontError` | ported |  |
| `EditorFontErrorCode` | re-export from core | | todo | |
| `EditorMode` | type | | todo | |
| `EditorQuery` | re-export from core | | todo | |
| `EditorScope` | re-export from core | `EditorScope` | ported |  |
| `EditorSnapshot` | re-export from core | `EditorSnapshot` | ported |  |
| `EditorValueCommandState` | interface | `EditorValueCommandState` | ported |  |
| `FontConfiguration` | re-export from core | `FontConfiguration` | ported |  |
| `FontConfigurationBase` | re-export from core | | todo | |
| `FontConfigurationFragment` | re-export from core | `FontConfigurationFragment` | ported |  |
| `FontFaceRequest` | re-export from core | | todo | |
| `FontFamilyItemProps` | interface | | dropped | Dropped: the `Trigger`, `Content`, and `Item` compound parts with `asChild` are a Vue mechanism. An Angular host replaces the slot with `docxSlot` and builds its own picker on `injectFontFamily`. |
| `FontFamilyNamespace` | interface | | dropped | Dropped: the `Trigger`, `Content`, and `Item` compound parts with `asChild` are a Vue mechanism. An Angular host replaces the slot with `docxSlot` and builds its own picker on `injectFontFamily`. |
| `FontFamilyPartProps` | interface | | dropped | Dropped: the `Trigger`, `Content`, and `Item` compound parts with `asChild` are a Vue mechanism. An Angular host replaces the slot with `docxSlot` and builds its own picker on `injectFontFamily`. |
| `FontFamilyProps` | interface | `DocxToolbarFontFamilyComponent` | ported | `div[docxToolbarFontFamily]`. `hidden` is a `docxSlotHidden` template in the toolbar. |
| `FontLoadFailure` | re-export from core | | todo | |
| `FontLoadFailureReason` | re-export from core | | todo | |
| `FontOption` | interface | | todo | |
| `FontOrigin` | re-export from core | `FontOrigin` | ported |  |
| `FontResolutionRequest` | re-export from core | | todo | |
| `FontResolver` | re-export from core | `FontResolver` | ported |  |
| `FontResolverMark` | re-export from core | | todo | |
| `FontsInput` | type | `FontOriginInput` | ported | Angular reads a signal origin with `isSignal`; every other function is a resolver. |
| `FontSource` | re-export from core | | todo | |
| `FontSourceSubstitution` | re-export from core | | todo | |
| `FontUrlSource` | re-export from core | | todo | |
| `generateRulerTicks` | re-export from core | | todo | |
| `GoToBookmarkProblem` | type | | todo | |
| `HeaderFooterState` | type | | todo | |
| `HIGHLIGHT_REFRESH_MS` | re-export from core | | todo | |
| `HighlightHit` | re-export from core | | todo | |
| `HighlightKind` | re-export from core | | todo | |
| `HighlightOptions` | re-export from core | | todo | |
| `HighlightRange` | re-export from core | | todo | |
| `HighlightRect` | re-export from core | | todo | |
| `HighlightResult` | re-export from core | | todo | |
| `HighlightSource` | re-export from core | | todo | |
| `HorizontalRuler` | const | | todo | |
| `HorizontalRulerProps` | interface | | todo | |
| `HyperLinkActionProps` | interface | `DocxHyperlinkPartDirective` | ported | `icon` is the template content of the action part. |
| `HyperLinkPartProps` | interface | `DocxHyperlinkPartDirective` | ported | `className`, `asChild`, and `hidden` are inputs of the part template, and `children` is its content. Added: `styles`, `attributes`, `keepContent`. |
| `HyperlinkPopupAnchor` | interface | `HyperlinkPopupAnchor` | ported | |
| `HyperlinkPopupMode` | type | `HyperlinkPopupMode` | ported | |
| `HyperlinkPopupState` | interface | `HyperlinkPopupState` | ported | |
| `HyperLinkProps` | interface | `DocxEditorHyperlinkComponent` | ported | `className`, `preset`, and `hidden` are inputs. Dropped: `asChild` on the panel, because the panel's class carries its position. |
| `ImageAltText` | const | | todo | |
| `ImageInsertProvider` | const | | todo | |
| `ImageInsertTrigger` | const | | todo | |
| `ImagePropertiesTrigger` | const | | todo | |
| `ImagePropertiesTriggerProps` | interface | | todo | |
| `ImageWrap` | const | | todo | |
| `ImageWrapTarget` | re-export from core | | todo | |
| `IndentUpdate` | interface | | todo | |
| `isFieldLink` | function | `isFieldLink` | ported | |
| `isFontResolver` | re-export from core | | todo | |
| `ListState` | interface | | todo | |
| `ListType` | type | | todo | |
| `loadFonts` | re-export from core | | todo | |
| `LoadFontsRequest` | re-export from core | | todo | |
| `LoadFontsResult` | re-export from core | | todo | |
| `LOADING_SNAPSHOT` | re-export from core | | todo | |
| `LocaleProvider` | const | `DocxEditorI18nDirective`, `provideDocxEditorI18n` | ported |  |
| `LocaleProviderProps` | interface | `DocxEditorI18nDirective`, `DocxEditorI18nOptions` | ported | `i18n` is the `docxEditorI18n` input, `dir` is the `docxEditorDir` input. |
| `MarkedFontResolver` | re-export from core | `MarkedFontResolver` | ported |  |
| `MAX_RESOLVER_FAMILIES` | re-export from core | | todo | |
| `MaybeRefOrGetter` | type | | todo | |
| `MenuActionProps` | interface | `DocxMenuItemComponent` | ported | `className` and `hidden` are inputs of `docx-menu-item`, which routes `file.open`, `file.save`, and the dialog slots to their handlers. |
| `MenuGroupProps` | interface | `DocxMenuGroupComponent` | ported | `div[docxMenuGroup]` with `label` and `labelKey`. A host hides it with `@if`. |
| `MenuId` | type | `DocxMenuComponent` | ported | The `menu` input is a string: a registry id or a host id. |
| `MenuItemBaseProps` | interface | `DocxMenuItemComponent` | ported | Inputs `labelKey`, `shortcutKey`, `className`, and `hidden`. |
| `MenuItemProps` | interface | `DocxMenuItemComponent` | ported | The `slotId` input. The deprecated `slot` prop has no Angular counterpart. |
| `MenuPartComponent` | interface |  | dropped | Dropped: Angular pins a menu with the `menu` input of `docx-menu`, not with a component per id. |
| `MenuProps` | interface | `DocxMenuComponent` | ported | Inputs `menu` (Vue `id`), `labelKey`, `label`, `icon` (Material Symbols paths), `className`, `hidden`, and `preset`. Row overrides are `docxSlot` templates. |
| `MenuReportIssueProps` | interface | `DocxEditorMenuComponent` | ported | The bar inputs `reportIssue` (`false` removes the row) and `reportIssueHandler` (Vue `onSelect`) replace the part. |
| `MenuReviewersProps` | type | `DocxMenuReviewersComponent` | ported | Inputs `className` and `hidden`. |
| `MenuRowProps` | interface | `DocxMenuRowComponent` | ported | `button[docxMenuRow]`: the attribute value is the `data-slot` marker, inputs `icon`, `shortcut`, `disabled`, `reason` (Vue `title`), `active`, `radio` (Vue `selected`), `rowClass`, and the `activate` output (Vue `onSelect`). An `[docxMenuRowIcon]` element replaces the icon. |
| `MenuSeparatorProps` | interface | `DocxMenuSeparatorComponent` | ported | `div[docxMenuSeparator]`; the host sets `class`. |
| `MenuSubmenuBaseProps` | interface | `DocxMenuSubmenuComponent` | ported | Inputs `paths` and `className`; the content is the panel. |
| `MenuSubmenuProps` | interface | `DocxMenuSubmenuComponent` | ported | Inputs `label` and `labelKey`. |
| `MenuTableGridProps` | interface |  | dropped | Dropped: the registry `table.insert` row renders the size grid itself. A host grid is the toolbar's `docx-table-size-grid`, which ports as a public part with task 8.1 if a host needs it. |
| `NavigateHighlightCommand` | re-export from core | | todo | |
| `NavigateToChangeOptions` | re-export from core | | todo | |
| `NAVIGATION_PANE_GAP` | const | | todo | |
| `NAVIGATION_PANE_INSET` | const | | todo | |
| `NAVIGATION_PANE_MIN_PAGE_ROOM` | const | | todo | |
| `NAVIGATION_PANE_WIDTH` | const | | todo | |
| `NavigationClose` | const | | todo | |
| `NavigationFind` | const | | todo | |
| `NavigationGoTo` | const | | todo | |
| `NavigationHeader` | const | | todo | |
| `NavigationHeadings` | const | | todo | |
| `navigationPaneOverlays` | function | | todo | |
| `navigationPaneReservation` | function | | todo | |
| `NavigationPartProps` | interface | | todo | |
| `navigationShift` | function | | todo | |
| `NavigationShiftInput` | interface | | todo | |
| `NavigationTab` | const | | todo | |
| `NavigationTabProps` | interface | | todo | |
| `NavigationTabs` | const | | todo | |
| `NavigationTabValue` | type | | todo | |
| `NavigationTitle` | const | | todo | |
| `NavigationToggle` | const | | todo | |
| `NormalizedImagePayload` | type | | todo | |
| `normalizeImageBytes` | function | | todo | |
| `NotePropertiesState` | type | | todo | |
| `OBJECT_SLOT_PARTS` | const | | todo | |
| `ObjectButtonPartProps` | interface | | todo | |
| `ObjectControlProps` | interface | | todo | |
| `ObjectControlState` | interface | | todo | |
| `ObjectControlVariant` | type | | todo | |
| `ObjectPartComponent` | type | | todo | |
| `ObjectSlotId` | type | | todo | |
| `ObjectSlotReadValueMap` | re-export from core | | todo | |
| `ObjectSlotValueMap` | re-export from core | `ObjectSlotValueMap` | ported |  |
| `ObjectToggleControl` | interface | | todo | |
| `ObjectToggleSlotId` | re-export from core | | todo | |
| `OUTLINE_BUTTON_LEFT_OFFSET` | const | | todo | |
| `OUTLINE_BUTTON_RESERVED_SPACE` | const | | todo | |
| `OUTLINE_LEFT_OFFSET` | const | | todo | |
| `OUTLINE_RESERVED_SPACE` | const | | todo | |
| `OutlineHeading` | type | | todo | |
| `OutlineHeadingItem` | interface | | todo | |
| `PageIndicator` | const | | todo | |
| `PageIndicatorProps` | interface | | todo | |
| `PageNumberTranslationContext` | const | | todo | |
| `PageSetup` | re-export from core | `PageSetup` | ported |  |
| `PageSetupDialogFields` | interface | `PageSetupDialogFields` | ported |  |
| `PageSetupUpdate` | interface | `PageSetupUpdate` | ported |  |
| `PaginatedDocxEditor` | const | | todo | The proposal lists it with the deprecated shell, but the snapshot has no `@deprecated` tag. |
| `PaginatedDocxEditorExpose` | type | | todo | |
| `PaginatedDocxEditorHandle` | interface | | dropped | Deprecated shell: the proposal does not port `PaginatedDocxEditor`. `DocxEditorComponent` has the same save, load, and command methods. |
| `PaginatedDocxEditorProps` | interface | | todo | |
| `PaginatedDocxEditorShellProps` | interface | | todo | |
| `PANE_SECTIONS` | re-export from core | | todo | |
| `PaneSection` | re-export from core | | todo | |
| `PaneSectionId` | re-export from core | | todo | |
| `PaneSectionIds` | re-export from core | | todo | |
| `paneSections` | re-export from core | | todo | |
| `ParagraphAlignment` | type | | todo | |
| `ParagraphFlagState` | re-export from core | `ParagraphFlagState` | ported |  |
| `ParagraphFormatRead` | re-export from core | `ParagraphFormatRead` | ported |  |
| `ParagraphFormatUpdate` | re-export from core | `ParagraphFormatUpdate` | ported |  |
| `ParagraphsChangeEvent` | re-export from core | | todo | |
| `ParagraphStyleItemProps` | interface | | dropped | Dropped: the `Trigger`, `Content`, and `Item` compound parts with `asChild` are a Vue mechanism. An Angular host replaces the slot with `docxSlot` and builds its own picker on `injectParagraphStyle`. |
| `ParagraphStyleNamespace` | interface | | dropped | Dropped: the `Trigger`, `Content`, and `Item` compound parts with `asChild` are a Vue mechanism. An Angular host replaces the slot with `docxSlot` and builds its own picker on `injectParagraphStyle`. |
| `ParagraphStyleOption` | interface | `ParagraphStyleOption` | ported |  |
| `ParagraphStylePartProps` | interface | | dropped | Dropped: the `Trigger`, `Content`, and `Item` compound parts with `asChild` are a Vue mechanism. An Angular host replaces the slot with `docxSlot` and builds its own picker on `injectParagraphStyle`. |
| `ParagraphStyleProps` | interface | `DocxToolbarParagraphStyleComponent` | ported | `div[docxToolbarParagraphStyle]`. |
| `ParagraphTabStop` | re-export from core | `ParagraphTabStop` | ported |  |
| `ParagraphText` | re-export from core | | todo | |
| `ParagraphTextProofing` | re-export from core | | todo | |
| `ParagraphTextReadOptions` | re-export from core | | todo | |
| `ParagraphTextSpan` | re-export from core | | todo | |
| `provideDocxEditor` | function | | dropped | Vue-only mechanism. In Angular, `docxEditorRoot` owns the engine, and a host reads it with `injectDocxEditor` or the directive's `context`. |
| `ProvideDocxEditorResult` | interface | | dropped | Vue-only mechanism, see `provideDocxEditor`. |
| `PX_PER_CM` | re-export from core | | todo | |
| `PX_PER_INCH` | re-export from core | | todo | |
| `RefreshChange` | re-export from core | | todo | |
| `RefreshChangeInput` | re-export from core | | todo | |
| `RefreshFailureCode` | re-export from core | | todo | |
| `RefreshHighlightAnimation` | re-export from core | | todo | |
| `RefreshHighlightOptions` | re-export from core | | todo | |
| `RefreshLocation` | re-export from core | | todo | |
| `RefreshRefusal` | re-export from core | | todo | |
| `RefreshRefusalFields` | re-export from core | | todo | |
| `RefreshRefusalReason` | re-export from core | | todo | |
| `RefreshResult` | re-export from core | | todo | |
| `RefreshSubmission` | re-export from core | | todo | |
| `RefreshUpdate` | re-export from core | | todo | |
| `RefreshUpdateMode` | re-export from core | | todo | |
| `ReplaceAcceptedTextCommand` | re-export from core | | todo | |
| `ReplaceDocumentErrorCode` | re-export from core | | todo | |
| `ReplaceDocumentOutcome` | re-export from core | | todo | |
| `ReplaceDocumentRefusal` | re-export from core | | todo | |
| `ReplaceDocumentRefusalReason` | re-export from core | | todo | |
| `ReplaceDocumentRequest` | re-export from core | | todo | |
| `ResolvedColorMode` | type | `ResolvedColorMode` | ported |  |
| `ResolvedRevisionMarkup` | re-export from core | | todo | |
| `REVIEW_MARKERS_GUTTER` | const | | todo | |
| `REVIEW_PANE_GUTTER` | const | | todo | |
| `ReviewAuthorInfo` | re-export from core | | todo | |
| `ReviewDisplayMode` | re-export from core | | todo | |
| `reviewGutter` | function | | todo | |
| `ReviewGutter` | interface | | todo | |
| `ReviewGutterInput` | interface | | todo | |
| `ReviewRailContext` | const | | todo | |
| `ReviewRailRegistry` | interface | | todo | |
| `RevisionAuthorAssignments` | re-export from core | | todo | |
| `RevisionAuthorStyle` | re-export from core | | todo | |
| `RevisionChangedLinesMark` | re-export from core | | todo | |
| `RevisionDeletionMark` | re-export from core | | todo | |
| `RevisionMarkupChromeHandlers` | re-export from core | | todo | |
| `RevisionMarkupColor` | re-export from core | | todo | |
| `RevisionMarkupDialogSession` | re-export from core | | todo | |
| `RevisionMarkupMark` | re-export from core | | todo | |
| `RevisionMarkupNamedColor` | re-export from core | | todo | |
| `RevisionMarkupOptions` | re-export from core | | todo | |
| `RevisionStyles` | re-export from core | | todo | |
| `RULER_WIDTH` | const | | todo | |
| `rulerPageBox` | re-export from core | | todo | |
| `RulerTick` | re-export from core | | todo | |
| `RulerUnit` | re-export from core | | todo | |
| `runToolbarCommand` | re-export from core | `injectEditorCommand` | ported | `execute` runs `runToolbarCommand`. See `toolbarCommandState`. |
| `ScopedChromeAnchor` | interface | | todo | |
| `ScopedChromeAnchorPlacement` | type | | todo | |
| `ScrollToAnchorOptions` | re-export from core | | todo | |
| `SEARCH_DEBOUNCE_MS` | re-export from core | | todo | |
| `SEARCH_HIGHLIGHT_PRIORITY` | re-export from core | | todo | |
| `SEARCH_HIGHLIGHT_SET` | re-export from core | | todo | |
| `SEARCH_MATCH_LIMIT` | re-export from core | | todo | |
| `SelectedChartCommand` | re-export from core | | todo | |
| `SelectedChartState` | re-export from core | | todo | |
| `SelectedShapeCommand` | re-export from core | | todo | |
| `SHAPE_OUTLINE_DASHES` | re-export from core | | todo | |
| `SHAPE_OUTLINE_WIDTHS` | re-export from core | | todo | |
| `ShapeColorPartProps` | type | | todo | |
| `ShapeFillState` | re-export from core | | todo | |
| `ShapeFormatState` | re-export from core | | todo | |
| `ShapeOutlineDash` | re-export from core | | todo | |
| `ShapeOutlineState` | re-export from core | | todo | |
| `ShapePaneChange` | re-export from core | | todo | |
| `ShapePartProps` | type | | todo | |
| `Slot` | const | | todo | |
| `SlotProps` | interface | | todo | |
| `SplitCellDialogFields` | interface | | todo | |
| `TableAction` | type | | todo | |
| `TableBorderColorNamespace` | interface | `DocxToolbarTableBorderColorComponent` | ported | `div[docxToolbarTableBorderColor]`. The swatch body is the toolbar hex color picker; the shared color picker compound ports with phase 8. |
| `TableBorderStyleNamespace` | interface | `DocxToolbarTableBorderStyleComponent` | ported | `div[docxToolbarTableBorderStyle]`. |
| `TableBorderTargetNamespace` | interface | `DocxToolbarTableBorderTargetComponent` | ported | `div[docxToolbarTableBorderTarget]`. |
| `TableBorderWidthNamespace` | interface | `DocxToolbarTableBorderWidthComponent` | ported | `div[docxToolbarTableBorderWidth]`. |
| `TableCellFillNamespace` | interface | `DocxToolbarTableCellFillComponent` | ported | `div[docxToolbarTableCellFill]`. |
| `TableChromeItemProps` | interface | | dropped | Dropped: the `Trigger`, `Content`, `Item`, and `Main` compound parts with `asChild` are a Vue mechanism. An Angular host replaces a table slot with `docxSlot`. |
| `TableChromePartComponent` | interface | | dropped | Dropped: the `Trigger`, `Content`, `Item`, and `Main` compound parts with `asChild` are a Vue mechanism. An Angular host replaces a table slot with `docxSlot`. |
| `TableChromePartProps` | interface | | dropped | Dropped: the `Trigger`, `Content`, `Item`, and `Main` compound parts with `asChild` are a Vue mechanism. An Angular host replaces a table slot with `docxSlot`. |
| `TextFormFieldDialogFields` | interface | | todo | |
| `TextMatch` | re-export from core | | todo | |
| `textOffsetFromAccepted` | re-export from core | | todo | |
| `ToolbarActionProps` | interface | `DocxToolbarActionComponent` | ported | The `select` emit is the `select` output. |
| `ToolbarAlignmentComponent` | interface | `DocxToolbarAlignmentComponent` | ported | `div[docxToolbarAlignment]`. |
| `ToolbarButtonProps` | interface | `DocxToolbarButtonComponent` | ported | The slot id is the `docxToolbarButton` input. The deprecated `slot` prop has no Angular counterpart. Content in the button replaces the registry icon. |
| `ToolbarChartInsert` | const | | todo | |
| `toolbarCommandState` | re-export from core | `injectEditorCommand` | ported | Angular does not re-export engine functions, so the engine stays one peer copy. `injectEditorCommand` reads `toolbarCommandState`; a host can import it from `@docx-editor.dev/core/editor`. |
| `ToolbarCommandState` | re-export from core | `EditorCommandState` | ported | See `toolbarCommandState`. |
| `ToolbarContext` | const | `DOCX_TOOLBAR_PLACEMENT`, `injectToolbarLabel` | ported | The toolbar provides its label function and the panel placement through Angular DI. The context object itself is internal (`DocxToolbarState`). |
| `ToolbarContextValue` | interface |  | dropped | Dropped: the Vue context value has `t` and `onSave`. Angular reads `t` with `injectToolbarLabel`; the save part ports with its menu in phase 5. |
| `ToolbarImageProperties` | const | | todo | |
| `ToolbarPartComponent` | interface |  | dropped | Dropped: the `docxSlot` static marker is a Vue mechanism. Angular reads the slot from the `docxSlot` template directive and the part selector. |
| `ToolbarPartProps` | type | `DocxToolbarButtonComponent` | ported | The props of a named part are the inputs of `button[docxToolbarButton]`. `hidden` is a `docxSlotHidden` template, and `asChild` is the host's own element with the part selector. |
| `ToolbarReviewersProps` | type | `DocxToolbarReviewersComponent` | ported | Inputs `className` and `hidden`. Content replaces the icon, in place of the `icon` prop. |
| `ToolbarSeparatorProps` | interface | `DocxToolbarSeparatorComponent` | ported |  |
| `ToolbarSlotPartComponent` | interface |  | dropped | Dropped: the `docxSlot` static marker is a Vue mechanism. Angular reads the slot from the `docxSlot` template directive and the part selector. |
| `ToolbarSlotPartProps` | interface | `DocxSlotDirective` | ported | `hidden` is `docxSlotHidden`. `class` is the class of the part element. |
| `ToolbarTranslate` | type | `DocxToolbarTranslate` | ported |  |
| `TranslationKey` | re-export from core | | todo | |
| `UnsupportedElement` | re-export from core | | todo | |
| `UnsupportedElementLayout` | re-export from core | | todo | |
| `UnsupportedElementRule` | re-export from core | | todo | |
| `useBookmarkDialog` | function | | todo | |
| `UseBookmarkDialogReturn` | interface | | todo | |
| `useBookmarks` | function | | todo | |
| `UseBookmarksReturn` | interface | | todo | |
| `useChromeDir` | function | `injectChromeDir` | ported |  |
| `useChromeTranslate` | function | | todo | |
| `useCommentDraftRequest` | function | | todo | |
| `useCompareDocumentsDialog` | function | | todo | |
| `UseCompareDocumentsDialogReturn` | interface | | todo | |
| `useContentControl` | function | | todo | |
| `useContentControlInstance` | function | | todo | |
| `UseContentControlResult` | interface | | todo | |
| `useContentControlWidget` | function | | todo | |
| `UseContentControlWidgetResult` | interface | | todo | |
| `useContextMenuAnchor` | function | `injectContextMenuAnchor` | ported |  |
| `useContextMenuHighlights` | function | `injectContextMenuHighlights` | ported |  |
| `useContextMenuTarget` | function | `injectContextMenuTarget` | ported |  |
| `UseDialogReturn` | interface | `DocxDialogState` | ported | Signals in the place of refs. `injectDialog()` reads it inside a dialog. |
| `useDocumentOutline` | function | | todo | |
| `UseDocumentOutlineResult` | interface | | todo | |
| `useDocumentSearch` | function | | todo | |
| `UseDocumentSearchOptions` | interface | | todo | |
| `UseDocumentSearchResult` | interface | | todo | |
| `useDocxEditor` | function | `injectDocxEditor` | ported |  |
| `useDocxSource` | function | `injectDocxSource` | ported |  |
| `UseDocxSourceOptions` | interface | `DocxSourceOptions` | ported |  |
| `UseDocxSourceResult` | interface | `DocxSourceState` | ported | The members are signals. |
| `useEditorCaret` | function | | todo | |
| `useEditorColorMode` | function | `injectEditorColorMode` | ported |  |
| `UseEditorColorModeReturn` | interface | `EditorColorModeState` | ported |  |
| `useEditorCommand` | function | `injectEditorCommand` | ported |  |
| `useEditorEvent` | function | `injectEditorEvent` | ported |  |
| `useEditorHostDocument` | function | `injectEditorHostDocument` | ported |  |
| `UseEditorHostDocumentReturn` | type | `injectEditorHostDocument` | ported | The function returns `Signal<Document | null>`. |
| `useEditorSnapshot` | function | | todo | |
| `useEditorState` | function | `injectEditorState` | ported |  |
| `useEditorValueCommand` | function | `injectEditorValueCommand` | ported |  |
| `useFontFamily` | function | `injectFontFamily` | ported |  |
| `UseFontFamilyResult` | interface | `FontFamilyState` | ported |  |
| `useFonts` | function | `injectFonts` | ported | It needs no injection context. |
| `useGoToBookmark` | function | | todo | |
| `UseGoToBookmarkOptions` | interface | | todo | |
| `UseGoToBookmarkReturn` | interface | | todo | |
| `useHeaderFooterState` | function | | todo | |
| `useHighlightAt` | function | | todo | |
| `useHighlights` | function | | todo | |
| `UseHighlightsReturn` | interface | | todo | |
| `useHistoryGroup` | function | `injectHistoryGroup` | ported |  |
| `UseHistoryGroupReturn` | interface | `HistoryGroupState` | ported |  |
| `useHyperlinkPopup` | function | `injectHyperlinkPopup` | ported | Inside a root, the caller counts as a view of the shared popup until Angular destroys it. |
| `useHyperlinkPopupInstance` | function | `injectHyperlinkPopupInstance` | ported | |
| `UseHyperlinkPopupResult` | interface | `HyperlinkPopup` | ported | `state` is a signal. |
| `useNavigationPane` | function | | todo | |
| `UseNavigationPaneOptions` | interface | | todo | |
| `UseNavigationPaneResult` | interface | | todo | |
| `useNavigationShift` | function | | todo | |
| `useNotePropertiesState` | function | | todo | |
| `useNoteScopeState` | function | | todo | |
| `useObjectControl` | function | | todo | |
| `UseObjectControlReturn` | type | | todo | |
| `useOpenFailure` | function | | todo | |
| `UseOpenFailureOptions` | interface | | todo | |
| `UseOpenFailureReturn` | interface | | todo | |
| `usePageSetup` | function | `injectPageSetup` | ported |  |
| `usePageSetupDialog` | function | `injectPageSetupDialog` | ported |  |
| `UsePageSetupDialogReturn` | interface | `DocxDialogState` | ported | `DocxDialogState<PageSetupDialogFields>`. |
| `UsePageSetupReturn` | interface | `PageSetupState` | ported | Signals in the place of computed refs. |
| `usePaneRequest` | function | | todo | |
| `useParagraphDialog` | function | `injectParagraphDialog` | ported |  |
| `UseParagraphDialogReturn` | interface | `ParagraphDialogState` | ported | `mixed` is a signal. |
| `useParagraphFormat` | function | `injectParagraphFormat` | ported |  |
| `UseParagraphFormatReturn` | interface | `ParagraphFormatState` | ported | Signals in the place of computed refs. |
| `useParagraphIndent` | function | | todo | |
| `UseParagraphIndentReturn` | interface | | todo | |
| `useParagraphStyle` | function | `injectParagraphStyle` | ported |  |
| `UseParagraphStyleResult` | interface | `ParagraphStyleState` | ported |  |
| `useReviewAuthors` | function | | todo | |
| `useReviewGutter` | function | | todo | |
| `useRevisionMarkupDialog` | function | | todo | |
| `UseRevisionMarkupDialogReturn` | interface | | todo | |
| `useScopeClassName` | function | `injectScopeClassName` | ported |  |
| `useScopedChromeAnchor` | function | | todo | |
| `useSelectedChart` | function | | todo | |
| `UseSelectedChartReturn` | interface | | todo | |
| `useSelectedShape` | function | | todo | |
| `UseSelectedShapeReturn` | interface | | todo | |
| `useSplitCellDialog` | function | | todo | |
| `UseSplitCellDialogReturn` | interface | | todo | |
| `useTableBorderTargetLabel` | function | `injectTableBorderTargetLabel` | ported |   |
| `useTextFormFieldDialog` | function | | todo | |
| `UseTextFormFieldDialogReturn` | interface | | todo | |
| `useToolbarContext` | function | `injectToolbarLabel` | ported | The only public reading of the toolbar context is its label function. |
| `useToolbarLabel` | function | `injectToolbarLabel` | ported |  |
| `useToolbarLabelFor` | function | `injectToolbarLabel` | ported | A host that needs its own translation passes the toolbar `t` input. |
| `useTranslation` | function | `injectTranslation` | ported | Returns `DocxEditorTranslation`. |
| `useWatermark` | function | `injectWatermark` | ported |  |
| `useWatermarkDialog` | function | `injectWatermarkDialog` | ported |  |
| `UseWatermarkDialogReturn` | interface | `DocxDialogState` | ported | `DocxDialogState<WatermarkDialogFields>`. |
| `UseWatermarkReturn` | interface | `WatermarkCommandState` | ported | The name avoids the core type `WatermarkState`. |
| `useZoom` | function | `injectZoom` | ported |  |
| `UseZoomResult` | interface | `ZoomState` | ported |  |
| `VERSION` | const | | todo | |
| `VerticalRuler` | const | | todo | |
| `VerticalRulerProps` | interface | | todo | |
| `Watermark` | re-export from core | `Watermark` | ported |  |
| `WatermarkDialogFieldName` | type | `WatermarkDialogFieldName` | ported |  |
| `WatermarkDialogFields` | interface | `WatermarkDialogFields` | ported |  |
| `WatermarkDialogKind` | type | `WatermarkDialogKind` | ported |  |
| `WatermarkDialogLayout` | type | `WatermarkDialogLayout` | ported |  |
| `WatermarkState` | re-export from core | `WatermarkState` | ported |  |
| `WORD_DEFAULT_FONT` | re-export from core | | todo | |

## Tests

87 rows.

| Vue test | Angular spec | Status | Notes |
| --- | --- | --- | --- |
| `packages/vue/test/bookmark-dialog.test.ts` | | todo | Task 6.4 |
| `packages/vue/test/build-external.test.ts` | | todo | |
| `packages/vue/test/chrome-composition.test.ts` | `packages/angular/src/lib/toolbar/toolbar.spec.ts`, `packages/angular/src/lib/toolbar/toolbar-paragraph.spec.ts`, `packages/angular/src/lib/toolbar/toolbar-pickers.spec.ts`, `packages/angular/src/lib/toolbar/toolbar-table.spec.ts` | todo | Ported (task 4.1): the default bar follows the registry with slot ids; a hidden slot leaves the bar; a slot override keeps registry order and its class; `preset=false` renders only the composed content; Bold reflects live engine state; Escape closes the More panel before the header scope and the format painter; More closes when focus leaves. Ported (tasks 4.2 to 4.4): the contextual table chrome shows in a table only, collapses before the ordinary groups, leads the More panel, and its menu takes its own Escape; the line-spacing menu ticks the spacing in force, also one that no preset matches; a dropdown inside More takes the first Escape and the panel the second; Escape from a host input closes the editing-mode menu and keeps its default. Dropped: the deprecated `slot` prop of `Toolbar.Button`, which Angular never had. Ported (task 6.3): More stays open behind the paragraph dialog that it opened, and the focus comes back to the line spacing trigger. The reviewer-icon claim ports with task 10.1. The menu, context menu, navigation, hyperlink, loading, image, content control, header and footer, and font notice claims port with phases 5 to 9. |
| `packages/vue/test/chrome-popup-dismissal.test.ts` | `packages/angular/src/lib/chrome-popup-dismissal.spec.ts` | ported | Keeps: for the menu bar and the reviewers menu, an Escape from the pages closes the popup and stops there, and an Escape from a host input closes it and keeps its default; focus that moves to another control closes the menu bar and a picker, and the Escape then reaches that control; focus that moves into a host's modal dialog closes a picker. Added: the reviewers menu lists the authors, and the arrow keys and Home move between its rows. |
| `packages/vue/test/content-control-widget-parts.test.ts` | | todo | Task 9.3 |
| `packages/vue/test/content-lock-policy-prop.test.ts` | `packages/angular/src/lib/live-inputs.spec.ts` | ported | Keeps: the root and `<docx-editor>` pass the policy, so bold over a locked control succeeds only with it and the text stays; an equal new object does not warn; a changed value is ignored and warns once. |
| `packages/vue/test/context-menu-add-comment.test.ts` | | todo | |
| `packages/vue/test/context-menu-chart.test.ts` | `packages/angular/src/lib/context-menu/context-menu.spec.ts` | ported |  |
| `packages/vue/test/context-menu-keyboard.test.ts` | `packages/angular/src/lib/context-menu/context-menu-keyboard.spec.ts` | ported | Keeps: a keyboard-opened menu opens below the caret and follows the caret; a right-click opens at the pointer and names its target; a row that stays mounted follows a keyboard reopen of an open menu; with no painted caret, the menu opens at the document corner. Changed: Angular creates the rows in `docx-context-menu-start` once with the menu, so the probe row reads each open from `injectContextMenuAnchor` and `injectContextMenuTarget`, and the spec counts its creations. |
| `packages/vue/test/context-menu-link.test.ts` | `packages/angular/src/lib/context-menu/context-menu-link.spec.ts`, `packages/angular/src/lib/hyperlink/link-row-chrome.spec.ts` | ported | Keeps (task 5.2): the row is enabled and opens the registered link UI; in viewing mode it is disabled with the engine reason; a `text.link` slot template replaces the row; on a link the row says Edit link. Keeps (task 6.2): on a link the row opens the edit fields with the address; Escape, Enter, and Cancel close the popup and give the focus back to the document, and Escape keeps the selection. Changed: jsdom cannot focus the pages, so the specs count the engine's `focus` calls. Added: an empty address refuses Apply with an alert, and a press outside closes the popup. |
| `packages/vue/test/context-menu-placement.test.ts` | `packages/angular/src/lib/context-menu/context-menu-placement.spec.ts` | ported |  |
| `packages/vue/test/context-menu-start.test.ts` | `packages/angular/src/lib/context-menu/context-menu-target.spec.ts` | ported | Angular has no fragments: the spec checks that each start moves up and that the host rows stay below. The menu adds one separator below all starts. |
| `packages/vue/test/context-menu-target.test.ts` | `packages/angular/src/lib/context-menu/context-menu-target.spec.ts`, `packages/angular/src/lib/hyperlink/link-row-chrome.spec.ts` | todo | Ported (task 5.2): the pages, the canvas, painted text, a painted drop-down field, and the content mount open the menu; a form control beside the pages keeps the browser menu, its focus, and its selection; a right-click on the open menu keeps it in place. Ported (task 6.2), in `packages/angular/src/lib/hyperlink/link-row-chrome.spec.ts`: the link popup renders inside the content mount; a right-click in its address field, on the panel, or on its buttons keeps the browser menu, the focus, the draft, and the selection. The dialog claims port with task 6.4. |
| `packages/vue/test/context-menu.test.ts` | `packages/angular/src/lib/context-menu/context-menu.spec.ts` | todo | Ported (task 5.2): a right-click opens the panel at the pointer; engine changes keep the one listener; the default rows; Paste closes before the clipboard read settles; a refused read disables Paste with the reason; the form field row opens the field dialog; the table of contents rows and Remove. Added: the table rows, Insert row below, and Split cells. The claim that the dialog saves the field ports with task 6.4. |
| `packages/vue/test/deferred-notifier.test.ts` | | todo | Task 10.5 |
| `packages/vue/test/dialog-customization.test.ts` | `packages/angular/src/lib/dialog/dialog-parts.spec.ts` | ported | Keeps, on a spec dialog with the Field Options session contract: a custom Apply keeps the field wiring and closes only after an accepted write; a refused write stays open with an alert, and an abort closes the dialog; host classes merge with the packaged classes, and `asChild` keeps the packaged classes off the host element; a part forwards host attributes but keeps its own wiring; replacement content keeps its markup, a live label, the element, and the wiring. Changed: a host's own `(click)` on its element runs beside the part's handler, because Angular adds listeners and cannot drop them; the attribute `onclick` does not apply. Dropped: the SFC `children` forwarding claim (Vue-only mechanism). Added: a hidden part, `keepContent`, a component in a part template that reads `injectDialog()`, `preset` false, and the native dialog's focus, Escape, and scope class. |
| `packages/vue/test/document-refresh-edit.test.ts` | | todo | Task 10.3 |
| `packages/vue/test/editing-mode-author.test.ts` | `packages/angular/src/lib/construction-inputs.spec.ts`, `packages/angular/src/lib/toolbar/toolbar-pickers.spec.ts` | ported | Keeps: `<docx-editor mode='suggesting'>` with the review module waits for an author, then records typing as tracked changes by that author; the open editing mode menu tracks author arrival and removal without a mode change; keyboard travel skips Suggesting without an author. |
| `packages/vue/test/editor-accessibility.test.ts` | | todo | Task 10.5 |
| `packages/vue/test/editor-composition.test.ts` | | todo | Task 10.5 |
| `packages/vue/test/engine-popup-renderers.test.ts` | `packages/angular/src/lib/popups/engine-popups.spec.ts` | todo | Ported (task 6.1), with a spec widget renderer: value widget ownership for the native, custom, and manual cases; checkbox presses reach a renderer only through `contentControlCheckbox`; picture presses reach a renderer only through `contentControlPicture`. Added: a live `false` entry cancels the next session, and a removed entry gives the presses back to the engine. The claims about the packaged widget's own UI (date normalization, focus cleanup, the dropdown list, the gallery) port with task 9.3, and the invalid-field acknowledgement with task 6.4. |
| `packages/vue/test/equation-popup.test.ts` | | todo | Task 6.6 |
| `packages/vue/test/export-popup.test.ts` | | todo | Task 6.5 |
| `packages/vue/test/field-results-prop.test.ts` | `packages/angular/src/lib/live-inputs.spec.ts` | ported | Keeps: the root and `<docx-editor>` pass `fieldResults`; `'editable'` types inside the saved result, the default types after it. A later change is ignored and warns once. |
| `packages/vue/test/font-origins.test.ts` | `packages/angular/src/lib/fonts.spec.ts` | ported | Keeps: an eager loader holds the document until it settles; a failing loader releases it and warns; eager origins compose first-wins and one rejection keeps the others; an on-demand resolver releases the document at once behind one stable identity that reads the latest origins; one resolver makes a list on demand, and a loader in it gets no argument; `injectFonts` gives the request to marked and unmarked resolvers of every arity, composes resolvers, returns a marked resolver, reports a function that returns an origin, reads `computed` and signal origins again on each resolve. |
| `packages/vue/test/font-shaping-lifecycle.test.ts` | `packages/angular/src/lib/construction-inputs.spec.ts` | ported | Keeps: an equal new `fontShaping` keeps edits, selection, and undo; a changed `maxFontBytes` creates a new engine, and later equal values keep that one. |
| `packages/vue/test/fonts-prop.test.ts` | `packages/angular/src/lib/construction-inputs.spec.ts` | ported | Keeps: a new `fonts` value on the root and on `<docx-editor>` reaches the same engine and session through `setFonts`. |
| `packages/vue/test/form-field-i18n.test.ts` | `packages/angular/src/lib/i18n.spec.ts`, `packages/angular/src/lib/live-inputs.spec.ts` | todo | Ported: engine form labels follow the catalog and an explicit `translate` without a new engine, on the root and on `<docx-editor>`. The text form field dialog claims (the open dialog keeps its input and focus) port with task 6.4. |
| `packages/vue/test/header-footer-chrome.test.ts` | | todo | Task 9.1 |
| `packages/vue/test/highlights-background-layout.test.ts` | | todo | Task 7.2 |
| `packages/vue/test/host-window-chrome.test.ts` | `packages/angular/src/lib/host-window.spec.ts`, `packages/angular/src/lib/toolbar/toolbar-pickers.spec.ts` | ported | Keeps: the window panel renders in the pop-out while open and comes back on close; turning `open` off moves the panel back before the window closes; a picker in the pop-out closes on a press in its own window, not in the opener. |
| `packages/vue/test/hydration-renderer-slots.test.ts` | | todo | Task 10.5 |
| `packages/vue/test/hyperlink-as-child.test.ts` | `packages/angular/src/lib/hyperlink/hyperlink.spec.ts` | ported | Keeps, over a stub popup: `asChild` keeps the packaged class off the host element and the wiring on it, and `className` reaches it; `asChild` with no `className` leaves only the host's class; without `asChild` the packaged part keeps its class beside `className`. Added: `keepContent` keeps the packaged icon. |
| `packages/vue/test/image-authoring.test.ts` | | todo | Task 8.1 |
| `packages/vue/test/image-decode-port-prop.test.ts` | `packages/angular/src/lib/construction-inputs.spec.ts` | ported | Keeps: the root and `<docx-editor>` decode the picture with the port and keep its bytes on save; a new port creates a new engine that uses it; destroying the editor aborts a running decode. |
| `packages/vue/test/image-insert-refusal.test.ts` | | todo | Task 8.1 |
| `packages/vue/test/image-overlay-history.test.ts` | | todo | Task 8.1 |
| `packages/vue/test/image-panel-rows.test.ts` | | todo | Task 8.1 |
| `packages/vue/test/lifecycle-regression.test.ts` | `packages/angular/src/lib/packaged-api.spec.ts`, `packages/angular/src/lib/fonts.spec.ts`, `packages/angular/src/lib/root.directive.spec.ts` | todo | Ported: `ready` fires after pages paint; the `'blank'` source opens without warnings; a new inline `translate` keeps the engine and undo; an unrelated render keeps the engine; `ready`, `change`, and `fontError` emit once; destroy releases the engine; `injectDocxSource` follows a byte signal and aborts a stale fetch; `injectFonts` reads the current origin. Dropped: the `provideDocxEditor` claims (Vue-only mechanism). The hyperlink popup state claim is in `packages/angular/src/lib/hyperlink/hyperlink.spec.ts` (task 6.2). The search debounce and navigation observer claims with task 7.1. |
| `packages/vue/test/link-row-chrome.test.ts` | `packages/angular/src/lib/hyperlink/link-row-chrome.spec.ts` | ported | Keeps: with the popup, the toolbar button and the context menu row are enabled; without a link UI, both are disabled with the engine reason and open nothing; a host's own link UI enables both and they open it, and its removal disables them again; a host part that calls `injectHyperlinkPopup()` keeps the controls working; a caret inside a word opens the popup with that word, and Apply links it. Changed: the specs compose the root with `<docx-editor-hyperlink>` in the place of the packaged `hyperlinkPopup` flag, which lands with task 6.7. |
| `packages/vue/test/live-props.test.ts` | `packages/angular/src/lib/live-inputs.spec.ts` | todo | Ported: author, mode, translate, and locale apply to the same engine; another input change does not restore the host mode; a catalog change and then `translate` reach the engine; an equal new `translate` publishes no selection change; `colorMode` applies. The `title` claim is in `packages/angular/src/lib/packaged-menus.spec.ts`. The `rulers` and `navigation` claims port with tasks 9.6 and 7.3. |
| `packages/vue/test/loading-continued.test.ts` | | todo | Task 9.5 |
| `packages/vue/test/loading-indicator.test.ts` | | todo | Task 9.5 |
| `packages/vue/test/locked-control-navigation-prop.test.ts` | `packages/angular/src/lib/live-inputs.spec.ts` | ported | Keeps: the root and `<docx-editor>` pass the mode; `'atomic'` steps over the control, the default steps into it. A later change is ignored and warns once, only for a real change. |
| `packages/vue/test/menu-composition.test.ts` | `packages/angular/src/lib/menu/menu-bar.spec.ts` | ported | Keeps: the default bar renders every registry menu in order; a cancelled press on the pages closes the open menu; a host save handler runs from File › Save; a row override reaches the export submenu without a duplicate row; direct bar actions stay tabbable and follow command state; right-to-left bar arrows follow the visual order and the direction change. Changed: the Page Setup popup claims are in `packages/angular/src/lib/page-setup/page-setup-dialog.spec.ts` (task 6.3): File › Page setup opens the packaged dialog, a `pageSetup` registry component replaces it, and `false` opens nothing. Moved: the export dialog claims (missing converters, Markdown export dialog) port with task 6.5. Added: keyboard open and Escape, engine state on rows, Insert › Table, Review › Markup options, host menus, and the "⋯" menu. |
| `packages/vue/test/menu-download-name.test.ts` | `packages/angular/src/lib/menu/download.spec.ts` | ported | Keeps: a long run of dots and spaces stays linear; device names with an extension and a padded suffix; truncation leaves no trailing dot or space. |
| `packages/vue/test/menu-opening.test.ts` | `packages/angular/src/lib/menu/menu-bar.spec.ts` | ported | Keeps: the bar closes and its triggers are disabled while the next document opens, and they come back after the open. |
| `packages/vue/test/merge-arrangement.test.ts` | | todo | Task 10.5 |
| `packages/vue/test/navigation-goto.test.ts` | | todo | Task 7.1 |
| `packages/vue/test/navigation-heading-page-break.test.ts` | | todo | Task 7.1 |
| `packages/vue/test/navigation-pane-host.test.ts` | | todo | Task 7.1 |
| `packages/vue/test/notes-context-menu.test.ts` | | todo | Task 9.2 |
| `packages/vue/test/open-failure-part.test.ts` | | todo | Task 9.5 |
| `packages/vue/test/package-dependencies.test.ts` | | todo | |
| `packages/vue/test/package-exports.test.ts` | | todo | |
| `packages/vue/test/package-version.test.ts` | | todo | |
| `packages/vue/test/page-number-opening.test.ts` | | todo | Task 9.4 |
| `packages/vue/test/page-number.test.ts` | | todo | Task 9.4 |
| `packages/vue/test/paginated-form-save.test.ts` | `packages/angular/src/lib/packaged-api.spec.ts` | ported | Keeps: `save()` refuses (`invalidArgs`) a form field value that its format does not accept, and saves the value again after undo. The Angular spec uses `<docx-editor>`, because `PaginatedDocxEditor` is not ported. |
| `packages/vue/test/paragraph-dialog.test.ts` | `packages/angular/src/lib/paragraph/paragraph-dialog.spec.ts`, `packages/angular/src/lib/toolbar/toolbar-paragraph.spec.ts` | ported | Keeps: the line spacing menu opens the dialog (a click test in the place of the source read); the dialog seeds every field from the selection; a mixed flag is indeterminate and survives OK; it opens focused, and Escape closes it; every mixed field can be resolved, and a mixed line spacing locks At until a rule is picked; only a changed field is written; an asChild Apply keeps the command and the packaged fields; a custom layout keeps Apply, places a packaged field (`docxDialogOutlet` for `Field`), and sets a value through `injectParagraphDialog`; a custom layout reads and sets the direction; a draft cannot apply to a replacement document; the direction field seeds and writes, opens blank when mixed, and stays mixed if untouched. |
| `packages/vue/test/paragraph-hooks-lock-state.test.ts` | `packages/angular/src/lib/paragraph/paragraph-format.spec.ts` | todo | Ported (task 6.3): `injectParagraphFormat().isEnabled` follows the caret out of and back into a content lock, a data binding, and both, and viewing mode takes it away. The `useParagraphIndent` half ports with the rulers in task 9.4. |
| `packages/vue/test/paragraph-style-labels.test.ts` | `packages/angular/src/lib/toolbar/toolbar-pickers.spec.ts` | ported | Keeps: built-in styles take the catalog label (also a localized style id such as `Ttulo1`), and the document's own styles keep their name. |
| `packages/vue/test/popup-config.test.ts` | `packages/angular/src/lib/popups/popup-config.spec.ts` | todo | Ported (task 6.1): the root hosts only configured popups, inside the content, and honors live `false` entries; components keep their state across new registry objects and stay with their editor; a new session resets a draft, and a new entry object with the same component keeps it. Added: static inputs of an entry, and `DOCX_EDITOR_POPUPS` under the root's entries. The legacy flag, packaged hyperlink, and context menu translation claims port with task 6.7, and the note popup claim with task 9.2. |
| `packages/vue/test/print-popup.test.ts` | | todo | Task 6.5 |
| `packages/vue/test/props-primitives.test.ts` | | todo | Task 10.5 |
| `packages/vue/test/review-gutter.test.ts` | | todo | Task 10.1 |
| `packages/vue/test/review-overflow-navigation.test.ts` | | todo | Task 10.2 |
| `packages/vue/test/revision-markup.test.ts` | | todo | Task 10.2 |
| `packages/vue/test/root-color-mode.test.ts` | `packages/angular/src/lib/color-mode.spec.ts`, `packages/angular/src/lib/host-window.spec.ts` | ported | Keeps: 'dark' puts `.dark` beside `.docx-editor`; omitted and 'light' stay light; a change applies without new elements; 'system' follows the OS setting and stops when it leaves 'system'; the pop-out follows the root; `injectEditorColorMode` and `injectScopeClassName` report the theme. The toolbar and menu bar check their own scope classes in tasks 4.1 and 5.1. |
| `packages/vue/test/runtime-audit.test.ts` | | todo | Task 10.5 |
| `packages/vue/test/scope-dispose.test.ts` | `packages/angular/src/lib/color-mode.spec.ts` | ported | Keeps: mounting the composed editor logs no warning and no error. |
| `packages/vue/test/scoped-chrome-anchor.test.ts` | | todo | Task 10.5 |
| `packages/vue/test/security.test.ts` | `packages/angular/src/lib/hyperlink/hyperlink.spec.ts` | todo | Ported (task 6.2): `sanitizeHref` refuses `javascript:` targets; the popup shows a `javascript:` link as inert, with no Copy action, and following it opens nothing. Added: an `http` link is a button that opens its address. The `v-html` scan claim ports with task 11.1, with the lint fixtures. |
| `packages/vue/test/selection-range-state.test.ts` | | todo | Task 10.5 |
| `packages/vue/test/slot-ref.test.ts` | | todo | Task 10.5 |
| `packages/vue/test/sugar-host.test.ts` | `packages/angular/src/lib/packaged-api.spec.ts` | todo | Ported: `ready` emits once and the methods exist; `change` and `fontError` emit. Not kept: a new `document` value keeps the engine and calls `load()`, which the Angular browser test checks. The menu, title, save, open, context menu, and chrome band claims are in `packages/angular/src/lib/packaged-menus.spec.ts`. The loading, ruler, and navigation claims port with tasks 9.5, 9.6, and 7.3. |
| `packages/vue/test/text-highlights.test.ts` | | todo | Task 7.2 |
| `packages/vue/test/toolbar-dev-warnings.test.ts` | `packages/angular/src/lib/toolbar/toolbar-pickers.spec.ts` | ported | Keeps: a development warning prints once. Dropped: the Vite and Node build signals. Angular decides development mode with `isDevMode()`, which the Angular build sets. |
| `packages/vue/test/toolbar-direction.test.ts` | `packages/angular/src/lib/i18n.spec.ts`, `packages/angular/src/lib/live-inputs.spec.ts`, `packages/angular/src/lib/toolbar/toolbar-link-direction.spec.ts` | ported | Keeps: the provider `dir` wins over the catalog; a right-to-left catalog mirrors the composed parts; a left-to-right catalog leaves the host direction; the packaged editor mirrors for Arabic and Hebrew and not for Portuguese; Right-to-left writes `w:bidi` and reads back pressed; the pressed state follows the caret; the default toolbar omits the direction buttons for a left-to-right language and shows them after alignment for a right-to-left one. |
| `packages/vue/test/toolbar-host-groups.test.ts` | `packages/angular/src/lib/toolbar/toolbar.spec.ts`, `packages/angular/src/lib/toolbar/toolbar-table.spec.ts` | todo | Ported (task 4.1): controls leave one at a time from the end of a group and More holds exactly those; a host group is measured, collapses, and shows its actions as panel rows; `after`, built-in ids, hidden groups, and all-hidden groups; label, labelKey, and id headings; `className` in the preset and alone; slot replacement in place with its panel content; a content-sized bar gets its groups back; the panel clamp, also right-to-left; `overflowContent` on a built-in group; the anchor, slot id, built-in setting, and reserved id warnings. Not kept: the template-style empty `hidden` attribute, which `booleanAttribute` covers. Ported (task 4.4): the table part opens a size grid and inserts the picked size; a second trigger click closes the grid and focuses the document; focus that leaves the grid closes it; the grid moves left to stay in the viewport; in More, Escape closes the grid and an insert closes the panel. The Add Comment claims port with task 10.1, the menu bar claims with task 5.1. |
| `packages/vue/test/toolbar-link-platform.test.ts` | `packages/angular/src/lib/toolbar/toolbar-link-direction.spec.ts` | ported | Keeps: the link control names the Mac modifier like the other buttons, keeps Ctrl on Windows, and normalizes a host translation; it sends the link request to the registered link UI, and without one it is disabled with the engine reason. Dropped: the server-render hydration claim. Angular hydration applies attribute bindings again on the client, so a label that differs between the server and a Mac client cannot cause a mismatch. |
| `packages/vue/test/toolbar-picker-accessibility.test.ts` | `packages/angular/src/lib/toolbar/toolbar-pickers.spec.ts`, `packages/angular/src/lib/toolbar/toolbar-paragraph.spec.ts` | ported | Keeps: the font family, style, and zoom pickers support arrows, Escape, and focus return; the font size input names its open list; the color popups are dialogs that give the focus back on Escape and to the pages after a pick; Escape closes a click-opened picker and stops there, and a host field keeps its Escape; Escape closes a picker before the header scope and the format painter; a canceled press on the pages closes each of the six popups, and a press inside keeps it. |
| `packages/vue/test/unsupported-elements.test.ts` | | todo | Task 9.5 |
| `packages/vue/test/use-comment-draft-request.test.ts` | | todo | Task 10.1 |
| `packages/vue/test/use-editor-command.test.ts` | `packages/angular/src/lib/editor-command.spec.ts` | ported | Keeps: state follows a payload change with the same type; `execute` returns false on a refusal; many commands share the root listeners; custom equality stops slice updates; a burst collapses through the deferred notifier; outside a root the command is disabled; range binding groups keyboard repeats and cleans up; a held stepper starts before its own keydown handler. |
| `packages/vue/test/viewport-rail.test.ts` | | todo | Task 10.1 |
| `packages/vue/test/watermark-dialog.test.ts` | `packages/angular/src/lib/watermark/watermark-dialog.spec.ts` | ported | Keeps: Insert › Watermark opens the dialog, and Apply sets a text watermark in one undo step; it seeds from the existing watermark, and No watermark removes it; Cancel writes nothing; a chosen picture applies as a picture watermark, and the preview follows Auto and a scale; custom presets fill the preset select, and typed text shows Custom; the preview draws the draft and follows every field; Select picture works for any kind and switches to Picture; the first focus lands on the checked kind; an unchanged Apply writes nothing; No watermark on a document without one is no change; a custom layout places the preview part, which follows the draft. Changed: the spec stubs object URLs and `File.arrayBuffer`, which jsdom lacks. Dropped: the `Body` part with children, because an Angular host writes its own layout element. |
