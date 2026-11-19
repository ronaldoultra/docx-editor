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
| `ChromeSlotId` | re-export from core | | todo | |
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
| `ContextMenuAnchor` | interface | | todo | |
| `ContextMenuCellVerticalAlignment` | const | | todo | |
| `ContextMenuCommandProps` | interface | | todo | |
| `ContextMenuContextValue` | interface | | todo | |
| `ContextMenuCopy` | const | | todo | |
| `ContextMenuCopyFormatting` | const | | todo | |
| `ContextMenuCut` | const | | todo | |
| `ContextMenuDelete` | const | | todo | |
| `ContextMenuDeleteTable` | const | | todo | |
| `ContextMenuDeleteTableColumn` | const | | todo | |
| `ContextMenuDeleteTableRow` | const | | todo | |
| `ContextMenuInsertColumnLeft` | const | | todo | |
| `ContextMenuInsertColumnRight` | const | | todo | |
| `ContextMenuInsertRowAbove` | const | | todo | |
| `ContextMenuInsertRowBelow` | const | | todo | |
| `ContextMenuItem` | const | | todo | |
| `ContextMenuItemProps` | interface | | todo | |
| `ContextMenuMergeCells` | const | | todo | |
| `ContextMenuPaste` | const | | todo | |
| `ContextMenuPasteFormatting` | const | | todo | |
| `ContextMenuPasteWithoutFormatting` | const | | todo | |
| `ContextMenuRefreshToc` | const | | todo | |
| `ContextMenuRefreshTocPageNumbers` | const | | todo | |
| `ContextMenuRemoveToc` | const | | todo | |
| `ContextMenuSelectAll` | const | | todo | |
| `ContextMenuSelectTable` | const | | todo | |
| `ContextMenuSplitCell` | const | | todo | |
| `ContextMenuStart` | const | | todo | |
| `ContextMenuStartProps` | interface | | todo | |
| `ContextMenuTableRowProps` | interface | | todo | |
| `createDocumentRefresh` | re-export from core | | todo | |
| `createFontSource` | re-export from core | | todo | |
| `DEFAULT_FONT` | re-export from core | | todo | |
| `DEFAULT_WATERMARK_PRESETS` | const | | todo | |
| `deferFonts` | re-export from core | | todo | |
| `defineFontResolver` | re-export from core | | todo | |
| `definePopup` | function | | todo | |
| `DialogCustomizationProps` | interface | | todo | |
| `DialogPartProps` | interface | | todo | |
| `DocumentOpenFailure` | re-export from core | | todo | |
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
| `DocxEditor` | const | | todo | |
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
| `DocxEditorContextMenu` | const | | todo | |
| `DocxEditorContextMenuNamespace` | interface | | todo | |
| `DocxEditorContextMenuProps` | interface | | todo | |
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
| `DocxEditorHyperLink` | const | | todo | |
| `DocxEditorHyperLinkNamespace` | interface | | todo | |
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
| `DocxEditorMenu` | const | | todo | |
| `DocxEditorMenuNamespace` | interface | | todo | |
| `DocxEditorMenuProps` | interface | | todo | |
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
| `DocxEditorPageSetupDialog` | const | | todo | |
| `DocxEditorPageSetupDialogProps` | interface | | todo | |
| `DocxEditorPaneSections` | const | | todo | |
| `DocxEditorPaneSectionsProps` | interface | | todo | |
| `DocxEditorParagraphDialog` | const | | todo | |
| `DocxEditorParagraphDialogProps` | interface | | todo | |
| `DocxEditorPopup` | type | | todo | |
| `DocxEditorPopups` | interface | | todo | |
| `DocxEditorPrintDialog` | const | | todo | |
| `DocxEditorPrintDialogProps` | interface | | todo | |
| `DocxEditorProps` | interface | | todo | |
| `DocxEditorRef` | interface | | todo | |
| `DocxEditorRefCallback` | type | | todo | |
| `DocxEditorRevisionMarkup` | const | | todo | |
| `DocxEditorRevisionMarkupDialog` | const | | todo | |
| `DocxEditorRevisionMarkupDialogProps` | interface | | todo | |
| `DocxEditorRevisionMarkupProps` | interface | | todo | |
| `DocxEditorRoot` | const | `DocxEditorRootDirective` | ported | Directive `docxEditorRoot`. The props and listeners rows track the inputs and outputs that are still missing. |
| `DocxEditorRootListeners` | interface | `DocxEditorRootDirective` | todo | Has the `ready` and `change` outputs. Missing: `fontError` (3.2), `revisionMarkupChange` (10.4). |
| `DocxEditorRootProps` | interface | `DocxEditorRootDirective` | todo | Has `document`, `mode`, `colorMode`, `locale`, `translate`, `author`. Missing: `fonts`, `fontShaping`, `imageDecodePort`, `modules`, `zoom`, `zoomMode`, `fieldResults`, `contentLockPolicy`, `lockedControlNavigation` (tasks 3.1, 3.2), `popups` (6.1), `revisionMarkup`, `reviewDisplayMode` (10.4), `tableInteractionLabel` (4.4). |
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
| `DocxEditorToolbar` | const | | todo | |
| `DocxEditorToolbarGroupProps` | interface | | todo | |
| `DocxEditorToolbarNamespace` | interface | | todo | |
| `DocxEditorToolbarProps` | interface | | todo | |
| `DocxEditorToolbarSlotProps` | interface | | todo | |
| `DocxEditorUnsupportedElement` | const | | todo | |
| `DocxEditorUnsupportedElementProps` | interface | | todo | |
| `DocxEditorVerticalRuler` | const | | todo | |
| `DocxEditorViewport` | const | `DocxEditorViewportComponent` | ported |  |
| `DocxEditorViewportProps` | interface | `DocxEditorViewportComponent` | ported | Angular hosts set `class` and `style` on the host element. Children are content projection. |
| `DocxEditorWatermarkDialog` | const | | todo | |
| `DocxEditorWatermarkDialogProps` | interface | | todo | |
| `DocxEditorWindow` | const | `DocxEditorWindowComponent` | ported |  |
| `DocxEditorWindowProps` | interface | `DocxEditorWindowComponent`, `DocxEditorWindowStyles` | ported | `onClose` is the `close` output. `open` has two-way binding through `openChange`. |
| `DocxFontOrigin` | type | `DocxFontOrigin` | ported |  |
| `DocxFontsInput` | type | `DocxFontsInput` | ported |  |
| `DocxFontsSource` | type | `DocxFontsSource` | ported |  |
| `DocxSource` | type | `DocxSource` | ported |  |
| `DrawingGraphicKind` | re-export from core | | todo | |
| `DrawingSupport` | re-export from core | | todo | |
| `DrawingSupportIssue` | re-export from core | | todo | |
| `Editor` | re-export from core | | todo | |
| `EditorCaret` | interface | | todo | |
| `EditorCommand` | re-export from core | | todo | |
| `EditorCommandState` | interface | `EditorCommandState` | ported |  |
| `EditorFontError` | re-export from core | | todo | |
| `EditorFontErrorCode` | re-export from core | | todo | |
| `EditorMode` | type | | todo | |
| `EditorQuery` | re-export from core | | todo | |
| `EditorScope` | re-export from core | | todo | |
| `EditorSnapshot` | re-export from core | | todo | |
| `EditorValueCommandState` | interface | `EditorValueCommandState` | ported |  |
| `FontConfiguration` | re-export from core | | todo | |
| `FontConfigurationBase` | re-export from core | | todo | |
| `FontConfigurationFragment` | re-export from core | | todo | |
| `FontFaceRequest` | re-export from core | | todo | |
| `FontFamilyItemProps` | interface | | todo | |
| `FontFamilyNamespace` | interface | | todo | |
| `FontFamilyPartProps` | interface | | todo | |
| `FontFamilyProps` | interface | | todo | |
| `FontLoadFailure` | re-export from core | | todo | |
| `FontLoadFailureReason` | re-export from core | | todo | |
| `FontOption` | interface | | todo | |
| `FontOrigin` | re-export from core | | todo | |
| `FontResolutionRequest` | re-export from core | | todo | |
| `FontResolver` | re-export from core | | todo | |
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
| `HyperLinkActionProps` | interface | | todo | |
| `HyperLinkPartProps` | interface | | todo | |
| `HyperlinkPopupAnchor` | interface | | todo | |
| `HyperlinkPopupMode` | type | | todo | |
| `HyperlinkPopupState` | interface | | todo | |
| `HyperLinkProps` | interface | | todo | |
| `ImageAltText` | const | | todo | |
| `ImageInsertProvider` | const | | todo | |
| `ImageInsertTrigger` | const | | todo | |
| `ImagePropertiesTrigger` | const | | todo | |
| `ImagePropertiesTriggerProps` | interface | | todo | |
| `ImageWrap` | const | | todo | |
| `ImageWrapTarget` | re-export from core | | todo | |
| `IndentUpdate` | interface | | todo | |
| `isFieldLink` | function | | todo | |
| `isFontResolver` | re-export from core | | todo | |
| `ListState` | interface | | todo | |
| `ListType` | type | | todo | |
| `loadFonts` | re-export from core | | todo | |
| `LoadFontsRequest` | re-export from core | | todo | |
| `LoadFontsResult` | re-export from core | | todo | |
| `LOADING_SNAPSHOT` | re-export from core | | todo | |
| `LocaleProvider` | const | `DocxEditorI18nDirective`, `provideDocxEditorI18n` | ported |  |
| `LocaleProviderProps` | interface | `DocxEditorI18nDirective`, `DocxEditorI18nOptions` | ported | `i18n` is the `docxEditorI18n` input, `dir` is the `docxEditorDir` input. |
| `MarkedFontResolver` | re-export from core | | todo | |
| `MAX_RESOLVER_FAMILIES` | re-export from core | | todo | |
| `MaybeRefOrGetter` | type | | todo | |
| `MenuActionProps` | interface | | todo | |
| `MenuGroupProps` | interface | | todo | |
| `MenuId` | type | | todo | |
| `MenuItemBaseProps` | interface | | todo | |
| `MenuItemProps` | interface | | todo | |
| `MenuPartComponent` | interface | | todo | |
| `MenuProps` | interface | | todo | |
| `MenuReportIssueProps` | interface | | todo | |
| `MenuReviewersProps` | type | | todo | |
| `MenuRowProps` | interface | | todo | |
| `MenuSeparatorProps` | interface | | todo | |
| `MenuSubmenuBaseProps` | interface | | todo | |
| `MenuSubmenuProps` | interface | | todo | |
| `MenuTableGridProps` | interface | | todo | |
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
| `ObjectSlotValueMap` | re-export from core | | todo | |
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
| `PageSetup` | re-export from core | | todo | |
| `PageSetupDialogFields` | interface | | todo | |
| `PageSetupUpdate` | interface | | todo | |
| `PaginatedDocxEditor` | const | | todo | The proposal lists it with the deprecated shell, but the snapshot has no `@deprecated` tag. |
| `PaginatedDocxEditorExpose` | type | | todo | |
| `PaginatedDocxEditorHandle` | interface | | todo | |
| `PaginatedDocxEditorProps` | interface | | todo | |
| `PaginatedDocxEditorShellProps` | interface | | todo | |
| `PANE_SECTIONS` | re-export from core | | todo | |
| `PaneSection` | re-export from core | | todo | |
| `PaneSectionId` | re-export from core | | todo | |
| `PaneSectionIds` | re-export from core | | todo | |
| `paneSections` | re-export from core | | todo | |
| `ParagraphAlignment` | type | | todo | |
| `ParagraphFlagState` | re-export from core | | todo | |
| `ParagraphFormatRead` | re-export from core | | todo | |
| `ParagraphFormatUpdate` | re-export from core | | todo | |
| `ParagraphsChangeEvent` | re-export from core | | todo | |
| `ParagraphStyleItemProps` | interface | | todo | |
| `ParagraphStyleNamespace` | interface | | todo | |
| `ParagraphStyleOption` | interface | | todo | |
| `ParagraphStylePartProps` | interface | | todo | |
| `ParagraphStyleProps` | interface | | todo | |
| `ParagraphTabStop` | re-export from core | | todo | |
| `ParagraphText` | re-export from core | | todo | |
| `ParagraphTextProofing` | re-export from core | | todo | |
| `ParagraphTextReadOptions` | re-export from core | | todo | |
| `ParagraphTextSpan` | re-export from core | | todo | |
| `provideDocxEditor` | function | | todo | |
| `ProvideDocxEditorResult` | interface | | todo | |
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
| `runToolbarCommand` | re-export from core | | todo | |
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
| `TableBorderColorNamespace` | interface | | todo | |
| `TableBorderStyleNamespace` | interface | | todo | |
| `TableBorderTargetNamespace` | interface | | todo | |
| `TableBorderWidthNamespace` | interface | | todo | |
| `TableCellFillNamespace` | interface | | todo | |
| `TableChromeItemProps` | interface | | todo | |
| `TableChromePartComponent` | interface | | todo | |
| `TableChromePartProps` | interface | | todo | |
| `TextFormFieldDialogFields` | interface | | todo | |
| `TextMatch` | re-export from core | | todo | |
| `textOffsetFromAccepted` | re-export from core | | todo | |
| `ToolbarActionProps` | interface | | todo | |
| `ToolbarAlignmentComponent` | interface | | todo | |
| `ToolbarButtonProps` | interface | | todo | |
| `ToolbarChartInsert` | const | | todo | |
| `toolbarCommandState` | re-export from core | | todo | |
| `ToolbarCommandState` | re-export from core | | todo | |
| `ToolbarContext` | const | | todo | |
| `ToolbarContextValue` | interface | | todo | |
| `ToolbarImageProperties` | const | | todo | |
| `ToolbarPartComponent` | interface | | todo | |
| `ToolbarPartProps` | type | | todo | |
| `ToolbarReviewersProps` | type | | todo | |
| `ToolbarSeparatorProps` | interface | | todo | |
| `ToolbarSlotPartComponent` | interface | | todo | |
| `ToolbarSlotPartProps` | interface | | todo | |
| `ToolbarTranslate` | type | | todo | |
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
| `useContextMenuAnchor` | function | | todo | |
| `useContextMenuHighlights` | function | | todo | |
| `useContextMenuTarget` | function | | todo | |
| `UseDialogReturn` | interface | | todo | |
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
| `useFontFamily` | function | | todo | |
| `UseFontFamilyResult` | interface | | todo | |
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
| `useHyperlinkPopup` | function | | todo | |
| `useHyperlinkPopupInstance` | function | | todo | |
| `UseHyperlinkPopupResult` | interface | | todo | |
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
| `usePageSetup` | function | | todo | |
| `usePageSetupDialog` | function | | todo | |
| `UsePageSetupDialogReturn` | interface | | todo | |
| `UsePageSetupReturn` | interface | | todo | |
| `usePaneRequest` | function | | todo | |
| `useParagraphDialog` | function | | todo | |
| `UseParagraphDialogReturn` | interface | | todo | |
| `useParagraphFormat` | function | | todo | |
| `UseParagraphFormatReturn` | interface | | todo | |
| `useParagraphIndent` | function | | todo | |
| `UseParagraphIndentReturn` | interface | | todo | |
| `useParagraphStyle` | function | | todo | |
| `UseParagraphStyleResult` | interface | | todo | |
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
| `useTableBorderTargetLabel` | function | | todo | |
| `useTextFormFieldDialog` | function | | todo | |
| `UseTextFormFieldDialogReturn` | interface | | todo | |
| `useToolbarContext` | function | | todo | |
| `useToolbarLabel` | function | | todo | |
| `useToolbarLabelFor` | function | | todo | |
| `useTranslation` | function | `injectTranslation` | ported | Returns `DocxEditorTranslation`. |
| `useWatermark` | function | | todo | |
| `useWatermarkDialog` | function | | todo | |
| `UseWatermarkDialogReturn` | interface | | todo | |
| `UseWatermarkReturn` | interface | | todo | |
| `useZoom` | function | | todo | |
| `UseZoomResult` | interface | | todo | |
| `VERSION` | const | | todo | |
| `VerticalRuler` | const | | todo | |
| `VerticalRulerProps` | interface | | todo | |
| `Watermark` | re-export from core | | todo | |
| `WatermarkDialogFieldName` | type | | todo | |
| `WatermarkDialogFields` | interface | | todo | |
| `WatermarkDialogKind` | type | | todo | |
| `WatermarkDialogLayout` | type | | todo | |
| `WatermarkState` | re-export from core | | todo | |
| `WORD_DEFAULT_FONT` | re-export from core | | todo | |

## Tests

87 rows.

| Vue test | Angular spec | Status | Notes |
| --- | --- | --- | --- |
| `packages/vue/test/bookmark-dialog.test.ts` | | todo | Task 6.4 |
| `packages/vue/test/build-external.test.ts` | | todo | |
| `packages/vue/test/chrome-composition.test.ts` | | todo | Task 4.1 |
| `packages/vue/test/chrome-popup-dismissal.test.ts` | | todo | Task 5.3 |
| `packages/vue/test/content-control-widget-parts.test.ts` | | todo | Task 9.3 |
| `packages/vue/test/content-lock-policy-prop.test.ts` | `packages/angular/src/lib/live-inputs.spec.ts` | ported | Keeps: the root and `<docx-editor>` pass the policy, so bold over a locked control succeeds only with it and the text stays; an equal new object does not warn; a changed value is ignored and warns once. |
| `packages/vue/test/context-menu-add-comment.test.ts` | | todo | |
| `packages/vue/test/context-menu-chart.test.ts` | | todo | Task 5.2 |
| `packages/vue/test/context-menu-keyboard.test.ts` | | todo | Task 5.3 |
| `packages/vue/test/context-menu-link.test.ts` | | todo | Task 5.2 |
| `packages/vue/test/context-menu-placement.test.ts` | | todo | Task 5.2 |
| `packages/vue/test/context-menu-start.test.ts` | | todo | Task 5.2 |
| `packages/vue/test/context-menu-target.test.ts` | | todo | Task 5.2 |
| `packages/vue/test/context-menu.test.ts` | | todo | Task 5.2 |
| `packages/vue/test/deferred-notifier.test.ts` | | todo | Task 10.5 |
| `packages/vue/test/dialog-customization.test.ts` | | todo | Task 6.1 |
| `packages/vue/test/document-refresh-edit.test.ts` | | todo | Task 10.3 |
| `packages/vue/test/editing-mode-author.test.ts` | `packages/angular/src/lib/construction-inputs.spec.ts` | todo | Ported: `<docx-editor mode='suggesting'>` with the review module waits for an author, then records typing as tracked changes by that author. The editing mode menu claims port with the toolbar in task 4.2. |
| `packages/vue/test/editor-accessibility.test.ts` | | todo | Task 10.5 |
| `packages/vue/test/editor-composition.test.ts` | | todo | Task 10.5 |
| `packages/vue/test/engine-popup-renderers.test.ts` | | todo | Task 6.1 |
| `packages/vue/test/equation-popup.test.ts` | | todo | Task 6.6 |
| `packages/vue/test/export-popup.test.ts` | | todo | Task 6.5 |
| `packages/vue/test/field-results-prop.test.ts` | `packages/angular/src/lib/live-inputs.spec.ts` | ported | Keeps: the root and `<docx-editor>` pass `fieldResults`; `'editable'` types inside the saved result, the default types after it. A later change is ignored and warns once. |
| `packages/vue/test/font-origins.test.ts` | `packages/angular/src/lib/fonts.spec.ts` | ported | Keeps: an eager loader holds the document until it settles; a failing loader releases it and warns; eager origins compose first-wins and one rejection keeps the others; an on-demand resolver releases the document at once behind one stable identity that reads the latest origins; one resolver makes a list on demand, and a loader in it gets no argument; `injectFonts` gives the request to marked and unmarked resolvers of every arity, composes resolvers, returns a marked resolver, reports a function that returns an origin, reads `computed` and signal origins again on each resolve. |
| `packages/vue/test/font-shaping-lifecycle.test.ts` | `packages/angular/src/lib/construction-inputs.spec.ts` | ported | Keeps: an equal new `fontShaping` keeps edits, selection, and undo; a changed `maxFontBytes` creates a new engine, and later equal values keep that one. |
| `packages/vue/test/fonts-prop.test.ts` | `packages/angular/src/lib/construction-inputs.spec.ts` | ported | Keeps: a new `fonts` value on the root and on `<docx-editor>` reaches the same engine and session through `setFonts`. |
| `packages/vue/test/form-field-i18n.test.ts` | `packages/angular/src/lib/i18n.spec.ts`, `packages/angular/src/lib/live-inputs.spec.ts` | todo | Ported: engine form labels follow the catalog and an explicit `translate` without a new engine, on the root and on `<docx-editor>`. The text form field dialog claims (the open dialog keeps its input and focus) port with task 6.4. |
| `packages/vue/test/header-footer-chrome.test.ts` | | todo | Task 9.1 |
| `packages/vue/test/highlights-background-layout.test.ts` | | todo | Task 7.2 |
| `packages/vue/test/host-window-chrome.test.ts` | `packages/angular/src/lib/host-window.spec.ts` | todo | Ported: the window panel renders in the pop-out while open and comes back on close; turning `open` off moves the panel back before the window closes. The picker-dismissal claim needs the font family control and ports with task 4.2. |
| `packages/vue/test/hydration-renderer-slots.test.ts` | | todo | Task 10.5 |
| `packages/vue/test/hyperlink-as-child.test.ts` | | todo | Task 6.2 |
| `packages/vue/test/image-authoring.test.ts` | | todo | Task 8.1 |
| `packages/vue/test/image-decode-port-prop.test.ts` | `packages/angular/src/lib/construction-inputs.spec.ts` | ported | Keeps: the root and `<docx-editor>` decode the picture with the port and keep its bytes on save; a new port creates a new engine that uses it; destroying the editor aborts a running decode. |
| `packages/vue/test/image-insert-refusal.test.ts` | | todo | Task 8.1 |
| `packages/vue/test/image-overlay-history.test.ts` | | todo | Task 8.1 |
| `packages/vue/test/image-panel-rows.test.ts` | | todo | Task 8.1 |
| `packages/vue/test/lifecycle-regression.test.ts` | | todo | Task 3.3 |
| `packages/vue/test/link-row-chrome.test.ts` | | todo | Task 6.2 |
| `packages/vue/test/live-props.test.ts` | `packages/angular/src/lib/live-inputs.spec.ts` | todo | Ported: author, mode, translate, and locale apply to the same engine; another input change does not restore the host mode; a catalog change and then `translate` reach the engine; an equal new `translate` publishes no selection change; `colorMode` applies. The `rulers`, `navigation`, and `title` claims port with tasks 9.6, 7.3, and 5.4. |
| `packages/vue/test/loading-continued.test.ts` | | todo | Task 9.5 |
| `packages/vue/test/loading-indicator.test.ts` | | todo | Task 9.5 |
| `packages/vue/test/locked-control-navigation-prop.test.ts` | `packages/angular/src/lib/live-inputs.spec.ts` | ported | Keeps: the root and `<docx-editor>` pass the mode; `'atomic'` steps over the control, the default steps into it. A later change is ignored and warns once, only for a real change. |
| `packages/vue/test/menu-composition.test.ts` | | todo | Task 5.1 |
| `packages/vue/test/menu-download-name.test.ts` | | todo | Task 5.1 |
| `packages/vue/test/menu-opening.test.ts` | | todo | Task 5.1 |
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
| `packages/vue/test/paginated-form-save.test.ts` | | todo | Task 3.3 |
| `packages/vue/test/paragraph-dialog.test.ts` | | todo | Task 6.3 |
| `packages/vue/test/paragraph-hooks-lock-state.test.ts` | | todo | Task 6.3 |
| `packages/vue/test/paragraph-style-labels.test.ts` |  | todo | Needs the paragraph style picker. Ports with task 4.2. |
| `packages/vue/test/popup-config.test.ts` | | todo | Task 6.1 |
| `packages/vue/test/print-popup.test.ts` | | todo | Task 6.5 |
| `packages/vue/test/props-primitives.test.ts` | | todo | Task 10.5 |
| `packages/vue/test/review-gutter.test.ts` | | todo | Task 10.1 |
| `packages/vue/test/review-overflow-navigation.test.ts` | | todo | Task 10.2 |
| `packages/vue/test/revision-markup.test.ts` | | todo | Task 10.2 |
| `packages/vue/test/root-color-mode.test.ts` | `packages/angular/src/lib/color-mode.spec.ts`, `packages/angular/src/lib/host-window.spec.ts` | ported | Keeps: 'dark' puts `.dark` beside `.docx-editor`; omitted and 'light' stay light; a change applies without new elements; 'system' follows the OS setting and stops when it leaves 'system'; the pop-out follows the root; `injectEditorColorMode` and `injectScopeClassName` report the theme. The toolbar and menu bar check their own scope classes in tasks 4.1 and 5.1. |
| `packages/vue/test/runtime-audit.test.ts` | | todo | Task 10.5 |
| `packages/vue/test/scope-dispose.test.ts` | `packages/angular/src/lib/color-mode.spec.ts` | ported | Keeps: mounting the composed editor logs no warning and no error. |
| `packages/vue/test/scoped-chrome-anchor.test.ts` | | todo | Task 10.5 |
| `packages/vue/test/security.test.ts` | | todo | Task 6.2 |
| `packages/vue/test/selection-range-state.test.ts` | | todo | Task 10.5 |
| `packages/vue/test/slot-ref.test.ts` | | todo | Task 10.5 |
| `packages/vue/test/sugar-host.test.ts` | | todo | Task 3.3 |
| `packages/vue/test/text-highlights.test.ts` | | todo | Task 7.2 |
| `packages/vue/test/toolbar-dev-warnings.test.ts` | | todo | Task 4.2 |
| `packages/vue/test/toolbar-direction.test.ts` | `packages/angular/src/lib/i18n.spec.ts` | todo | Ported: the provider `dir` wins over the catalog; a right-to-left catalog mirrors the composed parts; a left-to-right catalog leaves the host direction. The direction button claims port with task 4.5, the packaged-editor claim is in `live-inputs.spec.ts` (Arabic and Hebrew mirror the `.docx-editor` wrapper, Portuguese leaves it). |
| `packages/vue/test/toolbar-host-groups.test.ts` | | todo | Task 4.1 |
| `packages/vue/test/toolbar-link-platform.test.ts` | | todo | Task 4.5 |
| `packages/vue/test/toolbar-picker-accessibility.test.ts` | | todo | Task 4.2 |
| `packages/vue/test/unsupported-elements.test.ts` | | todo | Task 9.5 |
| `packages/vue/test/use-comment-draft-request.test.ts` | | todo | Task 10.1 |
| `packages/vue/test/use-editor-command.test.ts` | `packages/angular/src/lib/editor-command.spec.ts` | ported | Keeps: state follows a payload change with the same type; `execute` returns false on a refusal; many commands share the root listeners; custom equality stops slice updates; a burst collapses through the deferred notifier; outside a root the command is disabled; range binding groups keyboard repeats and cleans up; a held stepper starts before its own keydown handler. |
| `packages/vue/test/viewport-rail.test.ts` | | todo | Task 10.1 |
| `packages/vue/test/watermark-dialog.test.ts` | | todo | Task 6.3 |
