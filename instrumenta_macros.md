# Instrumenta Macro References

This document provides a list of **macro references** of Instrumenta that can be used in the **built-in** Shortcut Editor (or directly in the `shortcuts.csv` file) for configuring **keyboard shortcuts** in Instrumenta Keys.

## Supported Shortcut Keys
The keyboard shortcuts allow for a combination of the following keys:

### Modifier Keys
- Ctrl
- Shift
- Alt

### Function Keys
- F1 to F12

### Navigation Keys
- Up (Arrow Up)
- Down (Arrow Down)
- Left (Arrow Left)
- Right (Arrow Right)
- Home
- End
- PageUp
- PageDown

### Special Keys
- Esc
- Enter
- Tab
- Space
- Backspace
- Insert
- Del (Delete)
- Plus (`+`)

### Numeric Keys (Numpad)
- Num0 to Num9
- NumLock
- NumpadDivide (`/`)
- NumpadMultiply (`*`)
- NumpadSubtract (`-`)
- NumpadAdd (`+`)
- NumpadEnter

### Lock & System Keys
- CapsLock
- ScrollLock
- PrintScreen
- PauseBreak

### Standard ASCII Keys

|      |      |      |      |      |      |      |      |      |      |
|------|------|------|------|------|------|------|------|------|------|
| `!`  | `"`  | `#`  | `$`  | `%`  | `&`  | `'`  | `(`  | `)`  | `*`  |
| `,`  | `-`  | `.`  | `/`  | `0`  | `1`  | `2`  | `3`  | `4`  | `5`  |
| `6`  | `7`  | `8`  | `9`  | `:`  | `;`  | `<`  | `=`  | `>`  | `?`  |
| `@`  | `A`  | `B`  | `C`  | `D`  | `E`  | `F`  | `G`  | `H`  | `I`  |
| `J`  | `K`  | `L`  | `M`  | `N`  | `O`  | `P`  | `Q`  | `R`  | `S`  |
| `T`  | `U`  | `V`  | `W`  | `X`  | `Y`  | `Z`  | `[`  | `\`  | `]`  |
| `^`  | `_`  | `` ` `` |

---

## How to use 
To assign a shortcut to a macro, **use the built-in** Shortcut Editor (which you can open in PowerPoint by pressing CTRL+SHIFT+ALT+Q) or **edit the `shortcuts.csv` file** manually and specify:
- **Shortcut key combination** (with `+` between the different keys)
- **Macro reference** (from the table below)

Examples:
- Ctrl+Shift+Alt+Q,InstrumentaKeysEditor
- Ctrl+Shift+Alt+S,ObjectsSwapPosition
- Ctrl+Shift+Alt+L,ObjectsAlignLefts
- Ctrl+Shift+Alt+T,ObjectsAlignTops
- Ctrl+Shift+Alt+R,ObjectsAlignRights
- Ctrl+Shift+Alt+B,ObjectsAlignBottoms
- Ctrl+Shift+Alt+E,ObjectsAlignCenters
- Ctrl+Shift+Alt+M,ObjectsAlignMiddles
- Ctrl+Shift+Alt+H,ObjectsDistributeHorizontally
- Ctrl+Shift+Alt+V,ObjectsDistributeVertically
- Alt+Q,GenerateStickyNote

![image](https://github.com/user-attachments/assets/ffea3f59-6523-49e2-97e4-230f373e8344)


---

## Macro List
Below is the list of Instrumenta macro references you can use:

| Macro Description                                         | Macro Reference                  |
|-----------------------------------------------------------|----------------------------------|
| Open Instrumenta Keys Shortcut Editor | InstrumentaKeysEditor | 
| Copy selected slides to slide library | AddSelectedSlidesToLibraryFile | 
| Anonymize all/selected slides with Lorem Ipsum | AnonymizeWithLoremIpsum | 
| Apply same crop to selected pictures | ApplySameCropToSelectedImages | 
| Arrange shapes | ArrangeShapes | 
| Average based on selected star ratings | AverageFivePointStars | 
| Average based on selected Harvey Balls | AverageHarveyBall | 
| Average based on selected RAG status-shapes | AverageRAGStatus | 
| Add page numbers to all/selected slides (except the first) | CleanUpAddSlideNumbers | 
| Move to end and hide selected slides | CleanUpHideAndMoveSelectedSlides | 
| Remove animations from all/selected slides | CleanUpRemoveAnimationsFromAllSlides | 
| Remove comments from all/selected slides | CleanUpRemoveCommentsFromAllSlides | 
| Remove all hidden slides | CleanUpRemoveHiddenSlides | 
| Remove entry transitions from all/selected slides | CleanUpRemoveSlideShowTransitionsFromAllSlides | 
| Remove speaker notes from all/selected slides | CleanUpRemoveSpeakerNotesFromAllSlides | 
| Remove all unused master slides | CleanUpRemoveUnusedMasterSlides | 
| Color bold text (automatically) | ColorBoldTextColorAutomatically | 
| Color bold text (color picker) | ColorBoldTextColorPicker | 
| Connect shape 1 bottom side to shape 2 top side | ConnectRectangleShapesBottomToTop | 
| Connect shape 1 left side to shape 2 right side | ConnectRectangleShapesLeftToRight | 
| Connect shape 1 right side to shape 2 left side | ConnectRectangleShapesRightToLeft | 
| Connect shape 1 top side to shape 2 bottom side | ConnectRectangleShapesTopToBottom | 
| Convert comments on all slides to notes | ConvertAllCommentsToStickyNotes | 
| Convert comments on this slide to notes | ConvertCommentsToStickyNotes | 
| Convert shapes to table | ConvertShapesToTable | 
| Convert all/selected slides to pictures (readonly) | ConvertSlidesToPictures | 
| Convert table to shapes | ConvertTableToShapes | 
| Copy position and dimensions | CopyPosition | 
| Copy slide notes to clipboard | CopySlideNotesToClipboardOnly | 
| Export slide notes to Word | CopySlideNotesToWord | 
| Copy storyline to clipboard | CopyStorylineToClipBoardOnly | 
| Export storyline to Word | CopyStorylineToWord | 
| Create or update agenda pages | CreateOrUpdateMasterAgenda | 
| Delete Stamps on all slides | DeleteStampsOnAllSlides | 
| Delete Stamps on this slide | DeleteStampsOnSlide | 
| Delete notes on all slides | DeleteStickyNotesOnAllSlides | 
| Delete notes on this slide | DeleteStickyNotesOnSlide | 
| Delete selected multislide shape on all slides | DeleteTaggedShapes | 
| E-mail selected slides | EmailSelectedSlides | 
| E-mail selected slides as PDF | EmailSelectedSlidesAsPDF | 
| Mail merge full presentation based on Excel-file (duplicate presentation file) | ExcelFullFileMailMerge | 
| Mail merge this slide based on Excel-file (duplicate slide within this presentation) | ExcelMailMerge | 
| Add cross-slide steps counter | GenerateCrossSlideStepsCounter | 
| 0.5 star | GenerateFivePointStars05 | 
| 1 star | GenerateFivePointStars10 | 
| 1.5 star | GenerateFivePointStars15 | 
| 2 star | GenerateFivePointStars20 | 
| 2.5 star | GenerateFivePointStars25 | 
| 3 star | GenerateFivePointStars30 | 
| 3.5 star | GenerateFivePointStars35 | 
| 4 star | GenerateFivePointStars40 | 
| 4.5 star | GenerateFivePointStars45 | 
| 5 star | GenerateFivePointStars50 | 
| Harvey Ball 10% | GenerateHarveyBall10 | 
| Harvey Ball 100% | GenerateHarveyBall100 | 
| Harvey Ball 20% | GenerateHarveyBall20 | 
| Harvey Ball 25% | GenerateHarveyBall25 | 
| Harvey Ball 30% | GenerateHarveyBall30 | 
| Harvey Ball 33% | GenerateHarveyBall33 | 
| Harvey Ball 40% | GenerateHarveyBall40 | 
| Harvey Ball 50% | GenerateHarveyBall50 | 
| Harvey Ball 60% | GenerateHarveyBall60 | 
| Harvey Ball 67% | GenerateHarveyBall67 | 
| Harvey Ball 70% | GenerateHarveyBall70 | 
| Harvey Ball 75% | GenerateHarveyBall75 | 
| Harvey Ball 80% | GenerateHarveyBall80 | 
| Harvey Ball 90% | GenerateHarveyBall90 | 
| Custom Harvey Ball | GenerateHarveyBallCustom | 
| RAG status - Amber | GenerateRAGStatusAmber | 
| RAG status - Green | GenerateRAGStatusGreen | 
| RAG status - Red | GenerateRAGStatusRed | 
| Stamp CONFIDENTIAL | GenerateStampConfidential | 
| Stamp DO NOT DISTRIBUTE | GenerateStampDoNotDistribute | 
| Stamp DRAFT | GenerateStampDraft | 
| Stamp NEW | GenerateStampNew | 
| Stamp TO APPENDIX | GenerateStampToAppendix | 
| Stamp TO BE REMOVED | GenerateStampToBeRemoved | 
| Stamp UPDATED | GenerateStampUpdated | 
| Add steps counter | GenerateStepsCounter | 
| Insert sticky note | GenerateStickyNote | 
| Group shapes by columns | GroupShapesByColumns | 
| Group shapes by rows | GroupShapesByRows | 
| Hide slidetags on all slides | HideTagsOnSlide | 
| Insert merge fields from Excel-file | ImportHeadersFromExcel | 
| Increase shape transparency | IncreaseShapeTransparency | 
| Set position | InitialiseSetPositionAppEventHandler | 
| Insert caption for selected shape | InsertCaption | 
| Insert column left (preserve widths) | InsertColumnToLeftKeepOtherColumnWidths | 
| Insert column right (preserve widths) | InsertColumnToRightKeepOtherColumnWidths | 
| Insert empty merge field | InsertMergeField | 
| Insert process (SmartArt) | InsertProcessSmartArt | 
| Insert QR-code | InsertQRCode | 
| Watermark and then convert all/selected slides to pictures (readonly) | InsertWatermarkAndConvertSlidesToPictures | 
| Lock position of all shapes on all slides | LockAllShapesOnAllSlides | 
| Toggle lock aspect ratio of selected shape(s) | LockAspectRatioToggleSelectedShapes | 
| Toggle lock or unlock position all shapes on all slides | LockToggleAllShapesOnAllSlides | 
| Toggle lock or unlock position of selected shapes | LockToggleSelectedShapes | 
| Manually replace all merge fields on all slides | ManualMailMerge | 
| Move Stamps off all slides | MoveStampsOffAllSlides | 
| Move Stamps off this slide | MoveStampsOffSlide | 
| Move Stamps on all slides | MoveStampsOnAllSlides | 
| Move Stamps on this slide | MoveStampsOnSlide | 
| Move notes off all slides | MoveStickyNotesOffAllSlides | 
| Move notes off this slide | MoveStickyNotesOffSlide | 
| Move notes on all slides | MoveStickyNotesOnAllSlides | 
| Move notes on this slide | MoveStickyNotesOnSlide | 
| Move column left | MoveTableColumnLeft | 
| Move column left (ignore borders) | MoveTableColumnLeftIgnoreBorders | 
| Move column left (text only) | MoveTableColumnLeftTextOnly | 
| Move column right | MoveTableColumnRight | 
| Move column right (ignore borders) | MoveTableColumnRightIgnoreBorders | 
| Move column right (text only) | MoveTableColumnRightTextOnly | 
| Move row down | MoveTableRowDown | 
| Move row down (ignore borders) | MoveTableRowDownIgnoreBorders | 
| Move row down (text only) | MoveTableRowDownTextOnly | 
| Move row up | MoveTableRowUp | 
| Move row up (ignore borders) | MoveTableRowUpIgnoreBorders | 
| Move row up (text only) | MoveTableRowUpTextOnly | 
| Align bottom | ObjectsAlignBottoms | 
| Align center | ObjectsAlignCenters | 
| Align left | ObjectsAlignLefts | 
| Align middle | ObjectsAlignMiddles | 
| Align right | ObjectsAlignRights | 
| Align top | ObjectsAlignTops | 
| Center align shapes to table cells | ObjectsAlignToTable | 
| Center align shapes to columns of a table | ObjectsAlignToTableColumn | 
| Center align shapes to rows of a table | ObjectsAlignToTableRow | 
| Do not Autofit | ObjectsAutoSizeNone | 
| Resize shape to fit text | ObjectsAutoSizeShapeToFitText | 
| Resize text on overflow | ObjectsAutoSizeTextToFitShape | 
| Clone selection down | ObjectsCloneDown | 
| Clone selection to right | ObjectsCloneRight | 
| Copy rounded corner of first selected shape to selected shapes | ObjectsCopyRoundedCorner | 
| Copy shape type and all adjustments of first selected shape to selected shapes | ObjectsCopyShapeTypeAndAdjustments | 
| Decrease line spacing | ObjectsDecreaseLineSpacing | 
| Decrease line spacing before and after | ObjectsDecreaseLineSpacingBeforeAndAfter | 
| Decrease horizontal gap between shapes | ObjectsDecreaseSpacingHorizontal | 
| Decrease vertical gap between shapes | ObjectsDecreaseSpacingVertical | 
| Distribute horizontally | ObjectsDistributeHorizontally | 
| Distribute vertically | ObjectsDistributeVertically | 
| Increase line spacing | ObjectsIncreaseLineSpacing | 
| Increase line spacing before and after | ObjectsIncreaseLineSpacingBeforeAndAfter | 
| Increase horizontal gap between shapes | ObjectsIncreaseSpacingHorizontal | 
| Increase vertical gap between shapes | ObjectsIncreaseSpacingVertical | 
| Decrease margins | ObjectsMarginsDecrease | 
| Increase margins | ObjectsMarginsIncrease | 
| Remove margins | ObjectsMarginsToZero | 
| Remove hyperlinks | ObjectsRemoveHyperlinks | 
| Remove horizontal gap between shapes | ObjectsRemoveSpacingHorizontal | 
| Remove vertical gap between shapes | ObjectsRemoveSpacingVertical | 
| Remove text | ObjectsRemoveText | 
| Set same height | ObjectsSameHeight | 
| Set same height and width | ObjectsSameHeightAndWidth | 
| Set same width | ObjectsSameWidth | 
| Select shapes with same fill and line color | ObjectsSelectBySameFillAndLineColor | 
| Select shapes with same fill color | ObjectsSelectBySameFillColor | 
| Select shapes with same height | ObjectsSelectBySameHeight | 
| Select shapes with same line color | ObjectsSelectBySameLineColor | 
| Select shapes with same type | ObjectsSelectBySameType | 
| Select shapes with same width | ObjectsSelectBySameWidth | 
| Select shapes with same size | ObjectsSelectBySameWidthAndHeight | 
| Size to narrowest | ObjectsSizeToNarrowest | 
| Size to shortest | ObjectsSizeToShortest | 
| Size to tallest | ObjectsSizeToTallest | 
| Size to widest | ObjectsSizeToWidest | 
| Stretch shapes to bottom | ObjectsStretchBottom | 
| Stretch shapes to the top edge of the bottommost shape | ObjectsStretchBottomShapeTop | 
| Stretch shapes to left | ObjectsStretchLeft | 
| Stretch shapes to the right edge of the leftmost shape | ObjectsStretchLeftShapeRight | 
| Stretch shapes to right | ObjectsStretchRight | 
| Stretch shapes to the left edge of the rightmost shape | ObjectsStretchRightShapeLeft | 
| Stretch shapes to top | ObjectsStretchTop | 
| Stretch shapes to the bottom edge of the topmost shape | ObjectsStretchTopShapeBottom | 
| Swap position of two shapes | ObjectsSwapPosition | 
| Swap position of two shapes (centered) | ObjectsSwapPositionCentered | 
| Swap text (with formatting) | ObjectsSwapText | 
| Swap text (no formatting) | ObjectsSwapTextNoFormatting | 
| Delete strikethrough text | ObjectsTextDeleteStrikethrough | 
| Merge text of all shapes in first selected shape | ObjectsTextMerge | 
| Split (by paragraphs) into multiple shapes | ObjectsTextSplitByParagraph | 
| Toggle text wrap | ObjectsTextWordwrapToggle | 
| Toggle autofit | ObjectsToggleAutoSize | 
| Open slide library file | OpenSlideLibraryFile | 
| Optimized (10 iterations) | OptimizeTableHeight10Iterations | 
| Optimized (20 iterations) | OptimizeTableHeight20Iterations | 
| Optimized (3 iterations) | OptimizeTableHeight3Iterations | 
| Optimized (5 iterations) | OptimizeTableHeight5Iterations | 
| Quick | OptimizeTableHeightQuick | 
| Paste position | PastePosition | 
| Paste position and dimensions | PastePositionAndDimensions | 
| Paste storyline in shape | PasteStorylineInSelectedShape | 
| Crop picture or shape to slide | PictureCropToSlide | 
| Rectify lines | RectifyLines | 
| Renumber all captions | ReNumberCaptions | 
| Resize and space horizontally (equal spacing) | ResizeAndSpaceEvenHorizontal | 
| Resize and space horizontally (preserve first) | ResizeAndSpaceEvenHorizontalPreserveFirst | 
| Resize and space horizontally (preserve last) | ResizeAndSpaceEvenHorizontalPreserveLast | 
| Resize and space vertically (equal spacing) | ResizeAndSpaceEvenVertical | 
| Resize and space vertically (preserve first) | ResizeAndSpaceEvenVerticalPreserveFirst | 
| Resize and space vertically (preserve last) | ResizeAndSpaceEvenVerticalPreserveLast | 
| Save selected slides | SaveSelectedSlides | 
| Select all cross-slide step counters on this slide | SelectAllCrossSlideStepsCounter | 
| Select all step counters on this slide | SelectAllStepsCounter | 
| Show about dialog | ShowAboutDialog | 
| Set proofing language on all objects and all slides | ShowChangeSpellCheckLanguageForm | 
| Copy shape to multiple slides (multislide shape) | ShowFormCopyShapeToMultipleSlides | 
| Manage hidden tags to shapes or slides | ShowFormManageTags | 
| Select slides by tag(s) or specific stamp | ShowFormSelectSlidesByTag | 
| Instrumenta settings | ShowSettings | 
| Insert slide from slide library | ShowSlideLibrary | 
| Show slidetags on all slides | ShowTagsOnSlide | 
| Split table by column | SplitTableByColumn | 
| Split table by row | SplitTableByRow | 
| Decrease column gaps | TableColumnDecreaseGaps | 
| Add column gaps | TableColumnGapsEven | 
| Add column gaps (including left and right sides) | TableColumnGapsOdd | 
| Increase column gaps | TableColumnIncreaseGaps | 
| Remove column gaps | TableColumnRemoveGaps | 
| Distribute columns ignoring column gaps | TableDistributeColumnsWithGaps | 
| Distribute rows ignoring row gaps | TableDistributeRowsWithGaps | 
| Quick format table | TableQuickFormat | 
| Remove cell fills | TableRemoveBackgrounds | 
| Remove all borders | TableRemoveBorders | 
| Decrease row gaps | TableRowDecreaseGaps | 
| Add row gaps | TableRowGapsEven | 
| Add row gaps (including top and bottom sides) | TableRowGapsOdd | 
| Increase row gaps | TableRowIncreaseGaps | 
| Remove row gaps | TableRowRemoveGaps | 
| Sum row (values left from selected cells) | TableRowSum | 
| Decrease margins of selected table or selected cells | TablesMarginsDecrease | 
| Increase margins of selected table or selected cells | TablesMarginsIncrease | 
| Remove margins of selected table or selected cells | TablesMarginsToZero | 
| Sum column (values above selected cells) | TableSum | 
| Transpose table | TableTranspose | 
| Crosses | TextBulletsCrosses | 
| Ticks | TextBulletsTicks | 
| Insert Copyright | TextInsertCopyright | 
| Insert Euro | TextInsertEuro | 
| Insert NonBreakingSpace | TextInsertNoBreakSpace | 
| Unlock position of all shapes on all slides | UnLockAllShapesOnAllSlides | 
| Update position and dimensions of selected multislide shape on all slides | UpdateTaggedShapePositionAndDimensions | 
