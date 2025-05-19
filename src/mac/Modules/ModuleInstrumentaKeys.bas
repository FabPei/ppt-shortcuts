Attribute VB_Name = "ModuleInstrumentaKeys"
'MIT License

'Copyright (c) 2025 iappyx

'Permission is hereby granted, free of charge, to any person obtaining a copy
'of this software and associated documentation files (the "Software"), to deal
'in the Software without restriction, including without limitation the rights
'to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
'copies of the Software, and to permit persons to whom the Software is
'furnished to do so, subject to the following conditions:

'The above copyright notice and this permission notice shall be included in all
'copies or substantial portions of the Software.

'THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
'IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
'FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
'AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
'LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
'OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
'SOFTWARE.


Sub InstrumentaKeysHelperInitialize(Ribbon As IRibbonUI)
    
   #If Mac Then
   
   
   
   #Else
   
        MsgBox "This add-in is Mac only. For Instrumenta Keys for Windows see: https://github.com/iappyx/Instrumenta-Keys", vbCritical
        
    #End If
    
End Sub

Sub ShowInstrumentaKeysHelperAboutDialog()

    Dim InstrumentaKeysHelperVersion As String
    InstrumentaKeysHelperVersion = "0.10"
    MsgBox "Instrumenta Keys Helper v" & InstrumentaKeysHelperVersion, vbInformation
    
End Sub

Sub RunFunction(control As IRibbonControl)
    
    #If Mac Then
        
        ExecuteInstrumentaFunction
        
    #Else
        
        MsgBox "This add-in is Mac only. For Instrumenta Keys for Windows see: https://github.com/iappyx/Instrumenta-Keys", vbCritical
        
    #End If
    
End Sub

Function IsApprovedFunction(funcName As String) As Boolean

    Dim approvedFunctions As Variant
    approvedFunctions = Array( _
                        "AddSelectedSlidesToLibraryFile", "AnonymizeWithLoremIpsum", "ApplySameCropToSelectedImages", "ArrangeShapes", "AverageFivePointStars", "AverageHarveyBall", "AverageRAGStatus", "CleanUpAddSlideNumbers", "CleanUpHideAndMoveSelectedSlides", "CleanUpRemoveAnimationsFromAllSlides", "CleanUpRemoveCommentsFromAllSlides", _
                        "CleanUpRemoveHiddenSlides", "CleanUpRemoveSlideShowTransitionsFromAllSlides", "CleanUpRemoveSpeakerNotesFromAllSlides", "CleanUpRemoveUnusedMasterSlides", "ColorBoldTextColorAutomatically", "ColorBoldTextColorPicker", "ConnectRectangleShapesBottomToTop", "ConnectRectangleShapesLeftToRight", "ConnectRectangleShapesRightToLeft", _
                        "ConnectRectangleShapesTopToBottom", "ConvertAllCommentsToStickyNotes", "ConvertCommentsToStickyNotes", "ConvertShapesToTable", "ConvertSlidesToPictures", "ConvertTableToShapes", "CopyPosition", "CopySlideNotesToClipboardOnly", "CopySlideNotesToWord", "CopyStorylineToClipBoardOnly", "CopyStorylineToWord", "CreateOrUpdateMasterAgenda", _
                        "DeleteStampsOnAllSlides", "DeleteStampsOnSlide", "DeleteStickyNotesOnAllSlides", "DeleteStickyNotesOnSlide", "DeleteTaggedShapes", "EmailSelectedSlides", "EmailSelectedSlidesAsPDF", "ExcelFullFileMailMerge", "ExcelMailMerge", "GenerateCrossSlideStepsCounter", "GenerateFivePointStars05", "GenerateFivePointStars10", "GenerateFivePointStars15", _
                        "GenerateFivePointStars20", "GenerateFivePointStars25", "GenerateFivePointStars30", "GenerateFivePointStars35", "GenerateFivePointStars40", "GenerateFivePointStars45", "GenerateFivePointStars50", "GenerateHarveyBall10", "GenerateHarveyBall100", "GenerateHarveyBall20", "GenerateHarveyBall25", "GenerateHarveyBall30", "GenerateHarveyBall33", _
                        "GenerateHarveyBall40", "GenerateHarveyBall50", "GenerateHarveyBall60", "GenerateHarveyBall67", "GenerateHarveyBall70", "GenerateHarveyBall75", "GenerateHarveyBall80", "GenerateHarveyBall90", "GenerateHarveyBallCustom", "GenerateRAGStatusAmber", "GenerateRAGStatusGreen", "GenerateRAGStatusRed", "GenerateStampConfidential", "GenerateStampDoNotDistribute", _
                        "GenerateStampDraft", "GenerateStampNew", "GenerateStampToAppendix", "GenerateStampToBeRemoved", "GenerateStampUpdated", "GenerateStepsCounter", "GenerateStickyNote", "GroupShapesByColumns", "GroupShapesByRows", "HideTagsOnSlide", "ImportHeadersFromExcel", "IncreaseShapeTransparency", "InitialiseSetPositionAppEventHandler", "InsertCaption", _
                        "InsertColumnToLeftKeepOtherColumnWidths", "InsertColumnToRightKeepOtherColumnWidths", "InsertMergeField", "InsertProcessSmartArt", "InsertQRCode", "InsertWatermarkAndConvertSlidesToPictures", "LockAllShapesOnAllSlides", "LockAspectRatioToggleSelectedShapes", "LockToggleAllShapesOnAllSlides", "LockToggleSelectedShapes", "ManualMailMerge", _
                        "MoveStampsOffAllSlides", "MoveStampsOffSlide", "MoveStampsOnAllSlides", "MoveStampsOnSlide", "MoveStickyNotesOffAllSlides", "MoveStickyNotesOffSlide", "MoveStickyNotesOnAllSlides", "MoveStickyNotesOnSlide", "MoveTableColumnLeft", "MoveTableColumnLeftIgnoreBorders", "MoveTableColumnLeftTextOnly", "MoveTableColumnRight", "MoveTableColumnRightIgnoreBorders", _
                        "MoveTableColumnRightTextOnly", "MoveTableRowDown", "MoveTableRowDownIgnoreBorders", "MoveTableRowDownTextOnly", "MoveTableRowUp", "MoveTableRowUpIgnoreBorders", "MoveTableRowUpTextOnly", "ObjectsAlignBottoms", "ObjectsAlignCenters", "ObjectsAlignLefts", "ObjectsAlignMiddles", "ObjectsAlignRights", "ObjectsAlignTops", "ObjectsAlignToTable", _
                        "ObjectsAlignToTableColumn", "ObjectsAlignToTableRow", "ObjectsAutoSizeNone", "ObjectsAutoSizeShapeToFitText", "ObjectsAutoSizeTextToFitShape", "ObjectsCloneDown", "ObjectsCloneRight", "ObjectsCopyRoundedCorner", "ObjectsCopyShapeTypeAndAdjustments", "ObjectsDecreaseLineSpacing", "ObjectsDecreaseLineSpacingBeforeAndAfter", "ObjectsDecreaseSpacingHorizontal", _
                        "ObjectsDecreaseSpacingVertical", "ObjectsDistributeHorizontally", "ObjectsDistributeVertically", "ObjectsIncreaseLineSpacing", "ObjectsIncreaseLineSpacingBeforeAndAfter", "ObjectsIncreaseSpacingHorizontal", "ObjectsIncreaseSpacingVertical", "ObjectsMarginsDecrease", "ObjectsMarginsIncrease", "ObjectsMarginsToZero", "ObjectsRemoveHyperlinks", _
                        "ObjectsRemoveSpacingHorizontal", "ObjectsRemoveSpacingVertical", "ObjectsRemoveText", "ObjectsSameHeight", "ObjectsSameHeightAndWidth", "ObjectsSameWidth", "ObjectsSelectBySameFillAndLineColor", "ObjectsSelectBySameFillColor", "ObjectsSelectBySameHeight", "ObjectsSelectBySameLineColor", "ObjectsSelectBySameType", "ObjectsSelectBySameWidth", _
                        "ObjectsSelectBySameWidthAndHeight", "ObjectsSizeToNarrowest", "ObjectsSizeToShortest", "ObjectsSizeToTallest", "ObjectsSizeToWidest", "ObjectsStretchBottom", "ObjectsStretchBottomShapeTop", "ObjectsStretchLeft", "ObjectsStretchLeftShapeRight", "ObjectsStretchRight", "ObjectsStretchRightShapeLeft", "ObjectsStretchTop", "ObjectsStretchTopShapeBottom", _
                        "ObjectsSwapPosition", "ObjectsSwapPositionCentered", "ObjectsSwapText", "ObjectsSwapTextNoFormatting", "ObjectsTextDeleteStrikethrough", "ObjectsTextMerge", "ObjectsTextSplitByParagraph", "ObjectsTextWordwrapToggle", "ObjectsToggleAutoSize", "OpenSlideLibraryFile", "OptimizeTableHeight10Iterations", "OptimizeTableHeight20Iterations", _
                        "OptimizeTableHeight3Iterations", "OptimizeTableHeight5Iterations", "OptimizeTableHeightQuick", "PastePosition", "PastePositionAndDimensions", "PasteStorylineInSelectedShape", "PictureCropToSlide", "RectifyLines", "ReNumberCaptions", "ResizeAndSpaceEvenHorizontal", "ResizeAndSpaceEvenHorizontalPreserveFirst", "ResizeAndSpaceEvenHorizontalPreserveLast", _
                        "ResizeAndSpaceEvenVertical", "ResizeAndSpaceEvenVerticalPreserveFirst", "ResizeAndSpaceEvenVerticalPreserveLast", "SaveSelectedSlides", "SelectAllCrossSlideStepsCounter", "SelectAllStepsCounter", "ShowAboutDialog", "ShowChangeSpellCheckLanguageForm", "ShowFormCopyShapeToMultipleSlides", "ShowFormManageTags", "ShowFormSelectSlidesByTag", "ShowSettings", _
                        "ShowSlideLibrary", "ShowTagsOnSlide", "SplitTableByColumn", "SplitTableByRow", "TableColumnDecreaseGaps", "TableColumnGapsEven", "TableColumnGapsOdd", "TableColumnIncreaseGaps", "TableColumnRemoveGaps", "TableDistributeColumnsWithGaps", "TableDistributeRowsWithGaps", "TableQuickFormat", "TableRemoveBackgrounds", "TableRemoveBorders", "TableRowDecreaseGaps", _
                        "TableRowGapsEven", "TableRowGapsOdd", "TableRowIncreaseGaps", "TableRowRemoveGaps", "TableRowSum", "TablesMarginsDecrease", "TablesMarginsIncrease", "TablesMarginsToZero", "TableSum", "TableTranspose", "TextBulletsCrosses", "TextBulletsTicks", "TextInsertCopyright", "TextInsertEuro", "TextInsertNoBreakSpace", "UnLockAllShapesOnAllSlides", "UpdateTaggedShapePositionAndDimensions")
    Dim i           As Integer
    For i = LBound(approvedFunctions) To UBound(approvedFunctions)
        If approvedFunctions(i) = funcName Then
            IsApprovedFunction = True
            Exit Function
        End If
    Next i
    
    IsApprovedFunction = False
    
End Function

Sub ExecuteInstrumentaFunction()
    
    #If Mac Then
        
        Dim macroName As String
        macroName = ReadMacroNameFromFile()
        
        If macroName <> "" Then
            If IsApprovedFunction(macroName) Then
                Application.Run macroName
            Else
                MsgBox "Error: Function '" & macroName & "' is not approved.", vbCritical
            End If
        End If
        
    #End If
    
End Sub


Function ReadMacroNameFromFile() As String
    #If Mac Then
        Dim tempPath As String
        tempPath = MacScript("return posix path of (path To temporary items) As string")
        Dim filePath As String
        filePath = tempPath & "tmpInstrumentaKeys.txt"

        Dim fileNum As Integer
        Dim macroName As String

        If Dir(filePath) <> "" Then
            fileNum = FreeFile
            On Error Resume Next
            Open filePath For Input As fileNum
            If Err.Number <> 0 Then
                MsgBox "Error opening file! " & Err.Description, vbCritical
                Exit Function
            End If
            On Error GoTo 0

            If Not EOF(fileNum) Then
                Line Input #fileNum, macroName
            Else
                macroName = ""
            End If
            Close fileNum
            
        Else
            macroName = ""
        End If

        ReadMacroNameFromFile = macroName
    #End If
End Function


