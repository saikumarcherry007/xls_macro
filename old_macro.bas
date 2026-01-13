'==========================================
' PVT Processing generates PVT_STA
' Author: Saikumar Malluru
' Created: 2025-12-10
' Version: 1.0.0
' Description: Processes PVT data from the 'PVTs' sheet and creates a 'PVT_STA' report.
' Usage: Configure extraction types/focus columns on 'Generate PVT_STA' sheet and click the 'Generate PVT_STA Sheet' button.
' Note: The script supports wildcard matching for focus columns and searches header rows 1 and 2.
'==========================================

'==========================================
' DEFAULT CONFIGURATION - EDIT THESE VALUES TO CHANGE DEFAULTS
'==========================================
' Setup Pattern & Extraction Defaults
Const DEFAULT_SETUP_PATTERN As String = "SS*"
Const DEFAULT_SETUP_EXTRACTION As String = "Cworst_T,RCworst_T"

' Hold Pattern & Extraction Defaults
Const DEFAULT_HOLD_PATTERN As String = "FF*"
Const DEFAULT_HOLD_EXTRACTION As String = "Cbest,Cworst,RCbest,RCworst"

' Typical Pattern & Extraction Defaults
Const DEFAULT_TYPICAL_PATTERN As String = "TT*"
Const DEFAULT_TYPICAL_EXTRACTION As String = "Ctypical"

' Focus Columns Default (comma-separated list, leave empty for no filter)
Const DEFAULT_FOCUS_COLUMNS As String = ""

' Hold-only Mappings Default (semicolon-separated pattern:extraction pairs)
' Example: "SS*:Cbest;SSGNP*:Cbest,Cworst"
Const DEFAULT_HOLD_ONLY_MAPPINGS As String = ""

' Sheet Names
Const SHEET_PVT As String = "PVTs"
Const SHEET_INSTANCES As String = "N3P Instance List"
Const SHEET_OUTPUT As String = "PVT_STA"
Const SHEET_CONFIG As String = "Generate PVT_STA"

'==========================================

'==========================================
' INSTANCE MATCHER HELPER FUNCTIONS
'==========================================
Public gPVTProcessing_Completed As Boolean
Public gPVTProcessing_SuccessMsg As String
Public gPVTProcessing_MsgShown As Boolean

' Parse memory type and WA from header like "HDDP\n(WA=1)"
Private Function ParseMemoryConditionIM(ByVal headerText As String, ByRef memoryType As String, ByRef waValues() As String) As Boolean
    Dim cleanText As String
    Dim parenPos As Long
    Dim waText As String
    
    ParseMemoryCondition = False
    memoryType = ""
    ReDim waValues(0)
    
    cleanText = Replace(headerText, vbLf, " ")
    cleanText = Replace(cleanText, vbCr, " ")
    cleanText = Trim(cleanText)
    
    parenPos = InStr(cleanText, "(")
    If parenPos = 0 Then Exit Function
    
    memoryType = Trim(Left(cleanText, parenPos - 1))
    
    waText = Mid(cleanText, parenPos)
    waText = Replace(waText, "(", "")
    waText = Replace(waText, ")", "")
    waText = Replace(waText, "WA=", "")
    waText = Trim(waText)
    
    If InStr(waText, "/") > 0 Then
        Dim parts() As String
        parts = Split(waText, "/")
        ReDim waValues(LBound(parts) To UBound(parts))
        Dim i As Long
        For i = LBound(parts) To UBound(parts)
            waValues(i) = Trim(parts(i))
        Next i
    Else
        ReDim waValues(0)
        waValues(0) = waText
    End If
    
    ParseMemoryConditionIM = (memoryType <> "" And waValues(0) <> "")
End Function

' Check if instance matches condition
Private Function InstanceMatchesIM(instanceMemType As String, instanceWA As String, _
                                 condMemType As String, condWAValues() As String) As Boolean
    Dim i As Long
    InstanceMatchesIM = False
    
    If UCase(Trim(instanceMemType)) <> UCase(Trim(condMemType)) Then
        Exit Function
    End If
    
    For i = LBound(condWAValues) To UBound(condWAValues)
        If Trim(instanceWA) = Trim(condWAValues(i)) Then
            InstanceMatchesIM = True
            Exit Function
        End If
    Next i
End Function

'==========================================

Private Function GetCellValue(ws As Worksheet, rowNum As Long, colNum As Long) As String
    Dim cell As Range
    Set cell = ws.Cells(rowNum, colNum)
    
    On Error Resume Next
    If cell.MergeCells Then
        GetCellValue = Trim(CStr(cell.MergeArea.Cells(1, 1).Value))
    Else
        GetCellValue = Trim(CStr(cell.Value))
    End If
    
    If IsEmpty(GetCellValue) Or IsNull(GetCellValue) Or GetCellValue = "Error" Then
        GetCellValue = ""
    End If
    On Error GoTo 0
End Function

 
Private Function MatchesWildcard(ByVal text As String, ByVal pattern As String) As Boolean
    Dim textUpper As String
    Dim patternUpper As String
    
    textUpper = UCase(Trim(text))
    patternUpper = UCase(Trim(pattern))
    
    If InStr(patternUpper, "*") > 0 Then
        MatchesWildcard = textUpper Like patternUpper
    Else
        MatchesWildcard = (textUpper = patternUpper)
    End If
End Function

 
Private Function ColLetter(ByVal colNum As Long) As String
    Dim s As String
    Dim modVal As Long
    s = ""
    Do While colNum > 0
        modVal = (colNum - 1) Mod 26
        s = Chr(65 + modVal) & s
        colNum = (colNum - 1) \ 26
    Loop
    ColLetter = s
End Function

Private Function ExtractMemoryTypeFromColumnName(ByVal colName As String) As String
    ' Extract memory type from column name like "HDDP\n(WA=1)"
    ' Returns "HDDP*"
    ' For focus patterns like "hddp*", return as-is
    Dim memType As String
    memType = Trim(UCase(colName))
    
    ' If it already ends with *, it's a focus pattern - return as-is
    If Right(memType, 1) = "*" Then
        ExtractMemoryTypeFromColumnName = memType
        Exit Function
    End If
    
    ' Remove newline and everything after (
    Dim pos As Long
    pos = InStr(memType, vbLf)
    If pos > 0 Then
        memType = Left(memType, pos - 1)
    End If
    
    pos = InStr(memType, "(")
    If pos > 0 Then
        memType = Left(memType, pos - 1)
    End If
    
    memType = Trim(memType) & "*"
    ExtractMemoryTypeFromColumnName = memType
End Function



Private Function FindAllColumnIndices(ws As Worksheet, ByVal headerRow As Long, ByVal colNamePattern As String) As Collection
    Dim col As Long
    Dim headerValue As String
    Dim lastCol As Long
    Dim matchedCols As Collection
    
    Set matchedCols = New Collection
    
    lastCol = ws.Cells(headerRow, ws.Columns.Count).End(xlToLeft).Column
    
    For col = 1 To lastCol
        headerValue = Trim(CStr(ws.Cells(headerRow, col).Value))
        If headerValue <> "" And MatchesWildcard(headerValue, colNamePattern) Then
            matchedCols.Add col
        End If
    Next col
    
    Set FindAllColumnIndices = matchedCols
End Function

Private Function MatchesPatternList(ByVal text As String, patterns As Variant) As Boolean
    Dim i As Long
    Dim pat As String
    MatchesPatternList = False
    If text = "" Then Exit Function
    If IsEmpty(patterns) Then Exit Function
    For i = LBound(patterns) To UBound(patterns)
        pat = Trim(CStr(patterns(i)))
        If pat <> "" Then
            If MatchesWildcard(text, pat) Then
                MatchesPatternList = True
                Exit Function
            End If
        End If
    Next i
End Function

Private Function FindHoldOverrideExtractions(ByVal pvtName As String, holdOverrideArr As Variant) As Variant
    Dim i As Long
    Dim parts As Variant
    Dim pat As String
    Dim exStr As String
    Dim exArr As Variant
    FindHoldOverrideExtractions = Empty
    If IsEmpty(holdOverrideArr) Then Exit Function
    If UBound(holdOverrideArr) < LBound(holdOverrideArr) Then Exit Function
    For i = LBound(holdOverrideArr) To UBound(holdOverrideArr)
        If Trim(CStr(holdOverrideArr(i))) <> "" Then
            parts = Split(holdOverrideArr(i), ":")
            If UBound(parts) >= 1 Then
                pat = Trim(parts(0))
                exStr = Trim(parts(1))
                If pat <> "" And exStr <> "" Then
                    If MatchesWildcard(pvtName, pat) Then
                        exArr = Split(exStr, ",")
                        FindHoldOverrideExtractions = exArr
                        Exit Function
                    End If
                End If
            End If
        End If
    Next i
End Function

'' ExtractLabelFromPattern removed: no longer needed since labels are static

Private Sub UpdatePatternLabels(ws As Worksheet)
    Dim cur As String
    cur = Trim(CStr(ws.Range("B12").Value))
    If cur = "" Or InStr(1, cur, "(Setup)", vbTextCompare) > 0 Then
        ws.Range("B12").Value = "(Setup) patterns & extraction types"
    End If
    cur = Trim(CStr(ws.Range("B13").Value))
    If cur = "" Or InStr(1, cur, "(Hold)", vbTextCompare) > 0 Then
        ws.Range("B13").Value = "(Hold) patterns & extraction types"
    End If
    cur = Trim(CStr(ws.Range("B14").Value))
    If cur = "" Or InStr(1, cur, "(Typical)", vbTextCompare) > 0 Then
        ws.Range("B14").Value = "(Typical) patterns & extraction types"
    End If
End Sub

Public Sub RefreshPatternLabels()
    On Error Resume Next
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(SHEET_CONFIG)
    If ws Is Nothing Then Exit Sub
    UpdatePatternLabels ws
    On Error GoTo 0
End Sub

 
Private Function MatchesFocusFilter(ws As Worksheet, ByVal rowNum As Long, focusColOrder As Collection, focusColMapping As Object) As Boolean
    Dim i As Long
    Dim colName As String
    Dim sourceColIndex As Long
    Dim cellValue As String

    MatchesFocusFilter = True

    If focusColOrder Is Nothing Then
        Exit Function
    End If
    If focusColOrder.Count = 0 Then
        Exit Function
    End If

    For i = 1 To focusColOrder.Count
        colName = focusColOrder(i)
        If focusColMapping.Exists(colName) Then
            sourceColIndex = focusColMapping(colName)
            cellValue = UCase(GetCellValue(ws, rowNum, sourceColIndex))
            If cellValue <> "YES" Then
                MatchesFocusFilter = False
                Exit Function
            End If
        End If
    Next i
End Function

 
Sub CreatePVTSTASheet()
    Dim wsPVT As Worksheet
    Dim btn As Button
    Dim shpProcess As Shape
    
    
    On Error Resume Next
    Set wsPVT = ThisWorkbook.Sheets(SHEET_CONFIG)
    On Error GoTo 0
    If wsPVT Is Nothing Then
    
        Set wsPVT = ThisWorkbook.Sheets.Add(Before:=ThisWorkbook.Sheets(1))
        wsPVT.Name = SHEET_CONFIG
    End If
    
    
    With wsPVT
        
        .Range("A1:G1").Merge
        .Range("A1").Value = "PVT Data Processing Tool"
        .Range("A1").Font.Size = 20
        .Range("A1").Font.Bold = True
        .Range("A1").Font.Color = RGB(255, 255, 255)
        .Range("A1").Interior.Color = RGB(68, 114, 196)
        .Range("A1").HorizontalAlignment = xlCenter
        .Range("A1").VerticalAlignment = xlCenter
        .Range("A1").RowHeight = 35
        
        
        .Range("A3:G3").Merge
        .Range("A3").Value = "Click the button below to process PVT data and generate the PVT_STA sheet."
        .Range("A3").Font.Size = 11
        .Range("A3").WrapText = True
        .Range("A3").HorizontalAlignment = xlCenter
        .Range("A3").RowHeight = 30
        
    
    .Range("B5:F5").Merge
    .Range("B5").Value = "This process will:"
    .Range("B5").Font.Bold = True
    .Range("B5").Font.Size = 13
    .Range("B5").Font.Color = RGB(255, 255, 255)
    .Range("B5").Interior.Color = RGB(68, 114, 196)
    .Range("B5").HorizontalAlignment = xlLeft
    .Range("B5").IndentLevel = 1
    .Range("B5").VerticalAlignment = xlCenter
        
    
    .Range("B6:B9").Value = "" ' clear any prior icons
    .Range("B6").Value = ChrW(10003)
    .Range("B7").Value = ChrW(10003)
    .Range("B8").Value = ChrW(10003)
    .Range("B9").Value = ChrW(10003)
    .Range("B6:B9").Font.Size = 12
    .Range("B6:B9").Font.Color = RGB(56, 118, 29) ' green checks
    .Range("B6:B9").HorizontalAlignment = xlCenter
    .Range("C6").Value = "Read data from PVTs Sheet"
    .Range("C7").Value = "Process all PVT corners (SETUP and HOLD)"
    .Range("C8").Value = "Create organized output in PVT_STA sheet"
    .Range("C9").Value = "Apply formatting and filters automatically"
        
        
        Dim i As Integer
        For i = 6 To 9
            .Cells(i, 3).Font.Size = 11
            .Cells(i, 3).Font.Color = RGB(60, 60, 60)
            .Cells(i, 3).Font.Name = "Calibri"
            .Cells(i, 3).IndentLevel = 1
        Next i
        
        
        .Range("A10").RowHeight = 8
        
    
    .Range("B5:F9").Borders.LineStyle = xlContinuous
    .Range("B5:F9").Borders.Color = RGB(191, 191, 191)
    .Range("B6:F9").Interior.Color = RGB(249, 250, 252) ' very light gray background for context
    .Range("B5:F5").Borders(xlEdgeBottom).Weight = xlMedium
    .Range("B6:F9").HorizontalAlignment = xlLeft
        
    
    .Range("A11").RowHeight = 20
    .Range("A12").RowHeight = 30
    .Range("A13").RowHeight = 30
    .Range("A14").RowHeight = 30
    .Range("A15").RowHeight = 30
    .Range("A16").RowHeight = 30
        
    
    On Error Resume Next
    .Buttons("ProcessPVTButton").Delete
    .Shapes("ProcessPVTButton").Delete
    .Shapes("ProcessPVTShape").Delete
    On Error GoTo 0

    
    On Error Resume Next
    .Range("B19:F21").ClearContents
    .Range("B19:F21").UnMerge
    .Range("B19:F21").Borders.LineStyle = xlNone
    .Range("B19:F21").Interior.pattern = xlNone
    
    .Range("B22:F22").ClearContents
    .Range("B22:F22").UnMerge
    .Range("B22:F22").Borders.LineStyle = xlNone
    .Range("B22:F22").Interior.pattern = xlNone
    On Error GoTo 0

    
    
    If Trim(CStr(.Range("B11").Value)) = "" Then .Range("B11").Value = "Default Configurations: "
    .Range("B11:F11").Merge
    .Range("B11").Font.Bold = True
    .Range("B11").Font.Size = 11
    .Range("B11").HorizontalAlignment = xlLeft
    
    
    If Trim(CStr(.Range("B12").Value)) = "" Then .Range("B12").Value = "(Setup) patterns & extraction types"
    .Range("B12").Font.Bold = True
    ' Pattern input first, extraction types second (patterns left, extraction right)
    If Trim(CStr(.Range("C12").Value)) = "" Then .Range("C12").Value = DEFAULT_SETUP_PATTERN
    .Range("C12:D12").Merge
    .Range("C12:D12").WrapText = True
    .Range("C12:D12").ShrinkToFit = True
    .Range("C12:D12").HorizontalAlignment = xlLeft
    .Range("C12:D12").VerticalAlignment = xlCenter
    If Trim(CStr(.Range("E12").Value)) = "" Then .Range("E12").Value = DEFAULT_SETUP_EXTRACTION
    .Range("E12:F12").Merge
    .Range("E12:F12").WrapText = True
    .Range("E12:F12").ShrinkToFit = True
    .Range("E12:F12").HorizontalAlignment = xlLeft
    .Range("E12:F12").VerticalAlignment = xlCenter

    If Trim(CStr(.Range("B13").Value)) = "" Then .Range("B13").Value = "(Hold) patterns & extraction types"
    .Range("B13").Font.Bold = True
    If Trim(CStr(.Range("C13").Value)) = "" Then .Range("C13").Value = DEFAULT_HOLD_PATTERN
    .Range("C13:D13").Merge
    .Range("C13:D13").WrapText = True
    .Range("C13:D13").ShrinkToFit = True
    .Range("C13:D13").HorizontalAlignment = xlLeft
    .Range("C13:D13").VerticalAlignment = xlCenter
    If Trim(CStr(.Range("E13").Value)) = "" Then .Range("E13").Value = DEFAULT_HOLD_EXTRACTION
    .Range("E13:F13").Merge
    .Range("E13:F13").WrapText = True
    .Range("E13:F13").ShrinkToFit = True
    .Range("E13:F13").HorizontalAlignment = xlLeft
    .Range("E13:F13").VerticalAlignment = xlCenter

    If Trim(CStr(.Range("B14").Value)) = "" Then .Range("B14").Value = "(Typical) patterns & extraction types"
    .Range("B14").Font.Bold = True
    If Trim(CStr(.Range("C14").Value)) = "" Then .Range("C14").Value = DEFAULT_TYPICAL_PATTERN
    .Range("C14:D14").Merge
    .Range("C14:D14").WrapText = True
    .Range("C14:D14").ShrinkToFit = True
    .Range("C14:D14").HorizontalAlignment = xlLeft
    .Range("C14:D14").VerticalAlignment = xlCenter
    If Trim(CStr(.Range("E14").Value)) = "" Then .Range("E14").Value = DEFAULT_TYPICAL_EXTRACTION
    .Range("E14:F14").Merge
    .Range("E14:F14").WrapText = True
    .Range("E14:F14").ShrinkToFit = True
    .Range("E14:F14").HorizontalAlignment = xlLeft
    .Range("E14:F14").VerticalAlignment = xlCenter

    
    If Trim(CStr(.Range("B15").Value)) = "" Then .Range("B15").Value = "Focus Columns (filter)"
    .Range("B15").Font.Bold = True
    If Trim(CStr(.Range("C15").Value)) = "" Then .Range("C15").Value = DEFAULT_FOCUS_COLUMNS
    .Range("C15:F15").Merge
    .Range("C15:F15").WrapText = True
    .Range("C15:F15").ShrinkToFit = True
    .Range("C15:F15").HorizontalAlignment = xlLeft
    .Range("C15:F15").VerticalAlignment = xlCenter

    ' Hold-only mappings: pattern:extraction1,extraction2;pattern2:extrA,extrB
    If Trim(CStr(.Range("B16").Value)) = "" Then .Range("B16").Value = "Hold-only mappings"
    .Range("B16").Font.Bold = True
    If Trim(CStr(.Range("C16").Value)) = "" Then .Range("C16").Value = DEFAULT_HOLD_ONLY_MAPPINGS ' e.g. SS*:Cbest;SSGNP*:Cbest,Cworst"
    .Range("C16:F16").Merge
    .Range("C16:F16").WrapText = True
    .Range("C16:F16").ShrinkToFit = True
    .Range("C16:F16").HorizontalAlignment = xlLeft
    .Range("C16:F16").VerticalAlignment = xlCenter

    ' Patterns will be entered in C12:C14 (left) alongside extraction inputs in E12:E14 (right)
    .Range("C15:F15").HorizontalAlignment = xlLeft
    .Range("C15:F15").VerticalAlignment = xlCenter
    
    .Range("B12:F16").Borders.LineStyle = xlContinuous
    
    .Range("B12:B16").Interior.Color = RGB(221, 235, 255) ' light blue for labels
    .Range("C12:F16").Interior.Color = RGB(255, 255, 255) ' white input background
    .Range("B12:B16").Font.Color = RGB(68, 114, 196)
    .Range("C12:F16").Borders.Color = RGB(191, 191, 191)
    .Range("B12:B16").Borders.Color = RGB(191, 191, 191)
    .Range("B12:B16").Font.Bold = True
    .Range("B12:B16").WrapText = True
    .Range("B12:B16").HorizontalAlignment = xlLeft
    .Range("B12:B16").VerticalAlignment = xlCenter

    
    
    If Trim(CStr(.Range("B17").Value)) = "" Then .Range("B17").Value = "NOTE: You can override default values above. Enter comma-separated PVT name patterns (left) and extraction types (right). Examples: SS*, FF*, TT*. To force a PVT into HOLD with custom extractions use semicolon-separated mappings in 'Hold-only mappings' (e.g.: SS*:Cbest;SSGNP*:Cbest,Cworst). Wildcards supported."
    .Range("B17:F17").Merge
    .Range("B17").Font.Size = 9
    .Range("B17").Font.Color = RGB(100, 100, 100)
    .Range("B17").HorizontalAlignment = xlLeft
    .Range("B17").WrapText = True
    .Range("A17").RowHeight = 30
    
    .Range("A18").RowHeight = 12
    
    
    ' Set column widths first to ensure consistent positioning
    .Columns("A").ColumnWidth = 2
    .Columns("B").ColumnWidth = 35
    .Columns("C:F").ColumnWidth = 30
    .Columns("G").ColumnWidth = 2
    
    ' Set row height for button area
    .Range("A19").RowHeight = 60
    .Range("A20").RowHeight = 60
    
    Dim btnLeft As Double, btnTop As Double, btnWidth As Double, btnHeight As Double
    btnWidth = 240 ' Width in points - consistent smaller size
    btnHeight = 40 ' Height in points - consistent smaller size
    
    ' Calculate center position using actual column positions (now that widths are set)
    Dim centerPoint As Double
    centerPoint = .Range("B19").Left + (.Range("F19").Left + .Range("F19").Width - .Range("B19").Left) / 2
    btnLeft = centerPoint - (btnWidth / 2)
    btnTop = .Range("A19").Top + 10 ' 10 points padding from top of row 19

        Set shpProcess = .Shapes.AddShape(msoShapeRoundedRectangle, btnLeft, btnTop, btnWidth, btnHeight)
            With shpProcess
                .Name = "ProcessPVTShape"
                
                .Fill.ForeColor.RGB = RGB(68, 114, 196)
                .Fill.Transparency = 0
                .Line.Visible = msoFalse
                .Shadow.Type = msoShadow6
                
                .TextFrame.Characters.text = "Generate " & SHEET_OUTPUT & " Sheet"
                .TextFrame.Characters.Font.Size = 11
                .TextFrame.Characters.Font.Bold = True
                .TextFrame.HorizontalAlignment = xlHAlignCenter
                .TextFrame.VerticalAlignment = xlVAlignCenter
                .TextFrame.Characters.Font.Color = RGB(255, 255, 255)
                .OnAction = "RunPVTProcessing_Testing" ' Use wrapper to avoid ambiguous macro name when duplicate modules exist
                .Placement = xlFreeFloating ' Keep button size fixed, don't resize with cells
            End With
        
        
        ActiveWindow.DisplayGridlines = False
    End With
    
        ' Update labels to reflect patterns (won't overwrite custom labels)
        On Error Resume Next
        UpdatePatternLabels wsPVT
        On Error GoTo 0

        wsPVT.Activate
        wsPVT.Range("A1").Select
End Sub

 
Sub RunPVTProcessing_Testing()
    On Error GoTo Cleanup
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    
    ' Initialize completion flags (used as a robust fallback to show completion after UI restored)
    gPVTProcessing_Completed = False
    gPVTProcessing_SuccessMsg = ""
    gPVTProcessing_MsgShown = False

    ProcessPVTData_Final
Cleanup:
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    ' If the procedure completed but its MsgBox may have been suppressed by UI state, show a fallback
    If gPVTProcessing_Completed And Not gPVTProcessing_MsgShown Then
        MsgBox gPVTProcessing_SuccessMsg, vbInformation, "Processing Complete (wrapper)"
        gPVTProcessing_MsgShown = True
    End If
End Sub

Private Sub ProcessPVTData_Final()
    On Error GoTo ErrorHandler
    
    
    
    
    Dim wsSource As Worksheet
    Dim wsOutput As Worksheet
    Dim wsPVT As Worksheet
    Dim lastRow As Long
    Dim outputRow As Integer
    Dim i As Long
    Dim pvtName As String
    Dim vddp As Variant
    Dim vdda As Variant
    Dim extractionType As String
    Dim bandName As String
    Dim fastestRMValue As String
    
    Dim setupExtractionStr As String
    Dim holdExtractionStr As String
    Dim typicalExtractionStr As String
    Dim setupExtractionArr As Variant
    Dim holdExtractionArr As Variant
    Dim typicalExtractionArr As Variant
    Dim setupPatternStr As String
    Dim holdPatternStr As String
    Dim typicalPatternStr As String
    Dim setupPatternArr As Variant
    Dim holdPatternArr As Variant
    Dim typicalPatternArr As Variant
    Dim focusColumnsStr As String
    Dim focusColumnsArr As Variant
    Dim hasFocusFilter As Boolean
    
    Dim dataCol As Integer
    Dim colIdx As Integer
    Dim headerCol As Integer
    Dim lastCol As String
    Dim lastColIndex As Long
    Dim col As Integer
    Dim counter As Integer
    Dim focusCol As String
    Dim holdOverrideTotal As Integer
    Dim holdOverrideMatchesCount As Integer
    
    Dim patternValidationMsg As String
    Dim patIdx As Integer
    Dim testPat As String
    Dim patMatched As Boolean
    Dim testRow As Long
    Dim testPVT As String
    Dim patternsValid As Boolean
    Dim focusColMapping As Object
    Dim focusColOrder As Collection
    Dim pattern As String
    Dim matchedCols As Collection
    Dim foundColIndex As Long
    Dim foundColName As String
    Dim headerRow As Long
    Dim matchIdx As Integer
    Dim focusColOrderOutput As Collection
    Dim condition1Cols As Collection
    Dim condition2Cols As Collection
    Dim otherCols As Collection
    Dim colName As String
    Dim bandNames As Collection
    Dim bandDict As Object
    Dim currentPVT As String
    Dim currentProcess As String
    Dim isSETUP As Boolean
    Dim holdOverrideEx As Variant
    Dim extractionTypes As Variant
    Dim setupEndRow As Integer
    Dim setupInstanceListCol As Integer
    Dim holdStartRow As Integer
    Dim isHOLD As Boolean
    Dim colHeaderName As String
    Dim sourceColIndex As Long
    Dim holdOverrideStr As String
    Dim holdOverrideArr As Variant
    Dim hoIdx As Integer
    Dim hoParts As Variant
    Dim hoPat As String
    Dim holdOverrideEx2 As Variant
    Dim currentBand As String
    Dim previousBandHadData As Boolean
    Dim bandIndex As Integer
    Dim thisBandStartRow As Integer
    Dim bandStartRow As Long
    Dim bandEndRow As Long
    Dim j As Integer
    Dim debugMsg As String
    Dim debugIdx As Integer
    Dim cond1StartCol As Integer
    Dim cond1EndCol As Integer
    Dim cond2StartCol As Integer
    Dim cond2EndCol As Integer
    Dim holdExtractionTypes As Variant
    Dim kInt As Integer
    Dim holdInstanceListCol As Integer
    Dim setupCount As Integer
    Dim holdCount As Integer
    Dim yesRange As Range
    Dim focusStartCol As String
    Dim focusEndCol As String
    Dim rshade1 As Long, rshade2 As Long, r2 As Long
    Dim bandRow As Integer
    Dim currentBandName As String
    Dim bandColor As Long
    Dim colorIndex As Integer
    Dim bandCell As Range
    Dim sectionCell As Range
    Dim bandColor055 As Long
    Dim patternsStatus As String
    Dim focusStatus As String
    Dim holdMappingStatus As String
    Dim successMsg As String
    Dim errMsg As String
    
    
    On Error Resume Next
    Set wsSource = ThisWorkbook.Sheets(SHEET_PVT)
    If wsSource Is Nothing Then
        
        Exit Sub
    End If
    On Error GoTo 0

    
    setupExtractionStr = DEFAULT_SETUP_EXTRACTION
    holdExtractionStr = DEFAULT_HOLD_EXTRACTION
    typicalExtractionStr = DEFAULT_TYPICAL_EXTRACTION
    focusColumnsStr = DEFAULT_FOCUS_COLUMNS
    hasFocusFilter = (DEFAULT_FOCUS_COLUMNS <> "")
    On Error Resume Next
    Set wsPVT = ThisWorkbook.Sheets(SHEET_CONFIG)
    If Not wsPVT Is Nothing Then
    If Trim(CStr(wsPVT.Range("E12").Value)) <> "" Then setupExtractionStr = CStr(wsPVT.Range("E12").Value)
    If Trim(CStr(wsPVT.Range("E13").Value)) <> "" Then holdExtractionStr = CStr(wsPVT.Range("E13").Value)
    If Trim(CStr(wsPVT.Range("E14").Value)) <> "" Then typicalExtractionStr = CStr(wsPVT.Range("E14").Value)
        If Trim(CStr(wsPVT.Range("C15").Value)) <> "" Then
            focusColumnsStr = CStr(wsPVT.Range("C15").Value)
            hasFocusFilter = True
        End If
        ' Read PVT name patterns
        If Trim(CStr(wsPVT.Range("C12").Value)) <> "" Then
            setupPatternStr = CStr(wsPVT.Range("C12").Value)
        Else
            setupPatternStr = DEFAULT_SETUP_PATTERN
        End If
        If Trim(CStr(wsPVT.Range("C13").Value)) <> "" Then
            holdPatternStr = CStr(wsPVT.Range("C13").Value)
        Else
            holdPatternStr = DEFAULT_HOLD_PATTERN
        End If
        If Trim(CStr(wsPVT.Range("C14").Value)) <> "" Then
            typicalPatternStr = CStr(wsPVT.Range("C14").Value)
        Else
            typicalPatternStr = DEFAULT_TYPICAL_PATTERN
        End If
        ' Read optional hold-only mappings
        holdOverrideStr = DEFAULT_HOLD_ONLY_MAPPINGS
        If Trim(CStr(wsPVT.Range("C16").Value)) <> "" Then
            holdOverrideStr = CStr(wsPVT.Range("C16").Value)
        End If
        If holdOverrideStr <> "" Then
            holdOverrideArr = Split(holdOverrideStr, ";")
        Else
            ReDim holdOverrideArr(-1) ' empty
        End If
    End If
    On Error GoTo 0

    setupExtractionArr = Split(setupExtractionStr, ",")
    holdExtractionArr = Split(holdExtractionStr, ",")
    typicalExtractionArr = Split(typicalExtractionStr, ",")
    setupPatternArr = Split(setupPatternStr, ",")
    holdPatternArr = Split(holdPatternStr, ",")
    typicalPatternArr = Split(typicalPatternStr, ",")
    ' holdOverrideArr as variant array of strings like "pattern:extr1,extr2"
    
    ' Validate patterns against actual PVT names in source data
    lastRow = wsSource.Cells(wsSource.Rows.Count, "C").End(xlUp).Row
    patternValidationMsg = ""
    For patIdx = LBound(setupPatternArr) To UBound(setupPatternArr)
        testPat = Trim(CStr(setupPatternArr(patIdx)))
        If testPat <> "" Then
            patMatched = False
            For testRow = 2 To lastRow
                testPVT = Trim(CStr(wsSource.Cells(testRow, 3).Value))
                If testPVT <> "" And MatchesWildcard(testPVT, testPat) Then
                    patMatched = True
                    Exit For
                End If
            Next testRow
            If Not patMatched Then
                patternValidationMsg = patternValidationMsg & "  • Setup pattern: '" & testPat & "'" & vbCrLf
            End If
        End If
    Next patIdx
    
    ' Validate Hold patterns
    For patIdx = LBound(holdPatternArr) To UBound(holdPatternArr)
        testPat = Trim(CStr(holdPatternArr(patIdx)))
        If testPat <> "" Then
            patMatched = False
            For testRow = 2 To lastRow
                testPVT = Trim(CStr(wsSource.Cells(testRow, 3).Value))
                If testPVT <> "" And MatchesWildcard(testPVT, testPat) Then
                    patMatched = True
                    Exit For
                End If
            Next testRow
            If Not patMatched Then
                patternValidationMsg = patternValidationMsg & "  • Hold pattern: '" & testPat & "'" & vbCrLf
            End If
        End If
    Next patIdx
    
    ' Validate Typical patterns
    For patIdx = LBound(typicalPatternArr) To UBound(typicalPatternArr)
        testPat = Trim(CStr(typicalPatternArr(patIdx)))
        If testPat <> "" Then
            patMatched = False
            For testRow = 2 To lastRow
                testPVT = Trim(CStr(wsSource.Cells(testRow, 3).Value))
                If testPVT <> "" And MatchesWildcard(testPVT, testPat) Then
                    patMatched = True
                    Exit For
                End If
            Next testRow
            If Not patMatched Then
                patternValidationMsg = patternValidationMsg & "  • Typical pattern: '" & testPat & "'" & vbCrLf
            End If
        End If
    Next patIdx
    
    ' Validate Hold-only mapping patterns
    If Not IsEmpty(holdOverrideArr) Then
        If UBound(holdOverrideArr) >= LBound(holdOverrideArr) Then
            holdOverrideTotal = 0
            holdOverrideMatchesCount = 0
            For hoIdx = LBound(holdOverrideArr) To UBound(holdOverrideArr)
                If Trim(CStr(holdOverrideArr(hoIdx))) <> "" Then
                    hoParts = Split(holdOverrideArr(hoIdx), ":")
                    If UBound(hoParts) >= 1 Then
                        hoPat = Trim(hoParts(0))
                        If hoPat <> "" Then
                            holdOverrideTotal = holdOverrideTotal + 1
                            patMatched = False
                            For testRow = 2 To lastRow
                                testPVT = Trim(CStr(wsSource.Cells(testRow, 3).Value))
                                If testPVT <> "" And MatchesWildcard(testPVT, hoPat) Then
                                    patMatched = True
                                    holdOverrideMatchesCount = holdOverrideMatchesCount + 1
                                    Exit For
                                End If
                            Next testRow
                            If Not patMatched Then
                                patternValidationMsg = patternValidationMsg & "  • Hold-only mapping pattern: '" & hoPat & "'" & vbCrLf
                            End If
                        End If
                    End If
                End If
            Next hoIdx
        End If
    End If
    
    ' Show warning if any patterns don't match
    If patternValidationMsg <> "" Then
        If MsgBox("WARNING: The following patterns do not match any PVT names in the source data:" & vbCrLf & vbCrLf & _
               patternValidationMsg & vbCrLf & _
               "Please verify your pattern entries in the configuration." & vbCrLf & vbCrLf & _
               "Click OK to continue anyway, or Cancel to stop.", vbExclamation + vbOKCancel, "Pattern Validation Warning") = vbCancel Then
            Exit Sub
        End If
    End If

    ' patternsValid flag
    patternsValid = (patternValidationMsg = "")
    
    
    Set focusColMapping = CreateObject("Scripting.Dictionary")
    Set focusColOrder = New Collection
    
    If hasFocusFilter Then
        focusColumnsArr = Split(focusColumnsStr, ",")
        
        
        For colIdx = LBound(focusColumnsArr) To UBound(focusColumnsArr)
            pattern = Trim(focusColumnsArr(colIdx))
            
            
            Set matchedCols = FindAllColumnIndices(wsSource, 1, pattern)
            
            
            If matchedCols.Count = 0 Then
                Set matchedCols = FindAllColumnIndices(wsSource, 2, pattern)
                headerRow = 2
            Else
                headerRow = 1
            End If
            
            
            For matchIdx = 1 To matchedCols.Count
                foundColIndex = matchedCols(matchIdx)
                
                foundColName = Trim(UCase(CStr(wsSource.Cells(headerRow, foundColIndex).Value)))
                
                
                If Not focusColMapping.Exists(foundColName) Then
                    focusColMapping.Add foundColName, foundColIndex
                    focusColOrder.Add foundColName
                End If
            Next matchIdx
            
            
            If matchedCols.Count = 0 Then
                MsgBox "WARNING: No focus columns found matching pattern '" & pattern & "'" & vbCrLf & vbCrLf & _
                       "Pattern: " & pattern & vbCrLf & _
                       "Searched in: PVTs sheet (rows 1 and 2)" & vbCrLf & vbCrLf & _
                       "Please verify:" & vbCrLf & _
                       "  • Column names exist in header rows 1 or 2" & vbCrLf & _
                       "  • Pattern spelling and wildcards are correct" & vbCrLf & _
                       "  • Column names match exactly (case-insensitive)", vbExclamation, "Focus Column Not Found"
            End If
        Next colIdx
        
        
        If focusColOrder.Count = 0 Then
            MsgBox "ERROR: No focus columns were found!" & vbCrLf & _
                   "You entered: " & focusColumnsStr & vbCrLf & vbCrLf & _
                   "Please verify the column names in rows 1 or 2 of the PVTs sheet match your input." & vbCrLf & _
                   "Macro execution stopped.", vbCritical
            Exit Sub
        Else
            
            debugMsg = "Found " & focusColOrder.Count & " focus column(s):" & vbCrLf
            For debugIdx = 1 To focusColOrder.Count
                debugMsg = debugMsg & "  - " & focusColOrder(debugIdx) & vbCrLf
            Next debugIdx
            ' Focus Columns Discovered popup removed per user request (no MsgBox displayed)
        End If
    End If
    
    
    On Error Resume Next
    Application.DisplayAlerts = False
    ThisWorkbook.Sheets(SHEET_OUTPUT).Delete
    Application.DisplayAlerts = True
    On Error GoTo 0
    
    Set wsOutput = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
    wsOutput.Name = SHEET_OUTPUT
    
    ' Use a single header row for PVT_STA (no extra condition header row)
    
    With wsOutput
        ' Write main headers to single header row (row 1)
        .Range("A1").Value = "Band"
        .Range("B1").Value = "Section"
        .Range("C1").Value = "PVTs for timing closure"
        .Range("D1").Value = "Extraction Corner"
        .Range("E1").Value = "VDDP(V)"
        .Range("F1").Value = "VDDA (V)"
        .Range("G1").Value = "RM"
        
        
    
        headerCol = 8 ' Start after RM column (G)
        
        ' Create a new ordered collection for output that matches the reordered header structure
        Set focusColOrderOutput = New Collection
        
        If hasFocusFilter And focusColOrder.Count > 0 Then
            ' Organize columns by condition groups
            Set condition1Cols = New Collection
            Set condition2Cols = New Collection
            Set otherCols = New Collection
            
            ' Classify columns by condition
            For i = 1 To focusColOrder.Count
                colName = UCase(Trim(focusColOrder(i)))
                
                If colName = "ROM" Or colName = "TCAM" Then
                    condition1Cols.Add focusColOrder(i)
                ElseIf colName = "DR0" Or colName = "DR1" Then
                    condition2Cols.Add focusColOrder(i)
                Else
                    otherCols.Add focusColOrder(i)
                End If
            Next i
            
            ' Build the output order: Condition 1, then Condition 2, then others
            For i = 1 To condition1Cols.Count
                focusColOrderOutput.Add condition1Cols(i)
            Next i
            For i = 1 To condition2Cols.Count
                focusColOrderOutput.Add condition2Cols(i)
            Next i
            For i = 1 To otherCols.Count
                focusColOrderOutput.Add otherCols(i)
            Next i
            
            ' Track start columns for condition headers
            
            ' Write Condition 1 columns (ROM, TCAM)
            If condition1Cols.Count > 0 Then
                cond1StartCol = headerCol
                For i = 1 To condition1Cols.Count
                    .Cells(1, headerCol).Value = condition1Cols(i)
                    headerCol = headerCol + 1
                Next i
                cond1EndCol = headerCol - 1
            End If
            
            ' Write Condition 2 columns (DR0, DR1)
            If condition2Cols.Count > 0 Then
                cond2StartCol = headerCol
                For i = 1 To condition2Cols.Count
                    .Cells(1, headerCol).Value = condition2Cols(i)
                    headerCol = headerCol + 1
                Next i
                cond2EndCol = headerCol - 1
            End If
            
            ' Write other columns
            For i = 1 To otherCols.Count
                .Cells(1, headerCol).Value = otherCols(i)
                headerCol = headerCol + 1
            Next i
            
' No separate condition header row is needed for PVT_STA; memory-type headers are single-row only
        End If
        
        
        .Cells(1, headerCol).Value = "Instance List"
        .Cells(1, headerCol + 1).Value = "PNR"
        .Cells(1, headerCol + 2).Value = "STA"
        .Cells(1, headerCol + 3).Value = "IR"
        
        
    ' Use a single header row (row 1) for all headers
    .Cells(1, 1).Value = "Band"
    .Cells(1, 2).Value = "Section"
    .Cells(1, 3).Value = "PVTs for timing closure"
    .Cells(1, 4).Value = "Extraction Corner"
    .Cells(1, 5).Value = "VDDP(V)"
    .Cells(1, 6).Value = "VDDA (V)"
    .Cells(1, 7).Value = "RM"

    .Cells(1, headerCol).Value = "Instance List"
    .Cells(1, headerCol + 1).Value = "PNR"
    .Cells(1, headerCol + 2).Value = "STA"
    .Cells(1, headerCol + 3).Value = "IR"

    ' Merge each header column vertically (rows 1-2) and format with blue color
    Application.DisplayAlerts = False
    For col = 1 To (headerCol + 3)
        .Range(.Cells(1, col), .Cells(2, col)).Merge
        .Cells(1, col).Interior.Color = RGB(68, 114, 196) ' Professional blue
        .Cells(1, col).Font.Color = RGB(255, 255, 255) ' White text
        .Cells(1, col).Font.Bold = True
        .Cells(1, col).Font.Size = 11
        .Cells(1, col).HorizontalAlignment = xlCenter
        .Cells(1, col).VerticalAlignment = xlCenter
        .Cells(1, col).Borders.LineStyle = xlContinuous
        .Cells(1, col).Borders.Weight = xlMedium
        .Cells(1, col).WrapText = True
    Next col
    Application.DisplayAlerts = True
    
    ' Set row height for merged header rows
    .Rows("1:2").RowHeight = 28
    End With
    
    
    ' lastRow already assigned earlier during pattern validation
    
    
    Set bandNames = New Collection

    ' Apply AutoFilter on merged header row (row 1) so dropdowns appear on the header cells
    wsOutput.Range("A1:" & ColLetter(headerCol + 3) & "1").AutoFilter
    
    Set bandDict = CreateObject("Scripting.Dictionary")
    
    For i = 2 To lastRow
        currentBand = Trim(CStr(wsSource.Cells(i, 2).Value)) ' Column B - BAND
        If currentBand <> "" And Not bandDict.Exists(currentBand) Then
            bandDict.Add currentBand, True
            bandNames.Add currentBand
        End If
    Next i
    
    
    outputRow = 3  ' Start at row 3 (rows 1-2 = header rows with same blue color)
    
    
    previousBandHadData = False
    
    
    For bandIndex = 1 To bandNames.Count
        bandName = bandNames(bandIndex)
        
        
        If bandIndex > 1 And previousBandHadData Then
            outputRow = outputRow + 1
        End If
        
        
        thisBandStartRow = outputRow
    
    
        bandStartRow = 0
        bandEndRow = 0
        
        
        For i = 2 To lastRow
            currentBand = Trim(CStr(wsSource.Cells(i, 2).Value)) ' Column B - BAND
            
            If currentBand = bandName And bandStartRow = 0 Then
                bandStartRow = i
            ElseIf currentBand <> "" And currentBand <> bandName And bandStartRow > 0 Then
                bandEndRow = i - 1
                Exit For
            End If
        Next i
        
        
        If bandEndRow = 0 And bandStartRow > 0 Then
            bandEndRow = lastRow
        End If
        
        
        For i = bandStartRow To bandEndRow
        currentPVT = Trim(CStr(wsSource.Cells(i, 3).Value)) ' Column C - PVT Name
        currentProcess = Trim(CStr(wsSource.Cells(i, 4).Value)) ' Column D - Process
        
        
        
    isSETUP = False
    holdOverrideEx = FindHoldOverrideExtractions(currentPVT, holdOverrideArr)
        
        If currentPVT <> "" Then
            
                If MatchesPatternList(currentPVT, setupPatternArr) Or MatchesPatternList(currentPVT, typicalPatternArr) Then
                    isSETUP = True
                End If
            
            
            ' If currentProcess = "SSGNP_CCWT" Or currentProcess = "TT" Then
            '     isSETUP = True
            ' End If
        End If
        
        If isSETUP Then
            
            pvtName = wsSource.Cells(i, 3).Value ' Column C - PVT Name
            
            
            If MatchesPatternList(pvtName, setupPatternArr) Then
                extractionTypes = setupExtractionArr
            ElseIf MatchesPatternList(pvtName, typicalPatternArr) Then
                extractionTypes = typicalExtractionArr
            ElseIf MatchesPatternList(pvtName, holdPatternArr) Then
                extractionTypes = holdExtractionArr
            Else
                extractionTypes = Array("") ' Default empty
            End If
            
            vddp = wsSource.Cells(i, 5).Value ' Column E - VDDP
            vdda = wsSource.Cells(i, 6).Value ' Column F - VDDA
            
            
            fastestRMValue = GetCellValue(wsSource, i, 9) ' Column I - Fastest RM as
            
            
            If pvtName = "" Then
                GoTo NextRow
            End If
            
            
            For j = LBound(extractionTypes) To UBound(extractionTypes)
                extractionType = Trim(extractionTypes(j))
                
                
                With wsOutput
                    .Cells(outputRow, 3).Value = pvtName ' PVT Name
                    .Cells(outputRow, 4).Value = extractionType ' Extraction Corner
                    .Cells(outputRow, 5).Value = vddp ' VDDP(V)
                    .Cells(outputRow, 6).Value = vdda ' VDDA (V)
                    .Cells(outputRow, 7).Value = fastestRMValue ' Fastest RM as
                    
                    
                    
                    dataCol = 8 ' Start after RM column
                    
                    If hasFocusFilter And focusColOrderOutput.Count > 0 Then
                        
                        For focusIdx = 1 To focusColOrderOutput.Count
                            colHeaderName = focusColOrderOutput(focusIdx)
                            sourceColIndex = focusColMapping(colHeaderName)
                            
                            
                            .Cells(outputRow, dataCol).Value = GetCellValue(wsSource, i, sourceColIndex)
                            dataCol = dataCol + 1
                        Next focusIdx
                    End If
                    
                    
                    .Cells(outputRow, dataCol).Value = "" ' Instance List
                    .Cells(outputRow, dataCol + 1).Value = "" ' PNR
                    .Cells(outputRow, dataCol + 2).Value = "" ' STA
                    .Cells(outputRow, dataCol + 3).Value = "" ' IR
                    
                    
                    .Range("A" & outputRow & ":" & ColLetter(dataCol + 3) & outputRow).Borders.LineStyle = xlContinuous
                    .Range("A" & outputRow & ":" & ColLetter(dataCol + 3) & outputRow).HorizontalAlignment = xlCenter
                End With
                
                counter = counter + 1
                outputRow = outputRow + 1
            Next j
        End If
        
NextRow:
    Next i
    
    
    setupEndRow = outputRow - 1
    
    
    If setupEndRow >= thisBandStartRow Then
        With wsOutput
            .Range("B" & thisBandStartRow & ":B" & setupEndRow).Merge
            .Range("B" & thisBandStartRow).Value = "SETUP"
            .Range("B" & thisBandStartRow).Font.Bold = True
            .Range("B" & thisBandStartRow).HorizontalAlignment = xlCenter
            .Range("B" & thisBandStartRow & ":B" & setupEndRow).Borders.LineStyle = xlContinuous
            
            ' Unmerge Instance List column for SETUP section so each PVT row has its own cell
            If hasFocusFilter And focusColOrder.Count > 0 Then
                setupInstanceListCol = 7 + focusColOrder.Count + 1
            Else
                setupInstanceListCol = 8
            End If
            .Range(.Cells(thisBandStartRow, setupInstanceListCol), .Cells(setupEndRow, setupInstanceListCol)).UnMerge
            .Range(.Cells(thisBandStartRow, setupInstanceListCol), .Cells(setupEndRow, setupInstanceListCol)).HorizontalAlignment = xlCenter
            .Range(.Cells(thisBandStartRow, setupInstanceListCol), .Cells(setupEndRow, setupInstanceListCol)).VerticalAlignment = xlCenter
            .Range(.Cells(thisBandStartRow, setupInstanceListCol), .Cells(setupEndRow, setupInstanceListCol)).Borders.LineStyle = xlContinuous
        End With
    End If
    
    
    
    holdStartRow = outputRow
    
    
    For i = bandStartRow To bandEndRow
        currentPVT = Trim(CStr(wsSource.Cells(i, 3).Value)) ' Column C - PVT Name
        currentProcess = Trim(CStr(wsSource.Cells(i, 4).Value)) ' Column D - Process
        
        
        
        isHOLD = False
        ' Check for hold override first
        holdOverrideEx2 = FindHoldOverrideExtractions(currentPVT, holdOverrideArr)
        If Not IsEmpty(holdOverrideEx2) Then
            isHOLD = True
        ElseIf currentPVT <> "" Then
            If MatchesPatternList(currentPVT, holdPatternArr) Or MatchesPatternList(currentPVT, typicalPatternArr) Then
                isHOLD = True
            End If
        End If
            
            
            ' If currentProcess = "FFGNP_CCWT" Or currentProcess = "TT" Then
            '     isHOLD = True
            ' End If
        
        If isHOLD Then
            
            pvtName = wsSource.Cells(i, 3).Value ' Column C - PVT Name
            
            
            If Not IsEmpty(holdOverrideEx2) Then
                holdExtractionTypes = holdOverrideEx2
            ElseIf MatchesPatternList(pvtName, holdPatternArr) Then
                holdExtractionTypes = holdExtractionArr
            ElseIf MatchesPatternList(pvtName, typicalPatternArr) Then
                holdExtractionTypes = typicalExtractionArr
            Else
                holdExtractionTypes = Array("")
            End If
            
            vddp = wsSource.Cells(i, 5).Value ' Column E - VDDP
            vdda = wsSource.Cells(i, 6).Value ' Column F - VDDA
            
            
            fastestRMValue = GetCellValue(wsSource, i, 9) ' Column I - Fastest RM as
            
            
            If pvtName = "" Then
                GoTo NextRowHold
            End If
            
            
            For kInt = LBound(holdExtractionTypes) To UBound(holdExtractionTypes)
                extractionType = Trim(holdExtractionTypes(kInt))
                
                
                With wsOutput
                    .Cells(outputRow, 3).Value = pvtName ' PVT Name
                    .Cells(outputRow, 4).Value = extractionType ' Extraction Corner
                    .Cells(outputRow, 5).Value = vddp ' VDDP(V)
                    .Cells(outputRow, 6).Value = vdda ' VDDA (V)
                    .Cells(outputRow, 7).Value = fastestRMValue ' Fastest RM as
                    
                    
                    
                    dataCol = 8 ' Start after RM column
                    
                    If hasFocusFilter And focusColOrderOutput.Count > 0 Then
                        
                        For focusIdx = 1 To focusColOrderOutput.Count
                            colHeaderName = focusColOrderOutput(focusIdx)
                            sourceColIndex = focusColMapping(colHeaderName)
                            
                            
                            .Cells(outputRow, dataCol).Value = GetCellValue(wsSource, i, sourceColIndex)
                            dataCol = dataCol + 1
                        Next focusIdx
                    End If
                    
                    
                    .Cells(outputRow, dataCol).Value = "" ' Instance List
                    .Cells(outputRow, dataCol + 1).Value = "" ' PNR
                    .Cells(outputRow, dataCol + 2).Value = "" ' STA
                    .Cells(outputRow, dataCol + 3).Value = "" ' IR
                    
                    
                    .Range("A" & outputRow & ":" & ColLetter(dataCol + 3) & outputRow).Borders.LineStyle = xlContinuous
                    .Range("A" & outputRow & ":" & ColLetter(dataCol + 3) & outputRow).HorizontalAlignment = xlCenter
                End With
                
                counter = counter + 1
                outputRow = outputRow + 1
            Next kInt
        End If
        
NextRowHold:
    Next i
    
    
    If outputRow > holdStartRow Then
        With wsOutput
            .Range("B" & holdStartRow & ":B" & (outputRow - 1)).Merge
            .Range("B" & holdStartRow).Value = "HOLD"
            .Range("B" & holdStartRow).Font.Bold = True
            .Range("B" & holdStartRow).HorizontalAlignment = xlCenter
            .Range("B" & holdStartRow & ":B" & (outputRow - 1)).Borders.LineStyle = xlContinuous
            
            ' Unmerge Instance List column for HOLD section so each PVT row has its own cell
            If hasFocusFilter And focusColOrder.Count > 0 Then
                holdInstanceListCol = 7 + focusColOrder.Count + 1
            Else
                holdInstanceListCol = 8
            End If
            .Range(.Cells(holdStartRow, holdInstanceListCol), .Cells(outputRow - 1, holdInstanceListCol)).UnMerge
            .Range(.Cells(holdStartRow, holdInstanceListCol), .Cells(outputRow - 1, holdInstanceListCol)).HorizontalAlignment = xlCenter
            .Range(.Cells(holdStartRow, holdInstanceListCol), .Cells(outputRow - 1, holdInstanceListCol)).VerticalAlignment = xlCenter
            .Range(.Cells(holdStartRow, holdInstanceListCol), .Cells(outputRow - 1, holdInstanceListCol)).Borders.LineStyle = xlContinuous
        End With
    End If
    
    
    If outputRow - 1 >= thisBandStartRow Then
        With wsOutput
            .Range("A" & thisBandStartRow & ":A" & (outputRow - 1)).Merge
            .Range("A" & thisBandStartRow).Value = bandName
            .Range("A" & thisBandStartRow).Font.Bold = True
            .Range("A" & thisBandStartRow).HorizontalAlignment = xlCenter
            .Range("A" & thisBandStartRow & ":A" & (outputRow - 1)).Borders.LineStyle = xlContinuous
        End With
    End If
    
    
    If outputRow > thisBandStartRow Then
        previousBandHadData = True
    Else
        previousBandHadData = False
    End If
    
    setupCount = setupEndRow - thisBandStartRow + 1
    If outputRow > holdStartRow Then
        holdCount = outputRow - holdStartRow
    Else
        holdCount = 0
    End If
    
    
    
    Next bandIndex
    
    
    With wsOutput
        
    
        If hasFocusFilter And focusColOrder.Count > 0 Then
            lastColIndex = 7 + focusColOrder.Count + 4 ' A-G + focus cols + Instance List/PNR/STA/IR
        Else
            lastColIndex = 11 ' A-G + Instance List/PNR/STA/IR (no focus columns)
        End If
        lastCol = ColLetter(lastColIndex)
        
    .Columns("A:" & lastCol).AutoFit
    .Range("A1:" & lastCol & (outputRow - 1)).VerticalAlignment = xlCenter
        
        
        .Columns("A").ColumnWidth = 24
        .Columns("B").ColumnWidth = 18
        .Columns("C").ColumnWidth = 45
        .Columns("D").ColumnWidth = 24
        .Columns("E").ColumnWidth = 12
        .Columns("F").ColumnWidth = 12
        .Columns("G").ColumnWidth = 15
        
        
    
        For col = 8 To lastColIndex
            If col <= lastColIndex - 4 Then ' Focus columns
                .Columns(ColLetter(col)).ColumnWidth = 10
            Else ' Instance List, PNR, STA, IR columns
                .Columns(ColLetter(col)).ColumnWidth = 15
            End If
        Next col
        
        
        If hasFocusFilter And focusColOrder.Count > 0 Then
            focusStartCol = "H" ' Column H is first focus column
            focusEndCol = ColLetter(7 + focusColOrder.Count) ' Last focus column based on actual count
            Set yesRange = .Range(focusStartCol & "2:" & focusEndCol & (outputRow - 1)) ' Start from row 2 (data rows)
            yesRange.FormatConditions.Delete ' Clear existing rules
            
            With yesRange.FormatConditions.Add(Type:=xlCellValue, Operator:=xlEqual, Formula1:="=""Yes""")
                .Interior.Color = RGB(198, 239, 206) ' Light green
                .Font.Color = RGB(0, 97, 0) ' Dark green text
                .Font.Bold = True
            End With
            
            
            With yesRange.FormatConditions.Add(Type:=xlCellValue, Operator:=xlEqual, Formula1:="=""No""")
                .Interior.Color = RGB(255, 199, 206) ' Light red
                .Font.Color = RGB(156, 0, 6) ' Dark red text
            End With
        End If

        
        rshade1 = RGB(255, 255, 255)
        rshade2 = RGB(245, 245, 245)
        For r2 = 2 To outputRow - 1 ' Start from row 2 (data rows)
            If r2 Mod 2 = 0 Then
                .Range("C" & r2 & ":" & lastCol & r2).Interior.Color = rshade1
            Else
                .Range("C" & r2 & ":" & lastCol & r2).Interior.Color = rshade2
            End If
        Next r2

        
    .Range("A1:" & lastCol & (outputRow - 1)).Font.Name = "Calibri"
        
        
        colorIndex = 0
        
        For bandRow = 2 To outputRow - 1 ' Start from row 2 (data rows)
            If .Cells(bandRow, 1).Value <> "" Then
                currentBandName = .Cells(bandRow, 1).Value
                
                Select Case colorIndex Mod 3
                    Case 0
                        bandColor = RGB(217, 225, 242) ' Light blue
                    Case 1
                        bandColor = RGB(234, 209, 220) ' Light purple
                    Case 2
                        bandColor = RGB(226, 239, 218) ' Light green
                End Select
                
                
                For Each bandCell In .Range("A" & bandRow & ":A" & bandRow).Cells
                    If Not bandCell.MergeCells Then
                        bandCell.Interior.Color = bandColor
                    Else
                        bandCell.MergeArea.Interior.Color = bandColor
                        Exit For
                    End If
                Next bandCell
                
                colorIndex = colorIndex + 1
            End If
        Next bandRow
        
        
        For bandRow = 2 To outputRow - 1 ' Start from row 2 (data rows)
            If .Cells(bandRow, 2).Value <> "" Then
                Set sectionCell = .Cells(bandRow, 2)
                
                If sectionCell.MergeCells Then
                    If UCase(Trim(sectionCell.Value)) = "SETUP" Then
                        sectionCell.MergeArea.Interior.Color = RGB(252, 228, 214) ' Light orange
                        sectionCell.MergeArea.Font.Color = RGB(192, 80, 77) ' Dark orange text
                    ElseIf UCase(Trim(sectionCell.Value)) = "HOLD" Then
                        sectionCell.MergeArea.Interior.Color = RGB(217, 234, 211) ' Light green
                        sectionCell.MergeArea.Font.Color = RGB(56, 118, 29) ' Dark green text
                    End If
                End If
            End If
        Next bandRow
        
        
    .Range("A1:" & lastCol & (outputRow - 1)).Borders.LineStyle = xlContinuous
    .Range("A1:" & lastCol & (outputRow - 1)).Borders.Weight = xlThin
    .Range("A1:" & lastCol & (outputRow - 1)).Borders.Color = RGB(191, 191, 191) ' Gray borders
        
        ' Header formatting already done earlier, just ensure borders are proper
        .Range("A1:" & lastCol & "2").Borders.Weight = xlMedium
    End With
    
    
    ' Apply AutoFilter starting from header row (row 1)
wsOutput.Range("A1:" & lastCol & "1").AutoFilter ' Ensure filter dropdowns are on header row only (row 1)

' Activate worksheet and freeze panes properly
wsOutput.Activate
Application.Goto wsOutput.Range("A1"), True ' Ensure worksheet is fully activated
wsOutput.Range("A3").Select ' Select A3 to freeze merged header rows (1-2)
On Error Resume Next
ActiveWindow.FreezePanes = False ' Clear any existing freeze
ActiveWindow.FreezePanes = True  ' Freeze at A3 (freezes merged header rows 1-2)
If Err.Number <> 0 Then
    ' Freeze panes failed - continue without error
    Err.Clear
End If
On Error GoTo ErrorHandler
wsOutput.Range("A1").Select ' Return to top
    
    
    
    
    
    CreatePVTSTASheet
    
    ' Populate Instance List column if focus columns include memory types or DR columns
    Debug.Print "=== Calling PopulateInstanceListColumn ==="
    Call PopulateInstanceListColumn(wsOutput, focusColOrder, focusColMapping)
    Debug.Print "=== PopulateInstanceListColumn call completed ==="
    
    ' Show concise success message with simple Valid/Invalid/Not used statuses
    
    If patternsValid Then
        patternsStatus = "Valid"
    Else
        patternsStatus = "Invalid - see warnings"
    End If

    If hasFocusFilter Then
        If focusColOrder.Count > 0 Then
            focusStatus = "Valid"
        Else
            focusStatus = "Invalid"
        End If
    Else
        focusStatus = "Not used"
    End If

    If holdOverrideTotal > 0 Then
        If holdOverrideMatchesCount > 0 Then
            holdMappingStatus = "Used (" & holdOverrideMatchesCount & " match(es))"
        Else
            holdMappingStatus = "Used (no matches)"
        End If
    Else
        holdMappingStatus = "Not used"
    End If

    successMsg = SHEET_OUTPUT & " sheet generated." & vbCrLf & vbCrLf
    successMsg = successMsg & "Configuration Check:" & vbCrLf
    successMsg = successMsg & "  - PVT patterns: " & patternsStatus & vbCrLf
    successMsg = successMsg & "  - Focus columns: " & focusStatus & vbCrLf
    successMsg = successMsg & "  - Hold-only mappings: " & holdMappingStatus & vbCrLf & vbCrLf
    successMsg = successMsg & "Total rows generated: " & (outputRow - 2) & vbCrLf & vbCrLf
    successMsg = successMsg & "Refer to the '" & SHEET_OUTPUT & "' sheet for results."

    ' Memory-type cross-check: Instance List -> PVTs (report matched / unmatched instance memory types)
    Dim instMemTypesDict As Object
    Set instMemTypesDict = CreateObject("Scripting.Dictionary")
    On Error Resume Next
    Dim wsInstances As Worksheet
    Set wsInstances = ThisWorkbook.Sheets(SHEET_INSTANCES)
    On Error GoTo 0
    If Not wsInstances Is Nothing Then
        Dim instLastRow As Long, instMemTypeCol As Long, r As Long
        instLastRow = wsInstances.Cells(wsInstances.Rows.Count, 2).End(xlUp).Row
        instMemTypeCol = FindColumnInRow(wsInstances, "memory_type", 2)
        If instMemTypeCol <> -1 Then
            For r = 3 To instLastRow
                Dim mt As String
                mt = Trim(UCase(CStr(wsInstances.Cells(r, instMemTypeCol).Value)))
                If mt <> "" Then
                    If Not instMemTypesDict.Exists(mt) Then instMemTypesDict.Add mt, True
                End If
            Next r
        End If
    End If

    Dim pvtMemTypesDict As Object
    Set pvtMemTypesDict = CreateObject("Scripting.Dictionary")
    Dim condIdx As Long
    For condIdx = 1 To memoryConditions.Count
        Dim condMT As String
        condMT = Trim(UCase(CStr(memoryConditions(condIdx)("memory_type"))))
        If condMT <> "" Then
            If Not pvtMemTypesDict.Exists(condMT) Then pvtMemTypesDict.Add condMT, True
        End If
    Next condIdx

    If instMemTypesDict.Count > 0 Then
        Dim matchedList As String, unmatchedList As String
        matchedList = "": unmatchedList = ""
        Dim k As Variant, foundMatch As Boolean, p As Variant
        For Each k In instMemTypesDict.Keys
            foundMatch = False
            For Each p In pvtMemTypesDict.Keys
                ' flexible matching: substring match either way
                If InStr(1, k, p, vbTextCompare) > 0 Or InStr(1, p, k, vbTextCompare) > 0 Then
                    foundMatch = True
                    Exit For
                End If
            Next p
            If foundMatch Then
                If matchedList = "" Then matchedList = k Else matchedList = matchedList & ", " & k
            Else
                If unmatchedList = "" Then unmatchedList = k Else unmatchedList = unmatchedList & ", " & k
            End If
        Next k

        successMsg = successMsg & vbCrLf & vbCrLf & "Memory-type cross-check (Instance List → PVTs):" & vbCrLf
        If matchedList <> "" Then
            successMsg = successMsg & "  • Matched: " & matchedList & vbCrLf
        Else
            successMsg = successMsg & "  • Matched: (none)" & vbCrLf
        End If
        If unmatchedList <> "" Then
            successMsg = successMsg & "  • Unmatched (in Instance List but not found in PVTs): " & unmatchedList & vbCrLf
        Else
            successMsg = successMsg & "  • Unmatched: (none)"
        End If
    End If

    ' Mark completion and cache the message so wrapper can show it after UI restored
    gPVTProcessing_Completed = True
    gPVTProcessing_SuccessMsg = successMsg
    gPVTProcessing_MsgShown = False

    MsgBox successMsg, vbInformation, "Processing Complete"
    gPVTProcessing_MsgShown = True

    Exit Sub

ErrorHandler:
    errMsg = "Error " & Err.Number & ": " & Err.Description & vbCrLf & vbCrLf & "At approximately row: " & outputRow
    MsgBox errMsg, vbCritical, "Macro error"
    On Error Resume Next
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Exit Sub
End Sub

'==========================================
' INSTANCE MATCHER - Match instances for all bands with visual colors
'==========================================
Public Sub GenerateInstanceMatchingReportEnhanced()
    On Error GoTo ErrorHandler
    
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    
    Dim wsPVT As Worksheet, wsInstances As Worksheet, wsOutput As Worksheet
    
    On Error Resume Next
    Set wsPVT = ThisWorkbook.Sheets(SHEET_PVT)
    Set wsInstances = ThisWorkbook.Sheets(SHEET_INSTANCES)
    On Error GoTo ErrorHandler
    
    If wsPVT Is Nothing Or wsInstances Is Nothing Then
        MsgBox "Required sheets not found: '" & SHEET_PVT & "' or '" & SHEET_INSTANCES & "'", vbCritical
        GoTo Cleanup
    End If
    
    On Error Resume Next
    Application.DisplayAlerts = False
    ThisWorkbook.Sheets(SHEET_MATCH_REPORT).Delete
    Application.DisplayAlerts = True
    On Error GoTo ErrorHandler
    
    Set wsOutput = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
    wsOutput.Name = SHEET_MATCH_REPORT
    
    ' Get all unique bands from PVTs sheet
    Dim bandDict As Object, bands As Collection
    Set bandDict = CreateObject("Scripting.Dictionary")
    Set bands = New Collection
    
    Dim pvtLastRow As Long, pvtRow As Long, bandName As String
    pvtLastRow = wsPVT.Cells(wsPVT.Rows.Count, 2).End(xlUp).Row
    
    For pvtRow = 3 To pvtLastRow
        bandName = Trim(CStr(wsPVT.Cells(pvtRow, 2).Value))
        If bandName <> "" And Not bandDict.Exists(bandName) Then
            bandDict.Add bandName, True
            bands.Add bandName
        End If
    Next pvtRow
    
    ' Find DR columns in PVTs sheet
    Dim dr0Col As Long, dr1Col As Long, ldrCol As Long, lastCol As Long
    dr0Col = 0: dr1Col = 0: ldrCol = 0
    lastCol = wsPVT.Cells(2, wsPVT.Columns.Count).End(xlToLeft).Column
    
    Dim col As Long
    For col = 1 To lastCol
        Dim h As String
        h = UCase(Trim(CStr(wsPVT.Cells(2, col).Value)))
        If h = "DR0" Then dr0Col = col
        If h = "DR1" Then dr1Col = col
        If h = "LDR" Then ldrCol = col
    Next col
    
    ' Parse memory conditions from Row 1 of PVTs
    Dim memoryConditions As Collection
    Set memoryConditions = New Collection
    
    For col = 10 To lastCol
        Dim headerText As String, memType As String, waVals() As String
        headerText = Trim(CStr(wsPVT.Cells(1, col).Value))
        If headerText <> "" And InStr(headerText, "(") > 0 Then
            If ParseMemoryConditionIM(headerText, memType, waVals) Then
                Dim condDict As Object
                Set condDict = CreateObject("Scripting.Dictionary")
                condDict("column") = col
                condDict("memory_type") = memType
                condDict("wa_values") = waVals
                memoryConditions.Add condDict
            End If
        End If
    Next col
    
    ' Find Instance List columns
    Dim instMemTypeCol As Long, instNameCol As Long, instWACol As Long, instDRCol As Long
    Dim instPartNameCol As Long
    Dim instCol As Long, instLastCol As Long
    instMemTypeCol = 0: instNameCol = 0: instWACol = 0: instDRCol = 0: instPartNameCol = 0
    
    instLastCol = wsInstances.Cells(2, wsInstances.Columns.Count).End(xlToLeft).Column
    For instCol = 1 To instLastCol
        Dim hdr As String
        hdr = LCase(Trim(CStr(wsInstances.Cells(2, instCol).Value)))
        If hdr = "memory_type" Then instMemTypeCol = instCol
        If hdr = "instance_name" Then instNameCol = instCol
        If hdr = "write_assist" Then instWACol = instCol
        If hdr = "dual_rail" Then instDRCol = instCol
        If hdr = "part_name" Then instPartNameCol = instCol
    Next instCol
    
    If instMemTypeCol = 0 Or instDRCol = 0 Or instWACol = 0 Then
        MsgBox "Required columns not found in " & SHEET_INSTANCES & " sheet", vbCritical
        GoTo Cleanup
    End If
    
    ' Verify memory types in PVTs exist in Instance List - warn if mismatches
    ' Determine instance last row once and reuse
    Dim instDataLastRow As Long
    instDataLastRow = wsInstances.Cells(wsInstances.Rows.Count, instMemTypeCol).End(xlUp).Row
    Dim instMemTypesDict As Object
    Set instMemTypesDict = CreateObject("Scripting.Dictionary")
    Dim instRow2 As Long
    Dim mt As String
    For instRow2 = 3 To instDataLastRow
        mt = Trim(UCase(CStr(wsInstances.Cells(instRow2, instMemTypeCol).Value)))
        If mt <> "" Then
            If Not instMemTypesDict.Exists(mt) Then instMemTypesDict.Add mt, True
        End If
    Next instRow2

    Dim missingList As String
    missingList = ""
    Dim condIdx2 As Long
    For condIdx2 = 1 To memoryConditions.Count
        Dim condMT As String
        condMT = Trim(UCase(CStr(memoryConditions(condIdx2)("memory_type"))))
        If condMT <> "" Then
            If Not instMemTypesDict.Exists(condMT) Then
                If missingList <> "" Then missingList = missingList & ", "
                missingList = missingList & condMT
            End If
        End If
    Next condIdx2

    If missingList <> "" Then
        If MsgBox("WARNING: The following memory types are present in '" & SHEET_PVT & "' but not in '" & SHEET_INSTANCES & "': " & missingList & "." & vbCrLf & vbCrLf & _
                  "This may cause no matches for affected bands. Click OK to continue or Cancel to stop.", vbExclamation + vbOKCancel, "Memory type mismatch") = vbCancel Then
            GoTo Cleanup
        End If
    End If

    ' Create headers
    Dim outputRow As Long
    outputRow = 1
    With wsOutput
        .Cells(1, 1).Value = "Band"
        .Cells(1, 2).Value = "instance_name"
        .Cells(1, 3).Value = "Condition Status"
        .Cells(1, 4).Value = "memory_type"
        .Cells(1, 5).Value = "write_assist"
        .Cells(1, 6).Value = "dr0"
        .Cells(1, 7).Value = "dr1"
        .Range("A1:G1").Font.Bold = True
        .Range("A1:G1").Font.Size = 12
        .Range("A1:G1").Interior.Color = RGB(68, 114, 196)
        .Range("A1:G1").Font.Color = RGB(255, 255, 255)
        .Range("A1:G1").HorizontalAlignment = xlCenter
        .Range("A1:G1").VerticalAlignment = xlCenter
        .Rows(1).RowHeight = 30
    End With
    
    outputRow = 2
    
    ' Define color palette for bands (matching PVT_STA sheet dynamic assignment)
    ' Colors will be assigned based on band order, not band names
    Dim bandColorIndex As Integer
    bandColorIndex = 0
    
    Dim totalMatchCount As Long
    totalMatchCount = 0

    ' Track counts per band for summary
    Dim bandMatchCounts As Object
    Set bandMatchCounts = CreateObject("Scripting.Dictionary")
    
    ' Process each band
    Dim bandIdx As Long
    For bandIdx = 1 To bands.Count
        bandName = bands(bandIdx)
        
        ' Find first row for this band in PVTs
        bandRow = 0
        For pvtRow = 3 To pvtLastRow
            If Trim(CStr(wsPVT.Cells(pvtRow, 2).Value)) = bandName Then
                bandRow = pvtRow
                Exit For
            End If
        Next pvtRow
        
        If bandRow = 0 Then GoTo NextBand
        
        ' All bands now use dynamic logic from PVTs sheet
        ' (Removed special hardcoded handling for 0.55V-0.75V)
        Dim useHardcodedLogic As Boolean
        useHardcodedLogic = False
        
        If useHardcodedLogic Then
            ' Use hardcoded logic for this specific band
            Dim band055075StartRow As Long
            band055075StartRow = outputRow
            
            ' Hardcoded memory type conditions
            Dim memCond055 As Object
            Set memCond055 = CreateObject("Scripting.Dictionary")
            memCond055.Add "HDDP_1", True
            memCond055.Add "EHD1PRF_1", True
            memCond055.Add "EHD2PRF_0", True
            memCond055.Add "EHDP2PRF_0", True
            memCond055.Add "UHS1PRF_0", True
            memCond055.Add "HS2PRF_0", True
            memCond055.Add "HS2PRF_1", True
            memCond055.Add "HDSP_0", True
            memCond055.Add "HDSP_1", True
            memCond055.Add "HDSPSRAM_0", True
            memCond055.Add "HDSPSRAM_1", True
            memCond055.Add "HD1PRF_0", True
            memCond055.Add "HD1PRF_1", True
            memCond055.Add "HD1PRF_cr_1", False
            memCond055.Add "HD1PRF_gy_0", True
            
            Dim inst055Row As Long
            Dim skippedDR As Long, skippedWA As Long, skippedMemType As Long, matchedCount As Long
            skippedDR = 0: skippedWA = 0: skippedMemType = 0: matchedCount = 0
            
            For inst055Row = 3 To instDataLastRow
                Dim inst055MemType As String, inst055Name As String, inst055PartName As String
                Dim inst055WA As String, inst055DR As String, inst055WAInt As Integer
                
                inst055MemType = Trim(CStr(wsInstances.Cells(inst055Row, instMemTypeCol).Value))
                If inst055MemType = "" Then GoTo Next055Instance
                
                If instNameCol > 0 Then
                    inst055Name = Trim(CStr(wsInstances.Cells(inst055Row, instNameCol).Value))
                Else
                    inst055Name = ""
                End If
                
                If instPartNameCol > 0 Then
                    inst055PartName = Trim(CStr(wsInstances.Cells(inst055Row, instPartNameCol).Value))
                Else
                    inst055PartName = ""
                End If
                
                If inst055Name = "" Or Left(inst055Name, 1) = "=" Then
                    If inst055PartName <> "" Then
                        inst055Name = inst055PartName
                    Else
                        inst055Name = inst055MemType & "_Row" & inst055Row
                    End If
                End If
                
                inst055WA = Trim(CStr(wsInstances.Cells(inst055Row, instWACol).Value))
                inst055DR = LCase(Trim(CStr(wsInstances.Cells(inst055Row, instDRCol).Value)))
                
                ' DR must be dr1 (DR0=No, DR1=Yes)
                If inst055DR <> "dr1" Then
                    skippedDR = skippedDR + 1
                    GoTo Next055Instance
                End If
                
                ' Parse WA
                On Error Resume Next
                inst055WAInt = CInt(inst055WA)
                If Err.Number <> 0 Then
                    Err.Clear
                    skippedWA = skippedWA + 1
                    GoTo Next055Instance
                End If
                On Error GoTo ErrorHandler
                
                ' Check memory type and WA
                Dim condKey055 As String
                condKey055 = inst055MemType & "_" & CStr(inst055WAInt)
                
                Dim shouldMatch055 As Boolean
                shouldMatch055 = False
                
                If memCond055.Exists(condKey055) Then
                    shouldMatch055 = memCond055(condKey055)
                End If
                
                If Not shouldMatch055 Then
                    skippedMemType = skippedMemType + 1
                    GoTo Next055Instance
                End If
                
                If shouldMatch055 Then
                    matchedCount = matchedCount + 1
                    Select Case bandColorIndex Mod 3
                        Case 0
                            bandColor055 = RGB(217, 225, 242) ' Light blue
                        Case 1
                            bandColor055 = RGB(234, 209, 220) ' Light purple
                        Case 2
                            bandColor055 = RGB(226, 239, 218) ' Light green
                    End Select
                    
                    With wsOutput
                        .Cells(outputRow, 1).Value = bandName
                        .Cells(outputRow, 1).Interior.Color = bandColor055
                        
                        .Cells(outputRow, 2).Value = inst055Name
                        .Cells(outputRow, 2).Interior.Color = bandColor055
                        
                        .Cells(outputRow, 3).Value = " Match"
                        .Cells(outputRow, 3).Interior.Color = RGB(198, 239, 206)
                        .Cells(outputRow, 3).Font.Color = RGB(0, 97, 0)
                        .Cells(outputRow, 3).Font.Bold = True
                        
                        .Cells(outputRow, 4).Value = inst055MemType
                        .Cells(outputRow, 4).Interior.Color = bandColor055
                        
                        .Cells(outputRow, 5).Value = inst055WAInt
                        .Cells(outputRow, 5).Interior.Color = bandColor055
                        
                        .Cells(outputRow, 6).Value = "No"
                        .Cells(outputRow, 6).Interior.Color = bandColor055
                        
                        .Cells(outputRow, 7).Value = "Yes"
                        .Cells(outputRow, 7).Interior.Color = bandColor055
                    End With
                    
                    outputRow = outputRow + 1
                    totalMatchCount = totalMatchCount + 1
                End If
                
Next055Instance:
            Next inst055Row
            
            ' Debug message for 0.55V-0.75V band
            Debug.Print "0.55V-0.75V Band: Matched=" & matchedCount & ", SkippedDR=" & skippedDR & ", SkippedWA=" & skippedWA & ", SkippedMemType=" & skippedMemType
            
            ' Merge band column and append count onto the merged cell
            If outputRow > band055075StartRow Then
                Application.DisplayAlerts = False
                With wsOutput
                    .Range("A" & band055075StartRow & ":A" & (outputRow - 1)).Merge
                    .Cells(band055075StartRow, 1).Value = bandName & vbCrLf & "No.of Instances: " & matchedCount
                    .Range("A" & band055075StartRow & ":A" & (outputRow - 1)).VerticalAlignment = xlTop
                    .Range("A" & band055075StartRow & ":A" & (outputRow - 1)).HorizontalAlignment = xlCenter
                    .Range("A" & band055075StartRow & ":A" & (outputRow - 1)).Font.Bold = True
                    .Range("A" & band055075StartRow & ":A" & (outputRow - 1)).Font.Size = 11
                    .Range("A" & band055075StartRow & ":A" & (outputRow - 1)).WrapText = True
                    .Range("A" & band055075StartRow & ":A" & (outputRow - 1)).Interior.Color = bandColor055
                End With
                ' Record the count for summary
                If Not bandMatchCounts.Exists(bandName) Then bandMatchCounts.Add bandName, matchedCount Else bandMatchCounts(bandName) = matchedCount
                Application.DisplayAlerts = True
            End If
            
            ' Increment color index for next band (matching PVT_STA logic)
            bandColorIndex = bandColorIndex + 1
            
            GoTo NextBand
        End If
        
        ' For other bands, use dynamic logic from PVTs sheet
        ' Get DR conditions for this band
        Dim allowedDRs As Object
        Set allowedDRs = CreateObject("Scripting.Dictionary")
        
        If dr0Col > 0 Then
            Dim dr0Val As String
            dr0Val = UCase(Trim(CStr(wsPVT.Cells(bandRow, dr0Col).Value)))
            If dr0Val = "YES" Or dr0Val = "Y" Then allowedDRs.Add "dr0", True
        End If
        
        If dr1Col > 0 Then
            Dim dr1Val As String
            dr1Val = UCase(Trim(CStr(wsPVT.Cells(bandRow, dr1Col).Value)))
            If dr1Val = "YES" Or dr1Val = "Y" Then allowedDRs.Add "dr1", True
        End If
        
        If ldrCol > 0 Then
            Dim ldrVal As String
            ldrVal = UCase(Trim(CStr(wsPVT.Cells(bandRow, ldrCol).Value)))
            If ldrVal = "YES" Or ldrVal = "Y" Then allowedDRs.Add "ldr", True
        End If
        
        ' Get valid memory types for this band
        Dim validMemTypes As Collection
        Set validMemTypes = New Collection
        
        Dim condIdx As Long
        For condIdx = 1 To memoryConditions.Count
            Dim cond As Object
            Set cond = memoryConditions(condIdx)
            Dim cellVal As String
            cellVal = UCase(Trim(CStr(wsPVT.Cells(bandRow, cond("column")).Value)))
            
            If cellVal = "YES" Or cellVal = "Y" Then
                validMemTypes.Add cond
            End If
        Next condIdx
        
        If validMemTypes.Count = 0 Or allowedDRs.Count = 0 Then GoTo NextBand
        
        ' Track start row for this band
        Dim bandStartRow As Long
        bandStartRow = outputRow
        
        ' Track count for this band
        Dim bandInstanceCount As Long
        bandInstanceCount = 0
        
        ' Match instances for this band
        Dim instRow As Long
        For instRow = 3 To instDataLastRow
            Dim instMemType As String, instName As String, instPartName As String
            Dim instWA As String, instDR As String
            
            instMemType = Trim(CStr(wsInstances.Cells(instRow, instMemTypeCol).Value))
            If instMemType = "" Then GoTo NextInstance
            
            If instNameCol > 0 Then
                instName = Trim(CStr(wsInstances.Cells(instRow, instNameCol).Value))
            Else
                instName = ""
            End If
            
            If instPartNameCol > 0 Then
                instPartName = Trim(CStr(wsInstances.Cells(instRow, instPartNameCol).Value))
            Else
                instPartName = ""
            End If
            
            ' Use part_name as fallback
            If instName = "" Or Left(instName, 1) = "=" Then
                If instPartName <> "" Then
                    instName = instPartName
                Else
                    instName = instMemType & "_Row" & instRow
                End If
            End If
            
            instWA = Trim(CStr(wsInstances.Cells(instRow, instWACol).Value))
            instDR = LCase(Trim(CStr(wsInstances.Cells(instRow, instDRCol).Value)))
            
            ' Check if DR is allowed
            If Not allowedDRs.Exists(instDR) Then GoTo NextInstance
            
            ' Check memory type match
            Dim vIdx As Long, vCond As Object
            Dim vMemType As String, vWAVals() As String
            
            For vIdx = 1 To validMemTypes.Count
                Set vCond = validMemTypes(vIdx)
                vMemType = vCond("memory_type")
                vWAVals = vCond("wa_values")
                
                If InstanceMatchesIM(instMemType, instWA, vMemType, vWAVals) Then
                    ' Write match
                    Dim dr0Display As String, dr1Display As String
                    If instDR = "dr0" Then
                        dr0Display = "Yes"
                        dr1Display = "No"
                    ElseIf instDR = "dr1" Then
                        dr0Display = "No"
                        dr1Display = "Yes"
                    Else
                        dr0Display = "N/A"
                        dr1Display = "N/A"
                    End If
                    
                    ' Get band color (dynamic assignment matching PVT_STA)
                    Select Case bandColorIndex Mod 3
                        Case 0
                            bandColor = RGB(217, 225, 242) ' Light blue
                        Case 1
                            bandColor = RGB(234, 209, 220) ' Light purple
                        Case 2
                            bandColor = RGB(226, 239, 218) ' Light green
                    End Select
                    
                    With wsOutput
                        .Cells(outputRow, 1).Value = bandName
                        .Cells(outputRow, 1).Interior.Color = bandColor
                        
                        .Cells(outputRow, 2).Value = instName
                        .Cells(outputRow, 2).Interior.Color = bandColor
                        
                        .Cells(outputRow, 3).Value = "Match"
                        .Cells(outputRow, 3).Interior.Color = RGB(198, 239, 206)
                        .Cells(outputRow, 3).Font.Color = RGB(0, 97, 0)
                        .Cells(outputRow, 3).Font.Bold = True
                        
                        .Cells(outputRow, 4).Value = instMemType
                        .Cells(outputRow, 4).Interior.Color = bandColor
                        
                        .Cells(outputRow, 5).Value = instWA
                        .Cells(outputRow, 5).Interior.Color = bandColor
                        
                        .Cells(outputRow, 6).Value = dr0Display
                        .Cells(outputRow, 6).Interior.Color = bandColor
                        
                        .Cells(outputRow, 7).Value = dr1Display
                        .Cells(outputRow, 7).Interior.Color = bandColor
                    End With
                    
                    outputRow = outputRow + 1
                    totalMatchCount = totalMatchCount + 1
                    bandInstanceCount = bandInstanceCount + 1
                    Exit For
                End If
            Next vIdx
            
NextInstance:
        Next instRow
        
        ' Merge band column for this band's instances and append count to the merged header
        If outputRow > bandStartRow Then
            Application.DisplayAlerts = False
            With wsOutput
                .Range("A" & bandStartRow & ":A" & (outputRow - 1)).Merge
                .Range("A" & bandStartRow & ":A" & (outputRow - 1)).Value = bandName & vbCrLf & "No.of Instances: " & bandInstanceCount
                .Range("A" & bandStartRow & ":A" & (outputRow - 1)).VerticalAlignment = xlTop
                .Range("A" & bandStartRow & ":A" & (outputRow - 1)).HorizontalAlignment = xlCenter
                .Range("A" & bandStartRow & ":A" & (outputRow - 1)).Font.Bold = True
                .Range("A" & bandStartRow & ":A" & (outputRow - 1)).Font.Size = 11
                .Range("A" & bandStartRow & ":A" & (outputRow - 1)).WrapText = True
                .Range("A" & bandStartRow & ":A" & (outputRow - 1)).Interior.Color = bandColor
            End With
            ' Record the count for summary
            If Not bandMatchCounts.Exists(bandName) Then bandMatchCounts.Add bandName, bandInstanceCount Else bandMatchCounts(bandName) = bandInstanceCount
            Application.DisplayAlerts = True
        End If
        
        ' Increment color index for next band (matching PVT_STA logic)
        bandColorIndex = bandColorIndex + 1
        
NextBand:
    Next bandIdx
    
    ' Format output
    With wsOutput
        .Columns("A").ColumnWidth = 18
        ' Dynamically size Instance Name column: autosize then cap to a sensible range
        .Columns("B").AutoFit
        Dim instColWidth As Double
        instColWidth = .Columns("B").ColumnWidth
        If instColWidth < 30 Then
            .Columns("B").ColumnWidth = 30
        ElseIf instColWidth > 80 Then
            .Columns("B").ColumnWidth = 80
        End If
        .Columns("C").ColumnWidth = 20
        .Columns("D").ColumnWidth = 18
        .Columns("E").ColumnWidth = 15
        .Columns("F").ColumnWidth = 12
        .Columns("G").ColumnWidth = 10
        
        .Range("A1:G1").Borders.Weight = xlMedium
        .Range("A2:G" & (outputRow - 1)).Borders.LineStyle = xlContinuous
        .Range("A2:G" & (outputRow - 1)).Borders.Color = RGB(191, 191, 191)
        
        .Range("B2:B" & (outputRow - 1)).HorizontalAlignment = xlCenter
        .Range("C2:G" & (outputRow - 1)).HorizontalAlignment = xlCenter
        
        ' Add alternating row colors (matching PVT_STA sheet)
        rshade1 = RGB(255, 255, 255)  ' White
        rshade2 = RGB(245, 245, 245)  ' Light gray
        For r2 = 2 To outputRow - 1    ' Start from row 2 (data rows)
            If r2 Mod 2 = 0 Then
                .Range("B" & r2 & ":G" & r2).Interior.Color = rshade1
            Else
                .Range("B" & r2 & ":G" & r2).Interior.Color = rshade2
            End If
        Next r2
        
        ' Add autofilter
        .Range("A1:G1").AutoFilter
        
        ' Freeze panes
        .Range("A2").Select
        On Error Resume Next
        ActiveWindow.FreezePanes = False
        ActiveWindow.FreezePanes = True
        If Err.Number <> 0 Then Err.Clear
        On Error GoTo ErrorHandler
        
        .Range("A1").Select
    End With
    
    ' Apply consistent font formatting (matching PVT_STA sheet)
    With wsOutput
        .Range("A1:H" & (outputRow - 1)).Font.Name = "Calibri"
        .Range("A1:H" & (outputRow - 1)).VerticalAlignment = xlCenter
    End With
    
    wsOutput.Activate
    
    ' Create detailed summary by band
    Dim summaryMsg As String
    summaryMsg = "Instance Matching Process Completed Successfully" & vbCrLf & vbCrLf & _
                 "Summary:" & vbCrLf & _
                 "• Total instances matched: " & totalMatchCount & vbCrLf & vbCrLf & _
                 "Breakdown by Band:" & vbCrLf
    Dim k As Variant
    For Each k In bandMatchCounts.Keys
        summaryMsg = summaryMsg & "  • " & k & ": " & bandMatchCounts(k) & " instances" & vbCrLf
    Next k
    summaryMsg = summaryMsg & vbCrLf & "Please review the '" & SHEET_MATCH_REPORT & "' sheet for detailed results."
    
    MsgBox summaryMsg, vbInformation, "Instance Matching Complete"
    
Cleanup:
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Exit Sub

ErrorHandler:
    MsgBox "Error " & Err.Number & ": " & Err.Description, vbCritical
    GoTo Cleanup
End Sub

'==========================================
' POPULATE INSTANCE LIST COLUMN IN PVT_STA
'==========================================
Private Sub PopulateInstanceListColumn(wsPVT As Worksheet, focusColOrder As Collection, focusColMapping As Object)
    On Error GoTo ErrorHandler
    
    ' Declare all variables at the top
    Dim wsInstances As Worksheet
    Dim instanceListCol As Long
    Dim instanceData As Collection
    Dim dr0Col As Long, dr1Col As Long
    Dim memTypeColumns As Collection
    Dim vtCols As Collection
    Dim vtInfo As Object
    Dim hasDRFocus As Boolean
    Dim focusIdx As Long
    Dim focusName As String
    Dim lastRow As Long, pvtRow As Long
    Dim dr0 As String, dr1 As String
    Dim allInstances As Collection
    Dim i As Long
    Dim colInfo As Object
    Dim cellValue As String
    Dim memType As String
    Dim waValues() As String
    Dim waIdx As Long
    Dim matches As Collection
    Dim j As Long
    Dim instName As String
    Dim alreadyExists As Boolean
    Dim k As Long
    Dim instIdx As Long
    Dim instInfo As Object
    Dim instDR As String
    Dim drMatch As Boolean
    Dim instanceListStr As String
    Dim bandDRConditions As Object
    Dim bandName As String
    Dim bandCond As Object
    Dim checkRow As Long
    Dim checkDR0 As String, checkDR1 As String
    Dim focusedMemTypes As Collection
    Dim dr0Yes As Boolean, dr1Yes As Boolean
    Dim memMatch As Boolean
    Dim fm As Variant
    Dim rowMemMatch As Boolean
    Dim foundCol As Long
    Dim memTypePattern As String
    Dim vtType As String
    Dim vtMatched As Boolean
    Dim vtHasYes As Boolean
    Dim vtMismatch As Boolean
    Dim lastHeaderCol As Long
    Dim headerText As String
    ' Debugging: per-instance skip reasons
    Dim skipReasons As Object
    Dim reasonCounts As Object
    Dim reasonExamples As Object
    Dim reasonKey As String
    Dim exampleLimit As Integer
    Dim exStr As String
    Dim exCount As Integer
    
    Set wsInstances = ThisWorkbook.Sheets(SHEET_INSTANCES)
    
    If wsInstances Is Nothing Then
        Debug.Print "PopulateInstanceListColumn: N3P Instance List sheet not found"
        Exit Sub
    End If
    
    ' Find Instance List column
    instanceListCol = FindColumnInRow(wsPVT, "instance list", 1)
    If instanceListCol = -1 Then instanceListCol = FindColumnInRow(wsPVT, "instance_list", 1)
    If instanceListCol = -1 Then
        Debug.Print "PopulateInstanceListColumn: Instance List column not found in PVT_STA"
        Exit Sub
    End If
    Debug.Print "PopulateInstanceListColumn: Instance List column found at: " & instanceListCol
    
    ' Load instance data
    Set instanceData = LoadInstanceDataForMatching(wsInstances)
    If instanceData Is Nothing Then
        Debug.Print "PopulateInstanceListColumn: LoadInstanceDataForMatching returned Nothing"
        Exit Sub
    End If
    If instanceData.Count = 0 Then
        Debug.Print "PopulateInstanceListColumn: No instance data loaded"
        Exit Sub
    End If
    Debug.Print "PopulateInstanceListColumn: Loaded " & instanceData.Count & " instances"
    
    ' Find DR0 and DR1 columns in PVT_STA
    dr0Col = FindColumnInRow(wsPVT, "dr0", 1)
    dr1Col = FindColumnInRow(wsPVT, "dr1", 1)
    Debug.Print "PopulateInstanceListColumn: DR0 col=" & dr0Col & ", DR1 col=" & dr1Col
    
    ' Parse memory type columns from header
    Set memTypeColumns = ParseMemoryTypeColumns(wsPVT)
    
    Debug.Print "PopulateInstanceListColumn: Found " & memTypeColumns.Count & " memory type columns"

    ' Find VT columns (LVT / SVT / ULVT) in header row (supports names like LVT_1 etc.)
    Set vtCols = New Collection
    lastHeaderCol = wsPVT.Cells(1, wsPVT.Columns.Count).End(xlToLeft).Column
    For foundCol = 8 To lastHeaderCol
        headerText = Trim(CStr(wsPVT.Cells(1, foundCol).Value))
        If headerText <> "" Then
            If UCase(Left(headerText, 4)) = "ULVT" Then
                vtType = "ulvt"
            ElseIf UCase(Left(headerText, 3)) = "LVT" Then
                vtType = "lvt"
            ElseIf UCase(Left(headerText, 3)) = "SVT" Then
                vtType = "svt"
            Else
                vtType = ""
            End If

            If vtType <> "" Then
                Set vtInfo = CreateObject("Scripting.Dictionary")
                vtInfo("column") = foundCol
                vtInfo("vt_type") = vtType
                vtCols.Add vtInfo
            End If
        End If
    Next foundCol

    Debug.Print "PopulateInstanceListColumn: Found " & vtCols.Count & " VT columns"

    ' Pre-check: warn if memory types in PVTs don't exist in Instance List
    Dim instMemTypesDict As Object
    Set instMemTypesDict = CreateObject("Scripting.Dictionary")
    Dim instInfo2 As Object
    For instIdx = 1 To instanceData.Count
        Set instInfo2 = instanceData(instIdx)
        If Trim(CStr(instInfo2("memory_type"))) <> "" Then
            If Not instMemTypesDict.Exists(UCase(Trim(CStr(instInfo2("memory_type"))))) Then
                instMemTypesDict.Add UCase(Trim(CStr(instInfo2("memory_type")))), True
            End If
        End If
    Next instIdx

    Dim missingList As String
    missingList = ""
    ' Mark memory-type columns as valid/invalid based on whether any instance memory types match the pattern
    Dim instKey As Variant
    For Each colInfo In memTypeColumns
        memTypePattern = ExtractMemoryTypeFromColumnName(colInfo("memory_type")) ' e.g. "HDDP*"
        Dim foundInstanceMatch As Boolean
        foundInstanceMatch = False
        For Each instKey In instMemTypesDict.Keys
            If LCase(instKey) Like LCase(memTypePattern) Then
                foundInstanceMatch = True
                Exit For
            End If
        Next instKey
        ' Store validity flag on column info
        If colInfo.Exists("valid_in_instances") Then
            colInfo("valid_in_instances") = foundInstanceMatch
        Else
            colInfo.Add "valid_in_instances", foundInstanceMatch
        End If
        If Not foundInstanceMatch Then
            If missingList = "" Then
                missingList = Replace(memTypePattern, "*", "")
            Else
                missingList = missingList & ", " & Replace(memTypePattern, "*", "")
            End If
        End If
    Next

    If missingList <> "" Then
        MsgBox "WARNING: The following memory types are present in the PVT sheet but NOT in the N3P Instance List and will be ignored for matching: " & missingList & "." & vbCrLf & vbCrLf & _
               "This will NOT abort processing; these memory types are skipped so matches can still be found for other enabled types.", vbExclamation, "Memory type mismatch"
        Debug.Print "PopulateInstanceListColumn: Ignoring memory types (not present in instances): " & missingList
    End If

    ' Debugging support: create or find a debug column to record match counts/results
    Dim debugCol As Long
    debugCol = FindColumnInRow(wsPVT, "instance_match_debug", 1)
    If debugCol = -1 Then
        debugCol = wsPVT.Cells(1, wsPVT.Columns.Count).End(xlToLeft).Column + 1
        wsPVT.Cells(1, debugCol).Value = "Instance_Match_Debug"
    End If
    Debug.Print "PopulateInstanceListColumn: Debug column at " & debugCol
    
    ' Check if we have focus on DR columns
    hasDRFocus = False
    For focusIdx = 1 To focusColOrder.Count
        focusName = UCase(Trim(focusColOrder(focusIdx)))
        If focusName = "DR0" Or focusName = "DR1" Then
            hasDRFocus = True
            Exit For
        End If
    Next focusIdx
    
    ' Collect focused memory types
    Set focusedMemTypes = New Collection
    For focusIdx = 1 To focusColOrder.Count
        focusName = Trim(focusColOrder(focusIdx))
        If UCase(focusName) <> "DR0" And UCase(focusName) <> "DR1" Then
            focusedMemTypes.Add focusName
        End If
    Next focusIdx
    Debug.Print "PopulateInstanceListColumn: Focused memory types: " & focusedMemTypes.Count

    ' Prune focused memory types that have no valid memory-type columns present in Instance List
    If focusedMemTypes.Count > 0 Then
        Dim prunedFocused As Collection
        Set prunedFocused = New Collection
        Dim focusFound As Boolean
        For Each fm In focusedMemTypes
            focusFound = False
            For Each colInfo In memTypeColumns
                If colInfo.Exists("valid_in_instances") Then
                    If colInfo("valid_in_instances") Then
                        ' Compare header or base memory type
                        If UCase(Trim(colInfo("memory_type"))) = UCase(Trim(fm)) Then
                            focusFound = True
                            Exit For
                        End If
                        If UCase(Replace(ExtractMemoryTypeFromColumnName(colInfo("memory_type")), "*", "")) = _
                           UCase(Replace(ExtractMemoryTypeFromColumnName(fm), "*", "")) Then
                            focusFound = True
                            Exit For
                        End If
                    End If
                End If
            Next
            If focusFound Then
                prunedFocused.Add fm
            Else
                Debug.Print "PopulateInstanceListColumn: Ignoring focused memory type '" & fm & "' because it's not present in Instance List"
            End If
        Next fm
        Set focusedMemTypes = prunedFocused
        Debug.Print "PopulateInstanceListColumn: Focused memory types after pruning: " & focusedMemTypes.Count
    End If

    ' Build band DR conditions if DR focus
    Set bandDRConditions = CreateObject("Scripting.Dictionary")
    If hasDRFocus Then
        lastRow = wsPVT.Cells(wsPVT.Rows.Count, "A").End(xlUp).Row
        For checkRow = 3 To lastRow
            bandName = Trim(CStr(wsPVT.Cells(checkRow, 1).Value))
            If bandName <> "" Then
                If Not bandDRConditions.Exists(bandName) Then
                    bandDRConditions.Add bandName, CreateObject("Scripting.Dictionary")
                End If
                
                checkDR0 = ""
                checkDR1 = ""
                If dr0Col > 0 Then checkDR0 = Trim(UCase(CStr(wsPVT.Cells(checkRow, dr0Col).Value)))
                If dr1Col > 0 Then checkDR1 = Trim(UCase(CStr(wsPVT.Cells(checkRow, dr1Col).Value)))
                
                If checkDR0 = "YES" Or checkDR0 = "Y" Then bandDRConditions(bandName)("DR0") = True
                If checkDR1 = "YES" Or checkDR1 = "Y" Then bandDRConditions(bandName)("DR1") = True
            End If
        Next checkRow
    End If
    
    ' Process each data row
    ' Use Column C (PVT Name) instead of Column A (Band) because Column A is merged
    lastRow = wsPVT.Cells(wsPVT.Rows.Count, 3).End(xlUp).Row
    
    For pvtRow = 3 To lastRow
        ' Skip rows that do not contain a PVT name in column C or a Section in column B
        If Trim(CStr(wsPVT.Cells(pvtRow, 3).Value)) = "" And Trim(CStr(wsPVT.Cells(pvtRow, 2).Value)) = "" Then GoTo NextRowPop
        
        Debug.Print "PopulateInstanceListColumn: Processing row " & pvtRow
        
        ' Determine the band name for this row (band column A may be merged; use nearest non-empty above)
        Dim curBandName As String
        curBandName = Trim(CStr(wsPVT.Cells(pvtRow, 1).Value))
        If curBandName = "" Then curBandName = Trim(CStr(wsPVT.Cells(pvtRow, 1).End(xlUp).Value))
        Debug.Print "  Band=" & curBandName
        
        ' Get DR values
        dr0 = ""
        dr1 = ""
        If dr0Col > 0 Then dr0 = Trim(UCase(CStr(wsPVT.Cells(pvtRow, dr0Col).Value)))
        If dr1Col > 0 Then dr1 = Trim(UCase(CStr(wsPVT.Cells(pvtRow, dr1Col).Value)))
        Debug.Print "  DR0=" & dr0 & ", DR1=" & dr1
        
        ' Collect matching instances
        Set allInstances = New Collection
        ' Initialize skip reason trackers for diagnostics
        Set skipReasons = CreateObject("Scripting.Dictionary")
        Set reasonCounts = CreateObject("Scripting.Dictionary")
        Set reasonExamples = CreateObject("Scripting.Dictionary")
        exampleLimit = 3
        
        If hasDRFocus Then
            ' DR focus: match instances based on the DR values for THIS row
            ' This takes precedence over memory type logic when DR focus is specified
            Debug.Print "  Using DR focus logic for row"
            
            dr0Yes = (dr0 = "YES" Or dr0 = "Y")
            dr1Yes = (dr1 = "YES" Or dr1 = "Y")
            
            If Not dr0Yes And Not dr1Yes Then
                Debug.Print "  No DR condition present for this row - skipping DR-only matching"
            Else
                For instIdx = 1 To instanceData.Count
                    Set instInfo = instanceData(instIdx)
                    instDR = LCase(Trim(CStr(instInfo("dual_rail"))))
                    drMatch = False
                    If dr0Yes Then
                        If instDR = "dr0" Or instDR = "0" Then drMatch = True
                    End If
                    If dr1Yes Then
                        If instDR = "dr1" Or instDR = "1" Then drMatch = True
                    End If
                    
                    ' Check memory type match if focused
                    If drMatch And focusedMemTypes.Count > 0 Then
                        memMatch = False
                        For Each fm In focusedMemTypes
                            memTypePattern = ExtractMemoryTypeFromColumnName(fm)
                            If LCase(instInfo("memory_type")) Like LCase(memTypePattern) Then
                                memMatch = True
                                Exit For
                            End If
                        Next
                        If Not memMatch Then drMatch = False
                    End If
                    
                    ' Check write assist match if focused memory types have WA conditions
                    If drMatch And focusedMemTypes.Count > 0 Then
                        waMatch = False
                        For Each fm In focusedMemTypes
                            memTypePattern = ExtractMemoryTypeFromColumnName(fm)
                            If LCase(instInfo("memory_type")) Like LCase(memTypePattern) Then
                                ' Find the corresponding column to check for WA condition
                                        For Each colInfo In memTypeColumns
                                        If Not colInfo.Exists("valid_in_instances") Then
                                            ' assume valid if not marked
                                        ElseIf Not colInfo("valid_in_instances") Then
                                            GoTo NextCol_WA_Check
                                        End If

                                        colMemTypePattern = ExtractMemoryTypeFromColumnName(colInfo("memory_type"))
                                        If LCase(colMemTypePattern) Like LCase(memTypePattern) Then
                                            ' Check if column name contains (WA=1)
                                            If InStr(colInfo("memory_type"), "(WA=1)") > 0 Then
                                                If LCase(Trim(CStr(instInfo("write_assist")))) = "1" Then
                                                    waMatch = True
                                                End If
                                            Else
                                                waMatch = True ' No WA condition specified
                                            End If
                                            Exit For
                                        End If
NextCol_WA_Check:
                                    Next
                                    If waMatch Then Exit For
                            End If
                        Next
                        If Not waMatch Then drMatch = False
                    End If
                    
                    ' Check if the instance's memory type is enabled in this row
                    If drMatch And memMatch Then
                        ' Find the column for this instance's memory type and check if it's enabled
                        rowMemMatch = False
                        For Each colInfo In memTypeColumns
                            ' Skip columns that do not have matching instances
                            If colInfo.Exists("valid_in_instances") Then
                                If Not colInfo("valid_in_instances") Then GoTo NextCol_RowMem
                            End If

                            memTypePattern = ExtractMemoryTypeFromColumnName(colInfo("memory_type"))
                            If LCase(instInfo("memory_type")) Like LCase(memTypePattern) Then
                                ' Check if this memory type is enabled in the current row
                                cellVal = UCase(Trim(CStr(wsPVT.Cells(pvtRow, colInfo("column")).Value)))
                                If cellVal = "YES" Or cellVal = "Y" Then
                                    rowMemMatch = True
                                    Exit For
                                End If
                            End If
NextCol_RowMem:
                        Next
                        If Not rowMemMatch Then
                            drMatch = False
                        Else
                            ' VT check: if any VT columns are enabled (Yes) on this row, require instance name to contain p<vt>
                            vtMatched = False
                            vtHasYes = False
                            vtMismatch = False
                            If vtCols.Count > 0 Then
                                For Each vtInfo In vtCols
                                    cellVal = UCase(Trim(CStr(wsPVT.Cells(pvtRow, vtInfo("column")).Value)))
                                    If cellVal = "YES" Or cellVal = "Y" Then
                                        vtHasYes = True
                                        If InStr(1, LCase(instInfo("instance_name")), "p" & vtInfo("vt_type"), vbTextCompare) > 0 Then
                                            vtMatched = True
                                            Exit For
                                        End If
                                    End If
                                Next
                                If vtHasYes And Not vtMatched Then
                                    drMatch = False
                                    vtMismatch = True
                                End If
                            End If
                        End If
                    End If
                    
                    ' Collect skip reason if this instance won't be added
                    If Not drMatch Then
                        reasonKey = "DR_MISMATCH"
                    ElseIf vtMismatch Then
                        reasonKey = "VT_MISMATCH"
                    ElseIf focusedMemTypes.Count > 0 And Not memMatch Then
                        reasonKey = "MEMTYPE_MISMATCH"
                    ElseIf focusedMemTypes.Count > 0 And Not waMatch Then
                        reasonKey = "WA_MISMATCH"
                    ElseIf focusedMemTypes.Count > 0 And Not rowMemMatch Then
                        reasonKey = "ROW_MEM_DISABLED"
                    Else
                        reasonKey = "NO_MATCH_REASON_SPECIFIED"
                    End If
                    
                    If Not drMatch Or (focusedMemTypes.Count > 0 And (Not memMatch Or Not waMatch Or Not rowMemMatch)) Then
                        If Not reasonCounts.Exists(reasonKey) Then
                            reasonCounts.Add reasonKey, 1
                        Else
                            reasonCounts(reasonKey) = reasonCounts(reasonKey) + 1
                        End If
                        instName = Trim(CStr(instInfo("instance_name")))
                        If Not reasonExamples.Exists(reasonKey) Then
                            reasonExamples.Add reasonKey, instName
                        Else
                            exStr = reasonExamples(reasonKey)
                            If InStr(1, exStr, instName, vbTextCompare) = 0 Then
                                exCount = UBound(Split(exStr, ",")) + 1
                                If exCount < exampleLimit Then reasonExamples(reasonKey) = exStr & "," & instName
                            End If
                        End If
                    End If
                    
                    If drMatch Then
                        instName = Trim(CStr(instInfo("instance_name")))
                        If instName <> "" Then
                            ' prevent duplicates
                            alreadyExists = False
                            For k = 1 To allInstances.Count
                                If allInstances(k) = instName Then
                                    alreadyExists = True
                                    Exit For
                                End If
                            Next k
                            If Not alreadyExists Then
                                allInstances.Add instName
                                Debug.Print "  Added instance: " & instName & " (DR=" & instDR & ", MemType=" & instInfo("memory_type") & ")"
                            End If
                        End If
                    End If
                Next instIdx
            End If
        Else
            ' Memory type focus: match instances based on memory types that are YES in this row
            Debug.Print "  Using memory type focus logic for row"
            
            If memTypeColumns.Count > 0 Then
                For instIdx = 1 To instanceData.Count
                    Set instInfo = instanceData(instIdx)
                    memMatch = False
                    
                    ' Check if any focused memory type matches this instance
                    For Each colInfo In memTypeColumns
                        ' Skip columns with no matching instances
                        If colInfo.Exists("valid_in_instances") Then
                            If Not colInfo("valid_in_instances") Then GoTo NextCol_MemMatch
                        End If

                        memTypePattern = ExtractMemoryTypeFromColumnName(colInfo("memory_type"))
                        If LCase(instInfo("memory_type")) Like LCase(memTypePattern) Then
                            ' Check if this memory type is enabled in the current row
                            cellVal = UCase(Trim(CStr(wsPVT.Cells(pvtRow, colInfo("column")).Value)))
                            If cellVal = "YES" Or cellVal = "Y" Then
                                memMatch = True
                                Exit For
                            End If
                        End If
NextCol_MemMatch:
                    Next
                    
                    ' Also check write_assist if the column requires it
                    If memMatch Then
                        ' Check if any of the matching columns require write_assist
                        waRequired = False
                        waMatch = True
                        For Each colInfo In memTypeColumns
                            memTypePattern = ExtractMemoryTypeFromColumnName(colInfo("memory_type"))
                            If LCase(instInfo("memory_type")) Like LCase(memTypePattern) Then
                                ' Check if this column is enabled in the current row
                                cellVal = UCase(Trim(CStr(wsPVT.Cells(pvtRow, colInfo("column")).Value)))
                                If cellVal = "YES" Or cellVal = "Y" Then
                                    ' Check if this column requires WA=1
                                    If InStr(colInfo("memory_type"), "(WA=1)") > 0 Then
                                        waRequired = True
                                        If LCase(Trim(CStr(instInfo("write_assist")))) <> "1" Then
                                            waMatch = False
                                            Exit For
                                        End If
                                    End If
                                End If
                            End If
                        Next
                        If waRequired And Not waMatch Then memMatch = False
                    End If
                    
                    If memMatch Then
                        ' VT check for memory-type-focused rows: if any VT columns are enabled (Yes), require instance name to contain p<vt>
                        vtMatched = False
                        vtHasYes = False
                        If vtCols.Count > 0 Then
                            For Each vtInfo In vtCols
                                cellVal = UCase(Trim(CStr(wsPVT.Cells(pvtRow, vtInfo("column")).Value)))
                                If cellVal = "YES" Or cellVal = "Y" Then
                                    vtHasYes = True
                                    If InStr(1, LCase(instInfo("instance_name")), "p" & vtInfo("vt_type"), vbTextCompare) > 0 Then
                                        vtMatched = True
                                        Exit For
                                    End If
                                End If
                            Next
                            If vtHasYes And Not vtMatched Then
                                ' instance doesn't match enabled VT(s) - skip
                                GoTo NextInstInMemFocus
                            End If
                        End If

                        instName = Trim(CStr(instInfo("instance_name")))
                        If instName <> "" Then
                            ' prevent duplicates
                            alreadyExists = False
                            For k = 1 To allInstances.Count
                                If allInstances(k) = instName Then
                                    alreadyExists = True
                                    Exit For
                                End If
                            Next k
                            If Not alreadyExists Then
                                allInstances.Add instName
                                Debug.Print "  Added instance: " & instName & " (MemType=" & instInfo("memory_type") & ", WA=" & instInfo("write_assist") & ")"
                            End If
                        End If
                    End If
                    NextInstInMemFocus:
                Next instIdx
            End If
        End If
        
        ' Write instances to cell
        instanceListStr = ""
        For i = 1 To allInstances.Count
            If instanceListStr = "" Then
                instanceListStr = allInstances(i)
            Else
                instanceListStr = instanceListStr & ", " & allInstances(i)
            End If
        Next i
        
        If instanceListStr = "" Then
            ' Diagnostic: list enabled memory types and reason for no matches
            Dim enabledList As String
            enabledList = ""
            Dim cv As String
            For Each colInfo In memTypeColumns
                cv = UCase(Trim(CStr(wsPVT.Cells(pvtRow, colInfo("column")).Value)))
                If cv = "YES" Or cv = "Y" Then
                    If enabledList = "" Then
                        enabledList = colInfo("memory_type")
                    Else
                        enabledList = enabledList & ", " & colInfo("memory_type")
                    End If
                End If
            Next
            Debug.Print "  No instances for row " & pvtRow & " (Band=" & curBandName & "). DR0=" & dr0 & ", DR1=" & dr1 & ", EnabledMemTypes=" & enabledList & ", FocusedMemTypesCount=" & focusedMemTypes.Count & ", hasDRFocus=" & hasDRFocus
        End If
        
        Debug.Print "  Final instance list for row " & pvtRow & ": " & instanceListStr
        wsPVT.Cells(pvtRow, instanceListCol).Value = instanceListStr
        ' Write debug info to debug column
        On Error Resume Next
        wsPVT.Cells(pvtRow, debugCol).Value = allInstances.Count
        If allInstances.Count = 0 Then
            wsPVT.Cells(pvtRow, debugCol).Interior.Color = RGB(255, 199, 206) ' light red for zero
            ' Build reason summary and examples
            Dim reasonSummary As String
            reasonSummary = ""
            Dim rk As Variant
            For Each rk In reasonCounts.Keys
                If reasonSummary = "" Then
                    reasonSummary = rk & ":" & reasonCounts(rk) & " (" & reasonExamples(rk) & ")"
                Else
                    reasonSummary = reasonSummary & "; " & rk & ":" & reasonCounts(rk) & " (" & reasonExamples(rk) & ")"
                End If
            Next rk
            If reasonSummary = "" Then reasonSummary = "No matches and no skip reasons captured"
            ' Write to debug adjacent column
            wsPVT.Cells(pvtRow, debugCol + 1).Value = reasonSummary
            wsPVT.Cells(pvtRow, debugCol + 1).Interior.Color = RGB(255, 242, 204) ' light yellow
        Else
            wsPVT.Cells(pvtRow, debugCol).Interior.Color = RGB(198, 239, 206) ' light green for >0
        End If
        On Error GoTo ErrorHandler
        
NextRowPop:
    Next pvtRow
    
    Debug.Print "PopulateInstanceListColumn: Completed"
    Exit Sub
    
ErrorHandler:
    Debug.Print "PopulateInstanceListColumn ERROR: " & Err.Number & " - " & Err.Description
End Sub

Private Function FindColumnInRow(ws As Worksheet, headerName As String, headerRow As Long) As Long
    Dim col As Long
    For col = 1 To ws.Cells(headerRow, ws.Columns.Count).End(xlToLeft).Column
        If UCase(Trim(ws.Cells(headerRow, col).Value)) = UCase(headerName) Then
            FindColumnInRow = col
            Exit Function
        End If
    Next col
    FindColumnInRow = -1
End Function

Private Function LoadInstanceDataForMatching(ws As Worksheet) As Collection
    Dim data As Collection
    Set data = New Collection
    
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    
    Dim memCol As Long, waCol As Long, instCol As Long, drCol As Long
    memCol = FindColumnInRow(ws, "memory_type", 2)
    waCol = FindColumnInRow(ws, "write_assist", 2)
    instCol = FindColumnInRow(ws, "instance_name", 2)
    drCol = FindColumnInRow(ws, "dual_rail", 2)
    
    If memCol = -1 Or instCol = -1 Then
        Set LoadInstanceDataForMatching = data
        Exit Function
    End If
    
    Dim i As Long
    For i = 3 To lastRow
        Dim memType As String, wa As String, instName As String, dr As String
        memType = Trim(CStr(ws.Cells(i, memCol).Value))
        wa = Trim(CStr(ws.Cells(i, waCol).Value))
        instName = Trim(CStr(ws.Cells(i, instCol).Value))
        dr = Trim(CStr(ws.Cells(i, drCol).Value))
        
        If memType <> "" And instName <> "" And Not (LCase(Left(memType, 4)) = "hddp" And wa <> "1") Then
            Dim instanceInfo As Object
            Set instanceInfo = CreateObject("Scripting.Dictionary")
            instanceInfo("memory_type") = memType
            instanceInfo("write_assist") = wa
            instanceInfo("instance_name") = instName
            instanceInfo("dual_rail") = dr
            data.Add instanceInfo
        End If
    Next i
    
    Set LoadInstanceDataForMatching = data
End Function

Private Function ParseMemoryTypeColumns(ws As Worksheet) As Collection
    Dim memTypeColumns As Collection
    Set memTypeColumns = New Collection
    
    Dim lastCol As Long
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    
    Dim col As Long
    For col = 8 To lastCol
        Dim headerText As String
        headerText = Trim(CStr(ws.Cells(1, col).Value))
        
        If headerText <> "" And InStr(headerText, "(") > 0 And InStr(UCase(headerText), "WA=") > 0 Then
            Dim memType As String
            Dim waValues() As String
            If ParseMemTypeHeader(headerText, memType, waValues) Then
                Dim colInfo As Object
                Set colInfo = CreateObject("Scripting.Dictionary")
                colInfo("column") = col
                colInfo("memory_type") = memType
                colInfo("wa_values") = waValues
                memTypeColumns.Add colInfo
            End If
        End If
    Next col
    
    Set ParseMemoryTypeColumns = memTypeColumns
End Function

Private Function ParseMemTypeHeader(ByVal headerText As String, ByRef memoryType As String, ByRef waValues() As String) As Boolean
    Dim cleanText As String, parenPos As Long, waText As String
    
    ParseMemTypeHeader = False
    memoryType = ""
    ReDim waValues(0)
    
    cleanText = Replace(headerText, vbLf, " ")
    cleanText = Replace(cleanText, vbCr, " ")
    cleanText = Trim(cleanText)
    
    parenPos = InStr(cleanText, "(")
    If parenPos = 0 Then Exit Function
    
    memoryType = Trim(Left(cleanText, parenPos - 1))
    
    waText = Mid(cleanText, parenPos)
    waText = Replace(waText, "(", "")
    waText = Replace(waText, ")", "")
    waText = Replace(UCase(waText), "WA=", "")
    waText = Trim(waText)
    
    If InStr(waText, "/") > 0 Then
        Dim parts() As String
        parts = Split(waText, "/")
        ReDim waValues(LBound(parts) To UBound(parts))
        Dim i As Long
        For i = LBound(parts) To UBound(parts)
            waValues(i) = Trim(parts(i))
        Next i
    Else
        ReDim waValues(0)
        waValues(0) = waText
    End If
    
    ParseMemTypeHeader = (memoryType <> "" And waValues(0) <> "")
End Function

Private Function FindMatchingInstancesForRow(instanceData As Collection, memoryType As String, writeAssist As String, dr0 As String, dr1 As String) As Collection
    Dim matches As Collection
    Set matches = New Collection
    
    Dim i As Long
    For i = 1 To instanceData.Count
        Dim inst As Object
        Set inst = instanceData(i)
        
        If UCase(Trim(inst("memory_type"))) <> UCase(Trim(memoryType)) Then GoTo NextInstMatch
        
        Dim instWA As String
        instWA = Trim(CStr(inst("write_assist")))
        If writeAssist <> "" And instWA <> Trim(writeAssist) Then GoTo NextInstMatch
        
        Dim drMatch As Boolean, instDR As String
        drMatch = False
        instDR = Trim(CStr(inst("dual_rail")))
        
        Dim isDR0Yes As Boolean, isDR1Yes As Boolean
        isDR0Yes = (UCase(Trim(dr0)) = "YES" Or UCase(Trim(dr0)) = "Y")
        isDR1Yes = (UCase(Trim(dr1)) = "YES" Or UCase(Trim(dr1)) = "Y")
        
        If Not isDR0Yes And Not isDR1Yes Then
            drMatch = True
        Else
            If isDR0Yes And instDR = "0" Then drMatch = True
            If isDR1Yes And instDR = "1" Then drMatch = True
        End If
        
        If Not drMatch Then GoTo NextInstMatch
        
        matches.Add inst("instance_name")
        
NextInstMatch:
    Next i
    
    Set FindMatchingInstancesForRow = matches
End Function
