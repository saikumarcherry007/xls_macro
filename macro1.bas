'==========================================
' PVT Processing generates PVT_STA
' Author: Saikumar Malluru
' Created: 2025-12-10
' Version: 1.0.0
' Description: Processes PVT data from the sheet defined by SHEET_PVT and creates a SHEET_OUTPUT report.
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

' VT prefixes default (comma-separated patterns, wildcards allowed)
Const DEFAULT_VT_PREFIXES As String = ""

' DR (Dual Rail) prefixes default (comma-separated patterns, wildcards allowed)
Const DEFAULT_DR_PREFIXES As String = ""

' Hold-only Mappings Default (semicolon-separated pattern:extraction pairs)
' Example: "SS*:Cbest;SSGNP*:Cbest,Cworst"
Const DEFAULT_HOLD_ONLY_MAPPINGS As String = ""

' Sheet Names
Const SHEET_PVT As String = "PVTs"
Const SHEET_INSTANCES As String = "N3P Instance List"
Const SHEET_OUTPUT As String = "PVT_STA"
Const SHEET_CONFIG As String = "Generate PVT_STA"
Const SHEET_VARIANCE As String = "var_sheet"

'==========================================

Public Function GetSupportedVTs() As Collection
    ' Return VT column prefixes configured on the Generate PVT_STA sheet (cell C18),
    ' or fallback to defaults (ULVT,LVT,SVT) if the cell is blank or not present.
    Dim c As Collection
    Set c = New Collection
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(SHEET_CONFIG)
    On Error GoTo 0

    Dim raw As String
    raw = ""
    If Not ws Is Nothing Then
        raw = Trim(CStr(ws.Range("C17").Value))
    End If

    If raw <> "" Then
        Dim parts() As String, p As Variant
        parts = Split(raw, ",")
        For Each p In parts
            p = Trim(CStr(p))
            If p <> "" Then
                p = UCase(p)
                ' If user omitted wildcard, treat as prefix (append *) so LVT matches LVT_1 etc.
                If InStr(p, "*") = 0 Then p = p & "*"
                c.Add p
            End If
        Next p
    End If

    If c.Count = 0 Then
        Dim dvParts() As String, dv As Variant
        dvParts = Split(DEFAULT_VT_PREFIXES, ",")
        For Each dv In dvParts
            dv = Trim(CStr(dv))
            If dv <> "" Then
                If InStr(dv, "*") = 0 Then dv = dv & "*"
                c.Add UCase(dv)
            End If
        Next dv
    End If

    Set GetSupportedVTs = c
End Function

Public Function GetSupportedDRs() As Collection
    ' Return DR (Dual Rail) column prefixes configured on the Generate PVT_STA sheet (cell C19),
    ' or fallback to defaults (DR0,DR1) if the cell is blank or not present.
    Dim c As Collection
    Set c = New Collection
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(SHEET_CONFIG)
    On Error GoTo 0

    Dim raw As String
    raw = ""
    If Not ws Is Nothing Then
        raw = Trim(CStr(ws.Range("C18").Value))
    End If

    If raw <> "" Then
        Dim parts() As String, p As Variant
        parts = Split(raw, ",")
        For Each p In parts
            p = Trim(CStr(p))
            If p <> "" Then
                p = UCase(p)
                ' If user omitted wildcard, treat as prefix (append *) so DR0 matches DR0_1 etc.
                If InStr(p, "*") = 0 Then p = p & "*"
                c.Add p
            End If
        Next p
    End If

    If c.Count = 0 Then
        Dim dvParts() As String, dv As Variant
        dvParts = Split(DEFAULT_DR_PREFIXES, ",")
        For Each dv In dvParts
            dv = Trim(CStr(dv))
            If dv <> "" Then
                If InStr(dv, "*") = 0 Then dv = dv & "*"
                c.Add UCase(dv)
            End If
        Next dv
    End If

    Set GetSupportedDRs = c
End Function


'==========================================
' INSTANCE MATCHER HELPER FUNCTIONS
'==========================================



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
            sourceColIndex = CLng(focusColMapping(colName))
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
        
        ' Set column widths for better layout
        .Columns("A").ColumnWidth = 5
        .Columns("B").ColumnWidth = 25
        .Columns("C").ColumnWidth = 20
        .Columns("D").ColumnWidth = 20
        .Columns("E").ColumnWidth = 20
        .Columns("F").ColumnWidth = 20
        
        .Range("A3:G3").Merge
        .Range("A3").Value = "Click the button below to process PVT data and generate the " & SHEET_OUTPUT & " sheet."
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
    .Range("C6").Value = "Read data from " & SHEET_PVT & " Sheet"
    .Range("C7").Value = "Process all PVT corners (SETUP and HOLD)"
    .Range("C8").Value = "Create organized output in " & SHEET_OUTPUT & " sheet"
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
    .Range("B12:B18").WrapText = True
        
    
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

    
    ' Clearing of rows 19-22 removed to preserve Dual Rail Filter configuration
    On Error Resume Next
    ' Only clear if absolutely necessary, but here it was wiping user config
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

    
    ' Hold-only mappings: pattern:extraction1,extraction2;pattern2:extrA,extrB
    If Trim(CStr(.Range("B15").Value)) = "" Then .Range("B15").Value = "Hold-only mappings"
    .Range("B15").Font.Bold = True
    If Trim(CStr(.Range("C15").Value)) = "" Then .Range("C15").Value = DEFAULT_HOLD_ONLY_MAPPINGS ' e.g. SS*:Cbest;SSGNP*:Cbest,Cworst"
    .Range("C15:F15").Merge
    .Range("C15:F15").WrapText = True
    .Range("C15:F15").ShrinkToFit = True
    .Range("C15:F15").HorizontalAlignment = xlLeft
    .Range("C15:F15").VerticalAlignment = xlCenter

    If Trim(CStr(.Range("B16").Value)) = "" Then .Range("B16").Value = "Auto memory Filter"
    .Range("B16").Font.Bold = True
    
    Dim combinedMsg As String
    combinedMsg = ""
    Dim combinedWarningMsg As String
    combinedWarningMsg = ""
    
    ' Auto-populate Focus Columns if cell is empty or has default value
    If Trim(CStr(.Range("C16").Value)) = "" Or Trim(CStr(.Range("C16").Value)) = DEFAULT_FOCUS_COLUMNS Then
        Dim autoFocus As String
        Dim focusMsg As String
        Dim unmatchedMsg As String
        autoFocus = GetUniqueMemoryTypesFromInstanceList(focusMsg, unmatchedMsg)
        If autoFocus <> "" Then
            .Range("C16").Value = autoFocus
            If combinedMsg <> "" Then combinedMsg = combinedMsg & vbCrLf & vbCrLf
            combinedMsg = combinedMsg & "--- Auto memory Filter Discovery ---" & vbCrLf & focusMsg
            
            ' Collect warning for unmatched memory types
            If unmatchedMsg <> "" Then
                If combinedWarningMsg <> "" Then combinedWarningMsg = combinedWarningMsg & vbCrLf & vbCrLf
                combinedWarningMsg = combinedWarningMsg & "--- Unmatched Memory Types ---" & vbCrLf & _
                                     "The following memory types from Instance List were NOT found in " & SHEET_PVT & " sheet:" & vbCrLf & unmatchedMsg & vbCrLf & _
                                     "These will be excluded from the Auto memory Filter."
            End If
        Else
            .Range("C16").Value = DEFAULT_FOCUS_COLUMNS
        End If
    End If
    
    .Range("C16:F16").Merge
    .Range("C16:F16").WrapText = True
    .Range("C16:F16").ShrinkToFit = True
    .Range("C16:F16").HorizontalAlignment = xlLeft
    .Range("C16:F16").VerticalAlignment = xlCenter

    ' Patterns will be entered in C12:C14 (left) alongside extraction inputs in E12:E14 (right)
    
    ' VT prefixes configuration (comma-separated values, e.g. ULVT,LVT,SVT)
    If Trim(CStr(.Range("B17").Value)) = "" Or Left(Trim(CStr(.Range("B17").Value)), 4) = "NOTE" Then .Range("B17").Value = "Auto VT Filters"
    .Range("B17").Font.Bold = True
    
    ' Auto-populate VT values if cell is empty or has default value
    If Trim(CStr(.Range("C17").Value)) = "" Or Trim(CStr(.Range("C17").Value)) = DEFAULT_VT_PREFIXES Then
        Dim autoVTs As String
        Dim vtMsg As String
        Dim vtUnmatchedMsg As String
        autoVTs = GetUniqueVTValuesFromInstanceList(vtMsg, vtUnmatchedMsg)
        If autoVTs <> "" Then
            .Range("C17").Value = autoVTs
            If combinedMsg <> "" Then combinedMsg = combinedMsg & vbCrLf & vbCrLf
            combinedMsg = combinedMsg & "--- VT Discovery ---" & vbCrLf & vtMsg
            
            ' Collect warning for unmatched VT types
            If vtUnmatchedMsg <> "" Then
                If combinedWarningMsg <> "" Then combinedWarningMsg = combinedWarningMsg & vbCrLf & vbCrLf
                combinedWarningMsg = combinedWarningMsg & "--- VT Configuration Warnings ---" & vbCrLf & _
                                     "The following VT types found in instance names were NOT found in " & SHEET_PVT & " sheet:" & vbCrLf & vtUnmatchedMsg & vbCrLf & _
                                     "These will be excluded from the VT Filters."
            End If
        Else
            .Range("C17").Value = DEFAULT_VT_PREFIXES
        End If
    End If
    
    .Range("C17:F17").Merge
    .Range("C17:F17").WrapText = True
    .Range("C17:F17").ShrinkToFit = True
    .Range("C17:F17").HorizontalAlignment = xlLeft
    .Range("C17:F17").VerticalAlignment = xlCenter
    .Range("A17").RowHeight = 30
    
    ' DR (Dual Rail) prefixes configuration (comma-separated values, e.g. DR0,DR1)
    If Trim(CStr(.Range("B18").Value)) = "" Then .Range("B18").Value = "Auto Dual_Rail"
    .Range("B18").Font.Bold = True
    
    ' Auto-populate DR values if cell is empty or has default value
    If Trim(CStr(.Range("C18").Value)) = "" Or Trim(CStr(.Range("C18").Value)) = DEFAULT_DR_PREFIXES Then
        Dim autoDRs As String
        Dim drMsg As String
        autoDRs = GetUniqueDRValuesFromInstanceList(drMsg)
        If autoDRs <> "" Then
            .Range("C18").Value = autoDRs
            If combinedMsg <> "" Then combinedMsg = combinedMsg & vbCrLf & vbCrLf
            combinedMsg = combinedMsg & "--- Dual Rail Discovery ---" & vbCrLf & drMsg
        Else
            .Range("C18").Value = DEFAULT_DR_PREFIXES
            ' Collect warning if we expected auto-pop but got nothing (and default is empty)
            If DEFAULT_DR_PREFIXES = "" Then
                If combinedWarningMsg <> "" Then combinedWarningMsg = combinedWarningMsg & vbCrLf & vbCrLf
                combinedWarningMsg = combinedWarningMsg & "--- Dual Rail Auto-Population Failed ---" & vbCrLf & _
                                     "Could not auto-populate Dual Rail filters from '" & SHEET_INSTANCES & "'. Please check if the sheet exists and has a 'dual_rail' column with data."
            End If
        End If
    End If
    
    ' Show combined warning message if any warnings occurred
    If combinedWarningMsg <> "" Then
        MsgBox "Auto-Configuration Warnings:" & vbCrLf & vbCrLf & combinedWarningMsg, vbExclamation, "Auto-Configuration Warnings"
    End If
    
    ' Show combined info message if any discovery happened
    If combinedMsg <> "" Then
        MsgBox "Auto-Discovery Successful:" & vbCrLf & vbCrLf & combinedMsg, vbInformation, "Auto-Configuration"
    End If
    
    .Range("C18:F18").Merge
    .Range("C18:F18").WrapText = True
    .Range("C18:F18").ShrinkToFit = True
    .Range("C18:F18").HorizontalAlignment = xlLeft
    .Range("C18:F18").VerticalAlignment = xlCenter
    .Range("A18").RowHeight = 30
    
    ' Custom Condition Filter configuration (Row 19)
    If Trim(CStr(.Range("B19").Value)) = "" Or Left(Trim(CStr(.Range("B19").Value)), 4) = "NOTE" Then .Range("B19").Value = "Custom Cond Filter"
    .Range("B19").Font.Bold = True
    
    ' Default empty for custom filter
    If Trim(CStr(.Range("C19").Value)) = "" Then .Range("C19").Value = ""
    
    .Range("C19:F19").Merge
    .Range("C19:F19").WrapText = True
    .Range("C19:F19").ShrinkToFit = True
    .Range("C19:F19").HorizontalAlignment = xlLeft
    .Range("C19:F19").VerticalAlignment = xlCenter
    .Range("A19").RowHeight = 30

    ' Apply table formatting to the extended range (now including Custom Cond Filter row)
    .Range("B12:F19").Borders.LineStyle = xlContinuous
    .Range("B12:B19").Interior.Color = RGB(221, 235, 255) ' light blue for labels
    .Range("C12:F19").Interior.Color = RGB(255, 255, 255) ' white input background
    .Range("B12:B19").Font.Color = RGB(68, 114, 196)
    .Range("C12:F19").Borders.Color = RGB(191, 191, 191)
    .Range("B12:B19").Borders.Color = RGB(191, 191, 191)
    .Range("B12:B19").Font.Bold = True
    .Range("B12:B19").WrapText = True
    .Range("B12:B19").HorizontalAlignment = xlLeft
    .Range("B12:B19").VerticalAlignment = xlCenter

    ' NOTE section moved to Row 20
    If Trim(CStr(.Range("B20").Value)) = "" Or Left(Trim(CStr(.Range("B20").Value)), 4) <> "NOTE" Then .Range("B20").Value = "NOTE: You can override default values above. Enter comma-separated PVT name patterns (left) and extraction types (right). Examples: SS*, FF*, TT*. To force a PVT into HOLD with custom extractions use semicolon-separated mappings in 'Hold-only mappings' (e.g.: SS*:Cbest;SSGNP*:Cbest,Cworst). Wildcards supported."
    .Range("B20:F20").Merge
    .Range("B20").Font.Size = 9
    .Range("B20").Font.Color = RGB(100, 100, 100)
    .Range("B20").HorizontalAlignment = xlLeft
    .Range("B20").WrapText = True
    .Range("B20").Font.Bold = False ' Ensure note is not bold
    .Range("A20").RowHeight = 30
    
    
    ' Set column widths first to ensure consistent positioning
    .Columns("A").ColumnWidth = 2
    .Columns("B").ColumnWidth = 35
    .Columns("C:F").ColumnWidth = 30
    .Columns("G").ColumnWidth = 2
    
    ' Set row height for button area (moved to row 21)
    .Range("A21").RowHeight = 60
    .Range("A22").RowHeight = 60
    
    Dim btnLeft As Double, btnTop As Double, btnWidth As Double, btnHeight As Double
    btnWidth = 240 ' Width in points - consistent smaller size
    btnHeight = 40 ' Height in points - consistent smaller size
    
    ' Calculate center position using actual column positions (now that widths are set)
    Dim centerPoint As Double
    centerPoint = .Range("B21").Left + (.Range("F21").Left + .Range("F21").Width - .Range("B21").Left) / 2
    btnLeft = centerPoint - (btnWidth / 2)
    btnTop = .Range("A21").Top + 10 ' 10 points padding from top of row 21

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
        
        ' Add visual annotations (Braces and Text boxes)
        AddConfigAnnotations wsPVT
        On Error GoTo 0

        wsPVT.Activate
        wsPVT.Range("A1").Select
End Sub

Private Sub AddConfigAnnotations(ws As Worksheet)
    ' Adds curly braces and text boxes to the right of the config table
    ' Group 1: Rows 12-14 (Setup/Hold/Typical) -> "Manually need to update"
    ' Group 2: Rows 16-18 (Auto Filters) -> "Automatically filters will be applied"
    ' Group 3: Row 19 (Custom Cond) -> "Used as a condition"

    Dim shp As Shape
    Dim rStart As Range, rEnd As Range
    Dim topPos As Double, botPos As Double, heightVal As Double, centerPos As Double
    Dim braceLeft As Double, textLeft As Double
    
    ' Cleanup existing annotations
    On Error Resume Next
    For Each shp In ws.Shapes
        If Left(shp.Name, 12) = "ConfigAnnot_" Then shp.Delete
    Next shp
    On Error GoTo 0
    
    ' Column G is narrow (width 2), put things starting from H
    braceLeft = ws.Range("H1").Left
    textLeft = braceLeft + 20 ' 20 points gap
    
    ' --- Group 1: Rows 12-14 ---
    Set rStart = ws.Range("F12")
    Set rEnd = ws.Range("F14")
    
    topPos = rStart.Top
    botPos = rEnd.Top + rEnd.Height
    heightVal = botPos - topPos
    centerPos = topPos + (heightVal / 2)
    
    ' Draw Brace
    Set shp = ws.Shapes.AddShape(msoShapeRightBrace, braceLeft, topPos, 15, heightVal)
    With shp
        .Name = "ConfigAnnot_Brace1"
        .Line.ForeColor.RGB = RGB(68, 114, 196)
        .Line.Weight = 1.5
        .Fill.Visible = msoFalse
    End With
    
    ' Draw Text Box
    Set shp = ws.Shapes.AddShape(msoShapeRoundedRectangle, textLeft, centerPos - 15, 120, 40)
    With shp
        .Name = "ConfigAnnot_Text1"
        .TextFrame.Characters.Text = "Ex data are captured, Manually need to update "
        .TextFrame.Characters.Font.Size = 9
        .TextFrame.Characters.Font.Name = "Calibri"
        .TextFrame.Characters.Font.Color = RGB(0, 0, 0)
        .TextFrame.HorizontalAlignment = xlHAlignCenter
        .TextFrame.VerticalAlignment = xlVAlignCenter
        .Fill.ForeColor.RGB = RGB(255, 255, 255)
        .Line.ForeColor.RGB = RGB(0, 0, 0)
        .Line.Weight = 1
    End With
    
    ' --- Group 2: Rows 16-18 ---
    Set rStart = ws.Range("F16")
    Set rEnd = ws.Range("F18")
    
    topPos = rStart.Top
    botPos = rEnd.Top + rEnd.Height
    heightVal = botPos - topPos
    centerPos = topPos + (heightVal / 2)
    
    ' Draw Brace
    Set shp = ws.Shapes.AddShape(msoShapeRightBrace, braceLeft, topPos, 15, heightVal)
    With shp
        .Name = "ConfigAnnot_Brace2"
        .Line.ForeColor.RGB = RGB(68, 114, 196)
        .Line.Weight = 1.5
        .Fill.Visible = msoFalse
    End With
    
    ' Draw Text Box
    Set shp = ws.Shapes.AddShape(msoShapeRoundedRectangle, textLeft, centerPos - 20, 120, 50) ' Slightly taller
    With shp
        .Name = "ConfigAnnot_Text2"
        .TextFrame.Characters.Text = "Automatically filters will be applied,  Edit for manual (Not recommended)"
        .TextFrame.Characters.Font.Size = 9
        .TextFrame.Characters.Font.Name = "Calibri"
        .TextFrame.Characters.Font.Color = RGB(0, 0, 0)
        .TextFrame.HorizontalAlignment = xlHAlignCenter
        .TextFrame.VerticalAlignment = xlVAlignCenter
        .Fill.ForeColor.RGB = RGB(255, 255, 255)
        .Line.ForeColor.RGB = RGB(0, 0, 0)
        .Line.Weight = 1
    End With
    
    ' --- Group 3: Row 19 ---
    Set rStart = ws.Range("F19")
    
    topPos = rStart.Top
    heightVal = rStart.Height
    centerPos = topPos + (heightVal / 2)
    
    ' For single row, maybe just an arrow or small brace? User image shows arrow-like pointer but text says "Used as a condition".
    ' Let's use a small brace centered on the row
    Set shp = ws.Shapes.AddShape(msoShapeRightArrow, braceLeft, centerPos - 6, 20, 12) ' Arrow pointer
    With shp
        .Name = "ConfigAnnot_Arrow3"
        .Line.ForeColor.RGB = RGB(68, 114, 196)
        .Fill.ForeColor.RGB = RGB(68, 114, 196)
    End With
    
    ' Draw Text Box
    Set shp = ws.Shapes.AddShape(msoShapeRoundedRectangle, textLeft, centerPos - 15, 120, 30)
    With shp
        .Name = "ConfigAnnot_Text3"
        .TextFrame.Characters.Text = "Used as a condition explicitly"
        .TextFrame.Characters.Font.Size = 9
        .TextFrame.Characters.Font.Name = "Calibri"
        .TextFrame.Characters.Font.Color = RGB(0, 0, 0)
        .TextFrame.HorizontalAlignment = xlHAlignCenter
        .TextFrame.VerticalAlignment = xlVAlignCenter
        .Fill.ForeColor.RGB = RGB(255, 255, 255)
        .Line.ForeColor.RGB = RGB(0, 0, 0)
        .Line.Weight = 1
    End With

End Sub

 
Sub RunPVTProcessing_Testing()
    On Error GoTo Cleanup
    
    ' --- Animation/Loading Indicator Start ---
    Dim wsActive As Worksheet
    Dim loadingShape As Shape
    Dim centerX As Double
    Dim centerY As Double
    Dim shpExisted As Boolean
    
    Set wsActive = ActiveSheet
    
    ' Clean up any stuck previous shape
    On Error Resume Next
    wsActive.Shapes("PVT_Loading_Indicator").Delete
    On Error GoTo 0
    On Error GoTo Cleanup
    
    ' Calculate center of visible window to place the notification
    ' Use fallback if ActiveWindow is not available or weird
    If ActiveWindow.VisibleRange.Width > 0 Then
        centerX = ActiveWindow.VisibleRange.Left + (ActiveWindow.VisibleRange.Width / 2) - 100
        centerY = ActiveWindow.VisibleRange.Top + (ActiveWindow.VisibleRange.Height / 2) - 30
    Else
        centerX = 200
        centerY = 200
    End If
    
    ' Create "Processing..." shape
    Set loadingShape = wsActive.Shapes.AddShape(msoShapeRoundedRectangle, centerX, centerY, 200, 60)
    With loadingShape
        .Name = "PVT_Loading_Indicator"
        .Fill.ForeColor.RGB = RGB(68, 114, 196) ' Match theme blue
        .Line.Visible = msoFalse
        .Shadow.Type = msoShadow6
        .TextFrame.Characters.Text = "Processing Data..." & vbCrLf & "Please wait..."
        .TextFrame.Characters.Font.Color = vbWhite
        .TextFrame.Characters.Font.Bold = True
        .TextFrame.Characters.Font.Size = 12
        .TextFrame.HorizontalAlignment = xlHAlignCenter
        .TextFrame.VerticalAlignment = xlVAlignCenter
        .Placement = xlFreeFloating ' Don't move with cells
    End With
    
    Application.Cursor = xlWait
    DoEvents ' Force UI to update and show the shape
    ' --- Animation End ---

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    
    ProcessPVTData_Final
    
    ' Generate variance sheet after PVT_STA is created
    CreateVarianceSheet

Cleanup:
    ' --- Animation Cleanup ---
    On Error Resume Next
    Application.ScreenUpdating = True ' Must turn on to delete shape visible? No, but good practice
    wsActive.Shapes("PVT_Loading_Indicator").Delete
    Application.Cursor = xlDefault
    On Error GoTo 0
    ' --- Animation Cleanup End ---

    Application.ScreenUpdating = True
    Application.EnableEvents = True
End Sub

'==========================================
' CREATE VARIANCE SHEET - Shows applicable corners for each instance
'==========================================
Public Sub CreateVarianceSheet()
    On Error GoTo ErrorHandler
    
    Dim wsInstances As Worksheet
    Dim wsPVTSTA As Worksheet
    Dim wsVariance As Worksheet
    Dim instLastRow As Long
    Dim pvtLastRow As Long
    Dim instNameCol As Long
    Dim pvtInstanceListCol As Long
    Dim pvtCornerCol As Long
    Dim i As Long
    Dim j As Long
    Dim instanceName As String
    Dim cornerName As String
    Dim instanceList As String
    Dim applicableCorners As String
    Dim varRow As Long
    Dim instanceDict As Object
    
    Debug.Print "CreateVarianceSheet: Starting"
    
    ' Get references to required sheets
    On Error Resume Next
    Set wsInstances = ThisWorkbook.Sheets(SHEET_INSTANCES)
    Set wsPVTSTA = ThisWorkbook.Sheets(SHEET_OUTPUT)
    On Error GoTo ErrorHandler
    
    If wsInstances Is Nothing Then
        Debug.Print "CreateVarianceSheet: " & SHEET_INSTANCES & " sheet not found"
        Exit Sub
    End If
    
    If wsPVTSTA Is Nothing Then
        Debug.Print "CreateVarianceSheet: " & SHEET_OUTPUT & " sheet not found"
        Exit Sub
    End If
    
    ' Delete existing variance sheet if it exists
    On Error Resume Next
    Application.DisplayAlerts = False
    ThisWorkbook.Sheets(SHEET_VARIANCE).Delete
    Application.DisplayAlerts = True
    On Error GoTo ErrorHandler
    
    ' Create new variance sheet
    Set wsVariance = ThisWorkbook.Sheets.Add(After:=wsPVTSTA)
    wsVariance.Name = SHEET_VARIANCE
    
    ' Set up headers
    wsVariance.Range("A1").Value = "Instance Name"
    wsVariance.Range("B1").Value = "Corners"
    
    ' Format headers
    With wsVariance.Range("A1:B1")
        .Font.Bold = True
        .Font.Size = 11
        .Font.Color = RGB(255, 255, 255)
        .Interior.Color = RGB(68, 114, 196)
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
    
    ' Find instance_name column in SHEET_INSTANCES
    instNameCol = FindColumnInRow(wsInstances, "instance_name", 2)
    If instNameCol = -1 Then
        instNameCol = FindColumnInRow(wsInstances, "instance_name", 1)
    End If
    
    If instNameCol = -1 Then
        Debug.Print "CreateVarianceSheet: instance_name column not found in " & SHEET_INSTANCES
        Exit Sub
    End If
    
    ' Find Instance List column in PVT_STA (should be column 8 based on typical structure)
    pvtInstanceListCol = -1
    Dim pvtHeaderRow As Long
    pvtHeaderRow = 1
    Dim pvtLastColCheck As Long
    pvtLastColCheck = wsPVTSTA.Cells(pvtHeaderRow, wsPVTSTA.Columns.Count).End(xlToLeft).Column
    
    For i = 1 To pvtLastColCheck
        If UCase(Trim(wsPVTSTA.Cells(pvtHeaderRow, i).Value)) = "INSTANCE LIST" Then
            pvtInstanceListCol = i
            Exit For
        End If
    Next i
    
    If pvtInstanceListCol = -1 Then
        Debug.Print "CreateVarianceSheet: Instance List column not found in " & SHEET_OUTPUT
        Exit Sub
    End If
    
    ' Find PVT Name column in PVT_STA (should be column 3 with header "PVTs Name")
    pvtCornerCol = -1
    For i = 1 To 10 ' Check first 10 columns
        Dim hdrVal As String
        hdrVal = UCase(Trim(wsPVTSTA.Cells(pvtHeaderRow, i).Value))
        ' Check for various possible header names
        If hdrVal = "PVTS NAME" Or hdrVal = "PVT NAME" Or hdrVal = "CORNER" Or hdrVal = "PVT" Then
            pvtCornerCol = i
            Debug.Print "CreateVarianceSheet: Found PVT Name column at position " & i & " with header: " & wsPVTSTA.Cells(pvtHeaderRow, i).Value
            Exit For
        End If
    Next i
    
    ' If not found by header, use column 3 (default position for "PVTs Name")
    If pvtCornerCol = -1 Then
        pvtCornerCol = 3
        Debug.Print "CreateVarianceSheet: Using column 3 for PVT names (header not found, using default position)"
    End If
    
    ' Get all unique instances from SHEET_INSTANCES
    Set instanceDict = CreateObject("Scripting.Dictionary")
    instLastRow = wsInstances.Cells(wsInstances.Rows.Count, instNameCol).End(xlUp).Row
    
    For i = 3 To instLastRow ' Assuming data starts at row 3
        instanceName = Trim(CStr(wsInstances.Cells(i, instNameCol).Value))
        If instanceName <> "" And Not instanceDict.Exists(instanceName) Then
            instanceDict.Add instanceName, ""
        End If
    Next i
    
    Debug.Print "CreateVarianceSheet: Found " & instanceDict.Count & " unique instances"
    
    ' Get last row of PVT_STA
    pvtLastRow = wsPVTSTA.Cells(wsPVTSTA.Rows.Count, pvtCornerCol).End(xlUp).Row
    
    ' For each unique instance, find applicable corners
    varRow = 2
    Dim instKey As Variant
    Dim cornersDict As Object
    
    For Each instKey In instanceDict.Keys
        instanceName = CStr(instKey)
        applicableCorners = ""
        Set cornersDict = CreateObject("Scripting.Dictionary")
        
        ' Search through all PVT_STA rows
        For j = 2 To pvtLastRow ' Assuming data starts at row 2
            cornerName = Trim(CStr(wsPVTSTA.Cells(j, pvtCornerCol).Value))
            instanceList = Trim(CStr(wsPVTSTA.Cells(j, pvtInstanceListCol).Value))
            
            ' Check if this instance appears in the instance list
            If cornerName <> "" And instanceList <> "" Then
                ' Split the list and check each item to handle spaces correctly
                Dim instArray() As String
                Dim k As Long
                instArray = Split(instanceList, ",")
                
                For k = LBound(instArray) To UBound(instArray)
                    If UCase(Trim(instArray(k))) = UCase(instanceName) Then
                        ' Match found
                        
                        ' Only add if not already in dictionary (Duplicate Check)
                        If Not cornersDict.Exists(cornerName) Then
                            cornersDict.Add cornerName, True
                            
                            If applicableCorners = "" Then
                                applicableCorners = cornerName
                            Else
                                applicableCorners = applicableCorners & ", " & cornerName
                            End If
                        End If
                        
                        Exit For ' Found in this row, move to next row
                    End If
                Next k
            End If
        Next j
        
        ' Write to variance sheet
        wsVariance.Cells(varRow, 1).Value = instanceName
        wsVariance.Cells(varRow, 2).Value = applicableCorners
        varRow = varRow + 1
    Next instKey
    
    ' Format the variance sheet
    With wsVariance
        .Columns("A").ColumnWidth = 30
        .Columns("B").ColumnWidth = 80
        .Columns("B").WrapText = True
        
        ' Add borders
        Dim lastVarRow As Long
        lastVarRow = varRow - 1
        If lastVarRow > 1 Then
            .Range("A1:B" & lastVarRow).Borders.LineStyle = xlContinuous
            .Range("A1:B" & lastVarRow).Borders.Color = RGB(191, 191, 191)
            
            ' Alternate row colors for readability
            Dim r As Long
            For r = 2 To lastVarRow
                If r Mod 2 = 0 Then
                    .Range("A" & r & ":B" & r).Interior.Color = RGB(242, 242, 242)
                End If
            Next r
        End If
        
        ' Freeze top row
        .Rows(2).Select
        ActiveWindow.FreezePanes = True
        .Range("A1").Select
    End With
    
    Debug.Print "CreateVarianceSheet: Completed successfully. Created " & (varRow - 2) & " rows"
    Exit Sub
    
ErrorHandler:
    Debug.Print "CreateVarianceSheet ERROR: " & Err.Number & " - " & Err.Description
End Sub

Private Function ValidateDRsAgainstInstanceList(drPrefixes As Collection) As String
    Dim wsInst As Worksheet
    Dim drCol As Long
    Dim lastRow As Long
    Dim i As Long
    Dim drVal As String
    Dim drValues As Object
    Dim drp As Variant
    Dim matchFound As Boolean
    Dim missingDRs As String
    Dim drKey As Variant

    Set drValues = CreateObject("Scripting.Dictionary")
    
    On Error Resume Next
    Set wsInst = ThisWorkbook.Sheets(SHEET_INSTANCES)
    On Error GoTo 0
    
    If wsInst Is Nothing Then
        ' If instance sheet doesn't exist, we can't validate.
        Exit Function
    End If

    ' Find dual_rail column
    drCol = FindColumnInRow(wsInst, "dual_rail", 2)
    If drCol = -1 Then
        ' Column not found
        Exit Function
    End If

    lastRow = wsInst.Cells(wsInst.Rows.Count, drCol).End(xlUp).Row
    
    ' Collect unique DR values from Instance List
    For i = 3 To lastRow
        drVal = Trim(UCase(CStr(wsInst.Cells(i, drCol).Value)))
        If drVal <> "" Then
            If Not drValues.Exists(drVal) Then drValues.Add drVal, True
        End If
    Next i

    ' Check each configured prefix
    missingDRs = ""
    For Each drp In drPrefixes
        matchFound = False
        ' Check if this prefix matches ANY of the actual DR values
        For Each drKey In drValues.Keys
            If MatchesWildcard(CStr(drKey), CStr(drp)) Then
                matchFound = True
                Exit For
            End If
        Next drKey
        
        If Not matchFound Then
            If missingDRs = "" Then
                missingDRs = drp
            Else
                missingDRs = missingDRs & ", " & drp
            End If
        End If
    Next drp

    ValidateDRsAgainstInstanceList = missingDRs
End Function

Private Function GetUniqueDRValuesFromInstanceList(ByRef outMsg As String) As String
    Dim wsInst As Worksheet
    Dim wsPVT As Worksheet
    Dim drCol As Long
    Dim lastRow As Long
    Dim i As Long
    Dim drVal As String
    Dim drValues As Object
    Dim finalFilters As Object
    Dim drList As String
    Dim pvtHeaders As Collection
    Dim foundList As String
    foundList = ""
    
    Set drValues = CreateObject("Scripting.Dictionary")
    Set finalFilters = CreateObject("Scripting.Dictionary")
    
    On Error Resume Next
    Set wsInst = ThisWorkbook.Sheets(SHEET_INSTANCES)
    Set wsPVT = ThisWorkbook.Sheets(SHEET_PVT)
    On Error GoTo 0
    
    If wsInst Is Nothing Then
        GetUniqueDRValuesFromInstanceList = ""
        Exit Function
    End If
    
    ' Load PVT headers for validation
    ' Load PVT headers for validation (Row 1 and Row 2)
    Set pvtHeaders = New Collection
    If Not wsPVT Is Nothing Then
        Dim lc As Long, c As Long
        ' Row 1
        lc = wsPVT.Cells(1, wsPVT.Columns.Count).End(xlToLeft).Column
        For c = 1 To lc
            pvtHeaders.Add UCase(Trim(CStr(wsPVT.Cells(1, c).Value)))
        Next c
        ' Row 2
        lc = wsPVT.Cells(2, wsPVT.Columns.Count).End(xlToLeft).Column
        For c = 1 To lc
            pvtHeaders.Add UCase(Trim(CStr(wsPVT.Cells(2, c).Value)))
        Next c
    End If
    
    drCol = FindColumnInRow(wsInst, "dual_rail", 2)
    If drCol = -1 Then
        ' Try Row 1 as fallback
        drCol = FindColumnInRow(wsInst, "dual_rail", 1)
    End If
    
    If drCol = -1 Then
        GetUniqueDRValuesFromInstanceList = ""
        Exit Function
    End If
    
    lastRow = wsInst.Cells(wsInst.Rows.Count, drCol).End(xlUp).Row
    
    ' Step 1: Collect unique values from Instance List
    For i = 3 To lastRow
        drVal = Trim(CStr(wsInst.Cells(i, drCol).Value))
        If drVal <> "" Then
            If Not drValues.Exists(UCase(drVal)) Then
                drValues.Add UCase(drVal), drVal
            End If
        End If
    Next i
    
    ' Step 2: Generate smart filters
    Dim key As Variant
    Dim originalVal As String
    Dim prefix As String
    Dim wildcardPat As String
    Dim patMatches As Boolean
    Dim j As Long
    
    For Each key In drValues.Keys
        originalVal = drValues(key)
        
        ' Extract alpha prefix (e.g., "dr0" -> "dr", "ldr1" -> "ldr")
        prefix = ""
        For j = 1 To Len(originalVal)
            If IsNumeric(Mid(originalVal, j, 1)) Then Exit For
            prefix = prefix & Mid(originalVal, j, 1)
        Next j
        
        ' If prefix is empty or same as original (no numbers), just use original + *
        If prefix = "" Then prefix = originalVal
        
        wildcardPat = prefix & "*"
        
        ' Check if wildcard pattern matches any PVT header
        patMatches = False
        If pvtHeaders.Count > 0 Then
            Dim hdr As Variant
            For Each hdr In pvtHeaders
                If hdr Like UCase(wildcardPat) Then
                    patMatches = True
                    Exit For
                End If
            Next hdr
        Else
            ' If PVT sheet missing/empty, assume wildcard is good
            patMatches = True
        End If
        
        If patMatches Then
            If Not finalFilters.Exists(UCase(wildcardPat)) Then
                finalFilters.Add UCase(wildcardPat), wildcardPat
                If foundList = "" Then
                    foundList = wildcardPat & " (found for " & originalVal & ")"
                Else
                    foundList = foundList & vbCrLf & "  • " & wildcardPat & " (found for " & originalVal & ")"
                End If
            End If
        Else
            ' Fallback to original value if wildcard didn't match anything
            ' (Though if wildcard didn't match, original likely won't either, but safe to keep)
            If Not finalFilters.Exists(UCase(originalVal)) Then
                finalFilters.Add UCase(originalVal), originalVal
                If foundList = "" Then
                    foundList = originalVal & " (exact match)"
                Else
                    foundList = foundList & vbCrLf & "  • " & originalVal & " (exact match)"
                End If
            End If
        End If
    Next key
    
    ' Step 3: Build comma-separated string
    drList = ""
    For Each key In finalFilters.Keys
        If drList = "" Then
            drList = finalFilters(key)
        Else
            drList = drList & "," & finalFilters(key)
        End If
    Next key
    
    outMsg = "Found the following Dual Rail columns in " & SHEET_PVT & " sheet based on instance data:" & vbCrLf & vbCrLf & "  • " & foundList
    GetUniqueDRValuesFromInstanceList = drList
End Function

Private Function GetUniqueVTValuesFromInstanceList(ByRef outMsg As String, ByRef unmatchedMsg As String) As String
    Dim wsInst As Worksheet
    Dim wsPVT As Worksheet
    Dim instNameCol As Long
    Dim lastRow As Long
    Dim i As Long
    Dim instName As String
    Dim vtValues As Object
    Dim finalFilters As Object
    Dim vtList As String
    Dim pvtHeaders As Collection
    
    Set vtValues = CreateObject("Scripting.Dictionary")
    Set finalFilters = CreateObject("Scripting.Dictionary")
    
    On Error Resume Next
    Set wsInst = ThisWorkbook.Sheets(SHEET_INSTANCES)
    Set wsPVT = ThisWorkbook.Sheets(SHEET_PVT)
    On Error GoTo 0
    
    If wsInst Is Nothing Then
        GetUniqueVTValuesFromInstanceList = ""
        Exit Function
    End If
    
    ' Load PVT headers for validation (Row 1 and Row 2)
    Set pvtHeaders = New Collection
    If Not wsPVT Is Nothing Then
        Dim lc As Long, c As Long
        ' Row 1
        lc = wsPVT.Cells(1, wsPVT.Columns.Count).End(xlToLeft).Column
        For c = 1 To lc
            pvtHeaders.Add UCase(Trim(CStr(wsPVT.Cells(1, c).Value)))
        Next c
        ' Row 2
        lc = wsPVT.Cells(2, wsPVT.Columns.Count).End(xlToLeft).Column
        For c = 1 To lc
            pvtHeaders.Add UCase(Trim(CStr(wsPVT.Cells(2, c).Value)))
        Next c
    End If
    
    ' Find instance_name column (Row 1 or 2)
    instNameCol = FindColumnInRow(wsInst, "instance_name", 2)
    If instNameCol = -1 Then instNameCol = FindColumnInRow(wsInst, "instance_name", 1)
    
    If instNameCol = -1 Then
        GetUniqueVTValuesFromInstanceList = ""
        Exit Function
    End If
    
    lastRow = wsInst.Cells(wsInst.Rows.Count, instNameCol).End(xlUp).Row
    
    ' Step 1: Extract potential VT types from instance names
    ' Pattern: _p<vt>_ or _p<vt> (end of string)
    Dim pPos As Long
    Dim vtPart As String
    Dim nextUnderscore As Long
    
    For i = 3 To lastRow
        instName = Trim(CStr(wsInst.Cells(i, instNameCol).Value))
        If instName <> "" Then
            ' Look for "_p" (case insensitive)
            pPos = InStr(1, instName, "_p", vbTextCompare)
            If pPos > 0 Then
                ' Extract everything after _p
                vtPart = Mid(instName, pPos + 2)
                ' If there's another underscore, stop there
                nextUnderscore = InStr(1, vtPart, "_")
                If nextUnderscore > 0 Then
                    vtPart = Left(vtPart, nextUnderscore - 1)
                End If
                
                If vtPart <> "" Then
                    If Not vtValues.Exists(UCase(vtPart)) Then
                        vtValues.Add UCase(vtPart), vtPart
                    End If
                End If
            End If
        End If
    Next i
    
    ' Step 2: Validate against PVT headers
    Dim key As Variant
    Dim candidateVT As String
    Dim wildcardPat As String
    Dim patMatches As Boolean
    Dim foundList As String
    foundList = ""
    unmatchedMsg = ""
    
    For Each key In vtValues.Keys
        candidateVT = vtValues(key)
        wildcardPat = candidateVT & "*"
        
        ' Check if wildcard pattern matches any PVT header
        patMatches = False
        If pvtHeaders.Count > 0 Then
            Dim hdr As Variant
            For Each hdr In pvtHeaders
                If hdr Like UCase(wildcardPat) Then
                    patMatches = True
                    Exit For
                End If
            Next hdr
        Else
            ' If PVT sheet missing, assume valid? Or fail?
            ' User wants to search PVTs sheet. If no PVT sheet, can't verify.
            ' Let's be strict: if no PVT sheet, no auto-VT.
            patMatches = False
        End If
        
        If patMatches Then
            If Not finalFilters.Exists(UCase(wildcardPat)) Then
                finalFilters.Add UCase(wildcardPat), wildcardPat
                If foundList = "" Then
                    foundList = wildcardPat & " (found for " & candidateVT & ")"
                Else
                    foundList = foundList & vbCrLf & "  • " & wildcardPat & " (found for " & candidateVT & ")"
                End If
            End If
        Else
            ' Collect unmatched VT types
            If unmatchedMsg = "" Then
                unmatchedMsg = candidateVT
            Else
                unmatchedMsg = unmatchedMsg & ", " & candidateVT
            End If
        End If
    Next key
    
    ' Step 3: Build comma-separated string
    vtList = ""
    For Each key In finalFilters.Keys
        If vtList = "" Then
            vtList = finalFilters(key)
        Else
            vtList = vtList & "," & finalFilters(key)
        End If
    Next key
    
    outMsg = "Found the following VT columns in " & SHEET_PVT & " sheet based on instance names:" & vbCrLf & vbCrLf & "  • " & foundList
    GetUniqueVTValuesFromInstanceList = vtList
End Function

Private Function ValidateVTsAgainstInstanceList(vtPrefixes As Collection) As String
    Dim wsInst As Worksheet
    Dim instNameCol As Long
    Dim lastRow As Long
    Dim i As Long
    Dim instName As String
    Dim vtFoundDict As Object
    Dim vp As Variant
    Dim missingVTs As String
    
    Set vtFoundDict = CreateObject("Scripting.Dictionary")
    For Each vp In vtPrefixes
        vtFoundDict.Add vp, False
    Next vp
    
    On Error Resume Next
    Set wsInst = ThisWorkbook.Sheets(SHEET_INSTANCES)
    On Error GoTo 0
    
    If wsInst Is Nothing Then
        ValidateVTsAgainstInstanceList = ""
        Exit Function
    End If
    
    ' Find instance_name column (Row 1 or 2)
    instNameCol = FindColumnInRow(wsInst, "instance_name", 2)
    If instNameCol = -1 Then instNameCol = FindColumnInRow(wsInst, "instance_name", 1)
    
    If instNameCol = -1 Then
        ValidateVTsAgainstInstanceList = ""
        Exit Function
    End If
    
    lastRow = wsInst.Cells(wsInst.Rows.Count, instNameCol).End(xlUp).Row
    
    ' Scan instance names for VT patterns
    Dim pPos As Long, vtPart As String, nextUnderscore As Long
    Dim baseVT As String
    
    For i = 3 To lastRow
        instName = Trim(CStr(wsInst.Cells(i, instNameCol).Value))
        If instName <> "" Then
            ' Look for "_p" (case insensitive)
            pPos = InStr(1, instName, "_p", vbTextCompare)
            If pPos > 0 Then
                ' Extract everything after _p
                vtPart = Mid(instName, pPos + 2)
                ' If there's another underscore, stop there
                nextUnderscore = InStr(1, vtPart, "_")
                If nextUnderscore > 0 Then
                    vtPart = Left(vtPart, nextUnderscore - 1)
                End If
                
                If vtPart <> "" Then
                    ' Check if this VT part matches any of our configured prefixes
                    For Each vp In vtPrefixes
                        ' Extract base name from wildcard (e.g., "fvt*" -> "fvt")
                        baseVT = Replace(vp, "*", "")
                        If UCase(vtPart) = UCase(baseVT) Then
                            vtFoundDict(vp) = True
                        End If
                    Next vp
                End If
            End If
        End If
    Next i
    
    ' Collect missing VTs
    missingVTs = ""
    For Each vp In vtPrefixes
        If Not vtFoundDict(vp) Then
            If missingVTs = "" Then missingVTs = vp Else missingVTs = missingVTs & ", " & vp
        End If
    Next vp
    
    ValidateVTsAgainstInstanceList = missingVTs
End Function

Private Function ValidateFocusPatternsAgainstInstanceList(focusPatterns As Variant) As String
    Dim wsInst As Worksheet
    Dim memTypeCol As Long
    Dim lastRow As Long
    Dim i As Long
    Dim memType As String
    Dim pat As Variant
    Dim missingPatterns As String
    Dim patternFound As Boolean
    
    On Error Resume Next
    Set wsInst = ThisWorkbook.Sheets(SHEET_INSTANCES)
    On Error GoTo 0
    
    If wsInst Is Nothing Then
        ValidateFocusPatternsAgainstInstanceList = ""
        Exit Function
    End If
    
    ' Find memory_type column (Row 1 or 2)
    memTypeCol = FindColumnInRow(wsInst, "memory_type", 2)
    If memTypeCol = -1 Then memTypeCol = FindColumnInRow(wsInst, "memory_type", 1)
    
    If memTypeCol = -1 Then
        ValidateFocusPatternsAgainstInstanceList = ""
        Exit Function
    End If
    
    lastRow = wsInst.Cells(wsInst.Rows.Count, memTypeCol).End(xlUp).Row
    
    ' Load all unique memory types into a dictionary for fast lookup
    Dim memTypesDict As Object
    Set memTypesDict = CreateObject("Scripting.Dictionary")
    
    For i = 3 To lastRow
        memType = Trim(CStr(wsInst.Cells(i, memTypeCol).Value))
        If memType <> "" Then
            If Not memTypesDict.Exists(UCase(memType)) Then
                memTypesDict.Add UCase(memType), True
            End If
        End If
    Next i
    
    ' Check each pattern against the memory types
    missingPatterns = ""
    For Each pat In focusPatterns
        patternFound = False
        Dim mt As Variant
        For Each mt In memTypesDict.Keys
            If MatchesWildcard(CStr(mt), CStr(pat)) Then
                patternFound = True
                Exit For
            End If
        Next mt
        
        If Not patternFound Then
            If missingPatterns = "" Then
                missingPatterns = pat
            Else
                missingPatterns = missingPatterns & ", " & pat
            End If
        End If
    Next pat
    
    ValidateFocusPatternsAgainstInstanceList = missingPatterns
End Function


Private Function GetUniqueMemoryTypesFromInstanceList(ByRef outMsg As String, ByRef unmatchedMsg As String) As String
    Dim wsInst As Worksheet
    Dim wsPVT As Worksheet
    Dim memTypeCol As Long
    Dim lastRow As Long
    Dim i As Long
    Dim memType As String
    Dim memTypes As Object
    Dim matchedTypes As Object
    Dim unmatchedTypes As Object
    Dim memTypeList As String
    Dim pvtHeaders As Collection
    
    Set memTypes = CreateObject("Scripting.Dictionary")
    Set matchedTypes = CreateObject("Scripting.Dictionary")
    Set unmatchedTypes = CreateObject("Scripting.Dictionary")
    
    On Error Resume Next
    Set wsInst = ThisWorkbook.Sheets(SHEET_INSTANCES)
    Set wsPVT = ThisWorkbook.Sheets(SHEET_PVT)
    On Error GoTo 0
    
    If wsInst Is Nothing Then
        GetUniqueMemoryTypesFromInstanceList = ""
        Exit Function
    End If
    
    ' Load PVT headers for validation (Row 1 and Row 2)
    Set pvtHeaders = New Collection
    If Not wsPVT Is Nothing Then
        Dim lc As Long, c As Long
        ' Row 1
        lc = wsPVT.Cells(1, wsPVT.Columns.Count).End(xlToLeft).Column
        For c = 1 To lc
            pvtHeaders.Add UCase(Trim(CStr(wsPVT.Cells(1, c).Value)))
        Next c
        ' Row 2
        lc = wsPVT.Cells(2, wsPVT.Columns.Count).End(xlToLeft).Column
        For c = 1 To lc
            pvtHeaders.Add UCase(Trim(CStr(wsPVT.Cells(2, c).Value)))
        Next c
    End If
    
    ' Find memory_type column (Row 1 or 2)
    memTypeCol = FindColumnInRow(wsInst, "memory_type", 2)
    If memTypeCol = -1 Then memTypeCol = FindColumnInRow(wsInst, "memory_type", 1)
    
    If memTypeCol = -1 Then
        GetUniqueMemoryTypesFromInstanceList = ""
        Exit Function
    End If
    
    lastRow = wsInst.Cells(wsInst.Rows.Count, memTypeCol).End(xlUp).Row
    
    ' Step 1: Extract unique memory types from Instance List
    For i = 3 To lastRow
        memType = Trim(CStr(wsInst.Cells(i, memTypeCol).Value))
        If memType <> "" Then
            If Not memTypes.Exists(UCase(memType)) Then
                memTypes.Add UCase(memType), memType
            End If
        End If
    Next i
    
    ' Step 2: Validate against PVT headers (exact match only, no WA tolerance)
    Dim key As Variant
    Dim originalMemType As String
    Dim found As Boolean
    Dim foundList As String
    Dim unmatchedList As String
    foundList = ""
    unmatchedList = ""
    
    For Each key In memTypes.Keys
        originalMemType = memTypes(key)
        found = False
        
        ' Check if memory type exists in PVT headers (exact match only)
        If pvtHeaders.Count > 0 Then
            Dim hdr As Variant
            For Each hdr In pvtHeaders
                ' Only exact match counts
                If UCase(Trim(CStr(hdr))) = UCase(originalMemType) Then
                    found = True
                    Exit For
                End If
            Next hdr
        End If
        
        If found Then
            If Not matchedTypes.Exists(UCase(originalMemType)) Then
                matchedTypes.Add UCase(originalMemType), originalMemType
                If foundList = "" Then
                    foundList = originalMemType
                Else
                    foundList = foundList & vbCrLf & "  • " & originalMemType
                End If
            End If
        Else
            If Not unmatchedTypes.Exists(UCase(originalMemType)) Then
                unmatchedTypes.Add UCase(originalMemType), originalMemType
                If unmatchedList = "" Then
                    unmatchedList = originalMemType
                Else
                    unmatchedList = unmatchedList & vbCrLf & "  • " & originalMemType
                End If
            End If
        End If
    Next key
    
    ' Step 3: Build comma-separated string
    memTypeList = ""
    For Each key In matchedTypes.Keys
        If memTypeList = "" Then
            memTypeList = matchedTypes(key)
        Else
            memTypeList = memTypeList & "," & matchedTypes(key)
        End If
    Next key
    
    outMsg = "Found the following memory types in " & SHEET_PVT & " sheet:" & vbCrLf & vbCrLf & "  • " & foundList
    unmatchedMsg = unmatchedList
    GetUniqueMemoryTypesFromInstanceList = memTypeList
End Function

'==========================================
' CUSTOM CONDITION FILTER HELPER FUNCTIONS
'==========================================

Private Function ParseCustomFilters(ByVal configStr As String) As Collection
    Dim c As Collection
    Set c = New Collection
    
    If Trim(configStr) = "" Then
        Set ParseCustomFilters = c
        Exit Function
    End If
    
    Dim parts() As String
    Dim filterPart As Variant
    Dim subParts() As String
    Dim filterObj As Object
    Dim valueMapping As Object
    Dim mappingPairs() As String
    Dim pairPart As Variant
    Dim kvParts() As String
    
    parts = Split(configStr, ";")
    
    For Each filterPart In parts
        If Trim(CStr(filterPart)) <> "" Then
            subParts = Split(Trim(CStr(filterPart)), ":")
            
            ' Support 2 or 3 parts: PVTColumn:InstanceColumn[:ValueMapping]
            If UBound(subParts) >= 1 Then
                Set filterObj = CreateObject("Scripting.Dictionary")
                filterObj.Add "PVTColumn", Trim(subParts(0))
                filterObj.Add "InstanceColumn", Trim(subParts(1))
                
                ' Parse value mapping if present (3rd part)
                If UBound(subParts) >= 2 Then
                    Set valueMapping = CreateObject("Scripting.Dictionary")
                    mappingPairs = Split(Trim(subParts(2)), ",")
                    
                    For Each pairPart In mappingPairs
                        If InStr(CStr(pairPart), "=") > 0 Then
                            kvParts = Split(Trim(CStr(pairPart)), "=")
                            If UBound(kvParts) >= 1 Then
                                ' Store mapping: PVT Value -> Instance Value
                                valueMapping(UCase(Trim(kvParts(0)))) = Trim(kvParts(1))
                            End If
                        End If
                    Next pairPart
                    
                    filterObj.Add "ValueMapping", valueMapping
                End If
                
                c.Add filterObj
            End If
        End If
    Next filterPart
    
    Set ParseCustomFilters = c
End Function

Private Function LoadAllInstances(wsInst As Worksheet) As Collection
    Dim c As Collection
    Set c = New Collection
    
    If wsInst Is Nothing Then
        Set LoadAllInstances = c
        Exit Function
    End If
    
    Dim lastRow As Long
    Dim lastCol As Long
    Dim i As Long, j As Long
    Dim headerRow As Long
    Dim headers As Object
    Set headers = CreateObject("Scripting.Dictionary")
    
    ' Determine header row (1 or 2)
    If wsInst.Cells(2, 1).Value <> "" Then
        headerRow = 2
    Else
        headerRow = 1
    End If
    
    lastCol = wsInst.Cells(headerRow, wsInst.Columns.Count).End(xlToLeft).Column
    
    ' Map headers
    For j = 1 To lastCol
        headers(UCase(Trim(CStr(wsInst.Cells(headerRow, j).Value)))) = j
    Next j
    
    lastRow = wsInst.Cells(wsInst.Rows.Count, 1).End(xlUp).Row
    
    Dim instObj As Object
    Dim val As String
    
    For i = (headerRow + 1) To lastRow
        Set instObj = CreateObject("Scripting.Dictionary")
        
        ' Load all columns for this instance
        Dim key As Variant
        For Each key In headers.Keys
            val = Trim(CStr(wsInst.Cells(i, headers(key)).Value))
            instObj.Add key, val
        Next key
        
        ' Ensure instance_name and memory_type exist
        If instObj.Exists("INSTANCE_NAME") Then
            c.Add instObj
        End If
    Next i
    
    Set LoadAllInstances = c
End Function

Private Function GetMatchingInstances(pvtRowData As Object, _
                                    allInstances As Collection, _
                                    autoMemCols As Collection, _
                                    autoVTCols As Collection, _
                                    autoDRCols As Collection, _
                                    customFilters As Collection) As String
    Dim matches As String
    matches = ""
    
    Dim inst As Object
    Dim isMatch As Boolean
    Dim autoMatch As Boolean
    Dim customMatch As Boolean
    
    Dim colName As Variant
    Dim pvtVal As String
    Dim instVal As String
    
    ' Check if Auto Memory Filter is active for this row (any "Yes" values?)
    Dim autoFilterActive As Boolean
    autoFilterActive = False
    
    If Not autoMemCols Is Nothing Then
        For Each colName In autoMemCols
            If UCase(pvtRowData(colName)) = "YES" Then
                autoFilterActive = True
                Exit For
            End If
        Next colName
    End If
    
    For Each inst In allInstances
        isMatch = True
        
        ' 1. Auto Memory Filter Check
        If autoFilterActive Then
            autoMatch = False
            If inst.Exists("MEMORY_TYPE") Then
                instVal = UCase(inst("MEMORY_TYPE"))
                ' Check against all PVT columns that are "YES"
                For Each colName In autoMemCols
                    If UCase(pvtRowData(colName)) = "YES" Then
                        ' Check if instance memory type matches this column name
                        ' (Simple wildcard match or exact match depending on logic)
                        ' Existing logic implies exact match or simple mapping
                        ' For now, assume exact match or simple prefix match
                        If instVal = UCase(colName) Or MatchesWildcard(instVal, UCase(colName) & "*") Then
                            autoMatch = True
                            Exit For
                        End If
                    End If
                Next colName
            End If
            If Not autoMatch Then isMatch = False
        End If
        
        ' 2. Auto VT Filter Check
        If isMatch And Not autoVTCols Is Nothing Then
            Dim vtActive As Boolean
            vtActive = False
            Dim vtCol As Variant
            ' Check if any VT filter is active for this row
            For Each vtCol In autoVTCols
                If UCase(pvtRowData(vtCol)) = "YES" Then
                    vtActive = True
                    Exit For
                End If
            Next vtCol
            
            If vtActive Then
                Dim vtMatch As Boolean
                vtMatch = False
                
                ' Extract VT from instance name
                Dim instName As String
                instName = inst("INSTANCE_NAME")
                Dim pPos As Long
                Dim vtPart As String
                Dim nextUnderscore As Long
                
                pPos = InStr(1, instName, "_p", vbTextCompare)
                If pPos > 0 Then
                    vtPart = Mid(instName, pPos + 2)
                    nextUnderscore = InStr(1, vtPart, "_")
                    If nextUnderscore > 0 Then
                        vtPart = Left(vtPart, nextUnderscore - 1)
                    End If
                    
                    ' Check against active VT columns
                    If vtPart <> "" Then
                        For Each vtCol In autoVTCols
                            If UCase(pvtRowData(vtCol)) = "YES" Then
                                ' Check if column name contains VT part (case insensitive)
                                ' e.g. vtPart="fvt", col="fvt_fast" -> Match
                                If InStr(1, vtCol, vtPart, vbTextCompare) > 0 Then
                                    vtMatch = True
                                    Exit For
                                End If
                            End If
                        Next vtCol
                    End If
                End If
                
                If Not vtMatch Then isMatch = False
            End If
        End If
        
        ' 3. Auto DR Filter Check
        If isMatch And Not autoDRCols Is Nothing Then
            Dim drActive As Boolean
            drActive = False
            Dim drCol As Variant
            ' Check if any DR filter is active
            For Each drCol In autoDRCols
                If UCase(pvtRowData(drCol)) = "YES" Then
                    drActive = True
                    Exit For
                End If
            Next drCol
            
            If drActive Then
                Dim drMatch As Boolean
                drMatch = False
                
                If inst.Exists("DUAL_RAIL") Then
                    Dim drVal As String
                    drVal = UCase(inst("DUAL_RAIL"))
                    
                    If drVal <> "" Then
                        For Each drCol In autoDRCols
                            If UCase(pvtRowData(drCol)) = "YES" Then
                                ' Check if column name contains DR value
                                ' e.g. drVal="0.75", col="dr_0.75" -> Match
                                If InStr(1, drCol, drVal, vbTextCompare) > 0 Then
                                    drMatch = True
                                    Exit For
                                End If
                            End If
                        Next drCol
                    End If
                End If
                
                If Not drMatch Then isMatch = False
            End If
        End If
        
        ' 4. Custom Condition Filter Check
        If isMatch And customFilters.Count > 0 Then
            customMatch = True
            Dim cf As Object
            For Each cf In customFilters
                If cf.Exists("SourceColIndex") Then ' Only check if column was found in PVTs
                    pvtVal = pvtRowData(cf("PVTColumn"))
                    
                    ' Only filter if PVT value is not empty
                    If pvtVal <> "" Then
                        If inst.Exists(UCase(cf("InstanceColumn"))) Then
                            instVal = inst(UCase(cf("InstanceColumn")))
                            
                            ' Apply value mapping if it exists
                            Dim mappedVal As String
                            mappedVal = pvtVal ' Default to original value
                            
                            If cf.Exists("ValueMapping") Then
                                Dim valueMap As Object
                                Set valueMap = cf("ValueMapping")
                                
                                ' Check if mapping exists for this PVT value
                                If valueMap.Exists(UCase(pvtVal)) Then
                                    mappedVal = valueMap(UCase(pvtVal))
                                End If
                            End If
                            
                            ' Compare instance value against mapped value
                            If UCase(instVal) <> UCase(mappedVal) Then
                                customMatch = False
                                Exit For
                            End If
                        Else
                            ' Instance column missing - treat as mismatch
                            customMatch = False
                            Exit For
                        End If
                    End If
                End If
            Next cf
            If Not customMatch Then isMatch = False
        End If
        
        If isMatch Then
            If matches = "" Then
                matches = inst("INSTANCE_NAME")
            Else
                matches = matches & ", " & inst("INSTANCE_NAME")
            End If
        End If
    Next inst
    
    GetMatchingInstances = matches
End Function

Private Function GetCustomOnlyMatchingInstances(pvtRowData As Object, _
                                                allInstances As Collection, _
                                                customFilters As Collection) As String
    Dim matches As String
    matches = ""
    
    ' If no custom filters configured, return empty
    If customFilters.Count = 0 Then
        GetCustomOnlyMatchingInstances = ""
        Exit Function
    End If
    
    ' Check if any custom filters are active for this row
    Dim hasActiveFilter As Boolean
    hasActiveFilter = False
    Dim cf As Object
    For Each cf In customFilters
        If cf.Exists("SourceColIndex") Then
            Dim pvtVal As String
            pvtVal = pvtRowData(cf("PVTColumn"))
            If pvtVal <> "" Then
                hasActiveFilter = True
                Exit For
            End If
        End If
    Next cf
    
    ' If no filters are active for this row, return empty
    If Not hasActiveFilter Then
        GetCustomOnlyMatchingInstances = ""
        Exit Function
    End If
    
    ' Proceed with filtering
    Dim inst As Object
    Dim isMatch As Boolean
    Dim customMatch As Boolean
    
    For Each inst In allInstances
        isMatch = True
        
        ' Custom Condition Filter Check Only
        If isMatch And customFilters.Count > 0 Then
            customMatch = True
            For Each cf In customFilters
                If cf.Exists("SourceColIndex") Then ' Only check if column was found in PVTs
                    pvtVal = pvtRowData(cf("PVTColumn"))
                    
                    ' Only filter if PVT value is not empty
                    If pvtVal <> "" Then
                        If inst.Exists(UCase(cf("InstanceColumn"))) Then
                            Dim instVal As String
                            instVal = inst(UCase(cf("InstanceColumn")))
                            
                            ' Apply value mapping if it exists
                            Dim mappedVal As String
                            mappedVal = pvtVal ' Default to original value
                            
                            If cf.Exists("ValueMapping") Then
                                Dim valueMap As Object
                                Set valueMap = cf("ValueMapping")
                                
                                ' Check if mapping exists for this PVT value
                                If valueMap.Exists(UCase(pvtVal)) Then
                                    mappedVal = valueMap(UCase(pvtVal))
                                End If
                            End If
                            
                            ' Compare instance value against mapped value
                            If UCase(instVal) <> UCase(mappedVal) Then
                                customMatch = False
                                Exit For
                            End If
                        Else
                            ' Instance column missing - treat as mismatch
                            customMatch = False
                            Exit For
                        End If
                    End If
                End If
            Next cf
            If Not customMatch Then isMatch = False
        End If
        
        If isMatch Then
            If matches = "" Then
                matches = inst("INSTANCE_NAME")
            Else
                matches = matches & ", " & inst("INSTANCE_NAME")
            End If
        End If
    Next inst
    
    GetCustomOnlyMatchingInstances = matches
End Function

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
    Dim focusWarnings As Boolean
    Dim missingFocusList As String
    
    Dim dataCol As Integer
    Dim colIdx As Integer
    Dim headerCol As Integer
    Dim lastCol As String
    Dim lastColIndex As Long
    Dim col As Long
    Dim outLastCol As Long
    Dim vtCol As Long
    Dim vtRange As Range
    Dim vtPrefix As Variant
    Dim hdrTextVT As String
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
    Dim foundVTCols As Collection
    Set foundVTCols = New Collection
    Dim foundDRCols As Collection
    Set foundDRCols = New Collection
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
    Dim holdStartRow As Integer
    Dim isHOLD As Boolean
    Dim colHeaderName As String
    Dim sourceColIndex As Long
    Dim focusIdx As Long
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

    Dim holdExtractionTypes As Variant
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
    Dim processingWarnings As String
    processingWarnings = ""
    Dim cellVal As String, cellValue As String, instanceListStr As String
    Dim cellValHold As String, instanceListStrHold As String
    Dim memFilterWarnings As String: memFilterWarnings = ""
    Dim vtWarnings As String: vtWarnings = ""
    Dim drWarnings As String: drWarnings = ""
    
    
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
        If Trim(CStr(wsPVT.Range("C16").Value)) <> "" Then
            focusColumnsStr = CStr(wsPVT.Range("C16").Value)
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
        If Trim(CStr(wsPVT.Range("C15").Value)) <> "" Then
            holdOverrideStr = CStr(wsPVT.Range("C15").Value)
        End If
        If holdOverrideStr <> "" Then
            holdOverrideArr = Split(holdOverrideStr, ";")
        Else
            ReDim holdOverrideArr(-1) ' empty
        End If
        
        ' Read Custom Condition Filter
        Dim customFilterStr As String
        customFilterStr = ""
        If Trim(CStr(wsPVT.Range("C19").Value)) <> "" Then
            customFilterStr = CStr(wsPVT.Range("C19").Value)
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
        
        ' Validate patterns against Instance List memory types
        Dim missingFocusPatterns As String
        missingFocusPatterns = ValidateFocusPatternsAgainstInstanceList(focusColumnsArr)
        
        If missingFocusPatterns <> "" Then
            Dim missingPatArr() As String
            Dim mpIdx As Long
            missingPatArr = Split(missingFocusPatterns, ", ")
            
            If memFilterWarnings <> "" Then memFilterWarnings = memFilterWarnings & vbCrLf
            
            For mpIdx = LBound(missingPatArr) To UBound(missingPatArr)
                memFilterWarnings = memFilterWarnings & "  • " & missingPatArr(mpIdx) & " memory type is not part of memory_type in " & SHEET_INSTANCES & " Sheet" & vbCrLf
            Next mpIdx
        End If
        
        
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
                MsgBox "WARNING: No Auto memory Filter columns found matching pattern '" & pattern & "'" & vbCrLf & vbCrLf & _
                       "Pattern: " & pattern & vbCrLf & _
                       "Searched in: " & SHEET_INSTANCES & " (rows 1 and 2)" & vbCrLf & vbCrLf & _
                       "Please verify:" & vbCrLf & _
                       "  • Column names exist in header rows 1 or 2" & vbCrLf & _
                       "  • Pattern spelling and wildcards are correct" & vbCrLf & _
                       "  • Column names match exactly (case-insensitive)", vbExclamation, "Auto memory Filter Not Found"
                ' Record missing focus pattern so final status can show warnings
                focusWarnings = True
                If missingFocusList = "" Then
                    missingFocusList = pattern
                Else
                    missingFocusList = missingFocusList & ", " & pattern
                End If
            End If
        Next colIdx
        
        
        If focusColOrder.Count = 0 Then
            MsgBox "ERROR: No Auto memory Filter columns were found!" & vbCrLf & _
                   "You entered: " & focusColumnsStr & vbCrLf & vbCrLf & _
                   "Please verify the column names in rows 1 or 2 of the " & SHEET_INSTANCES & " match your input." & vbCrLf & _
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

    ' Final output sheet logic removed — PVT_STA is now the authoritative source for Instance List/Count
    ' (Previously created a separate "Final output" sheet and populated it; that behavior has been removed.)
    
    ' Use a single header row for PVT_STA (no extra condition header row)
    
    With wsOutput
        ' Write main headers to single header row (row 1)
        .Range("A1").Value = "Band"
        .Range("B1").Value = "Section"
        .Range("C1").Value = "PVT Name"
        .Range("D1").Value = "Extraction Corner"
        .Range("E1").Value = "VDDP(V)"
        .Range("F1").Value = "VDDA (V)"
        .Range("G1").Value = "RM"
        
        
    
        headerCol = 8 ' Start after RM column (G)
        
        ' Create a new ordered collection for output that matches the reordered header structure
        Set focusColOrderOutput = New Collection
        Dim customFilters As Collection
        Dim customFilterCols As Collection
        Set customFilterCols = New Collection
        
        ' Parse Custom Filters
        Set customFilters = ParseCustomFilters(customFilterStr)
        
        ' Discover Custom Filter Columns in SHEET_PVT Sheet
        Dim cf As Object
        Dim cfMatchedCols As Collection
        Dim cfColIdx As Long
        Dim cfColName As String
        
        For Each cf In customFilters
            cfColName = cf("PVTColumn")
            ' Find column in PVTs sheet (Row 1 or 2)
            Set cfMatchedCols = FindAllColumnIndices(wsSource, 1, cfColName)
            If cfMatchedCols.Count = 0 Then
                Set cfMatchedCols = FindAllColumnIndices(wsSource, 2, cfColName)
            End If
            
            If cfMatchedCols.Count > 0 Then
                cfColIdx = cfMatchedCols(1)
                cf("SourceColIndex") = cfColIdx
                
                ' Add to mapping if not exists
                If Not focusColMapping.Exists(UCase(cfColName)) Then
                    focusColMapping.Add UCase(cfColName), cfColIdx
                    customFilterCols.Add cfColName
                End If
            Else
                ' Warning: Custom filter column not found
                If processingWarnings <> "" Then processingWarnings = processingWarnings & vbCrLf
                processingWarnings = processingWarnings & "WARNING: Custom filter column '" & cfColName & "' not found in " & SHEET_PVT & " sheet."
            End If
        Next cf
        
        If (hasFocusFilter And focusColOrder.Count > 0) Or customFilterCols.Count > 0 Then
            ' Organize columns by condition groups
            Set condition2Cols = New Collection
            Set otherCols = New Collection
            
            ' Classify columns by condition
            For i = 1 To focusColOrder.Count
                colName = UCase(Trim(focusColOrder(i)))
                
                ' Check if it matches any configured DR prefix
                Dim drPrefixesCheck As Collection, drpCheck As Variant
                Dim isDR As Boolean
                isDR = False
                Set drPrefixesCheck = GetSupportedDRs()
                For Each drpCheck In drPrefixesCheck
                    If MatchesWildcard(colName, drpCheck) Then
                        isDR = True
                        Exit For
                    End If
                Next drpCheck
                
                If isDR Then
                    condition2Cols.Add focusColOrder(i)
                Else
                    otherCols.Add focusColOrder(i)
                End If
            Next i
            
            ' Add any VT columns from the source PVTs sheet into output columns (so VT headers exist in PVT_STA)
            Dim vtPrefixes As Collection, vp As Variant
            Set vtPrefixes = GetSupportedVTs()
            Dim srcLastCol As Long, srcColIdx As Long, srcHdr As String
            Dim ocIdx As Long, alreadyInOther As Boolean
            ' Track which configured VT prefixes were actually found in the source headers
            Dim vtFoundDict As Object, vp2 As Variant, missingVTs As String
            Set vtFoundDict = CreateObject("Scripting.Dictionary")
            For Each vp2 In vtPrefixes
                vtFoundDict.Add vp2, False
            Next vp2

            srcLastCol = wsSource.Cells(1, wsSource.Columns.Count).End(xlToLeft).Column
            For srcColIdx = 1 To srcLastCol
                ' Check Row 1
                srcHdr = Trim(CStr(wsSource.Cells(1, srcColIdx).Value))
                Dim matchFound As Boolean
                matchFound = False
                
                If srcHdr <> "" Then
                    For Each vp In vtPrefixes
                        If MatchesWildcard(UCase(srcHdr), vp) Then
                            matchFound = True
                            If vtFoundDict.Exists(vp) Then vtFoundDict(vp) = True
                            srcHdr = UCase(srcHdr)
                            If Not focusColMapping.Exists(srcHdr) Then
                                focusColMapping.Add srcHdr, srcColIdx
                                otherCols.Add srcHdr
                                Debug.Print "ProcessPVTData_Final: Added VT column to outputs (Row 1): " & srcHdr & " (col " & srcColIdx & ")"
                            Else
                                ' Ensure it's present in otherCols if not already
                                alreadyInOther = False
                                For ocIdx = 1 To otherCols.Count
                                    If UCase(otherCols(ocIdx)) = srcHdr Then
                                        alreadyInOther = True
                                        Exit For
                                    End If
                                Next ocIdx
                                If Not alreadyInOther Then otherCols.Add srcHdr
                            End If
                            On Error Resume Next
                            foundVTCols.Add srcHdr, srcHdr
                            On Error GoTo 0
                            Exit For
                        End If
                    Next vp
                End If
                
                ' If not found in Row 1, check Row 2
                If Not matchFound Then
                    srcHdr = Trim(CStr(wsSource.Cells(2, srcColIdx).Value))
                    If srcHdr <> "" Then
                        For Each vp In vtPrefixes
                            If MatchesWildcard(UCase(srcHdr), vp) Then
                                If vtFoundDict.Exists(vp) Then vtFoundDict(vp) = True
                                srcHdr = UCase(srcHdr)
                                If Not focusColMapping.Exists(srcHdr) Then
                                    focusColMapping.Add srcHdr, srcColIdx
                                    otherCols.Add srcHdr
                                    Debug.Print "ProcessPVTData_Final: Added VT column to outputs (Row 2): " & srcHdr & " (col " & srcColIdx & ")"
                                Else
                                    ' Ensure it's present in otherCols if not already
                                    alreadyInOther = False
                                    For ocIdx = 1 To otherCols.Count
                                        If UCase(otherCols(ocIdx)) = srcHdr Then
                                            alreadyInOther = True
                                            Exit For
                                        End If
                                    Next ocIdx
                                    If Not alreadyInOther Then otherCols.Add srcHdr
                                End If
                            On Error Resume Next
                            foundVTCols.Add srcHdr, srcHdr
                            On Error GoTo 0
                            Exit For
                            End If
                        Next vp
                    End If
                End If
            Next srcColIdx

            ' Collect warnings for missing VT prefixes
            missingVTs = ""
            For Each vp2 In vtPrefixes
                If Not vtFoundDict(vp2) Then
                    If missingVTs = "" Then missingVTs = vp2 Else missingVTs = missingVTs & ", " & vp2
                End If
            Next vp2
            
            ' Check against Instance List
            Dim missingInstVTs As String
            missingInstVTs = ValidateVTsAgainstInstanceList(vtPrefixes)
            
            ' Build comprehensive VT warnings
            Dim vtWarningMsg As String
            vtWarningMsg = ""
            
            For Each vp2 In vtPrefixes
                Dim inPVT As Boolean, inInst As Boolean
                inPVT = vtFoundDict(vp2)
                inInst = (InStr(1, ", " & missingInstVTs & ",", ", " & vp2 & ",", vbTextCompare) = 0)
                
                If inInst And Not inPVT Then
                    ' Case 1: Instances exist but no PVT coverage
                    If vtWarningMsg <> "" Then vtWarningMsg = vtWarningMsg & vbCrLf
                    vtWarningMsg = vtWarningMsg & "  • " & vp2 & ": Instances found in '" & SHEET_INSTANCES & "' but NO matching columns in '" & SHEET_PVT & "'."
                ElseIf Not inInst And inPVT Then
                    ' Case 2: PVT columns exist but no instances use them (Optional warning, maybe less critical?)
                    ' User asked for: "check for name like fvt* in PVTs Sheet coloumns if its not found then mention that also"
                    ' So primarily focused on missing PVT columns.
                ElseIf Not inInst And Not inPVT Then
                    ' Case 3: Filter is completely unused
                    If vtWarningMsg <> "" Then vtWarningMsg = vtWarningMsg & vbCrLf
                    vtWarningMsg = vtWarningMsg & "  • " & vp2 & ": No instances found in '" & SHEET_INSTANCES & "' AND no columns found in '" & SHEET_PVT & "'."
                End If
            Next vp2
            
            If vtWarningMsg <> "" Then
                If vtWarnings <> "" Then vtWarnings = vtWarnings & vbCrLf
                vtWarnings = vtWarnings & vtWarningMsg
            End If

            ' Add any DR columns from the source PVTs sheet into output columns (condition2Cols)
            Dim drPrefixes As Collection, drp As Variant
            Set drPrefixes = GetSupportedDRs()
            Dim drFoundDict As Object, drp2 As Variant, missingDRs As String
            Set drFoundDict = CreateObject("Scripting.Dictionary")
            For Each drp2 In drPrefixes
                drFoundDict.Add drp2, False
            Next drp2

            For srcColIdx = 1 To srcLastCol
                ' Check Row 1
                srcHdr = Trim(CStr(wsSource.Cells(1, srcColIdx).Value))
                matchFound = False
                
                If srcHdr <> "" Then
                    For Each drp In drPrefixes
                        If MatchesWildcard(UCase(srcHdr), drp) Then
                            matchFound = True
                            If drFoundDict.Exists(drp) Then drFoundDict(drp) = True
                            srcHdr = UCase(srcHdr)
                            If Not focusColMapping.Exists(srcHdr) Then
                                focusColMapping.Add srcHdr, srcColIdx
                                condition2Cols.Add srcHdr
                                Debug.Print "ProcessPVTData_Final: Added DR column to outputs (Row 1): " & srcHdr & " (col " & srcColIdx & ")"
                            Else
                                ' Ensure it's present in condition2Cols if not already
                                alreadyInOther = False
                                For ocIdx = 1 To condition2Cols.Count
                                    If UCase(condition2Cols(ocIdx)) = srcHdr Then
                                        alreadyInOther = True
                                        Exit For
                                    End If
                                Next ocIdx
                                If Not alreadyInOther Then condition2Cols.Add srcHdr
                            End If
                            On Error Resume Next
                            foundDRCols.Add srcHdr, srcHdr
                            On Error GoTo 0
                            Exit For
                        End If
                    Next drp
                End If
                
                ' If not found in Row 1, check Row 2
                If Not matchFound Then
                    srcHdr = Trim(CStr(wsSource.Cells(2, srcColIdx).Value))
                    If srcHdr <> "" Then
                        For Each drp In drPrefixes
                            If MatchesWildcard(UCase(srcHdr), drp) Then
                                If drFoundDict.Exists(drp) Then drFoundDict(drp) = True
                                srcHdr = UCase(srcHdr)
                                If Not focusColMapping.Exists(srcHdr) Then
                                    focusColMapping.Add srcHdr, srcColIdx
                                    condition2Cols.Add srcHdr
                                    Debug.Print "ProcessPVTData_Final: Added DR column to outputs (Row 2): " & srcHdr & " (col " & srcColIdx & ")"
                                Else
                                    ' Ensure it's present in condition2Cols if not already
                                    alreadyInOther = False
                                    For ocIdx = 1 To condition2Cols.Count
                                        If UCase(condition2Cols(ocIdx)) = srcHdr Then
                                            alreadyInOther = True
                                            Exit For
                                        End If
                                    Next ocIdx
                                    If Not alreadyInOther Then condition2Cols.Add srcHdr
                                End If
                            On Error Resume Next
                            foundDRCols.Add srcHdr, srcHdr
                            On Error GoTo 0
                            Exit For
                            End If
                        Next drp
                    End If
                End If
            Next srcColIdx

            ' Collect warnings for missing DR prefixes
            missingDRs = ""
            For Each drp2 In drPrefixes
                If Not drFoundDict.Exists(drp2) Or drFoundDict(drp2) = False Then
                    If missingDRs = "" Then missingDRs = drp2 Else missingDRs = missingDRs & ", " & drp2
                End If
            Next drp2
            
            ' If missingDRs <> "" Then
            '     If processingWarnings <> "" Then processingWarnings = processingWarnings & vbCrLf & vbCrLf
            '     processingWarnings = processingWarnings & "--- Dual Rail Column Warnings ---" & vbCrLf & _
            '                          "The following DR prefix(es) configured on '" & SHEET_CONFIG & "' (cell C18) were not found in the '" & SHEET_PVT & "' headers: " & missingDRs & "." & vbCrLf & _
            '                          "These columns will not appear in the output."
            ' End If

            ' Warn if any configured DR prefixes are not present in the N3P Instance List 'dual_rail' column
            Dim missingInstDRs As String
            missingInstDRs = ValidateDRsAgainstInstanceList(drPrefixes)
            If missingInstDRs <> "" Then
                If drWarnings <> "" Then drWarnings = drWarnings & vbCrLf
                drWarnings = drWarnings & "  • The following Dual Rail " & missingInstDRs & " configured names are  not found in the '" & SHEET_INSTANCES & "' 'dual_rail' column." & vbCrLf & _
                                     "These entries may not match any instances."
            End If

            ' Build the output order: Condition 2 (DR), then others
            For i = 1 To condition2Cols.Count
                focusColOrderOutput.Add condition2Cols(i)
            Next i
            For i = 1 To otherCols.Count
                focusColOrderOutput.Add otherCols(i)
            Next i
            
            ' Add Custom Filter Columns to Output Order
            For i = 1 To customFilterCols.Count
                focusColOrderOutput.Add customFilterCols(i)
            Next i
            
            ' Write fixed headers first (columns 1-7)
            .Cells(1, 1).Value = "Band"
            .Cells(1, 2).Value = "Section"
            .Cells(1, 3).Value = "PVTs Name"
            .Cells(1, 4).Value = "Extraction Corner"
            .Cells(1, 5).Value = "VDDP(V)"
            .Cells(1, 6).Value = "VDDA (V)"
            .Cells(1, 7).Value = "RM"
            
            ' Track start columns for condition headers
            
            ' Write Condition 2 columns (DR0, DR1)
            If condition2Cols.Count > 0 Then
                For i = 1 To condition2Cols.Count
                    .Cells(1, headerCol).Value = condition2Cols(i)
                    headerCol = headerCol + 1
                Next i
            End If
            
            ' Write other columns
            For i = 1 To otherCols.Count
                .Cells(1, headerCol).Value = otherCols(i)
                headerCol = headerCol + 1
            Next i
            
            ' Write Custom Filter columns
            If customFilterCols.Count > 0 Then
                For i = 1 To customFilterCols.Count
                    .Cells(1, headerCol).Value = customFilterCols(i)
                    headerCol = headerCol + 1
                Next i
            End If
            
' No separate condition header row is needed for PVT_STA; memory-type headers are single-row only
        End If
        
        ' Add Instance List and Instance Count columns
        .Cells(1, headerCol).Value = "Instance List"
        .Cells(1, headerCol + 1).Value = "Instance Count"


    ' Merge each header column vertically (rows 1-2) and format with blue color
    Application.DisplayAlerts = False
    For col = 1 To (headerCol + 1)
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
    wsOutput.Range("A1:" & ColLetter(headerCol + 4) & "1").AutoFilter
    
    Set bandDict = CreateObject("Scripting.Dictionary")
    
    For i = 2 To lastRow
        currentBand = Trim(CStr(wsSource.Cells(i, 2).Value)) ' Column B - BAND
        If currentBand <> "" And Not bandDict.Exists(currentBand) Then
            bandDict.Add currentBand, True
            bandNames.Add currentBand
        End If
    Next i
    
    
    ' Load all instances for filtering
    Dim allInstances As Collection
    Dim wsInstanceList As Worksheet
    On Error Resume Next
    Set wsInstanceList = ThisWorkbook.Sheets(SHEET_INSTANCES)
    On Error GoTo 0
    Set allInstances = LoadAllInstances(wsInstanceList)
    
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
                    
                    ' Build PVT Row Data for filtering
                    Dim pvtRowData As Object
                    Set pvtRowData = CreateObject("Scripting.Dictionary")
                    
                    If hasFocusFilter And focusColOrderOutput.Count > 0 Then
                        For focusIdx = 1 To focusColOrderOutput.Count
                            colHeaderName = focusColOrderOutput(focusIdx)
                            sourceColIndex = CLng(focusColMapping(colHeaderName))
                            
                            cellVal = GetCellValue(wsSource, i, sourceColIndex)
                            .Cells(outputRow, dataCol).Value = cellVal
                            pvtRowData(colHeaderName) = cellVal
                            
                            dataCol = dataCol + 1
                        Next focusIdx
                    End If
                    
                    ' Get Matching Instances (always run, regardless of focus columns)
                    instanceListStr = GetMatchingInstances(pvtRowData, allInstances, otherCols, foundVTCols, foundDRCols, customFilters)
                    
                    ' Write Instance List and Count to PVT_STA
                    .Cells(outputRow, dataCol).Value = instanceListStr
                    
                    ' Compute instance count
                    Dim instArr() As String
                    Dim instCount As Long
                    If instanceListStr <> "" Then
                        instArr = Split(instanceListStr, ",")
                        instCount = UBound(instArr) - LBound(instArr) + 1
                    Else
                        instCount = 0
                    End If
                    .Cells(outputRow, dataCol + 1).Value = instCount
                    
                    ' Previously dumped to Final output sheet; now skipped (PVT_STA holds instance data).
                    
                    ' Get Custom Only Matching Instances
                    Dim customOnlyInstanceListStr As String
                    customOnlyInstanceListStr = GetCustomOnlyMatchingInstances(pvtRowData, allInstances, customFilters)
                    
                    
                    .Range("A" & outputRow & ":" & ColLetter(dataCol + 1) & outputRow).Borders.LineStyle = xlContinuous
                    .Range("A" & outputRow & ":" & ColLetter(dataCol + 1) & outputRow).HorizontalAlignment = xlCenter
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
            
            
            For k = LBound(holdExtractionTypes) To UBound(holdExtractionTypes)
                extractionType = Trim(holdExtractionTypes(k))
                
                
                With wsOutput
                    .Cells(outputRow, 3).Value = pvtName ' PVT Name
                    .Cells(outputRow, 4).Value = extractionType ' Extraction Corner
                    .Cells(outputRow, 5).Value = vddp ' VDDP(V)
                    .Cells(outputRow, 6).Value = vdda ' VDDA (V)
                    .Cells(outputRow, 7).Value = fastestRMValue ' Fastest RM as
                    
                    
                    
                    dataCol = 8 ' Start after RM column
                    
                    ' Build PVT Row Data for filtering
                    Dim pvtRowDataHold As Object
                    Set pvtRowDataHold = CreateObject("Scripting.Dictionary")
                    
                    If hasFocusFilter And focusColOrderOutput.Count > 0 Then
                        For focusIdx = 1 To focusColOrderOutput.Count
                            colHeaderName = focusColOrderOutput(focusIdx)
                            sourceColIndex = CLng(focusColMapping(colHeaderName))
                            
                            cellValHold = GetCellValue(wsSource, i, sourceColIndex)
                            .Cells(outputRow, dataCol).Value = cellValHold
                            pvtRowDataHold(colHeaderName) = cellValHold
                            
                            dataCol = dataCol + 1
                        Next focusIdx
                    End If
                    
                    ' Get Matching Instances (always run, regardless of focus columns)
                    instanceListStrHold = GetMatchingInstances(pvtRowDataHold, allInstances, otherCols, foundVTCols, foundDRCols, customFilters)
                    
                    ' Write Instance List and Count to PVT_STA
                    .Cells(outputRow, dataCol).Value = instanceListStrHold
                    
                    ' Compute instance count
                    Dim instArrH() As String
                    Dim instCountH As Long
                    If instanceListStrHold <> "" Then
                        instArrH = Split(instanceListStrHold, ",")
                        instCountH = UBound(instArrH) - LBound(instArrH) + 1
                    Else
                        instCountH = 0
                    End If
                    .Cells(outputRow, dataCol + 1).Value = instCountH
                    
                    ' Previously dumped to Final output sheet for HOLD rows; now skipped (PVT_STA holds instance data).
                    
                    ' Get Custom Only Matching Instances
                    Dim customOnlyInstanceListStrHold As String
                    customOnlyInstanceListStrHold = GetCustomOnlyMatchingInstances(pvtRowDataHold, allInstances, customFilters)
                    
                    
                    .Range("A" & outputRow & ":" & ColLetter(dataCol + 1) & outputRow).Borders.LineStyle = xlContinuous
                    .Range("A" & outputRow & ":" & ColLetter(dataCol + 1) & outputRow).HorizontalAlignment = xlCenter
                End With
                
                counter = counter + 1
                outputRow = outputRow + 1
            Next k
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
            ' Use focusColOrderOutput.Count to include any VT/DR columns added during processing
            lastColIndex = 7 + focusColOrderOutput.Count + 2 ' A-G + focus cols + Instance List + Instance Count
        Else
            lastColIndex = 9 ' A-G + Instance List + Instance Count (no focus columns)
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
            .Columns(ColLetter(col)).ColumnWidth = 10
        Next col

        
        If hasFocusFilter And focusColOrder.Count > 0 Then
            focusStartCol = "H" ' Column H is first focus column
            ' Use focusColOrderOutput.Count to include any VT/DR columns added during processing
            focusEndCol = ColLetter(7 + focusColOrderOutput.Count)
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

        ' Apply VT-specific conditional formatting (Yes=green, No=red) for any VT columns present in the output
        If outputRow > 2 Then
            outLastCol = .Cells(1, .Columns.Count).End(xlToLeft).Column
            For vtCol = 8 To outLastCol
                hdrTextVT = UCase(Trim(CStr(.Cells(1, vtCol).Value)))
                If hdrTextVT <> "" Then
                    For Each vtPrefix In GetSupportedVTs()
                        If MatchesWildcard(hdrTextVT, vtPrefix) Then
                            On Error Resume Next
                            Set vtRange = .Range(.Cells(2, vtCol), .Cells(outputRow - 1, vtCol))
                            vtRange.FormatConditions.Delete
                            With vtRange.FormatConditions.Add(Type:=xlCellValue, Operator:=xlEqual, Formula1:="=""Yes""")
                                .Interior.Color = RGB(198, 239, 206) ' Light green
                                .Font.Color = RGB(0, 97, 0) ' Dark green text
                                .Font.Bold = True
                            End With
                            With vtRange.FormatConditions.Add(Type:=xlCellValue, Operator:=xlEqual, Formula1:="=""No""")
                                .Interior.Color = RGB(255, 199, 206) ' Light red
                                .Font.Color = RGB(156, 0, 6) ' Dark red text
                            End With
                            Debug.Print "ProcessPVTData_Final: Applied VT formatting to " & hdrTextVT & " at col " & vtCol
                            On Error GoTo ErrorHandler
                            Exit For
                        End If
                    Next vtPrefix
                End If
            Next vtCol
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
            If focusWarnings Then
                focusStatus = "Warnings - Missing: " & missingFocusList
            Else
                focusStatus = "Valid"
            End If
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

    ' Build memory-type matching summary (compare " & SHEET_INSTANCES & " vs " & SHEET_PVT & ")
    Dim instMemTypesDict As Object
    Set instMemTypesDict = CreateObject("Scripting.Dictionary")

    ' --- Refined Memory Type Coverage Summary (Driven by Instance List) ---
    Dim wsInst As Worksheet
    On Error Resume Next
    Set wsInst = ThisWorkbook.Sheets(SHEET_INSTANCES)
    On Error GoTo 0
    
    If Not wsInst Is Nothing Then
        Dim instMemCol As Long, instLastRow As Long
        ' Find memory_type column (check Row 2 then Row 1)
        instMemCol = FindColumnInRow(wsInst, "memory_type", 2)
        If instMemCol = -1 Then instMemCol = FindColumnInRow(wsInst, "memory_type", 1)
        
        If instMemCol <> -1 Then
            ' Load unique memory types from Instance List
            instLastRow = wsInst.Cells(wsInst.Rows.Count, instMemCol).End(xlUp).Row
            Dim imr As Long, imt As String
            For imr = 3 To instLastRow
                imt = Trim(CStr(wsInst.Cells(imr, instMemCol).Value))
                If imt <> "" Then
                    If Not instMemTypesDict.Exists(UCase(imt)) Then instMemTypesDict.Add UCase(imt), imt
                End If
            Next imr
            
            ' Load all PVT headers for fast lookup
            Dim pvtHeadersDict As Object
            Set pvtHeadersDict = CreateObject("Scripting.Dictionary")
            Dim lastHdrCol As Long, hcol As Long, hdrTxt As String
            lastHdrCol = wsSource.Cells(1, wsSource.Columns.Count).End(xlToLeft).Column
            For hcol = 8 To lastHdrCol
                hdrTxt = Trim(CStr(wsSource.Cells(1, hcol).Value))
                If hdrTxt <> "" Then
                    ' Store the base memory type
                    Dim pvtMT As String
                    If ParseMemTypeHeader(hdrTxt, pvtMT) Then
                        If pvtMT <> "" Then
                            If Not pvtHeadersDict.Exists(UCase(pvtMT)) Then pvtHeadersDict.Add UCase(pvtMT), True
                        End If
                    Else
                        ' Fallback for simple headers
                        If Not pvtHeadersDict.Exists(UCase(hdrTxt)) Then pvtHeadersDict.Add UCase(hdrTxt), True
                    End If
                End If
            Next hcol
            
            ' Compare and build warning message
            Dim memKey As Variant, unmatchedList As String
            unmatchedList = ""
            For Each memKey In instMemTypesDict.Keys
                If Not pvtHeadersDict.Exists(memKey) Then
                    If unmatchedList <> "" Then unmatchedList = unmatchedList & vbCrLf
                    unmatchedList = unmatchedList & "  • " & instMemTypesDict(memKey) & " memory_type is not matching in " & SHEET_PVT & " Sheet"
                End If
            Next memKey
            
            ' Add to memory filter warnings
            If unmatchedList <> "" Then
                If memFilterWarnings <> "" Then memFilterWarnings = memFilterWarnings & vbCrLf
                memFilterWarnings = memFilterWarnings & unmatchedList
            End If
        End If
    End If

    ' --- Assemble Consolidated Warnings ---
    processingWarnings = ""
    
    If memFilterWarnings <> "" Then
        processingWarnings = processingWarnings & "--- Auto memory Filter Warnings ---" & vbCrLf & memFilterWarnings
    End If
    
    If vtWarnings <> "" Then
        If processingWarnings <> "" Then processingWarnings = processingWarnings & vbCrLf & vbCrLf
        processingWarnings = processingWarnings & "--- VT Configuration Warnings ---" & vbCrLf & vtWarnings
    End If
    
    If drWarnings <> "" Then
        If processingWarnings <> "" Then processingWarnings = processingWarnings & vbCrLf & vbCrLf
        processingWarnings = processingWarnings & "--- Dual Rail Configuration Warnings ---" & vbCrLf & drWarnings
    End If

    ' --- Final Success Message ---
    successMsg = "Processing Complete." & vbCrLf & vbCrLf & _
                 "Total rows generated: " & (outputRow - 2) & vbCrLf & vbCrLf
                 
    If processingWarnings <> "" Then
        successMsg = successMsg & "WARNINGS:" & vbCrLf & processingWarnings & vbCrLf & vbCrLf
    End If

    successMsg = successMsg & "Refer to the '" & SHEET_OUTPUT & "' sheet for results."

    MsgBox successMsg, vbInformation, "Processing Complete"

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
    
    Dim wsOutput As Worksheet
    Dim cellVal As String
    
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
    Dim bands As Collection
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
    
    ' Find DR columns in PVTs sheet dynamically
    Dim drCols As Object, lastCol As Long, col As Long
    Dim h As String
    Dim drKeyName As Variant, drTypeHeader As Variant
    Set drCols = CreateObject("Scripting.Dictionary")
    lastCol = wsPVT.Cells(2, wsPVT.Columns.Count).End(xlToLeft).Column
    
    Dim supportedDRs As Collection
    Set supportedDRs = GetSupportedDRs()
    
    For col = 1 To lastCol
        h = UCase(Trim(CStr(wsPVT.Cells(2, col).Value)))
        If h <> "" Then
            ' Check if this header matches any supported DR prefix
            Dim sdr As Variant
            For Each sdr In supportedDRs
                If MatchesWildcard(h, CStr(sdr)) Then
                    If Not drCols.Exists(h) Then
                        drCols.Add h, col
                    End If
                    Exit For
                End If
            Next sdr
        End If
    Next col
    
    ' Parse memory conditions from Row 1 of PVTs
    Dim memoryConditions As Collection
    Dim headerText As String, memType As String
    Dim condDict As Object
    Set memoryConditions = New Collection
    
    For col = 10 To lastCol
        headerText = Trim(CStr(wsPVT.Cells(1, col).Value))
        If headerText <> "" Then
            If ParseMemTypeHeader(headerText, memType) Then
                Set condDict = CreateObject("Scripting.Dictionary")
                condDict("column") = col
                condDict("memory_type") = memType
                memoryConditions.Add condDict
            End If
        End If
    Next col
    
    instLastCol = wsInstances.Cells(2, wsInstances.Columns.Count).End(xlToLeft).Column
    For instCol = 1 To instLastCol
        hdr = LCase(Trim(CStr(wsInstances.Cells(2, instCol).Value)))
        If hdr = "memory_type" Then instMemTypeCol = instCol
        If hdr = "instance_name" Then instNameCol = instCol
        If hdr = "dual_rail" Then instDRCol = instCol
        If hdr = "part_name" Then instPartNameCol = instCol
    Next instCol
    
    If instMemTypeCol = 0 Or instDRCol = 0 Then
        MsgBox "Required columns not found in " & SHEET_INSTANCES & " sheet", vbCritical
        GoTo Cleanup
    End If
    
    ' Verify memory types in PVTs exist in Instance List - warn if mismatches
    ' Determine instance last row once and reuse
    Dim instDataLastRow As Long
    instDataLastRow = wsInstances.Cells(wsInstances.Rows.Count, instMemTypeCol).End(xlUp).Row
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
    Dim condIdx2 As Long
    Dim condMT As String
    missingList = ""
    For condIdx2 = 1 To memoryConditions.Count
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
        
        ' Dynamic DR headers
        Dim drOutCol As Long
        drOutCol = 6
        If drCols.Count > 0 Then
            For Each drTypeHeader In drCols.Keys
                .Cells(1, drOutCol).Value = drTypeHeader
                drOutCol = drOutCol + 1
            Next drTypeHeader
        Else
            ' Fallback if no DRs found in PVTs
            .Cells(1, 6).Value = "DR (None found)"
            drOutCol = 7
        End If
        
        Dim headerRange As Range
        Set headerRange = .Range(.Cells(1, 1), .Cells(1, drOutCol - 1))
        headerRange.Font.Bold = True
        headerRange.Font.Size = 12
        headerRange.Interior.Color = RGB(68, 114, 196)
        headerRange.Font.Color = RGB(255, 255, 255)
        headerRange.HorizontalAlignment = xlCenter
        headerRange.VerticalAlignment = xlCenter
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
        
        ' For other bands, use dynamic logic from PVTs sheet
        ' Get DR conditions for this band
        Dim allowedDRs As Object
        Set allowedDRs = CreateObject("Scripting.Dictionary")
        
        For Each drKeyName In drCols.Keys
            cellVal = UCase(Trim(CStr(wsPVT.Cells(bandRow, CLng(drCols(drKeyName))).Value)))
            If cellVal = "YES" Or cellVal = "Y" Then
                allowedDRs.Add LCase(CStr(drKeyName)), True
            End If
        Next drKeyName
        
        ' Get valid memory types for this band
        Dim validMemTypes As Collection
        Dim cond As Object
        Set validMemTypes = New Collection
        
        Dim condIdx As Long
        For condIdx = 1 To memoryConditions.Count
            Set cond = memoryConditions(condIdx)
            cellVal = UCase(Trim(CStr(wsPVT.Cells(bandRow, CLng(cond("column"))).Value)))
            
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
        Dim instMemType As String, instName As String, instPartName As String
        Dim instDR As String
        For instRow = 3 To instDataLastRow
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
            
            instDR = LCase(Trim(CStr(wsInstances.Cells(instRow, instDRCol).Value)))
            
            ' Check if DR is allowed
            If Not allowedDRs.Exists(instDR) Then GoTo NextInstance
            
            For vIdx = 1 To validMemTypes.Count
                Set vCond = validMemTypes(vIdx)
                vMemType = vCond("memory_type")
                
                If LCase(Trim(instMemType)) = LCase(Trim(vMemType)) Then
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
                        
                        ' Dynamic DR display
                        Dim dOutCol As Long
                        dOutCol = 6
                        For Each drKeyName In drCols.Keys
                            If instDR = LCase(CStr(drKeyName)) Then
                                .Cells(outputRow, dOutCol).Value = "Yes"
                            Else
                                .Cells(outputRow, dOutCol).Value = "No"
                            End If
                            .Cells(outputRow, dOutCol).Interior.Color = bandColor
                            dOutCol = dOutCol + 1
                        Next drKeyName
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
    Dim kBand As Variant
    For Each kBand In bandMatchCounts.Keys
        summaryMsg = summaryMsg & "  • " & kBand & ": " & bandMatchCounts(kBand) & " instances" & vbCrLf
    Next kBand
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
    Dim drCols As Object
    Dim memTypeColumns As Collection
    Dim vtCols As Collection
    Dim vtInfo As Object
    Dim hasDRFocus As Boolean
    Dim focusIdx As Long
    Dim focusName As String
    Dim lastRow As Long, pvtRow As Long
    Dim drVal As String
    Dim allInstances As Collection
    Dim i As Long
    Dim colInfo As Object
    Dim cellVal As String, cellValue As String
    Dim memType As String
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
    Dim focusedMemTypes As Collection
    Dim instDRValue As String
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
    Dim supportedVTs As Collection
    Dim vtCandidate As Variant
    Dim ofci As Object
    Dim extraFocusStr As String
    Dim tmpVal As String
    Dim drPass As Boolean
    Dim instDRVal As String
    
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
    
    ' Find DR columns dynamically in PVT_STA
    Set drCols = CreateObject("Scripting.Dictionary")
    Dim supportedDRsList As Collection
    Set supportedDRsList = GetSupportedDRs()
    
    Dim cIdx As Long, headerVal As String
    lastHeaderCol = wsPVT.Cells(1, wsPVT.Columns.Count).End(xlToLeft).Column
    For cIdx = 1 To lastHeaderCol
        headerVal = UCase(Trim(CStr(wsPVT.Cells(1, cIdx).Value)))
        If headerVal = "" Then headerVal = UCase(Trim(CStr(wsPVT.Cells(2, cIdx).Value)))
        
        If headerVal <> "" Then
            Dim sdrEntry As Variant
            For Each sdrEntry In supportedDRsList
                If MatchesWildcard(headerVal, CStr(sdrEntry)) Then
                    If Not drCols.Exists(headerVal) Then
                        drCols.Add headerVal, cIdx
                    End If
                    Exit For
                End If
            Next sdrEntry
        End If
    Next cIdx
    Debug.Print "PopulateInstanceListColumn: Found " & drCols.Count & " DR columns"
    
    ' Extract unique memory types from Instance List for validation
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

    ' Parse memory type columns from header (driven by Instance List valid types)
    Set memTypeColumns = ParseMemoryTypeColumns(wsPVT, instMemTypesDict)
    
    
    Debug.Print "PopulateInstanceListColumn: Found " & memTypeColumns.Count & " valid memory type columns"

    ' Find VT columns (configured types) in header row (supports names like LVT_1 etc.)
    Set vtCols = New Collection
    Set supportedVTs = GetSupportedVTs()
    lastHeaderCol = wsPVT.Cells(1, wsPVT.Columns.Count).End(xlToLeft).Column
    For foundCol = 8 To lastHeaderCol
        headerText = Trim(CStr(wsPVT.Cells(1, foundCol).Value))
        If headerText <> "" Then
            vtType = ""
            For Each vtCandidate In supportedVTs
                If MatchesWildcard(UCase(headerText), vtCandidate) Then
                    ' Normalize vt_type to exclude any wildcard characters (used for instance name matching)
                    vtType = LCase(Replace(vtCandidate, "*", ""))
                    Exit For
                End If
            Next vtCandidate

            If vtType <> "" Then
                Set vtInfo = CreateObject("Scripting.Dictionary")
                vtInfo("column") = foundCol
                vtInfo("vt_type") = vtType
                vtCols.Add vtInfo
            End If
        End If
    Next foundCol

    Debug.Print "PopulateInstanceListColumn: Found " & vtCols.Count & " VT columns"

    ' Debugging support: Count of Instances column removed (now using Instance Count in PVT_STA)
    ' No debug column will be created or formatted.
    
    ' Check if we have focus on DR columns
    hasDRFocus = False
    Dim drName As Variant
    For focusIdx = 1 To focusColOrder.Count
        focusName = UCase(Trim(focusColOrder(focusIdx)))
        For Each drName In drCols.Keys
            If focusName = UCase(CStr(drName)) Then
                hasDRFocus = True
                Exit For
            End If
        Next drName
        If hasDRFocus Then Exit For
    Next focusIdx
    
    ' Collect focused memory types (exclude DR columns)
    Set focusedMemTypes = New Collection
    For focusIdx = 1 To focusColOrder.Count
        focusName = Trim(focusColOrder(focusIdx))
        Dim isDRFocus As Boolean
        isDRFocus = False
        For Each drName In drCols.Keys
            If UCase(focusName) = UCase(CStr(drName)) Then
                isDRFocus = True
                Exit For
            End If
        Next drName
        
        If Not isDRFocus Then
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

    ' Capture extra focus columns (present in focus list but not DR/memory-type/VT) for diagnostics
    Dim otherFocusCols As Collection
    Set otherFocusCols = New Collection
    Dim colIdxInPVT As Long, hdrVal As String
    For focusIdx = 1 To focusColOrder.Count
        focusName = focusColOrder(focusIdx)
        isDRFocus = False
        For Each drName In drCols.Keys
            If UCase(focusName) = UCase(CStr(drName)) Then
                isDRFocus = True
                Exit For
            End If
        Next drName
        
        If isDRFocus Then
            ' skip dual rail focus
        Else
            ' Find corresponding header in PVT_STA (row 1 or 2)
            colIdxInPVT = FindColumnInRow(wsPVT, focusName, 1)
            If colIdxInPVT = -1 Then colIdxInPVT = FindColumnInRow(wsPVT, focusName, 2)
            If colIdxInPVT <> -1 Then
                hdrVal = Trim(CStr(wsPVT.Cells(1, colIdxInPVT).Value))
                ' Skip memory-type columns (they contain '(')
                If InStr(hdrVal, "(") > 0 Then
                    ' memory-type column - skip
                Else
                    ' Determine if it's a VT column by checking configured VT prefixes
                    Dim isVTCol As Boolean
                    isVTCol = False
                    For Each vtCandidate In supportedVTs
                        If MatchesWildcard(UCase(hdrVal), vtCandidate) Then
                            isVTCol = True
                            Exit For
                        End If
                    Next vtCandidate
                    If isVTCol Then
                        ' VT column - skip
                    Else
                        ' Treat as extra focus column
                        Dim ofc As Object
                        Set ofc = CreateObject("Scripting.Dictionary")
                        ofc("name") = focusName
                        ofc("column") = colIdxInPVT
                        otherFocusCols.Add ofc
                    End If
                End If
            End If
        End If
    Next focusIdx
    Debug.Print "PopulateInstanceListColumn: Extra focus columns captured: " & otherFocusCols.Count

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
                
                ' Capture YES/NO for all discovered DR columns
                Dim drColKey As Variant
                For Each drColKey In drCols.Keys
                    Dim drColIdx As Long: drColIdx = drCols(drColKey)
                    drVal = Trim(UCase(CStr(wsPVT.Cells(checkRow, drColIdx).Value)))
                    If drVal = "YES" Or drVal = "Y" Then
                        bandDRConditions(bandName)(CStr(drColKey)) = True
                    End If
                Next drColKey
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
        
        ' If Instance List already populated by previous processing, skip to avoid overwriting
        If Trim(CStr(wsPVT.Cells(pvtRow, instanceListCol).Value)) <> "" Then
            Debug.Print "PopulateInstanceListColumn: Skipping row " & pvtRow & " because Instance List already populated"
            GoTo NextRowPop
        End If
        
        ' Determine the band name for this row (band column A may be merged; use nearest non-empty above)
        Dim curBandName As String
        curBandName = Trim(CStr(wsPVT.Cells(pvtRow, 1).Value))
        If curBandName = "" Then curBandName = Trim(CStr(wsPVT.Cells(pvtRow, 1).End(xlUp).Value))
        Debug.Print "  Band=" & curBandName
        
        ' Get DR status for this row
        Dim activeDRsInRow As Object
        Set activeDRsInRow = CreateObject("Scripting.Dictionary")
        For Each drColKey In drCols.Keys
            drVal = Trim(UCase(CStr(wsPVT.Cells(pvtRow, CLng(drCols(drColKey))).Value)))
            If drVal = "YES" Or drVal = "Y" Then
                activeDRsInRow.Add LCase(CStr(drColKey)), True
            End If
        Next drColKey
        
        ' Capture extra focus column values for this row (diagnostic only)
        extraFocusStr = ""
        If otherFocusCols.Count > 0 Then
            For Each ofci In otherFocusCols
                tmpVal = Trim(CStr(wsPVT.Cells(pvtRow, CLng(ofci("column"))).Value))
                If tmpVal <> "" Then
                    If extraFocusStr = "" Then
                        extraFocusStr = ofci("name") & "=" & tmpVal
                    Else
                        extraFocusStr = extraFocusStr & ", " & ofci("name") & "=" & tmpVal
                    End If
                End If
            Next ofci
        End If
        
        ' Collect matching instances
        Set allInstances = New Collection
        ' Use DR focus logic only when DR focus is enabled AND at least one DR is YES for this row
        If hasDRFocus And activeDRsInRow.Count > 0 Then
            ' DR focus: match instances based on any DR value that is YES for THIS row
            Debug.Print "  Using DR focus logic for row (Active DRs=" & activeDRsInRow.Count & ")"
                For instIdx = 1 To instanceData.Count
                    Set instInfo = instanceData(instIdx)
                    instDRValue = LCase(Trim(CStr(instInfo("dual_rail"))))
                    drMatch = False
                    
                    ' Check if instance DR matches any DR that is YES for this row
                    If activeDRsInRow.Exists(instDRValue) Then
                        drMatch = True
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
                    
                    ' Match check completed
                    
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
                                cellVal = UCase(Trim(CStr(wsPVT.Cells(pvtRow, CLng(colInfo("column"))).Value)))
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
                                    cellVal = UCase(Trim(CStr(wsPVT.Cells(pvtRow, CLng(vtInfo("column"))).Value)))
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
                    
                    ' Match check completed
                    
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
        Else
            ' Memory type focus: match instances based on memory types that are YES in this row
            ' This includes: (1) when DR focus is not enabled, or (2) when DR focus is enabled but neither DR is YES
            Debug.Print "  Using memory type focus logic for row (hasDRFocus=" & hasDRFocus & ", ActiveDRs=" & activeDRsInRow.Count & ")"
            
            If memTypeColumns.Count > 0 Then
                For instIdx = 1 To instanceData.Count
                    Set instInfo = instanceData(instIdx)
                    
                    ' DR Check: apply DR filtering if at least one DR column is enabled on this row
                    
                    ' If at least one DR column is enabled, enforce DR matching
                    If activeDRsInRow.Count > 0 Then
                        drPass = False
                        instDRVal = LCase(Trim(CStr(instInfo("dual_rail"))))
                        
                        If activeDRsInRow.Exists(instDRVal) Then
                            drPass = True
                        End If
                    Else
                        ' No DR columns enabled for this row - skip DR filtering
                        drPass = True
                    End If
                    
                    If Not drPass Then
                        GoTo NextInstInMemFocus
                    End If

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
                            cellVal = UCase(Trim(CStr(wsPVT.Cells(pvtRow, CLng(colInfo("column"))).Value)))
                            If cellVal = "YES" Or cellVal = "Y" Then
                                memMatch = True
                                Exit For
                            End If
                        End If
NextCol_MemMatch:
                    Next
                    
                    ' WA check removed
                    
                    If memMatch Then
                        ' VT check for memory-type-focused rows: if any VT columns are enabled (Yes), require instance name to contain p<vt>
                        vtMatched = False
                        vtHasYes = False
                        If vtCols.Count > 0 Then
                            For Each vtInfo In vtCols
                                cellVal = UCase(Trim(CStr(wsPVT.Cells(pvtRow, CLng(vtInfo("column"))).Value)))
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
                                Debug.Print "  Added instance: " & instName & " (MemType=" & instInfo("memory_type") & ")"
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
        
        
        Debug.Print "  Final instance list for row " & pvtRow & ": " & instanceListStr
        wsPVT.Cells(pvtRow, instanceListCol).Value = instanceListStr
        ' No separate debug Count of Instances column — Instance Count in PVT_STA is authoritative.
        
NextRowPop:
    Next pvtRow
    
    ' No final Count of Instances column to adjust (removed).

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
    
    Dim memCol As Long, instCol As Long, drCol As Long
    
    ' Try finding headers in Row 2 first (default)
    memCol = FindColumnInRow(ws, "memory_type", 2)
    instCol = FindColumnInRow(ws, "instance_name", 2)
    drCol = FindColumnInRow(ws, "dual_rail", 2)
    
    ' If not found, try Row 1
    If memCol = -1 Then memCol = FindColumnInRow(ws, "memory_type", 1)
    If instCol = -1 Then instCol = FindColumnInRow(ws, "instance_name", 1)
    If drCol = -1 Then drCol = FindColumnInRow(ws, "dual_rail", 1)
    
    If memCol = -1 Or instCol = -1 Then
        Set LoadInstanceDataForMatching = data
        Exit Function
    End If
    
    Dim i As Long
    For i = 3 To lastRow
        Dim memType As String, instName As String, dr As String
        memType = Trim(CStr(ws.Cells(i, memCol).Value))
        instName = Trim(CStr(ws.Cells(i, instCol).Value))
        dr = Trim(CStr(ws.Cells(i, drCol).Value))
        
        If memType <> "" And instName <> "" Then
            Dim instanceInfo As Object
            Set instanceInfo = CreateObject("Scripting.Dictionary")
            instanceInfo("memory_type") = memType
            instanceInfo("instance_name") = instName
            instanceInfo("dual_rail") = dr
            data.Add instanceInfo
        End If
    Next i
    
    Set LoadInstanceDataForMatching = data
End Function

Private Function ParseMemoryTypeColumns(ws As Worksheet, Optional validMemTypes As Object = Nothing) As Collection
    Dim memTypeColumns As Collection
    Set memTypeColumns = New Collection
    
    Dim lastCol As Long
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    
    Dim col As Long
    For col = 8 To lastCol
        Dim headerText As String
        headerText = Trim(CStr(ws.Cells(1, col).Value))
        
        If headerText <> "" Then
            Dim memType As String
            
            If ParseMemTypeHeader(headerText, memType) Then
                ' Check if this memory type is valid (exists in Instance List)
                Dim isValid As Boolean
                isValid = True
                
                If Not validMemTypes Is Nothing Then
                    isValid = False
                    ' Check for exact match or match with wildcard
                    Dim vKey As Variant
                    For Each vKey In validMemTypes.Keys
                        ' Check if header memType matches valid type (case insensitive)
                        If UCase(memType) = UCase(vKey) Then
                            isValid = True
                            Exit For
                        End If
                    Next vKey
                End If
                
                If isValid Then
                    Dim colInfo As Object
                    Set colInfo = CreateObject("Scripting.Dictionary")
                    colInfo("column") = col
                    colInfo("memory_type") = memType
                    ' Since we validated against Instance List, it is valid by definition
                    colInfo("valid_in_instances") = True
                    memTypeColumns.Add colInfo
                End If
            End If
        End If
    Next col
    
    Set ParseMemoryTypeColumns = memTypeColumns
End Function

Private Function ParseMemTypeHeader(ByVal headerText As String, ByRef memoryType As String) As Boolean
    Dim cleanText As String, parenPos As Long
    
    ParseMemTypeHeader = False
    memoryType = ""
    
    cleanText = Replace(headerText, vbLf, " ")
    cleanText = Replace(cleanText, vbCr, " ")
    cleanText = Trim(cleanText)
    
    ' If it doesn't have parentheses, it might just be the memory type (e.g. "HPP")
    parenPos = InStr(cleanText, "(")
    If parenPos = 0 Then
        If cleanText <> "" Then
            memoryType = cleanText
            ParseMemTypeHeader = True
        End If
        Exit Function
    End If
    
    memoryType = Trim(Left(cleanText, parenPos - 1))
    ParseMemTypeHeader = (memoryType <> "")
End Function




