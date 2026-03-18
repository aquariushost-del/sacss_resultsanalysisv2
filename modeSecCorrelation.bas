Attribute VB_Name = "modeSecCorrelation"
Option Explicit

'=========================================================
' Module: modSecCorrelation
'
' PURPOSE:
'   Build SEC subject-score correlation matrices for a
'   given year, one block per assessment sheet (e.g.
'   S2_WA2_2025), and output them into a single sheet:
'
'       SEC_Correl_<Year>   e.g. SEC_Correl_2025
'
' HOW TO RUN:
'   - General-purpose:
'       BuildSecCorrelationForYear 2025
'       BuildSecCorrelationForYear 2026
'
'   - Prompt user for year (easiest for button/menu):
'       BuildSecCorrelation_PromptYear
'
'   - Convenience wrapper for 2025:
'       BuildSecCorrelation_2025
'
' KEY FEATURES:
'   - Uses SCORE columns (e.g. "Eng - O (Score)").
'   - Pairwise deletion: each correlation uses only
'     students who have both scores.
'   - Only includes subjects with at least
'         MIN_N_SUBJECT (default 30)
'     valid scores.
'   - Only computes correlations for pairs with at least
'         MIN_N_PAIR (default 30)
'     overlapping students.
'   - Auto-detects SEC result sheets:
'       * Name starts with "S"
'       * Name contains the target year (e.g. "2025")
'       * Row 1 contains "Class" as a header
'       * Name does NOT contain "Subj Analysis"
'       * Name is not "Dashboard", "Settings",
'         or any *_Correl_* sheet.
'   - Produces a block per assessment:
'       "Level: S2   Assessment: S2_WA2_2025"
'         [mini summary panel]
'         [correlation matrix]
'         [blank row]
'         [Key Insights]
'
'   - Matrix display:
'       * SHORT subject labels (e.g. "Eng", "Math", "Sci")
'         extracted from header before first "-".
'       * Diagonal is displayed as "-" (dash) instead of 1.00.
'       * Cells with |r| ³ MOD_R_THRESHOLD (0.50) are
'         highlighted light green as "meaningful" correlations.
'       * Subject header row + row headers shaded.
'
'   - Each block appears as a "boxed card":
'       border around title + summary + matrix + insights.
'
'=========================================================

'------------------------------
' Correlation thresholds
'------------------------------
Private Const MIN_N_SUBJECT As Long = 30     ' minimum N to include subject
Private Const MIN_N_PAIR As Long = 30        ' minimum N to compute correlation

Private Const HIGH_R_THRESHOLD As Double = 0.65  ' |r| ³ 0.65 => strong
Private Const MOD_R_THRESHOLD As Double = 0.5    ' |r| ³ 0.50 => moderate/significant
Private Const LOW_R_THRESHOLD As Double = 0.3    ' |r| < 0.30 => "independent"

'------------------------------
' Helper types
'------------------------------
Private Type SubjectInfo
    FullName   As String   ' full header in source sheet (e.g. "Eng - O (Score)")
    ShortName  As String   ' display label in matrix (e.g. "Eng")
    ColIndex   As Long     ' column index in source sheet
    nValid     As Long     ' number of valid numeric scores
End Type

Private Type CorrPair
    IndexA As Long     ' subject index (1..p)
    IndexB As Long     ' subject index (1..p)
    r      As Double   ' correlation coefficient
    n      As Long     ' overlapping N used for this r
End Type

'=========================================================
' PUBLIC ENTRY POINTS
'=========================================================

' Convenience wrapper for a fixed year (e.g. legacy buttons)
Public Sub BuildSecCorrelation_2025()
    BuildSecCorrelationForYear 2025
End Sub

' General-purpose entry point:
' Call e.g. BuildSecCorrelationForYear 2025, BuildSecCorrelationForYear 2026, etc.
Public Sub BuildSecCorrelationForYear(ByVal targetYear As Long)
    Dim wb As Workbook
    Dim wsCorr As Worksheet
    Dim ws As Worksheet
    Dim corrSheetName As String
    Dim nextRow As Long
    Dim anyBlocks As Boolean
    Dim lastUsedCol As Long
    
    Set wb = ThisWorkbook
    corrSheetName = "SEC_Correl_" & CStr(targetYear)
    
    ' Ensure / create correlation sheet
    On Error Resume Next
    Set wsCorr = wb.Worksheets(corrSheetName)
    On Error GoTo 0
    
    If wsCorr Is Nothing Then
        Set wsCorr = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.count))
        wsCorr.Name = corrSheetName
    End If
    
    ' Clear existing content
    wsCorr.Cells.Clear
    
    ' Header
    wsCorr.Range("A1").value = "SEC Subject Score Correlation Matrix Ð " & targetYear
    wsCorr.Range("A1").Font.Bold = True
    wsCorr.Range("A1").Font.Size = 14
    
    wsCorr.Range("A2").value = _
        "Correlations are based on subject SCORE columns only, using pairwise deletion, " & _
        "and are computed only when at least " & MIN_N_PAIR & _
        " students have valid scores in both subjects. Cells with |r| ³ " & _
        MOD_R_THRESHOLD & " are highlighted light green."
    
    nextRow = 4
    anyBlocks = False
    
    ' Loop through worksheets and build blocks
    For Each ws In wb.Worksheets
        If IsSecSourceSheet(ws, targetYear, corrSheetName) Then
            BuildCorrelationBlockForSheet wsCorr, nextRow, ws, targetYear, anyBlocks
        End If
    Next ws
    
    If Not anyBlocks Then
        wsCorr.Cells(nextRow, 1).value = "No suitable SEC result sheets found for year " & targetYear & "."
    End If
    
    ' Autofit columns B onwards, keep column A at a fixed smaller width
    lastUsedCol = wsCorr.Cells(1, wsCorr.Columns.count).End(xlToLeft).Column
    If lastUsedCol < 2 Then lastUsedCol = 2
    wsCorr.Range(wsCorr.Columns(2), wsCorr.Columns(lastUsedCol)).EntireColumn.AutoFit
    
    wsCorr.Columns(1).ColumnWidth = 12   ' narrow subject-name column
End Sub

' Prompt-based runner Ð easiest to hook to a button/menu
Public Sub BuildSecCorrelation_PromptYear()
    Dim s As String
    Dim yr As Long
    
    s = InputBox("Enter SEC assessment year (e.g. 2025):", "SEC Correlation Year")
    If Trim$(s) = "" Then Exit Sub
    
    If Not IsNumeric(s) Then
        MsgBox "Please enter a valid numeric year (e.g. 2025).", vbExclamation, "SEC Correlation"
        Exit Sub
    End If
    
    yr = CLng(s)
    If yr < 2000 Or yr > 2100 Then
        MsgBox "Please enter a realistic year between 2000 and 2100.", vbExclamation, "SEC Correlation"
        Exit Sub
    End If
    
    BuildSecCorrelationForYear yr
End Sub

'=========================================================
' SOURCE SHEET DETECTION (SEC)
'=========================================================

Private Function IsSecSourceSheet(ByVal ws As Worksheet, _
                                  ByVal targetYear As Long, _
                                  ByVal corrSheetName As String) As Boolean
    Dim nm As String
    nm = ws.Name
    
    IsSecSourceSheet = False
    
    ' Skip obvious non-data sheets
    If nm = corrSheetName Then Exit Function
    If nm = "Dashboard" Then Exit Function
    If nm = "Settings" Then Exit Function
    If nm Like "SEC_Correl_*" Then Exit Function
    If nm Like "IP_Correl_*" Then Exit Function
    If InStr(1, nm, "Subj Analysis", vbTextCompare) > 0 Then Exit Function
    
    ' Must start with "S" (SEC levels like S1, S2...)
    If Left$(nm, 1) <> "S" Then Exit Function
    
    ' Must contain the target year somewhere in the name
    If InStr(1, nm, CStr(targetYear), vbTextCompare) = 0 Then Exit Function
    
    ' Must have "Class" in header row
    If FindHeaderColumn(ws, "Class") = 0 Then Exit Function
    
    IsSecSourceSheet = True
End Function

Private Function FindHeaderColumn(ByVal ws As Worksheet, _
                                  ByVal headerText As String) As Long
    Dim lastCol As Long
    Dim c As Long
    Dim cellVal As String
    
    FindHeaderColumn = 0
    lastCol = ws.Cells(1, ws.Columns.count).End(xlToLeft).Column
    
    For c = 1 To lastCol
        cellVal = Trim$(CStr(ws.Cells(1, c).value))
        If LCase$(cellVal) = LCase$(headerText) Then
            FindHeaderColumn = c
            Exit Function
        End If
    Next c
End Function

'=========================================================
' BUILD BLOCK FOR ONE ASSESSMENT SHEET
'=========================================================

Private Sub BuildCorrelationBlockForSheet(ByVal wsCorr As Worksheet, _
                                          ByRef nextRow As Long, _
                                          ByVal src As Worksheet, _
                                          ByVal targetYear As Long, _
                                          ByRef anyBlocks As Boolean)
    Dim subjects() As SubjectInfo
    Dim p As Long
    
    Dim lastRow As Long
    Dim lastCol As Long
    
    Dim levelKey As String
    
    Dim corr() As Double
    Dim corrValid() As Boolean
    Dim nPair() As Long
    
    Dim rTitleRow As Long
    Dim rSummaryRow As Long
    Dim rHeaderRow As Long
    Dim rMatrixTop As Long
    Dim rMatrixBottom As Long
    Dim rInsightsStart As Long
    Dim rInsightsEnd As Long
    
    Dim i As Long, j As Long
    Dim matrixRow As Long, matrixCol As Long
    Dim rVal As Double
    
    Dim rngMatrix As Range
    Dim rngHeader As Range
    Dim rngRowHeader As Range
    Dim rngCard As Range
    
    ' Determine used range in source
    lastRow = src.Cells(src.Rows.count, 1).End(xlUp).Row
    If lastRow < 3 Then Exit Sub  ' not enough rows
    
    lastCol = src.Cells(1, src.Columns.count).End(xlToLeft).Column
    
    ' Collect subject SCORE columns
    subjects = CollectScoreSubjects(src, lastRow, lastCol, p)
    
    If p < 2 Then
        wsCorr.Cells(nextRow, 1).value = "Level: " & GetLevelKeyFromSheet(src) & _
                                         "   Assessment: " & src.Name
        wsCorr.Cells(nextRow, 1).Font.Bold = True
        wsCorr.Cells(nextRow + 1, 1).value = "Not enough subjects with N ³ " & MIN_N_SUBJECT & _
                                             " valid scores to compute correlations."
        nextRow = nextRow + 4
        Exit Sub
    End If
    
    ' Compute correlation matrix
    ReDim corr(1 To p, 1 To p)
    ReDim corrValid(1 To p, 1 To p)
    ReDim nPair(1 To p, 1 To p)
    
    ComputeCorrelationMatrix src, subjects, p, lastRow, corr, corrValid, nPair
    
    ' Check if we have at least one off-diagonal valid pair
    If Not HasAnyValidPair(corrValid, p) Then
        wsCorr.Cells(nextRow, 1).value = "Level: " & GetLevelKeyFromSheet(src) & _
                                         "   Assessment: " & src.Name
        wsCorr.Cells(nextRow, 1).Font.Bold = True
        wsCorr.Cells(nextRow + 1, 1).value = "Not enough overlapping students (N ³ " & _
                                             MIN_N_PAIR & ") to compute meaningful correlations."
        nextRow = nextRow + 4
        Exit Sub
    End If
    
    anyBlocks = True
    
    '-----------------------------
    ' Title + mini summary panel
    '-----------------------------
    levelKey = GetLevelKeyFromSheet(src)
    
    rTitleRow = nextRow
    wsCorr.Cells(rTitleRow, 1).value = "Level: " & levelKey & "   Assessment: " & src.Name
    With wsCorr.Cells(rTitleRow, 1)
        .Font.Bold = True
        .Font.Size = 13
        .Font.Color = RGB(0, 51, 102) ' dark blue
    End With
    
    ' Mini summary row (panel)
    rSummaryRow = rTitleRow + 1
    WriteMiniSummary wsCorr, rSummaryRow, subjects, p, corr, corrValid, nPair
    
    '-----------------------------
    ' Matrix header + body
    '-----------------------------
    rHeaderRow = rSummaryRow + 1        ' subjects across top
    rMatrixTop = rHeaderRow + 1         ' first data row
    rMatrixBottom = rMatrixTop + p - 1
    
    ' Column headers (subject names across)
    For j = 1 To p
        wsCorr.Cells(rHeaderRow, 1 + j).value = subjects(j).ShortName
        wsCorr.Cells(rHeaderRow, 1 + j).Font.Bold = True
        wsCorr.Cells(rHeaderRow, 1 + j).HorizontalAlignment = xlCenter
    Next j
    
    ' Row headers + values
    Set rngMatrix = wsCorr.Range(wsCorr.Cells(rMatrixTop, 2), wsCorr.Cells(rMatrixBottom, 1 + p))
    rngMatrix.Interior.ColorIndex = xlNone  ' clear previous color
    rngMatrix.ClearFormats
    
    For i = 1 To p
        matrixRow = rMatrixTop + (i - 1)
        
        ' Row header in col A
        wsCorr.Cells(matrixRow, 1).value = subjects(i).ShortName
        wsCorr.Cells(matrixRow, 1).Font.Bold = True
        
        For j = 1 To p
            matrixCol = 1 + j
            
            If i = j Then
                ' Diagonal: dash
                wsCorr.Cells(matrixRow, matrixCol).value = "-"
                wsCorr.Cells(matrixRow, matrixCol).HorizontalAlignment = xlCenter
                wsCorr.Cells(matrixRow, matrixCol).Interior.ColorIndex = xlNone
            ElseIf corrValid(i, j) Then
                rVal = corr(i, j)
                wsCorr.Cells(matrixRow, matrixCol).value = Round(rVal, 2)
                wsCorr.Cells(matrixRow, matrixCol).NumberFormat = "0.00"
                wsCorr.Cells(matrixRow, matrixCol).HorizontalAlignment = xlCenter
                
                ' Highlight meaningful correlations |r| ³ MOD_R_THRESHOLD (0.50)
                If Abs(rVal) >= MOD_R_THRESHOLD Then
                    wsCorr.Cells(matrixRow, matrixCol).Interior.Color = RGB(198, 239, 206) ' light green
                Else
                    wsCorr.Cells(matrixRow, matrixCol).Interior.ColorIndex = xlNone
                End If
            Else
                ' Not enough N or undefined
                wsCorr.Cells(matrixRow, matrixCol).value = ""
                wsCorr.Cells(matrixRow, matrixCol).HorizontalAlignment = xlCenter
                wsCorr.Cells(matrixRow, matrixCol).Interior.ColorIndex = xlNone
            End If
        Next j
    Next i
    
    ' Shade header row + row header column
    Set rngHeader = wsCorr.Range(wsCorr.Cells(rHeaderRow, 1), wsCorr.Cells(rHeaderRow, 1 + p))
    With rngHeader
        .Interior.Color = RGB(242, 242, 242) ' light grey
        .Font.Bold = True
    End With
    
    Set rngRowHeader = wsCorr.Range(wsCorr.Cells(rMatrixTop, 1), wsCorr.Cells(rMatrixBottom, 1))
    With rngRowHeader
        .Interior.Color = RGB(242, 242, 242)
        .Font.Bold = True
    End With
    
    '-----------------------------
    ' Key Insights section
    ' Leave ONE blank row after matrix before insights
    '-----------------------------
    rInsightsStart = rMatrixBottom + 2   ' +1 = blank row, +2 = Key Insights line
    WriteCorrelationInsights wsCorr, rInsightsStart, subjects, p, corr, corrValid, nPair, rInsightsEnd
    
    '-----------------------------
    ' Boxed card around the block
    '-----------------------------
    Dim cardRightCol As Long
    cardRightCol = 1 + p ' up to last subject column
    
    Set rngCard = wsCorr.Range(wsCorr.Cells(rTitleRow, 1), wsCorr.Cells(rInsightsEnd, cardRightCol))
    With rngCard.Borders
        .LineStyle = xlContinuous
        .Weight = xlThin
        .Color = RGB(191, 191, 191) ' grey border
    End With
    ' Slightly thicker top border to emphasise card
    With rngCard.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Weight = xlMedium
    End With
    
    ' Extra spacing before next block
    nextRow = rInsightsEnd + 4
End Sub

'=========================================================
' MINI SUMMARY PANEL
'=========================================================

Private Sub WriteMiniSummary(ByVal wsCorr As Worksheet, _
                             ByVal rSummaryRow As Long, _
                             ByRef subjects() As SubjectInfo, _
                             ByVal p As Long, _
                             ByRef corr() As Double, _
                             ByRef corrValid() As Boolean, _
                             ByRef nPair() As Long)
    Dim text As String
    Dim i As Long, j As Long
    Dim subjectList As String
    Dim maxR As Double
    Dim maxI As Long, maxJ As Long
    Dim hasMax As Boolean
    Dim countMeaningful As Long
    Dim r As Double
    
    ' List of included subjects (short names)
    For i = 1 To p
        If subjectList <> "" Then subjectList = subjectList & ", "
        subjectList = subjectList & subjects(i).ShortName
    Next i
    
    ' Find strongest pair and count meaningful correlations (|r| ³ 0.50)
    hasMax = False
    countMeaningful = 0
    For i = 1 To p - 1
        For j = i + 1 To p
            If corrValid(i, j) Then
                r = corr(i, j)
                If Abs(r) >= MOD_R_THRESHOLD Then
                    countMeaningful = countMeaningful + 1
                End If
                If Not hasMax Then
                    maxR = r
                    maxI = i
                    maxJ = j
                    hasMax = True
                ElseIf Abs(r) > Abs(maxR) Then
                    maxR = r
                    maxI = i
                    maxJ = j
                End If
            End If
        Next j
    Next i
    
    If hasMax Then
        text = "Summary: " & p & " subjects included (N ³ " & MIN_N_SUBJECT & _
               "); " & countMeaningful & " meaningful correlations (|r| ³ " & MOD_R_THRESHOLD & _
               "). Strongest: " & subjects(maxI).ShortName & "Ð" & subjects(maxJ).ShortName & _
               " (r = " & Format$(maxR, "0.00") & ", N = " & nPair(maxI, maxJ) & ")."
    Else
        text = "Summary: " & p & " subjects included (N ³ " & MIN_N_SUBJECT & _
               "), but no valid subject pairs with N ³ " & MIN_N_PAIR & "."
    End If
    
    ' Write summary across the block and shade it lightly
    Dim lastCol As Long
    lastCol = 1 + p ' up to last subject column
    
    With wsCorr.Range(wsCorr.Cells(rSummaryRow, 1), wsCorr.Cells(rSummaryRow, lastCol))
        .Merge
        .value = text
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
        .Interior.Color = RGB(255, 242, 204) ' light yellow
        .Font.Bold = True
        .RowHeight = 18
    End With
End Sub

'=========================================================
' COLLECT SUBJECT SCORE COLUMNS
'=========================================================

Private Function CollectScoreSubjects(ByVal src As Worksheet, _
                                      ByVal lastRow As Long, _
                                      ByVal lastCol As Long, _
                                      ByRef p As Long) As SubjectInfo()
    Dim tmp() As SubjectInfo
    Dim countSubjects As Long
    Dim c As Long
    Dim headerText As String
    Dim r As Long
    Dim val As Variant
    Dim nValid As Long
    
    ReDim tmp(1 To lastCol)  ' max possible; will shrink later
    countSubjects = 0
    
    ' Identify candidate SCORE columns
    For c = 1 To lastCol
        headerText = Trim$(CStr(src.Cells(1, c).value))
        
        ' Heuristic: SCORE columns contain "(Score"
        If headerText <> "" And InStr(1, headerText, "(Score", vbTextCompare) > 0 Then
            
            ' Count how many numeric entries (valid scores) from row 2..lastRow
            nValid = 0
            For r = 2 To lastRow
                val = src.Cells(r, c).value
                If IsNumeric(val) Then
                    If Not IsEmpty(val) Then
                        nValid = nValid + 1
                    End If
                End If
            Next r
            
            ' Only include subjects with N ³ MIN_N_SUBJECT
            If nValid >= MIN_N_SUBJECT Then
                countSubjects = countSubjects + 1
                tmp(countSubjects).FullName = headerText
                tmp(countSubjects).ShortName = GetShortSubjectName(headerText)
                tmp(countSubjects).ColIndex = c
                tmp(countSubjects).nValid = nValid
            End If
        End If
    Next c
    
    ' Shrink array to actual size
    If countSubjects > 0 Then
        ReDim Preserve tmp(1 To countSubjects)
        p = countSubjects
        CollectScoreSubjects = tmp
    Else
        ReDim tmp(1 To 1)
        p = 0
        CollectScoreSubjects = tmp
    End If
End Function

Private Function GetShortSubjectName(ByVal fullHeader As String) As String
    Dim parts() As String
    
    ' Use text before first "-" as short name, else entire header
    If InStr(1, fullHeader, "-", vbTextCompare) > 0 Then
        parts = Split(fullHeader, "-")
        GetShortSubjectName = Trim$(parts(0))
    Else
        GetShortSubjectName = Trim$(fullHeader)
    End If
End Function

'=========================================================
' CORRELATION MATRIX COMPUTATION
'=========================================================

Private Sub ComputeCorrelationMatrix(ByVal src As Worksheet, _
                                     ByRef subjects() As SubjectInfo, _
                                     ByVal p As Long, _
                                     ByVal lastRow As Long, _
                                     ByRef corr() As Double, _
                                     ByRef corrValid() As Boolean, _
                                     ByRef nPair() As Long)
    Dim i As Long, j As Long, r As Long
    Dim colI As Long, colJ As Long
    Dim v1 As Variant, v2 As Variant
    
    Dim n As Long
    Dim sumX As Double, sumY As Double
    Dim sumX2 As Double, sumY2 As Double
    Dim sumXY As Double
    Dim num As Double, den As Double
    
    ' Initialise diagonal (we compute them, but will display "-" later)
    For i = 1 To p
        corr(i, i) = 1#
        corrValid(i, i) = True
        nPair(i, i) = subjects(i).nValid
    Next i
    
    ' Compute off-diagonal pairs with pairwise deletion
    For i = 1 To p - 1
        colI = subjects(i).ColIndex
        
        For j = i + 1 To p
            colJ = subjects(j).ColIndex
            
            n = 0
            sumX = 0#
            sumY = 0#
            sumX2 = 0#
            sumY2 = 0#
            sumXY = 0#
            
            For r = 2 To lastRow
                v1 = src.Cells(r, colI).value
                v2 = src.Cells(r, colJ).value
                
                If IsNumeric(v1) And IsNumeric(v2) Then
                    If Not IsEmpty(v1) And Not IsEmpty(v2) Then
                        n = n + 1
                        sumX = sumX + CDbl(v1)
                        sumY = sumY + CDbl(v2)
                        sumX2 = sumX2 + CDbl(v1) * CDbl(v1)
                        sumY2 = sumY2 + CDbl(v2) * CDbl(v2)
                        sumXY = sumXY + CDbl(v1) * CDbl(v2)
                    End If
                End If
            Next r
            
            If n >= MIN_N_PAIR Then
                num = n * sumXY - (sumX * sumY)
                den = (n * sumX2 - sumX * sumX) * (n * sumY2 - sumY * sumY)
                
                If den > 0 Then
                    corr(i, j) = num / Sqr(den)
                    corr(j, i) = corr(i, j)
                    corrValid(i, j) = True
                    corrValid(j, i) = True
                    nPair(i, j) = n
                    nPair(j, i) = n
                Else
                    ' No variance; correlation undefined -> treat as blank
                    corrValid(i, j) = False
                    corrValid(j, i) = False
                End If
            Else
                ' Not enough overlapping students
                corrValid(i, j) = False
                corrValid(j, i) = False
            End If
        Next j
    Next i
End Sub

Private Function HasAnyValidPair(ByRef corrValid() As Boolean, ByVal p As Long) As Boolean
    Dim i As Long, j As Long
    
    HasAnyValidPair = False
    For i = 1 To p - 1
        For j = i + 1 To p
            If corrValid(i, j) Then
                HasAnyValidPair = True
                Exit Function
            End If
        Next j
    Next i
End Function

'=========================================================
' AUTO-GENERATED KEY INSIGHTS
'=========================================================

Private Sub WriteCorrelationInsights(ByVal wsCorr As Worksheet, _
                                     ByVal startRow As Long, _
                                     ByRef subjects() As SubjectInfo, _
                                     ByVal p As Long, _
                                     ByRef corr() As Double, _
                                     ByRef corrValid() As Boolean, _
                                     ByRef nPair() As Long, _
                                     ByRef endRow As Long)
    Dim highPairs() As CorrPair
    Dim modPairs() As CorrPair
    Dim lowPairs() As CorrPair
    Dim countHigh As Long
    Dim countMod As Long
    Dim countLow As Long
    
    Dim i As Long, j As Long
    Dim r As Double
    
    ' Prepare arrays (max possible size; will shrink later)
    ReDim highPairs(1 To p * p)
    ReDim modPairs(1 To p * p)
    ReDim lowPairs(1 To p * p)
    
    countHigh = 0
    countMod = 0
    countLow = 0
    
    ' Categorise pairs
    For i = 1 To p - 1
        For j = i + 1 To p
            If corrValid(i, j) Then
                r = corr(i, j)
                
                If Abs(r) >= HIGH_R_THRESHOLD Then
                    countHigh = countHigh + 1
                    highPairs(countHigh).IndexA = i
                    highPairs(countHigh).IndexB = j
                    highPairs(countHigh).r = r
                    highPairs(countHigh).n = nPair(i, j)
                ElseIf Abs(r) >= MOD_R_THRESHOLD Then
                    countMod = countMod + 1
                    modPairs(countMod).IndexA = i
                    modPairs(countMod).IndexB = j
                    modPairs(countMod).r = r
                    modPairs(countMod).n = nPair(i, j)
                ElseIf Abs(r) < LOW_R_THRESHOLD Then
                    countLow = countLow + 1
                    lowPairs(countLow).IndexA = i
                    lowPairs(countLow).IndexB = j
                    lowPairs(countLow).r = r
                    lowPairs(countLow).n = nPair(i, j)
                End If
            End If
        Next j
    Next i
    
    ' Shrink arrays
    If countHigh > 0 Then ReDim Preserve highPairs(1 To countHigh)
    If countMod > 0 Then ReDim Preserve modPairs(1 To countMod)
    If countLow > 0 Then ReDim Preserve lowPairs(1 To countLow)
    
    Dim rLine As Long
    rLine = startRow
    
    wsCorr.Cells(rLine, 1).value = "Key Insights:"
    wsCorr.Cells(rLine, 1).Font.Bold = True
    rLine = rLine + 1
    
    ' Line 1: strong pairs
    If countHigh > 0 Then
        wsCorr.Cells(rLine, 1).value = "¥ Strong alignment pairs (|r| ³ " & HIGH_R_THRESHOLD & "): " & _
            FormatPairList(highPairs, subjects, 3)
        rLine = rLine + 1
    Else
        wsCorr.Cells(rLine, 1).value = "¥ Strong alignment pairs (|r| ³ " & HIGH_R_THRESHOLD & "): None identified."
        rLine = rLine + 1
    End If
    
    ' Line 2: moderate/significant pairs (0.50 ² |r| < 0.65)
    If countMod > 0 Then
        wsCorr.Cells(rLine, 1).value = "¥ Moderate/significant pairs (" & MOD_R_THRESHOLD & " ² |r| < " & _
            HIGH_R_THRESHOLD & "): " & FormatPairList(modPairs, subjects, 3)
        rLine = rLine + 1
    Else
        wsCorr.Cells(rLine, 1).value = "¥ Moderate/significant pairs: None highlighted."
        rLine = rLine + 1
    End If
    
    ' Line 3: relatively independent pairs
    If countLow > 0 Then
        wsCorr.Cells(rLine, 1).value = "¥ Relatively independent pairs (|r| < " & LOW_R_THRESHOLD & "): " & _
            FormatPairList(lowPairs, subjects, 3)
        rLine = rLine + 1
    Else
        wsCorr.Cells(rLine, 1).value = "¥ Relatively independent pairs (|r| < " & LOW_R_THRESHOLD & "): None highlighted."
        rLine = rLine + 1
    End If
    
    ' Line 4: methodological note
    wsCorr.Cells(rLine, 1).value = "Note: Correlations are based on subject SCORE columns only, " & _
                                   "and only computed when at least " & MIN_N_PAIR & _
                                   " students have valid scores in both subjects."
    
    endRow = rLine
End Sub

Private Function FormatPairList(ByRef pairs() As CorrPair, _
                                ByRef subjects() As SubjectInfo, _
                                ByVal maxItems As Long) As String
    Dim result As String
    Dim i As Long
    Dim limit As Long
    Dim nameA As String, nameB As String
    
    On Error GoTo EmptyList
    If UBound(pairs) < 1 Then GoTo EmptyList
    
    limit = Application.WorksheetFunction.Min(maxItems, UBound(pairs))
    
    For i = 1 To limit
        nameA = subjects(pairs(i).IndexA).ShortName
        nameB = subjects(pairs(i).IndexB).ShortName
        
        If result <> "" Then result = result & "; "
        result = result & nameA & "Ð" & nameB & " (r = " & _
                 Format$(pairs(i).r, "0.00") & ", N = " & pairs(i).n & ")"
    Next i
    
    If UBound(pairs) > limit Then
        result = result & " É"
    End If
    
    FormatPairList = result
    Exit Function
    
EmptyList:
    FormatPairList = "None."
End Function

'=========================================================
' LEVEL KEY (e.g. "S1", "S2")
'=========================================================

Private Function GetLevelKeyFromSheet(ByVal src As Worksheet) As String
    Dim classCol As Long
    Dim lastRow As Long
    Dim r As Long
    Dim val As String
    Dim firstChar As String
    
    classCol = FindHeaderColumn(src, "Class")
    If classCol = 0 Then
        GetLevelKeyFromSheet = "S?"
        Exit Function
    End If
    
    lastRow = src.Cells(src.Rows.count, classCol).End(xlUp).Row
    
    For r = 2 To lastRow
        val = Trim$(CStr(src.Cells(r, classCol).value))
        If val <> "" Then
            firstChar = Left$(val, 1)
            If IsNumeric(firstChar) Then
                GetLevelKeyFromSheet = "S" & firstChar
                Exit Function
            End If
        End If
    Next r
    
    GetLevelKeyFromSheet = "S?"
End Function

