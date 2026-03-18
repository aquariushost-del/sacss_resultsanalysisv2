Attribute VB_Name = "modPhysicsMsgDebug"
Option Explicit

'=========================================================
' Module: modPhysicsMsgDebug
'
' PURPOSE:
'   Standalone debug tool to compute **Sec 4 Physics cohort MSG**
'   from a specific source sheet and log every step.
'
' SOURCE:
'   Sheet name : "S4_PRELIMINARYEXAM_2025"
'   Class col  : header "Class"
'   Physics col: header "Phy - O (Grade)"
'
' OUTPUT:
'   Sheet "PhyLogs" (created/cleared automatically) with:
'     - Per-row log:
'         Row, Class, GradeText, IsSec4, IsValidGrade,
'         IncludedInCohort, NumericValue
'     - Summary:
'         Counts of A1..F9, Total included, Weighted sum, MSG
'
' NOTES:
'   - SEC Grade mapping:
'       A1=1, A2=2, B3=3, B4=4, C5=5, C6=6, D7=7, E8=8, F9=9
'   - Only rows where:
'       * Class starts with "4" (e.g. 4A, 4B...)
'       * AND grade is a valid SEC grade (A1..F9)
'     are included in the MSG computation.
'=========================================================

Private Const SRC_SHEET_NAME As String = "S4_PRELIMINARYEXAM_2025"
Private Const LOG_SHEET_NAME As String = "PhyLogs"
Private Const CLASS_HEADER As String = "Class"
Private Const PHY_HEADER As String = "Phy - O (Grade)"

'---------------------------------------------------------
' ENTRY POINT
'---------------------------------------------------------
Public Sub ComputeSec4PhysicsCohortMSG()
    Dim wb As Workbook
    Dim wsSrc As Worksheet
    Dim wsLog As Worksheet
    Dim classCol As Long, phyCol As Long
    Dim lastRow As Long, r As Long
    Dim logRow As Long
    
    Dim className As String
    Dim gradeText As String
    Dim isSec4 As Boolean
    Dim isValidGrade As Boolean
    Dim included As Boolean
    Dim gradeValue As Long
    
    Dim counts(1 To 9) As Long
    Dim totalIncluded As Long
    Dim weightedSum As Long
    Dim msgValue As Double
    
    On Error GoTo ErrHandler
    
    Set wb = ThisWorkbook
    
    '-----------------------------
    ' 1) Get source sheet
    '-----------------------------
    On Error Resume Next
    Set wsSrc = wb.Worksheets(SRC_SHEET_NAME)
    On Error GoTo ErrHandler
    
    If wsSrc Is Nothing Then
        MsgBox "Source sheet '" & SRC_SHEET_NAME & "' not found.", vbCritical
        Exit Sub
    End If
    
    '-----------------------------
    ' 2) Find columns
    '-----------------------------
    classCol = FindHeaderColumn_Local(wsSrc, 1, CLASS_HEADER)
    phyCol = FindHeaderColumn_Local(wsSrc, 1, PHY_HEADER)
    
    If classCol = 0 Then
        MsgBox "Header '" & CLASS_HEADER & "' not found in row 1 of '" & SRC_SHEET_NAME & "'.", vbCritical
        Exit Sub
    End If
    
    If phyCol = 0 Then
        MsgBox "Header '" & PHY_HEADER & "' not found in row 1 of '" & SRC_SHEET_NAME & "'.", vbCritical
        Exit Sub
    End If
    
    lastRow = wsSrc.Cells(wsSrc.Rows.count, classCol).End(xlUp).Row
    If lastRow < 2 Then
        MsgBox "No data rows found in '" & SRC_SHEET_NAME & "'.", vbInformation
        Exit Sub
    End If
    
    '-----------------------------
    ' 3) Prepare log sheet
    '-----------------------------
    Set wsLog = PrepareLogSheet()
    
    ' Header row for detailed log
    logRow = 1
    With wsLog
        .Cells(logRow, 1).value = "Row"
        .Cells(logRow, 2).value = "Class"
        .Cells(logRow, 3).value = "GradeText"
        .Cells(logRow, 4).value = "IsSec4"
        .Cells(logRow, 5).value = "IsValidGrade(A1ÐF9)"
        .Cells(logRow, 6).value = "IncludedInCohort"
        .Cells(logRow, 7).value = "NumericValue(1Ð9)"
        .Rows(logRow).Font.Bold = True
    End With
    
    '-----------------------------
    ' 4) Scan each row and log
    '-----------------------------
    For r = 2 To lastRow
        className = Trim(CStr(wsSrc.Cells(r, classCol).value))
        gradeText = UCase$(Trim$(CStr(wsSrc.Cells(r, phyCol).value)))
        
        isSec4 = IsSec4Class(className)
        gradeValue = GradeToValueSEC(gradeText)
        isValidGrade = (gradeValue >= 1 And gradeValue <= 9)
        included = (isSec4 And isValidGrade)
        
        ' Log this row
        logRow = logRow + 1
        With wsLog
            .Cells(logRow, 1).value = r
            .Cells(logRow, 2).value = className
            .Cells(logRow, 3).value = gradeText
            .Cells(logRow, 4).value = IIf(isSec4, "Yes", "No")
            .Cells(logRow, 5).value = IIf(isValidGrade, "Yes", "No")
            .Cells(logRow, 6).value = IIf(included, "Yes", "No")
            If isValidGrade Then
                .Cells(logRow, 7).value = gradeValue
            Else
                .Cells(logRow, 7).ClearContents
            End If
        End With
        
        ' Update counts if included
        If included Then
            counts(gradeValue) = counts(gradeValue) + 1
            totalIncluded = totalIncluded + 1
            weightedSum = weightedSum + gradeValue
        End If
    Next r
    
    '-----------------------------
    ' 5) Summary and MSG
    '-----------------------------
    wsLog.Columns("A:G").AutoFit
    
    logRow = logRow + 2
    With wsLog
        .Cells(logRow, 1).value = "SUMMARY"
        .Cells(logRow, 1).Font.Bold = True
    End With
    logRow = logRow + 1
    
    ' Grade count header
    With wsLog
        .Cells(logRow, 1).value = "Grade"
        .Cells(logRow, 2).value = "Count"
        .Rows(logRow).Font.Bold = True
    End With
    logRow = logRow + 1
    
    ' A1..F9 counts
    Call WriteGradeCount(wsLog, logRow, "A1", counts(1)): logRow = logRow + 1
    Call WriteGradeCount(wsLog, logRow, "A2", counts(2)): logRow = logRow + 1
    Call WriteGradeCount(wsLog, logRow, "B3", counts(3)): logRow = logRow + 1
    Call WriteGradeCount(wsLog, logRow, "B4", counts(4)): logRow = logRow + 1
    Call WriteGradeCount(wsLog, logRow, "C5", counts(5)): logRow = logRow + 1
    Call WriteGradeCount(wsLog, logRow, "C6", counts(6)): logRow = logRow + 1
    Call WriteGradeCount(wsLog, logRow, "D7", counts(7)): logRow = logRow + 1
    Call WriteGradeCount(wsLog, logRow, "E8", counts(8)): logRow = logRow + 1
    Call WriteGradeCount(wsLog, logRow, "F9", counts(9)): logRow = logRow + 1
    
    ' Totals + MSG
    logRow = logRow + 1
    
    If totalIncluded > 0 Then
        msgValue = weightedSum / totalIncluded
    Else
        msgValue = 0
    End If
    
    With wsLog
        .Cells(logRow, 1).value = "Total Included (Sec 4 & valid Physics grades):"
        .Cells(logRow, 2).value = totalIncluded
        logRow = logRow + 1
        
        .Cells(logRow, 1).value = "Weighted Sum (S gradeValue):"
        .Cells(logRow, 2).value = weightedSum
        logRow = logRow + 1
        
        .Cells(logRow, 1).value = "MSG (WeightedSum / TotalIncluded):"
        .Cells(logRow, 2).value = msgValue
        .Cells(logRow, 2).NumberFormat = "0.00"
        .Cells(logRow, 1).Font.Bold = True
        .Cells(logRow, 2).Font.Bold = True
    End With
    
    wsLog.Columns("A:B").AutoFit
    
    MsgBox "Sec 4 Physics Cohort MSG = " & Format(msgValue, "0.00") & vbCrLf & _
           "(see '" & LOG_SHEET_NAME & "' for detailed logs).", vbInformation

    Exit Sub

ErrHandler:
    MsgBox "Error in ComputeSec4PhysicsCohortMSG: " & Err.Description, vbCritical
End Sub

'---------------------------------------------------------
' Is this a Sec 4 class?
'   Rule: first character = "4"
'   (You can tighten this if needed.)
'---------------------------------------------------------
Private Function IsSec4Class(ByVal className As String) As Boolean
    className = Trim$(className)
    If className = "" Then
        IsSec4Class = False
    Else
        IsSec4Class = (Left$(className, 1) = "4")
    End If
End Function

'---------------------------------------------------------
' Map SEC grade text -> numeric value 1..9
'   A1=1, A2=2, B3=3, B4=4, C5=5, C6=6, D7=7, E8=8, F9=9
'   Returns 0 if invalid.
'---------------------------------------------------------
Private Function GradeToValueSEC(ByVal gradeText As String) As Long
    Select Case gradeText
        Case "A1": GradeToValueSEC = 1
        Case "A2": GradeToValueSEC = 2
        Case "B3": GradeToValueSEC = 3
        Case "B4": GradeToValueSEC = 4
        Case "C5": GradeToValueSEC = 5
        Case "C6": GradeToValueSEC = 6
        Case "D7": GradeToValueSEC = 7
        Case "E8": GradeToValueSEC = 8
        Case "F9": GradeToValueSEC = 9
        Case Else: GradeToValueSEC = 0
    End Select
End Function

'---------------------------------------------------------
' Create / clear log sheet "PhyLogs"
'---------------------------------------------------------
Private Function PrepareLogSheet() As Worksheet
    Dim wb As Workbook
    Dim ws As Worksheet
    
    Set wb = ThisWorkbook
    
    On Error Resume Next
    Set ws = wb.Worksheets(LOG_SHEET_NAME)
    On Error GoTo 0
    
    If ws Is Nothing Then
        Set ws = wb.Worksheets.Add(After:=wb.Sheets(wb.Sheets.count))
        ws.Name = LOG_SHEET_NAME
    Else
        ws.Cells.Clear
    End If
    
    Set PrepareLogSheet = ws
End Function

'---------------------------------------------------------
' Find header column by exact name (case-insensitive) in row
'---------------------------------------------------------
Private Function FindHeaderColumn_Local(ws As Worksheet, _
                                       ByVal headerRow As Long, _
                                       ByVal headerName As String) As Long
    Dim lastCol As Long, c As Long
    Dim h As String
    
    lastCol = ws.Cells(headerRow, ws.Columns.count).End(xlToLeft).Column
    For c = 1 To lastCol
        h = Trim$(CStr(ws.Cells(headerRow, c).value))
        If StrComp(h, headerName, vbTextCompare) = 0 Then
            FindHeaderColumn_Local = c
            Exit Function
        End If
    Next c
    
    FindHeaderColumn_Local = 0
End Function

'---------------------------------------------------------
' Helper to write a single grade count line in summary
'---------------------------------------------------------
Private Sub WriteGradeCount(ws As Worksheet, ByVal rowNum As Long, _
                            ByVal gradeLabel As String, _
                            ByVal countValue As Long)
    ws.Cells(rowNum, 1).value = gradeLabel
    ws.Cells(rowNum, 2).value = countValue
End Sub


