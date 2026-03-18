Attribute VB_Name = "modSecGradeDistribution"
Option Explicit

Private Const DEFAULT_MIN_SUBJECT_N As Long = 10
Private Const DEFAULT_AT_RISK_FAIL_THRESHOLD As Long = 3
Private Const SHAPE_ROUNDED_RECTANGLE As Long = 5
Private Const ATRISK_NAV_SHEET_NAME As String = "Dashboard"
Private Const ATRISK_NAV_START_CELL As String = "M3"
Private Const ATRISK_NAV_BTN_PREFIX As String = "Nav_AtRisk_"
Private Const TOP_NAV_START_CELL As String = "T3"
Private Const TOP_NAV_BTN_PREFIX As String = "Nav_TopQual_"

Private Type TopStudentRec
    LevelCode As String
    ClassName As String
    RegNo As String
    StudentName As String
    GroupCode As String
    TopCount As Long
    TopPrimaryCount As Long
    TopSecondaryCount As Long
End Type

'=========================================================
' Module: modSecGradeDistribution
'
' PURPOSE:
'   Automatic subject analysis for G1 / G2 / G3 grade tracks.
'
' ENTRY POINT:
'   BuildAllSec_SubjectAnalysis
'=========================================================

'---------------------------------------------------------
' ENTRY POINT - RUN ONCE, DOES ALL ELIGIBLE SHEETS
'---------------------------------------------------------
Public Sub BuildAllSec_SubjectAnalysis()
    Dim wb As Workbook
    Dim ws As Worksheet

    On Error GoTo ErrHandler

    Set wb = ThisWorkbook

    For Each ws In wb.Worksheets
        ProcessSecSourceSheet ws
    Next ws

    MsgBox "Subject Analysis generated for all eligible sheets.", vbInformation
    Exit Sub

ErrHandler:
    MsgBox "Error in BuildAllSec_SubjectAnalysis: " & Err.Description, vbCritical
End Sub

Public Sub BuildSec_TopQualityByLevel()
    Dim wb As Workbook
    Dim ws As Worksheet, wsOut As Worksheet
    Dim lvl As Variant
    Dim recs() As TopStudentRec
    Dim recCount As Long
    Dim outRow As Long
    Dim groupThresholdPct As Double

    On Error GoTo ErrHandler

    Set wb = ThisWorkbook
    groupThresholdPct = GetGroupThresholdPercent()

    For Each lvl In Array("S1", "S2", "S3", "S4", "S5")
        Set wsOut = GetOrCreateWorksheet("TopQual_" & CStr(lvl))
        PrepareTopQualitySheet wsOut, CStr(lvl)

        recCount = 0
        For Each ws In wb.Worksheets
            AppendTopQualityFromSourceSheet ws, CStr(lvl), recs, recCount, groupThresholdPct
        Next ws

        outRow = 5
        outRow = WriteTopGroupSection(wsOut, outRow, CStr(lvl), "G3", 20, recs, recCount)
        outRow = WriteTopGroupSection(wsOut, outRow, CStr(lvl), "G2", 10, recs, recCount)
        outRow = WriteTopGroupSection(wsOut, outRow, CStr(lvl), "G1", 10, recs, recCount)

        FormatTopQualitySheet wsOut, outRow - 1
        AddAtRiskHomeButton wsOut
    Next lvl

    BuildTopQualityNavigation

    MsgBox "Top quality sheets built: TopQual_S1 to TopQual_S5.", vbInformation
    Exit Sub

ErrHandler:
    MsgBox "Error in BuildSec_TopQualityByLevel: " & Err.Description, vbCritical
End Sub

Public Sub BuildTopQualityNavigation()
    Dim wsNav As Worksheet
    Dim startCell As Range
    Dim startRow As Long, startCol As Long
    Dim rowPtr As Long
    Dim lvl As Variant
    Dim sheetName As String
    Dim shp As Shape
    Dim k As Long

    On Error GoTo ErrHandler

    On Error Resume Next
    Set wsNav = ThisWorkbook.Worksheets(ATRISK_NAV_SHEET_NAME)
    On Error GoTo ErrHandler
    If wsNav Is Nothing Then Exit Sub

    Set startCell = wsNav.Range(TOP_NAV_START_CELL)
    startRow = startCell.Row
    startCol = startCell.Column

    wsNav.Range(wsNav.Cells(startRow, startCol), wsNav.Cells(startRow + 120, startCol + 5)).Clear
    For k = wsNav.Shapes.count To 1 Step -1
        Set shp = wsNav.Shapes(k)
        If Left$(shp.Name, Len(TOP_NAV_BTN_PREFIX)) = TOP_NAV_BTN_PREFIX Then shp.Delete
    Next k

    wsNav.Cells(startRow, startCol).value = "Top Students Menu"
    wsNav.Cells(startRow, startCol).Font.Bold = True
    wsNav.Cells(startRow, startCol).Font.Size = 12
    wsNav.Cells(startRow, startCol).Font.Color = RGB(31, 73, 125)
    rowPtr = startRow + 1

    For Each lvl In Array("S1", "S2", "S3", "S4", "S5")
        sheetName = "TopQual_" & CStr(lvl)
        If WorksheetExistsByName(sheetName) Then
            CreateTopQualityNavButton wsNav, sheetName, CStr(lvl) & " Top Students", rowPtr, startCol
        Else
            wsNav.Cells(rowPtr, startCol).value = CStr(lvl) & " Top Students (not built)"
            wsNav.Cells(rowPtr, startCol).Font.Italic = True
        End If
        rowPtr = rowPtr + 2
    Next lvl
    Exit Sub

ErrHandler:
    ' Silent fallback
End Sub

Private Sub CreateTopQualityNavButton(ByVal wsNav As Worksheet, _
                                      ByVal targetSheetName As String, _
                                      ByVal displayText As String, _
                                      ByVal rowNum As Long, _
                                      ByVal firstCol As Long)
    Dim shp As Shape
    Dim leftPos As Double, topPos As Double
    Dim btnWidth As Double, btnHeight As Double

    leftPos = wsNav.Cells(rowNum, firstCol).Left
    topPos = wsNav.Cells(rowNum, firstCol).Top
    btnWidth = wsNav.Columns(firstCol).Resize(, 5).Width
    btnHeight = wsNav.Rows(rowNum).Height * 1.3

    Set shp = wsNav.Shapes.AddShape( _
        Type:=SHAPE_ROUNDED_RECTANGLE, _
        Left:=leftPos, _
        Top:=topPos, _
        Width:=btnWidth, _
        Height:=btnHeight)

    With shp
        .Name = TOP_NAV_BTN_PREFIX & targetSheetName
        .Fill.ForeColor.RGB = RGB(197, 217, 241)
        .Fill.Transparency = 0#
        .line.ForeColor.RGB = RGB(84, 141, 212)
        .line.Weight = 1.5
        With .TextFrame2
            .TextRange.text = displayText
            .TextRange.Font.Name = "Calibri"
            .TextRange.Font.Size = 10.5
            .TextRange.Font.Fill.ForeColor.RGB = RGB(31, 73, 125)
            .TextRange.ParagraphFormat.Alignment = msoAlignCenter
            .VerticalAnchor = msoAnchorMiddle
            .MarginLeft = 6
            .MarginRight = 6
            .MarginTop = 3
            .MarginBottom = 3
        End With
    End With

    wsNav.Hyperlinks.Add Anchor:=shp, Address:="", SubAddress:="'" & targetSheetName & "'!A1"
End Sub

'---------------------------------------------------------
' ENTRY POINT - BUILD STUDENTS AT RISK SUMMARY (SEC)
'---------------------------------------------------------
Public Sub BuildSec_AtRiskSummary()
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim wsOut As Worksheet
    Dim lvl As Variant
    Dim outRow As Long
    Dim addedRows As Long
    Dim threshold As Long

    On Error GoTo ErrHandler

    Set wb = ThisWorkbook
    threshold = GetAtRiskFailThreshold()

    For Each lvl In Array("S1", "S2", "S3", "S4", "S5")
        Set wsOut = GetOrCreateWorksheet("AtRisk_" & CStr(lvl))
        PrepareAtRiskSheet wsOut, CStr(lvl), threshold

        outRow = 5
        For Each ws In wb.Worksheets
            addedRows = AppendSecAtRiskFromSourceSheet(ws, wsOut, outRow, threshold, CStr(lvl))
            If addedRows > 0 Then outRow = outRow + addedRows
        Next ws

        If outRow = 5 Then
            wsOut.Cells(outRow, 1).value = "No eligible SEC result rows found for " & CStr(lvl) & "."
            outRow = outRow + 1
        End If

        FinalizeAtRiskSheet wsOut, outRow - 1
    Next lvl

    BuildSec_AtRiskNavigation
    BuildAllAtRiskHomeButtons

    wb.Worksheets("AtRisk_S1").Activate
    wb.Worksheets("AtRisk_S1").Range("A1").Select

    MsgBox "SEC at-risk sheets built: AtRisk_S1 to AtRisk_S5.", vbInformation
    Exit Sub

ErrHandler:
    MsgBox "Error in BuildSec_AtRiskSummary: " & Err.Description, vbCritical
End Sub

'---------------------------------------------------------
' PROCESS ONE SOURCE SHEET
'---------------------------------------------------------
Private Sub ProcessSecSourceSheet(ByVal wsSrc As Worksheet)
    Dim classCol As Long
    Dim lastRow As Long
    Dim firstClass As String
    Dim levelCode As String
    Dim lastCol As Long, c As Long
    Dim header As String

    Dim subjectCols() As Long
    Dim subjectNames() As String
    Dim subjectSchemeKeys() As String
    Dim subjCount As Long

    Dim examLabel As String
    Dim destSheetName As String
    Dim wb As Workbook
    Dim wsDest As Worksheet
    Dim i As Long
    Dim destRowHeader As Long
    Dim destTopLeft As String
    Dim titleText As String

    Const TABLE_GAP_ROWS As Long = 5

    On Error GoTo ErrHandler

    Set wb = ThisWorkbook

    If LCase$(wsSrc.Name) Like "*settings*" _
       Or LCase$(wsSrc.Name) Like "*config*" _
       Or LCase$(wsSrc.Name) Like "*menu*" _
       Or LCase$(wsSrc.Name) Like "*lookup*" _
       Or LCase$(wsSrc.Name) Like "*summary*" _
       Or LCase$(wsSrc.Name) Like "*template*" Then
        Exit Sub
    End If

    classCol = FindHeaderColumn(wsSrc, 1, "Class")
    If classCol = 0 Then Exit Sub

    lastRow = wsSrc.Cells(wsSrc.Rows.count, classCol).End(xlUp).Row
    firstClass = ""
    For i = 2 To lastRow
        firstClass = Trim$(CStr(wsSrc.Cells(i, classCol).value))
        If firstClass <> "" Then Exit For
    Next i
    If firstClass = "" Then Exit Sub

    levelCode = InferLevelCodeFromClass(firstClass)
    If levelCode = "" Then Exit Sub

    ' Detect subject grade columns and their schemes.
    lastCol = wsSrc.Cells(1, wsSrc.Columns.count).End(xlToLeft).Column
    subjCount = 0

    For c = 1 To lastCol
        If c <> classCol Then
            header = Trim$(CStr(wsSrc.Cells(1, c).value))
            If header <> "" And IsLikelySubjectGradeColumn(header) Then
                Dim schemeKey As String
                Dim subjectName As String
                schemeKey = GetGradeSchemeKey(wsSrc, c, header)
                subjectName = StripGradeHeaderSuffix(header)

                If schemeKey <> "" And Not SubjectAlreadyAdded(subjectNames, subjCount, subjectName) Then
                    subjCount = subjCount + 1
                    ReDim Preserve subjectCols(1 To subjCount)
                    ReDim Preserve subjectNames(1 To subjCount)
                    ReDim Preserve subjectSchemeKeys(1 To subjCount)

                    subjectCols(subjCount) = c
                    subjectNames(subjCount) = subjectName
                    subjectSchemeKeys(subjCount) = schemeKey
                End If
            End If
        End If
    Next c

    If subjCount = 0 Then Exit Sub

    examLabel = wsSrc.Name
    destSheetName = BuildSecDestSheetName(levelCode, examLabel)

    On Error Resume Next
    Set wsDest = wb.Worksheets(destSheetName)
    On Error GoTo ErrHandler

    If wsDest Is Nothing Then
        Set wsDest = wb.Worksheets.Add(After:=wb.Sheets(wb.Sheets.count))
        wsDest.Name = destSheetName
    Else
        Dim k As Long
        Dim shp As Shape

        wsDest.Cells.Clear

        For k = wsDest.ChartObjects.count To 1 Step -1
            wsDest.ChartObjects(k).Delete
        Next k

        For k = wsDest.Shapes.count To 1 Step -1
            Set shp = wsDest.Shapes(k)
            If Left$(shp.Name, 10) = "FlagPanel_" Then shp.Delete
        Next k
    End If

    titleText = levelCode & " Subject Grade Distribution (" & examLabel & ") - G1/G2/G3"
    With wsDest.Range("A1")
        .value = titleText
        .Font.Bold = True
        .Font.Size = 14
    End With

    destRowHeader = 3
    For i = 1 To subjCount
        destTopLeft = wsDest.Cells(destRowHeader, 1).Address(False, False)

        Dim tableEndRow As Long
        BuildSecSubjectGradeDistribution _
            srcSheetName:=wsSrc.Name, _
            srcClassCol:=classCol, _
            srcGradeCol:=subjectCols(i), _
            destSheetName:=wsDest.Name, _
            destTopLeft:=destTopLeft, _
            subjectTitle:=subjectNames(i), _
            schemeKey:=subjectSchemeKeys(i), _
            outEndRow:=tableEndRow

        If tableEndRow > 0 Then
            ' Keep exactly 5 blank rows between tables.
            destRowHeader = tableEndRow + TABLE_GAP_ROWS + 2
        End If
    Next i

    Exit Sub

ErrHandler:
    ' Skip broken sheets and continue with the rest.
End Sub

Private Sub AppendTopQualityFromSourceSheet(ByVal wsSrc As Worksheet, _
                                            ByVal targetLevel As String, _
                                            ByRef recs() As TopStudentRec, _
                                            ByRef recCount As Long, _
                                            ByVal groupThresholdPct As Double)
    Dim classCol As Long, nameCol As Long, regCol As Long
    Dim lastRow As Long, lastCol As Long
    Dim firstClass As String, levelCode As String
    Dim subjectCols() As Long, subjectNames() As String, subjectSchemeKeys() As String, subjectScoreCols() As Long
    Dim subjCount As Long
    Dim c As Long, r As Long, i As Long
    Dim header As String, schemeKey As String, subjectName As String
    Dim className As String, studentName As String, regNo As String
    Dim rawGrade As String, rawScore As String, gradeStr As String
    Dim isVrMc As Boolean
    Dim topCount As Long
    Dim topPrimaryCount As Long, topSecondaryCount As Long
    Dim g1GroupCount As Long, g2GroupCount As Long, g3GroupCount As Long, groupTotalCount As Long
    Dim fsbbGroup As String

    On Error GoTo FailSafe

    If LCase$(wsSrc.Name) Like "*settings*" _
       Or LCase$(wsSrc.Name) Like "*config*" _
       Or LCase$(wsSrc.Name) Like "*menu*" _
       Or LCase$(wsSrc.Name) Like "*lookup*" _
       Or LCase$(wsSrc.Name) Like "*summary*" _
       Or LCase$(wsSrc.Name) Like "*template*" _
       Or InStr(1, LCase$(wsSrc.Name), "_subj analysis_") > 0 _
       Or InStr(1, LCase$(wsSrc.Name), "dashboard") > 0 _
       Or InStr(1, LCase$(wsSrc.Name), "atrisk_") > 0 _
       Or InStr(1, LCase$(wsSrc.Name), "topqual_") > 0 Then
        Exit Sub
    End If

    classCol = FindHeaderColumn(wsSrc, 1, "Class")
    If classCol = 0 Then Exit Sub
    nameCol = FindFirstHeaderColumn(wsSrc, 1, Array("Name", "Student Name", "Student"))
    regCol = FindFirstHeaderColumn(wsSrc, 1, Array("RegNo", "Reg No", "Register No", "Index No", "Adm No"))

    lastRow = wsSrc.Cells(wsSrc.Rows.count, classCol).End(xlUp).Row
    firstClass = ""
    For r = 2 To lastRow
        firstClass = Trim$(CStr(wsSrc.Cells(r, classCol).value))
        If firstClass <> "" Then Exit For
    Next r
    If firstClass = "" Then Exit Sub

    levelCode = InferLevelCodeFromClass(firstClass)
    If levelCode = "" Then Exit Sub
    If UCase$(levelCode) <> UCase$(targetLevel) Then Exit Sub

    lastCol = wsSrc.Cells(1, wsSrc.Columns.count).End(xlToLeft).Column
    subjCount = 0
    For c = 1 To lastCol
        If c <> classCol Then
            header = Trim$(CStr(wsSrc.Cells(1, c).value))
            If header <> "" And IsLikelySubjectGradeColumn(header) Then
                schemeKey = GetGradeSchemeKey(wsSrc, c, header)
                subjectName = StripGradeHeaderSuffix(header)
                If schemeKey <> "" And Not SubjectAlreadyAdded(subjectNames, subjCount, subjectName) Then
                    subjCount = subjCount + 1
                    ReDim Preserve subjectCols(1 To subjCount)
                    ReDim Preserve subjectNames(1 To subjCount)
                    ReDim Preserve subjectSchemeKeys(1 To subjCount)
                    ReDim Preserve subjectScoreCols(1 To subjCount)
                    subjectCols(subjCount) = c
                    subjectNames(subjCount) = subjectName
                    subjectSchemeKeys(subjCount) = schemeKey
                    subjectScoreCols(subjCount) = FindScoreColumnForSubject(wsSrc, 1, subjectName)
                End If
            End If
        End If
    Next c
    If subjCount = 0 Then Exit Sub

    For r = 2 To lastRow
        className = Trim$(CStr(wsSrc.Cells(r, classCol).value))
        If className = "" Then GoTo NextStudent
        If UCase$(Left$(className, 1)) = "Y" Then GoTo NextStudent

        If nameCol > 0 Then
            studentName = Trim$(CStr(wsSrc.Cells(r, nameCol).value))
        Else
            studentName = ""
        End If
        If regCol > 0 Then
            regNo = Trim$(CStr(wsSrc.Cells(r, regCol).value))
        Else
            regNo = ""
        End If

        topCount = 0
        topPrimaryCount = 0
        topSecondaryCount = 0
        g1GroupCount = 0
        g2GroupCount = 0
        g3GroupCount = 0
        groupTotalCount = 0

        For i = 1 To subjCount
            rawGrade = UCase$(Trim$(CStr(wsSrc.Cells(r, subjectCols(i)).value)))
            gradeStr = NormalizeGradeForScheme(CStr(wsSrc.Cells(r, subjectCols(i)).value), subjectSchemeKeys(i))
            rawScore = ""
            If subjectScoreCols(i) > 0 Then rawScore = UCase$(Trim$(CStr(wsSrc.Cells(r, subjectScoreCols(i)).value)))

            isVrMc = (rawGrade = "VR" Or rawScore = "VR" Or rawGrade = "MC" Or rawScore = "MC")

            If gradeStr <> "" Or isVrMc Then
                groupTotalCount = groupTotalCount + 1
                Select Case UCase$(Trim$(subjectSchemeKeys(i)))
                    Case "G1": g1GroupCount = g1GroupCount + 1
                    Case "G2": g2GroupCount = g2GroupCount + 1
                    Case "G3": g3GroupCount = g3GroupCount + 1
                End Select
            End If

        Next i

        If groupTotalCount > 0 Then
            fsbbGroup = ResolveFsbbGroup(g1GroupCount, g2GroupCount, g3GroupCount, groupTotalCount, groupThresholdPct)

            If fsbbGroup = "G1" Or fsbbGroup = "G2" Or fsbbGroup = "G3" Then
                topPrimaryCount = 0
                topSecondaryCount = 0
                For i = 1 To subjCount
                    rawGrade = UCase$(Trim$(CStr(wsSrc.Cells(r, subjectCols(i)).value)))
                    gradeStr = NormalizeGradeForScheme(CStr(wsSrc.Cells(r, subjectCols(i)).value), subjectSchemeKeys(i))
                    rawScore = ""
                    If subjectScoreCols(i) > 0 Then rawScore = UCase$(Trim$(CStr(wsSrc.Cells(r, subjectScoreCols(i)).value)))
                    isVrMc = (rawGrade = "VR" Or rawScore = "VR" Or rawGrade = "MC" Or rawScore = "MC")

                    If Not isVrMc And gradeStr <> "" Then
                        Select Case GetTopBandByDownwardMap(gradeStr, subjectSchemeKeys(i), fsbbGroup)
                            Case 1
                                topPrimaryCount = topPrimaryCount + 1
                            Case 2
                                topSecondaryCount = topSecondaryCount + 1
                        End Select
                    End If
                Next i
                topCount = topPrimaryCount + topSecondaryCount

                recCount = recCount + 1
                If recCount = 1 Then
                    ReDim recs(1 To 1)
                Else
                    ReDim Preserve recs(1 To recCount)
                End If

                recs(recCount).LevelCode = levelCode
                recs(recCount).ClassName = className
                recs(recCount).RegNo = regNo
                recs(recCount).StudentName = studentName
                recs(recCount).GroupCode = fsbbGroup
                recs(recCount).TopCount = topCount
                recs(recCount).TopPrimaryCount = topPrimaryCount
                recs(recCount).TopSecondaryCount = topSecondaryCount
            End If
        End If

NextStudent:
    Next r
    Exit Sub

FailSafe:
    ' Skip broken source sheet
End Sub

Private Function WriteTopGroupSection(ByVal wsOut As Worksheet, _
                                      ByVal startRow As Long, _
                                      ByVal levelCode As String, _
                                      ByVal groupCode As String, _
                                      ByVal topN As Long, _
                                      ByRef recs() As TopStudentRec, _
                                      ByVal recCount As Long) As Long
    Dim idx() As Long
    Dim idxCount As Long
    Dim i As Long, j As Long, tmp As Long
    Dim cutoffTop As Long, cutoffPrimary As Long, cutoffSecondary As Long
    Dim r As Long
    Dim primaryLbl As String, secondaryLbl As String

    wsOut.Cells(startRow, 1).value = groupCode & " Top Students (Top " & topN & ", ties included)"
    wsOut.Cells(startRow, 1).Font.Bold = True
    wsOut.Cells(startRow, 1).Font.Color = RGB(79, 33, 33)
    startRow = startRow + 1

    For i = 1 To recCount
        If UCase$(recs(i).GroupCode) = UCase$(groupCode) Then
            idxCount = idxCount + 1
            If idxCount = 1 Then
                ReDim idx(1 To 1)
            Else
                ReDim Preserve idx(1 To idxCount)
            End If
            idx(idxCount) = i
        End If
    Next i

    If idxCount = 0 Then
        wsOut.Cells(startRow, 1).value = "(No students found for " & groupCode & ")"
        wsOut.Cells(startRow, 1).Font.Italic = True
        WriteTopGroupSection = startRow + 2
        Exit Function
    End If

    ' Sort by TopCount desc, then primary top-band count, then secondary.
    For i = 1 To idxCount - 1
        For j = i + 1 To idxCount
            If recs(idx(j)).TopCount > recs(idx(i)).TopCount _
               Or (recs(idx(j)).TopCount = recs(idx(i)).TopCount And recs(idx(j)).TopPrimaryCount > recs(idx(i)).TopPrimaryCount) _
               Or (recs(idx(j)).TopCount = recs(idx(i)).TopCount And recs(idx(j)).TopPrimaryCount = recs(idx(i)).TopPrimaryCount _
                   And recs(idx(j)).TopSecondaryCount > recs(idx(i)).TopSecondaryCount) Then
                tmp = idx(i)
                idx(i) = idx(j)
                idx(j) = tmp
            End If
        Next j
    Next i

    GetTopBandLabels groupCode, primaryLbl, secondaryLbl

    wsOut.Cells(startRow, 1).value = "Level"
    wsOut.Cells(startRow, 2).value = "Class"
    wsOut.Cells(startRow, 3).value = "RegNo"
    wsOut.Cells(startRow, 4).value = "Name"
    wsOut.Cells(startRow, 5).value = "Group"
    wsOut.Cells(startRow, 6).NumberFormat = "@"
    wsOut.Cells(startRow, 7).NumberFormat = "@"
    wsOut.Cells(startRow, 8).NumberFormat = "@"
    wsOut.Cells(startRow, 6).value = primaryLbl & "/" & secondaryLbl
    wsOut.Cells(startRow, 7).value = primaryLbl
    wsOut.Cells(startRow, 8).value = secondaryLbl
    wsOut.Range(wsOut.Cells(startRow, 1), wsOut.Cells(startRow, 8)).Font.Bold = True
    startRow = startRow + 1

    If idxCount <= topN Then
        cutoffTop = recs(idx(idxCount)).TopCount
        cutoffPrimary = recs(idx(idxCount)).TopPrimaryCount
        cutoffSecondary = recs(idx(idxCount)).TopSecondaryCount
    Else
        cutoffTop = recs(idx(topN)).TopCount
        cutoffPrimary = recs(idx(topN)).TopPrimaryCount
        cutoffSecondary = recs(idx(topN)).TopSecondaryCount
    End If

    r = startRow
    For i = 1 To idxCount
        If recs(idx(i)).TopCount < cutoffTop Then Exit For
        If recs(idx(i)).TopCount = cutoffTop Then
            If recs(idx(i)).TopPrimaryCount < cutoffPrimary Then Exit For
            If recs(idx(i)).TopPrimaryCount = cutoffPrimary And recs(idx(i)).TopSecondaryCount < cutoffSecondary Then Exit For
        End If
        wsOut.Cells(r, 1).value = levelCode
        wsOut.Cells(r, 2).value = recs(idx(i)).ClassName
        wsOut.Cells(r, 3).value = recs(idx(i)).RegNo
        wsOut.Cells(r, 4).value = recs(idx(i)).StudentName
        wsOut.Cells(r, 5).value = recs(idx(i)).GroupCode
        wsOut.Cells(r, 6).value = recs(idx(i)).TopCount
        wsOut.Cells(r, 7).value = recs(idx(i)).TopPrimaryCount
        wsOut.Cells(r, 8).value = recs(idx(i)).TopSecondaryCount
        r = r + 1
    Next i

    WriteTopGroupSection = r + 1
End Function

Private Sub PrepareTopQualitySheet(ByVal wsOut As Worksheet, ByVal levelCode As String)
    Dim explainer As String

    wsOut.Range("A1").value = levelCode & " Top Students by Top Grades"
    wsOut.Range("A1").Font.Bold = True
    wsOut.Range("A1").Font.Size = 14

    explainer = "How ranking works: first by total top grades, then by the first top-grade column, " & _
                "then by the second top-grade column. G3 top 20; G2/G1 top 10; ties included." & vbLf & _
                "Columns used: G3 uses A1/A2 (A1 then A2); G2 uses 1/2 (1 then 2); " & _
                "G1 uses A/B (A then B)." & vbLf & _
                "Downward conversion for mixed-level subjects: " & _
                "G3->G2: A1/A2/B3=>1, B4/C5/C6=>2. " & _
                "G2->G1: 1/2/3=>A, 4=>B. " & _
                "G3->G1: A1/A2/B3/B4/C5/C6/D7=>A, E8=>B."

    With wsOut.Range("A2:H2")
        .Merge
        .value = explainer
        .WrapText = True
        .Font.Italic = True
        .VerticalAlignment = xlTop
    End With
    wsOut.Rows(2).RowHeight = 60
End Sub

Private Sub FormatTopQualitySheet(ByVal wsOut As Worksheet, ByVal lastRow As Long)
    Dim rngTable As Range

    wsOut.Columns("A:H").AutoFit
    wsOut.Columns("A").ColumnWidth = 8
    wsOut.Columns("B").ColumnWidth = 12
    wsOut.Columns("C").ColumnWidth = 5
    wsOut.Columns("D").ColumnWidth = 24
    wsOut.Columns("E").ColumnWidth = 8
    wsOut.Columns("F").ColumnWidth = 10
    wsOut.Columns("G").ColumnWidth = 10
    wsOut.Columns("H").ColumnWidth = 10
    wsOut.Columns("C").HorizontalAlignment = xlCenter
    wsOut.Columns("E:H").HorizontalAlignment = xlCenter

    If lastRow >= 4 Then
        Set rngTable = wsOut.Range("A4:H" & lastRow)
        With rngTable.Borders
            .LineStyle = xlContinuous
            .Color = RGB(200, 200, 200)
            .Weight = xlThin
        End With
    End If
End Sub

Private Sub GetTopBandLabels(ByVal groupCode As String, ByRef primaryLbl As String, ByRef secondaryLbl As String)
    Select Case UCase$(Trim$(groupCode))
        Case "G3"
            primaryLbl = "A1"
            secondaryLbl = "A2"
        Case "G2"
            primaryLbl = "1"
            secondaryLbl = "2"
        Case "G1"
            primaryLbl = "A"
            secondaryLbl = "B"
        Case Else
            primaryLbl = "Top1"
            secondaryLbl = "Top2"
    End Select
End Sub

Private Function GetTopBandByDownwardMap(ByVal gradeStr As String, _
                                         ByVal sourceScheme As String, _
                                         ByVal targetGroup As String) As Long
    Dim g As String, src As String, tgt As String
    g = UCase$(Trim$(gradeStr))
    src = UCase$(Trim$(sourceScheme))
    tgt = UCase$(Trim$(targetGroup))

    Select Case tgt
        Case "G3"
            If src = "G3" Then
                If g = "A1" Then
                    GetTopBandByDownwardMap = 1
                ElseIf g = "A2" Then
                    GetTopBandByDownwardMap = 2
                End If
            End If

        Case "G2"
            If src = "G2" Then
                If g = "1" Then
                    GetTopBandByDownwardMap = 1
                ElseIf g = "2" Then
                    GetTopBandByDownwardMap = 2
                End If
            ElseIf src = "G3" Then
                ' G3 -> G2 mapping: A1/A2/B3->1 ; B4/C5/C6->2 ; D7->3 ; E8->4 ; 9/F9->5
                If g = "A1" Or g = "A2" Or g = "B3" Then
                    GetTopBandByDownwardMap = 1
                ElseIf g = "B4" Or g = "C5" Or g = "C6" Then
                    GetTopBandByDownwardMap = 2
                End If
            End If

        Case "G1"
            If src = "G1" Then
                If g = "A" Then
                    GetTopBandByDownwardMap = 1
                ElseIf g = "B" Then
                    GetTopBandByDownwardMap = 2
                End If
            ElseIf src = "G2" Then
                ' G2 -> G1 mapping: 1/2/3->A ; 4->B ; 5->C ; 6->D
                If g = "1" Or g = "2" Or g = "3" Then
                    GetTopBandByDownwardMap = 1
                ElseIf g = "4" Then
                    GetTopBandByDownwardMap = 2
                End If
            ElseIf src = "G3" Then
                ' G3 -> G1 mapping: A1/A2/B3/B4/C5/C6/D7->A ; E8->B ; 9/F9->C
                If g = "A1" Or g = "A2" Or g = "B3" Or g = "B4" _
                   Or g = "C5" Or g = "C6" Or g = "D7" Then
                    GetTopBandByDownwardMap = 1
                ElseIf g = "E8" Then
                    GetTopBandByDownwardMap = 2
                End If
            End If
    End Select
End Function

Private Function AppendSecAtRiskFromSourceSheet(ByVal wsSrc As Worksheet, _
                                                ByVal wsOut As Worksheet, _
                                                ByVal startOutRow As Long, _
                                                ByVal atRiskFailThreshold As Long, _
                                                ByVal targetLevel As String) As Long
    Dim classCol As Long, nameCol As Long, regCol As Long
    Dim lastRow As Long, lastCol As Long
    Dim firstClass As String, levelCode As String
    Dim subjectCols() As Long
    Dim subjectNames() As String
    Dim subjectSchemeKeys() As String
    Dim subjectScoreCols() As Long
    Dim subjCount As Long
    Dim c As Long, r As Long, i As Long
    Dim header As String, schemeKey As String
    Dim className As String, studentName As String, regNo As String
    Dim gradeStr As String
    Dim attemptedCount As Long, passCount As Long, failCount As Long
    Dim outRow As Long
    Dim riskBand As String, failedSubjects As String, attemptedSubjects As String
    Dim vrSubjects As String, rawGrade As String, rawScore As String
    Dim subjectName As String
    Dim isVrSubject As Boolean
    Dim g1Taken As Long, g2Taken As Long, g3Taken As Long
    Dim g1GroupCount As Long, g2GroupCount As Long, g3GroupCount As Long
    Dim groupTotalCount As Long
    Dim fsbbGroup As String
    Dim groupThresholdPct As Double

    On Error GoTo FailSafe

    If LCase$(wsSrc.Name) Like "*settings*" _
       Or LCase$(wsSrc.Name) Like "*config*" _
       Or LCase$(wsSrc.Name) Like "*menu*" _
       Or LCase$(wsSrc.Name) Like "*lookup*" _
       Or LCase$(wsSrc.Name) Like "*summary*" _
       Or LCase$(wsSrc.Name) Like "*template*" _
       Or InStr(1, LCase$(wsSrc.Name), "_subj analysis_") > 0 _
       Or InStr(1, LCase$(wsSrc.Name), "dashboard") > 0 Then
        Exit Function
    End If

    classCol = FindHeaderColumn(wsSrc, 1, "Class")
    If classCol = 0 Then Exit Function

    nameCol = FindFirstHeaderColumn(wsSrc, 1, Array("Name", "Student Name", "Student"))
    regCol = FindFirstHeaderColumn(wsSrc, 1, Array("RegNo", "Reg No", "Register No", "Index No", "Adm No"))

    lastRow = wsSrc.Cells(wsSrc.Rows.count, classCol).End(xlUp).Row
    firstClass = ""
    For r = 2 To lastRow
        firstClass = Trim$(CStr(wsSrc.Cells(r, classCol).value))
        If firstClass <> "" Then Exit For
    Next r
    If firstClass = "" Then Exit Function

    levelCode = InferLevelCodeFromClass(firstClass)
    If levelCode = "" Then Exit Function
    If UCase$(levelCode) <> UCase$(Trim$(targetLevel)) Then Exit Function

    lastCol = wsSrc.Cells(1, wsSrc.Columns.count).End(xlToLeft).Column
    subjCount = 0

    For c = 1 To lastCol
        If c <> classCol Then
            header = Trim$(CStr(wsSrc.Cells(1, c).value))
            If header <> "" And IsLikelySubjectGradeColumn(header) Then
                schemeKey = GetGradeSchemeKey(wsSrc, c, header)
                subjectName = StripGradeHeaderSuffix(header)
                If schemeKey <> "" And Not SubjectAlreadyAdded(subjectNames, subjCount, subjectName) Then
                    subjCount = subjCount + 1
                    ReDim Preserve subjectCols(1 To subjCount)
                    ReDim Preserve subjectNames(1 To subjCount)
                    ReDim Preserve subjectSchemeKeys(1 To subjCount)
                    ReDim Preserve subjectScoreCols(1 To subjCount)
                    subjectCols(subjCount) = c
                    subjectNames(subjCount) = subjectName
                    subjectSchemeKeys(subjCount) = schemeKey
                    subjectScoreCols(subjCount) = FindScoreColumnForSubject(wsSrc, 1, subjectName)
                End If
            End If
        End If
    Next c

    If subjCount = 0 Then Exit Function

    groupThresholdPct = GetGroupThresholdPercent()

    outRow = startOutRow
    For r = 2 To lastRow
        className = Trim$(CStr(wsSrc.Cells(r, classCol).value))
        If className = "" Then GoTo NextStudent

        If UCase$(Left$(className, 1)) = "Y" Then GoTo NextStudent

        If nameCol > 0 Then
            studentName = Trim$(CStr(wsSrc.Cells(r, nameCol).value))
        Else
            studentName = ""
        End If

        If regCol > 0 Then
            regNo = Trim$(CStr(wsSrc.Cells(r, regCol).value))
        Else
            regNo = ""
        End If

        attemptedCount = 0
        passCount = 0
        failCount = 0
        failedSubjects = ""
        attemptedSubjects = ""
        vrSubjects = ""
        g1Taken = 0
        g2Taken = 0
        g3Taken = 0
        g1GroupCount = 0
        g2GroupCount = 0
        g3GroupCount = 0
        groupTotalCount = 0

        For i = 1 To subjCount
            rawGrade = UCase$(Trim$(CStr(wsSrc.Cells(r, subjectCols(i)).value)))
            gradeStr = NormalizeGradeForScheme(CStr(wsSrc.Cells(r, subjectCols(i)).value), subjectSchemeKeys(i))
            rawScore = ""
            If subjectScoreCols(i) > 0 Then
                rawScore = UCase$(Trim$(CStr(wsSrc.Cells(r, subjectScoreCols(i)).value)))
            End If

            ' Group base includes attempted subjects and those marked VR/MC.
            If gradeStr <> "" Or rawGrade = "VR" Or rawScore = "VR" _
               Or rawGrade = "MC" Or rawScore = "MC" Then
                groupTotalCount = groupTotalCount + 1
                Select Case UCase$(Trim$(subjectSchemeKeys(i)))
                    Case "G1": g1GroupCount = g1GroupCount + 1
                    Case "G2": g2GroupCount = g2GroupCount + 1
                    Case "G3": g3GroupCount = g3GroupCount + 1
                End Select
            End If

            isVrSubject = (rawGrade = "VR" Or rawScore = "VR")
            If isVrSubject Then
                If vrSubjects <> "" Then vrSubjects = vrSubjects & ", "
                vrSubjects = vrSubjects & subjectNames(i)
                GoTo NextSubject
            End If

            If gradeStr <> "" Then
                attemptedCount = attemptedCount + 1
                Select Case UCase$(Trim$(subjectSchemeKeys(i)))
                    Case "G1": g1Taken = g1Taken + 1
                    Case "G2": g2Taken = g2Taken + 1
                    Case "G3": g3Taken = g3Taken + 1
                End Select
                If attemptedSubjects <> "" Then attemptedSubjects = attemptedSubjects & ", "
                attemptedSubjects = attemptedSubjects & subjectNames(i)
                If IsFailGradeByScheme(gradeStr, subjectSchemeKeys(i)) Then
                    failCount = failCount + 1
                    If failedSubjects <> "" Then failedSubjects = failedSubjects & ", "
                    failedSubjects = failedSubjects & subjectNames(i)
                Else
                    passCount = passCount + 1
                End If
            End If
NextSubject:
        Next i

        If attemptedCount > 0 Or groupTotalCount > 0 Then
            If failCount >= atRiskFailThreshold Then
                riskBand = "AT RISK"
            ElseIf failCount >= 1 Then
                riskBand = "MONITOR"
            Else
                riskBand = "OK"
            End If

            wsOut.Cells(outRow, 1).value = levelCode
            wsOut.Cells(outRow, 2).value = className
            wsOut.Cells(outRow, 3).value = regNo
            wsOut.Cells(outRow, 4).value = studentName
            fsbbGroup = ResolveFsbbGroup(g1GroupCount, g2GroupCount, g3GroupCount, groupTotalCount, groupThresholdPct)
            wsOut.Cells(outRow, 5).value = attemptedCount
            wsOut.Cells(outRow, 6).value = passCount
            wsOut.Cells(outRow, 7).value = failCount
            wsOut.Cells(outRow, 8).value = failedSubjects
            wsOut.Cells(outRow, 9).value = riskBand
            wsOut.Cells(outRow, 10).value = RiskBandRank(riskBand)
            wsOut.Cells(outRow, 12).value = attemptedSubjects
            wsOut.Cells(outRow, 13).value = vrSubjects
            wsOut.Cells(outRow, 14).value = fsbbGroup

            If riskBand = "AT RISK" Then
                wsOut.Range(wsOut.Cells(outRow, 1), wsOut.Cells(outRow, 9)).Interior.Color = RGB(255, 230, 230)
                wsOut.Cells(outRow, 9).Font.Color = RGB(192, 0, 0)
                wsOut.Cells(outRow, 9).Font.Bold = True
            ElseIf riskBand = "MONITOR" Then
                wsOut.Cells(outRow, 9).Font.Color = RGB(156, 101, 0)
            Else
                wsOut.Cells(outRow, 9).Font.Color = RGB(0, 97, 0)
            End If

            outRow = outRow + 1
        End If

NextStudent:
    Next r

    AppendSecAtRiskFromSourceSheet = outRow - startOutRow
    Exit Function

FailSafe:
    AppendSecAtRiskFromSourceSheet = 0
End Function

Public Sub BuildSec_AtRiskNavigation()
    Dim wsNav As Worksheet
    Dim startCell As Range
    Dim startRow As Long, startCol As Long
    Dim rowPtr As Long
    Dim lvl As Variant
    Dim sheetName As String
    Dim shp As Shape
    Dim k As Long

    On Error GoTo ErrHandler

    On Error Resume Next
    Set wsNav = ThisWorkbook.Worksheets(ATRISK_NAV_SHEET_NAME)
    On Error GoTo ErrHandler
    If wsNav Is Nothing Then Exit Sub

    Set startCell = wsNav.Range(ATRISK_NAV_START_CELL)
    startRow = startCell.Row
    startCol = startCell.Column

    wsNav.Range(wsNav.Cells(startRow, startCol), wsNav.Cells(startRow + 120, startCol + 5)).Clear
    For k = wsNav.Shapes.count To 1 Step -1
        Set shp = wsNav.Shapes(k)
        If Left$(shp.Name, Len(ATRISK_NAV_BTN_PREFIX)) = ATRISK_NAV_BTN_PREFIX Then
            shp.Delete
        End If
    Next k

    wsNav.Cells(startRow, startCol).value = "SEC At-Risk Menu"
    wsNav.Cells(startRow, startCol).Font.Bold = True
    wsNav.Cells(startRow, startCol).Font.Size = 12
    wsNav.Cells(startRow, startCol).Font.Color = RGB(156, 0, 6)
    rowPtr = startRow + 1

    For Each lvl In Array("S1", "S2", "S3", "S4", "S5")
        sheetName = "AtRisk_" & CStr(lvl)
        If WorksheetExistsByName(sheetName) Then
            CreateAtRiskNavButton wsNav, sheetName, CStr(lvl) & " At Risk", rowPtr, startCol
        Else
            wsNav.Cells(rowPtr, startCol).value = CStr(lvl) & " At Risk (not built)"
            wsNav.Cells(rowPtr, startCol).Font.Italic = True
        End If
        rowPtr = rowPtr + 2
    Next lvl

    Exit Sub

ErrHandler:
    ' Silent fallback
End Sub

Public Sub BuildAllAtRiskHomeButtons()
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        If Left$(ws.Name, 7) = "AtRisk_" Then
            AddAtRiskHomeButton ws
        End If
    Next ws
End Sub

Private Sub AddAtRiskHomeButton(ByVal ws As Worksheet)
    Dim shp As Shape
    Dim tgtCell As Range
    Dim leftPos As Double, topPos As Double
    Dim btnWidth As Double, btnHeight As Double

    Set tgtCell = ws.Range("E1")
    leftPos = tgtCell.Left
    topPos = tgtCell.Top
    btnWidth = tgtCell.Width * 1.2
    btnHeight = tgtCell.Height * 1.2

    On Error Resume Next
    ws.Shapes("HomeBtn").Delete
    On Error GoTo 0

    Set shp = ws.Shapes.AddShape( _
        Type:=SHAPE_ROUNDED_RECTANGLE, _
        Left:=leftPos, _
        Top:=topPos, _
        Width:=btnWidth, _
        Height:=btnHeight)

    With shp
        .Name = "HomeBtn"
        .Fill.ForeColor.RGB = RGB(244, 204, 204)
        .line.ForeColor.RGB = RGB(192, 80, 77)
        .line.Weight = 1.5
        With .TextFrame2
            .TextRange.text = "Home"
            .TextRange.Font.Name = "Calibri"
            .TextRange.Font.Size = 11
            .TextRange.Font.Fill.ForeColor.RGB = RGB(156, 0, 6)
            .VerticalAnchor = msoAnchorMiddle
            .TextRange.ParagraphFormat.Alignment = msoAlignCenter
            .MarginLeft = 4
            .MarginRight = 4
        End With
    End With

    ws.Hyperlinks.Add Anchor:=shp, Address:="", SubAddress:="'Dashboard'!A1"
End Sub

Private Sub CreateAtRiskNavButton(ByVal wsNav As Worksheet, _
                                  ByVal targetSheetName As String, _
                                  ByVal displayText As String, _
                                  ByVal rowNum As Long, _
                                  ByVal firstCol As Long)
    Dim shp As Shape
    Dim leftPos As Double, topPos As Double
    Dim btnWidth As Double, btnHeight As Double

    leftPos = wsNav.Cells(rowNum, firstCol).Left
    topPos = wsNav.Cells(rowNum, firstCol).Top
    btnWidth = wsNav.Columns(firstCol).Resize(, 5).Width
    btnHeight = wsNav.Rows(rowNum).Height * 1.3

    Set shp = wsNav.Shapes.AddShape( _
        Type:=SHAPE_ROUNDED_RECTANGLE, _
        Left:=leftPos, _
        Top:=topPos, _
        Width:=btnWidth, _
        Height:=btnHeight)

    With shp
        .Name = ATRISK_NAV_BTN_PREFIX & targetSheetName
        .Fill.ForeColor.RGB = RGB(244, 204, 204)
        .Fill.Transparency = 0#
        .line.ForeColor.RGB = RGB(192, 80, 77)
        .line.Weight = 1.5
        With .TextFrame2
            .TextRange.text = displayText
            .TextRange.Font.Name = "Calibri"
            .TextRange.Font.Size = 10.5
            .TextRange.Font.Fill.ForeColor.RGB = RGB(156, 0, 6)
            .TextRange.ParagraphFormat.Alignment = msoAlignCenter
            .VerticalAnchor = msoAnchorMiddle
            .MarginLeft = 6
            .MarginRight = 6
            .MarginTop = 3
            .MarginBottom = 3
        End With
    End With

    wsNav.Hyperlinks.Add Anchor:=shp, Address:="", SubAddress:="'" & targetSheetName & "'!A1"
End Sub

Private Function WorksheetExistsByName(ByVal sheetName As String) As Boolean
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0
    WorksheetExistsByName = Not ws Is Nothing
End Function

Private Function GetOrCreateWorksheet(ByVal sheetName As String) As Worksheet
    Dim ws As Worksheet

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0

    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.count))
        ws.Name = sheetName
    Else
        ws.Cells.Clear
    End If

    Set GetOrCreateWorksheet = ws
End Function

Private Sub PrepareAtRiskSheet(ByVal wsOut As Worksheet, ByVal levelCode As String, ByVal threshold As Long)
    wsOut.Range("A1").value = levelCode & " Students At Risk"
    wsOut.Range("A1").Font.Bold = True
    wsOut.Range("A1").Font.Size = 14

    wsOut.Range("A2").value = "At-risk rule: Failed Subjects >= " & threshold & " (VR excluded)"
    wsOut.Range("A2").Font.Italic = True

    wsOut.Cells(4, 1).value = "Level"
    wsOut.Cells(4, 2).value = "Class"
    wsOut.Cells(4, 3).value = "RegNo"
    wsOut.Cells(4, 4).value = "Name"
    wsOut.Cells(4, 5).value = "Subjects Attempted"
    wsOut.Cells(4, 6).value = "Subjects Passed"
    wsOut.Cells(4, 7).value = "Subjects Failed"
    wsOut.Cells(4, 8).value = "Failed Subjects"
    wsOut.Cells(4, 9).value = "Risk Band"
    wsOut.Cells(4, 10).value = "SortKey"
    wsOut.Cells(4, 12).value = "Attempted Subjects"
    wsOut.Cells(4, 13).value = "VR Subjects"
    wsOut.Cells(4, 14).value = "Group"
    wsOut.Rows(4).Font.Bold = True
End Sub

Private Sub FinalizeAtRiskSheet(ByVal wsOut As Worksheet, ByVal lastRow As Long)
    Dim sortRange As Range
    Dim rngTable As Range

    If lastRow >= 5 Then
        Set sortRange = wsOut.Range("A4:N" & lastRow)
        sortRange.Sort Key1:=wsOut.Range("J5"), Order1:=xlAscending, _
                       Key2:=wsOut.Range("G5"), Order2:=xlDescending, _
                       Key3:=wsOut.Range("D5"), Order3:=xlAscending, _
                       Header:=xlYes
    End If

    wsOut.Columns("A:N").AutoFit
    wsOut.Columns("A").ColumnWidth = 8
    wsOut.Columns("B").ColumnWidth = 15
    wsOut.Columns("C").ColumnWidth = 5
    wsOut.Columns("D").ColumnWidth = 24
    wsOut.Columns("E").ColumnWidth = 10
    wsOut.Columns("F").ColumnWidth = 9
    wsOut.Columns("G").ColumnWidth = 9
    wsOut.Columns("E:G").HorizontalAlignment = xlCenter
    wsOut.Columns("H").ColumnWidth = 40
    wsOut.Columns("H").WrapText = True
    wsOut.Columns("K").ColumnWidth = 10
    wsOut.Columns("L").ColumnWidth = 40
    wsOut.Columns("L").WrapText = True
    wsOut.Columns("M").ColumnWidth = 15
    wsOut.Columns("M").WrapText = True
    wsOut.Columns("N").ColumnWidth = 10
    wsOut.Columns("N").HorizontalAlignment = xlCenter
    wsOut.Columns("J").EntireColumn.Hidden = True
    wsOut.Range("E4:G4").WrapText = True
    wsOut.Range("A4:N4").VerticalAlignment = xlCenter

    If lastRow >= 4 Then
        Set rngTable = wsOut.Range("A4:N" & lastRow)
        With rngTable.Borders
            .LineStyle = xlContinuous
            .Color = RGB(200, 200, 200)
            .Weight = xlThin
        End With
    End If
End Sub

'---------------------------------------------------------
' ENGINE - BUILD ONE SUBJECT TABLE + CHART
'---------------------------------------------------------
Public Sub BuildSecSubjectGradeDistribution( _
    ByVal srcSheetName As String, _
    ByVal srcClassCol As Long, _
    ByVal srcGradeCol As Long, _
    ByVal destSheetName As String, _
    ByVal destTopLeft As String, _
    ByVal subjectTitle As String, _
    Optional ByVal schemeKey As String = "G3", _
    Optional ByRef outEndRow As Long = 0)

    Dim wb As Workbook
    Dim wsSrc As Worksheet, wsDest As Worksheet
    Dim lastRow As Long, r As Long
    Dim className As String, gradeStr As String

    Dim gradeLabels() As String
    Dim numBands As Long
    Dim passMaxIdx As Long, failMinIdx As Long, topMaxIdx As Long
    Dim pctPassLabel As String, pctFailLabel As String, pctTopLabel As String, meanLabel As String

    Dim countsArr() As Long
    Dim totalArr() As Long

    Dim classList() As String
    Dim sortedClassList() As String
    Dim classCounts() As Long   ' (gradeBand, classIndex)
    Dim classCount As Long
    Dim classNameKey As String
    Dim classIdx As Long
    Dim subjectTotalN As Long
    Dim minSubjectN As Long
    Dim i As Long, j As Long

    Dim destRowHeader As Long, destColFirst As Long
    Dim rowPtr As Long, cohortRow As Long

    Dim total As Long
    Dim passCount As Long, failCount As Long, topCount As Long
    Dim meanValue As Double

    Dim colNo As Long, colPctPass As Long, colPctFail As Long, colPctTop As Long, colMean As Long

    Dim rngTable As Range
    Dim rngHeader As Range, rngData As Range
    Dim rngCohortRow As Range

    Dim co As ChartObject
    Dim ch As Chart
    Dim leftPos As Double, topPos As Double, chartWidth As Double, chartHeight As Double

    Dim pastelColors(1 To 9) As Long
    Dim s As Series, pt As Point

    Dim titleCell As Range
    Dim validityFlag As String, patternType As String
    Dim line1 As String, line2 As String, line3 As String

    outEndRow = 0
    On Error GoTo ErrHandler

    Set wb = ThisWorkbook
    Set wsSrc = wb.Worksheets(srcSheetName)
    Set wsDest = wb.Worksheets(destSheetName)

    If Not InitGradeScheme(schemeKey, gradeLabels, passMaxIdx, failMinIdx, topMaxIdx, _
                           pctPassLabel, pctFailLabel, pctTopLabel, meanLabel) Then
        Exit Sub
    End If

    numBands = UBound(gradeLabels)
    InitPastelPalette pastelColors

    lastRow = wsSrc.Cells(wsSrc.Rows.count, srcClassCol).End(xlUp).Row

    For r = 2 To lastRow
        className = Trim$(CStr(wsSrc.Cells(r, srcClassCol).value))
        gradeStr = NormalizeGradeForScheme(CStr(wsSrc.Cells(r, srcGradeCol).value), schemeKey)

        If className <> "" And gradeStr <> "" Then
            ' Legacy safeguard: keep excluding Y-track classes if present.
            If UCase$(Left$(className, 1)) <> "Y" Then
                j = GradeIndexByScheme(gradeStr, gradeLabels)
                If j >= 1 And j <= numBands Then
                    classIdx = FindClassIndex(classList, classCount, className)
                    If classIdx = 0 Then
                        classCount = classCount + 1

                        If classCount = 1 Then
                            ReDim classList(1 To 1)
                            ReDim classCounts(1 To numBands, 1 To 1)
                        Else
                            ReDim Preserve classList(1 To classCount)
                            ReDim Preserve classCounts(1 To numBands, 1 To classCount)
                        End If

                        classList(classCount) = className
                        classIdx = classCount
                    End If

                    classCounts(j, classIdx) = classCounts(j, classIdx) + 1
                End If
            End If
        End If
    Next r

    If classCount = 0 Then Exit Sub

    minSubjectN = GetMinSubjectN()
    subjectTotalN = 0
    For classIdx = 1 To classCount
        For j = 1 To numBands
            subjectTotalN = subjectTotalN + classCounts(j, classIdx)
        Next j
    Next classIdx
    If subjectTotalN < minSubjectN Then Exit Sub

    ReDim sortedClassList(1 To classCount)
    For i = 1 To classCount
        sortedClassList(i) = classList(i)
    Next i
    SortStringArray sortedClassList

    destRowHeader = wsDest.Range(destTopLeft).Row
    destColFirst = wsDest.Range(destTopLeft).Column

    Set titleCell = wsDest.Cells(destRowHeader - 1, destColFirst)
    With titleCell
        .value = subjectTitle & " [" & UCase$(schemeKey) & "]"
        .Font.Bold = True
        .Font.Size = 11
    End With

    colNo = destColFirst + numBands + 1
    colPctPass = destColFirst + numBands + 2
    colPctFail = destColFirst + numBands + 3
    colPctTop = destColFirst + numBands + 4
    colMean = destColFirst + numBands + 5

    With wsDest
        .Cells(destRowHeader, destColFirst).value = "Class"

        For j = 1 To numBands
            .Cells(destRowHeader, destColFirst + j).value = gradeLabels(j)
        Next j

        .Cells(destRowHeader, colNo).value = "No."
        .Cells(destRowHeader, colPctPass).value = pctPassLabel
        .Cells(destRowHeader, colPctFail).value = pctFailLabel
        .Cells(destRowHeader, colPctTop).value = pctTopLabel
        .Cells(destRowHeader, colMean).value = meanLabel
        .Rows(destRowHeader).Font.Bold = True
    End With

    ReDim totalArr(1 To numBands)
    rowPtr = destRowHeader + 1

    For i = LBound(sortedClassList) To UBound(sortedClassList)
        classNameKey = sortedClassList(i)
        classIdx = FindClassIndex(classList, classCount, classNameKey)
        ReDim countsArr(1 To numBands)
        For j = 1 To numBands
            countsArr(j) = classCounts(j, classIdx)
        Next j

        total = 0
        passCount = 0
        failCount = 0
        topCount = 0

        For j = 1 To numBands
            total = total + countsArr(j)
            totalArr(j) = totalArr(j) + countsArr(j)

            If j <= passMaxIdx Then passCount = passCount + countsArr(j)
            If j >= failMinIdx Then failCount = failCount + countsArr(j)
            If j <= topMaxIdx Then topCount = topCount + countsArr(j)
        Next j

        If total > 0 Then
            meanValue = ComputeMeanBand(countsArr)
        Else
            meanValue = 0
        End If

        With wsDest
            .Cells(rowPtr, destColFirst).value = classNameKey

            For j = 1 To numBands
                .Cells(rowPtr, destColFirst + j).value = countsArr(j)
            Next j

            .Cells(rowPtr, colNo).value = total

            If total > 0 Then
                .Cells(rowPtr, colPctPass).value = Round(passCount * 100# / total, 1)
                .Cells(rowPtr, colPctFail).value = Round(failCount * 100# / total, 1)
                .Cells(rowPtr, colPctTop).value = Round(topCount * 100# / total, 1)
                .Cells(rowPtr, colMean).value = Round(meanValue, 1)
            Else
                .Cells(rowPtr, colPctPass).ClearContents
                .Cells(rowPtr, colPctFail).ClearContents
                .Cells(rowPtr, colPctTop).ClearContents
                .Cells(rowPtr, colMean).ClearContents
            End If
        End With

        ColourSubjectRow wsDest, rowPtr, destColFirst, numBands, topMaxIdx, failMinIdx, _
                        colPctPass, colPctFail, colPctTop, colMean

        rowPtr = rowPtr + 1
    Next i

    cohortRow = rowPtr
    total = 0
    passCount = 0
    failCount = 0
    topCount = 0

    For j = 1 To numBands
        total = total + totalArr(j)

        If j <= passMaxIdx Then passCount = passCount + totalArr(j)
        If j >= failMinIdx Then failCount = failCount + totalArr(j)
        If j <= topMaxIdx Then topCount = topCount + totalArr(j)
    Next j

    If total > 0 Then
        ReDim countsArr(1 To numBands)
        For j = 1 To numBands
            countsArr(j) = totalArr(j)
        Next j
        meanValue = ComputeMeanBand(countsArr)
    Else
        meanValue = 0
    End If

    With wsDest
        .Cells(cohortRow, destColFirst).value = "COHORT"

        For j = 1 To numBands
            .Cells(cohortRow, destColFirst + j).value = totalArr(j)
        Next j

        .Cells(cohortRow, colNo).value = total

        If total > 0 Then
            .Cells(cohortRow, colPctPass).value = Round(passCount * 100# / total, 1)
            .Cells(cohortRow, colPctFail).value = Round(failCount * 100# / total, 1)
            .Cells(cohortRow, colPctTop).value = Round(topCount * 100# / total, 1)
            .Cells(cohortRow, colMean).value = Round(meanValue, 1)
        Else
            .Cells(cohortRow, colPctPass).ClearContents
            .Cells(cohortRow, colPctFail).ClearContents
            .Cells(cohortRow, colPctTop).ClearContents
            .Cells(cohortRow, colMean).ClearContents
        End If
    End With

    ColourSubjectRow wsDest, cohortRow, destColFirst, numBands, topMaxIdx, failMinIdx, _
                    colPctPass, colPctFail, colPctTop, colMean

    Set rngTable = wsDest.Range(wsDest.Cells(destRowHeader, destColFirst), _
                                wsDest.Cells(cohortRow, colMean))

    With rngTable.Borders
        .LineStyle = xlContinuous
        .Color = RGB(200, 200, 200)
        .Weight = xlThin
    End With

    wsDest.Range(wsDest.Cells(destRowHeader + 1, colPctPass), _
                 wsDest.Cells(cohortRow, colPctTop)).NumberFormat = "0.0"
    wsDest.Range(wsDest.Cells(destRowHeader + 1, colMean), _
                 wsDest.Cells(cohortRow, colMean)).NumberFormat = "0.0"

    wsDest.Columns(destColFirst + 1).Resize(, colMean - destColFirst).AutoFit
    wsDest.Columns(destColFirst).ColumnWidth = 15

    Set rngCohortRow = wsDest.Range(wsDest.Cells(cohortRow, destColFirst), _
                                    wsDest.Cells(cohortRow, colMean))
    rngCohortRow.Interior.Color = RGB(255, 242, 204)
    rngCohortRow.Font.Bold = True

    Set rngHeader = wsDest.Range(wsDest.Cells(destRowHeader, destColFirst + 1), _
                                 wsDest.Cells(destRowHeader, destColFirst + numBands))
    Set rngData = wsDest.Range(wsDest.Cells(cohortRow, destColFirst + 1), _
                               wsDest.Cells(cohortRow, destColFirst + numBands))

    leftPos = wsDest.Columns(colMean + 2).Left
    topPos = wsDest.Rows(destRowHeader - 1).Top
    chartWidth = wsDest.Columns(colMean + 2).Resize(, 6).Width
    chartHeight = wsDest.Rows(destRowHeader - 1).Resize(cohortRow - destRowHeader + 3).Height

    Set co = wsDest.ChartObjects.Add(leftPos, topPos, chartWidth, chartHeight)
    Set ch = co.Chart

    With ch
        .ChartType = xlColumnClustered
        .HasTitle = False
        .SetSourceData Source:=rngData
        .SeriesCollection(1).XValues = rngHeader
        .Legend.Delete

        On Error Resume Next
        .Axes(xlValue).HasMajorGridlines = False
        .Axes(xlCategory).HasMajorGridlines = False
        On Error GoTo ErrHandler

        .ChartArea.Format.line.Visible = msoFalse
        .PlotArea.Format.line.Visible = msoFalse
        .ChartArea.Format.Fill.ForeColor.RGB = RGB(255, 255, 255)
        .PlotArea.Format.Fill.Visible = msoFalse

        .SeriesCollection(1).HasDataLabels = True

        Set s = .SeriesCollection(1)
        For j = 1 To numBands
            Set pt = s.Points(j)
            pt.Format.Fill.ForeColor.RGB = pastelColors(j)
            pt.Format.Fill.Solid
        Next j

        .ChartGroups(1).GapWidth = 30
    End With

    ' Validity panel (scheme-aware)
    EvaluateDistributionForScheme totalArr, total, schemeKey, validityFlag, patternType, line1, line2, line3
    DrawValidityPanel wsDest, co, validityFlag, patternType, line1, line2, line3

    outEndRow = cohortRow
    Exit Sub

ErrHandler:
    ' Skip this subject block quietly.
End Sub

Private Sub DrawValidityPanel(ByVal ws As Worksheet, ByVal co As ChartObject, _
                              ByVal validityFlag As String, ByVal patternType As String, _
                              ByVal line1 As String, ByVal line2 As String, ByVal line3 As String)
    Dim panelLeft As Double, panelTop As Double
    Dim panelWidth As Double, panelHeight As Double
    Dim shp As Shape
    Dim fullText As String
    Dim fillColor As Long, borderColor As Long, fontColor As Long

    If co Is Nothing Then Exit Sub
    If co.Width <= 0 Or co.Height <= 0 Then Exit Sub

    panelHeight = co.Height
    panelWidth = co.Width * 1.65
    panelLeft = co.Left + co.Width + 10
    panelTop = co.Top

    fullText = "Flag: " & validityFlag & " | Pattern: " & patternType & vbCrLf & vbCrLf & _
               line1 & vbCrLf & vbCrLf & line2 & vbCrLf & vbCrLf & line3

    Select Case UCase$(Trim$(validityFlag))
        Case "LOW N"
            fillColor = RGB(255, 242, 204): borderColor = RGB(191, 144, 0): fontColor = RGB(120, 63, 4)
        Case "SKEWED"
            fillColor = RGB(252, 228, 214): borderColor = RGB(192, 80, 77): fontColor = RGB(148, 55, 49)
        Case "MIXED"
            fillColor = RGB(217, 225, 242): borderColor = RGB(79, 129, 189): fontColor = RGB(47, 84, 150)
        Case "VALID"
            fillColor = RGB(226, 240, 217): borderColor = RGB(118, 146, 60): fontColor = RGB(55, 86, 35)
        Case Else
            fillColor = RGB(242, 242, 242): borderColor = RGB(166, 166, 166): fontColor = RGB(89, 89, 89)
    End Select

    Set shp = ws.Shapes.AddShape(SHAPE_ROUNDED_RECTANGLE, panelLeft, panelTop, panelWidth, panelHeight)
    shp.Name = "FlagPanel_" & co.Name & "_" & ws.Index

    With shp
        .Fill.ForeColor.RGB = fillColor
        .line.ForeColor.RGB = borderColor
        .line.Weight = 1

        With .TextFrame2
            .TextRange.text = fullText
            .TextRange.Font.Size = 10
            .TextRange.Font.Name = "Calibri"
            .TextRange.Font.Fill.ForeColor.RGB = fontColor
            .MarginLeft = 8
            .MarginRight = 8
            .MarginTop = 6
            .MarginBottom = 6
            .WordWrap = True
            .AutoSize = msoFalse
            .TextRange.ParagraphFormat.Alignment = msoAlignLeft
        End With
    End With
End Sub

'---------------------------------------------------------
' ROW COLOURING
'---------------------------------------------------------
Private Sub ColourSubjectRow(ByVal ws As Worksheet, ByVal rowNum As Long, ByVal firstCol As Long, _
                             ByVal numBands As Long, ByVal topMaxIdx As Long, ByVal failMinIdx As Long, _
                             ByVal colPctPass As Long, ByVal colPctFail As Long, _
                             ByVal colPctTop As Long, ByVal colMean As Long)
    Dim v As Variant
    Dim j As Long
    Dim bandCol As Long

    For j = 1 To numBands
        bandCol = firstCol + j
        v = ws.Cells(rowNum, bandCol).value

        If IsNumeric(v) And v > 0 Then
            If j <= topMaxIdx Then
                ws.Cells(rowNum, bandCol).Font.Color = RGB(0, 128, 0)
            ElseIf j >= failMinIdx Then
                ws.Cells(rowNum, bandCol).Font.Color = RGB(192, 0, 0)
            Else
                ws.Cells(rowNum, bandCol).Font.Color = RGB(0, 0, 0)
            End If
        Else
            ws.Cells(rowNum, bandCol).Font.Color = RGB(0, 0, 0)
        End If
    Next j

    ws.Cells(rowNum, colPctPass).Font.Color = RGB(0, 0, 0)

    v = ws.Cells(rowNum, colPctFail).value
    If IsNumeric(v) And v > 0 Then
        ws.Cells(rowNum, colPctFail).Font.Color = RGB(192, 0, 0)
    Else
        ws.Cells(rowNum, colPctFail).Font.Color = RGB(0, 0, 0)
    End If

    v = ws.Cells(rowNum, colPctTop).value
    If IsNumeric(v) And v > 0 Then
        ws.Cells(rowNum, colPctTop).Font.Color = RGB(0, 128, 0)
    Else
        ws.Cells(rowNum, colPctTop).Font.Color = RGB(0, 0, 0)
    End If

    ws.Cells(rowNum, colMean).Font.Color = RGB(0, 0, 0)
End Sub

'---------------------------------------------------------
' GRADE SCHEME HELPERS
'---------------------------------------------------------
Private Function InitGradeScheme(ByVal schemeKey As String, _
                                 ByRef gradeLabels() As String, _
                                 ByRef passMaxIdx As Long, _
                                 ByRef failMinIdx As Long, _
                                 ByRef topMaxIdx As Long, _
                                 ByRef pctPassLabel As String, _
                                 ByRef pctFailLabel As String, _
                                 ByRef pctTopLabel As String, _
                                 ByRef meanLabel As String) As Boolean
    Select Case UCase$(Trim$(schemeKey))
        Case "G3"
            ReDim gradeLabels(1 To 9)
            gradeLabels(1) = "A1"
            gradeLabels(2) = "A2"
            gradeLabels(3) = "B3"
            gradeLabels(4) = "B4"
            gradeLabels(5) = "C5"
            gradeLabels(6) = "C6"
            gradeLabels(7) = "D7"
            gradeLabels(8) = "E8"
            gradeLabels(9) = "F9"

            passMaxIdx = 6
            failMinIdx = 7
            topMaxIdx = 2

            pctPassLabel = "%A1 - C6"
            pctFailLabel = "%D7 - F9"
            pctTopLabel = "%A1 - A2"
            meanLabel = "MSG"

        Case "G2"
            ReDim gradeLabels(1 To 6)
            gradeLabels(1) = "1"
            gradeLabels(2) = "2"
            gradeLabels(3) = "3"
            gradeLabels(4) = "4"
            gradeLabels(5) = "5"
            gradeLabels(6) = "6"

            passMaxIdx = 5
            failMinIdx = 6
            topMaxIdx = 2

            pctPassLabel = "%1 - 5"
            pctFailLabel = "%6"
            pctTopLabel = "%1 - 2"
            meanLabel = "Mean"

        Case "G1"
            ReDim gradeLabels(1 To 5)
            gradeLabels(1) = "A"
            gradeLabels(2) = "B"
            gradeLabels(3) = "C"
            gradeLabels(4) = "D"
            gradeLabels(5) = "E"

            passMaxIdx = 4
            failMinIdx = 5
            topMaxIdx = 2

            pctPassLabel = "%A - D"
            pctFailLabel = "%E"
            pctTopLabel = "%A - B"
            meanLabel = "Mean"

        Case Else
            InitGradeScheme = False
            Exit Function
    End Select

    InitGradeScheme = True
End Function

Private Sub InitPastelPalette(ByRef pastelColors() As Long)
    pastelColors(1) = RGB(0, 150, 136)
    pastelColors(2) = RGB(77, 182, 172)
    pastelColors(3) = RGB(129, 199, 132)
    pastelColors(4) = RGB(200, 230, 201)
    pastelColors(5) = RGB(255, 245, 157)
    pastelColors(6) = RGB(255, 224, 130)
    pastelColors(7) = RGB(255, 204, 128)
    pastelColors(8) = RGB(255, 171, 145)
    pastelColors(9) = RGB(239, 83, 80)
End Sub

Private Function NormalizeGradeForScheme(ByVal gradeRaw As String, ByVal schemeKey As String) As String
    Dim g As String
    g = UCase$(Trim$(gradeRaw))

    If g = "-" Or g = "AB" Or g = "VR" Then g = ""

    Select Case UCase$(schemeKey)
        Case "G3"
            If g = "9" Then g = "F9"
        Case Else
            ' No special mapping needed.
    End Select

    NormalizeGradeForScheme = g
End Function

Private Function GradeIndexByScheme(ByVal gradeStr As String, ByRef gradeLabels() As String) As Long
    Dim k As Long
    For k = LBound(gradeLabels) To UBound(gradeLabels)
        If gradeStr = gradeLabels(k) Then
            GradeIndexByScheme = k
            Exit Function
        End If
    Next k
    GradeIndexByScheme = 0
End Function

Private Function ComputeMeanBand(ByRef countsArr() As Long) As Double
    Dim i As Long
    Dim total As Long
    Dim weightedSum As Long

    For i = LBound(countsArr) To UBound(countsArr)
        weightedSum = weightedSum + countsArr(i) * i
        total = total + countsArr(i)
    Next i

    If total > 0 Then
        ComputeMeanBand = weightedSum / total
    Else
        ComputeMeanBand = 0
    End If
End Function

'---------------------------------------------------------
' DETECTION HELPERS
'---------------------------------------------------------
Private Function IsLikelySubjectGradeColumn(ByVal header As String) As Boolean
    Dim h As String
    h = UCase$(Trim$(header))

    ' Never treat score columns as grade columns.
    If InStr(1, h, "SCORE", vbTextCompare) > 0 Then
        IsLikelySubjectGradeColumn = False
        Exit Function
    End If

    If Right$(h, 7) = "(GRADE)" Then
        IsLikelySubjectGradeColumn = True
        Exit Function
    End If

    ' Backward compatibility for older staging sheets with no suffix.
    If InStr(1, h, " - G1", vbTextCompare) > 0 _
       Or InStr(1, h, " - G2", vbTextCompare) > 0 _
       Or InStr(1, h, " - G3", vbTextCompare) > 0 Then
        IsLikelySubjectGradeColumn = True
    End If
End Function

Private Function StripGradeHeaderSuffix(ByVal header As String) As String
    Dim h As String
    h = Trim$(header)

    If Len(h) >= 7 Then
        If UCase$(Right$(h, 7)) = "(GRADE)" Then
            StripGradeHeaderSuffix = Trim$(Left$(h, Len(h) - 7))
            Exit Function
        End If
    End If

    StripGradeHeaderSuffix = h
End Function

Private Function GetGradeSchemeKey(ByVal ws As Worksheet, ByVal gradeCol As Long, ByVal header As String) As String
    Dim keyFromHeader As String

    keyFromHeader = InferSchemeFromHeader(header)
    If keyFromHeader <> "" Then
        GetGradeSchemeKey = keyFromHeader
        Exit Function
    End If

    GetGradeSchemeKey = InferSchemeFromValues(ws, gradeCol)
End Function

Private Function InferSchemeFromHeader(ByVal header As String) As String
    Dim h As String
    h = UCase$(Trim$(header))

    If InStr(1, h, "- G1", vbTextCompare) > 0 Then
        InferSchemeFromHeader = "G1"
    ElseIf InStr(1, h, "- G2", vbTextCompare) > 0 Then
        InferSchemeFromHeader = "G2"
    ElseIf InStr(1, h, "- G3", vbTextCompare) > 0 Then
        InferSchemeFromHeader = "G3"
    End If
End Function

Private Function InferSchemeFromValues(ByVal ws As Worksheet, ByVal gradeCol As Long) As String
    Dim r As Long, lastRow As Long
    Dim v As String
    Dim g1Hits As Long, g2Hits As Long, g3Hits As Long
    Dim maxSamples As Long

    lastRow = ws.Cells(ws.Rows.count, gradeCol).End(xlUp).Row
    maxSamples = 200

    For r = 2 To lastRow
        v = UCase$(Trim$(CStr(ws.Cells(r, gradeCol).value)))

        If v <> "" And v <> "-" And v <> "AB" Then
            If IsGradeInScheme(v, "G1") Then g1Hits = g1Hits + 1
            If IsGradeInScheme(v, "G2") Then g2Hits = g2Hits + 1
            If IsGradeInScheme(v, "G3") Then g3Hits = g3Hits + 1
        End If

        If r - 1 >= maxSamples Then Exit For
    Next r

    If g3Hits >= g2Hits And g3Hits >= g1Hits And g3Hits >= 3 Then
        InferSchemeFromValues = "G3"
    ElseIf g2Hits >= g1Hits And g2Hits >= 3 Then
        InferSchemeFromValues = "G2"
    ElseIf g1Hits >= 3 Then
        InferSchemeFromValues = "G1"
    Else
        InferSchemeFromValues = ""
    End If
End Function

Private Function IsGradeInScheme(ByVal v As String, ByVal schemeKey As String) As Boolean
    Dim g As String
    g = UCase$(Trim$(v))

    Select Case UCase$(schemeKey)
        Case "G1"
            IsGradeInScheme = (g = "A" Or g = "B" Or g = "C" Or g = "D" Or g = "E")
        Case "G2"
            IsGradeInScheme = (g = "1" Or g = "2" Or g = "3" Or g = "4" Or g = "5" Or g = "6")
        Case "G3"
            IsGradeInScheme = (g = "A1" Or g = "A2" Or g = "B3" Or g = "B4" Or g = "C5" Or _
                               g = "C6" Or g = "D7" Or g = "E8" Or g = "F9" Or g = "9")
    End Select
End Function

'---------------------------------------------------------
' SHEET NAME HELPERS
'---------------------------------------------------------
Private Function CleanSheetNameFragment(ByVal txt As String) As String
    Dim s As String
    s = txt
    s = Replace(s, ":", "")
    s = Replace(s, "\", "")
    s = Replace(s, "/", "")
    s = Replace(s, "?", "")
    s = Replace(s, "*", "")
    s = Replace(s, "[", "")
    s = Replace(s, "]", "")
    CleanSheetNameFragment = s
End Function

Private Function BuildSecDestSheetName(ByVal levelCode As String, ByVal examLabel As String) As String
    Dim prefix As String
    Dim yearPart As String
    Dim baseLabel As String
    Dim safeBase As String
    Dim maxShort As Long
    Dim safeName As String
    Dim yearCandidate As String

    prefix = levelCode & "_Subj Analysis_"
    yearPart = ""

    If Len(examLabel) >= 4 Then
        yearCandidate = Right$(examLabel, 4)
        If IsNumeric(yearCandidate) Then yearPart = yearCandidate
    End If

    If yearPart <> "" Then
        baseLabel = Left$(examLabel, Len(examLabel) - 4)

        Do While Len(baseLabel) > 0 And _
              (Right$(baseLabel, 1) = "_" Or Right$(baseLabel, 1) = " " Or Right$(baseLabel, 1) = "-")
            baseLabel = Left$(baseLabel, Len(baseLabel) - 1)
        Loop

        safeBase = CleanSheetNameFragment(baseLabel)
        If safeBase = "" Then safeBase = "Exam"

        maxShort = 31 - Len(prefix) - 1 - Len(yearPart)
        If maxShort < 1 Then maxShort = 1
        If Len(safeBase) > maxShort Then safeBase = Left$(safeBase, maxShort)

        safeName = prefix & safeBase & "_" & yearPart
        If Len(safeName) > 31 Then safeName = Left$(safeName, 31)
    Else
        safeName = prefix & examLabel
        safeName = CleanSheetNameFragment(safeName)
        If Len(safeName) > 31 Then safeName = Left$(safeName, 31)
    End If

    BuildSecDestSheetName = safeName
End Function

Private Function GetMinSubjectN() As Long
    Dim ws As Worksheet
    Dim v As Variant

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("Settings")
    On Error GoTo 0

    If ws Is Nothing Then
        GetMinSubjectN = DEFAULT_MIN_SUBJECT_N
        Exit Function
    End If

    ' Optional override: Settings!L6
    v = ws.Range("L6").value
    If IsNumeric(v) Then
        GetMinSubjectN = CLng(v)
        If GetMinSubjectN < 1 Then GetMinSubjectN = DEFAULT_MIN_SUBJECT_N
    Else
        GetMinSubjectN = DEFAULT_MIN_SUBJECT_N
    End If
End Function

Private Function InferLevelCodeFromClass(ByVal className As String) As String
    Dim s As String
    Dim i As Long
    Dim ch As String

    s = UCase$(Trim$(className))
    If s = "" Then Exit Function

    ' Preferred match for class names like S1-..., S2 ..., etc.
    For i = 1 To Len(s) - 1
        If Mid$(s, i, 1) = "S" Then
            ch = Mid$(s, i + 1, 1)
            If ch >= "1" And ch <= "5" Then
                InferLevelCodeFromClass = "S" & ch
                Exit Function
            End If
        End If
    Next i

    ' Fallback: first standalone digit 1..5 in the class string.
    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        If ch >= "1" And ch <= "5" Then
            InferLevelCodeFromClass = "S" & ch
            Exit Function
        End If
    Next i
End Function

Private Function SubjectAlreadyAdded(ByRef subjectNames() As String, _
                                     ByVal subjCount As Long, _
                                     ByVal subjectName As String) As Boolean
    Dim i As Long
    For i = 1 To subjCount
        If StrComp(Trim$(subjectNames(i)), Trim$(subjectName), vbTextCompare) = 0 Then
            SubjectAlreadyAdded = True
            Exit Function
        End If
    Next i
End Function

Private Function FindScoreColumnForSubject(ByVal ws As Worksheet, _
                                           ByVal headerRow As Long, _
                                           ByVal subjectName As String) As Long
    Dim lastCol As Long, c As Long
    Dim h As String
    Dim baseName As String

    lastCol = ws.Cells(headerRow, ws.Columns.count).End(xlToLeft).Column
    For c = 1 To lastCol
        h = Trim$(CStr(ws.Cells(headerRow, c).value))
        If InStr(1, UCase$(h), "SCORE", vbTextCompare) > 0 Then
            baseName = NormalizeSubjectHeaderBase(h)
            If StrComp(baseName, NormalizeSubjectHeaderBase(subjectName), vbTextCompare) = 0 Then
                FindScoreColumnForSubject = c
                Exit Function
            End If
        End If
    Next c
End Function

Private Function NormalizeSubjectHeaderBase(ByVal headerText As String) As String
    Dim h As String
    h = Trim$(headerText)

    If UCase$(Right$(h, 7)) = "(GRADE)" Then
        h = Trim$(Left$(h, Len(h) - 7))
    End If

    If UCase$(Right$(h, 7)) = "(SCORE)" Then
        h = Trim$(Left$(h, Len(h) - 7))
    End If

    NormalizeSubjectHeaderBase = h
End Function

Private Function FindClassIndex(ByRef classList() As String, ByVal classCount As Long, ByVal className As String) As Long
    Dim i As Long
    For i = 1 To classCount
        If StrComp(classList(i), className, vbTextCompare) = 0 Then
            FindClassIndex = i
            Exit Function
        End If
    Next i
End Function

'---------------------------------------------------------
' GENERIC HELPERS
'---------------------------------------------------------
Private Function IsFailGradeByScheme(ByVal gradeStr As String, ByVal schemeKey As String) As Boolean
    Dim g As String
    g = UCase$(Trim$(gradeStr))

    Select Case UCase$(Trim$(schemeKey))
        Case "G3"
            IsFailGradeByScheme = (g = "D7" Or g = "E8" Or g = "F9")
        Case "G2"
            IsFailGradeByScheme = (g = "6")
        Case "G1"
            IsFailGradeByScheme = (g = "E")
        Case Else
            IsFailGradeByScheme = False
    End Select
End Function

Private Function IsTopGradeByScheme(ByVal gradeStr As String, ByVal schemeKey As String) As Boolean
    Dim g As String
    g = UCase$(Trim$(gradeStr))

    Select Case UCase$(Trim$(schemeKey))
        Case "G3"
            IsTopGradeByScheme = (g = "A1" Or g = "A2")
        Case "G2"
            IsTopGradeByScheme = (g = "1" Or g = "2")
        Case "G1"
            IsTopGradeByScheme = (g = "A" Or g = "B")
        Case Else
            IsTopGradeByScheme = False
    End Select
End Function

Private Function RiskBandRank(ByVal riskBand As String) As Long
    Select Case UCase$(Trim$(riskBand))
        Case "AT RISK"
            RiskBandRank = 1
        Case "MONITOR"
            RiskBandRank = 2
        Case Else
            RiskBandRank = 3
    End Select
End Function

Private Function GetAtRiskFailThreshold() As Long
    Dim ws As Worksheet
    Dim v As Variant

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("Settings")
    On Error GoTo 0

    If ws Is Nothing Then
        GetAtRiskFailThreshold = DEFAULT_AT_RISK_FAIL_THRESHOLD
        Exit Function
    End If

    ' Optional override: Settings!L7
    v = ws.Range("L7").value
    If IsNumeric(v) Then
        GetAtRiskFailThreshold = CLng(v)
        If GetAtRiskFailThreshold < 1 Then GetAtRiskFailThreshold = DEFAULT_AT_RISK_FAIL_THRESHOLD
    Else
        GetAtRiskFailThreshold = DEFAULT_AT_RISK_FAIL_THRESHOLD
    End If
End Function

Private Function GetGroupThresholdPercent() As Double
    Dim ws As Worksheet
    Dim v As Variant
    Dim p As Double

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("Settings")
    On Error GoTo 0

    If ws Is Nothing Then
        GetGroupThresholdPercent = 70#
        Exit Function
    End If

    ' Optional override: Settings!L8
    v = ws.Range("L8").value
    If IsNumeric(v) Then
        p = CDbl(v)
        If p <= 1# Then p = p * 100#
        If p < 1# Or p > 100# Then p = 70#
        GetGroupThresholdPercent = p
    Else
        GetGroupThresholdPercent = 70#
    End If
End Function

Private Function ResolveFsbbGroup(ByVal g1Taken As Long, _
                                  ByVal g2Taken As Long, _
                                  ByVal g3Taken As Long, _
                                  ByVal attemptedCount As Long, _
                                  ByVal thresholdPct As Double) As String
    Dim p1 As Double, p2 As Double, p3 As Double

    If attemptedCount <= 0 Then
        ResolveFsbbGroup = ""
        Exit Function
    End If

    p1 = (CDbl(g1Taken) / CDbl(attemptedCount)) * 100#
    p2 = (CDbl(g2Taken) / CDbl(attemptedCount)) * 100#
    p3 = (CDbl(g3Taken) / CDbl(attemptedCount)) * 100#

    If p3 >= thresholdPct Then
        ResolveFsbbGroup = "G3"
    ElseIf p2 >= thresholdPct Then
        ResolveFsbbGroup = "G2"
    ElseIf p1 >= thresholdPct Then
        ResolveFsbbGroup = "G1"
    Else
        ResolveFsbbGroup = "MIXED"
    End If
End Function

Private Function FindFirstHeaderColumn(ByVal ws As Worksheet, _
                                       ByVal headerRow As Long, _
                                       ByVal headerCandidates As Variant) As Long
    Dim i As Long
    Dim col As Long

    For i = LBound(headerCandidates) To UBound(headerCandidates)
        col = FindHeaderColumn(ws, headerRow, CStr(headerCandidates(i)))
        If col > 0 Then
            FindFirstHeaderColumn = col
            Exit Function
        End If
    Next i
End Function

Private Function FindHeaderColumn(ByVal ws As Worksheet, ByVal headerRow As Long, ByVal headerName As String) As Long
    Dim lastCol As Long, c As Long
    Dim h As String

    lastCol = ws.Cells(headerRow, ws.Columns.count).End(xlToLeft).Column
    For c = 1 To lastCol
        h = Trim$(CStr(ws.Cells(headerRow, c).value))
        If StrComp(h, headerName, vbTextCompare) = 0 Then
            FindHeaderColumn = c
            Exit Function
        End If
    Next c

    FindHeaderColumn = 0
End Function

Private Sub SortStringArray(ByRef arr() As String)
    Dim i As Long, j As Long
    Dim temp As String

    For i = LBound(arr) To UBound(arr) - 1
        For j = i + 1 To UBound(arr)
            If arr(j) < arr(i) Then
                temp = arr(i)
                arr(i) = arr(j)
                arr(j) = temp
            End If
        Next j
    Next i
End Sub

