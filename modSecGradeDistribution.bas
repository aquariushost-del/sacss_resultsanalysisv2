Attribute VB_Name = "modSecGradeDistribution"
Option Explicit

Private Const DEFAULT_MIN_SUBJECT_N As Long = 10
Private Const DEFAULT_AT_RISK_FAIL_THRESHOLD As Long = 3
Private Const SHAPE_ROUNDED_RECTANGLE As Long = 5
Private Const ATRISK_NAV_SHEET_NAME As String = "Dashboard"
Private Const ATRISK_NAV_START_CELL As String = "M3"
Private Const ATRISK_NAV_BTN_PREFIX As String = "Nav_AtRisk_"
Private Const TOP_NAV_START_CELL As String = "P3"
Private Const TOP_NAV_BTN_PREFIX As String = "Nav_TopQual_"
Private Const NAV_BTN_WIDTH_FACTOR As Double = 0.5
Private Const LEVEL_MODE_AUTO_FSBB As String = "AUTO_FSBB"
Private Const LEVEL_MODE_LEGACY_NO_DOWNWARD As String = "LEGACY_NO_DOWNWARD"

Private gSecBatchMode As Boolean

Private Type TopStudentRec
    LevelCode As String
    ClassName As String
    RegNo As String
    StudentName As String
    GroupCode As String
    TopCount As Long
    TopPrimaryCount As Long
    TopSecondaryCount As Long
    DownwardRemarks As String
    RawTopGrades As String
    SubjectMix As String
End Type

Private Type SubjectTopRec
    SubjectName As String
    SchemeKey As String
    ClassName As String
    RegNo As String
    StudentName As String
    ScorePct As Double
    GradeText As String
End Type

Private Type SecReportGroup
    LevelCode As String
    AssessmentKey As String
    AssessmentLabel As String
    YearText As String
    SheetName As String
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
Public Sub BuildAllSecReportsAndMenus()
    Dim previousScreenUpdating As Boolean
    Dim errorText As String

    On Error GoTo ErrHandler

    previousScreenUpdating = Application.ScreenUpdating
    Application.ScreenUpdating = False
    gSecBatchMode = True

    Application.StatusBar = "1 of 5: Analysing SEC subjects..."
    BuildAllSec_SubjectAnalysis

    Application.StatusBar = "2 of 5: Building SEC AtRisk summaries..."
    BuildSec_AtRiskSummary

    Application.StatusBar = "3 of 5: Building SEC Top Students summaries..."
    BuildSec_TopQualityByLevel

    Application.StatusBar = "4 of 5: Building class and Subject x Class reports..."
    BuildSecClassAndSubjectClassAnalysis True

    Application.StatusBar = "5 of 5: Refreshing the SEC menu..."
    BuildSubjectAnalysisNavigation

    gSecBatchMode = False
    Application.StatusBar = False
    Application.ScreenUpdating = previousScreenUpdating
    MsgBox "SEC subject analysis, AtRisk summaries, Top Students, Class Analysis, Subject x Class reports and menus have been refreshed.", _
           vbInformation, "SEC Reports Complete"
    Exit Sub

ErrHandler:
    errorText = Err.Description
    gSecBatchMode = False
    Application.StatusBar = False
    Application.ScreenUpdating = previousScreenUpdating
    MsgBox "The SEC report run stopped: " & errorText, vbCritical, "SEC Reports"
End Sub

Public Sub BuildAllSec_SubjectAnalysis()
    Dim wb As Workbook
    Dim ws As Worksheet

    On Error GoTo ErrHandler

    Set wb = ThisWorkbook

    For Each ws In wb.Worksheets
        ProcessSecSourceSheet ws
    Next ws

    If Not gSecBatchMode Then MsgBox "Subject Analysis generated for all eligible sheets.", vbInformation
    Exit Sub

ErrHandler:
    If gSecBatchMode Then
        Err.Raise Err.Number, "BuildAllSec_SubjectAnalysis", Err.Description
    Else
        MsgBox "Error in BuildAllSec_SubjectAnalysis: " & Err.Description, vbCritical
    End If
End Sub

Private Sub CollectSecReportGroups(ByRef groups() As SecReportGroup, _
                                   ByRef groupCount As Long, _
                                   ByVal reportPrefix As String)
    Dim ws As Worksheet
    Dim levelCode As String, assessmentKey As String
    Dim assessmentLabel As String, yearText As String
    Dim i As Long, foundIndex As Long

    For Each ws In ThisWorkbook.Worksheets
        levelCode = "": assessmentKey = "": assessmentLabel = "": yearText = ""
        If GetSecSourceReportLabels(ws, levelCode, assessmentKey, assessmentLabel, yearText) Then
            foundIndex = 0
            For i = 1 To groupCount
                If groups(i).LevelCode = levelCode And _
                   groups(i).AssessmentKey = assessmentKey And _
                   groups(i).YearText = yearText Then
                    foundIndex = i
                    Exit For
                End If
            Next i
            If foundIndex = 0 Then
                groupCount = groupCount + 1
                ReDim Preserve groups(1 To groupCount)
                With groups(groupCount)
                    .LevelCode = levelCode
                    .AssessmentKey = assessmentKey
                    .AssessmentLabel = assessmentLabel
                    .YearText = yearText
                    .SheetName = BuildSecReportSheetName(reportPrefix, levelCode, assessmentKey, yearText)
                End With
            End If
        End If
    Next ws

    SortSecReportGroups groups, groupCount
End Sub

Private Function GetSecSourceReportLabels(ByVal ws As Worksheet, _
                                          ByRef levelCode As String, _
                                          ByRef assessmentKey As String, _
                                          ByRef assessmentLabel As String, _
                                          ByRef yearText As String) As Boolean
    Dim classCol As Long, assessmentCol As Long, yearCol As Long
    Dim lastRow As Long, lastCol As Long, r As Long, c As Long
    Dim firstClass As String, headerText As String
    Dim hasGradeColumn As Boolean

    If LCase$(ws.Name) Like "*settings*" _
       Or LCase$(ws.Name) Like "*config*" _
       Or LCase$(ws.Name) Like "*menu*" _
       Or LCase$(ws.Name) Like "*lookup*" _
       Or LCase$(ws.Name) Like "*summary*" _
       Or LCase$(ws.Name) Like "*template*" _
       Or InStr(1, LCase$(ws.Name), "_subj analysis_") > 0 _
       Or InStr(1, LCase$(ws.Name), "dashboard") > 0 _
       Or InStr(1, LCase$(ws.Name), "atrisk_") > 0 _
       Or InStr(1, LCase$(ws.Name), "topqual_") > 0 _
       Or InStr(1, LCase$(ws.Name), "classan_") > 0 _
       Or InStr(1, LCase$(ws.Name), "subjclass_") > 0 Then Exit Function

    classCol = FindHeaderColumn(ws, 1, "Class")
    assessmentCol = FindHeaderColumn(ws, 1, "Assessment")
    yearCol = FindHeaderColumn(ws, 1, "Year")
    If classCol = 0 Or assessmentCol = 0 Or yearCol = 0 Then Exit Function

    lastRow = ws.Cells(ws.Rows.count, classCol).End(xlUp).Row
    lastCol = ws.Cells(1, ws.Columns.count).End(xlToLeft).Column
    For c = 1 To lastCol
        headerText = Trim$(CStr(ws.Cells(1, c).value))
        If headerText <> "" And IsLikelySubjectGradeColumn(headerText) Then
            hasGradeColumn = True
            Exit For
        End If
    Next c
    If Not hasGradeColumn Then Exit Function

    For r = 2 To lastRow
        If firstClass = "" Then firstClass = Trim$(CStr(ws.Cells(r, classCol).value))
        If assessmentLabel = "" Then assessmentLabel = Trim$(CStr(ws.Cells(r, assessmentCol).value))
        If yearText = "" Then yearText = NormalizeSecReportYear(CStr(ws.Cells(r, yearCol).value))
        If firstClass <> "" And assessmentLabel <> "" And yearText <> "" Then Exit For
    Next r

    If firstClass = "" Or UCase$(Left$(firstClass, 1)) = "Y" Then Exit Function
    levelCode = InferLevelCodeFromClass(firstClass)
    assessmentKey = CanonicalSecReportAssessmentKey(assessmentLabel)
    If levelCode = "" Or assessmentKey = "" Or yearText = "" Then Exit Function
    GetSecSourceReportLabels = True
End Function

Private Function SecSourceMatchesGroup(ByVal ws As Worksheet, _
                                       ByRef targetGroup As SecReportGroup) As Boolean
    Dim levelCode As String, assessmentKey As String
    Dim assessmentLabel As String, yearText As String
    If GetSecSourceReportLabels(ws, levelCode, assessmentKey, assessmentLabel, yearText) Then
        SecSourceMatchesGroup = (levelCode = targetGroup.LevelCode And _
                                 assessmentKey = targetGroup.AssessmentKey And _
                                 yearText = targetGroup.YearText)
    End If
End Function

Private Function CanonicalSecReportAssessmentKey(ByVal assessmentLabel As String) As String
    Dim compact As String, i As Long, ch As String
    For i = 1 To Len(UCase$(assessmentLabel))
        ch = Mid$(UCase$(assessmentLabel), i, 1)
        If ch Like "[A-Z0-9]" Then compact = compact & ch
    Next i

    Select Case True
        Case InStr(compact, "WA1") > 0 Or InStr(compact, "TERM1WA") > 0 Or InStr(compact, "TERM1NWA") > 0
            CanonicalSecReportAssessmentKey = "WA1"
        Case InStr(compact, "WA2") > 0 Or InStr(compact, "TERM2WA") > 0 Or InStr(compact, "TERM2NWA") > 0
            CanonicalSecReportAssessmentKey = "WA2"
        Case InStr(compact, "FIRSTCOMBINED") > 0 Or InStr(compact, "1STCOMBINED") > 0 Or _
             InStr(compact, "COMBINED1") > 0 Or InStr(compact, "SEMESTER1") > 0 Or _
             InStr(compact, "TERM2COMBINED") > 0
            CanonicalSecReportAssessmentKey = "1COMB"
        Case InStr(compact, "WA3") > 0 Or InStr(compact, "TERM3WA") > 0 Or InStr(compact, "TERM3NWA") > 0
            CanonicalSecReportAssessmentKey = "WA3"
        Case InStr(compact, "PRELIM") > 0
            CanonicalSecReportAssessmentKey = "PRELIM"
        Case InStr(compact, "SECONDCOMBINED") > 0 Or InStr(compact, "2NDCOMBINED") > 0 Or _
             InStr(compact, "COMBINED2") > 0 Or InStr(compact, "SEMESTER2") > 0 Or _
             InStr(compact, "TERM3COMBINED") > 0 Or InStr(compact, "TERM4COMBINED") > 0
            CanonicalSecReportAssessmentKey = "2COMB"
        Case InStr(compact, "EYE") > 0 Or InStr(compact, "ENDOFYEAR") > 0
            CanonicalSecReportAssessmentKey = "EYE"
        Case InStr(compact, "MIDYEAR") > 0 Or compact = "MYE"
            CanonicalSecReportAssessmentKey = "MYE"
        Case Else
            CanonicalSecReportAssessmentKey = compact
    End Select
End Function

Private Function NormalizeSecReportYear(ByVal valueText As String) As String
    Dim i As Long, token As String
    For i = 1 To Len(valueText) - 3
        token = Mid$(valueText, i, 4)
        If IsNumeric(token) Then
            If CLng(token) >= 2000 And CLng(token) <= 2099 Then
                NormalizeSecReportYear = token
                Exit Function
            End If
        End If
    Next i
End Function

Private Function BuildSecReportSheetName(ByVal reportPrefix As String, _
                                         ByVal levelCode As String, _
                                         ByVal assessmentKey As String, _
                                         ByVal yearText As String) As String
    Dim maxAssessmentLength As Long, safeAssessment As String
    maxAssessmentLength = 31 - Len(reportPrefix) - Len(levelCode) - Len(yearText) - 2
    If maxAssessmentLength < 1 Then maxAssessmentLength = 1
    safeAssessment = Left$(assessmentKey, maxAssessmentLength)
    BuildSecReportSheetName = reportPrefix & levelCode & "_" & safeAssessment & "_" & yearText
End Function

Private Sub SortSecReportGroups(ByRef groups() As SecReportGroup, ByVal groupCount As Long)
    Dim i As Long, j As Long
    For i = 1 To groupCount - 1
        For j = i + 1 To groupCount
            If SecReportGroupBefore(groups(j), groups(i)) Then SwapSecReportGroups groups(i), groups(j)
        Next j
    Next i
End Sub

Private Function SecReportGroupBefore(ByRef a As SecReportGroup, ByRef b As SecReportGroup) As Boolean
    Dim ao As Long, bo As Long
    If a.YearText <> b.YearText Then SecReportGroupBefore = (a.YearText > b.YearText): Exit Function
    ao = SecReportAssessmentOrder(a.AssessmentKey): bo = SecReportAssessmentOrder(b.AssessmentKey)
    If ao <> bo Then SecReportGroupBefore = (ao < bo): Exit Function
    If a.LevelCode <> b.LevelCode Then SecReportGroupBefore = (a.LevelCode < b.LevelCode): Exit Function
    SecReportGroupBefore = (a.AssessmentKey < b.AssessmentKey)
End Function

Private Function SecReportAssessmentOrder(ByVal assessmentKey As String) As Long
    Select Case assessmentKey
        Case "WA1": SecReportAssessmentOrder = 1
        Case "WA2": SecReportAssessmentOrder = 2
        Case "1COMB": SecReportAssessmentOrder = 3
        Case "WA3": SecReportAssessmentOrder = 4
        Case "PRELIM": SecReportAssessmentOrder = 5
        Case "2COMB": SecReportAssessmentOrder = 6
        Case "EYE": SecReportAssessmentOrder = 7
        Case Else: SecReportAssessmentOrder = 100
    End Select
End Function

Private Sub SwapSecReportGroups(ByRef a As SecReportGroup, ByRef b As SecReportGroup)
    Dim textValue As String
    textValue = a.LevelCode: a.LevelCode = b.LevelCode: b.LevelCode = textValue
    textValue = a.AssessmentKey: a.AssessmentKey = b.AssessmentKey: b.AssessmentKey = textValue
    textValue = a.AssessmentLabel: a.AssessmentLabel = b.AssessmentLabel: b.AssessmentLabel = textValue
    textValue = a.YearText: a.YearText = b.YearText: b.YearText = textValue
    textValue = a.SheetName: a.SheetName = b.SheetName: b.SheetName = textValue
End Sub

Public Sub BuildSec_TopQualityByLevel()
    Dim wb As Workbook
    Dim ws As Worksheet, wsOut As Worksheet
    Dim groups() As SecReportGroup
    Dim groupCount As Long, groupIndex As Long
    Dim recs() As TopStudentRec
    Dim recCount As Long
    Dim subjectTopRecs() As SubjectTopRec
    Dim subjectTopCount As Long
    Dim outRow As Long
    Dim groupThresholdPct As Double
    Dim levelMode As String

    On Error GoTo ErrHandler

    Set wb = ThisWorkbook
    groupThresholdPct = GetGroupThresholdPercent()

    CollectSecReportGroups groups, groupCount, "TopQual_"
    For groupIndex = 1 To groupCount
        Set wsOut = GetOrCreateWorksheet(groups(groupIndex).SheetName)
        wsOut.Cells.UnMerge
        wsOut.Cells.Clear
        PrepareTopQualitySheet wsOut, groups(groupIndex).LevelCode, _
                               groups(groupIndex).AssessmentLabel, groups(groupIndex).YearText
        levelMode = GetLevelMode(groups(groupIndex).LevelCode)

        recCount = 0
        subjectTopCount = 0
        For Each ws In wb.Worksheets
            If SecSourceMatchesGroup(ws, groups(groupIndex)) Then
                AppendTopQualityFromSourceSheet ws, groups(groupIndex).LevelCode, recs, recCount, _
                                                subjectTopRecs, subjectTopCount, groupThresholdPct
            End If
        Next ws

        outRow = 5
        outRow = WriteSubjectTopPerformersSection(wsOut, outRow, subjectTopRecs, subjectTopCount)
        outRow = WriteTopGroupSection(wsOut, outRow, groups(groupIndex).LevelCode, "G3", 5, recs, recCount, levelMode)
        outRow = WriteTopGroupSection(wsOut, outRow, groups(groupIndex).LevelCode, "G2", 5, recs, recCount, levelMode)
        outRow = WriteTopGroupSection(wsOut, outRow, groups(groupIndex).LevelCode, "G1", 5, recs, recCount, levelMode)
        If UCase$(levelMode) <> LEVEL_MODE_LEGACY_NO_DOWNWARD Then
            outRow = WriteTopGroupSection(wsOut, outRow, groups(groupIndex).LevelCode, "MIXED", 5, recs, recCount, levelMode)
        End If

        FormatTopQualitySheet wsOut, outRow - 1
        AddTopQualityHomeButton wsOut
        FreezeReportHeaderRows wsOut, 2
    Next groupIndex

    BuildTopQualityNavigation

    If Not gSecBatchMode Then _
        MsgBox groupCount & " assessment-specific top-quality sheet(s) built.", vbInformation
    Exit Sub

ErrHandler:
    If gSecBatchMode Then
        Err.Raise Err.Number, "BuildSec_TopQualityByLevel", Err.Description
    Else
        MsgBox "Error in BuildSec_TopQualityByLevel: " & Err.Description, vbCritical
    End If
End Sub

Public Sub BuildTopQualityNavigation()
    Dim wsNav As Worksheet
    Dim startCell As Range
    Dim startRow As Long, startCol As Long
    Dim rowPtr As Long
    Dim reportSheets() As String
    Dim reportCount As Long, i As Long
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

    ' Rows 15 onward are reserved for the Subject x Class menu below.
    wsNav.Range(wsNav.Cells(startRow, startCol), wsNav.Cells(startRow + 11, startCol + 3)).Clear
    For k = wsNav.Shapes.count To 1 Step -1
        Set shp = wsNav.Shapes(k)
        If Left$(shp.Name, Len(TOP_NAV_BTN_PREFIX)) = TOP_NAV_BTN_PREFIX Then shp.Delete
    Next k

    wsNav.Cells(startRow, startCol).value = "Top Students Menu"
    wsNav.Cells(startRow, startCol).Font.Bold = True
    wsNav.Cells(startRow, startCol).Font.Size = 12
    wsNav.Cells(startRow, startCol).Font.Color = RGB(31, 73, 125)
    rowPtr = startRow + 1

    CollectVersionedReportSheetNames "TopQual_", reportSheets, reportCount
    If reportCount = 0 Then
        wsNav.Cells(rowPtr, startCol).value = "No assessment-specific top-student reports built."
        wsNav.Cells(rowPtr, startCol).Font.Italic = True
    Else
        For i = 1 To reportCount
            CreateTopQualityNavButton wsNav, reportSheets(i), ReportNavigationLabel(reportSheets(i), "TopQual_"), rowPtr, startCol
            rowPtr = rowPtr + 2
        Next i
    End If
    Exit Sub

ErrHandler:
    MsgBox "Error in BuildTopQualityNavigation: " & Err.Description, vbExclamation
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
    btnWidth = wsNav.Columns(firstCol).Resize(, 5).Width * NAV_BTN_WIDTH_FACTOR
    btnHeight = wsNav.Rows(rowNum).Height * 1.3

    Set shp = wsNav.Shapes.AddShape( _
        Type:=SHAPE_ROUNDED_RECTANGLE, _
        Left:=leftPos, _
        Top:=topPos, _
        Width:=btnWidth, _
        Height:=btnHeight)

    With shp
        .Name = TOP_NAV_BTN_PREFIX & targetSheetName
        .Fill.ForeColor.RGB = RGB(217, 234, 211)
        .Fill.Transparency = 0#
        .line.ForeColor.RGB = RGB(106, 168, 79)
        .line.Weight = 1.5
        With .TextFrame2
            .TextRange.text = displayText
            .TextRange.Font.Name = "Calibri"
            .TextRange.Font.Size = 10.5
            .TextRange.Font.Fill.ForeColor.RGB = RGB(39, 78, 19)
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

Private Sub AddTopQualityHomeButton(ByVal ws As Worksheet)
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
        .Fill.ForeColor.RGB = RGB(217, 234, 211)
        .line.ForeColor.RGB = RGB(106, 168, 79)
        .line.Weight = 1.5
        With .TextFrame2
            .TextRange.text = "Home"
            .TextRange.Font.Name = "Calibri"
            .TextRange.Font.Size = 11
            .TextRange.Font.Fill.ForeColor.RGB = RGB(39, 78, 19)
            .VerticalAnchor = msoAnchorMiddle
            .TextRange.ParagraphFormat.Alignment = msoAlignCenter
            .MarginLeft = 4
            .MarginRight = 4
        End With
    End With

    ws.Hyperlinks.Add Anchor:=shp, Address:="", SubAddress:="'Dashboard'!A1"
End Sub

'---------------------------------------------------------
' ENTRY POINT - BUILD STUDENTS AT RISK SUMMARY (SEC)
'---------------------------------------------------------
Public Sub BuildSec_AtRiskSummary()
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim wsOut As Worksheet
    Dim groups() As SecReportGroup
    Dim groupCount As Long, groupIndex As Long
    Dim outRow As Long
    Dim addedRows As Long
    Dim threshold As Long

    On Error GoTo ErrHandler

    Set wb = ThisWorkbook
    threshold = GetAtRiskFailThreshold()

    CollectSecReportGroups groups, groupCount, "AtRisk_"
    For groupIndex = 1 To groupCount
        Set wsOut = GetOrCreateWorksheet(groups(groupIndex).SheetName)
        PrepareAtRiskSheet wsOut, groups(groupIndex).LevelCode, threshold, _
                           groups(groupIndex).AssessmentLabel, groups(groupIndex).YearText

        outRow = 5
        For Each ws In wb.Worksheets
            If SecSourceMatchesGroup(ws, groups(groupIndex)) Then
                addedRows = AppendSecAtRiskFromSourceSheet(ws, wsOut, outRow, threshold, groups(groupIndex).LevelCode)
                If addedRows > 0 Then outRow = outRow + addedRows
            End If
        Next ws

        If outRow = 5 Then
            wsOut.Cells(outRow, 1).value = "No eligible SEC result rows found for " & groups(groupIndex).LevelCode & "."
            outRow = outRow + 1
        End If

        FinalizeAtRiskSheet wsOut, outRow - 1
        FreezeReportHeaderRows wsOut, 4
    Next groupIndex

    BuildSec_AtRiskNavigation
    BuildAllAtRiskHomeButtons

    If groupCount > 0 Then
        wb.Worksheets(groups(1).SheetName).Activate
        wb.Worksheets(groups(1).SheetName).Range("A1").Select
    End If

    If Not gSecBatchMode Then _
        MsgBox groupCount & " assessment-specific at-risk sheet(s) built.", vbInformation
    Exit Sub

ErrHandler:
    If gSecBatchMode Then
        Err.Raise Err.Number, "BuildSec_AtRiskSummary", Err.Description
    Else
        MsgBox "Error in BuildSec_AtRiskSummary: " & Err.Description, vbCritical
    End If
End Sub

Private Sub FreezeReportHeaderRows(ByVal ws As Worksheet, ByVal headerRowCount As Long)
    If headerRowCount < 1 Then Exit Sub

    On Error GoTo CleanExit
    ws.Activate
    With ActiveWindow
        .FreezePanes = False
        .SplitColumn = 0
        .SplitRow = 0
    End With
    ws.Cells(headerRowCount + 1, 1).Select
    ActiveWindow.FreezePanes = True

CleanExit:
    On Error GoTo 0
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
                subjectName = StripGradeHeaderSuffix(header)
                If UCase$(GetLevelMode(levelCode)) = LEVEL_MODE_LEGACY_NO_DOWNWARD Then
                    schemeKey = GetLegacySchemeFromHeader(header)
                    If schemeKey = "" Then schemeKey = GetGradeSchemeKey(wsSrc, c, header)
                Else
                    schemeKey = GetGradeSchemeKey(wsSrc, c, header)
                End If

                If schemeKey <> "" And Not IsExcludedSecSubject(subjectName) _
                   And Not SubjectAlreadyAdded(subjectNames, subjCount, subjectName) Then
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

    ' Use an explicit row height for every generated layout row. This avoids
    ' the workbook/platform mismatch between the stored default row height and
    ' the coordinate grid Excel uses for floating drawing objects.
    If destRowHeader > 3 Then
        wsDest.Rows("2:" & CStr(destRowHeader + 5)).RowHeight = 15#
    End If

    ' Column AutoFit calls made while later subjects are built can move or
    ' resize earlier floating objects in Excel for Mac. Snap every chart and
    ' panel back to its encoded worksheet anchors after the sheet is complete.
    wsDest.Activate
    DoEvents
    RealignSecAnalysisObjects wsDest
    DoEvents
    RealignSecAnalysisObjects wsDest

    Exit Sub

ErrHandler:
    ' Skip broken sheets and continue with the rest.
End Sub

Private Sub AppendTopQualityFromSourceSheet(ByVal wsSrc As Worksheet, _
                                            ByVal targetLevel As String, _
                                            ByRef recs() As TopStudentRec, _
                                            ByRef recCount As Long, _
                                            ByRef subjectTopRecs() As SubjectTopRec, _
                                            ByRef subjectTopCount As Long, _
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
    Dim isVrMc As Boolean, isAb As Boolean, hasNumericScore As Boolean
    Dim topCount As Long
    Dim topPrimaryCount As Long, topSecondaryCount As Long
    Dim g1GroupCount As Long, g2GroupCount As Long, g3GroupCount As Long, groupTotalCount As Long
    Dim fsbbGroup As String
    Dim remarksText As String
    Dim rawTopText As String
    Dim mappedBand As Long
    Dim levelMode As String
    Dim scoreValue As Double

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
    levelMode = GetLevelMode(levelCode)

    lastCol = wsSrc.Cells(1, wsSrc.Columns.count).End(xlToLeft).Column
    subjCount = 0
    For c = 1 To lastCol
        If c <> classCol Then
            header = Trim$(CStr(wsSrc.Cells(1, c).value))
            If header <> "" And IsLikelySubjectGradeColumn(header) Then
                If UCase$(levelMode) = LEVEL_MODE_LEGACY_NO_DOWNWARD Then
                    schemeKey = GetLegacySchemeFromHeader(header)
                    If schemeKey = "" Then schemeKey = GetGradeSchemeKey(wsSrc, c, header)
                Else
                    schemeKey = GetGradeSchemeKey(wsSrc, c, header)
                End If
                subjectName = StripGradeHeaderSuffix(header)
                If schemeKey <> "" And Not IsExcludedSecSubject(subjectName) _
                   And Not SubjectAlreadyAdded(subjectNames, subjCount, subjectName) Then
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
            isAb = (rawGrade = "AB" Or rawScore = "AB")
            hasNumericScore = False
            If subjectScoreCols(i) > 0 And Not isVrMc And Not isAb Then
                hasNumericScore = TryGetPercentageScore(wsSrc.Cells(r, subjectScoreCols(i)), scoreValue)
            End If

            If gradeStr <> "" Or isVrMc Or isAb Or hasNumericScore Then
                groupTotalCount = groupTotalCount + 1
                Select Case UCase$(Trim$(subjectSchemeKeys(i)))
                    Case "G1": g1GroupCount = g1GroupCount + 1
                    Case "G2": g2GroupCount = g2GroupCount + 1
                    Case "G3": g3GroupCount = g3GroupCount + 1
                End Select
            End If

            If hasNumericScore Then
                AppendSubjectTopRecord subjectTopRecs, subjectTopCount, subjectNames(i), _
                                       subjectSchemeKeys(i), className, regNo, studentName, _
                                       scoreValue, gradeStr
            End If

        Next i

        If groupTotalCount > 0 Then
            fsbbGroup = GetConfiguredLegacyGroup(className, levelCode)
            If fsbbGroup = "" Then
                fsbbGroup = ResolveFsbbGroup(g1GroupCount, g2GroupCount, g3GroupCount, groupTotalCount, groupThresholdPct)
            End If

            If fsbbGroup = "G1" Or fsbbGroup = "G2" Or fsbbGroup = "G3" Or fsbbGroup = "MIXED" Then
                topPrimaryCount = 0
                topSecondaryCount = 0
                remarksText = ""
                rawTopText = ""
                For i = 1 To subjCount
                    rawGrade = UCase$(Trim$(CStr(wsSrc.Cells(r, subjectCols(i)).value)))
                    gradeStr = NormalizeGradeForScheme(CStr(wsSrc.Cells(r, subjectCols(i)).value), subjectSchemeKeys(i))
                    rawScore = ""
                    If subjectScoreCols(i) > 0 Then rawScore = UCase$(Trim$(CStr(wsSrc.Cells(r, subjectScoreCols(i)).value)))
                    isVrMc = (rawGrade = "VR" Or rawScore = "VR" Or rawGrade = "MC" Or rawScore = "MC")

                    If Not isVrMc And gradeStr <> "" Then
                        mappedBand = GetNativeTopBand(gradeStr, subjectSchemeKeys(i))
                        Select Case mappedBand
                            Case 1
                                topPrimaryCount = topPrimaryCount + 1
                            Case 2
                                topSecondaryCount = topSecondaryCount + 1
                        End Select

                        If mappedBand > 0 Then
                            If rawTopText <> "" Then rawTopText = rawTopText & ", "
                            rawTopText = rawTopText & subjectNames(i) & " (" & gradeStr & ")"
                        End If

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
                recs(recCount).DownwardRemarks = remarksText
                recs(recCount).RawTopGrades = rawTopText
                recs(recCount).SubjectMix = "G3: " & g3GroupCount & " | G2: " & g2GroupCount & " | G1: " & g1GroupCount
            End If
        End If

NextStudent:
    Next r
    Exit Sub

FailSafe:
    ' Skip broken source sheet
End Sub

Private Function TryGetPercentageScore(ByVal scoreCell As Range, _
                                       ByRef scorePct As Double) As Boolean
    Dim rawValue As Variant
    Dim pctValue As Double
    Dim rawText As String

    rawValue = scoreCell.value
    If IsError(rawValue) Or IsEmpty(rawValue) Then Exit Function

    rawText = Trim$(CStr(rawValue))
    If Right$(rawText, 1) = "%" Then
        rawText = Trim$(Left$(rawText, Len(rawText) - 1))
        If Not IsNumeric(rawText) Then Exit Function
        pctValue = CDbl(rawText)
    Else
        If Not IsNumeric(rawValue) Then Exit Function
        pctValue = CDbl(rawValue)
        If InStr(1, scoreCell.NumberFormat, "%", vbBinaryCompare) > 0 Then pctValue = pctValue * 100#
    End If
    If pctValue < 0# Or pctValue > 100# Then Exit Function

    scorePct = pctValue
    TryGetPercentageScore = True
End Function

Private Sub AppendSubjectTopRecord(ByRef recs() As SubjectTopRec, _
                                   ByRef recCount As Long, _
                                   ByVal subjectName As String, _
                                   ByVal schemeKey As String, _
                                   ByVal className As String, _
                                   ByVal regNo As String, _
                                   ByVal studentName As String, _
                                   ByVal scorePct As Double, _
                                   ByVal gradeText As String)
    recCount = recCount + 1
    If recCount = 1 Then
        ReDim recs(1 To 1)
    Else
        ReDim Preserve recs(1 To recCount)
    End If

    With recs(recCount)
        .SubjectName = subjectName
        .SchemeKey = UCase$(Trim$(schemeKey))
        .ClassName = className
        .RegNo = regNo
        .StudentName = studentName
        .ScorePct = scorePct
        .GradeText = gradeText
    End With
End Sub

Private Function WriteSubjectTopPerformersSection(ByVal wsOut As Worksheet, _
                                                  ByVal startRow As Long, _
                                                  ByRef recs() As SubjectTopRec, _
                                                  ByVal recCount As Long) As Long
    Dim subjectMap As Object
    Dim eligibleSubjectMap As Object
    Dim keys As Variant
    Dim subjectKey As Variant
    Dim idx() As Long
    Dim idxCount As Long
    Dim i As Long, j As Long
    Dim r As Long, rankNo As Long
    Dim tableHeaderRow As Long, lastDataRow As Long
    Dim key As String, tmpKey As String
    Dim currentScore As Double
    Dim subjectName As String, schemeKey As String
    Dim minSubjectN As Long

    minSubjectN = GetMinSubjectN()

    With wsOut.Range(wsOut.Cells(startRow, 1), wsOut.Cells(startRow, 6))
        .Merge
        .value = "Top 3 in Each Subject by Percentage (ties included)"
        .Font.Bold = True
        .Font.Size = 12
        .Font.Color = RGB(255, 255, 255)
        .Interior.Color = RGB(31, 78, 121)
        .VerticalAlignment = xlCenter
        .Borders.LineStyle = xlContinuous
        .Borders.Color = RGB(31, 78, 121)
    End With
    wsOut.Rows(startRow).RowHeight = 24
    With wsOut.Range(wsOut.Cells(startRow + 1, 1), wsOut.Cells(startRow + 1, 6))
        .Merge
        .value = "Subjects with fewer than " & minSubjectN & _
                 " valid numeric scores are excluded (Settings!L6)."
        .Font.Italic = True
        .Font.Size = 10
        .Font.Color = RGB(89, 89, 89)
        .VerticalAlignment = xlCenter
    End With
    wsOut.Rows(startRow + 1).RowHeight = 20
    startRow = startRow + 3

    If recCount = 0 Then
        wsOut.Cells(startRow, 1).value = "(No valid numeric subject scores found.)"
        wsOut.Cells(startRow, 1).Font.Italic = True
        WriteSubjectTopPerformersSection = startRow + 2
        Exit Function
    End If

    r = startRow

    Set subjectMap = CreateObject("Scripting.Dictionary")
    subjectMap.CompareMode = vbTextCompare
    For i = 1 To recCount
        key = SubjectTopSortKey(recs(i).SubjectName, recs(i).SchemeKey)
        If subjectMap.Exists(key) Then
            subjectMap(key) = CLng(subjectMap(key)) + 1
        Else
            subjectMap.Add key, 1
        End If
    Next i

    Set eligibleSubjectMap = CreateObject("Scripting.Dictionary")
    eligibleSubjectMap.CompareMode = vbTextCompare
    For Each subjectKey In subjectMap.Keys
        If CLng(subjectMap(subjectKey)) >= minSubjectN Then _
            eligibleSubjectMap.Add CStr(subjectKey), CStr(subjectKey)
    Next subjectKey

    If eligibleSubjectMap.Count = 0 Then
        wsOut.Cells(startRow, 1).value = "(No subjects meet the minimum candidature of " & minSubjectN & ".)"
        wsOut.Cells(startRow, 1).Font.Italic = True
        WriteSubjectTopPerformersSection = startRow + 2
        Exit Function
    End If

    keys = eligibleSubjectMap.Keys

    For i = LBound(keys) To UBound(keys) - 1
        For j = i + 1 To UBound(keys)
            If StrComp(CStr(keys(j)), CStr(keys(i)), vbTextCompare) < 0 Then
                tmpKey = CStr(keys(i)): keys(i) = keys(j): keys(j) = tmpKey
            End If
        Next j
    Next i

    For i = LBound(keys) To UBound(keys)
        idxCount = 0
        Erase idx
        For j = 1 To recCount
            key = SubjectTopSortKey(recs(j).SubjectName, recs(j).SchemeKey)
            If StrComp(key, CStr(keys(i)), vbTextCompare) = 0 Then
                idxCount = idxCount + 1
                If idxCount = 1 Then
                    ReDim idx(1 To 1)
                Else
                    ReDim Preserve idx(1 To idxCount)
                End If
                idx(idxCount) = j
            End If
        Next j

        SortSubjectTopIndexes recs, idx, idxCount
        subjectName = recs(idx(1)).SubjectName
        schemeKey = recs(idx(1)).SchemeKey

        With wsOut.Range(wsOut.Cells(r, 1), wsOut.Cells(r, 6))
            .Merge
            .value = subjectName & " [" & schemeKey & "] - Top 3 by Percentage"
            .Font.Bold = True
            .Font.Size = 11
            .HorizontalAlignment = xlLeft
            .VerticalAlignment = xlCenter
            .IndentLevel = 1
            .Borders.LineStyle = xlContinuous
            .Borders.Weight = xlThin
        End With
        StyleSubjectTopBlockTitle wsOut.Range(wsOut.Cells(r, 1), wsOut.Cells(r, 6)), schemeKey
        wsOut.Rows(r).RowHeight = 22
        r = r + 1

        tableHeaderRow = r
        wsOut.Cells(r, 1).value = "Rank"
        wsOut.Cells(r, 2).value = "Class"
        wsOut.Cells(r, 3).value = "RegNo"
        wsOut.Cells(r, 4).value = "Name"
        wsOut.Cells(r, 5).value = "Score %"
        wsOut.Cells(r, 6).value = "Grade"
        With wsOut.Range(wsOut.Cells(r, 1), wsOut.Cells(r, 6))
            .Font.Bold = True
            .Interior.Color = RGB(242, 246, 250)
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
        End With
        wsOut.Rows(r).RowHeight = 20
        r = r + 1

        rankNo = 0
        currentScore = -1#
        For j = 1 To idxCount
            If j = 1 Or Abs(recs(idx(j)).ScorePct - currentScore) > 0.0000001 Then rankNo = j
            If rankNo > 3 Then Exit For
            currentScore = recs(idx(j)).ScorePct

            wsOut.Cells(r, 1).value = rankNo
            wsOut.Cells(r, 2).value = recs(idx(j)).ClassName
            wsOut.Cells(r, 3).value = recs(idx(j)).RegNo
            wsOut.Cells(r, 4).value = recs(idx(j)).StudentName
            wsOut.Cells(r, 5).value = recs(idx(j)).ScorePct
            wsOut.Cells(r, 5).NumberFormat = "0.0"
            wsOut.Cells(r, 6).value = recs(idx(j)).GradeText
            StyleSubjectTopRankRow wsOut.Range(wsOut.Cells(r, 1), wsOut.Cells(r, 6)), rankNo
            wsOut.Range(wsOut.Cells(r, 1), wsOut.Cells(r, 3)).HorizontalAlignment = xlCenter
            wsOut.Range(wsOut.Cells(r, 5), wsOut.Cells(r, 6)).HorizontalAlignment = xlCenter
            wsOut.Cells(r, 4).HorizontalAlignment = xlLeft
            wsOut.Rows(r).RowHeight = 20
            r = r + 1
        Next j
        lastDataRow = r - 1
        With wsOut.Range(wsOut.Cells(tableHeaderRow, 1), wsOut.Cells(lastDataRow, 6)).Borders
            .LineStyle = xlContinuous
            .Color = RGB(205, 215, 225)
            .Weight = xlThin
        End With
        r = r + 1
    Next i

    WriteSubjectTopPerformersSection = r + 1
End Function

Private Function SubjectTopSortKey(ByVal subjectName As String, _
                                   ByVal schemeKey As String) As String
    Dim schemeOrder As String
    Select Case UCase$(Trim$(schemeKey))
        Case "G3": schemeOrder = "1"
        Case "G2": schemeOrder = "2"
        Case "G1": schemeOrder = "3"
        Case Else: schemeOrder = "9"
    End Select
    SubjectTopSortKey = schemeOrder & "|" & UCase$(Trim$(subjectName)) & "|" & UCase$(Trim$(schemeKey))
End Function

Private Sub StyleSubjectTopBlockTitle(ByVal titleRange As Range, _
                                      ByVal schemeKey As String)
    Select Case UCase$(Trim$(schemeKey))
        Case "G3"
            titleRange.Interior.Color = RGB(221, 235, 247)
            titleRange.Font.Color = RGB(31, 78, 121)
        Case "G2"
            titleRange.Interior.Color = RGB(226, 240, 217)
            titleRange.Font.Color = RGB(55, 86, 35)
        Case "G1"
            titleRange.Interior.Color = RGB(255, 242, 204)
            titleRange.Font.Color = RGB(127, 96, 0)
        Case Else
            titleRange.Interior.Color = RGB(242, 242, 242)
            titleRange.Font.Color = RGB(89, 89, 89)
    End Select
End Sub

Private Sub StyleSubjectTopRankRow(ByVal rowRange As Range, ByVal rankNo As Long)
    Select Case rankNo
        Case 1: rowRange.Interior.Color = RGB(255, 242, 204)
        Case 2: rowRange.Interior.Color = RGB(242, 242, 242)
        Case 3: rowRange.Interior.Color = RGB(252, 228, 214)
    End Select
End Sub

Private Sub SortSubjectTopIndexes(ByRef recs() As SubjectTopRec, _
                                  ByRef idx() As Long, _
                                  ByVal idxCount As Long)
    Dim i As Long, j As Long, tmp As Long
    For i = 1 To idxCount - 1
        For j = i + 1 To idxCount
            If recs(idx(j)).ScorePct > recs(idx(i)).ScorePct Or _
               (Abs(recs(idx(j)).ScorePct - recs(idx(i)).ScorePct) < 0.0000001 And _
                StrComp(recs(idx(j)).StudentName, recs(idx(i)).StudentName, vbTextCompare) < 0) Then
                tmp = idx(i): idx(i) = idx(j): idx(j) = tmp
            End If
        Next j
    Next i
End Sub

Private Function GetNativeTopBand(ByVal gradeStr As String, _
                                  ByVal schemeKey As String) As Long
    Dim g As String
    g = UCase$(Trim$(gradeStr))

    Select Case UCase$(Trim$(schemeKey))
        Case "G3"
            If g = "A1" Then
                GetNativeTopBand = 1
            ElseIf g = "A2" Then
                GetNativeTopBand = 2
            End If
        Case "G2"
            If g = "1" Then
                GetNativeTopBand = 1
            ElseIf g = "2" Then
                GetNativeTopBand = 2
            End If
        Case "G1"
            If g = "A" Then
                GetNativeTopBand = 1
            ElseIf g = "B" Then
                GetNativeTopBand = 2
            End If
    End Select
End Function

Private Function WriteTopGroupSection(ByVal wsOut As Worksheet, _
                                      ByVal startRow As Long, _
                                      ByVal levelCode As String, _
                                      ByVal groupCode As String, _
                                      ByVal topN As Long, _
                                      ByRef recs() As TopStudentRec, _
                                      ByVal recCount As Long, _
                                      ByVal levelMode As String) As Long
    Dim idx() As Long
    Dim idxCount As Long
    Dim i As Long, j As Long, tmp As Long
    Dim cutoffTop As Long, cutoffPrimary As Long, cutoffSecondary As Long
    Dim r As Long, tableHeaderRow As Long
    Dim displayGroup As String
    Dim sectionTitle As String

    displayGroup = MapGroupLabelForMode(groupCode, levelMode)
    If UCase$(levelMode) = LEVEL_MODE_LEGACY_NO_DOWNWARD Then
        sectionTitle = "Top Performers in " & displayGroup & " Subjects"
    ElseIf UCase$(groupCode) = "MIXED" Then
        sectionTitle = "Top Performers with Mixed Subject Levels"
    Else
        sectionTitle = "Top Performers in Predominantly " & UCase$(groupCode) & " Subjects"
    End If
    With wsOut.Range(wsOut.Cells(startRow, 1), wsOut.Cells(startRow, 10))
        .Merge
        .value = sectionTitle & " (Top " & topN & ", ties included)"
        .Font.Bold = True
        .Font.Color = RGB(79, 33, 33)
        .Interior.Color = RGB(252, 228, 214)
    End With
    startRow = startRow + 1

    For i = 1 To recCount
        If UCase$(recs(i).GroupCode) = UCase$(groupCode) And recs(i).TopCount > 0 Then
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

    tableHeaderRow = startRow
    wsOut.Cells(startRow, 1).value = "Level"
    wsOut.Cells(startRow, 2).value = "Class"
    wsOut.Cells(startRow, 3).value = "RegNo"
    wsOut.Cells(startRow, 4).value = "Name"
    If UCase$(levelMode) = LEVEL_MODE_LEGACY_NO_DOWNWARD Then
        wsOut.Cells(startRow, 5).value = "Group"
    Else
        wsOut.Cells(startRow, 5).value = "Predominant Group"
    End If
    wsOut.Cells(startRow, 6).NumberFormat = "@"
    wsOut.Cells(startRow, 7).NumberFormat = "@"
    wsOut.Cells(startRow, 8).NumberFormat = "@"
    wsOut.Cells(startRow, 6).value = "Top Grades" & vbLf & "(1st + 2nd Band)"
    wsOut.Cells(startRow, 7).value = "1st Band" & vbLf & "(A1 / 1 / A)"
    wsOut.Cells(startRow, 8).value = "2nd Band" & vbLf & "(A2 / 2 / B)"
    wsOut.Cells(startRow, 9).value = "Subject Mix"
    wsOut.Cells(startRow, 10).value = "Native Top Grades"
    With wsOut.Range(wsOut.Cells(startRow, 1), wsOut.Cells(startRow, 10))
        .Font.Bold = True
        .WrapText = True
        .VerticalAlignment = xlCenter
        .Interior.Color = RGB(250, 242, 238)
    End With
    wsOut.Rows(startRow).RowHeight = 32
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
        wsOut.Cells(r, 5).value = MapGroupLabelForMode(recs(idx(i)).GroupCode, levelMode)
        wsOut.Cells(r, 6).value = recs(idx(i)).TopCount
        wsOut.Cells(r, 7).value = recs(idx(i)).TopPrimaryCount
        wsOut.Cells(r, 8).value = recs(idx(i)).TopSecondaryCount
        wsOut.Cells(r, 9).value = recs(idx(i)).SubjectMix
        wsOut.Cells(r, 10).value = recs(idx(i)).RawTopGrades
        r = r + 1
    Next i

    With wsOut.Range(wsOut.Cells(tableHeaderRow, 1), wsOut.Cells(r - 1, 10)).Borders
        .LineStyle = xlContinuous
        .Color = RGB(210, 200, 195)
        .Weight = xlThin
    End With

    WriteTopGroupSection = r + 1
End Function

Private Sub PrepareTopQualitySheet(ByVal wsOut As Worksheet, ByVal levelCode As String, _
                                   ByVal assessmentLabel As String, ByVal yearText As String)
    Dim explainer As String
    Dim levelMode As String
    Dim thresholdPct As Double

    With wsOut.Range("A1:D1")
        .Merge
        .value = levelCode & " " & assessmentLabel & IIf(yearText <> "", " " & yearText, "") & _
                 " - Top Performers"
        .Font.Bold = True
        .Font.Size = 14
        .Font.Color = RGB(31, 78, 121)
        .VerticalAlignment = xlCenter
    End With
    wsOut.Rows(1).RowHeight = 26

    levelMode = GetLevelMode(levelCode)
    thresholdPct = GetGroupThresholdPercent()
    If UCase$(levelMode) = LEVEL_MODE_LEGACY_NO_DOWNWARD Then
        explainer = "Part 1 lists the top 3 students in each subject by numeric score percentage; ties are included." & vbLf & _
                    "Part 2 lists the top 5 performers in each mapped EX/NA/NT group. Ranking uses native top grades: " & _
                    "G3/EX A1-A2, G2/NA 1-2 and G1/NT A-B; no downward conversion is applied." & vbLf & _
                    "AB, MC and VR are excluded from percentage ranking."
    Else
        explainer = "Part 1 lists the top 3 students in each subject by numeric score percentage; ties are included." & vbLf & _
                    "Part 2 lists the top 5 performers for predominantly G3, G2 and G1 subject loads, plus Mixed where needed. " & _
                    "Predominant means at least " & Format$(thresholdPct, "0.#") & "% of registered subjects at that level." & vbLf & _
                    "Ranking uses native top grades across the student's subjects: G3 A1-A2, G2 1-2 and G1 A-B; " & _
                    "no downward conversion is applied. AB, MC and VR count toward subject mix but not performance ranking."
    End If

    With wsOut.Range("A2:J2")
        .Merge
        .value = explainer
        .WrapText = True
        .Font.Italic = True
        .VerticalAlignment = xlTop
    End With
    wsOut.Rows(2).RowHeight = 58
End Sub

Private Sub FormatTopQualitySheet(ByVal wsOut As Worksheet, ByVal lastRow As Long)
    wsOut.Columns("A:J").AutoFit
    wsOut.Columns("A").ColumnWidth = 8
    wsOut.Columns("B").ColumnWidth = 18
    wsOut.Columns("C").ColumnWidth = 8
    wsOut.Columns("D").ColumnWidth = 28
    wsOut.Columns("E").ColumnWidth = 17
    wsOut.Columns("F").ColumnWidth = 14
    wsOut.Columns("G").ColumnWidth = 12
    wsOut.Columns("H").ColumnWidth = 12
    wsOut.Columns("I").ColumnWidth = 28
    wsOut.Columns("C").HorizontalAlignment = xlCenter
    wsOut.Columns("E:H").HorizontalAlignment = xlCenter
    wsOut.Columns("I").WrapText = True
    wsOut.Columns("J").ColumnWidth = 45
    wsOut.Columns("J").WrapText = True
End Sub

Private Function GetLevelMode(ByVal levelCode As String) As String
    Dim ws As Worksheet
    Dim r As Long
    Dim lvl As String, modeVal As String

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("Settings")
    On Error GoTo 0

    If ws Is Nothing Then
        GetLevelMode = LEVEL_MODE_AUTO_FSBB
        Exit Function
    End If

    ' Settings table (optional): N2:O20
    '   N = Level (e.g. S4), O = Mode (AUTO_FSBB / LEGACY_NO_DOWNWARD)
    For r = 2 To 20
        lvl = UCase$(Trim$(CStr(ws.Cells(r, "N").value)))
        modeVal = UCase$(Trim$(CStr(ws.Cells(r, "O").value)))
        If lvl = UCase$(Trim$(levelCode)) Then
            If modeVal = LEVEL_MODE_LEGACY_NO_DOWNWARD Then
                GetLevelMode = LEVEL_MODE_LEGACY_NO_DOWNWARD
            Else
                GetLevelMode = LEVEL_MODE_AUTO_FSBB
            End If
            Exit Function
        End If
    Next r

    GetLevelMode = LEVEL_MODE_AUTO_FSBB
End Function

Private Function GetLegacySchemeFromHeader(ByVal header As String) As String
    Dim h As String
    h = UCase$(Replace(Trim$(header), " ", ""))

    ' Legacy tracks:
    '   - EX / Express / O-Level -> G3 equivalent
    '   - NA / N(A)             -> G2 equivalent
    '   - NT / N(T)             -> G1 equivalent
    If InStr(h, "N(T)") > 0 Or InStr(h, "-NT") > 0 _
       Or InStr(h, "(NT)") > 0 Or InStr(h, "NORMALTECH") > 0 Then
        GetLegacySchemeFromHeader = "G1"
    ElseIf InStr(h, "N(A)") > 0 Or InStr(h, "-NA") > 0 _
       Or InStr(h, "(NA)") > 0 Or InStr(h, "NORMALACADEMIC") > 0 Then
        GetLegacySchemeFromHeader = "G2"
    ElseIf InStr(h, "-EX") > 0 Or InStr(h, "(EX)") > 0 _
       Or InStr(h, "EXPRESS") > 0 Or InStr(h, "-O") > 0 _
       Or InStr(h, "(O)") > 0 Or InStr(h, "OLEVEL") > 0 Then
        GetLegacySchemeFromHeader = "G3"
    End If
End Function

Private Function MapGroupLabelForMode(ByVal fsbbGroup As String, ByVal levelMode As String) As String
    Dim g As String
    g = UCase$(Trim$(fsbbGroup))

    If UCase$(Trim$(levelMode)) = LEVEL_MODE_LEGACY_NO_DOWNWARD Then
        Select Case g
            Case "G3": MapGroupLabelForMode = "EX"
            Case "G2": MapGroupLabelForMode = "NA"
            Case "G1": MapGroupLabelForMode = "NT"
            Case Else: MapGroupLabelForMode = g
        End Select
    Else
        MapGroupLabelForMode = g
    End If
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
    Dim attemptedCount As Long, passCount As Long, failCount As Long, distinctionCount As Long
    Dim outRow As Long
    Dim riskBand As String, failedSubjects As String, attemptedSubjects As String
    Dim abSubjects As String
    Dim vrSubjects As String, rawGrade As String, rawScore As String
    Dim subjectName As String
    Dim isVrSubject As Boolean, isAbSubject As Boolean
    Dim g1Taken As Long, g2Taken As Long, g3Taken As Long
    Dim g1GroupCount As Long, g2GroupCount As Long, g3GroupCount As Long
    Dim groupTotalCount As Long
    Dim countedOutcomeCount As Long
    Dim fsbbGroup As String
    Dim displayGroup As String
    Dim groupThresholdPct As Double
    Dim levelMode As String

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
    levelMode = GetLevelMode(levelCode)

    lastCol = wsSrc.Cells(1, wsSrc.Columns.count).End(xlToLeft).Column
    subjCount = 0

    For c = 1 To lastCol
        If c <> classCol Then
            header = Trim$(CStr(wsSrc.Cells(1, c).value))
            If header <> "" And IsLikelySubjectGradeColumn(header) Then
                If UCase$(levelMode) = LEVEL_MODE_LEGACY_NO_DOWNWARD Then
                    schemeKey = GetLegacySchemeFromHeader(header)
                    If schemeKey = "" Then schemeKey = GetGradeSchemeKey(wsSrc, c, header)
                Else
                    schemeKey = GetGradeSchemeKey(wsSrc, c, header)
                End If
                subjectName = StripGradeHeaderSuffix(header)
                If schemeKey <> "" And Not IsExcludedSecSubject(subjectName) _
                   And Not SubjectAlreadyAdded(subjectNames, subjCount, subjectName) Then
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
        distinctionCount = 0
        failedSubjects = ""
        attemptedSubjects = ""
        vrSubjects = ""
        abSubjects = ""
        g1Taken = 0
        g2Taken = 0
        g3Taken = 0
        g1GroupCount = 0
        g2GroupCount = 0
        g3GroupCount = 0
        groupTotalCount = 0
        countedOutcomeCount = 0

        For i = 1 To subjCount
            rawGrade = UCase$(Trim$(CStr(wsSrc.Cells(r, subjectCols(i)).value)))
            gradeStr = NormalizeGradeForScheme(CStr(wsSrc.Cells(r, subjectCols(i)).value), subjectSchemeKeys(i))
            rawScore = ""
            If subjectScoreCols(i) > 0 Then
                rawScore = UCase$(Trim$(CStr(wsSrc.Cells(r, subjectScoreCols(i)).value)))
            End If

            ' Predominant subject mix includes registered subjects with a
            ' result or an AB/VR/MC status; absence does not change G-level.
            If gradeStr <> "" Or rawGrade = "VR" Or rawScore = "VR" _
               Or rawGrade = "MC" Or rawScore = "MC" _
               Or rawGrade = "AB" Or rawScore = "AB" Then
                groupTotalCount = groupTotalCount + 1
                Select Case UCase$(Trim$(subjectSchemeKeys(i)))
                    Case "G1": g1GroupCount = g1GroupCount + 1
                    Case "G2": g2GroupCount = g2GroupCount + 1
                    Case "G3": g3GroupCount = g3GroupCount + 1
                End Select
            End If

            isVrSubject = (rawGrade = "VR" Or rawScore = "VR" Or _
                           rawGrade = "MC" Or rawScore = "MC")
            If isVrSubject Then
                If vrSubjects <> "" Then vrSubjects = vrSubjects & ", "
                If rawGrade = "MC" Or rawScore = "MC" Then
                    vrSubjects = vrSubjects & subjectNames(i) & " (MC)"
                Else
                    vrSubjects = vrSubjects & subjectNames(i) & " (VR)"
                End If
                GoTo NextSubject
            End If

            isAbSubject = (rawGrade = "AB" Or rawScore = "AB")
            If isAbSubject Then
                If abSubjects <> "" Then abSubjects = abSubjects & ", "
                abSubjects = abSubjects & subjectNames(i)
                failCount = failCount + 1
                If failedSubjects <> "" Then failedSubjects = failedSubjects & ", "
                failedSubjects = failedSubjects & subjectNames(i) & " (AB)"
                countedOutcomeCount = countedOutcomeCount + 1
                GoTo NextSubject
            End If

            If gradeStr <> "" Then
                attemptedCount = attemptedCount + 1
                countedOutcomeCount = countedOutcomeCount + 1
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
                If IsDistinctionGradeByScheme(gradeStr, subjectSchemeKeys(i)) Then
                    distinctionCount = distinctionCount + 1
                End If
            End If
NextSubject:
        Next i

        ' Keep students whose release contains only VR/MC for follow-up
        ' follow-up, but classify them as MONITOR rather than OK.
        ' Completely blank rows remain excluded. AB is a counted
        ' outcome and contributes one failure.
        If countedOutcomeCount > 0 Or groupTotalCount > 0 Then
            If countedOutcomeCount = 0 Then
                riskBand = "MONITOR"
            ElseIf failCount >= atRiskFailThreshold Then
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
            fsbbGroup = GetConfiguredLegacyGroup(className, levelCode)
            If fsbbGroup = "" Then
                fsbbGroup = ResolveFsbbGroup(g1GroupCount, g2GroupCount, g3GroupCount, groupTotalCount, groupThresholdPct)
            End If
            wsOut.Cells(outRow, 5).value = attemptedCount
            wsOut.Cells(outRow, 6).value = passCount
            wsOut.Cells(outRow, 7).value = failCount
            wsOut.Cells(outRow, 8).value = failedSubjects
            wsOut.Cells(outRow, 9).value = riskBand
            wsOut.Cells(outRow, 10).value = RiskBandRank(riskBand)
            wsOut.Cells(outRow, 12).value = attemptedSubjects
            wsOut.Cells(outRow, 13).value = vrSubjects
            wsOut.Cells(outRow, 14).value = abSubjects
            ' Legacy Sec 5 is the 5NA cohort even though its subjects use
            ' the G3/O-Level grade scheme. Keep the cohort label separate
            ' from the subject scheme used for pass/distinction calculations.
            If UCase$(Trim$(levelCode)) = "S5" Then
                displayGroup = "NA"
            Else
                displayGroup = MapGroupLabelForMode(fsbbGroup, levelMode)
            End If
            wsOut.Cells(outRow, 15).value = displayGroup
            wsOut.Cells(outRow, 16).value = distinctionCount

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
    Dim reportSheets() As String
    Dim reportCount As Long, i As Long
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

    ' Rows 15 onward are reserved for the Class Analysis menu below.
    wsNav.Range(wsNav.Cells(startRow, startCol), wsNav.Cells(startRow + 11, startCol + 3)).Clear
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

    CollectVersionedReportSheetNames "AtRisk_", reportSheets, reportCount
    If reportCount = 0 Then
        wsNav.Cells(rowPtr, startCol).value = "No assessment-specific at-risk reports built."
        wsNav.Cells(rowPtr, startCol).Font.Italic = True
    Else
        For i = 1 To reportCount
            CreateAtRiskNavButton wsNav, reportSheets(i), ReportNavigationLabel(reportSheets(i), "AtRisk_"), rowPtr, startCol
            rowPtr = rowPtr + 2
        Next i
    End If

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
    btnWidth = wsNav.Columns(firstCol).Resize(, 5).Width * NAV_BTN_WIDTH_FACTOR
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

Private Sub CollectVersionedReportSheetNames(ByVal reportPrefix As String, _
                                             ByRef sheetNames() As String, _
                                             ByRef sheetCount As Long)
    Dim ws As Worksheet, i As Long, j As Long, tmp As String
    Dim parts As Variant

    For Each ws In ThisWorkbook.Worksheets
        If Left$(ws.Name, Len(reportPrefix)) = reportPrefix Then
            parts = Split(ws.Name, "_")
            ' Versioned form: Prefix_S1_Assessment_Year
            If UBound(parts) >= 3 Then
                sheetCount = sheetCount + 1
                ReDim Preserve sheetNames(1 To sheetCount)
                sheetNames(sheetCount) = ws.Name
            End If
        End If
    Next ws

    For i = 1 To sheetCount - 1
        For j = i + 1 To sheetCount
            If VersionedReportSheetBefore(sheetNames(j), sheetNames(i)) Then
                tmp = sheetNames(i): sheetNames(i) = sheetNames(j): sheetNames(j) = tmp
            End If
        Next j
    Next i
End Sub

Private Function VersionedReportSheetBefore(ByVal aName As String, ByVal bName As String) As Boolean
    Dim aParts As Variant, bParts As Variant
    Dim ao As Long, bo As Long
    aParts = Split(aName, "_"): bParts = Split(bName, "_")
    If StrComp(CStr(aParts(1)), CStr(bParts(1)), vbTextCompare) <> 0 Then
        VersionedReportSheetBefore = (StrComp(CStr(aParts(1)), CStr(bParts(1)), vbTextCompare) < 0)
        Exit Function
    End If
    ao = SecReportAssessmentOrder(CStr(aParts(2))): bo = SecReportAssessmentOrder(CStr(bParts(2)))
    If ao <> bo Then VersionedReportSheetBefore = (ao < bo): Exit Function
    If StrComp(CStr(aParts(2)), CStr(bParts(2)), vbTextCompare) <> 0 Then
        VersionedReportSheetBefore = (StrComp(CStr(aParts(2)), CStr(bParts(2)), vbTextCompare) < 0)
        Exit Function
    End If
    VersionedReportSheetBefore = (CStr(aParts(3)) > CStr(bParts(3)))
End Function

Private Function ReportNavigationLabel(ByVal sheetName As String, _
                                       ByVal reportPrefix As String) As String
    Dim labelText As String
    labelText = Replace(Mid$(sheetName, Len(reportPrefix) + 1), "_", " ")
    If reportPrefix = "AtRisk_" Then
        ReportNavigationLabel = labelText & " At Risk"
    Else
        ReportNavigationLabel = labelText & " Top Students"
    End If
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

Private Sub PrepareAtRiskSheet(ByVal wsOut As Worksheet, ByVal levelCode As String, ByVal threshold As Long, _
                               ByVal assessmentLabel As String, ByVal yearText As String)
    wsOut.Range("A1").value = levelCode & " " & assessmentLabel & IIf(yearText <> "", " " & yearText, "") & _
                              " - Students At Risk"
    wsOut.Range("A1").Font.Bold = True
    wsOut.Range("A1").Font.Size = 14

    wsOut.Range("A2").value = "At-risk rule: Failed Subjects >= " & threshold & _
                              " (AB considered 0 marks; all-VR/MC students retained as MONITOR)"
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
    wsOut.Cells(4, 13).value = "VR/MC Subjects"
    wsOut.Cells(4, 14).value = "AB Subjects"
    wsOut.Cells(4, 15).value = "Group"
    wsOut.Cells(4, 16).value = "Distinctions"
    wsOut.Rows(4).Font.Bold = True
End Sub

Private Sub FinalizeAtRiskSheet(ByVal wsOut As Worksheet, ByVal lastRow As Long)
    Dim sortRange As Range
    Dim rngTable As Range

    If lastRow >= 5 Then
        Set sortRange = wsOut.Range("A4:P" & lastRow)
        sortRange.Sort Key1:=wsOut.Range("J5"), Order1:=xlAscending, _
                       Key2:=wsOut.Range("G5"), Order2:=xlDescending, _
                       Key3:=wsOut.Range("D5"), Order3:=xlAscending, _
                       Header:=xlYes
    End If

    wsOut.Columns("A:P").AutoFit
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
    wsOut.Columns("N").ColumnWidth = 25
    wsOut.Columns("N").WrapText = True
    wsOut.Columns("O").ColumnWidth = 10
    wsOut.Columns("O").HorizontalAlignment = xlCenter
    wsOut.Columns("P").ColumnWidth = 10
    wsOut.Columns("P").HorizontalAlignment = xlCenter
    wsOut.Columns("J").EntireColumn.Hidden = True
    wsOut.Columns("O:P").EntireColumn.Hidden = True
    wsOut.Range("E4:G4").WrapText = True
    wsOut.Range("A4:P4").VerticalAlignment = xlCenter

    If lastRow >= 4 Then
        Set rngTable = wsOut.Range("A4:P" & lastRow)
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
    Dim titleRow As Long, visualEndRow As Long, chartRowCount As Long

    Dim total As Long
    Dim passCount As Long, failCount As Long, topCount As Long
    Dim meanValue As Double

    Dim colNo As Long, colPctPass As Long, colPctFail As Long, colPctTop As Long, colMean As Long
    Dim colTableLast As Long, chartStartCol As Long
    Dim showMean As Boolean

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

    Const MIN_SEC_CHART_ROWS As Long = 6

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
    showMean = (UCase$(Trim$(schemeKey)) <> "G1")
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
    If showMean Then
        colMean = destColFirst + numBands + 5
        colTableLast = colMean
    Else
        colMean = 0
        colTableLast = colPctTop
    End If
    chartStartCol = colTableLast + 2

    With wsDest
        .Cells(destRowHeader, destColFirst).value = "Class"

        For j = 1 To numBands
            .Cells(destRowHeader, destColFirst + j).value = gradeLabels(j)
        Next j

        .Cells(destRowHeader, colNo).value = "No."
        .Cells(destRowHeader, colPctPass).value = pctPassLabel
        .Cells(destRowHeader, colPctFail).value = pctFailLabel
        .Cells(destRowHeader, colPctTop).value = pctTopLabel
        If showMean Then .Cells(destRowHeader, colMean).value = meanLabel
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

        If total > 0 And showMean Then
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
                If showMean Then .Cells(rowPtr, colMean).value = Round(meanValue, 1)
            Else
                .Cells(rowPtr, colPctPass).ClearContents
                .Cells(rowPtr, colPctFail).ClearContents
                .Cells(rowPtr, colPctTop).ClearContents
                If showMean Then .Cells(rowPtr, colMean).ClearContents
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

    If total > 0 And showMean Then
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
            If showMean Then .Cells(cohortRow, colMean).value = Round(meanValue, 1)
        Else
            .Cells(cohortRow, colPctPass).ClearContents
            .Cells(cohortRow, colPctFail).ClearContents
            .Cells(cohortRow, colPctTop).ClearContents
            If showMean Then .Cells(cohortRow, colMean).ClearContents
        End If
    End With

    ColourSubjectRow wsDest, cohortRow, destColFirst, numBands, topMaxIdx, failMinIdx, _
                    colPctPass, colPctFail, colPctTop, colMean

    Set rngTable = wsDest.Range(wsDest.Cells(destRowHeader, destColFirst), _
                                wsDest.Cells(cohortRow, colTableLast))

    With rngTable.Borders
        .LineStyle = xlContinuous
        .Color = RGB(200, 200, 200)
        .Weight = xlThin
    End With

    wsDest.Range(wsDest.Cells(destRowHeader + 1, colPctPass), _
                 wsDest.Cells(cohortRow, colPctTop)).NumberFormat = "0.0"
    If showMean Then
        wsDest.Range(wsDest.Cells(destRowHeader + 1, colMean), _
                     wsDest.Cells(cohortRow, colMean)).NumberFormat = "0.0"
    End If

    wsDest.Columns(destColFirst + 1).Resize(, colTableLast - destColFirst).AutoFit
    wsDest.Columns(destColFirst).ColumnWidth = 15

    Set rngCohortRow = wsDest.Range(wsDest.Cells(cohortRow, destColFirst), _
                                    wsDest.Cells(cohortRow, colTableLast))
    rngCohortRow.Interior.Color = RGB(255, 242, 204)
    rngCohortRow.Font.Bold = True

    Set rngHeader = wsDest.Range(wsDest.Cells(destRowHeader, destColFirst + 1), _
                                 wsDest.Cells(destRowHeader, destColFirst + numBands))
    Set rngData = wsDest.Range(wsDest.Cells(cohortRow, destColFirst + 1), _
                               wsDest.Cells(cohortRow, destColFirst + numBands))

    ' Excel charts have a practical minimum height for axes and data labels,
    ' especially on Excel for Mac. A one-class table naturally occupies only
    ' four rows including its title; a two-class table occupies five. Reserve
    ' at least six rows for the visual block so the next subject is positioned
    ' after the chart rather than merely after the shorter table.
    titleRow = destRowHeader - 1
    chartRowCount = cohortRow - titleRow + 1
    If chartRowCount < MIN_SEC_CHART_ROWS Then chartRowCount = MIN_SEC_CHART_ROWS
    visualEndRow = titleRow + chartRowCount - 1

    ' Report the larger table/chart footprint before creating floating
    ' objects so a rendering error cannot make the next subject reuse it.
    outEndRow = visualEndRow

    leftPos = wsDest.Columns(chartStartCol).Left
    ' Strict alignment rule: the chart and narrative panel start at the top
    ' of the table header, not at the separate subject-title row.
    topPos = ExactSecRowTop(wsDest, destRowHeader)
    chartWidth = wsDest.Columns(chartStartCol).Resize(, 6).Width
    chartHeight = ExactSecRowsHeight(wsDest, destRowHeader, visualEndRow)

    Set co = wsDest.ChartObjects.Add(leftPos, topPos, chartWidth, chartHeight)
    co.Name = "SecChart_R" & CStr(destRowHeader) & _
              "_E" & CStr(visualEndRow) & "_C" & CStr(chartStartCol)
    co.Placement = xlMove
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

    SetSecChartPlotArea co

    ' Validity panel (scheme-aware)
    EvaluateDistributionForScheme totalArr, total, schemeKey, validityFlag, patternType, line1, line2, line3
    DrawValidityPanel wsDest, co, validityFlag, patternType, line1, line2, line3

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

    On Error GoTo PanelFail

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
    shp.Name = "FlagPanel_" & co.Name
    shp.Placement = xlMove

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

PanelFail:
    ' A panel-formatting limitation in a particular Excel version must not
    ' abort the enclosing subject block or disturb the next table position.
End Sub

Private Sub RealignSecAnalysisObjects(ByVal ws As Worksheet)
    Dim co As ChartObject
    Dim shp As Shape
    Dim titleRow As Long, endRow As Long, anchorCol As Long
    Dim chartName As String

    On Error GoTo AlignmentDone

    For Each co In ws.ChartObjects
        If ParseSecChartAnchor(co.Name, titleRow, endRow, anchorCol) Then
            co.Placement = xlMove
            co.Top = ExactSecRowTop(ws, titleRow)
            co.Left = ws.Columns(anchorCol).Left
            co.Width = ws.Columns(anchorCol).Resize(, 6).Width
            co.Height = ExactSecRowsHeight(ws, titleRow, endRow)
            SetSecChartPlotArea co
        End If
    Next co

    For Each shp In ws.Shapes
        If Left$(shp.Name, Len("FlagPanel_")) = "FlagPanel_" Then
            chartName = Mid$(shp.Name, Len("FlagPanel_") + 1)
            Set co = Nothing
            On Error Resume Next
            Set co = ws.ChartObjects(chartName)
            On Error GoTo AlignmentDone

            If Not co Is Nothing Then
                shp.Placement = xlMove
                shp.Left = co.Left + co.Width + 10
                shp.Width = co.Width * 1.65

                If ParseSecChartAnchor(chartName, titleRow, endRow, anchorCol) Then
                    ' Keep the narrative panel aligned to the subject block,
                    ' independently of the chart's measured plot correction.
                    shp.Top = ExactSecRowTop(ws, titleRow)
                    shp.Height = ExactSecRowsHeight(ws, titleRow, endRow)
                Else
                    shp.Top = co.Top
                    shp.Height = co.Height
                End If
            End If
        End If
    Next shp

AlignmentDone:
End Sub

Private Function ExactSecRowTop(ByVal ws As Worksheet, ByVal targetRow As Long) As Double
    Dim r As Long
    Dim totalPoints As Double

    If targetRow <= 1 Then Exit Function

    ' On the supplied Mac workbook, Rows(targetRow).Top accumulated roughly
    ' 0.25 extra points per default-height row. Summing the stored individual
    ' row heights follows the worksheet drawing grid exactly.
    For r = 1 To targetRow - 1
        totalPoints = totalPoints + CDbl(ws.Rows(r).RowHeight)
    Next r

    ExactSecRowTop = totalPoints
End Function

Private Function ExactSecRowsHeight(ByVal ws As Worksheet, _
                                    ByVal firstRow As Long, _
                                    ByVal lastRow As Long) As Double
    Dim r As Long
    Dim totalPoints As Double

    If firstRow < 1 Or lastRow < firstRow Then Exit Function

    For r = firstRow To lastRow
        totalPoints = totalPoints + CDbl(ws.Rows(r).RowHeight)
    Next r

    ExactSecRowsHeight = totalPoints
End Function

Private Sub SetSecChartPlotArea(ByVal co As ChartObject)
    Dim insideLeft As Double, insideTop As Double
    Dim insideWidth As Double, insideHeight As Double

    If co Is Nothing Then Exit Sub

    ' Excel otherwise chooses a different internal top margin as chart height
    ' changes. Keep the visible y-axis/plot near the table-header row, leaving
    ' enough room above bars for their data labels and below for categories.
    insideLeft = 32
    insideTop = 15
    insideWidth = co.Width - insideLeft - 6
    insideHeight = co.Height - insideTop - 22
    If insideWidth < 30 Or insideHeight < 20 Then Exit Sub

    On Error Resume Next
    co.Chart.Refresh
    With co.Chart.PlotArea
        .InsideLeft = insideLeft
        .InsideTop = insideTop
        .InsideWidth = insideWidth
        .InsideHeight = insideHeight
    End With
    On Error GoTo 0
End Sub

Private Function ParseSecChartAnchor(ByVal chartName As String, _
                                     ByRef titleRow As Long, _
                                     ByRef endRow As Long, _
                                     ByRef anchorCol As Long) As Boolean
    Const PREFIX As String = "SecChart_R"
    Dim body As String
    Dim endMarker As Long, colMarker As Long
    Dim titleText As String, endText As String, colText As String

    If Left$(chartName, Len(PREFIX)) <> PREFIX Then Exit Function

    body = Mid$(chartName, Len(PREFIX) + 1)
    endMarker = InStr(1, body, "_E", vbBinaryCompare)
    colMarker = InStr(1, body, "_C", vbBinaryCompare)
    If endMarker <= 1 Or colMarker <= endMarker + 2 Then Exit Function

    titleText = Left$(body, endMarker - 1)
    endText = Mid$(body, endMarker + 2, colMarker - endMarker - 2)
    colText = Mid$(body, colMarker + 2)
    If Not IsNumeric(titleText) Or Not IsNumeric(endText) Or Not IsNumeric(colText) Then Exit Function

    titleRow = CLng(titleText)
    endRow = CLng(endText)
    anchorCol = CLng(colText)
    If titleRow < 1 Or endRow < titleRow Or anchorCol < 1 Then Exit Function

    ParseSecChartAnchor = True
End Function

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

    If colMean > 0 Then ws.Cells(rowNum, colMean).Font.Color = RGB(0, 0, 0)
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
            topMaxIdx = 1

            pctPassLabel = "%A - D"
            pctFailLabel = "%E"
            pctTopLabel = "%A"
            meanLabel = ""

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

Private Function IsExcludedSecSubject(ByVal subjectName As String) As Boolean
    Dim ws As Worksheet
    Dim r As Long
    Dim subjectKey As String, excludedKey As String

    subjectKey = NormalizeSecSubjectKey(subjectName)
    If subjectKey = "" Then Exit Function

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("Settings")
    On Error GoTo 0
    If ws Is Nothing Then Exit Function

    ' One excluded subject per cell in Settings!V2:V100.
    For r = 2 To 100
        excludedKey = NormalizeSecSubjectKey(CStr(ws.Cells(r, "V").value))
        If excludedKey <> "" And excludedKey = subjectKey Then
            IsExcludedSecSubject = True
            Exit Function
        End If
    Next r
End Function

Private Function NormalizeSecSubjectKey(ByVal subjectName As String) As String
    Dim s As String
    Dim suffix As Variant

    s = Trim$(subjectName)
    If UCase$(Right$(s, 7)) = "(GRADE)" Then s = Trim$(Left$(s, Len(s) - 7))
    If UCase$(Right$(s, 7)) = "(SCORE)" Then s = Trim$(Left$(s, Len(s) - 7))
    s = UCase$(Replace(Trim$(s), " ", ""))

    For Each suffix In Array("-G1", "-G2", "-G3", "-N(T)", "-N(A)", _
                             "-NT", "-NA", "-EX", "-O")
        If Right$(s, Len(CStr(suffix))) = CStr(suffix) Then
            s = Trim$(Left$(s, Len(s) - Len(CStr(suffix))))
            Exit For
        End If
    Next suffix

    NormalizeSecSubjectKey = s
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
    Else
        InferSchemeFromHeader = GetLegacySchemeFromHeader(header)
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

        ' Avoid repeating level code in output sheet name, e.g. S1_Subj Analysis_S1_TERM1...
        If UCase$(Left$(baseLabel, Len(levelCode) + 1)) = UCase$(levelCode & "_") _
           Or UCase$(Left$(baseLabel, Len(levelCode) + 1)) = UCase$(levelCode & " ") _
           Or UCase$(Left$(baseLabel, Len(levelCode) + 1)) = UCase$(levelCode & "-") Then
            baseLabel = Mid$(baseLabel, Len(levelCode) + 2)
        End If

        ' Improve readability for compact source names like TERM1WA.
        baseLabel = Replace(baseLabel, "_", " ")
        baseLabel = Replace(baseLabel, "WA", " WA")
        Do While InStr(baseLabel, "  ") > 0
            baseLabel = Replace(baseLabel, "  ", " ")
        Loop
        baseLabel = Trim$(baseLabel)

        safeBase = CleanSheetNameFragment(baseLabel)
        If safeBase = "" Then safeBase = "Exam"

        maxShort = 31 - Len(prefix) - 1 - Len(yearPart)
        If maxShort < 1 Then maxShort = 1
        If Len(safeBase) > maxShort Then safeBase = Left$(safeBase, maxShort)

        safeName = prefix & safeBase & "_" & yearPart
        If Len(safeName) > 31 Then safeName = Left$(safeName, 31)
    Else
        baseLabel = examLabel
        If UCase$(Left$(baseLabel, Len(levelCode) + 1)) = UCase$(levelCode & "_") _
           Or UCase$(Left$(baseLabel, Len(levelCode) + 1)) = UCase$(levelCode & " ") _
           Or UCase$(Left$(baseLabel, Len(levelCode) + 1)) = UCase$(levelCode & "-") Then
            baseLabel = Mid$(baseLabel, Len(levelCode) + 2)
        End If
        baseLabel = Replace(baseLabel, "_", " ")
        baseLabel = Trim$(baseLabel)

        safeName = prefix & baseLabel
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

Private Function IsDistinctionGradeByScheme(ByVal gradeStr As String, ByVal schemeKey As String) As Boolean
    Dim g As String
    g = UCase$(Trim$(gradeStr))

    Select Case UCase$(Trim$(schemeKey))
        Case "G3"
            IsDistinctionGradeByScheme = (g = "A1" Or g = "A2")
        Case "G2"
            IsDistinctionGradeByScheme = (g = "1" Or g = "2")
        Case "G1"
            IsDistinctionGradeByScheme = (g = "A")
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

Private Function GetConfiguredLegacyGroup(ByVal className As String, _
                                          ByVal levelCode As String) As String
    Dim ws As Worksheet
    Dim r As Long
    Dim classKey As String, pattern As String, streamValue As String
    Dim bestMatchLength As Long, candidateGroup As String

    ' Explicit legacy streams apply only to the present Sec 4/5 cohorts.
    If UCase$(Trim$(levelCode)) <> "S4" And UCase$(Trim$(levelCode)) <> "S5" Then Exit Function

    classKey = UCase$(Replace(Trim$(className), " ", ""))
    If classKey = "" Then Exit Function

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("Settings")
    On Error GoTo 0
    If ws Is Nothing Then Exit Function

    ' Settings!D2:E50 supports either the existing Class -> Level mapping
    ' or a Sec 4/5 Class -> Stream mapping. Longest matching class pattern wins.
    For r = 2 To 50
        pattern = UCase$(Replace(Trim$(CStr(ws.Cells(r, "D").value)), " ", ""))
        streamValue = CStr(ws.Cells(r, "E").value)
        candidateGroup = LegacyStreamToGroup(streamValue)

        If pattern <> "" And candidateGroup <> "" Then
            If Left$(classKey, Len(pattern)) = pattern Then
                If Len(pattern) > bestMatchLength Then
                    bestMatchLength = Len(pattern)
                    GetConfiguredLegacyGroup = candidateGroup
                End If
            End If
        End If
    Next r
End Function

Private Function LegacyStreamToGroup(ByVal streamValue As String) As String
    Dim s As String
    s = UCase$(Replace(Trim$(streamValue), " ", ""))
    s = Replace(s, ".", "")
    s = Replace(s, "(", "")
    s = Replace(s, ")", "")

    Select Case s
        Case "EX", "EXPRESS"
            LegacyStreamToGroup = "G3"
        Case "NA", "NORMALACADEMIC"
            LegacyStreamToGroup = "G2"
        Case "NT", "NORMALTECHNICAL"
            LegacyStreamToGroup = "G1"
    End Select
End Function

Private Function ResolveFsbbGroup(ByVal g1Taken As Long, _
                                  ByVal g2Taken As Long, _
                                  ByVal g3Taken As Long, _
                                  ByVal attemptedCount As Long, _
                                  ByVal thresholdPct As Double) As String
    Dim maxTaken As Long
    Dim maxPct As Double
    Dim maxCount As Long

    If attemptedCount <= 0 Then
        ResolveFsbbGroup = ""
        Exit Function
    End If

    maxTaken = g3Taken
    If g2Taken > maxTaken Then maxTaken = g2Taken
    If g1Taken > maxTaken Then maxTaken = g1Taken
    maxPct = maxTaken * 100# / attemptedCount

    If maxPct < thresholdPct Then
        ResolveFsbbGroup = "MIXED"
        Exit Function
    End If

    If g3Taken = maxTaken Then maxCount = maxCount + 1
    If g2Taken = maxTaken Then maxCount = maxCount + 1
    If g1Taken = maxTaken Then maxCount = maxCount + 1
    If maxCount > 1 Then
        ResolveFsbbGroup = "MIXED"
    ElseIf g3Taken = maxTaken Then
        ResolveFsbbGroup = "G3"
    ElseIf g2Taken = maxTaken Then
        ResolveFsbbGroup = "G2"
    Else
        ResolveFsbbGroup = "G1"
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

