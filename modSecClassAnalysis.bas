Attribute VB_Name = "modSecClassAnalysis"
Option Explicit

'============================================================
' Module: modSecClassAnalysis
'
' PURPOSE
'   Build assessment-specific SEC Class Analysis and
'   Subject x Class pass-rate heatmaps as static-value sheets.
'
' DATA SOURCES
'   Class Analysis   <- matching AtRisk_* sheet
'   Subject x Class  <- matching S*_Subj Analysis_* sheet
'
' No worksheet formulas or hidden calculation blocks are used.
'============================================================

Private Const CLASS_REPORT_PREFIX As String = "ClassAn_"
Private Const SUBJCLASS_REPORT_PREFIX As String = "SubjClass_"
Private Const CLASS_NAV_PREFIX As String = "Nav_ClassAn_"
Private Const SUBJCLASS_NAV_PREFIX As String = "Nav_SubjClass_"
Private Const DASHBOARD_SHEET As String = "Dashboard"
Private Const CLASS_NAV_START_CELL As String = "Z3"
Private Const SUBJCLASS_NAV_START_CELL As String = "AC3"
Private Const SHAPE_ROUNDED_RECTANGLE As Long = 5

Private Const DEFAULT_MONITOR_PCT As Double = 10#
Private Const DEFAULT_ELEVATED_PCT As Double = 20#
Private Const DEFAULT_PRIORITY_PCT As Double = 30#

'------------------------------------------------------------
' PUBLIC ENTRY POINTS
'------------------------------------------------------------
Public Sub BuildSecClassAndSubjectClassReports()
    BuildSecClassAndSubjectClassAnalysis False
End Sub

Public Sub BuildSecClassAndSubjectClassAnalysis(ByVal suppressMessage As Boolean)
    Dim oldScreenUpdating As Boolean

    On Error GoTo ErrHandler
    oldScreenUpdating = Application.ScreenUpdating
    Application.ScreenUpdating = False

    EnsureClassRiskSettings
    BuildSecClassAnalysisReports
    BuildSecSubjectClassReports
    BuildSecClassAnalysisNavigation
    BuildSecSubjectClassNavigation

    Application.ScreenUpdating = oldScreenUpdating
    If Not suppressMessage Then
        MsgBox "SEC Class Analysis, Subject x Class analysis and menus have been refreshed.", _
               vbInformation, "SEC Class Reports Complete"
    End If
    Exit Sub

ErrHandler:
    Application.ScreenUpdating = oldScreenUpdating
    If suppressMessage Then
        Err.Raise Err.Number, "BuildSecClassAndSubjectClassAnalysis", Err.Description
    Else
        MsgBox "Could not build the SEC class reports: " & Err.Description, _
               vbCritical, "SEC Class Reports"
    End If
End Sub

'------------------------------------------------------------
' CLASS ANALYSIS FROM ATRISK
'------------------------------------------------------------
Private Sub BuildSecClassAnalysisReports()
    Dim sourceNames As Collection
    Dim sourceName As Variant
    Dim wsSrc As Worksheet, wsOut As Worksheet
    Dim suffix As String, levelCode As String, outName As String

    Set sourceNames = CollectSheetNamesWithPrefix("AtRisk_")
    For Each sourceName In sourceNames
        suffix = Mid$(CStr(sourceName), Len("AtRisk_") + 1)
        levelCode = FirstToken(suffix)
        If IsSecLevelCode(levelCode) Then
            outName = SafeWorksheetName(CLASS_REPORT_PREFIX & suffix)
            Set wsSrc = ThisWorkbook.Worksheets(CStr(sourceName))
            Set wsOut = GetOrCreateStaticReportSheet(outName)
            WriteClassAnalysis wsSrc, wsOut, levelCode, suffix
        End If
    Next sourceName
End Sub

Private Sub WriteClassAnalysis(ByVal wsSrc As Worksheet, _
                               ByVal wsOut As Worksheet, _
                               ByVal levelCode As String, _
                               ByVal reportSuffix As String)
    Dim classCol As Long, regCol As Long, nameCol As Long
    Dim attemptedCol As Long, passedCol As Long, failedCol As Long
    Dim vrmcCol As Long, abCol As Long
    Dim lastRow As Long, r As Long, outRow As Long
    Dim className As String, regText As String, nameText As String, studentKey As String
    Dim attemptedCount As Long, passedCount As Long, failedCount As Long
    Dim seen As Object, classes As Object
    Dim cohort As Object, passAll As Object, failOne As Object, failTwo As Object
    Dim failThree As Object, failFour As Object, failFivePlus As Object
    Dim failThreePlus As Object, abCases As Object, vrmcCases As Object
    Dim classNames() As String, classCount As Long, i As Long
    Dim totalCohort As Long, totalPassAll As Long, totalFailOne As Long, totalFailTwo As Long
    Dim totalFailThree As Long, totalFailFour As Long, totalFailFivePlus As Long
    Dim totalFailThreePlus As Long, totalAbCases As Long, totalVrmcCases As Long
    Dim monitorPct As Double, elevatedPct As Double, priorityPct As Double
    Dim riskRate As Double, riskShare As Double, readingText As String
    Dim titleText As String

    classCol = FindHeaderAtRow(wsSrc, 4, "Class")
    regCol = FindHeaderAtRow(wsSrc, 4, "RegNo")
    nameCol = FindHeaderAtRow(wsSrc, 4, "Name")
    attemptedCol = FindHeaderAtRow(wsSrc, 4, "Subjects Attempted")
    passedCol = FindHeaderAtRow(wsSrc, 4, "Subjects Passed")
    failedCol = FindHeaderAtRow(wsSrc, 4, "Subjects Failed")
    vrmcCol = FindHeaderAtRow(wsSrc, 4, "VR/MC Subjects")
    abCol = FindHeaderAtRow(wsSrc, 4, "AB Subjects")

    If classCol = 0 Or nameCol = 0 Or attemptedCol = 0 Or passedCol = 0 Or failedCol = 0 Then
        Err.Raise vbObjectError + 3201, "WriteClassAnalysis", _
                  "The sheet '" & wsSrc.Name & "' does not have the expected AtRisk columns."
    End If

    Set seen = NewTextDictionary()
    Set classes = NewTextDictionary()
    Set cohort = NewTextDictionary(): Set passAll = NewTextDictionary()
    Set failOne = NewTextDictionary(): Set failTwo = NewTextDictionary()
    Set failThree = NewTextDictionary(): Set failFour = NewTextDictionary()
    Set failFivePlus = NewTextDictionary(): Set failThreePlus = NewTextDictionary()
    Set abCases = NewTextDictionary(): Set vrmcCases = NewTextDictionary()

    lastRow = wsSrc.Cells(wsSrc.Rows.count, nameCol).End(xlUp).Row
    For r = 5 To lastRow
        className = Trim$(CStr(wsSrc.Cells(r, classCol).value))
        nameText = Trim$(CStr(wsSrc.Cells(r, nameCol).value))
        regText = ""
        If regCol > 0 Then regText = Trim$(CStr(wsSrc.Cells(r, regCol).value))

        If className <> "" And nameText <> "" Then
            If regText <> "" Then
                studentKey = className & "|REG|" & regText
            Else
                studentKey = className & "|NAME|" & nameText
            End If

            If Not seen.Exists(studentKey) Then
                attemptedCount = LongValue(wsSrc.Cells(r, attemptedCol).value)
                passedCount = LongValue(wsSrc.Cells(r, passedCol).value)
                failedCount = LongValue(wsSrc.Cells(r, failedCol).value)

                ' All-VR/MC 0/0/0 students remain in AtRisk for follow-up but
                ' are excluded from class outcome denominators.
                If attemptedCount > 0 Or passedCount > 0 Or failedCount > 0 Then
                    seen.Add studentKey, True
                    If Not classes.Exists(className) Then classes.Add className, True
                    IncrementDict cohort, className
                    Select Case failedCount
                        Case 0: IncrementDict passAll, className
                        Case 1: IncrementDict failOne, className
                        Case 2: IncrementDict failTwo, className
                        Case 3: IncrementDict failThree, className: IncrementDict failThreePlus, className
                        Case 4: IncrementDict failFour, className: IncrementDict failThreePlus, className
                        Case Else: IncrementDict failFivePlus, className: IncrementDict failThreePlus, className
                    End Select
                    If abCol > 0 Then
                        If Trim$(CStr(wsSrc.Cells(r, abCol).value)) <> "" Then IncrementDict abCases, className
                    End If
                    If vrmcCol > 0 Then
                        If Trim$(CStr(wsSrc.Cells(r, vrmcCol).value)) <> "" Then IncrementDict vrmcCases, className
                    End If
                End If
            End If
        End If
    Next r

    DictionaryKeysToSortedArray classes, classNames, classCount
    monitorPct = GetClassRiskSetting("ClassRiskMonitorAtLeast", DEFAULT_MONITOR_PCT)
    elevatedPct = GetClassRiskSetting("ClassRiskElevatedAtLeast", DEFAULT_ELEVATED_PCT)
    priorityPct = GetClassRiskSetting("ClassRiskPriorityAtLeast", DEFAULT_PRIORITY_PCT)
    NormalizeClassRiskThresholds monitorPct, elevatedPct, priorityPct

    titleText = Replace(reportSuffix, "_", " ") & " - Class Concentration Analysis"
    With wsOut
        .Range("A1:N1").Merge
        .Range("A1").value = titleText
        .Range("A1").Font.Bold = True
        .Range("A1").Font.Size = 16
        .Range("A1").Font.Color = RGB(31, 78, 121)
        .Range("A2:N2").Merge
        .Range("A2").value = "Class outcomes use each student's own subject combination. AB is counted as a failure; VR/MC is excluded. Students with 0 attempted, 0 passed and 0 failed are excluded from outcome rates."
        .Range("A2").Font.Italic = True
        .Range("A2").Font.Color = RGB(96, 120, 142)
        .Range("A2").WrapText = True
        .Rows(2).RowHeight = 32
    End With

    WriteStaticHeader wsOut, 4, Array("Class", "Cohort", "Pass All", "Fail 1", "Fail 2", _
                     "Fail 3", "Fail 4", "Fail 5+", "Fail 3+", "Fail 3+ Rate", _
                     "Share of Level Risk", "AB Cases", "VR/MC Cases", "KM Reading")

    outRow = 5
    For i = 1 To classCount
        className = classNames(i)
        riskRate = SafeRatio(DictLong(failThreePlus, className), DictLong(cohort, className))
        wsOut.Cells(outRow, 1).value = className
        wsOut.Cells(outRow, 2).value = DictLong(cohort, className)
        wsOut.Cells(outRow, 3).value = DictLong(passAll, className)
        wsOut.Cells(outRow, 4).value = DictLong(failOne, className)
        wsOut.Cells(outRow, 5).value = DictLong(failTwo, className)
        wsOut.Cells(outRow, 6).value = DictLong(failThree, className)
        wsOut.Cells(outRow, 7).value = DictLong(failFour, className)
        wsOut.Cells(outRow, 8).value = DictLong(failFivePlus, className)
        wsOut.Cells(outRow, 9).value = DictLong(failThreePlus, className)
        wsOut.Cells(outRow, 10).value = riskRate
        wsOut.Cells(outRow, 11).value = 0#
        wsOut.Cells(outRow, 12).value = DictLong(abCases, className)
        wsOut.Cells(outRow, 13).value = DictLong(vrmcCases, className)
        readingText = ClassRiskReading(riskRate * 100#, monitorPct, elevatedPct, priorityPct)
        wsOut.Cells(outRow, 14).value = readingText
        StyleClassRiskReading wsOut.Cells(outRow, 14), readingText

        totalCohort = totalCohort + DictLong(cohort, className)
        totalPassAll = totalPassAll + DictLong(passAll, className)
        totalFailOne = totalFailOne + DictLong(failOne, className)
        totalFailTwo = totalFailTwo + DictLong(failTwo, className)
        totalFailThree = totalFailThree + DictLong(failThree, className)
        totalFailFour = totalFailFour + DictLong(failFour, className)
        totalFailFivePlus = totalFailFivePlus + DictLong(failFivePlus, className)
        totalFailThreePlus = totalFailThreePlus + DictLong(failThreePlus, className)
        totalAbCases = totalAbCases + DictLong(abCases, className)
        totalVrmcCases = totalVrmcCases + DictLong(vrmcCases, className)
        outRow = outRow + 1
    Next i

    If totalFailThreePlus > 0 Then
        For r = 5 To outRow - 1
            wsOut.Cells(r, 11).value = wsOut.Cells(r, 9).value / totalFailThreePlus
        Next r
    End If

    wsOut.Cells(outRow, 1).value = "COHORT"
    wsOut.Cells(outRow, 2).value = totalCohort
    wsOut.Cells(outRow, 3).value = totalPassAll
    wsOut.Cells(outRow, 4).value = totalFailOne
    wsOut.Cells(outRow, 5).value = totalFailTwo
    wsOut.Cells(outRow, 6).value = totalFailThree
    wsOut.Cells(outRow, 7).value = totalFailFour
    wsOut.Cells(outRow, 8).value = totalFailFivePlus
    wsOut.Cells(outRow, 9).value = totalFailThreePlus
    wsOut.Cells(outRow, 10).value = SafeRatio(totalFailThreePlus, totalCohort)
    wsOut.Cells(outRow, 11).value = IIf(totalFailThreePlus > 0, 1#, 0#)
    wsOut.Cells(outRow, 12).value = totalAbCases
    wsOut.Cells(outRow, 13).value = totalVrmcCases
    wsOut.Cells(outRow, 14).value = levelCode & " overall"
    StyleStaticCohortRow wsOut.Range(wsOut.Cells(outRow, 1), wsOut.Cells(outRow, 14))

    FinalizeClassAnalysis wsOut, outRow
    AddStaticReportHomeButton wsOut, RGB(217, 234, 211), RGB(106, 168, 79), RGB(39, 78, 19)
End Sub

Private Sub FinalizeClassAnalysis(ByVal ws As Worksheet, ByVal lastRow As Long)
    With ws.Range("A4:N" & lastRow)
        .Borders.LineStyle = xlContinuous
        .Borders.Color = RGB(210, 220, 230)
        .Borders.Weight = xlThin
        .VerticalAlignment = xlCenter
    End With
    ws.Range("B5:I" & lastRow).HorizontalAlignment = xlCenter
    ws.Range("J5:K" & lastRow).NumberFormat = "0.0%"
    ws.Range("J5:N" & lastRow).HorizontalAlignment = xlCenter
    ws.Columns("A").ColumnWidth = 20
    ws.Columns("B:I").ColumnWidth = 10
    ws.Columns("J:K").ColumnWidth = 15
    ws.Columns("L:M").ColumnWidth = 11
    ws.Columns("N").ColumnWidth = 18
    ws.Rows(4).RowHeight = 30
    ws.Range("A4:N4").WrapText = True
    FreezeStaticReport ws, "A5"
    LimitStaticReportArea ws, lastRow + 3, 14
End Sub

'------------------------------------------------------------
' SUBJECT x CLASS FROM SUBJECT ANALYSIS
'------------------------------------------------------------
Private Sub BuildSecSubjectClassReports()
    Dim sourceNames As Collection
    Dim sourceName As Variant
    Dim wsSrc As Worksheet, wsOut As Worksheet
    Dim markerPos As Long, levelCode As String, suffix As String, outName As String

    Set sourceNames = CollectSubjectAnalysisSheetNames()
    For Each sourceName In sourceNames
        markerPos = InStr(1, CStr(sourceName), "_Subj Analysis_", vbTextCompare)
        If markerPos > 0 Then
            levelCode = Left$(CStr(sourceName), 2)
            suffix = Mid$(CStr(sourceName), markerPos + Len("_Subj Analysis_"))
            If IsSecLevelCode(levelCode) And suffix <> "" Then
                outName = SafeWorksheetName(SUBJCLASS_REPORT_PREFIX & levelCode & "_" & suffix)
                Set wsSrc = ThisWorkbook.Worksheets(CStr(sourceName))
                Set wsOut = GetOrCreateStaticReportSheet(outName)
                WriteSubjectClassAnalysis wsSrc, wsOut, levelCode, suffix
            End If
        End If
    Next sourceName
End Sub

Private Sub WriteSubjectClassAnalysis(ByVal wsSrc As Worksheet, _
                                      ByVal wsOut As Worksheet, _
                                      ByVal levelCode As String, _
                                      ByVal reportSuffix As String)
    Dim subjectRates As Object, subjectCohort As Object, classes As Object
    Dim rates As Object
    Dim lastRow As Long, titleRow As Long, detailRow As Long
    Dim titleText As String, subjectName As String, className As String
    Dim noCol As Long, passCol As Long, nValue As Long
    Dim passValue As Variant
    Dim subjectNames() As String, subjectCount As Long
    Dim classNames() As String, classCount As Long
    Dim i As Long, j As Long, outRow As Long, lastCol As Long

    Set subjectRates = NewTextDictionary()
    Set subjectCohort = NewTextDictionary()
    Set classes = NewTextDictionary()

    lastRow = LastPopulatedRow(wsSrc)
    For titleRow = 1 To lastRow
        titleText = Trim$(CStr(wsSrc.Cells(titleRow, 1).value))
        If IsSubjectTableTitle(titleText) Then
            subjectName = Trim$(Left$(titleText, Len(titleText) - 4))
            noCol = FindHeaderAtRow(wsSrc, titleRow + 1, "No.")
            If noCol > 0 Then
                passCol = noCol + 1
                Set rates = NewTextDictionary()
                detailRow = titleRow + 2
                Do While detailRow <= lastRow
                    className = Trim$(CStr(wsSrc.Cells(detailRow, 1).value))
                    If className = "" Then Exit Do
                    nValue = LongValue(wsSrc.Cells(detailRow, noCol).value)
                    passValue = wsSrc.Cells(detailRow, passCol).value
                    If UCase$(className) = "COHORT" Then
                        If IsNumeric(passValue) And nValue > 0 Then
                            subjectCohort(subjectName) = CDbl(passValue) / 100#
                        End If
                        Exit Do
                    ElseIf nValue > 0 And IsNumeric(passValue) Then
                        rates(className) = CDbl(passValue) / 100#
                        If Not classes.Exists(className) Then classes.Add className, True
                    End If
                    detailRow = detailRow + 1
                Loop
                If rates.count > 0 And subjectCohort.Exists(subjectName) Then
                    If subjectRates.Exists(subjectName) Then subjectRates.Remove subjectName
                    subjectRates.Add subjectName, rates
                End If
            End If
        End If
    Next titleRow

    DictionaryKeysToSortedArray classes, classNames, classCount
    DictionaryKeysToSortedArray subjectRates, subjectNames, subjectCount
    SortSubjectsByCohortRate subjectNames, subjectCount, subjectCohort
    lastCol = classCount + 2
    If lastCol < 3 Then lastCol = 3

    With wsOut
        .Range(.Cells(1, 1), .Cells(1, lastCol)).Merge
        .Cells(1, 1).value = levelCode & " " & Replace(reportSuffix, "_", " ") & _
                              " - Subject x Class Pass-Rate Heatmap"
        .Cells(1, 1).Font.Bold = True
        .Cells(1, 1).Font.Size = 16
        .Cells(1, 1).Font.Color = RGB(31, 78, 121)
        .Range(.Cells(2, 1), .Cells(2, lastCol)).Merge
        .Cells(2, 1).value = "Each percentage is the share of students in that class who took the stated subject/G-level and achieved a passing grade. The Cohort column combines all " & levelCode & " classes."
        .Cells(2, 1).Font.Italic = True
        .Cells(2, 1).Font.Color = RGB(96, 120, 142)
        .Range(.Cells(3, 1), .Cells(3, lastCol)).Merge
        .Cells(3, 1).value = "Colours: red below 50%; orange 50-69%; yellow 70-79%; green 80% or higher. '-' means no students in that class took the subject/G-level. Small candidature should be interpreted cautiously."
        .Cells(3, 1).Font.Italic = True
        .Cells(3, 1).Font.Color = RGB(96, 120, 142)
    End With

    wsOut.Cells(5, 1).value = "Subject / G-Level"
    For i = 1 To classCount
        wsOut.Cells(5, i + 1).value = classNames(i)
    Next i
    wsOut.Cells(5, lastCol).value = "Cohort"
    StyleStaticHeader wsOut.Range(wsOut.Cells(5, 1), wsOut.Cells(5, lastCol))

    outRow = 6
    For i = 1 To subjectCount
        subjectName = subjectNames(i)
        Set rates = subjectRates(subjectName)
        wsOut.Cells(outRow, 1).value = subjectName
        For j = 1 To classCount
            className = classNames(j)
            If rates.Exists(className) Then
                wsOut.Cells(outRow, j + 1).value = CDbl(rates(className))
                StyleHeatmapCell wsOut.Cells(outRow, j + 1), CDbl(rates(className))
            Else
                wsOut.Cells(outRow, j + 1).value = "-"
                wsOut.Cells(outRow, j + 1).Font.Color = RGB(140, 140, 140)
                wsOut.Cells(outRow, j + 1).Interior.Color = RGB(245, 245, 245)
            End If
        Next j
        wsOut.Cells(outRow, lastCol).value = CDbl(subjectCohort(subjectName))
        StyleHeatmapCell wsOut.Cells(outRow, lastCol), CDbl(subjectCohort(subjectName))
        wsOut.Cells(outRow, lastCol).Font.Bold = True
        outRow = outRow + 1
    Next i

    FinalizeSubjectClassAnalysis wsOut, outRow - 1, lastCol
    AddStaticReportHomeButton wsOut, RGB(207, 226, 243), RGB(61, 133, 198), RGB(11, 83, 148)
End Sub

Private Sub FinalizeSubjectClassAnalysis(ByVal ws As Worksheet, _
                                         ByVal lastRow As Long, _
                                         ByVal lastCol As Long)
    If lastRow >= 6 Then
        With ws.Range(ws.Cells(5, 1), ws.Cells(lastRow, lastCol))
            .Borders.LineStyle = xlContinuous
            .Borders.Color = RGB(210, 220, 230)
            .Borders.Weight = xlThin
            .VerticalAlignment = xlCenter
        End With
        ws.Range(ws.Cells(6, 2), ws.Cells(lastRow, lastCol)).NumberFormat = "0.0%"
        ws.Range(ws.Cells(6, 2), ws.Cells(lastRow, lastCol)).HorizontalAlignment = xlCenter
    End If
    ws.Columns(1).ColumnWidth = 29
    If lastCol >= 2 Then
        ws.Range(ws.Cells(1, 2), ws.Cells(1, lastCol)).EntireColumn.ColumnWidth = 14
    End If
    ws.Rows(5).RowHeight = 44
    ws.Range(ws.Cells(5, 1), ws.Cells(5, lastCol)).WrapText = True
    FreezeStaticReport ws, "B6"
    LimitStaticReportArea ws, lastRow + 3, lastCol
End Sub

'------------------------------------------------------------
' DASHBOARD NAVIGATION
'------------------------------------------------------------
Public Sub BuildSecClassAnalysisNavigation()
    BuildStaticReportNavigation CLASS_REPORT_PREFIX, CLASS_NAV_PREFIX, CLASS_NAV_START_CELL, _
                                "Class Analysis Menu", " Class Analysis", _
                                RGB(234, 209, 220), RGB(166, 77, 121), RGB(116, 27, 71)
End Sub

Public Sub BuildSecSubjectClassNavigation()
    BuildStaticReportNavigation SUBJCLASS_REPORT_PREFIX, SUBJCLASS_NAV_PREFIX, SUBJCLASS_NAV_START_CELL, _
                                "Subject x Class Menu", " Subject x Class", _
                                RGB(208, 224, 227), RGB(69, 129, 142), RGB(19, 79, 92)
End Sub

Private Sub BuildStaticReportNavigation(ByVal reportPrefix As String, _
                                        ByVal shapePrefix As String, _
                                        ByVal startCellAddress As String, _
                                        ByVal headingText As String, _
                                        ByVal labelSuffix As String, _
                                        ByVal fillColor As Long, _
                                        ByVal lineColor As Long, _
                                        ByVal fontColor As Long)
    Dim ws As Worksheet, startCell As Range, shp As Shape
    Dim names As Collection, reportName As Variant
    Dim rowPtr As Long, k As Long

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(DASHBOARD_SHEET)
    On Error GoTo 0
    If ws Is Nothing Then Exit Sub

    Set startCell = ws.Range(startCellAddress)
    ws.Range(ws.Cells(startCell.Row, startCell.Column), _
             ws.Cells(startCell.Row + 220, startCell.Column + 2)).Clear
    For k = ws.Shapes.count To 1 Step -1
        Set shp = ws.Shapes(k)
        If Left$(shp.Name, Len(shapePrefix)) = shapePrefix Then shp.Delete
    Next k

    ws.Cells(startCell.Row, startCell.Column).value = headingText
    ws.Cells(startCell.Row, startCell.Column).Font.Bold = True
    ws.Cells(startCell.Row, startCell.Column).Font.Size = 12
    ws.Cells(startCell.Row, startCell.Column).Font.Color = fontColor
    rowPtr = startCell.Row + 1

    Set names = CollectSheetNamesWithPrefix(reportPrefix)
    If names.count = 0 Then
        ws.Cells(rowPtr, startCell.Column).value = "No reports built."
        ws.Cells(rowPtr, startCell.Column).Font.Italic = True
    Else
        For Each reportName In names
            CreateStaticNavButton ws, CStr(reportName), shapePrefix, _
                                  Replace(Mid$(CStr(reportName), Len(reportPrefix) + 1), "_", " ") & labelSuffix, _
                                  rowPtr, startCell.Column, fillColor, lineColor, fontColor
            rowPtr = rowPtr + 2
        Next reportName
    End If
End Sub

Private Sub CreateStaticNavButton(ByVal ws As Worksheet, _
                                  ByVal targetSheet As String, _
                                  ByVal shapePrefix As String, _
                                  ByVal displayText As String, _
                                  ByVal rowNum As Long, _
                                  ByVal firstCol As Long, _
                                  ByVal fillColor As Long, _
                                  ByVal lineColor As Long, _
                                  ByVal fontColor As Long)
    Dim shp As Shape
    Dim btnWidth As Double, btnHeight As Double

    btnWidth = ws.Columns(firstCol).Resize(, 3).Width * 0.95
    btnHeight = ws.Rows(rowNum).Height * 1.3
    Set shp = ws.Shapes.AddShape(SHAPE_ROUNDED_RECTANGLE, _
                                 ws.Cells(rowNum, firstCol).Left, _
                                 ws.Cells(rowNum, firstCol).Top, _
                                 btnWidth, btnHeight)
    With shp
        .Name = shapePrefix & targetSheet
        .Fill.ForeColor.RGB = fillColor
        .line.ForeColor.RGB = lineColor
        .line.Weight = 1.5
        With .TextFrame2
            .TextRange.text = displayText
            .TextRange.Font.Name = "Calibri"
            .TextRange.Font.Size = 9.5
            .TextRange.Font.Fill.ForeColor.RGB = fontColor
            .TextRange.ParagraphFormat.Alignment = msoAlignCenter
            .VerticalAnchor = msoAnchorMiddle
            .MarginLeft = 4
            .MarginRight = 4
        End With
    End With
    ws.Hyperlinks.Add Anchor:=shp, Address:="", SubAddress:="'" & targetSheet & "'!A1"
End Sub

'------------------------------------------------------------
' FORMATTING HELPERS
'------------------------------------------------------------
Private Sub WriteStaticHeader(ByVal ws As Worksheet, ByVal rowNum As Long, ByVal headers As Variant)
    Dim i As Long, lastCol As Long
    lastCol = UBound(headers) - LBound(headers) + 1
    For i = LBound(headers) To UBound(headers)
        ws.Cells(rowNum, i - LBound(headers) + 1).value = headers(i)
    Next i
    StyleStaticHeader ws.Range(ws.Cells(rowNum, 1), ws.Cells(rowNum, lastCol))
End Sub

Private Sub StyleStaticHeader(ByVal rng As Range)
    With rng
        .Interior.Color = RGB(68, 114, 196)
        .Font.Color = RGB(255, 255, 255)
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Borders.LineStyle = xlContinuous
        .Borders.Color = RGB(183, 204, 221)
    End With
End Sub

Private Sub StyleStaticCohortRow(ByVal rng As Range)
    With rng
        .Interior.Color = RGB(255, 242, 204)
        .Font.Bold = True
        .Borders.LineStyle = xlContinuous
        .Borders.Color = RGB(191, 143, 0)
    End With
End Sub

Private Sub StyleClassRiskReading(ByVal targetCell As Range, ByVal readingText As String)
    targetCell.Font.Bold = True
    targetCell.HorizontalAlignment = xlCenter
    Select Case readingText
        Case "Priority class"
            targetCell.Interior.Color = RGB(244, 204, 204): targetCell.Font.Color = RGB(156, 0, 6)
        Case "Elevated"
            targetCell.Interior.Color = RGB(252, 229, 205): targetCell.Font.Color = RGB(180, 95, 6)
        Case "Monitor"
            targetCell.Interior.Color = RGB(255, 242, 204): targetCell.Font.Color = RGB(127, 96, 0)
        Case Else
            targetCell.Interior.Color = RGB(217, 234, 211): targetCell.Font.Color = RGB(39, 78, 19)
    End Select
End Sub

Private Sub StyleHeatmapCell(ByVal targetCell As Range, ByVal passRate As Double)
    targetCell.HorizontalAlignment = xlCenter
    Select Case passRate
        Case Is < 0.5
            targetCell.Interior.Color = RGB(244, 204, 204): targetCell.Font.Color = RGB(156, 0, 6)
        Case Is < 0.7
            targetCell.Interior.Color = RGB(252, 229, 205): targetCell.Font.Color = RGB(180, 95, 6)
        Case Is < 0.8
            targetCell.Interior.Color = RGB(255, 242, 204): targetCell.Font.Color = RGB(127, 96, 0)
        Case Else
            targetCell.Interior.Color = RGB(217, 234, 211): targetCell.Font.Color = RGB(39, 78, 19)
    End Select
End Sub

Private Sub AddStaticReportHomeButton(ByVal ws As Worksheet, _
                                      ByVal fillColor As Long, _
                                      ByVal lineColor As Long, _
                                      ByVal fontColor As Long)
    Dim shp As Shape, tgt As Range
    On Error Resume Next
    ws.Shapes("HomeBtn").Delete
    On Error GoTo 0

    Set tgt = ws.Range("E1")
    Set shp = ws.Shapes.AddShape(SHAPE_ROUNDED_RECTANGLE, tgt.Left, tgt.Top, tgt.Width * 1.5, tgt.Height * 1.25)
    With shp
        .Name = "HomeBtn"
        .Fill.ForeColor.RGB = fillColor
        .line.ForeColor.RGB = lineColor
        .line.Weight = 1.5
        With .TextFrame2
            .TextRange.text = "Home"
            .TextRange.Font.Name = "Calibri"
            .TextRange.Font.Size = 10.5
            .TextRange.Font.Fill.ForeColor.RGB = fontColor
            .TextRange.ParagraphFormat.Alignment = msoAlignCenter
            .VerticalAnchor = msoAnchorMiddle
        End With
    End With
    ws.Hyperlinks.Add Anchor:=shp, Address:="", SubAddress:="'Dashboard'!A1"
End Sub

Private Sub FreezeStaticReport(ByVal ws As Worksheet, ByVal selectAddress As String)
    ws.Activate
    ActiveWindow.FreezePanes = False
    ws.Range(selectAddress).Select
    ActiveWindow.FreezePanes = True
    ActiveWindow.DisplayGridlines = False
End Sub

Private Sub LimitStaticReportArea(ByVal ws As Worksheet, ByVal lastRow As Long, ByVal lastCol As Long)
    ws.ScrollArea = ws.Range(ws.Cells(1, 1), ws.Cells(lastRow, lastCol)).Address
    ws.PageSetup.PrintArea = ws.Range(ws.Cells(1, 1), ws.Cells(lastRow, lastCol)).Address
    ws.PageSetup.Orientation = xlLandscape
    ws.PageSetup.Zoom = False
    ws.PageSetup.FitToPagesWide = 1
    ws.PageSetup.FitToPagesTall = False
End Sub

'------------------------------------------------------------
' SETTINGS
'------------------------------------------------------------
Private Sub EnsureClassRiskSettings()
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("Settings")
    On Error GoTo 0
    If ws Is Nothing Then Exit Sub

    If Trim$(CStr(ws.Range("S8").value)) = "" Then ws.Range("S8").value = "ClassRiskMonitorAtLeast=10%"
    If Trim$(CStr(ws.Range("S9").value)) = "" Then ws.Range("S9").value = "ClassRiskElevatedAtLeast=20%"
    If Trim$(CStr(ws.Range("S10").value)) = "" Then ws.Range("S10").value = "ClassRiskPriorityAtLeast=30%"
    ws.Columns("S").AutoFit
End Sub

Private Function GetClassRiskSetting(ByVal settingKey As String, ByVal defaultValue As Double) As Double
    Dim ws As Worksheet, r As Long
    Dim rawText As String, compactKey As String, p As Long, v As Double
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("Settings")
    On Error GoTo 0
    If ws Is Nothing Then GetClassRiskSetting = defaultValue: Exit Function

    compactKey = UCase$(Replace(settingKey, " ", ""))
    For r = 2 To 30
        rawText = Trim$(CStr(ws.Cells(r, "S").value))
        p = InStr(1, rawText, "=", vbBinaryCompare)
        If p > 0 Then
            If UCase$(Replace(Trim$(Left$(rawText, p - 1)), " ", "")) = compactKey Then
                rawText = Replace(Trim$(Mid$(rawText, p + 1)), "%", "")
                If IsNumeric(rawText) Then
                    v = CDbl(rawText)
                    If v <= 1# Then v = v * 100#
                    If v >= 0# And v <= 100# Then GetClassRiskSetting = v: Exit Function
                End If
            End If
        End If
    Next r
    GetClassRiskSetting = defaultValue
End Function

Private Sub NormalizeClassRiskThresholds(ByRef monitorPct As Double, _
                                         ByRef elevatedPct As Double, _
                                         ByRef priorityPct As Double)
    If monitorPct < 0# Or elevatedPct <= monitorPct Or priorityPct <= elevatedPct Or priorityPct > 100# Then
        monitorPct = DEFAULT_MONITOR_PCT
        elevatedPct = DEFAULT_ELEVATED_PCT
        priorityPct = DEFAULT_PRIORITY_PCT
    End If
End Sub

Private Function ClassRiskReading(ByVal riskPct As Double, _
                                  ByVal monitorPct As Double, _
                                  ByVal elevatedPct As Double, _
                                  ByVal priorityPct As Double) As String
    If riskPct >= priorityPct Then
        ClassRiskReading = "Priority class"
    ElseIf riskPct >= elevatedPct Then
        ClassRiskReading = "Elevated"
    ElseIf riskPct >= monitorPct Then
        ClassRiskReading = "Monitor"
    Else
        ClassRiskReading = "Lower incidence"
    End If
End Function

'------------------------------------------------------------
' GENERAL HELPERS
'------------------------------------------------------------
Private Function CollectSheetNamesWithPrefix(ByVal prefixText As String) As Collection
    Dim result As New Collection
    Dim names() As String, count As Long, i As Long, j As Long, tmp As String
    Dim ws As Worksheet

    For Each ws In ThisWorkbook.Worksheets
        If Left$(ws.Name, Len(prefixText)) = prefixText Then
            count = count + 1
            ReDim Preserve names(1 To count)
            names(count) = ws.Name
        End If
    Next ws
    For i = 1 To count - 1
        For j = i + 1 To count
            If StrComp(names(j), names(i), vbTextCompare) < 0 Then
                tmp = names(i): names(i) = names(j): names(j) = tmp
            End If
        Next j
    Next i
    For i = 1 To count
        result.Add names(i)
    Next i
    Set CollectSheetNamesWithPrefix = result
End Function

Private Function CollectSubjectAnalysisSheetNames() As Collection
    Dim result As New Collection
    Dim names() As String, count As Long, i As Long, j As Long, tmp As String
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        If IsSecLevelCode(Left$(ws.Name, 2)) And _
           InStr(1, ws.Name, "_Subj Analysis_", vbTextCompare) > 0 Then
            count = count + 1
            ReDim Preserve names(1 To count)
            names(count) = ws.Name
        End If
    Next ws
    For i = 1 To count - 1
        For j = i + 1 To count
            If StrComp(names(j), names(i), vbTextCompare) < 0 Then
                tmp = names(i): names(i) = names(j): names(j) = tmp
            End If
        Next j
    Next i
    For i = 1 To count
        result.Add names(i)
    Next i
    Set CollectSubjectAnalysisSheetNames = result
End Function

Private Function GetOrCreateStaticReportSheet(ByVal sheetName As String) As Worksheet
    Dim ws As Worksheet, k As Long
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.count))
        ws.Name = sheetName
    Else
        ws.ScrollArea = ""
        ws.Cells.UnMerge
        ws.Cells.Clear
        For k = ws.Shapes.count To 1 Step -1
            ws.Shapes(k).Delete
        Next k
    End If
    Set GetOrCreateStaticReportSheet = ws
End Function

Private Function FindHeaderAtRow(ByVal ws As Worksheet, ByVal headerRow As Long, ByVal headerText As String) As Long
    Dim lastCol As Long, c As Long
    lastCol = LastPopulatedColumn(ws, headerRow)
    For c = 1 To lastCol
        If StrComp(Trim$(CStr(ws.Cells(headerRow, c).value)), headerText, vbTextCompare) = 0 Then
            FindHeaderAtRow = c
            Exit Function
        End If
    Next c
End Function

Private Function LastPopulatedColumn(ByVal ws As Worksheet, ByVal rowNum As Long) As Long
    Dim hit As Range
    On Error Resume Next
    Set hit = ws.Rows(rowNum).Find(What:="*", After:=ws.Cells(rowNum, 1), _
                                   LookIn:=xlValues, LookAt:=xlPart, _
                                   SearchOrder:=xlByColumns, SearchDirection:=xlPrevious, _
                                   MatchCase:=False)
    On Error GoTo 0
    If Not hit Is Nothing Then LastPopulatedColumn = hit.Column
End Function

Private Function LastPopulatedRow(ByVal ws As Worksheet) As Long
    Dim hit As Range
    On Error Resume Next
    Set hit = ws.Cells.Find(What:="*", After:=ws.Cells(1, 1), _
                            LookIn:=xlValues, LookAt:=xlPart, _
                            SearchOrder:=xlByRows, SearchDirection:=xlPrevious, _
                            MatchCase:=False)
    On Error GoTo 0
    If Not hit Is Nothing Then LastPopulatedRow = hit.Row
End Function

Private Function IsSubjectTableTitle(ByVal valueText As String) As Boolean
    Dim suffixText As String
    If Len(valueText) < 4 Then Exit Function
    suffixText = UCase$(Right$(Trim$(valueText), 4))
    IsSubjectTableTitle = (suffixText = "[G1]" Or suffixText = "[G2]" Or suffixText = "[G3]")
End Function

Private Function IsSecLevelCode(ByVal levelCode As String) As Boolean
    Dim d As String
    If Len(levelCode) <> 2 Then Exit Function
    If UCase$(Left$(levelCode, 1)) <> "S" Then Exit Function
    d = Right$(levelCode, 1)
    IsSecLevelCode = (d >= "1" And d <= "5")
End Function

Private Function FirstToken(ByVal valueText As String) As String
    Dim parts As Variant
    parts = Split(valueText, "_")
    FirstToken = CStr(parts(0))
End Function

Private Function SafeWorksheetName(ByVal candidateName As String) As String
    Dim resultText As String
    resultText = Replace(candidateName, ":", "-")
    resultText = Replace(resultText, "\", "-")
    resultText = Replace(resultText, "/", "-")
    resultText = Replace(resultText, "?", "")
    resultText = Replace(resultText, "*", "")
    resultText = Replace(resultText, "[", "(")
    resultText = Replace(resultText, "]", ")")
    If Len(resultText) > 31 Then resultText = Left$(resultText, 31)
    SafeWorksheetName = resultText
End Function

Private Function NewTextDictionary() As Object
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    dict.CompareMode = vbTextCompare
    Set NewTextDictionary = dict
End Function

Private Sub IncrementDict(ByVal dict As Object, ByVal keyText As String)
    If dict.Exists(keyText) Then
        dict(keyText) = CLng(dict(keyText)) + 1
    Else
        dict.Add keyText, 1&
    End If
End Sub

Private Function DictLong(ByVal dict As Object, ByVal keyText As String) As Long
    If dict.Exists(keyText) Then DictLong = CLng(dict(keyText))
End Function

Private Function LongValue(ByVal valueIn As Variant) As Long
    If IsNumeric(valueIn) Then LongValue = CLng(valueIn)
End Function

Private Function SafeRatio(ByVal numerator As Long, ByVal denominator As Long) As Double
    If denominator > 0 Then SafeRatio = numerator / denominator
End Function

Private Sub DictionaryKeysToSortedArray(ByVal dict As Object, _
                                        ByRef values() As String, _
                                        ByRef valueCount As Long)
    Dim keyValue As Variant, i As Long, j As Long, tmp As String
    valueCount = dict.count
    If valueCount = 0 Then Exit Sub
    ReDim values(1 To valueCount)
    i = 0
    For Each keyValue In dict.Keys
        i = i + 1
        values(i) = CStr(keyValue)
    Next keyValue
    For i = 1 To valueCount - 1
        For j = i + 1 To valueCount
            If StrComp(values(j), values(i), vbTextCompare) < 0 Then
                tmp = values(i): values(i) = values(j): values(j) = tmp
            End If
        Next j
    Next i
End Sub

Private Sub SortSubjectsByCohortRate(ByRef subjectNames() As String, _
                                     ByVal subjectCount As Long, _
                                     ByVal cohortRates As Object)
    Dim i As Long, j As Long, tmp As String
    Dim aRate As Double, bRate As Double
    For i = 1 To subjectCount - 1
        For j = i + 1 To subjectCount
            aRate = CDbl(cohortRates(subjectNames(i)))
            bRate = CDbl(cohortRates(subjectNames(j)))
            If bRate < aRate Or _
               (bRate = aRate And StrComp(subjectNames(j), subjectNames(i), vbTextCompare) < 0) Then
                tmp = subjectNames(i): subjectNames(i) = subjectNames(j): subjectNames(j) = tmp
            End If
        Next j
    Next i
End Sub
