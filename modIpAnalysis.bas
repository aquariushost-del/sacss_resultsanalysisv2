Attribute VB_Name = "modIpAnalysis"

Option Explicit

'===========================================================
' Module: modIpAnalysis
'
' PURPOSE:
'   100% automatic IP subject analysis (A+ÐU GPA system).
'
' FEATURES:
'   - Detects IP classes (1FÐ1I, 2FÐ2I, 3FÐ3J, 4FÐ4J)
'   - Identifies IP grade columns (A+, A, B+, B, C+, C, D+, D, U)
'   - Builds distribution tables:
'       Class | A+ | A | B+ | B | C+ | C | D+ | D | U | No. | %Pass | %Fail | %Top | GPA
'   - Adds vertical pastel bar chart for COHORT
'   - Adds IP GPA validity panel (with interpretation text)
'   - Fully parallel to SEC architecture
'
' OUTPUT:
'   - Y1_Subj Analysis_<ExamName>
'   - Y2_Subj Analysis_<ExamName>
'   - Y3_Subj Analysis_<ExamName>
'   - Y4_Subj Analysis_<ExamName>
'
'===========================================================

Private Const SHAPE_ROUNDED_RECTANGLE As Long = 5

Private Const LOW_N_THRESHOLD As Long = 10       ' Cohort < 10 ? LOW N
Private Const TOP_HEAVY_PCT As Double = 30#      ' A+ + A = 30% = Top-heavy
Private Const WEAK_TAIL_PCT As Double = 25#      ' D+ + D + U = 25% = Weak tail
Private Const FAT_MIDDLE_PCT As Double = 50#     ' B+..C = 50%
Private Const THIN_MIDDLE_PCT As Double = 20#
Private Const SKEWED_HIGH_GPA As Double = 3.2
Private Const SKEWED_LOW_GPA As Double = 2.3

'===========================================================
' ENTRY POINT Ð RUN THIS TO GENERATE ALL IP ANALYSIS
'===========================================================
Public Sub BuildAllIp_SubjectAnalysis()
    Dim ws As Worksheet
    Dim wb As Workbook
    
    Set wb = ThisWorkbook
    
    For Each ws In wb.Worksheets
        Call ProcessIpSourceSheet(ws)
    Next ws
    
    MsgBox "IP Subject Analysis generated for all eligible sheets.", vbInformation
End Sub

'===========================================================
' PROCESS ONE SOURCE SHEET (IF IT IS IP RESULTS)
'===========================================================
Private Sub ProcessIpSourceSheet(ByVal wsSrc As Worksheet)
    Dim classCol As Long
    Dim levelCode As String
    Dim lastRow As Long, i As Long
    Dim firstClass As String
    Dim subjectCols() As Long
    Dim subjectNames() As String
    Dim subjCount As Long
    Dim header As String
    Dim examLabel As String
    Dim lastCol As Long
    Dim c As Long
    Dim wsDest As Worksheet
    Dim destSheetName As String
    Dim titleText As String
    Dim wb As Workbook
    Const BLOCK_HEIGHT As Long = 13
    
    On Error GoTo ExitPoint
    Set wb = ThisWorkbook
    
    '-----------------------------------------------------------
    ' 0) Skip irrelevant sheets
    '-----------------------------------------------------------
    If LCase$(wsSrc.Name) Like "*settings*" _
       Or LCase$(wsSrc.Name) Like "*config*" _
       Or LCase$(wsSrc.Name) Like "*menu*" _
       Or LCase$(wsSrc.Name) Like "*lookup*" _
       Or LCase$(wsSrc.Name) Like "*summary*" _
       Or LCase$(wsSrc.Name) Like "*template*" Then
           GoTo ExitPoint
    End If
    
    '-----------------------------------------------------------
    ' 1) Find Class column
    '-----------------------------------------------------------
    classCol = FindHeaderColumnIp(wsSrc, 1, "Class")
    If classCol = 0 Then GoTo ExitPoint
    
    '-----------------------------------------------------------
    ' 2) Get first class to determine IP level
    '-----------------------------------------------------------
    lastRow = wsSrc.Cells(wsSrc.Rows.count, classCol).End(xlUp).Row
    firstClass = ""
    For i = 2 To lastRow
        firstClass = Trim(CStr(wsSrc.Cells(i, classCol).value))
        If firstClass <> "" Then Exit For
    Next i
    If firstClass = "" Then GoTo ExitPoint
    
    If Not IsIpClass(firstClass) Then GoTo ExitPoint
    
    levelCode = "Y" & Left$(firstClass, 1) ' e.g. 3F ? Y3
    
    '-----------------------------------------------------------
    ' 3) Detect IP grade columns (A+..U)
    '-----------------------------------------------------------
    lastCol = wsSrc.Cells(1, wsSrc.Columns.count).End(xlToLeft).Column
    subjCount = 0
    
    For c = 1 To lastCol
        If c <> classCol Then
            header = Trim(CStr(wsSrc.Cells(1, c).value))
            If header <> "" Then
                If LooksLikeIpGradeColumn(wsSrc, c) Then
                    subjCount = subjCount + 1
                    ReDim Preserve subjectCols(1 To subjCount)
                    ReDim Preserve subjectNames(1 To subjCount)
                    subjectCols(subjCount) = c
                    subjectNames(subjCount) = header
                End If
            End If
        End If
    Next c
    
    If subjCount = 0 Then GoTo ExitPoint
    
    '-----------------------------------------------------------
    ' 4) exam label = sheet name
    '-----------------------------------------------------------
    examLabel = wsSrc.Name
    
    '-----------------------------------------------------------
    ' 5) Create / clear destination sheet
    '-----------------------------------------------------------
    destSheetName = BuildIpDestSheetName(levelCode, examLabel)

    
    On Error Resume Next
    Set wsDest = wb.Worksheets(destSheetName)
    On Error GoTo ExitPoint
    
    If wsDest Is Nothing Then
        Set wsDest = wb.Worksheets.Add(After:=wb.Sheets(wb.Sheets.count))
        wsDest.Name = destSheetName
    Else
        wsDest.Cells.Clear
        Dim k As Long
        Dim shp As Shape
        
        For k = wsDest.ChartObjects.count To 1 Step -1
            wsDest.ChartObjects(k).Delete
        Next k
        
        For k = wsDest.Shapes.count To 1 Step -1
            Set shp = wsDest.Shapes(k)
            If Left$(shp.Name, 12) = "IpFlagPanel_" Then
                shp.Delete
            End If
        Next k
    End If
    
    '-----------------------------------------------------------
    ' 6) Write main heading
    '-----------------------------------------------------------
    titleText = levelCode & " Subject Grade Distribution (" & examLabel & ") Ð IP Track"
    With wsDest.Range("A1")
        .value = titleText
        .Font.Bold = True
        .Font.Size = 14
    End With
    
    '-----------------------------------------------------------
    ' 7) Build all subject blocks
    '-----------------------------------------------------------
    Dim destRowHeader As Long
    Dim topLeft As String
    For i = 1 To subjCount
        destRowHeader = 3 + (i - 1) * BLOCK_HEIGHT
        topLeft = wsDest.Cells(destRowHeader, 1).Address(False, False)
        
        Call BuildIpSubjectGradeDistribution( _
                wsSrc.Name, classCol, subjectCols(i), _
                wsDest.Name, topLeft, subjectNames(i))
    Next i
    
ExitPoint:
End Sub

'===========================================================
' BUILD ONE IP SUBJECT TABLE + CHART + VALIDITY PANEL
'===========================================================
Private Sub BuildIpSubjectGradeDistribution( _
    ByVal srcSheetName As String, _
    ByVal srcClassCol As Long, _
    ByVal srcGradeCol As Long, _
    ByVal destSheetName As String, _
    ByVal destTopLeft As String, _
    ByVal subjectTitle As String)

    Dim wb As Workbook
    Dim wsSrc As Worksheet, wsDest As Worksheet
    Dim lastRow As Long, r As Long
    Dim className As String, gradeStr As String
    
    Dim gradeLabels(1 To 9) As String
    Dim pastelColors(1 To 9) As Long
    
    Dim classDict As Object          ' Scripting.Dictionary
    Dim countsArr() As Long          ' per-class counts A+..U
    Dim totalArr() As Long           ' cohort totals A+..U
    
    Dim classKey As Variant
    Dim classList() As String
    Dim i As Long, j As Long
    
    Dim destRowHeader As Long, destColFirst As Long
    Dim rowPtr As Long, cohortRow As Long
    
    Dim total As Long
    Dim passCount As Long, failCount As Long, topCount As Long
    Dim gpaValue As Double
    
    Dim rngTable As Range
    Dim rngHeader As Range, rngData As Range
    Dim rngCohortRow As Range
    
    Dim co As ChartObject
    Dim ch As Chart
    Dim leftPos As Double, topPos As Double, chartWidth As Double, chartHeight As Double
    
    Dim s As Series, pt As Point
    Dim titleCell As Range
    
    ' validity engine outputs
    Dim validityFlag As String, patternType As String
    Dim l1 As String, l2 As String, l3 As String
    
    On Error GoTo ExitPoint
    
    Set wb = ThisWorkbook
    Set wsSrc = wb.Worksheets(srcSheetName)
    Set wsDest = wb.Worksheets(destSheetName)
    
    '--- Grade labels mapping (IP)
    gradeLabels(1) = "A+"
    gradeLabels(2) = "A"
    gradeLabels(3) = "B+"
    gradeLabels(4) = "B"
    gradeLabels(5) = "C+"
    gradeLabels(6) = "C"
    gradeLabels(7) = "D+"
    gradeLabels(8) = "D"
    gradeLabels(9) = "U"
    
    '--- Performance palette (IP Ð echoing SEC pastel feel)
    pastelColors(1) = RGB(0, 150, 136)   ' A+
    pastelColors(2) = RGB(77, 182, 172)  ' A
    pastelColors(3) = RGB(129, 199, 132) ' B+
    pastelColors(4) = RGB(200, 230, 201) ' B
    pastelColors(5) = RGB(255, 245, 157) ' C+
    pastelColors(6) = RGB(255, 224, 130) ' C
    pastelColors(7) = RGB(255, 204, 128) ' D+
    pastelColors(8) = RGB(255, 171, 145) ' D
    pastelColors(9) = RGB(239, 83, 80)   ' U
    
    '--- Build dictionary: key = Class, item = Long(1..9) counts
    Set classDict = CreateObject("Scripting.Dictionary")
    classDict.CompareMode = 1 ' TextCompare
    
    lastRow = wsSrc.Cells(wsSrc.Rows.count, srcClassCol).End(xlUp).Row
    
    For r = 2 To lastRow
        className = Trim(CStr(wsSrc.Cells(r, srcClassCol).value))
        gradeStr = UCase$(Trim$(CStr(wsSrc.Cells(r, srcGradeCol).value)))
        
        If className <> "" And gradeStr <> "" Then
            If IsIpClass(className) Then
                j = GradeIndexIp(gradeStr, gradeLabels)
                If j >= 1 And j <= 9 Then
                    If Not classDict.Exists(className) Then
                        ReDim countsArr(1 To 9)
                        classDict.Add className, countsArr
                    End If
                    
                    countsArr = classDict(className)
                    countsArr(j) = countsArr(j) + 1
                    classDict(className) = countsArr
                End If
            End If
        End If
    Next r
    
    If classDict.count = 0 Then GoTo ExitPoint
    
    '--- Sort class keys alphabetically
    ReDim classList(1 To classDict.count)
    i = 1
    For Each classKey In classDict.Keys
        classList(i) = CStr(classKey)
        i = i + 1
    Next classKey
    Call SortStringArrayIp(classList)
    
    '--- Determine header position from DEST_TOP_LEFT
    destRowHeader = wsDest.Range(destTopLeft).Row
    destColFirst = wsDest.Range(destTopLeft).Column
    
    '--- Subject title (one row above header)
    Set titleCell = wsDest.Cells(destRowHeader - 1, destColFirst)
    With titleCell
        .value = subjectTitle
        .Font.Bold = True
        .Font.Size = 11
    End With
    
    '--- Write header row
    With wsDest
        .Cells(destRowHeader, destColFirst + 0).value = "Class"
        .Cells(destRowHeader, destColFirst + 1).value = "A+"
        .Cells(destRowHeader, destColFirst + 2).value = "A"
        .Cells(destRowHeader, destColFirst + 3).value = "B+"
        .Cells(destRowHeader, destColFirst + 4).value = "B"
        .Cells(destRowHeader, destColFirst + 5).value = "C+"
        .Cells(destRowHeader, destColFirst + 6).value = "C"
        .Cells(destRowHeader, destColFirst + 7).value = "D+"
        .Cells(destRowHeader, destColFirst + 8).value = "D"
        .Cells(destRowHeader, destColFirst + 9).value = "U"
        .Cells(destRowHeader, destColFirst + 10).value = "No."
        .Cells(destRowHeader, destColFirst + 11).value = "%Pass (A+ÐC)"
        .Cells(destRowHeader, destColFirst + 12).value = "%Fail (D+ÐU)"
        .Cells(destRowHeader, destColFirst + 13).value = "%Top (A+ / A)"
        .Cells(destRowHeader, destColFirst + 14).value = "GPA"
        .Rows(destRowHeader).Font.Bold = True
    End With
    
    '--- Write class rows + accumulate cohort totals
    ReDim totalArr(1 To 9)
    rowPtr = destRowHeader + 1
    
    For i = LBound(classList) To UBound(classList)
        classKey = classList(i)
        countsArr = classDict(classKey)
        
        total = 0
        passCount = 0
        failCount = 0
        topCount = 0
        
        For j = 1 To 9
            total = total + countsArr(j)
            totalArr(j) = totalArr(j) + countsArr(j)
            
            Select Case j
                Case 1 To 6       ' A+..C = pass
                    passCount = passCount + countsArr(j)
                Case 7 To 9       ' D+..U = fail
                    failCount = failCount + countsArr(j)
            End Select
            
            If j = 1 Or j = 2 Then
                topCount = topCount + countsArr(j)
            End If
        Next j
        
        If total > 0 Then
            gpaValue = ComputeIpGpa(countsArr)
        Else
            gpaValue = 0
        End If
        
        With wsDest
            .Cells(rowPtr, destColFirst + 0).value = classKey
            For j = 1 To 9
                .Cells(rowPtr, destColFirst + j).value = countsArr(j)
            Next j
            
            .Cells(rowPtr, destColFirst + 10).value = total  ' No.
            
            If total > 0 Then
                .Cells(rowPtr, destColFirst + 11).value = Round(passCount * 100# / total, 1)
                .Cells(rowPtr, destColFirst + 12).value = Round(failCount * 100# / total, 1)
                .Cells(rowPtr, destColFirst + 13).value = Round(topCount * 100# / total, 1)
                .Cells(rowPtr, destColFirst + 14).value = Round(gpaValue, 2)
            Else
                .Cells(rowPtr, destColFirst + 11).ClearContents
                .Cells(rowPtr, destColFirst + 12).ClearContents
                .Cells(rowPtr, destColFirst + 13).ClearContents
                .Cells(rowPtr, destColFirst + 14).ClearContents
            End If
        End With
        
        ColourIpSubjectRow wsDest, rowPtr, destColFirst
        
        rowPtr = rowPtr + 1
    Next i
    
    '--- Cohort row
    cohortRow = rowPtr
    total = 0
    passCount = 0
    failCount = 0
    topCount = 0
    
    For j = 1 To 9
        total = total + totalArr(j)
        
        Select Case j
            Case 1 To 6
                passCount = passCount + totalArr(j)
            Case 7 To 9
                failCount = failCount + totalArr(j)
        End Select
        
        If j = 1 Or j = 2 Then
            topCount = topCount + totalArr(j)
        End If
    Next j
    
    If total > 0 Then
        ReDim countsArr(1 To 9)
        For j = 1 To 9
            countsArr(j) = totalArr(j)
        Next j
        gpaValue = ComputeIpGpa(countsArr)
    Else
        gpaValue = 0
    End If
    
    With wsDest
        .Cells(cohortRow, destColFirst + 0).value = "COHORT"
        For j = 1 To 9
            .Cells(cohortRow, destColFirst + j).value = totalArr(j)
        Next j
        
        .Cells(cohortRow, destColFirst + 10).value = total
        
        If total > 0 Then
            .Cells(cohortRow, destColFirst + 11).value = Round(passCount * 100# / total, 1)
            .Cells(cohortRow, destColFirst + 12).value = Round(failCount * 100# / total, 1)
            .Cells(cohortRow, destColFirst + 13).value = Round(topCount * 100# / total, 1)
            .Cells(cohortRow, destColFirst + 14).value = Round(gpaValue, 2)
        Else
            .Cells(cohortRow, destColFirst + 11).ClearContents
            .Cells(cohortRow, destColFirst + 12).ClearContents
            .Cells(cohortRow, destColFirst + 13).ClearContents
            .Cells(cohortRow, destColFirst + 14).ClearContents
        End If
    End With
    
    ColourIpSubjectRow wsDest, cohortRow, destColFirst
    
    '--- Formatting: borders, number formats
    Set rngTable = wsDest.Range( _
        wsDest.Cells(destRowHeader, destColFirst), _
        wsDest.Cells(cohortRow, destColFirst + 14))
    
    With rngTable.Borders
        .LineStyle = xlContinuous
        .Color = RGB(200, 200, 200)
        .Weight = xlThin
    End With
    
    With wsDest
        .Range(.Cells(destRowHeader + 1, destColFirst + 11), _
               .Cells(cohortRow, destColFirst + 13)).NumberFormat = "0.0"
        .Range(.Cells(destRowHeader + 1, destColFirst + 14), _
               .Cells(cohortRow, destColFirst + 14)).NumberFormat = "0.00"
    End With
    
    wsDest.Columns(destColFirst + 1).Resize(, 14).AutoFit
    wsDest.Columns(destColFirst).ColumnWidth = 7
    
    ' Cohort row background
    Set rngCohortRow = wsDest.Range( _
        wsDest.Cells(cohortRow, destColFirst), _
        wsDest.Cells(cohortRow, destColFirst + 14))
    rngCohortRow.Interior.Color = RGB(255, 242, 204)
    rngCohortRow.Font.Bold = True
    
    '--- Build chart using COHORT row
    Set rngHeader = wsDest.Range( _
        wsDest.Cells(destRowHeader, destColFirst + 1), _
        wsDest.Cells(destRowHeader, destColFirst + 9))
    
    Set rngData = wsDest.Range( _
        wsDest.Cells(cohortRow, destColFirst + 1), _
        wsDest.Cells(cohortRow, destColFirst + 9))
    
    leftPos = wsDest.Columns(destColFirst + 16).Left
    topPos = wsDest.Rows(destRowHeader - 1).Top
    chartWidth = wsDest.Columns(destColFirst + 17).Resize(, 6).Width
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
        On Error GoTo 0
        
        .ChartArea.Format.line.Visible = msoFalse
        .PlotArea.Format.line.Visible = msoFalse
        .ChartArea.Format.Fill.ForeColor.RGB = RGB(255, 255, 255)
        .PlotArea.Format.Fill.Visible = msoFalse
        
        .SeriesCollection(1).HasDataLabels = True
        
        Set s = .SeriesCollection(1)
        For j = 1 To 9
            Set pt = s.Points(j)
            pt.Format.Fill.ForeColor.RGB = pastelColors(j)
            pt.Format.Fill.Solid
        Next j
        .ChartGroups(1).GapWidth = 30
    End With
    
    '=====================================================
    ' VALIDITY ENGINE CALL (IP Ð GPA-based)
    '=====================================================
    EvaluateIpDistribution totalArr, total, validityFlag, patternType, l1, l2, l3
    DrawIpValidityPanel wsDest, co, validityFlag, patternType, l1, l2, l3
    
ExitPoint:
End Sub

'---------------------------------------------------------
' HELPER: colour one IP subject row (class or cohort)
'---------------------------------------------------------
Private Sub ColourIpSubjectRow(ws As Worksheet, ByVal rowNum As Long, ByVal firstCol As Long)
    Dim v As Variant
    Dim c As Long
    
    ' A+ & A counts (cols firstCol+1, +2) ? green if >0
    For c = firstCol + 1 To firstCol + 2
        v = ws.Cells(rowNum, c).value
        If IsNumeric(v) And v > 0 Then
            ws.Cells(rowNum, c).Font.Color = RGB(0, 128, 0)
        Else
            ws.Cells(rowNum, c).Font.Color = RGB(0, 0, 0)
        End If
    Next c
    
    ' Fail counts D+ / D / U (cols firstCol+7 to firstCol+9) ? red if >0
    For c = firstCol + 7 To firstCol + 9
        v = ws.Cells(rowNum, c).value
        If IsNumeric(v) And v > 0 Then
            ws.Cells(rowNum, c).Font.Color = RGB(192, 0, 0)
        Else
            ws.Cells(rowNum, c).Font.Color = RGB(0, 0, 0)
        End If
    Next c
    
    ' %Pass (A+ÐC) Ð always black
    ws.Cells(rowNum, firstCol + 11).Font.Color = RGB(0, 0, 0)
    
    ' %Fail (D+ÐU) Ð red if >0
    v = ws.Cells(rowNum, firstCol + 12).value
    If IsNumeric(v) And v > 0 Then
        ws.Cells(rowNum, firstCol + 12).Font.Color = RGB(192, 0, 0)
    Else
        ws.Cells(rowNum, firstCol + 12).Font.Color = RGB(0, 0, 0)
    End If
    
    ' %Top (A+ / A) Ð green if >0
    v = ws.Cells(rowNum, firstCol + 13).value
    If IsNumeric(v) And v > 0 Then
        ws.Cells(rowNum, firstCol + 13).Font.Color = RGB(0, 128, 0)
    Else
        ws.Cells(rowNum, firstCol + 13).Font.Color = RGB(0, 0, 0)
    End If
    
    ' GPA Ð always black
    ws.Cells(rowNum, firstCol + 14).Font.Color = RGB(0, 0, 0)
End Sub


'===========================================================
' IP VALIDITY ENGINE (GPAÐBASED)
'===========================================================

Private Sub EvaluateIpDistribution( _
    ByRef gradeCounts() As Long, _
    ByVal totalN As Long, _
    ByRef validityFlag As String, _
    ByRef patternType As String, _
    ByRef line1 As String, _
    ByRef line2 As String, _
    ByRef line3 As String)

    Dim percents() As Double

    validityFlag = "NO DATA"
    patternType = "No Data"
    line1 = ""
    line2 = ""
    line3 = ""

    If totalN <= 0 Then
        GetIpInterpretationText validityFlag, patternType, line1, line2, line3
        Exit Sub
    End If

    percents = BuildPercentArrayIp(gradeCounts, totalN)

    validityFlag = GetIpValidityFlag(gradeCounts, percents, totalN)
    patternType = GetIpPatternType(percents, totalN, validityFlag)

    GetIpInterpretationText validityFlag, patternType, line1, line2, line3
End Sub


'===========================================================
' VALIDITY FLAG Ð MAIN FLAG DECISION
'===========================================================
Private Function GetIpValidityFlag( _
    ByRef counts() As Long, _
    ByRef pct() As Double, _
    ByVal totalN As Long) As String

    Dim topPct As Double, midPct As Double, botPct As Double
    Dim gpa As Double
    Dim i As Long

    If totalN <= LOW_N_THRESHOLD Then
        GetIpValidityFlag = "LOW N"
        Exit Function
    End If

    ' top bands A+ + A
    topPct = pct(1) + pct(2)

    ' middle bands B+..C
    midPct = pct(3) + pct(4) + pct(5) + pct(6)

    ' weak tail D+..U
    botPct = pct(7) + pct(8) + pct(9)

    ' compute GPA
    gpa = ComputeIpGpa(counts)

    ' SKEWED (high or low)
    If gpa >= SKEWED_HIGH_GPA Then
        GetIpValidityFlag = "SKEWED HIGH"
        Exit Function
    End If

    If gpa <= SKEWED_LOW_GPA Then
        GetIpValidityFlag = "SKEWED LOW"
        Exit Function
    End If

    ' TOP HEAVY
    If topPct >= TOP_HEAVY_PCT And botPct < 5 Then
        GetIpValidityFlag = "TOP HEAVY"
        Exit Function
    End If

    ' WEAK TAIL
    If botPct >= WEAK_TAIL_PCT Then
        GetIpValidityFlag = "WEAK TAIL"
        Exit Function
    End If

    ' MIXED (strong top + weak tail)
    If topPct >= 20# And botPct >= 15# Then
        GetIpValidityFlag = "MIXED"
        Exit Function
    End If

    GetIpValidityFlag = "VALID"
End Function


'===========================================================
' PATTERN TYPE Ð DETAILED PATTERN CLASSIFICATION
'===========================================================
Private Function GetIpPatternType( _
    ByRef pct() As Double, _
    ByVal totalN As Long, _
    ByVal validityFlag As String) As String

    Dim topPct As Double, midPct As Double, botPct As Double
    Dim i As Long
    Dim signChanges As Long
    Dim delta As Double, lastDelta As Double

    If totalN <= 0 Then
        GetIpPatternType = "No Data"
        Exit Function
    End If

    topPct = pct(1) + pct(2)
    midPct = pct(3) + pct(4) + pct(5) + pct(6)
    botPct = pct(7) + pct(8) + pct(9)

    ' LOW N
    If UCase$(validityFlag) = "LOW N" Then
        GetIpPatternType = "Small Cohort"
        Exit Function
    End If

    ' TOP HEAVY
    If UCase$(validityFlag) = "TOP HEAVY" Then
        GetIpPatternType = "Top-Heavy"
        Exit Function
    End If

    ' WEAK TAIL
    If UCase$(validityFlag) = "WEAK TAIL" Then
        GetIpPatternType = "Weak Tail"
        Exit Function
    End If

    ' FAT MIDDLE
    If midPct >= FAT_MIDDLE_PCT Then
        GetIpPatternType = "Fat Middle"
        Exit Function
    End If

    ' THIN MIDDLE
    If midPct <= THIN_MIDDLE_PCT And topPct > 20# And botPct > 10# Then
        GetIpPatternType = "Thin Middle"
        Exit Function
    End If

    ' WIDE SPREAD
    Dim countBands As Long
    countBands = 0
    For i = 1 To 9
        If pct(i) >= 5# Then countBands = countBands + 1
    Next i
    If countBands >= 6 Then
        GetIpPatternType = "Wide Spread"
        Exit Function
    End If

    ' STEPPED vs SPIKY
    lastDelta = 0#
    signChanges = 0
    
    For i = 1 To 8
        delta = pct(i + 1) - pct(i)
        If Abs(delta) > 5# Then
            If lastDelta <> 0# Then
                If Sgn(delta) <> Sgn(lastDelta) Then
                    signChanges = signChanges + 1
                End If
            End If
            lastDelta = delta
        End If
    Next i

    If signChanges >= 3 Then
        GetIpPatternType = "Spiky"
        Exit Function
    End If

    If signChanges = 0 And pct(1) >= pct(9) Then
        GetIpPatternType = "Stepped"
        Exit Function
    End If

    GetIpPatternType = "Balanced"
End Function


'===========================================================
' INTERPRETATION TEXT (What you see / What it means / What to do)
'===========================================================
Private Sub GetIpInterpretationText( _
    ByVal validityFlag As String, _
    ByVal patternType As String, _
    ByRef l1 As String, _
    ByRef l2 As String, _
    ByRef l3 As String)

    Dim f As String, p As String
    f = UCase$(Trim$(validityFlag))
    p = Trim$(patternType)

    Select Case f

        Case "NO DATA"
            l1 = "What you see: No grades recorded for this subject."
            l2 = "What it means: No distribution can be formed."
            l3 = "What you can do: Ensure results have been entered correctly."

        Case "LOW N"
            l1 = "What you see: Very small cohort."
            l2 = "What it means: The pattern can shift easily with a few students."
            l3 = "What you can do: Focus more on individual learning needs."

        Case "SKEWED HIGH"
            l1 = "What you see: Distribution heavily concentrated in A+/A/B+."
            l2 = "What it means: The cohort shows strong mastery."
            l3 = "What you can do: Provide extension tasks to stretch learners."

        Case "SKEWED LOW"
            l1 = "What you see: Many students in D+ / D / U."
            l2 = "What it means: Overall attainment is low for this assessment."
            l3 = "What you can do: Reinforce core concepts; revisit prerequisites."

        Case "TOP HEAVY"
            l1 = "What you see: Large cluster at A+ and A."
            l2 = "What it means: Very strong performance concentrated at the top."
            l3 = "What you can do: Differentiate for high-ability learners."

        Case "WEAK TAIL"
            l1 = "What you see: Significant proportion in D+ / D / U."
            l2 = "What it means: Many learners are struggling with key concepts."
            l3 = "What you can do: Target remediation and scaffolded practice."

        Case "MIXED"
            l1 = "What you see: Strong top end and significant weak tail."
            l2 = "What it means: Wide variation in readiness."
            l3 = "What you can do: Use flexible grouping and differentiated tasks."

        Case "VALID"
            Select Case p
                Case "Fat Middle"
                    l1 = "What you see: Many students in B+ to C bands."
                    l2 = "What it means: Assessment matched cohort readiness."
                    l3 = "What you can do: Strengthen the middle and lift into A bands."

                Case "Thin Middle"
                    l1 = "What you see: Few students in B/C range."
                    l2 = "What it means: Cohort polarised between high and low."
                    l3 = "What you can do: Support struggling learners while extending top."

                Case "Wide Spread"
                    l1 = "What you see: Results across many grade bands."
                    l2 = "What it means: Highly diverse cohort."
                    l3 = "What you can do: Use tiered instruction and task variety."

                Case "Spiky"
                    l1 = "What you see: Uneven bars across grades."
                    l2 = "What it means: Some items much easier/harder."
                    l3 = "What you can do: Review item difficulty balance."

                Case "Stepped"
                    l1 = "What you see: Gradual decline from A+ to U."
                    l2 = "What it means: Cohort performance spreads normally."
                    l3 = "What you can do: Focus on shifting middle bands upward."

                Case Else
                    l1 = "What you see: Grades form a stable pattern."
                    l2 = "What it means: Assessment differentiates appropriately."
                    l3 = "What you can do: Use data confidently for planning."
            End Select

        Case Else
            l1 = "What you see: Pattern does not fit predefined categories."
            l2 = "What it means: Interpretation may require more context."
            l3 = "What you can do: Combine data with classroom evidence."

    End Select

End Sub


'===========================================================
' BUILD PERCENT ARRAY (IP)
'===========================================================
Private Function BuildPercentArrayIp( _
    ByRef counts() As Long, _
    ByVal totalN As Long) As Double()

    Dim arr() As Double
    Dim i As Long
    
    ReDim arr(1 To 9)
    
    If totalN <= 0 Then
        For i = 1 To 9
            arr(i) = 0#
        Next i
        BuildPercentArrayIp = arr
        Exit Function
    End If

    For i = 1 To 9
        arr(i) = (counts(i) / totalN) * 100#
    Next i
    
    BuildPercentArrayIp = arr
End Function


'===========================================================
' SAFE SUM (IP)
'===========================================================
Private Function SafeSumIp( _
    ByRef arr() As Double, _
    ByVal a As Long, _
    ByVal b As Long) As Double

    Dim i As Long, lo As Long, hi As Long, s As Double
    lo = 1: hi = 9
    If a < lo Then a = lo
    If b > hi Then b = hi
    
    For i = a To b
        s = s + arr(i)
    Next i
    
    SafeSumIp = s
End Function



'===========================================================
' DRAW VALIDITY PANEL (ROUNDED RECTANGLE)
'===========================================================
Private Sub DrawIpValidityPanel( _
    ws As Worksheet, _
    co As ChartObject, _
    validityFlag As String, _
    patternType As String, _
    l1 As String, _
    l2 As String, _
    l3 As String)

    Dim panelLeft As Double, panelTop As Double
    Dim panelWidth As Double, panelHeight As Double
    Dim widthFactor As Double
    Dim shp As Shape
    Dim txt As String
    Dim fontSize As Single
    Dim fillColor As Long, borderColor As Long, fontColor As Long

    If co Is Nothing Then Exit Sub
    If co.Width <= 0 Or co.Height <= 0 Then Exit Sub

    widthFactor = GetIpPanelWidthFactor()

    panelHeight = co.Height
    panelWidth = 5 * co.Width * widthFactor
    panelLeft = co.Left + co.Width + 10
    panelTop = co.Top

    txt = l1 & vbCrLf & vbCrLf & l2 & vbCrLf & vbCrLf & l3

    If Len(txt) > 380 Then
        fontSize = 10
    ElseIf Len(txt) > 260 Then
        fontSize = 11
    Else
        fontSize = 12
    End If

    Select Case UCase$(validityFlag)
        Case "LOW N"
            fillColor = RGB(255, 242, 204)
            borderColor = RGB(191, 144, 0)
            fontColor = RGB(120, 63, 4)

        Case "SKEWED HIGH"
            fillColor = RGB(226, 240, 217)
            borderColor = RGB(118, 146, 60)
            fontColor = RGB(56, 87, 35)

        Case "SKEWED LOW"
            fillColor = RGB(252, 228, 214)
            borderColor = RGB(192, 80, 77)
            fontColor = RGB(148, 55, 49)

        Case "TOP HEAVY"
            fillColor = RGB(217, 225, 242)
            borderColor = RGB(79, 129, 189)
            fontColor = RGB(47, 84, 150)

        Case "WEAK TAIL"
            fillColor = RGB(255, 229, 229)
            borderColor = RGB(166, 77, 77)
            fontColor = RGB(128, 0, 0)

        Case "MIXED"
            fillColor = RGB(237, 234, 246)
            borderColor = RGB(112, 48, 160)
            fontColor = RGB(76, 37, 115)

        Case "VALID"
            fillColor = RGB(226, 240, 217)
            borderColor = RGB(118, 146, 60)
            fontColor = RGB(55, 86, 35)

        Case Else
            fillColor = RGB(242, 242, 242)
            borderColor = RGB(166, 166, 166)
            fontColor = RGB(89, 89, 89)

    End Select

    Set shp = ws.Shapes.AddShape(SHAPE_ROUNDED_RECTANGLE, panelLeft, panelTop, panelWidth, panelHeight)

    With shp
        .Name = "IpFlagPanel_" & co.Name & "_" & ws.Index
        .Fill.ForeColor.RGB = fillColor
        .line.ForeColor.RGB = borderColor
        .line.Weight = 1

        With .TextFrame2
            .TextRange.text = txt
            .TextRange.Font.Size = fontSize
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


'===========================================================
' PANEL WIDTH FACTOR FROM SETTINGS (CELL L5)
'===========================================================
Private Function GetIpPanelWidthFactor() As Double
    Dim ws As Worksheet
    Dim v As Variant

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("Settings")
    On Error GoTo 0

    If ws Is Nothing Then
        GetIpPanelWidthFactor = 0.35
        Exit Function
    End If

    v = ws.Range("L5").value
    If IsNumeric(v) And v > 0 And v < 1 Then
        GetIpPanelWidthFactor = v
    Else
        GetIpPanelWidthFactor = 0.35
    End If
End Function


'===========================================================
' IP GPA CALCULATION
'===========================================================
Private Function ComputeIpGpa(counts() As Long) As Double
    Dim weights(1 To 9) As Double
    Dim i As Long, total As Long
    Dim sum As Double

    weights(1) = 4#    ' A+
    weights(2) = 4#    ' A
    weights(3) = 3.5   ' B+
    weights(4) = 3#    ' B
    weights(5) = 2.5   ' C+
    weights(6) = 2#    ' C
    weights(7) = 1.5   ' D+
    weights(8) = 1#    ' D
    weights(9) = 0#    ' U

    For i = 1 To 9
        sum = sum + counts(i) * weights(i)
        total = total + counts(i)
    Next i

    If total > 0 Then
        ComputeIpGpa = sum / total
    Else
        ComputeIpGpa = 0
    End If
End Function


'===========================================================
' MAP IP GRADE STRING ? INDEX 1..9
'===========================================================
Private Function GradeIndexIp(ByVal g As String, ByRef labels() As String) As Long
    Dim i As Long
    For i = 1 To 9
        If g = labels(i) Then
            GradeIndexIp = i
            Exit Function
        End If
    Next i
    GradeIndexIp = 0
End Function


'===========================================================
' IS THIS CLASS AN IP CLASS? (1FÐ1I, 2FÐ2I, 3FÐ3J, 4FÐ4J)
'===========================================================
Private Function IsIpClass(ByVal className As String) As Boolean
    Dim lvl As Long
    Dim sec As String

    If Len(className) < 2 Then Exit Function

    lvl = val(Left$(className, 1))
    sec = UCase$(Mid$(className, 2, 1))

    If lvl < 1 Or lvl > 4 Then Exit Function

    Select Case lvl
        Case 1
            If sec >= "F" And sec <= "I" Then IsIpClass = True
        Case 2
            If sec >= "F" And sec <= "I" Then IsIpClass = True
        Case 3
            If sec >= "F" And sec <= "J" Then IsIpClass = True
        Case 4
            If sec >= "F" And sec <= "J" Then IsIpClass = True
    End Select
End Function


'===========================================================
' DETECT IF COLUMN LOOKS LIKE IP GRADE COLUMN
'===========================================================
Private Function LooksLikeIpGradeColumn(ws As Worksheet, gradeCol As Long) As Boolean
    Dim allowed As Object
    Dim r As Long, lastRow As Long
    Dim v As String
    Dim countValid As Long
    Dim maxSamples As Long

    Set allowed = CreateObject("Scripting.Dictionary")
    allowed.CompareMode = 1

    allowed("A+") = True
    allowed("A") = True
    allowed("B+") = True
    allowed("B") = True
    allowed("C+") = True
    allowed("C") = True
    allowed("D+") = True
    allowed("D") = True
    allowed("U") = True

    lastRow = ws.Cells(ws.Rows.count, gradeCol).End(xlUp).Row
    maxSamples = 200

    For r = 2 To lastRow
        v = UCase$(Trim$(CStr(ws.Cells(r, gradeCol).value)))

        If v <> "" Then
            If allowed.Exists(v) Then
                countValid = countValid + 1
                If countValid >= 3 Then Exit For
            End If
        End If

        If r - 1 >= maxSamples Then Exit For
    Next r

    LooksLikeIpGradeColumn = (countValid >= 3)
End Function


'===========================================================
' FIND HEADER IN ROW 1
'===========================================================
Private Function FindHeaderColumnIp(ws As Worksheet, headerRow As Long, headerName As String) As Long
    Dim lastCol As Long, c As Long
    Dim h As String

    lastCol = ws.Cells(headerRow, ws.Columns.count).End(xlToLeft).Column

    For c = 1 To lastCol
        h = Trim(CStr(ws.Cells(headerRow, c).value))
        If StrComp(h, headerName, vbTextCompare) = 0 Then
            FindHeaderColumnIp = c
            Exit Function
        End If
    Next c

    FindHeaderColumnIp = 0
End Function


'===========================================================
' SIMPLE SORT FOR STRING ARRAY
'===========================================================
Private Sub SortStringArrayIp(ByRef arr() As String)
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

'-----------------------------------------------------------
' HELPER: clean fragment of a sheet name (remove invalid chars)
'-----------------------------------------------------------
Private Function CleanSheetNameFragmentIp(ByVal txt As String) As String
    Dim s As String
    s = txt
    s = Replace(s, ":", "")
    s = Replace(s, "\", "")
    s = Replace(s, "/", "")
    s = Replace(s, "?", "")
    s = Replace(s, "*", "")
    s = Replace(s, "[", "")
    s = Replace(s, "]", "")
    CleanSheetNameFragmentIp = s
End Function

'-----------------------------------------------------------
' HELPER: build safe IP destination sheet name (<=31 chars)
'   Pattern:
'     <levelCode>_Subj Analysis_<shortLabel>_<year>
'   where <year> is preserved if examLabel ends in 4 digits.
'-----------------------------------------------------------
Private Function BuildIpDestSheetName(ByVal levelCode As String, _
                                      ByVal examLabel As String) As String
    Dim prefix As String
    Dim yearPart As String
    Dim baseLabel As String
    Dim safeBase As String
    Dim yearCandidate As String
    Dim maxShort As Long
    Dim safeName As String
    
    prefix = levelCode & "_Subj Analysis_"
    yearPart = ""
    
    ' Detect 4-digit year at end
    If Len(examLabel) >= 4 Then
        yearCandidate = Right$(examLabel, 4)
        If IsNumeric(yearCandidate) Then yearPart = yearCandidate
    End If
    
    If yearPart <> "" Then
        ' Remove year
        baseLabel = Left$(examLabel, Len(examLabel) - 4)
        
        ' Strip trailing separators
        Do While Len(baseLabel) > 0 And _
              (Right$(baseLabel, 1) = "_" Or _
               Right$(baseLabel, 1) = " " Or _
               Right$(baseLabel, 1) = "-")
            baseLabel = Left$(baseLabel, Len(baseLabel) - 1)
        Loop
        
        safeBase = CleanSheetNameFragmentIp(baseLabel)
        If safeBase = "" Then safeBase = "Exam"
        
        maxShort = 31 - Len(prefix) - 1 - Len(yearPart)
        If maxShort < 1 Then maxShort = 1
        
        If Len(safeBase) > maxShort Then safeBase = Left$(safeBase, maxShort)
        
        safeName = prefix & safeBase & "_" & yearPart
        
        If Len(safeName) > 31 Then safeName = Left$(safeName, 31)
    Else
        safeName = prefix & examLabel
        safeName = CleanSheetNameFragmentIp(safeName)
        If Len(safeName) > 31 Then safeName = Left$(safeName, 31)
    End If
    
    BuildIpDestSheetName = safeName
End Function

