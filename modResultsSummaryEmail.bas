Attribute VB_Name = "modResultsSummaryEmail"
Option Explicit

'============================================================
' Module: modResultsSummaryEmail
'
' PURPOSE
'   Build a compact HTML results summary from one or more SEC
'   staging sheets and open it as a Microsoft Outlook draft.
'
' SAFETY
'   This module never calls Send. The message is displayed as
'   a draft so that the user can review recipients and results.
'
' OPTIONAL SETTINGS (Settings!Q2:R30)
'   SchoolName          | School display name
'   PreparedBy          | Name/role shown in footer
'   EmailTo             | Default To recipients
'   EmailCC             | Default CC recipients
'   EmailSubjectPrefix  | Default: Results Summary
'
' OPTIONAL SUBJECT DISPLAY NAMES (Settings!T2:U100)
'   T = staging subject name/header; U = email display name
'
' MANAGEMENT EMAIL CONFIGURATION (Settings!S1:S7)
'   S2  MinimumCandidature=10
'   S3  ConcernBelow=70%
'   S4  MonitorBelow=80%
'   S5  StrongPassAtLeast=95%
'   S6  StrongDistinctionAtLeast=40%
'   S7  PrelimSummaryMode=ASK   (ASK / LEVEL / STREAM)
'============================================================

Private Type tEmailSubject
    DisplayName As String
    SourceName As String
    Scheme As String
    GradeCol As Long
    N As Long
    DistCount As Long
    PassCount As Long
    FailCount As Long
    PointSum As Double
    InvalidCount As Long
    AbCount As Long
End Type

Private Type tEmailStudent
    StudentName As String
    ClassName As String
    GroupCode As String
    DistCount As Long
    PassCount As Long
    FailCount As Long
    CountedSubjectCount As Long
    G3TakenCount As Long
    G3FailCount As Long
    G3DistCount As Long
    G2TakenCount As Long
    G2FailCount As Long
    G2DistCount As Long
    G1TakenCount As Long
    G1FailCount As Long
    G1DistCount As Long
End Type

Private Type tEmailGroupProfile
    StudentCount As Long
    PassAllCount As Long
    FailOneCount As Long
    FailTwoCount As Long
    FailThreePlusCount As Long
    DistOnePlusCount As Long
    DistTwoPlusCount As Long
    DistThreePlusCount As Long
End Type

Private Type tEmailSubjectResult
    LevelText As String
    DisplayName As String
    Scheme As String
    N As Long
    PassCount As Long
    DistCount As Long
End Type

Private Type tEmailManagementConfig
    MinCandidature As Long
    ConcernBelowPct As Double
    MonitorBelowPct As Double
    StrongPassAtLeastPct As Double
    StrongDistAtLeastPct As Double
End Type

Private Type tEmailLevelSummary
    LevelText As String
    YearText As String
    CandidateCount As Long
    SubjectCount As Long
    ValidEntries As Long
    PassEntries As Long
    PerfectSubjects As Long
    BelowNinetySubjects As Long
    ValidStudentCount As Long
    PassAllStudentCount As Long
    FailOneStudentCount As Long
    FailTwoStudentCount As Long
    FailThreePlusStudentCount As Long
    DistinctionStudentCount As Long
    NoDistinctionStudentCount As Long
    OneDistinctionStudentCount As Long
    TwoDistinctionStudentCount As Long
    ThreeDistinctionStudentCount As Long
    FourDistinctionStudentCount As Long
    FivePlusDistinctionStudentCount As Long
    G3Profile As tEmailGroupProfile
    G2Profile As tEmailGroupProfile
    G1Profile As tEmailGroupProfile
    HighlightsHtml As String
End Type

Private Const SETTINGS_SHEET As String = "Settings"
Private Const EMAIL_KEY_COL As String = "Q"
Private Const EMAIL_VALUE_COL As String = "R"
Private Const SUBJECT_MAP_KEY_COL As String = "T"
Private Const SUBJECT_MAP_VALUE_COL As String = "U"
Private Const DEFAULT_MIN_N As Long = 10
Private Const MGMT_MIN_N_CELL As String = "S2"
Private Const MGMT_CONCERN_CELL As String = "S3"
Private Const MGMT_MONITOR_CELL As String = "S4"
Private Const MGMT_STRONG_PASS_CELL As String = "S5"
Private Const MGMT_STRONG_DIST_CELL As String = "S6"
Private Const MGMT_PRELIM_MODE_CELL As String = "S7"
Private Const SUMMARY_DASHBOARD_BUTTON_NAME As String = "Nav_Summary"
Private Const SUMMARY_DASHBOARD_BUTTON_RANGE As String = "M1:S2"
Private Const SUMMARY_HOME_BUTTON_NAME As String = "HomeBtn"
Private Const SUMMARY_HOME_BUTTON_RANGE As String = "G1:H2"

Private gDraftAllLevels As Boolean
Private gAllLevelSheets As Collection
Private gSelectedExamLabel As String
Private gSelectedYear As String

'------------------------------------------------------------
' PUBLIC ENTRY POINT
'------------------------------------------------------------
Public Sub DraftResultsSummaryEmail()
    Dim wsSrc As Worksheet
    Dim subjects() As tEmailSubject
    Dim subjectCount As Long
    Dim students() As tEmailStudent
    Dim studentCount As Long
    Dim candidateCount As Long
    Dim warningText As String
    Dim htmlBody As String
    Dim assessmentName As String, yearText As String, levelText As String

    On Error GoTo ErrHandler

    Set wsSrc = SelectEmailSourceByAssessment()
    If wsSrc Is Nothing Then Exit Sub

    If gDraftAllLevels Then
        DraftAllLevelsManagementSummary
        Exit Sub
    End If

    CollectEmailSubjects wsSrc, subjects, subjectCount, warningText
    If subjectCount = 0 Then
        MsgBox "No recognised SEC grade columns were found on '" & wsSrc.Name & "'." & vbCrLf & _
               "Confirm that grades were imported and headers end in (Grade).", vbExclamation
        Exit Sub
    End If

    candidateCount = CountCandidates(wsSrc)
    CollectEmailStudents wsSrc, subjects, subjectCount, students, studentCount
    GetSourceLabels wsSrc, assessmentName, yearText, levelText
    assessmentName = PreferredExamLabel(CanonicalExamKey(assessmentName), assessmentName)

    htmlBody = BuildResultsEmailHtml(wsSrc.Name, assessmentName, yearText, levelText, _
                                     candidateCount, subjects, subjectCount, _
                                     students, studentCount, warningText)

    CreateOutlookDraft htmlBody, assessmentName, yearText, levelText
    Exit Sub

ErrHandler:
    MsgBox "Could not create the results-summary draft: " & Err.Description, vbCritical
End Sub

'------------------------------------------------------------
' SOURCE VALIDATION AND LABELS
'------------------------------------------------------------
Private Function IsEligibleEmailSourceSheet(ByVal ws As Object) As Boolean
    Dim lastCol As Long, lastRow As Long, classCol As Long, c As Long
    Dim hasName As Boolean, hasClass As Boolean, hasGrade As Boolean
    Dim hasSecGrade As Boolean
    Dim h As String

    If ws Is Nothing Then Exit Function
    If TypeName(ws) <> "Worksheet" Then Exit Function
    If IsIpEmailSheetName(ws.Name) Then Exit Function

    If FindEmailHeader(ws, "Name") > 0 Then hasName = True
    classCol = FindEmailHeader(ws, "Class")
    If classCol > 0 Then
        hasClass = True
        lastRow = ws.Cells(ws.Rows.count, classCol).End(xlUp).Row
    End If

    lastCol = ws.Cells(1, ws.Columns.count).End(xlToLeft).Column
    For c = 1 To lastCol
        h = UCase$(Trim$(CStr(ws.Cells(1, c).value)))
        If Right$(h, 7) = "(GRADE)" Then
            hasGrade = True
            If DetectEmailScheme(ws, c, lastRow, h) <> "" Then hasSecGrade = True
        End If
    Next c

    IsEligibleEmailSourceSheet = hasName And hasClass And hasGrade And hasSecGrade
End Function

Private Function IsIpEmailSheetName(ByVal sheetName As String) As Boolean
    Dim upperName As String, levelNo As Variant
    upperName = UCase$(Trim$(sheetName))

    For Each levelNo In Array("1", "2", "3", "4")
        If Left$(upperName, 2) = "Y" & CStr(levelNo) Then
            IsIpEmailSheetName = True
            Exit Function
        End If
        If InStr(1, upperName, "_Y" & CStr(levelNo) & "_", vbBinaryCompare) > 0 Then
            IsIpEmailSheetName = True
            Exit Function
        End If
    Next levelNo
End Function

'------------------------------------------------------------
' FRIENDLY ASSESSMENT / COHORT SELECTOR
'------------------------------------------------------------
Private Function SelectEmailSourceByAssessment() As Worksheet
    Dim ws As Worksheet, activeWs As Worksheet, candidateWs As Worksheet
    Dim examKeys() As String, examLabels() As String, examOrders() As Long
    Dim examCount As Long, i As Long, j As Long, existingIndex As Long
    Dim assessmentName As String, yearText As String, levelText As String
    Dim examKey As String, answerText As String, promptText As String
    Dim selectedExamIndex As Long, defaultExam As String
    Dim matchingSheets As Collection
    Dim tmpText As String, tmpOrder As Long
    Dim defaultCohort As String, selectedCohort As Long
    Dim matchedLevel As String, matchedYear As String, matchedAssessment As String
    Dim levelAnswer As String, matchCount As Long, oneMatch As Long
    Dim canOfferAll As Boolean, commonYear As String

    gDraftAllLevels = False
    Set gAllLevelSheets = Nothing
    gSelectedExamLabel = ""
    gSelectedYear = ""

    On Error Resume Next
    If TypeName(ActiveSheet) = "Worksheet" Then Set activeWs = ActiveSheet
    On Error GoTo 0

    ' Find the distinct assessments that really exist in staging sheets.
    For Each ws In ThisWorkbook.Worksheets
        If IsEligibleEmailSourceSheet(ws) Then
            assessmentName = "": yearText = "": levelText = ""
            GetSourceLabels ws, assessmentName, yearText, levelText
            examKey = CanonicalExamKey(assessmentName)

            If examKey <> "" Then
                existingIndex = FindExamKeyIndex(examKeys, examCount, examKey)
                If existingIndex = 0 Then
                    examCount = examCount + 1
                    ReDim Preserve examKeys(1 To examCount)
                    ReDim Preserve examLabels(1 To examCount)
                    ReDim Preserve examOrders(1 To examCount)
                    examKeys(examCount) = examKey
                    examLabels(examCount) = PreferredExamLabel(examKey, assessmentName)
                    examOrders(examCount) = PreferredExamOrder(examKey)
                End If

                If Not activeWs Is Nothing Then
                    If ws.Name = activeWs.Name Then defaultExam = examKey
                End If
            End If
        End If
    Next ws

    If examCount = 0 Then
        MsgBox "No available SEC assessments were found." & vbCrLf & _
               "Import results with grade columns before drafting the email.", _
               vbExclamation, "Draft Results Summary Email"
        Exit Function
    End If

    ' Preferred management order: WA1, WA2, First Combined, WA3,
    ' Prelim, 2nd Combined, EYE; then any other assessment names.
    For i = 1 To examCount - 1
        For j = i + 1 To examCount
            If examOrders(j) < examOrders(i) Or _
               (examOrders(j) = examOrders(i) And _
                StrComp(examLabels(j), examLabels(i), vbTextCompare) < 0) Then
                tmpText = examKeys(i): examKeys(i) = examKeys(j): examKeys(j) = tmpText
                tmpText = examLabels(i): examLabels(i) = examLabels(j): examLabels(j) = tmpText
                tmpOrder = examOrders(i): examOrders(i) = examOrders(j): examOrders(j) = tmpOrder
            End If
        Next j
    Next i

    promptText = "Available assessments:" & vbCrLf & vbCrLf
    For i = 1 To examCount
        promptText = promptText & i & ". " & examLabels(i) & vbCrLf
        If examKeys(i) = defaultExam Then defaultExam = CStr(i)
    Next i
    promptText = promptText & vbCrLf & "Type the number or assessment name (for example, WA3):"

    answerText = Trim$(InputBox(promptText, "Select Results Assessment", defaultExam))
    If answerText = "" Then Exit Function

    If IsNumeric(answerText) Then
        If CDbl(answerText) = Fix(CDbl(answerText)) Then selectedExamIndex = CLng(answerText)
        If selectedExamIndex < 1 Or selectedExamIndex > examCount Then selectedExamIndex = 0
    Else
        examKey = CanonicalExamKey(answerText)
        selectedExamIndex = FindExamKeyIndex(examKeys, examCount, examKey)
    End If

    If selectedExamIndex = 0 Then
        MsgBox "'" & answerText & "' is not one of the available assessments shown.", _
               vbExclamation, "Assessment Not Found"
        Exit Function
    End If

    Set matchingSheets = New Collection
    For Each ws In ThisWorkbook.Worksheets
        If IsEligibleEmailSourceSheet(ws) Then
            assessmentName = "": yearText = "": levelText = ""
            GetSourceLabels ws, assessmentName, yearText, levelText
            If CanonicalExamKey(assessmentName) = examKeys(selectedExamIndex) Then
                matchingSheets.Add ws
            End If
        End If
    Next ws

    Set matchingSheets = SortEmailSheetsByLevel(matchingSheets)

    If matchingSheets.count = 1 Then
        Set SelectEmailSourceByAssessment = matchingSheets(1)
        Exit Function
    End If

    ' More than one level/year has this assessment. Ask using friendly
    ' cohort labels so the user never has to know a staging-sheet name.
    promptText = examLabels(selectedExamIndex) & " is available for:" & vbCrLf & vbCrLf
    canOfferAll = MatchingSheetsShareOneYear(matchingSheets, commonYear)
    If canOfferAll Then
        promptText = promptText & "0. ALL LEVELS - Management Summary" & vbCrLf
    Else
        promptText = promptText & "ALL LEVELS is unavailable because these results span multiple years." & vbCrLf
    End If
    For i = 1 To matchingSheets.count
        Set candidateWs = matchingSheets(i)
        matchedAssessment = "": matchedYear = "": matchedLevel = ""
        GetSourceLabels candidateWs, matchedAssessment, matchedYear, matchedLevel
        promptText = promptText & i & ". " & matchedLevel
        If matchedYear <> "" Then promptText = promptText & " - " & matchedYear
        promptText = promptText & " (" & CountCandidates(candidateWs) & " candidates)" & vbCrLf
        If Not activeWs Is Nothing Then
            If candidateWs.Name = activeWs.Name Then defaultCohort = CStr(i)
        End If
    Next i
    promptText = promptText & vbCrLf & "Type 0 for all levels, or a number/level (for example, S4):"

    levelAnswer = Trim$(InputBox(promptText, "Select Results Cohort", defaultCohort))
    If levelAnswer = "" Then Exit Function

    If canOfferAll Then
        If UCase$(levelAnswer) = "ALL" Or UCase$(levelAnswer) = "ALL LEVELS" Or levelAnswer = "0" Then
            gDraftAllLevels = True
            Set gAllLevelSheets = matchingSheets
            gSelectedExamLabel = examLabels(selectedExamIndex)
            gSelectedYear = commonYear
            Set SelectEmailSourceByAssessment = matchingSheets(1)
            Exit Function
        End If
    End If

    If IsNumeric(levelAnswer) Then
        If CDbl(levelAnswer) = Fix(CDbl(levelAnswer)) Then selectedCohort = CLng(levelAnswer)
        If selectedCohort >= 1 And selectedCohort <= matchingSheets.count Then
            Set SelectEmailSourceByAssessment = matchingSheets(selectedCohort)
            Exit Function
        End If
    Else
        For i = 1 To matchingSheets.count
            Set candidateWs = matchingSheets(i)
            matchedAssessment = "": matchedYear = "": matchedLevel = ""
            GetSourceLabels candidateWs, matchedAssessment, matchedYear, matchedLevel
            If StrComp(Trim$(levelAnswer), matchedLevel, vbTextCompare) = 0 Then
                matchCount = matchCount + 1
                oneMatch = i
            End If
        Next i
        If matchCount = 1 Then
            Set SelectEmailSourceByAssessment = matchingSheets(oneMatch)
            Exit Function
        End If
    End If

    MsgBox "The cohort selection was not recognised. No draft was created.", _
           vbExclamation, "Cohort Not Found"
End Function

Private Function SortEmailSheetsByLevel(ByVal sourceSheets As Collection) As Collection
    Dim sortedSheets As New Collection
    Dim added As Object
    Dim levelNo As Long, i As Long
    Dim ws As Worksheet
    Dim assessmentName As String, yearText As String, levelText As String

    Set added = CreateObject("Scripting.Dictionary")
    added.CompareMode = vbTextCompare

    For levelNo = 1 To 5
        For i = 1 To sourceSheets.count
            Set ws = sourceSheets(i)
            assessmentName = "": yearText = "": levelText = ""
            GetSourceLabels ws, assessmentName, yearText, levelText
            If UCase$(levelText) = "S" & CStr(levelNo) Then
                sortedSheets.Add ws
                added(ws.Name) = True
            End If
        Next i
    Next levelNo

    For i = 1 To sourceSheets.count
        Set ws = sourceSheets(i)
        If Not added.Exists(ws.Name) Then sortedSheets.Add ws
    Next i

    Set SortEmailSheetsByLevel = sortedSheets
End Function

Private Function MatchingSheetsShareOneYear(ByVal matchingSheets As Collection, _
                                            ByRef commonYear As String) As Boolean
    Dim i As Long
    Dim ws As Worksheet
    Dim assessmentName As String, yearText As String, levelText As String

    commonYear = ""
    For i = 1 To matchingSheets.count
        Set ws = matchingSheets(i)
        assessmentName = "": yearText = "": levelText = ""
        GetSourceLabels ws, assessmentName, yearText, levelText

        If i = 1 Then
            commonYear = yearText
        ElseIf StrComp(commonYear, yearText, vbTextCompare) <> 0 Then
            MatchingSheetsShareOneYear = False
            Exit Function
        End If
    Next i

    MatchingSheetsShareOneYear = True
End Function

Private Function FindExamKeyIndex(ByRef examKeys() As String, _
                                  ByVal examCount As Long, ByVal examKey As String) As Long
    Dim i As Long
    For i = 1 To examCount
        If StrComp(examKeys(i), examKey, vbTextCompare) = 0 Then
            FindExamKeyIndex = i
            Exit Function
        End If
    Next i
End Function

Private Function CanonicalExamKey(ByVal assessmentName As String) As String
    Dim s As String, i As Long, ch As String, compact As String
    s = UCase$(Trim$(assessmentName))

    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        If (ch >= "A" And ch <= "Z") Or (ch >= "0" And ch <= "9") Then compact = compact & ch
    Next i

    If InStr(1, compact, "WA1", vbBinaryCompare) > 0 Or _
       InStr(1, compact, "TERM1WA", vbBinaryCompare) > 0 Or _
       InStr(1, compact, "TERM1NWA", vbBinaryCompare) > 0 Then
        CanonicalExamKey = "WA1"
    ElseIf InStr(1, compact, "WA2", vbBinaryCompare) > 0 Or _
           InStr(1, compact, "TERM2WA", vbBinaryCompare) > 0 Or _
           InStr(1, compact, "TERM2NWA", vbBinaryCompare) > 0 Then
        CanonicalExamKey = "WA2"
    ElseIf InStr(1, compact, "FIRSTCOMBINED", vbBinaryCompare) > 0 Or _
           InStr(1, compact, "1STCOMBINED", vbBinaryCompare) > 0 Or _
           InStr(1, compact, "COMBINED1", vbBinaryCompare) > 0 Or _
           InStr(1, compact, "SEMESTER1", vbBinaryCompare) > 0 Or _
           InStr(1, compact, "TERM2COMBINED", vbBinaryCompare) > 0 Then
        CanonicalExamKey = "FIRSTCOMBINED"
    ElseIf InStr(1, compact, "WA3", vbBinaryCompare) > 0 Or _
           InStr(1, compact, "TERM3WA", vbBinaryCompare) > 0 Or _
           InStr(1, compact, "TERM3NWA", vbBinaryCompare) > 0 Then
        CanonicalExamKey = "WA3"
    ElseIf InStr(1, compact, "PRELIM", vbBinaryCompare) > 0 Then
        CanonicalExamKey = "PRELIM"
    ElseIf InStr(1, compact, "SECONDCOMBINED", vbBinaryCompare) > 0 Or _
           InStr(1, compact, "2NDCOMBINED", vbBinaryCompare) > 0 Or _
           InStr(1, compact, "COMBINED2", vbBinaryCompare) > 0 Or _
           InStr(1, compact, "SEMESTER2", vbBinaryCompare) > 0 Or _
           InStr(1, compact, "TERM3COMBINED", vbBinaryCompare) > 0 Or _
           InStr(1, compact, "TERM4COMBINED", vbBinaryCompare) > 0 Then
        CanonicalExamKey = "SECONDCOMBINED"
    ElseIf InStr(1, compact, "EYE", vbBinaryCompare) > 0 Or _
           InStr(1, compact, "ENDOFYEAR", vbBinaryCompare) > 0 Then
        CanonicalExamKey = "EYE"
    Else
        CanonicalExamKey = compact
    End If
End Function

Private Function PreferredExamLabel(ByVal examKey As String, _
                                    ByVal originalLabel As String) As String
    Select Case examKey
        Case "WA1", "WA2", "WA3", "EYE": PreferredExamLabel = examKey
        Case "FIRSTCOMBINED": PreferredExamLabel = "First Combined"
        Case "PRELIM": PreferredExamLabel = "PRELIM"
        Case "SECONDCOMBINED": PreferredExamLabel = "2nd Combined"
        Case Else: PreferredExamLabel = originalLabel
    End Select
End Function

Private Function PreferredExamOrder(ByVal examKey As String) As Long
    Select Case examKey
        Case "WA1": PreferredExamOrder = 1
        Case "WA2": PreferredExamOrder = 2
        Case "FIRSTCOMBINED": PreferredExamOrder = 3
        Case "WA3": PreferredExamOrder = 4
        Case "PRELIM": PreferredExamOrder = 5
        Case "SECONDCOMBINED": PreferredExamOrder = 6
        Case "EYE": PreferredExamOrder = 7
        Case Else: PreferredExamOrder = 100
    End Select
End Function

Private Sub GetSourceLabels(ByVal ws As Worksheet, _
                            ByRef assessmentName As String, _
                            ByRef yearText As String, _
                            ByRef levelText As String)
    Dim assCol As Long, yearCol As Long, classCol As Long
    Dim lastRow As Long, r As Long
    Dim classText As String, ch As String

    assCol = FindEmailHeader(ws, "Assessment")
    yearCol = FindEmailHeader(ws, "Year")
    classCol = FindEmailHeader(ws, "Class")
    lastRow = ws.Cells(ws.Rows.count, classCol).End(xlUp).Row

    For r = 2 To lastRow
        If assessmentName = "" And assCol > 0 Then assessmentName = Trim$(CStr(ws.Cells(r, assCol).value))
        If yearText = "" And yearCol > 0 Then yearText = Trim$(CStr(ws.Cells(r, yearCol).value))
        If classText = "" Then classText = UCase$(Trim$(CStr(ws.Cells(r, classCol).value)))
        If assessmentName <> "" And yearText <> "" And classText <> "" Then Exit For
    Next r

    If assessmentName = "" Then assessmentName = ws.Name
    If yearText = "" Then yearText = ExtractEmailYear(ws.Name)

    ch = FirstLevelDigit(classText)
    If ch <> "" Then
        levelText = "S" & ch
    Else
        levelText = "SEC"
    End If
End Sub

Private Function FirstLevelDigit(ByVal valueText As String) As String
    Dim i As Long, ch As String
    For i = 1 To Len(valueText)
        ch = Mid$(valueText, i, 1)
        If ch >= "1" And ch <= "5" Then
            FirstLevelDigit = ch
            Exit Function
        End If
    Next i
End Function

Private Function ExtractEmailYear(ByVal valueText As String) As String
    Dim i As Long, token As String
    For i = 1 To Len(valueText) - 3
        token = Mid$(valueText, i, 4)
        If IsNumeric(token) Then
            If CLng(token) >= 2000 And CLng(token) <= 2100 Then
                ExtractEmailYear = token
                Exit Function
            End If
        End If
    Next i
End Function

'------------------------------------------------------------
' SUBJECT METRICS
'------------------------------------------------------------
Private Sub CollectEmailSubjects(ByVal ws As Worksheet, _
                                 ByRef subjects() As tEmailSubject, _
                                 ByRef subjectCount As Long, _
                                 ByRef warningText As String)
    Dim lastCol As Long, lastRow As Long, classCol As Long
    Dim c As Long, r As Long
    Dim headerText As String, sourceName As String, scheme As String
    Dim gradeText As String
    Dim isValid As Boolean, isPass As Boolean, isDist As Boolean
    Dim pointValue As Double
    Dim minN As Long

    classCol = FindEmailHeader(ws, "Class")
    lastRow = ws.Cells(ws.Rows.count, classCol).End(xlUp).Row
    lastCol = ws.Cells(1, ws.Columns.count).End(xlToLeft).Column
    minN = GetEmailMinN()

    For c = 1 To lastCol
        headerText = Trim$(CStr(ws.Cells(1, c).value))
        If UCase$(Right$(headerText, 7)) = "(GRADE)" Then
            sourceName = StripEmailGradeSuffix(headerText)
            If Not IsExcludedEmailSubject(sourceName) Then
                scheme = DetectEmailScheme(ws, c, lastRow, headerText)
                If scheme <> "" Then
                subjectCount = subjectCount + 1
                ReDim Preserve subjects(1 To subjectCount)

                With subjects(subjectCount)
                    .SourceName = sourceName
                    .DisplayName = GetEmailSubjectDisplayName(sourceName)
                    .Scheme = scheme
                    .GradeCol = c
                End With

                For r = 2 To lastRow
                    gradeText = UCase$(Trim$(CStr(ws.Cells(r, c).value)))
                    If gradeText <> "" Then
                        If gradeText = "AB" Then
                            With subjects(subjectCount)
                                .AbCount = .AbCount + 1
                                .N = .N + 1
                                .FailCount = .FailCount + 1
                            End With
                        ElseIf gradeText <> "VR" And gradeText <> "MC" And gradeText <> "-" Then
                            EvaluateEmailGrade scheme, gradeText, isValid, isPass, isDist, pointValue
                            If isValid Then
                                With subjects(subjectCount)
                                    .N = .N + 1
                                    If isPass Then .PassCount = .PassCount + 1
                                    If Not isPass Then .FailCount = .FailCount + 1
                                    If isDist Then .DistCount = .DistCount + 1
                                    .PointSum = .PointSum + pointValue
                                End With
                            Else
                                subjects(subjectCount).InvalidCount = subjects(subjectCount).InvalidCount + 1
                            End If
                        End If
                    End If
                Next r

                If subjects(subjectCount).N < minN Then
                    AppendWarning warningText, subjects(subjectCount).DisplayName & " (" & scheme & _
                                  ") has a small valid N of " & subjects(subjectCount).N & "."
                End If
                If subjects(subjectCount).InvalidCount > 0 Then
                    AppendWarning warningText, subjects(subjectCount).DisplayName & " (" & scheme & _
                                  ") contains " & subjects(subjectCount).InvalidCount & " unrecognised grade value(s)."
                End If
                If subjects(subjectCount).AbCount > 0 Then
                    AppendWarning warningText, subjects(subjectCount).DisplayName & " (" & scheme & _
                                  ") contains " & subjects(subjectCount).AbCount & " AB value(s), counted as failures."
                End If
                Else
                    AppendWarning warningText, "Skipped unrecognised grade column: " & headerText & "."
                End If
            End If
        End If
    Next c
End Sub

Private Function DetectEmailScheme(ByVal ws As Worksheet, ByVal gradeCol As Long, _
                                   ByVal lastRow As Long, ByVal headerText As String) As String
    Dim h As String, g As String
    Dim r As Long

    h = UCase$(Replace(headerText, " ", ""))
    If InStr(1, h, "-G3", vbTextCompare) > 0 Or InStr(1, h, "-EX", vbTextCompare) > 0 _
       Or InStr(1, h, "(EX)", vbTextCompare) > 0 Or InStr(1, h, "EXPRESS", vbTextCompare) > 0 _
       Or InStr(1, h, "-O", vbTextCompare) > 0 Then
        DetectEmailScheme = "G3"
        Exit Function
    End If
    If InStr(1, h, "-G2", vbTextCompare) > 0 Or InStr(1, h, "N(A)", vbTextCompare) > 0 _
       Or InStr(1, h, "-NA", vbTextCompare) > 0 Or InStr(1, h, "NORMALACADEMIC", vbTextCompare) > 0 Then
        DetectEmailScheme = "G2"
        Exit Function
    End If
    If InStr(1, h, "-G1", vbTextCompare) > 0 Or InStr(1, h, "N(T)", vbTextCompare) > 0 _
       Or InStr(1, h, "-NT", vbTextCompare) > 0 Or InStr(1, h, "NORMALTECH", vbTextCompare) > 0 Then
        DetectEmailScheme = "G1"
        Exit Function
    End If
    If InStr(1, h, "IP", vbTextCompare) > 0 Then Exit Function

    For r = 2 To lastRow
        g = UCase$(Trim$(CStr(ws.Cells(r, gradeCol).value)))
        If g <> "" And g <> "AB" And g <> "VR" And g <> "-" Then
            Select Case g
                Case "A1", "A2", "B3", "B4", "C5", "C6", "D7", "E8", "F9", "9"
                    DetectEmailScheme = "G3"
                Case "1", "2", "3", "4", "5", "6"
                    DetectEmailScheme = "G2"
                Case "A", "B", "C", "D", "E"
                    DetectEmailScheme = "G1"
            End Select
            If DetectEmailScheme <> "" Then Exit Function
        End If
    Next r
End Function

Private Sub EvaluateEmailGrade(ByVal scheme As String, ByVal gradeText As String, _
                               ByRef isValid As Boolean, ByRef isPass As Boolean, _
                               ByRef isDist As Boolean, ByRef pointValue As Double)
    ' Some source exports store the lowest G3 grade as the numeric text "9".
    ' Treat it as F9 so it remains in the candidature and counts as a failure,
    ' matching the Subject Analysis calculation.
    If UCase$(scheme) = "G3" And UCase$(Trim$(gradeText)) = "9" Then gradeText = "F9"

    isValid = True
    isPass = False
    isDist = False
    pointValue = 0

    Select Case UCase$(scheme)
        Case "G3"
            Select Case gradeText
                Case "A1": pointValue = 1: isPass = True: isDist = True
                Case "A2": pointValue = 2: isPass = True: isDist = True
                Case "B3": pointValue = 3: isPass = True
                Case "B4": pointValue = 4: isPass = True
                Case "C5": pointValue = 5: isPass = True
                Case "C6": pointValue = 6: isPass = True
                Case "D7": pointValue = 7
                Case "E8": pointValue = 8
                Case "F9": pointValue = 9
                Case Else: isValid = False
            End Select
        Case "G2"
            Select Case gradeText
                Case "1": pointValue = 1: isPass = True: isDist = True
                Case "2": pointValue = 2: isPass = True: isDist = True
                Case "3": pointValue = 3: isPass = True
                Case "4": pointValue = 4: isPass = True
                Case "5": pointValue = 5: isPass = True
                Case "6": pointValue = 6
                Case Else: isValid = False
            End Select
        Case "G1"
            Select Case gradeText
                Case "A": pointValue = 1: isPass = True: isDist = True
                Case "B": pointValue = 2: isPass = True
                Case "C": pointValue = 3: isPass = True
                Case "D": pointValue = 4: isPass = True
                Case "E": pointValue = 5
                Case Else: isValid = False
            End Select
        Case Else
            isValid = False
    End Select
End Sub

'------------------------------------------------------------
' CANDIDATE AND TOP-STUDENT METRICS
'------------------------------------------------------------
Private Function CountCandidates(ByVal ws As Worksheet) As Long
    Dim regCol As Long, nameCol As Long, classCol As Long
    Dim lastRow As Long, r As Long
    Dim keyText As String
    Dim regText As String, nameText As String, classText As String
    Dim seen As Object

    Set seen = CreateObject("Scripting.Dictionary")
    seen.CompareMode = vbTextCompare

    regCol = FindEmailHeader(ws, "RegNo")
    nameCol = FindEmailHeader(ws, "Name")
    classCol = FindEmailHeader(ws, "Class")
    lastRow = ws.Cells(ws.Rows.count, classCol).End(xlUp).Row

    For r = 2 To lastRow
        regText = "": nameText = "": classText = "": keyText = ""
        If regCol > 0 Then regText = Trim$(CStr(ws.Cells(r, regCol).value))
        If nameCol > 0 Then nameText = Trim$(CStr(ws.Cells(r, nameCol).value))
        classText = Trim$(CStr(ws.Cells(r, classCol).value))

        ' RegNo is normally a class register number (for example 1..40),
        ' so it is not unique across an entire level. Class + RegNo is.
        If regText <> "" Then
            keyText = classText & "|REG|" & regText
        ElseIf nameText <> "" Then
            keyText = classText & "|NAME|" & nameText
        End If

        If keyText <> "" Then
            If Not seen.Exists(keyText) Then seen.Add keyText, True
        End If
    Next r

    CountCandidates = seen.count
End Function

Private Sub CollectEmailStudents(ByVal ws As Worksheet, _
                                 ByRef subjects() As tEmailSubject, ByVal subjectCount As Long, _
                                 ByRef students() As tEmailStudent, ByRef studentCount As Long)
    Dim nameCol As Long, classCol As Long, regCol As Long, lastRow As Long
    Dim r As Long, i As Long
    Dim studentName As String, className As String, regText As String, keyText As String, gradeText As String
    Dim g1Taken As Long, g2Taken As Long, g3Taken As Long
    Dim g1Fail As Long, g2Fail As Long, g3Fail As Long
    Dim g1Dist As Long, g2Dist As Long, g3Dist As Long
    Dim isValid As Boolean, isPass As Boolean, isDist As Boolean
    Dim pointValue As Double
    Dim seen As Object

    nameCol = FindEmailHeader(ws, "Name")
    classCol = FindEmailHeader(ws, "Class")
    regCol = FindEmailHeader(ws, "RegNo")
    lastRow = ws.Cells(ws.Rows.count, classCol).End(xlUp).Row
    Set seen = CreateObject("Scripting.Dictionary")
    seen.CompareMode = vbTextCompare

    For r = 2 To lastRow
        studentName = Trim$(CStr(ws.Cells(r, nameCol).value))
        className = Trim$(CStr(ws.Cells(r, classCol).value))
        If studentName <> "" Then
            regText = ""
            If regCol > 0 Then regText = Trim$(CStr(ws.Cells(r, regCol).value))
            If regText <> "" Then
                keyText = className & "|REG|" & regText
            Else
                keyText = className & "|NAME|" & studentName
            End If
            If seen.Exists(keyText) Then GoTo NextStudentRow
            seen.Add keyText, True

            studentCount = studentCount + 1
            ReDim Preserve students(1 To studentCount)
            students(studentCount).StudentName = studentName
            students(studentCount).ClassName = className

            g1Taken = 0: g2Taken = 0: g3Taken = 0
            g1Fail = 0: g2Fail = 0: g3Fail = 0
            g1Dist = 0: g2Dist = 0: g3Dist = 0
            For i = 1 To subjectCount
                gradeText = UCase$(Trim$(CStr(ws.Cells(r, subjects(i).GradeCol).value)))
                If gradeText = "AB" Then
                    students(studentCount).CountedSubjectCount = students(studentCount).CountedSubjectCount + 1
                    students(studentCount).FailCount = students(studentCount).FailCount + 1
                    Select Case subjects(i).Scheme
                        Case "G1": g1Taken = g1Taken + 1: g1Fail = g1Fail + 1
                        Case "G2": g2Taken = g2Taken + 1: g2Fail = g2Fail + 1
                        Case "G3": g3Taken = g3Taken + 1: g3Fail = g3Fail + 1
                    End Select
                ElseIf gradeText <> "" And gradeText <> "VR" And gradeText <> "MC" And gradeText <> "-" Then
                    EvaluateEmailGrade subjects(i).Scheme, gradeText, isValid, isPass, isDist, pointValue
                    If isValid Then
                        students(studentCount).CountedSubjectCount = students(studentCount).CountedSubjectCount + 1
                        Select Case subjects(i).Scheme
                            Case "G1"
                                g1Taken = g1Taken + 1
                                If Not isPass Then g1Fail = g1Fail + 1
                                If isDist Then g1Dist = g1Dist + 1
                            Case "G2"
                                g2Taken = g2Taken + 1
                                If Not isPass Then g2Fail = g2Fail + 1
                                If isDist Then g2Dist = g2Dist + 1
                            Case "G3"
                                g3Taken = g3Taken + 1
                                If Not isPass Then g3Fail = g3Fail + 1
                                If isDist Then g3Dist = g3Dist + 1
                        End Select
                        If isPass Then students(studentCount).PassCount = students(studentCount).PassCount + 1
                        If Not isPass Then students(studentCount).FailCount = students(studentCount).FailCount + 1
                        If isDist Then students(studentCount).DistCount = students(studentCount).DistCount + 1
                    End If
                End If
            Next i

            If g1Taken > 0 Then
                students(studentCount).GroupCode = "G1"
            ElseIf g2Taken > 0 Then
                students(studentCount).GroupCode = "G2"
            ElseIf g3Taken > 0 Then
                students(studentCount).GroupCode = "G3"
            Else
                students(studentCount).GroupCode = ""
            End If
            With students(studentCount)
                .G3TakenCount = g3Taken
                .G3FailCount = g3Fail
                .G3DistCount = g3Dist
                .G2TakenCount = g2Taken
                .G2FailCount = g2Fail
                .G2DistCount = g2Dist
                .G1TakenCount = g1Taken
                .G1FailCount = g1Fail
                .G1DistCount = g1Dist
            End With
        End If
NextStudentRow:
    Next r
End Sub

'------------------------------------------------------------
' HTML DOCUMENT
'------------------------------------------------------------
Private Function BuildResultsEmailHtml(ByVal sourceSheetName As String, _
                                       ByVal assessmentName As String, ByVal yearText As String, _
                                       ByVal levelText As String, ByVal candidateCount As Long, _
                                       ByRef subjects() As tEmailSubject, ByVal subjectCount As Long, _
                                       ByRef students() As tEmailStudent, ByVal studentCount As Long, _
                                       ByVal warningText As String) As String
    Dim html As String
    Dim schoolName As String, preparedBy As String, embargoText As String
    Dim totalEntries As Long, totalPass As Long
    Dim perfectCount As Long, belowCount As Long
    Dim i As Long

    schoolName = GetEmailSetting("SchoolName", RemoveWorkbookExtension(ThisWorkbook.Name))
    preparedBy = GetEmailSetting("PreparedBy", Application.UserName)
    embargoText = "For Internal Use only."

    For i = 1 To subjectCount
        totalEntries = totalEntries + subjects(i).N
        totalPass = totalPass + subjects(i).PassCount
        If subjects(i).N > 0 And subjects(i).PassCount = subjects(i).N Then perfectCount = perfectCount + 1
        If subjects(i).N > 0 And EmailPct(subjects(i).PassCount, subjects(i).N) < 90# Then belowCount = belowCount + 1
    Next i
    AppendHtml html, "<html><body style='margin:0;padding:0;background:#f5f9fd;font-family:Arial,Helvetica,sans-serif;color:#23384d;'>"
    AppendHtml html, "<table role='presentation' width='100%' cellspacing='0' cellpadding='0' border='0' style='background:#f5f9fd;'><tr><td align='center' style='padding:14px;'>"
    AppendHtml html, "<table role='presentation' width='780' cellspacing='0' cellpadding='0' border='0' style='width:100%;max-width:780px;'>"

    AppendHtml html, "<tr><td style='background:#b71c1c;color:#ffffff;font-weight:bold;text-align:center;padding:10px 14px;border:1px solid #9f1818;'>" & HtmlEncode(embargoText) & "</td></tr>"
    AppendHtml html, SpacerRow(10)
    AppendHtml html, "<tr><td style='background:#eef5fb;border:1px solid #d7e4ef;padding:18px 20px;'>" & _
                     "<div style='font-size:24px;line-height:29px;font-weight:bold;color:#1f4e79;'>" & HtmlEncode(levelText & " " & assessmentName & IIf(yearText <> "", " " & yearText, "")) & "</div>" & _
                     "<div style='font-size:14px;color:#385a78;margin-top:5px;'>" & HtmlEncode(schoolName) & "</div>" & _
                     "<div style='font-size:12px;color:#60788e;margin-top:7px;'>Results Summary based on imported Cockpit data.</div></td></tr>"
    AppendHtml html, SpacerRow(10)

    AppendHtml html, "<tr><td>" & BuildKpiGrid(candidateCount, subjectCount, totalEntries, totalPass, perfectCount, belowCount) & "</td></tr>"
    AppendHtml html, SpacerRow(10)
    AppendHtml html, CardStart("Subject Highlights", "")
    AppendHtml html, "<div style='font-size:13px;font-weight:bold;color:#548235;margin:2px 0 7px;'>100% passes</div>"
    AppendHtml html, BuildHighlightTable(subjects, subjectCount, True)
    AppendHtml html, "<div style='font-size:13px;font-weight:bold;color:#c00000;margin:15px 0 7px;'>Pass rate below 90%</div>"
    AppendHtml html, BuildHighlightTable(subjects, subjectCount, False)
    AppendHtml html, CardEnd()

    AppendSchemePerformance html, subjects, subjectCount, "G3"
    AppendSchemePerformance html, subjects, subjectCount, "G2"
    AppendSchemePerformance html, subjects, subjectCount, "G1"

    AppendHtml html, SpacerRow(10)
    AppendHtml html, CardStart("Top Students", "Ranked by distinctions, then passes and student name. Current imported release only.")
    AppendHtml html, BuildTopStudentHtml(students, studentCount, "G3")
    AppendHtml html, BuildTopStudentHtml(students, studentCount, "G2")
    AppendHtml html, BuildTopStudentHtml(students, studentCount, "G1")
    AppendHtml html, CardEnd()

    AppendHtml html, SpacerRow(10)
    AppendHtml html, "<tr><td style='background:#ffffff;border:1px solid #d7e4ef;padding:14px 16px;font-size:11px;line-height:16px;color:#60788e;'>" & _
                     "<div><b>Prepared by:</b> " & HtmlEncode(preparedBy) & "</div>" & _
                     "<div><b>Generated:</b> " & Format$(Now, "dd mmm yyyy, hh:mm AM/PM") & "</div>" & _
                     "<div><b>Source:</b> " & HtmlEncode(sourceSheetName) & "</div>" & _
                     "<div style='margin-top:7px;'>Subject performance and student counts are based on the selected current results release. Grades are evaluated within their native G3, G2 or G1 scheme. This email does not combine earlier sittings.</div>" & _
                     "</td></tr>"

    AppendHtml html, "</table></td></tr></table></body></html>"
    BuildResultsEmailHtml = html
End Function

'------------------------------------------------------------
' ALL-LEVELS MANAGEMENT SUMMARY
'------------------------------------------------------------
Private Sub DraftAllLevelsManagementSummary()
    Dim summaries() As tEmailLevelSummary
    Dim summaryCount As Long
    Dim streamSummaries() As tEmailLevelSummary
    Dim streamSummaryCount As Long
    Dim subjectResults() As tEmailSubjectResult
    Dim subjectResultCount As Long
    Dim config As tEmailManagementConfig
    Dim htmlBody As String
    Dim summaryMode As String

    If gAllLevelSheets Is Nothing Then Exit Sub
    If gAllLevelSheets.count = 0 Then Exit Sub

    ReadEmailManagementConfig config
    BuildLevelSummaries gAllLevelSheets, summaries, summaryCount, subjectResults, subjectResultCount
    If summaryCount = 0 Then
        MsgBox "No eligible level summaries could be built for " & gSelectedExamLabel & ".", vbExclamation
        Exit Sub
    End If

    SortLevelSummaries summaries, summaryCount
    summaryMode = ResolvePrelimManagementSummaryMode(gSelectedExamLabel)
    If summaryMode = "" Then Exit Sub

    If summaryMode = "STREAM" Then
        BuildPrelimStreamSummaries gAllLevelSheets, streamSummaries, streamSummaryCount
        If streamSummaryCount = 0 Then
            MsgBox "No 4EX, 4NA, 4NT or 5NA student groups were found in the matching AtRisk sheets." & vbCrLf & _
                   "Check Settings column D:E and rebuild the AtRisk summaries.", _
                   vbExclamation, "PRELIM Stream Breakdown"
            Exit Sub
        End If
        BuildManagementSummaryWorksheet streamSummaries, streamSummaryCount, subjectResults, subjectResultCount, _
                                        config, gSelectedExamLabel, gSelectedYear, "Stream"
        htmlBody = BuildAllLevelsEmailHtml(streamSummaries, streamSummaryCount, subjectResults, subjectResultCount, _
                                           config, gSelectedExamLabel, gSelectedYear, "Stream")
    Else
        BuildManagementSummaryWorksheet summaries, summaryCount, subjectResults, subjectResultCount, _
                                        config, gSelectedExamLabel, gSelectedYear, "Level"
        htmlBody = BuildAllLevelsEmailHtml(summaries, summaryCount, subjectResults, subjectResultCount, _
                                           config, gSelectedExamLabel, gSelectedYear, "Level")
    End If
    CreateOutlookDraft htmlBody, gSelectedExamLabel, gSelectedYear, "All Levels", True
End Sub

Private Sub BuildLevelSummaries(ByVal sourceSheets As Collection, _
                                ByRef summaries() As tEmailLevelSummary, _
                                ByRef summaryCount As Long, _
                                ByRef subjectResults() As tEmailSubjectResult, _
                                ByRef subjectResultCount As Long)
    Dim ws As Worksheet
    Dim subjects() As tEmailSubject
    Dim subjectCount As Long
    Dim warningText As String
    Dim assessmentName As String, yearText As String, levelText As String
    Dim i As Long

    For i = 1 To sourceSheets.count
        Set ws = sourceSheets(i)
        Erase subjects
        subjectCount = 0: warningText = ""
        assessmentName = "": yearText = "": levelText = ""

        CollectEmailSubjects ws, subjects, subjectCount, warningText
        If subjectCount > 0 Then
            GetSourceLabels ws, assessmentName, yearText, levelText

            summaryCount = summaryCount + 1
            ReDim Preserve summaries(1 To summaryCount)
            With summaries(summaryCount)
                .LevelText = levelText
                .YearText = yearText
                .CandidateCount = CountCandidates(ws)
                .SubjectCount = subjectCount
            End With

            AddSubjectMetricsToLevelSummary summaries(summaryCount), subjects, subjectCount
            AddStudentMetricsFromAtRiskSheet summaries(summaryCount), levelText, assessmentName, yearText
            AddEmailSubjectResults subjectResults, subjectResultCount, levelText, subjects, subjectCount
        End If
    Next i
End Sub

Private Sub AddStudentMetricsFromAtRiskSheet(ByRef summary As tEmailLevelSummary, _
                                             ByVal levelText As String, _
                                             ByVal assessmentName As String, _
                                             ByVal yearText As String)
    Dim ws As Worksheet
    Dim expectedSheetName As String
    Dim classCol As Long, regCol As Long, nameCol As Long
    Dim attemptedCol As Long, passedCol As Long, failedCol As Long, distinctionCol As Long
    Dim lastRow As Long, r As Long
    Dim attemptedCount As Long, passedCount As Long, failedCount As Long, distinctionCount As Long
    Dim classText As String, regText As String, nameText As String, keyText As String
    Dim seen As Object

    expectedSheetName = BuildManagementAtRiskSheetName(levelText, assessmentName, yearText)
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(expectedSheetName)
    On Error GoTo 0
    If ws Is Nothing Then
        Err.Raise vbObjectError + 2101, "AddStudentMetricsFromAtRiskSheet", _
                  "Missing '" & expectedSheetName & "'. Run BuildSec_AtRiskSummary before drafting the management email."
    End If

    classCol = FindEmailHeaderAtRow(ws, 4, "Class")
    regCol = FindEmailHeaderAtRow(ws, 4, "RegNo")
    nameCol = FindEmailHeaderAtRow(ws, 4, "Name")
    attemptedCol = FindEmailHeaderAtRow(ws, 4, "Subjects Attempted")
    passedCol = FindEmailHeaderAtRow(ws, 4, "Subjects Passed")
    failedCol = FindEmailHeaderAtRow(ws, 4, "Subjects Failed")
    distinctionCol = FindEmailHeaderAtRow(ws, 4, "Distinctions")
    If classCol = 0 Or nameCol = 0 Or attemptedCol = 0 Or passedCol = 0 _
       Or failedCol = 0 Or distinctionCol = 0 Then
        Err.Raise vbObjectError + 2102, "AddStudentMetricsFromAtRiskSheet", _
                  "The sheet '" & expectedSheetName & "' does not have the expected AtRisk columns. Rebuild it with BuildSec_AtRiskSummary."
    End If

    Set seen = CreateObject("Scripting.Dictionary")
    seen.CompareMode = vbTextCompare
    lastRow = ws.Cells(ws.Rows.count, nameCol).End(xlUp).Row

    For r = 5 To lastRow
        classText = Trim$(CStr(ws.Cells(r, classCol).value))
        nameText = Trim$(CStr(ws.Cells(r, nameCol).value))
        regText = ""
        If regCol > 0 Then regText = Trim$(CStr(ws.Cells(r, regCol).value))
        If nameText <> "" Then
            If regText <> "" Then
                keyText = classText & "|REG|" & regText
            Else
                keyText = classText & "|NAME|" & nameText
            End If

            attemptedCount = EmailLongValue(ws.Cells(r, attemptedCol).value)
            passedCount = EmailLongValue(ws.Cells(r, passedCol).value)
            failedCount = EmailLongValue(ws.Cells(r, failedCol).value)
            distinctionCount = EmailLongValue(ws.Cells(r, distinctionCol).value)

            ' Retain all-VR/MC students in AtRisk for follow-up, but do
            ' not include 0/0/0 rows in management-email outcomes.
            If attemptedCount > 0 Or passedCount > 0 Or failedCount > 0 Then
                If Not seen.Exists(keyText) Then
                    seen.Add keyText, True
                    AddManagementStudentOutcome summary, failedCount, distinctionCount
                End If
            End If
        End If
    Next r
End Sub

Private Sub AddManagementStudentOutcome(ByRef summary As tEmailLevelSummary, _
                                        ByVal failedCount As Long, _
                                        ByVal distinctionCount As Long)
    summary.ValidStudentCount = summary.ValidStudentCount + 1
    summary.DistinctionStudentCount = summary.DistinctionStudentCount + 1

    Select Case failedCount
        Case 0: summary.PassAllStudentCount = summary.PassAllStudentCount + 1
        Case 1: summary.FailOneStudentCount = summary.FailOneStudentCount + 1
        Case 2: summary.FailTwoStudentCount = summary.FailTwoStudentCount + 1
        Case Else: summary.FailThreePlusStudentCount = summary.FailThreePlusStudentCount + 1
    End Select

    Select Case distinctionCount
        Case 0: summary.NoDistinctionStudentCount = summary.NoDistinctionStudentCount + 1
        Case 1: summary.OneDistinctionStudentCount = summary.OneDistinctionStudentCount + 1
        Case 2: summary.TwoDistinctionStudentCount = summary.TwoDistinctionStudentCount + 1
        Case 3: summary.ThreeDistinctionStudentCount = summary.ThreeDistinctionStudentCount + 1
        Case 4: summary.FourDistinctionStudentCount = summary.FourDistinctionStudentCount + 1
        Case Else: summary.FivePlusDistinctionStudentCount = summary.FivePlusDistinctionStudentCount + 1
    End Select
End Sub

Private Sub BuildPrelimStreamSummaries(ByVal sourceSheets As Collection, _
                                       ByRef summaries() As tEmailLevelSummary, _
                                       ByRef summaryCount As Long)
    Dim sourceWs As Worksheet, atRiskWs As Worksheet
    Dim assessmentName As String, yearText As String, levelText As String
    Dim expectedSheetName As String
    Dim classCol As Long, regCol As Long, nameCol As Long, groupCol As Long
    Dim attemptedCol As Long, passedCol As Long, failedCol As Long, distinctionCol As Long
    Dim lastRow As Long, r As Long, i As Long, streamIndex As Long
    Dim attemptedCount As Long, passedCount As Long, failedCount As Long, distinctionCount As Long
    Dim classText As String, regText As String, nameText As String, keyText As String
    Dim streamCode As String, levelDigit As String
    Dim groupedCount As Long, unmappedCount As Long
    Dim seen As Object

    summaryCount = 0
    If sourceSheets Is Nothing Then Exit Sub

    ReDim summaries(1 To 4)
    summaries(1).LevelText = "4EX"
    summaries(2).LevelText = "4NA"
    summaries(3).LevelText = "4NT"
    summaries(4).LevelText = "5NA"

    Set seen = CreateObject("Scripting.Dictionary")
    seen.CompareMode = vbTextCompare

    For i = 1 To sourceSheets.count
        Set sourceWs = sourceSheets(i)
        assessmentName = "": yearText = "": levelText = ""
        GetSourceLabels sourceWs, assessmentName, yearText, levelText
        levelDigit = FirstLevelDigit(levelText)
        If levelDigit <> "4" And levelDigit <> "5" Then GoTo NextSourceSheet
        expectedSheetName = BuildManagementAtRiskSheetName(levelText, assessmentName, yearText)

        Set atRiskWs = Nothing
        On Error Resume Next
        Set atRiskWs = ThisWorkbook.Worksheets(expectedSheetName)
        On Error GoTo 0
        If atRiskWs Is Nothing Then
            Err.Raise vbObjectError + 2111, "BuildPrelimStreamSummaries", _
                      "Missing '" & expectedSheetName & "'. Run BuildSec_AtRiskSummary before drafting the PRELIM stream email."
        End If

        classCol = FindEmailHeaderAtRow(atRiskWs, 4, "Class")
        regCol = FindEmailHeaderAtRow(atRiskWs, 4, "RegNo")
        nameCol = FindEmailHeaderAtRow(atRiskWs, 4, "Name")
        attemptedCol = FindEmailHeaderAtRow(atRiskWs, 4, "Subjects Attempted")
        passedCol = FindEmailHeaderAtRow(atRiskWs, 4, "Subjects Passed")
        failedCol = FindEmailHeaderAtRow(atRiskWs, 4, "Subjects Failed")
        groupCol = FindEmailHeaderAtRow(atRiskWs, 4, "Group")
        distinctionCol = FindEmailHeaderAtRow(atRiskWs, 4, "Distinctions")
        If classCol = 0 Or nameCol = 0 Or attemptedCol = 0 Or passedCol = 0 _
           Or failedCol = 0 Or groupCol = 0 Or distinctionCol = 0 Then
            Err.Raise vbObjectError + 2112, "BuildPrelimStreamSummaries", _
                      "The sheet '" & expectedSheetName & "' does not have the expected AtRisk columns. Rebuild it with BuildSec_AtRiskSummary."
        End If

        lastRow = atRiskWs.Cells(atRiskWs.Rows.count, nameCol).End(xlUp).Row
        For r = 5 To lastRow
            classText = Trim$(CStr(atRiskWs.Cells(r, classCol).value))
            nameText = Trim$(CStr(atRiskWs.Cells(r, nameCol).value))
            regText = ""
            If regCol > 0 Then regText = Trim$(CStr(atRiskWs.Cells(r, regCol).value))

            attemptedCount = EmailLongValue(atRiskWs.Cells(r, attemptedCol).value)
            passedCount = EmailLongValue(atRiskWs.Cells(r, passedCol).value)
            failedCount = EmailLongValue(atRiskWs.Cells(r, failedCol).value)
            distinctionCount = EmailLongValue(atRiskWs.Cells(r, distinctionCol).value)

            If nameText <> "" And (attemptedCount > 0 Or passedCount > 0 Or failedCount > 0) Then
                ' The PRELIM Sec 5 cohort is 5NA even though its subjects use
                ' the G3 grade scheme and an older AtRisk output may therefore
                ' show G3/EX in the Group column. Subject-level distinction
                ' rules remain unchanged and are already reflected in column P.
                If levelDigit = "5" Then
                    streamIndex = 4
                Else
                    streamCode = NormalizeManagementStream(CStr(atRiskWs.Cells(r, groupCol).value))
                    streamIndex = ManagementStreamIndex(levelDigit, streamCode)
                End If
                If streamIndex = 0 Then
                    unmappedCount = unmappedCount + 1
                Else
                    If regText <> "" Then
                        keyText = atRiskWs.Name & "|" & classText & "|REG|" & regText
                    Else
                        keyText = atRiskWs.Name & "|" & classText & "|NAME|" & nameText
                    End If
                    If Not seen.Exists(keyText) Then
                        seen.Add keyText, True
                        AddManagementStudentOutcome summaries(streamIndex), failedCount, distinctionCount
                        groupedCount = groupedCount + 1
                    End If
                End If
            End If
        Next r
NextSourceSheet:
    Next i

    If unmappedCount > 0 Then
        Err.Raise vbObjectError + 2113, "BuildPrelimStreamSummaries", _
                  CStr(unmappedCount) & " Sec 4 student(s) with valid results do not map to 4EX, 4NA or 4NT." & vbCrLf & _
                  "Complete the Sec 4 class-to-stream mappings in Settings column D:E and rebuild the AtRisk summaries."
    End If

    If groupedCount > 0 Then summaryCount = 4
End Sub

Private Function NormalizeManagementStream(ByVal rawGroup As String) As String
    Dim g As String
    g = UCase$(Trim$(rawGroup))
    g = Replace(g, " ", "")
    g = Replace(g, "-", "")
    g = Replace(g, "(", "")
    g = Replace(g, ")", "")
    g = Replace(g, ".", "")

    Select Case g
        Case "EX", "EXPRESS", "G3": NormalizeManagementStream = "EX"
        Case "NA", "NORMALACADEMIC", "G2": NormalizeManagementStream = "NA"
        Case "NT", "NORMALTECHNICAL", "G1": NormalizeManagementStream = "NT"
    End Select
End Function

Private Function ManagementStreamIndex(ByVal levelDigit As String, _
                                       ByVal streamCode As String) As Long
    Select Case Trim$(levelDigit) & UCase$(Trim$(streamCode))
        Case "4EX": ManagementStreamIndex = 1
        Case "4NA": ManagementStreamIndex = 2
        Case "4NT": ManagementStreamIndex = 3
        Case "5NA": ManagementStreamIndex = 4
    End Select
End Function

Private Function BuildManagementAtRiskSheetName(ByVal levelText As String, _
                                                ByVal assessmentName As String, _
                                                ByVal yearText As String) As String
    Dim levelCode As String, assessmentKey As String, reportYear As String
    Dim maxAssessmentLength As Long
    levelCode = "S" & FirstLevelDigit(levelText)
    assessmentKey = ManagementAtRiskAssessmentKey(assessmentName)
    reportYear = ExtractEmailYear(yearText)
    If reportYear = "" Then reportYear = Trim$(yearText)
    maxAssessmentLength = 31 - Len("AtRisk_") - Len(levelCode) - Len(reportYear) - 2
    If maxAssessmentLength < 1 Then maxAssessmentLength = 1
    BuildManagementAtRiskSheetName = "AtRisk_" & levelCode & "_" & _
                                     Left$(assessmentKey, maxAssessmentLength) & "_" & reportYear
End Function

Private Function ManagementAtRiskAssessmentKey(ByVal assessmentName As String) As String
    Select Case CanonicalExamKey(assessmentName)
        Case "FIRSTCOMBINED": ManagementAtRiskAssessmentKey = "1COMB"
        Case "SECONDCOMBINED": ManagementAtRiskAssessmentKey = "2COMB"
        Case Else: ManagementAtRiskAssessmentKey = CanonicalExamKey(assessmentName)
    End Select
End Function

Private Function EmailLongValue(ByVal valueData As Variant) As Long
    If IsNumeric(valueData) Then
        If CDbl(valueData) > 0# Then EmailLongValue = CLng(valueData)
    End If
End Function

Private Sub AddSubjectMetricsToLevelSummary(ByRef summary As tEmailLevelSummary, _
                                            ByRef subjects() As tEmailSubject, _
                                            ByVal subjectCount As Long)
    Dim i As Long, passPct As Double
    For i = 1 To subjectCount
        summary.ValidEntries = summary.ValidEntries + subjects(i).N
        summary.PassEntries = summary.PassEntries + subjects(i).PassCount
        If subjects(i).N > 0 Then
            passPct = EmailPct(subjects(i).PassCount, subjects(i).N)
            If passPct = 100# Then summary.PerfectSubjects = summary.PerfectSubjects + 1
            If passPct < 90# Then summary.BelowNinetySubjects = summary.BelowNinetySubjects + 1
        End If
    Next i
End Sub

Private Sub AddStudentMetricsToLevelSummary(ByRef summary As tEmailLevelSummary, _
                                            ByRef students() As tEmailStudent, _
                                            ByVal studentCount As Long)
    Dim i As Long
    For i = 1 To studentCount
        If students(i).CountedSubjectCount > 0 Then
            summary.ValidStudentCount = summary.ValidStudentCount + 1
            Select Case students(i).FailCount
                Case 0: summary.PassAllStudentCount = summary.PassAllStudentCount + 1
                Case 1: summary.FailOneStudentCount = summary.FailOneStudentCount + 1
                Case 2: summary.FailTwoStudentCount = summary.FailTwoStudentCount + 1
                Case Else: summary.FailThreePlusStudentCount = summary.FailThreePlusStudentCount + 1
            End Select
        End If
    Next i
End Sub

Private Sub AddEmailSubjectResults(ByRef results() As tEmailSubjectResult, _
                                   ByRef resultCount As Long, ByVal levelText As String, _
                                   ByRef subjects() As tEmailSubject, ByVal subjectCount As Long)
    Dim i As Long
    For i = 1 To subjectCount
        If subjects(i).N > 0 Then
            resultCount = resultCount + 1
            ReDim Preserve results(1 To resultCount)
            With results(resultCount)
                .LevelText = levelText
                .DisplayName = subjects(i).DisplayName
                .Scheme = subjects(i).Scheme
                .N = subjects(i).N
                .PassCount = subjects(i).PassCount
                .DistCount = subjects(i).DistCount
            End With
        End If
    Next i
End Sub

Private Sub UpdateEmailGroupProfile(ByRef profile As tEmailGroupProfile, _
                                    ByVal failCount As Long, ByVal distCount As Long)
    profile.StudentCount = profile.StudentCount + 1
    If failCount = 0 Then
        profile.PassAllCount = profile.PassAllCount + 1
    ElseIf failCount = 1 Then
        profile.FailOneCount = profile.FailOneCount + 1
    ElseIf failCount = 2 Then
        profile.FailTwoCount = profile.FailTwoCount + 1
    Else
        profile.FailThreePlusCount = profile.FailThreePlusCount + 1
    End If
    If distCount >= 1 Then profile.DistOnePlusCount = profile.DistOnePlusCount + 1
    If distCount >= 2 Then profile.DistTwoPlusCount = profile.DistTwoPlusCount + 1
    If distCount >= 3 Then profile.DistThreePlusCount = profile.DistThreePlusCount + 1
End Sub

Private Sub SortLevelSummaries(ByRef summaries() As tEmailLevelSummary, ByVal summaryCount As Long)
    Dim i As Long, j As Long

    For i = 1 To summaryCount - 1
        For j = i + 1 To summaryCount
            If LevelSummaryBefore(summaries(j), summaries(i)) Then
                SwapLevelSummaries summaries(i), summaries(j)
            End If
        Next j
    Next i
End Sub

Private Sub SwapLevelSummaries(ByRef a As tEmailLevelSummary, _
                               ByRef b As tEmailLevelSummary)
    Dim textValue As String
    Dim longValue As Long

    textValue = a.LevelText: a.LevelText = b.LevelText: b.LevelText = textValue
    textValue = a.YearText: a.YearText = b.YearText: b.YearText = textValue
    longValue = a.CandidateCount: a.CandidateCount = b.CandidateCount: b.CandidateCount = longValue
    longValue = a.SubjectCount: a.SubjectCount = b.SubjectCount: b.SubjectCount = longValue
    longValue = a.ValidEntries: a.ValidEntries = b.ValidEntries: b.ValidEntries = longValue
    longValue = a.PassEntries: a.PassEntries = b.PassEntries: b.PassEntries = longValue
    longValue = a.PerfectSubjects: a.PerfectSubjects = b.PerfectSubjects: b.PerfectSubjects = longValue
    longValue = a.BelowNinetySubjects: a.BelowNinetySubjects = b.BelowNinetySubjects: b.BelowNinetySubjects = longValue
    longValue = a.ValidStudentCount: a.ValidStudentCount = b.ValidStudentCount: b.ValidStudentCount = longValue
    longValue = a.PassAllStudentCount: a.PassAllStudentCount = b.PassAllStudentCount: b.PassAllStudentCount = longValue
    longValue = a.FailOneStudentCount: a.FailOneStudentCount = b.FailOneStudentCount: b.FailOneStudentCount = longValue
    longValue = a.FailTwoStudentCount: a.FailTwoStudentCount = b.FailTwoStudentCount: b.FailTwoStudentCount = longValue
    longValue = a.FailThreePlusStudentCount: a.FailThreePlusStudentCount = b.FailThreePlusStudentCount: b.FailThreePlusStudentCount = longValue
    longValue = a.DistinctionStudentCount: a.DistinctionStudentCount = b.DistinctionStudentCount: b.DistinctionStudentCount = longValue
    longValue = a.NoDistinctionStudentCount: a.NoDistinctionStudentCount = b.NoDistinctionStudentCount: b.NoDistinctionStudentCount = longValue
    longValue = a.OneDistinctionStudentCount: a.OneDistinctionStudentCount = b.OneDistinctionStudentCount: b.OneDistinctionStudentCount = longValue
    longValue = a.TwoDistinctionStudentCount: a.TwoDistinctionStudentCount = b.TwoDistinctionStudentCount: b.TwoDistinctionStudentCount = longValue
    longValue = a.ThreeDistinctionStudentCount: a.ThreeDistinctionStudentCount = b.ThreeDistinctionStudentCount: b.ThreeDistinctionStudentCount = longValue
    longValue = a.FourDistinctionStudentCount: a.FourDistinctionStudentCount = b.FourDistinctionStudentCount: b.FourDistinctionStudentCount = longValue
    longValue = a.FivePlusDistinctionStudentCount: a.FivePlusDistinctionStudentCount = b.FivePlusDistinctionStudentCount: b.FivePlusDistinctionStudentCount = longValue
    SwapEmailGroupProfiles a.G3Profile, b.G3Profile
    SwapEmailGroupProfiles a.G2Profile, b.G2Profile
    SwapEmailGroupProfiles a.G1Profile, b.G1Profile
    textValue = a.HighlightsHtml: a.HighlightsHtml = b.HighlightsHtml: b.HighlightsHtml = textValue
End Sub

Private Sub SwapEmailGroupProfiles(ByRef a As tEmailGroupProfile, ByRef b As tEmailGroupProfile)
    Dim longValue As Long
    longValue = a.StudentCount: a.StudentCount = b.StudentCount: b.StudentCount = longValue
    longValue = a.PassAllCount: a.PassAllCount = b.PassAllCount: b.PassAllCount = longValue
    longValue = a.FailOneCount: a.FailOneCount = b.FailOneCount: b.FailOneCount = longValue
    longValue = a.FailTwoCount: a.FailTwoCount = b.FailTwoCount: b.FailTwoCount = longValue
    longValue = a.FailThreePlusCount: a.FailThreePlusCount = b.FailThreePlusCount: b.FailThreePlusCount = longValue
    longValue = a.DistOnePlusCount: a.DistOnePlusCount = b.DistOnePlusCount: b.DistOnePlusCount = longValue
    longValue = a.DistTwoPlusCount: a.DistTwoPlusCount = b.DistTwoPlusCount: b.DistTwoPlusCount = longValue
    longValue = a.DistThreePlusCount: a.DistThreePlusCount = b.DistThreePlusCount: b.DistThreePlusCount = longValue
End Sub

Private Function LevelSummaryBefore(ByRef a As tEmailLevelSummary, _
                                    ByRef b As tEmailLevelSummary) As Boolean
    Dim aLevel As String, bLevel As String
    aLevel = FirstLevelDigit(a.LevelText)
    bLevel = FirstLevelDigit(b.LevelText)
    If aLevel <> bLevel Then
        LevelSummaryBefore = (aLevel < bLevel)
    Else
        LevelSummaryBefore = (StrComp(a.LevelText, b.LevelText, vbTextCompare) < 0)
    End If
End Function

Private Function BuildAllLevelsEmailHtml(ByRef summaries() As tEmailLevelSummary, _
                                         ByVal summaryCount As Long, _
                                         ByRef subjectResults() As tEmailSubjectResult, _
                                         ByVal subjectResultCount As Long, _
                                         ByRef config As tEmailManagementConfig, _
                                         ByVal assessmentName As String, _
                                         ByVal yearText As String, _
                                         ByVal breakdownLabel As String) As String
    Dim html As String
    Dim schoolName As String, embargoText As String

    schoolName = GetEmailSetting("SchoolName", RemoveWorkbookExtension(ThisWorkbook.Name))
    embargoText = "For Internal Use only."

    AppendHtml html, "<html><body style='margin:0;padding:0;background:#f5f9fd;font-family:Arial,Helvetica,sans-serif;color:#23384d;'>"
    AppendHtml html, "<table role='presentation' width='100%' cellspacing='0' cellpadding='0' border='0' style='background:#f5f9fd;'><tr><td align='center' style='padding:14px;'>"
    AppendHtml html, "<table role='presentation' width='900' cellspacing='0' cellpadding='0' border='0' style='width:100%;max-width:900px;'>"
    AppendHtml html, "<tr><td style='background:#b71c1c;color:#ffffff;font-weight:bold;text-align:center;padding:10px 14px;border:1px solid #9f1818;'>" & HtmlEncode(embargoText) & "</td></tr>"
    AppendHtml html, SpacerRow(10)
    AppendHtml html, "<tr><td style='background:#eef5fb;border:1px solid #d7e4ef;padding:18px 20px;'>" & _
                     "<div style='font-size:24px;line-height:29px;font-weight:bold;color:#1f4e79;'>" & HtmlEncode(assessmentName & IIf(yearText <> "", " " & yearText, "") & " - Results Summary") & "</div>" & _
                     "<div style='font-size:14px;color:#385a78;margin-top:5px;'>" & HtmlEncode(schoolName) & "</div>" & _
                     "<div style='font-size:13px;line-height:19px;color:#60788e;margin-top:8px;'>Please find attached the detailed " & HtmlEncode(assessmentName) & " Results Analysis. A summary of the key results is provided below.</div></td></tr>"
    AppendHtml html, SpacerRow(10)

    AppendHtml html, CardStart("1. Student Outcomes by " & breakdownLabel, "")
    AppendHtml html, BuildManagementStudentOutcomesTable(summaries, summaryCount, breakdownLabel)
    AppendHtml html, "<div style='font-size:11px;line-height:16px;color:#60788e;margin-top:9px;font-style:italic;'>The table shows whether students are succeeding across their own subject combination, regardless of whether subjects are taken at G1, G2 or G3. Failure figures are cumulative: a student who failed three subjects appears under Failed &ge;1, &ge;2 and &ge;3.</div>"
    AppendHtml html, CardEnd()

    AppendHtml html, SpacerRow(10)
    AppendHtml html, CardStart("2. Student Distinction Profile by " & breakdownLabel, "")
    AppendHtml html, BuildManagementDistinctionOutcomesTable(summaries, summaryCount, breakdownLabel)
    AppendHtml html, "<div style='font-size:11px;line-height:16px;color:#60788e;margin-top:9px;font-style:italic;'>The table recognises strong attainment within students&rsquo; own subject combinations and does not compare performance across G1, G2 and G3. Figures are cumulative: a student with three distinctions is included under &ge;1, &ge;2 and &ge;3.</div>"
    AppendHtml html, CardEnd()

    AppendHtml html, SpacerRow(10)
    AppendHtml html, CardStart("3. Subject-Level Areas of Concern", "")
    AppendHtml html, BuildManagementConcernTable(subjectResults, subjectResultCount, config)
    AppendHtml html, BuildManagementConcernCriteria(config)
    AppendHtml html, CardEnd()

    AppendHtml html, SpacerRow(10)
    AppendHtml html, CardStart("4. Strong Subject-Level Outcomes", "")
    AppendHtml html, BuildManagementStrongTable(subjectResults, subjectResultCount, config)
    AppendHtml html, BuildManagementStrongCriteria(config)
    AppendHtml html, CardEnd()

    AppendHtml html, SpacerRow(10)
    AppendHtml html, "<tr><td style='background:#ffffff;border:1px solid #d7e4ef;padding:14px 16px;font-size:11px;line-height:16px;color:#60788e;'>" & _
                     "<div style='font-size:13px;line-height:19px;color:#385a78;'>Detailed subject-level results, students at risk and top performers are available in the attached Excel workbook.</div>" & _
                     "<div style='font-size:13px;line-height:19px;color:#385a78;margin-top:8px;'>Thank you.</div>" & _
                     "</td></tr>"

    AppendHtml html, "</table></td></tr></table></body></html>"
    BuildAllLevelsEmailHtml = html
End Function

'------------------------------------------------------------
' MANAGEMENT SUMMARY WORKSHEET
'------------------------------------------------------------
Private Sub BuildManagementSummaryWorksheet(ByRef summaries() As tEmailLevelSummary, _
                                            ByVal summaryCount As Long, _
                                            ByRef subjectResults() As tEmailSubjectResult, _
                                            ByVal subjectResultCount As Long, _
                                            ByRef config As tEmailManagementConfig, _
                                            ByVal assessmentName As String, _
                                            ByVal yearText As String, _
                                            ByVal breakdownLabel As String)
    Dim ws As Worksheet
    Dim rowPtr As Long, headerRow As Long, lastRow As Long
    Dim i As Long
    Dim atLeastOne As Long, atLeastTwo As Long, atLeastThree As Long
    Dim atLeastFour As Long, atLeastFive As Long
    Dim geSymbol As String

    Set ws = GetOrCreateManagementSummarySheet()
    geSymbol = ChrW(&H2265)

    With ws
        .Cells.Font.Name = "Calibri"
        .Cells.Font.Size = 11
        .Columns("A").ColumnWidth = 16
        .Columns("B").ColumnWidth = 33
        .Columns("C:F").ColumnWidth = 20

        .Range("A1:F1").Merge
        .Range("A1").value = assessmentName & IIf(yearText <> "", " " & yearText, "") & " - Results Summary"
        .Range("A1").Font.Size = 20
        .Range("A1").Font.Bold = True
        .Range("A1").Font.Color = RGB(31, 78, 121)
        .Rows(1).RowHeight = 30
    End With

    rowPtr = 3
    WriteSummarySectionTitle ws, rowPtr, "1. Student Outcomes by " & breakdownLabel
    rowPtr = rowPtr + 1
    headerRow = rowPtr
    WriteSummaryHeaderRow ws, rowPtr, Array(breakdownLabel, "Pass All Subjects", _
                          "Failed " & geSymbol & "1 Subject", "Failed " & geSymbol & "2 Subjects", _
                          "Failed " & geSymbol & "3 Subjects")
    rowPtr = rowPtr + 1
    For i = 1 To summaryCount
        If summaries(i).ValidStudentCount > 0 Then
            atLeastThree = summaries(i).FailThreePlusStudentCount
            atLeastTwo = summaries(i).FailTwoStudentCount + atLeastThree
            atLeastOne = summaries(i).FailOneStudentCount + atLeastTwo
            ws.Cells(rowPtr, 1).value = ManagementLevelLabel(summaries(i).LevelText)
            ws.Cells(rowPtr, 2).value = SummaryOutcomeText(summaries(i).PassAllStudentCount, summaries(i).ValidStudentCount)
            ws.Cells(rowPtr, 3).value = SummaryOutcomeText(atLeastOne, summaries(i).ValidStudentCount)
            ws.Cells(rowPtr, 4).value = SummaryOutcomeText(atLeastTwo, summaries(i).ValidStudentCount)
            ws.Cells(rowPtr, 5).value = SummaryOutcomeText(atLeastThree, summaries(i).ValidStudentCount)
            rowPtr = rowPtr + 1
        End If
    Next i
    lastRow = rowPtr - 1
    StyleSummaryDataTable ws, headerRow, lastRow, 5
    StyleSummaryOutcomeColumns ws, headerRow, lastRow, False
    WriteSummaryNote ws, rowPtr, "The table shows whether students are succeeding across their own subject combination, " & _
                     "regardless of whether subjects are taken at G1, G2 or G3. Failure figures are cumulative: " & _
                     "a student who failed three subjects appears under Failed " & geSymbol & "1, " & _
                     geSymbol & "2 and " & geSymbol & "3."
    rowPtr = rowPtr + 2

    WriteSummarySectionTitle ws, rowPtr, "2. Student Distinction Profile by " & breakdownLabel
    rowPtr = rowPtr + 1
    headerRow = rowPtr
    WriteSummaryHeaderRow ws, rowPtr, Array(breakdownLabel, geSymbol & "1 Distinction", _
                          geSymbol & "2 Distinctions", geSymbol & "3 Distinctions", _
                          geSymbol & "4 Distinctions", geSymbol & "5 Distinctions")
    rowPtr = rowPtr + 1
    For i = 1 To summaryCount
        If summaries(i).DistinctionStudentCount > 0 Then
            atLeastFive = summaries(i).FivePlusDistinctionStudentCount
            atLeastFour = summaries(i).FourDistinctionStudentCount + atLeastFive
            atLeastThree = summaries(i).ThreeDistinctionStudentCount + atLeastFour
            atLeastTwo = summaries(i).TwoDistinctionStudentCount + atLeastThree
            atLeastOne = summaries(i).OneDistinctionStudentCount + atLeastTwo
            ws.Cells(rowPtr, 1).value = ManagementLevelLabel(summaries(i).LevelText)
            ws.Cells(rowPtr, 2).value = SummaryOutcomeText(atLeastOne, summaries(i).DistinctionStudentCount)
            ws.Cells(rowPtr, 3).value = SummaryOutcomeText(atLeastTwo, summaries(i).DistinctionStudentCount)
            ws.Cells(rowPtr, 4).value = SummaryOutcomeText(atLeastThree, summaries(i).DistinctionStudentCount)
            ws.Cells(rowPtr, 5).value = SummaryOutcomeText(atLeastFour, summaries(i).DistinctionStudentCount)
            ws.Cells(rowPtr, 6).value = SummaryOutcomeText(atLeastFive, summaries(i).DistinctionStudentCount)
            rowPtr = rowPtr + 1
        End If
    Next i
    lastRow = rowPtr - 1
    StyleSummaryDataTable ws, headerRow, lastRow, 6
    StyleSummaryOutcomeColumns ws, headerRow, lastRow, True
    WriteSummaryNote ws, rowPtr, "The table recognises strong attainment within students" & ChrW$(8217) & _
                     " own subject combinations and does not compare performance across G1, G2 and G3. " & _
                     "Figures are cumulative: a student with three distinctions is included under " & _
                     geSymbol & "1, " & geSymbol & "2 and " & geSymbol & "3."
    rowPtr = rowPtr + 2

    WriteSummarySectionTitle ws, rowPtr, "3. Subject-Level Areas of Concern"
    rowPtr = rowPtr + 1
    rowPtr = WriteSummaryConcernTable(ws, rowPtr, subjectResults, subjectResultCount, config)
    WriteSummaryNote ws, rowPtr, "Criteria: Concern = pass rate below " & FormatThreshold(config.ConcernBelowPct) & _
                     "; Monitor = pass rate from " & FormatThreshold(config.ConcernBelowPct) & " to below " & _
                     FormatThreshold(config.MonitorBelowPct) & ". Subjects with fewer than " & config.MinCandidature & _
                     " students are excluded due to the small candidature."
    rowPtr = rowPtr + 2

    WriteSummarySectionTitle ws, rowPtr, "4. Strong Subject-Level Outcomes"
    rowPtr = rowPtr + 1
    rowPtr = WriteSummaryStrongTable(ws, rowPtr, subjectResults, subjectResultCount, config)
    WriteSummaryNote ws, rowPtr, "Criteria: Subjects are highlighted if the pass rate is at least " & _
                     FormatThreshold(config.StrongPassAtLeastPct) & " and/or the distinction rate is at least " & _
                     FormatThreshold(config.StrongDistAtLeastPct) & ". Groups with fewer than " & _
                     config.MinCandidature & " students are excluded from this summary."
    rowPtr = rowPtr + 1

    AddManagementSummaryHomeButton ws
    AddManagementSummaryDashboardButton

    With ws
        .Rows("1:" & rowPtr).VerticalAlignment = xlCenter
        .Tab.Color = RGB(31, 78, 121)
        .PageSetup.Orientation = xlLandscape
        .PageSetup.Zoom = False
        .PageSetup.FitToPagesWide = 1
        .PageSetup.FitToPagesTall = False
        .Activate
    End With
    ActiveWindow.DisplayGridlines = False
    ActiveWindow.FreezePanes = False
    ws.Range("A5").Select
    ActiveWindow.FreezePanes = True
End Sub

Private Sub AddManagementSummaryHomeButton(ByVal ws As Worksheet)
    Dim wsDashboard As Worksheet
    Dim targetRange As Range
    Dim shp As Shape

    On Error Resume Next
    Set wsDashboard = ThisWorkbook.Worksheets("Dashboard")
    On Error GoTo 0
    If wsDashboard Is Nothing Then Exit Sub

    Set targetRange = ws.Range(SUMMARY_HOME_BUTTON_RANGE)

    On Error Resume Next
    ws.Shapes(SUMMARY_HOME_BUTTON_NAME).Delete
    On Error GoTo 0

    Set shp = ws.Shapes.AddShape( _
        Type:=5, _
        Left:=targetRange.Left, _
        Top:=targetRange.Top + 1, _
        Width:=targetRange.Width, _
        Height:=targetRange.Height - 2)

    With shp
        .Name = SUMMARY_HOME_BUTTON_NAME
        .Placement = xlMoveAndSize
        .Fill.ForeColor.RGB = RGB(217, 234, 211)
        .Fill.Transparency = 0#
        .line.ForeColor.RGB = RGB(106, 168, 79)
        .line.Weight = 1.5
        With .TextFrame2
            .TextRange.text = "Home"
            .TextRange.Font.Name = "Calibri"
            .TextRange.Font.Size = 11
            .TextRange.Font.Bold = True
            .TextRange.Font.Fill.ForeColor.RGB = RGB(39, 78, 19)
            .TextRange.ParagraphFormat.Alignment = msoAlignCenter
            .VerticalAnchor = msoAnchorMiddle
            .MarginLeft = 6
            .MarginRight = 6
            .MarginTop = 3
            .MarginBottom = 3
        End With
    End With

    ws.Hyperlinks.Add Anchor:=shp, Address:="", SubAddress:="'Dashboard'!A1"
End Sub

Private Sub AddManagementSummaryDashboardButton()
    Dim wsDashboard As Worksheet
    Dim targetRange As Range
    Dim shp As Shape

    On Error Resume Next
    Set wsDashboard = ThisWorkbook.Worksheets("Dashboard")
    On Error GoTo 0
    If wsDashboard Is Nothing Then Exit Sub

    Set targetRange = wsDashboard.Range(SUMMARY_DASHBOARD_BUTTON_RANGE)

    On Error Resume Next
    wsDashboard.Shapes(SUMMARY_DASHBOARD_BUTTON_NAME).Delete
    On Error GoTo 0

    Set shp = wsDashboard.Shapes.AddShape( _
        Type:=5, _
        Left:=targetRange.Left, _
        Top:=targetRange.Top + 1, _
        Width:=targetRange.Width, _
        Height:=targetRange.Height - 2)

    With shp
        .Name = SUMMARY_DASHBOARD_BUTTON_NAME
        .Placement = xlMoveAndSize
        .Fill.ForeColor.RGB = RGB(31, 78, 121)
        .Fill.Transparency = 0#
        .line.ForeColor.RGB = RGB(21, 55, 85)
        .line.Weight = 1.5
        With .TextFrame2
            .TextRange.text = "Management Summary"
            .TextRange.Font.Name = "Calibri"
            .TextRange.Font.Size = 12
            .TextRange.Font.Bold = True
            .TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
            .TextRange.ParagraphFormat.Alignment = msoAlignCenter
            .VerticalAnchor = msoAnchorMiddle
            .MarginLeft = 6
            .MarginRight = 6
            .MarginTop = 3
            .MarginBottom = 3
        End With
    End With

    wsDashboard.Hyperlinks.Add Anchor:=shp, Address:="", SubAddress:="'Summary'!A1"
End Sub

Private Sub StyleSummaryOutcomeColumns(ByVal ws As Worksheet, ByVal headerRow As Long, _
                                       ByVal lastRow As Long, ByVal distinctionTable As Boolean)
    If lastRow <= headerRow Then Exit Sub

    If distinctionTable Then
        ws.Range(ws.Cells(headerRow + 1, 2), ws.Cells(lastRow, 2)).Interior.Color = RGB(255, 255, 255)
        ws.Range(ws.Cells(headerRow + 1, 3), ws.Cells(lastRow, 3)).Interior.Color = RGB(238, 245, 251)
        ws.Range(ws.Cells(headerRow + 1, 4), ws.Cells(lastRow, 4)).Interior.Color = RGB(234, 240, 251)
        ws.Range(ws.Cells(headerRow + 1, 5), ws.Cells(lastRow, 5)).Interior.Color = RGB(237, 246, 232)
        ws.Range(ws.Cells(headerRow + 1, 6), ws.Cells(lastRow, 6)).Interior.Color = RGB(226, 240, 217)
        ws.Range(ws.Cells(headerRow + 1, 2), ws.Cells(lastRow, 4)).Font.Color = RGB(31, 78, 121)
        ws.Range(ws.Cells(headerRow + 1, 5), ws.Cells(lastRow, 6)).Font.Color = RGB(47, 107, 47)
    Else
        ws.Range(ws.Cells(headerRow + 1, 2), ws.Cells(lastRow, 2)).Interior.Color = RGB(237, 246, 232)
        ws.Range(ws.Cells(headerRow + 1, 2), ws.Cells(lastRow, 2)).Font.Color = RGB(84, 130, 53)
        ws.Range(ws.Cells(headerRow + 1, 3), ws.Cells(lastRow, 3)).Interior.Color = RGB(255, 251, 237)
        ws.Range(ws.Cells(headerRow + 1, 3), ws.Cells(lastRow, 3)).Font.Color = RGB(138, 100, 16)
        ws.Range(ws.Cells(headerRow + 1, 4), ws.Cells(lastRow, 4)).Interior.Color = RGB(255, 243, 214)
        ws.Range(ws.Cells(headerRow + 1, 4), ws.Cells(lastRow, 4)).Font.Color = RGB(179, 107, 0)
        ws.Range(ws.Cells(headerRow + 1, 5), ws.Cells(lastRow, 5)).Interior.Color = RGB(255, 240, 240)
        ws.Range(ws.Cells(headerRow + 1, 5), ws.Cells(lastRow, 5)).Font.Color = RGB(192, 0, 0)
    End If
End Sub

Private Function GetOrCreateManagementSummarySheet() As Worksheet
    Dim ws As Worksheet, dashboardWs As Worksheet
    Dim k As Long

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("Summary")
    On Error GoTo 0

    If ws Is Nothing Then
        On Error Resume Next
        Set dashboardWs = ThisWorkbook.Worksheets("Dashboard")
        On Error GoTo 0
        If dashboardWs Is Nothing Then
            Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.count))
        Else
            Set ws = ThisWorkbook.Worksheets.Add(After:=dashboardWs)
        End If
        ws.Name = "Summary"
    Else
        ws.Cells.UnMerge
        ws.Cells.Clear
        For k = ws.Shapes.count To 1 Step -1
            ws.Shapes(k).Delete
        Next k
    End If

    Set GetOrCreateManagementSummarySheet = ws
End Function

Private Sub WriteSummarySectionTitle(ByVal ws As Worksheet, ByVal rowNum As Long, ByVal titleText As String)
    With ws.Range(ws.Cells(rowNum, 1), ws.Cells(rowNum, 6))
        .Merge
        .value = titleText
        .Interior.Color = RGB(221, 235, 247)
        .Font.Color = RGB(31, 78, 121)
        .Font.Bold = True
        .Font.Size = 13
        .Borders.LineStyle = xlContinuous
        .Borders.Color = RGB(183, 204, 221)
    End With
    ws.Rows(rowNum).RowHeight = 24
End Sub

Private Sub WriteSummaryHeaderRow(ByVal ws As Worksheet, ByVal rowNum As Long, ByVal headers As Variant)
    Dim i As Long, lastCol As Long
    lastCol = UBound(headers) - LBound(headers) + 1
    For i = LBound(headers) To UBound(headers)
        ws.Cells(rowNum, i - LBound(headers) + 1).value = headers(i)
    Next i
    With ws.Range(ws.Cells(rowNum, 1), ws.Cells(rowNum, lastCol))
        .Interior.Color = RGB(238, 245, 251)
        .Font.Color = RGB(56, 90, 120)
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
        .WrapText = True
    End With
    ws.Rows(rowNum).RowHeight = 32
End Sub

Private Sub StyleSummaryDataTable(ByVal ws As Worksheet, ByVal headerRow As Long, _
                                  ByVal lastRow As Long, ByVal lastCol As Long)
    If lastRow < headerRow Then Exit Sub
    With ws.Range(ws.Cells(headerRow, 1), ws.Cells(lastRow, lastCol))
        .Borders.LineStyle = xlContinuous
        .Borders.Color = RGB(215, 228, 239)
        .Borders.Weight = xlThin
        .VerticalAlignment = xlCenter
    End With
    If lastRow > headerRow Then
        ws.Range(ws.Cells(headerRow + 1, 1), ws.Cells(lastRow, 1)).Font.Bold = True
        ws.Range(ws.Cells(headerRow + 1, 1), ws.Cells(lastRow, 1)).Font.Color = RGB(31, 78, 121)
        If lastCol >= 2 Then
            ws.Range(ws.Cells(headerRow + 1, 2), ws.Cells(lastRow, lastCol)).HorizontalAlignment = xlCenter
        End If
    End If
End Sub

Private Sub WriteSummaryNote(ByVal ws As Worksheet, ByVal rowNum As Long, ByVal noteText As String)
    With ws.Range(ws.Cells(rowNum, 1), ws.Cells(rowNum, 6))
        .Merge
        .value = noteText
        .Font.Italic = True
        .Font.Size = 10
        .Font.Color = RGB(96, 120, 142)
        .WrapText = True
    End With
    ws.Rows(rowNum).RowHeight = 34
End Sub

Private Function SummaryOutcomeText(ByVal studentCount As Long, ByVal denominator As Long) As String
    SummaryOutcomeText = Format$(EmailPct(studentCount, denominator), "0.0") & "% (" & CStr(studentCount) & ")"
End Function

Private Function WriteSummaryConcernTable(ByVal ws As Worksheet, ByVal startRow As Long, _
                                          ByRef results() As tEmailSubjectResult, _
                                          ByVal resultCount As Long, _
                                          ByRef config As tEmailManagementConfig) As Long
    Dim idx() As Long, n As Long, i As Long, j As Long, tmp As Long
    Dim rowPtr As Long, passPct As Double, statusText As String
    Dim statusColor As Long, statusFill As Long

    For i = 1 To resultCount
        If results(i).N >= config.MinCandidature Then
            If EmailPct(results(i).PassCount, results(i).N) < config.MonitorBelowPct Then
                n = n + 1
                ReDim Preserve idx(1 To n)
                idx(n) = i
            End If
        End If
    Next i

    If n = 0 Then
        ws.Range(ws.Cells(startRow, 1), ws.Cells(startRow, 6)).Merge
        ws.Cells(startRow, 1).value = "No subject/G-level groups met the criteria for Areas of Concern."
        ws.Cells(startRow, 1).Font.Color = RGB(96, 120, 142)
        WriteSummaryConcernTable = startRow + 1
        Exit Function
    End If

    For i = 1 To n - 1
        For j = i + 1 To n
            If ManagementConcernBefore(results(idx(j)), results(idx(i))) Then
                tmp = idx(i): idx(i) = idx(j): idx(j) = tmp
            End If
        Next j
    Next i

    WriteSummaryHeaderRow ws, startRow, Array("Level", "Subject / G-Level", "No. Taking", "Pass Rate", "Status")
    rowPtr = startRow + 1
    For i = 1 To n
        j = idx(i)
        passPct = EmailPct(results(j).PassCount, results(j).N)
        If passPct < config.ConcernBelowPct Then
            statusText = "Concern": statusColor = RGB(192, 0, 0): statusFill = RGB(255, 240, 240)
        Else
            statusText = "Monitor": statusColor = RGB(138, 100, 16): statusFill = RGB(255, 243, 214)
        End If
        ws.Cells(rowPtr, 1).value = ManagementLevelLabel(results(j).LevelText)
        ws.Cells(rowPtr, 2).value = results(j).DisplayName & " " & results(j).Scheme
        ws.Cells(rowPtr, 3).value = results(j).N
        ws.Cells(rowPtr, 4).value = passPct / 100#
        ws.Cells(rowPtr, 4).NumberFormat = "0.0%"
        ws.Cells(rowPtr, 5).value = statusText
        ws.Range(ws.Cells(rowPtr, 4), ws.Cells(rowPtr, 5)).Font.Color = statusColor
        ws.Range(ws.Cells(rowPtr, 4), ws.Cells(rowPtr, 5)).Interior.Color = statusFill
        rowPtr = rowPtr + 1
    Next i
    StyleSummaryDataTable ws, startRow, rowPtr - 1, 5
    ws.Range(ws.Cells(startRow + 1, 2), ws.Cells(rowPtr - 1, 2)).HorizontalAlignment = xlLeft
    WriteSummaryConcernTable = rowPtr
End Function

Private Function WriteSummaryStrongTable(ByVal ws As Worksheet, ByVal startRow As Long, _
                                         ByRef results() As tEmailSubjectResult, _
                                         ByVal resultCount As Long, _
                                         ByRef config As tEmailManagementConfig) As Long
    Dim idx() As Long, n As Long, i As Long, j As Long, tmp As Long
    Dim rowPtr As Long, passPct As Double, distPct As Double

    For i = 1 To resultCount
        If results(i).N >= config.MinCandidature Then
            passPct = EmailPct(results(i).PassCount, results(i).N)
            distPct = EmailPct(results(i).DistCount, results(i).N)
            If passPct >= config.StrongPassAtLeastPct Or distPct >= config.StrongDistAtLeastPct Then
                n = n + 1
                ReDim Preserve idx(1 To n)
                idx(n) = i
            End If
        End If
    Next i

    If n = 0 Then
        ws.Range(ws.Cells(startRow, 1), ws.Cells(startRow, 6)).Merge
        ws.Cells(startRow, 1).value = "No subject/G-level groups met the criteria for Strong Subject-Level Outcomes."
        ws.Cells(startRow, 1).Font.Color = RGB(96, 120, 142)
        WriteSummaryStrongTable = startRow + 1
        Exit Function
    End If

    For i = 1 To n - 1
        For j = i + 1 To n
            If ManagementStrongBefore(results(idx(j)), results(idx(i))) Then
                tmp = idx(i): idx(i) = idx(j): idx(j) = tmp
            End If
        Next j
    Next i

    WriteSummaryHeaderRow ws, startRow, Array("Level", "Subject / G-Level", "No. Taking", "Pass Rate", "Distinction Rate")
    rowPtr = startRow + 1
    For i = 1 To n
        j = idx(i)
        passPct = EmailPct(results(j).PassCount, results(j).N)
        distPct = EmailPct(results(j).DistCount, results(j).N)
        ws.Cells(rowPtr, 1).value = ManagementLevelLabel(results(j).LevelText)
        ws.Cells(rowPtr, 2).value = results(j).DisplayName & " " & results(j).Scheme
        ws.Cells(rowPtr, 3).value = results(j).N
        ws.Cells(rowPtr, 4).value = passPct / 100#
        ws.Cells(rowPtr, 5).value = distPct / 100#
        ws.Range(ws.Cells(rowPtr, 4), ws.Cells(rowPtr, 5)).NumberFormat = "0.0%"
        ws.Range(ws.Cells(rowPtr, 4), ws.Cells(rowPtr, 5)).Font.Color = RGB(84, 130, 53)
        ws.Range(ws.Cells(rowPtr, 4), ws.Cells(rowPtr, 5)).Interior.Color = RGB(237, 246, 232)
        rowPtr = rowPtr + 1
    Next i
    StyleSummaryDataTable ws, startRow, rowPtr - 1, 5
    ws.Range(ws.Cells(startRow + 1, 2), ws.Cells(rowPtr - 1, 2)).HorizontalAlignment = xlLeft
    WriteSummaryStrongTable = rowPtr
End Function

Private Function BuildManagementStudentOutcomesTable(ByRef summaries() As tEmailLevelSummary, _
                                                     ByVal summaryCount As Long, _
                                                     ByVal breakdownLabel As String) As String
    Dim html As String, i As Long
    Dim atLeastOneFail As Long, atLeastTwoFails As Long, atLeastThreeFails As Long

    html = ManagementTableStart() & "<tr style='background:#eef5fb;'>" & HeaderTd(breakdownLabel) & _
           CenterHeaderTd("Pass All Subjects") & CenterHeaderTdHtml("Failed &ge;1 Subject") & _
           CenterHeaderTdHtml("Failed &ge;2 Subjects") & CenterHeaderTdHtml("Failed &ge;3 Subjects") & "</tr>"

    For i = 1 To summaryCount
        If summaries(i).ValidStudentCount > 0 Then
            atLeastThreeFails = summaries(i).FailThreePlusStudentCount
            atLeastTwoFails = summaries(i).FailTwoStudentCount + atLeastThreeFails
            atLeastOneFail = summaries(i).FailOneStudentCount + atLeastTwoFails

            html = html & "<tr>" & TextTd(ManagementLevelLabel(summaries(i).LevelText), "#1f4e79") & _
                   OutcomeTd(summaries(i).PassAllStudentCount, summaries(i).ValidStudentCount, "#548235", "#edf6e8") & _
                   OutcomeTd(atLeastOneFail, summaries(i).ValidStudentCount, "#8a6410", "#fffbed") & _
                   OutcomeTd(atLeastTwoFails, summaries(i).ValidStudentCount, "#b36b00", "#fff3d6") & _
                   OutcomeTd(atLeastThreeFails, summaries(i).ValidStudentCount, "#c00000", "#fff0f0") & "</tr>"
        End If
    Next i

    BuildManagementStudentOutcomesTable = html & "</table>"
End Function

Private Function BuildManagementDistinctionOutcomesTable(ByRef summaries() As tEmailLevelSummary, _
                                                          ByVal summaryCount As Long, _
                                                          ByVal breakdownLabel As String) As String
    Dim html As String, i As Long
    Dim atLeastOne As Long, atLeastTwo As Long, atLeastThree As Long
    Dim atLeastFour As Long, atLeastFive As Long

    html = ManagementTableStart() & "<tr style='background:#eef5fb;'>" & HeaderTd(breakdownLabel) & _
           CenterHeaderTdHtml("&ge;1 Distinction") & CenterHeaderTdHtml("&ge;2 Distinctions") & _
           CenterHeaderTdHtml("&ge;3 Distinctions") & CenterHeaderTdHtml("&ge;4 Distinctions") & _
           CenterHeaderTdHtml("&ge;5 Distinctions") & "</tr>"

    For i = 1 To summaryCount
        If summaries(i).DistinctionStudentCount > 0 Then
            atLeastFive = summaries(i).FivePlusDistinctionStudentCount
            atLeastFour = summaries(i).FourDistinctionStudentCount + atLeastFive
            atLeastThree = summaries(i).ThreeDistinctionStudentCount + atLeastFour
            atLeastTwo = summaries(i).TwoDistinctionStudentCount + atLeastThree
            atLeastOne = summaries(i).OneDistinctionStudentCount + atLeastTwo

            html = html & "<tr>" & TextTd(ManagementLevelLabel(summaries(i).LevelText), "#1f4e79") & _
                   OutcomeTd(atLeastOne, summaries(i).DistinctionStudentCount, "#385a78", "#ffffff") & _
                   OutcomeTd(atLeastTwo, summaries(i).DistinctionStudentCount, "#1f4e79", "#eef5fb") & _
                   OutcomeTd(atLeastThree, summaries(i).DistinctionStudentCount, "#4472c4", "#eaf0fb") & _
                   OutcomeTd(atLeastFour, summaries(i).DistinctionStudentCount, "#548235", "#edf6e8") & _
                   OutcomeTd(atLeastFive, summaries(i).DistinctionStudentCount, "#2f6b2f", "#e2f0d9") & "</tr>"
        End If
    Next i

    BuildManagementDistinctionOutcomesTable = html & "</table>"
End Function

Private Function OutcomeTd(ByVal studentCount As Long, ByVal denominator As Long, _
                           ByVal colorText As String, ByVal bgColor As String) As String
    OutcomeTd = CenterNumTd(Format$(EmailPct(studentCount, denominator), "0.0") & "% (" & _
                             CStr(studentCount) & ")", colorText, bgColor)
End Function

Private Function BuildManagementConcernTable(ByRef results() As tEmailSubjectResult, _
                                             ByVal resultCount As Long, _
                                             ByRef config As tEmailManagementConfig) As String
    Dim idx() As Long, n As Long, i As Long, j As Long, tmp As Long
    Dim html As String, passPct As Double, statusText As String
    Dim statusColor As String, statusBg As String

    For i = 1 To resultCount
        If results(i).N >= config.MinCandidature Then
            If EmailPct(results(i).PassCount, results(i).N) < config.MonitorBelowPct Then
                n = n + 1
                ReDim Preserve idx(1 To n)
                idx(n) = i
            End If
        End If
    Next i

    If n = 0 Then
        BuildManagementConcernTable = ManagementEmptyStatement("No subject/G-level groups met the criteria for Areas of Concern.")
        Exit Function
    End If

    For i = 1 To n - 1
        For j = i + 1 To n
            If ManagementConcernBefore(results(idx(j)), results(idx(i))) Then
                tmp = idx(i): idx(i) = idx(j): idx(j) = tmp
            End If
        Next j
    Next i

    html = ManagementTableStart() & "<tr style='background:#eef5fb;'>" & HeaderTd("Level") & HeaderTd("Subject / G-Level") & _
           CenterHeaderTd("No. Taking") & CenterHeaderTd("Pass Rate") & CenterHeaderTd("Status") & "</tr>"
    For i = 1 To n
        j = idx(i)
        passPct = EmailPct(results(j).PassCount, results(j).N)
        If passPct < config.ConcernBelowPct Then
            statusText = "Concern": statusColor = "#c00000": statusBg = "#fff0f0"
        Else
            statusText = "Monitor": statusColor = "#8a6410": statusBg = "#fff3d6"
        End If
        html = html & "<tr>" & TextTd(ManagementLevelLabel(results(j).LevelText), "#1f4e79") & _
               TextTd(results(j).DisplayName & " " & results(j).Scheme, "#23384d") & _
               CenterNumTd(CStr(results(j).N), "#23384d", "#ffffff") & _
               CenterNumTd(Format$(passPct, "0.0") & "%", statusColor, statusBg) & _
               CenterNumTd(statusText, statusColor, statusBg) & "</tr>"
    Next i
    BuildManagementConcernTable = html & "</table>"
End Function

Private Function ManagementConcernBefore(ByRef a As tEmailSubjectResult, _
                                         ByRef b As tEmailSubjectResult) As Boolean
    Dim ap As Double, bp As Double, al As String, bl As String
    ap = EmailPct(a.PassCount, a.N): bp = EmailPct(b.PassCount, b.N)
    If ap <> bp Then ManagementConcernBefore = (ap < bp): Exit Function
    al = FirstLevelDigit(a.LevelText): bl = FirstLevelDigit(b.LevelText)
    If al <> bl Then ManagementConcernBefore = (al < bl): Exit Function
    If StrComp(a.DisplayName, b.DisplayName, vbTextCompare) <> 0 Then
        ManagementConcernBefore = (StrComp(a.DisplayName, b.DisplayName, vbTextCompare) < 0)
        Exit Function
    End If
    ManagementConcernBefore = (StrComp(a.Scheme, b.Scheme, vbTextCompare) < 0)
End Function

Private Function BuildManagementStrongTable(ByRef results() As tEmailSubjectResult, _
                                            ByVal resultCount As Long, _
                                            ByRef config As tEmailManagementConfig) As String
    Dim idx() As Long, n As Long, i As Long, j As Long, tmp As Long
    Dim html As String, passPct As Double, distPct As Double

    For i = 1 To resultCount
        If results(i).N >= config.MinCandidature Then
            passPct = EmailPct(results(i).PassCount, results(i).N)
            distPct = EmailPct(results(i).DistCount, results(i).N)
            If passPct >= config.StrongPassAtLeastPct Or distPct >= config.StrongDistAtLeastPct Then
                n = n + 1
                ReDim Preserve idx(1 To n)
                idx(n) = i
            End If
        End If
    Next i

    If n = 0 Then
        BuildManagementStrongTable = ManagementEmptyStatement("No subject/G-level groups met the criteria for Strong Subject-Level Outcomes.")
        Exit Function
    End If

    For i = 1 To n - 1
        For j = i + 1 To n
            If ManagementStrongBefore(results(idx(j)), results(idx(i))) Then
                tmp = idx(i): idx(i) = idx(j): idx(j) = tmp
            End If
        Next j
    Next i

    html = ManagementTableStart() & "<tr style='background:#eef5fb;'>" & HeaderTd("Level") & HeaderTd("Subject / G-Level") & _
           CenterHeaderTd("No. Taking") & CenterHeaderTd("Pass Rate") & CenterHeaderTd("Distinction Rate") & "</tr>"
    For i = 1 To n
        j = idx(i)
        passPct = EmailPct(results(j).PassCount, results(j).N)
        distPct = EmailPct(results(j).DistCount, results(j).N)
        html = html & "<tr>" & TextTd(ManagementLevelLabel(results(j).LevelText), "#1f4e79") & _
               TextTd(results(j).DisplayName & " " & results(j).Scheme, "#23384d") & _
               CenterNumTd(CStr(results(j).N), "#23384d", "#ffffff") & _
               CenterNumTd(Format$(passPct, "0.0") & "%", "#548235", "#edf6e8") & _
               CenterNumTd(Format$(distPct, "0.0") & "%", "#548235", "#edf6e8") & "</tr>"
    Next i
    BuildManagementStrongTable = html & "</table>"
End Function

Private Function ManagementStrongBefore(ByRef a As tEmailSubjectResult, _
                                        ByRef b As tEmailSubjectResult) As Boolean
    Dim ad As Double, bd As Double, ap As Double, bp As Double
    Dim al As String, bl As String
    ad = EmailPct(a.DistCount, a.N): bd = EmailPct(b.DistCount, b.N)
    If ad <> bd Then ManagementStrongBefore = (ad > bd): Exit Function
    ap = EmailPct(a.PassCount, a.N): bp = EmailPct(b.PassCount, b.N)
    If ap <> bp Then ManagementStrongBefore = (ap > bp): Exit Function
    al = FirstLevelDigit(a.LevelText): bl = FirstLevelDigit(b.LevelText)
    If al <> bl Then ManagementStrongBefore = (al < bl): Exit Function
    If StrComp(a.DisplayName, b.DisplayName, vbTextCompare) <> 0 Then
        ManagementStrongBefore = (StrComp(a.DisplayName, b.DisplayName, vbTextCompare) < 0)
        Exit Function
    End If
    ManagementStrongBefore = (StrComp(a.Scheme, b.Scheme, vbTextCompare) < 0)
End Function

Private Function BuildManagementConcernCriteria(ByRef config As tEmailManagementConfig) As String
    BuildManagementConcernCriteria = "<div style='font-size:11px;line-height:16px;color:#60788e;margin-top:9px;font-style:italic;'>Criteria: Concern = pass rate below " & _
        FormatThreshold(config.ConcernBelowPct) & "; Monitor = pass rate from " & FormatThreshold(config.ConcernBelowPct) & _
        " to below " & FormatThreshold(config.MonitorBelowPct) & ". Subjects with fewer than " & _
        config.MinCandidature & " students are excluded due to the small candidature. Full results remain available in the attached Excel workbook.</div>"
End Function

Private Function BuildManagementStrongCriteria(ByRef config As tEmailManagementConfig) As String
    BuildManagementStrongCriteria = "<div style='font-size:11px;line-height:16px;color:#60788e;margin-top:9px;font-style:italic;'>Criteria: Subjects are highlighted if the pass rate is at least " & _
        FormatThreshold(config.StrongPassAtLeastPct) & " and/or the distinction rate is at least " & _
        FormatThreshold(config.StrongDistAtLeastPct) & ". Groups with fewer than " & config.MinCandidature & _
        " students are excluded from this summary.</div>"
End Function

Private Function FormatThreshold(ByVal valuePct As Double) As String
    FormatThreshold = Format$(valuePct, "0.#") & "%"
End Function

Private Function ManagementEmptyStatement(ByVal valueText As String) As String
    ManagementEmptyStatement = "<div style='font-size:13px;line-height:19px;color:#60788e;padding:5px 0;'>" & _
                               HtmlEncode(valueText) & "</div>"
End Function

Private Function BuildStudentPerformanceProfileTable(ByRef summaries() As tEmailLevelSummary, _
                                                     ByVal summaryCount As Long) As String
    Dim html As String, i As Long
    html = TableStart() & "<tr style='background:#eef5fb;'>" & HeaderTd("Level") & HeaderTd("Scheme") & _
           HeaderTd("Pass All Subjects") & HeaderTd("Fail 1 Subject") & _
           HeaderTd("Fail 2 Subjects") & HeaderTd("Fail 3 or More Subjects") & "</tr>"

    For i = 1 To summaryCount
        AppendStudentProfileRow html, ManagementLevelLabel(summaries(i).LevelText), "G3", summaries(i).G3Profile
        AppendStudentProfileRow html, ManagementLevelLabel(summaries(i).LevelText), "G2", summaries(i).G2Profile
        AppendStudentProfileRow html, ManagementLevelLabel(summaries(i).LevelText), "G1", summaries(i).G1Profile
    Next i

    BuildStudentPerformanceProfileTable = html & "</table>"
End Function

Private Sub AppendStudentProfileRow(ByRef html As String, ByVal levelLabel As String, _
                                    ByVal scheme As String, ByRef profile As tEmailGroupProfile)
    If profile.StudentCount = 0 Then Exit Sub
    html = html & "<tr>" & TextTd(levelLabel, "#1f4e79") & TextTd(scheme, "#385a78") & _
           ProfileTd(profile.PassAllCount, profile.StudentCount, "#548235", "#edf6e8") & _
           ProfileTd(profile.FailOneCount, profile.StudentCount, "#385a78", "#ffffff") & _
           ProfileTd(profile.FailTwoCount, profile.StudentCount, "#b36b00", "#fff8e8") & _
           ProfileTd(profile.FailThreePlusCount, profile.StudentCount, "#c00000", "#fff0f0") & "</tr>"
End Sub

Private Function BuildDistinctionProfileTable(ByRef summaries() As tEmailLevelSummary, _
                                              ByVal summaryCount As Long) As String
    Dim html As String, i As Long
    html = TableStart() & "<tr style='background:#eef5fb;'>" & HeaderTd("Level") & HeaderTd("Scheme") & _
           HeaderTd("At Least 1 Distinction") & HeaderTd("At Least 2 Distinctions") & _
           HeaderTd("At Least 3 Distinctions") & "</tr>"

    For i = 1 To summaryCount
        AppendDistinctionProfileRow html, ManagementLevelLabel(summaries(i).LevelText), "G3", summaries(i).G3Profile
        AppendDistinctionProfileRow html, ManagementLevelLabel(summaries(i).LevelText), "G2", summaries(i).G2Profile
        AppendDistinctionProfileRow html, ManagementLevelLabel(summaries(i).LevelText), "G1", summaries(i).G1Profile
    Next i

    BuildDistinctionProfileTable = html & "</table>"
End Function

Private Sub AppendDistinctionProfileRow(ByRef html As String, ByVal levelLabel As String, _
                                        ByVal scheme As String, ByRef profile As tEmailGroupProfile)
    If profile.StudentCount = 0 Then Exit Sub
    html = html & "<tr>" & TextTd(levelLabel, "#1f4e79") & TextTd(scheme, "#385a78") & _
           ProfileTd(profile.DistOnePlusCount, profile.StudentCount, "#385a78", "#ffffff") & _
           ProfileTd(profile.DistTwoPlusCount, profile.StudentCount, "#1f4e79", "#eef5fb") & _
           ProfileTd(profile.DistThreePlusCount, profile.StudentCount, "#548235", "#edf6e8") & "</tr>"
End Sub

Private Function ProfileTd(ByVal studentCount As Long, ByVal denominator As Long, _
                           ByVal colorText As String, ByVal bgColor As String) As String
    ProfileTd = NumTd(Format$(EmailPct(studentCount, denominator), "0.0") & "% (" & _
                       CStr(studentCount) & ")", colorText, bgColor)
End Function

Private Function ManagementLevelLabel(ByVal levelText As String) As String
    Dim levelDigit As String
    Select Case UCase$(Trim$(levelText))
        Case "4EX", "4NA", "4NT", "5NA"
            ManagementLevelLabel = UCase$(Trim$(levelText))
            Exit Function
    End Select

    levelDigit = FirstLevelDigit(levelText)
    If levelDigit <> "" Then
        ManagementLevelLabel = "Sec " & levelDigit
    Else
        ManagementLevelLabel = levelText
    End If
End Function

Private Function BuildLevelManagementHighlights(ByRef subjects() As tEmailSubject, _
                                                ByVal subjectCount As Long, _
                                                ByRef summary As tEmailLevelSummary) As String
    Dim bestPassIndex As Long, bestDistIndex As Long, worstPassIndex As Long
    Dim i As Long, passPct As Double, distPct As Double
    Dim bestPassPct As Double, bestDistPct As Double, worstPassPct As Double
    Dim html As String, detailText As String

    worstPassPct = 101#
    For i = 1 To subjectCount
        If subjects(i).N > 0 Then
            passPct = EmailPct(subjects(i).PassCount, subjects(i).N)
            distPct = EmailPct(subjects(i).DistCount, subjects(i).N)
            If bestPassIndex = 0 Or passPct > bestPassPct Then
                bestPassIndex = i: bestPassPct = passPct
            End If
            If bestDistIndex = 0 Or distPct > bestDistPct Then
                bestDistIndex = i: bestDistPct = distPct
            End If
            If worstPassIndex = 0 Or passPct < worstPassPct Then
                worstPassIndex = i: worstPassPct = passPct
            End If
        End If
    Next i

    If bestPassIndex > 0 Then
        detailText = subjects(bestPassIndex).DisplayName & " (" & subjects(bestPassIndex).Scheme & ") recorded " & _
                     Format$(bestPassPct, "0.0") & "% pass, N=" & subjects(bestPassIndex).N & "."
        If summary.PerfectSubjects > 0 Then detailText = detailText & " " & summary.PerfectSubjects & " subject(s) achieved 100% pass."
        html = html & ManagementBullet(detailText, "#548235")
    End If

    If bestDistIndex > 0 And subjects(bestDistIndex).DistCount > 0 Then
        html = html & ManagementBullet(subjects(bestDistIndex).DisplayName & " (" & subjects(bestDistIndex).Scheme & _
               ") recorded " & Format$(bestDistPct, "0.0") & "% distinctions, N=" & subjects(bestDistIndex).N & ".", "#385a78")
    End If

    If worstPassIndex > 0 And worstPassPct < 90# Then
        html = html & ManagementBullet(subjects(worstPassIndex).DisplayName & " (" & subjects(worstPassIndex).Scheme & _
               ") requires monitoring: " & Format$(100# - worstPassPct, "0.0") & "% failed, N=" & subjects(worstPassIndex).N & ".", "#c00000")
    End If

    If LevelFailThreePlusCount(summary) > 0 Then
        html = html & ManagementBullet("Students failing three or more subjects require follow-up: " & _
               LevelFailThreeProfileText(summary) & ".", "#c00000")
    End If

    If html = "" Then html = ManagementBullet("No significant exception was identified from the available valid results.", "#60788e")
    BuildLevelManagementHighlights = html
End Function

Private Function BuildManagementTakeaways(ByRef summaries() As tEmailLevelSummary, _
                                          ByVal summaryCount As Long) As String
    Dim i As Long, totalFailThreePlus As Long, totalDistThreePlus As Long
    Dim totalPerfect As Long, totalBelowNinety As Long, html As String
    Dim failG3 As Long, failG2 As Long, failG1 As Long
    Dim distG3 As Long, distG2 As Long, distG1 As Long

    For i = 1 To summaryCount
        failG3 = failG3 + summaries(i).G3Profile.FailThreePlusCount
        failG2 = failG2 + summaries(i).G2Profile.FailThreePlusCount
        failG1 = failG1 + summaries(i).G1Profile.FailThreePlusCount
        distG3 = distG3 + summaries(i).G3Profile.DistThreePlusCount
        distG2 = distG2 + summaries(i).G2Profile.DistThreePlusCount
        distG1 = distG1 + summaries(i).G1Profile.DistThreePlusCount
        totalPerfect = totalPerfect + summaries(i).PerfectSubjects
        totalBelowNinety = totalBelowNinety + summaries(i).BelowNinetySubjects
    Next i
    totalFailThreePlus = failG3 + failG2 + failG1
    totalDistThreePlus = distG3 + distG2 + distG1

    If totalFailThreePlus > 0 Then
        html = html & ManagementBullet(totalFailThreePlus & " student(s) failed three or more subjects (" & _
               SchemeCountText(failG3, failG2, failG1) & "); School Leaders should review the named at-risk list in the workbook and coordinate follow-up.", "#c00000")
    Else
        html = html & ManagementBullet("No student with valid results failed three or more subjects.", "#548235")
    End If

    If totalBelowNinety > 0 Then
        html = html & ManagementBullet(totalBelowNinety & " level-subject result(s) recorded below 90% pass and should be reviewed by the relevant HODs.", "#c00000")
    Else
        html = html & ManagementBullet("No level-subject result recorded below 90% pass.", "#548235")
    End If

    If totalPerfect > 0 Then
        html = html & ManagementBullet(totalPerfect & " level-subject result(s) achieved 100% pass and merit recognition.", "#548235")
    Else
        html = html & ManagementBullet("No level-subject result achieved 100% pass; subject-level strengths remain detailed in the workbook.", "#60788e")
    End If

    If totalDistThreePlus > 0 Then
        html = html & ManagementBullet(totalDistThreePlus & " student(s) attained at least three distinctions (" & _
               SchemeCountText(distG3, distG2, distG1) & "); top-performer details are available in the workbook.", "#1f4e79")
    Else
        html = html & ManagementBullet("No student attained at least three distinctions in the available valid results.", "#60788e")
    End If

    BuildManagementTakeaways = html
End Function

Private Function LevelFailThreePlusCount(ByRef summary As tEmailLevelSummary) As Long
    LevelFailThreePlusCount = summary.G3Profile.FailThreePlusCount + _
                              summary.G2Profile.FailThreePlusCount + _
                              summary.G1Profile.FailThreePlusCount
End Function

Private Function LevelFailThreeProfileText(ByRef summary As tEmailLevelSummary) As String
    Dim valueText As String
    If summary.G3Profile.StudentCount > 0 Then
        valueText = valueText & "G3 " & ProfileText(summary.G3Profile.FailThreePlusCount, summary.G3Profile.StudentCount)
    End If
    If summary.G2Profile.StudentCount > 0 Then
        If valueText <> "" Then valueText = valueText & "; "
        valueText = valueText & "G2 " & ProfileText(summary.G2Profile.FailThreePlusCount, summary.G2Profile.StudentCount)
    End If
    If summary.G1Profile.StudentCount > 0 Then
        If valueText <> "" Then valueText = valueText & "; "
        valueText = valueText & "G1 " & ProfileText(summary.G1Profile.FailThreePlusCount, summary.G1Profile.StudentCount)
    End If
    LevelFailThreeProfileText = valueText
End Function

Private Function ProfileText(ByVal studentCount As Long, ByVal denominator As Long) As String
    ProfileText = Format$(EmailPct(studentCount, denominator), "0.0") & "% (" & studentCount & ")"
End Function

Private Function SchemeCountText(ByVal g3Count As Long, ByVal g2Count As Long, _
                                 ByVal g1Count As Long) As String
    SchemeCountText = "G3: " & g3Count & ", G2: " & g2Count & ", G1: " & g1Count
End Function

Private Function ManagementBullet(ByVal valueText As String, ByVal colorText As String) As String
    ManagementBullet = "<div style='font-size:11px;line-height:17px;color:" & colorText & ";margin:3px 0;'>&bull; " & _
                       HtmlEncode(valueText) & "</div>"
End Function

Private Function BuildKpiGrid(ByVal candidates As Long, ByVal subjectCount As Long, _
                              ByVal totalEntries As Long, ByVal totalPass As Long, _
                              ByVal perfectCount As Long, ByVal belowCount As Long) As String
    Dim overallPass As String
    overallPass = IIf(totalEntries > 0, Format$(EmailPct(totalPass, totalEntries), "0.0") & "%", "-")

    BuildKpiGrid = "<table role='presentation' width='100%' cellspacing='0' cellpadding='0' border='0'>" & _
        "<tr>" & KpiCell("Candidates", CStr(candidates), "#eef5fb", "#1f4e79", False) & _
        KpiGap() & KpiCell("Subjects", CStr(subjectCount), "#f5f5f5", "#1f4e79", False) & _
        KpiGap() & KpiCell("Overall pass", overallPass, "#e2f0d9", "#548235", False) & "</tr>" & _
        "<tr><td colspan='5' height='8' style='height:8px;'></td></tr>" & _
        "<tr>" & KpiCell("100% pass", CStr(perfectCount), "#e2f0d9", "#548235", False) & _
        KpiGap() & KpiCell("Below 90%", CStr(belowCount), "#fff0f0", "#c00000", False) & _
        KpiGap() & KpiCell("Valid subject results", CStr(totalEntries), "#f5f5f5", "#1f4e79", False) & "</tr></table>"
End Function

Private Function KpiCell(ByVal labelText As String, ByVal valueText As String, _
                         ByVal bgColor As String, ByVal valueColor As String, _
                         ByVal unusedFlag As Boolean) As String
    KpiCell = "<td width='32%' valign='top' style='background:" & bgColor & ";border:1px solid #d7e4ef;padding:11px 13px;'>" & _
              "<div style='font-size:10px;font-weight:bold;letter-spacing:.5px;text-transform:uppercase;color:#60788e;'>" & HtmlEncode(labelText) & "</div>" & _
              "<div style='font-size:23px;line-height:27px;font-weight:bold;color:" & valueColor & ";margin-top:3px;'>" & HtmlEncode(valueText) & "</div></td>"
End Function

Private Function KpiGap() As String
    KpiGap = "<td width='2%'>&nbsp;</td>"
End Function

Private Sub AppendSchemePerformance(ByRef html As String, _
                                    ByRef subjects() As tEmailSubject, ByVal subjectCount As Long, _
                                    ByVal scheme As String)
    Dim criteria As String
    Select Case scheme
        Case "G3": criteria = "Distinction: A1-A2; Pass: A1-C6; Fail: D7-F9. Lower MSG is stronger."
        Case "G2": criteria = "Distinction: 1-2; Pass: 1-5; Fail: 6. Lower MSG is stronger."
        Case "G1": criteria = "Distinction: A only; Pass: A-D; Fail: E."
    End Select

    AppendHtml html, SpacerRow(10)
    AppendHtml html, CardStart(scheme & " subject performance", criteria)
    AppendHtml html, BuildPerformanceTable(subjects, subjectCount, scheme)
    AppendHtml html, CardEnd()
End Sub

Private Function BuildPerformanceTable(ByRef subjects() As tEmailSubject, _
                                       ByVal subjectCount As Long, ByVal scheme As String) As String
    Dim idx() As Long, n As Long, i As Long, j As Long, tmp As Long
    Dim html As String, k As Long
    Dim distPct As Double, passPct As Double, failPct As Double
    Dim meanText As String

    For i = 1 To subjectCount
        If subjects(i).Scheme = scheme Then
            n = n + 1
            ReDim Preserve idx(1 To n)
            idx(n) = i
        End If
    Next i

    For i = 1 To n - 1
        For j = i + 1 To n
            If SubjectPerformanceBefore(subjects(idx(j)), subjects(idx(i)), scheme) Then
                tmp = idx(i): idx(i) = idx(j): idx(j) = tmp
            End If
        Next j
    Next i

    html = TableStart()
    If scheme = "G1" Then
        html = html & PerformanceHeader(False)
    Else
        html = html & PerformanceHeader(True)
    End If

    If n = 0 Then
        html = html & "<tr><td colspan='6' style='padding:8px;color:#60788e;'>No " & scheme & " subjects found.</td></tr>"
    Else
        For k = 1 To n
            i = idx(k)
            distPct = EmailPct(subjects(i).DistCount, subjects(i).N)
            passPct = EmailPct(subjects(i).PassCount, subjects(i).N)
            failPct = EmailPct(subjects(i).FailCount, subjects(i).N)
            If subjects(i).N > 0 Then meanText = Format$(subjects(i).PointSum / subjects(i).N, "0.00") Else meanText = "-"

            html = html & "<tr>" & TextTd(subjects(i).DisplayName, "#23384d") & _
                   NumTd(CStr(subjects(i).N), "#23384d", "#ffffff") & _
                   NumTd(Format$(distPct, "0.0") & "%", "#548235", "#edf6e8") & _
                   NumTd(Format$(passPct, "0.0") & "%", "#385a78", "#f5f5f5") & _
                   NumTd(Format$(failPct, "0.0") & "%", IIf(failPct > 0, "#c00000", "#60788e"), IIf(failPct > 0, "#fff0f0", "#ffffff"))
            If scheme <> "G1" Then html = html & NumTd(meanText, "#23384d", "#ffffff")
            html = html & "</tr>"
        Next k
    End If
    html = html & "</table>"
    BuildPerformanceTable = html
End Function

Private Function SubjectPerformanceBefore(ByRef a As tEmailSubject, _
                                          ByRef b As tEmailSubject, ByVal scheme As String) As Boolean
    Dim aDist As Double, bDist As Double, aPass As Double, bPass As Double
    Dim aMean As Double, bMean As Double

    aDist = EmailPct(a.DistCount, a.N): bDist = EmailPct(b.DistCount, b.N)
    aPass = EmailPct(a.PassCount, a.N): bPass = EmailPct(b.PassCount, b.N)
    If aDist <> bDist Then SubjectPerformanceBefore = (aDist > bDist): Exit Function
    If aPass <> bPass Then SubjectPerformanceBefore = (aPass > bPass): Exit Function
    If scheme <> "G1" Then
        If a.N > 0 Then aMean = a.PointSum / a.N Else aMean = 999
        If b.N > 0 Then bMean = b.PointSum / b.N Else bMean = 999
        If aMean <> bMean Then SubjectPerformanceBefore = (aMean < bMean): Exit Function
    End If
    SubjectPerformanceBefore = (StrComp(a.DisplayName, b.DisplayName, vbTextCompare) < 0)
End Function

Private Function BuildHighlightTable(ByRef subjects() As tEmailSubject, _
                                     ByVal subjectCount As Long, ByVal perfectOnly As Boolean) As String
    Dim html As String, scheme As Variant
    Dim idx() As Long, n As Long, i As Long, j As Long, tmp As Long, k As Long
    Dim qualifies As Boolean, passPct As Double

    html = TableStart() & "<tr style='background:#eef5fb;'>" & HeaderTd("Subject") & HeaderTd("N") & HeaderTd("% Pass") & "</tr>"

    For Each scheme In Array("G3", "G2", "G1")
        n = 0
        Erase idx
        For i = 1 To subjectCount
            If subjects(i).Scheme = CStr(scheme) And subjects(i).N > 0 Then
                passPct = EmailPct(subjects(i).PassCount, subjects(i).N)
                qualifies = (perfectOnly And passPct = 100#) Or ((Not perfectOnly) And passPct < 90#)
                If qualifies Then
                    n = n + 1
                    ReDim Preserve idx(1 To n)
                    idx(n) = i
                End If
            End If
        Next i

        If n > 0 Then
            For i = 1 To n - 1
                For j = i + 1 To n
                    If HighlightBefore(subjects(idx(j)), subjects(idx(i)), perfectOnly) Then
                        tmp = idx(i): idx(i) = idx(j): idx(j) = tmp
                    End If
                Next j
            Next i

            html = html & "<tr><td colspan='3' style='padding:6px 8px;background:#eef5fb;color:#385a78;font-size:11px;font-weight:bold;'>" & scheme & " subjects</td></tr>"
            For k = 1 To n
                i = idx(k)
                html = html & "<tr>" & TextTd(subjects(i).DisplayName, "#23384d") & _
                       NumTd(CStr(subjects(i).N), "#23384d", "#ffffff") & _
                       NumTd(Format$(EmailPct(subjects(i).PassCount, subjects(i).N), "0.0") & "%", _
                             IIf(perfectOnly, "#548235", "#c00000"), "#ffffff") & "</tr>"
            Next k
        End If
    Next scheme

    If InStr(1, html, "subjects</td>", vbTextCompare) = 0 Then
        html = html & "<tr><td colspan='3' style='padding:8px;color:#60788e;'>None.</td></tr>"
    End If
    BuildHighlightTable = html & "</table>"
End Function

Private Function HighlightBefore(ByRef a As tEmailSubject, ByRef b As tEmailSubject, _
                                 ByVal perfectOnly As Boolean) As Boolean
    Dim ap As Double, bp As Double
    If perfectOnly Then
        If a.N <> b.N Then HighlightBefore = (a.N > b.N): Exit Function
    Else
        ap = EmailPct(a.PassCount, a.N): bp = EmailPct(b.PassCount, b.N)
        If ap <> bp Then HighlightBefore = (ap > bp): Exit Function
    End If
    HighlightBefore = (StrComp(a.DisplayName, b.DisplayName, vbTextCompare) < 0)
End Function

Private Function BuildTopStudentHtml(ByRef students() As tEmailStudent, _
                                     ByVal studentCount As Long, ByVal groupCode As String) As String
    Dim idx() As Long, n As Long, i As Long, j As Long, tmp As Long
    Dim shown As Long, html As String

    For i = 1 To studentCount
        If students(i).GroupCode = groupCode And students(i).DistCount > 0 Then
            n = n + 1
            ReDim Preserve idx(1 To n)
            idx(n) = i
        End If
    Next i

    For i = 1 To n - 1
        For j = i + 1 To n
            If StudentBefore(students(idx(j)), students(idx(i))) Then
                tmp = idx(i): idx(i) = idx(j): idx(j) = tmp
            End If
        Next j
    Next i

    html = "<div style='font-size:13px;font-weight:bold;color:#1f4e79;margin:14px 0 6px;'>" & groupCode & "</div>" & TableStart() & _
           "<tr style='background:#eef5fb;'>" & HeaderTd("Student") & HeaderTd("Class") & HeaderTd("Distinctions") & HeaderTd("Passes") & "</tr>"

    shown = n
    If shown > 10 Then shown = 10
    If shown = 0 Then
        html = html & "<tr><td colspan='4' style='padding:8px;color:#60788e;'>No qualifying students.</td></tr>"
    Else
        For i = 1 To shown
            j = idx(i)
            html = html & "<tr>" & TextTd(students(j).StudentName, "#23384d") & _
                   TextTd(students(j).ClassName, "#385a78") & _
                   NumTd(CStr(students(j).DistCount), "#548235", "#edf6e8") & _
                   NumTd(CStr(students(j).PassCount), "#385a78", "#f5f5f5") & "</tr>"
        Next i
    End If
    html = html & "</table>"
    If n > 10 Then html = html & "<div style='font-size:10px;color:#60788e;margin-top:4px;'>Top 10 shown.</div>"
    BuildTopStudentHtml = html
End Function

Private Function StudentBefore(ByRef a As tEmailStudent, ByRef b As tEmailStudent) As Boolean
    If a.DistCount <> b.DistCount Then StudentBefore = (a.DistCount > b.DistCount): Exit Function
    If a.PassCount <> b.PassCount Then StudentBefore = (a.PassCount > b.PassCount): Exit Function
    StudentBefore = (StrComp(a.StudentName, b.StudentName, vbTextCompare) < 0)
End Function

'------------------------------------------------------------
' OUTLOOK DRAFT (NEVER SEND)
'------------------------------------------------------------
Private Sub CreateOutlookDraft(ByVal htmlBody As String, _
                               ByVal assessmentName As String, ByVal yearText As String, _
                               ByVal levelText As String, _
                               Optional ByVal remindManualAttachment As Boolean = False)
    Dim outlookApp As Object, mailItem As Object
    Dim subjectPrefix As String, subjectText As String
    Dim existingSignature As String

    On Error Resume Next
    Set outlookApp = GetObject(, "Outlook.Application")
    If outlookApp Is Nothing Then Set outlookApp = CreateObject("Outlook.Application")
    On Error GoTo DraftFail

    If outlookApp Is Nothing Then GoTo DraftFail

    Set mailItem = outlookApp.CreateItem(0)
    subjectPrefix = GetEmailSetting("EmailSubjectPrefix", "Results Summary")
    subjectText = subjectPrefix & " - " & levelText & " " & assessmentName
    If yearText <> "" Then subjectText = subjectText & " " & yearText

    With mailItem
        .To = GetEmailSetting("EmailTo", "")
        .CC = GetEmailSetting("EmailCC", "")
        .Subject = subjectText
        .Display
        existingSignature = .HTMLBody
        .HTMLBody = htmlBody & existingSignature
    End With

    If remindManualAttachment Then
        MsgBox "Outlook draft created. Attach the detailed Excel workbook manually, then review the results, recipients and embargo before sending.", vbInformation
    Else
        MsgBox "Outlook draft created. Review the results, recipients and embargo before sending.", vbInformation
    End If
    Exit Sub

DraftFail:
    MsgBox "The summary was built, but an Outlook draft could not be opened." & vbCrLf & _
           "This feature requires desktop Microsoft Outlook with VBA/COM automation available.", vbExclamation
End Sub

'------------------------------------------------------------
' HTML HELPERS
'------------------------------------------------------------
Private Function CardStart(ByVal heading As String, ByVal subtitle As String) As String
    CardStart = "<tr><td style='background:#ffffff;border:1px solid #d7e4ef;padding:15px 16px;'>" & _
                "<div style='font-size:16px;font-weight:bold;color:#1f4e79;margin-bottom:" & IIf(subtitle = "", "10", "3") & "px;'>" & HtmlEncode(heading) & "</div>"
    If subtitle <> "" Then CardStart = CardStart & "<div style='font-size:10px;line-height:14px;color:#60788e;margin-bottom:10px;'>" & HtmlEncode(subtitle) & "</div>"
End Function

Private Function CardEnd() As String
    CardEnd = "</td></tr>"
End Function

Private Function SpacerRow(ByVal heightPx As Long) As String
    SpacerRow = "<tr><td height='" & heightPx & "' style='height:" & heightPx & "px;font-size:1px;line-height:1px;'>&nbsp;</td></tr>"
End Function

Private Function TableStart() As String
    TableStart = "<table role='presentation' width='100%' cellspacing='0' cellpadding='0' border='0' style='border-collapse:collapse;font-size:11px;'>"
End Function

Private Function ManagementTableStart() As String
    ManagementTableStart = "<table role='presentation' width='100%' cellspacing='0' cellpadding='0' border='0' style='border-collapse:collapse;font-size:13px;'>"
End Function

Private Function PerformanceHeader(ByVal includeMsg As Boolean) As String
    PerformanceHeader = "<tr style='background:#eef5fb;'>" & HeaderTd("Subject") & HeaderTd("N") & HeaderTd("% Dist") & HeaderTd("% Pass") & HeaderTd("% Fail")
    If includeMsg Then PerformanceHeader = PerformanceHeader & HeaderTd("MSG")
    PerformanceHeader = PerformanceHeader & "</tr>"
End Function

Private Function HeaderTd(ByVal valueText As String) As String
    HeaderTd = "<td style='padding:7px 8px;border-bottom:1px solid #d7e4ef;color:#385a78;font-weight:bold;'>" & HtmlEncode(valueText) & "</td>"
End Function

Private Function HeaderTdHtml(ByVal trustedHtml As String) As String
    HeaderTdHtml = "<td style='padding:7px 8px;border-bottom:1px solid #d7e4ef;color:#385a78;font-weight:bold;'>" & trustedHtml & "</td>"
End Function

Private Function CenterHeaderTd(ByVal valueText As String) As String
    CenterHeaderTd = "<td align='center' style='padding:7px 8px;border-bottom:1px solid #d7e4ef;" & _
                     "color:#385a78;font-weight:bold;text-align:center;'>" & HtmlEncode(valueText) & "</td>"
End Function

Private Function CenterHeaderTdHtml(ByVal trustedHtml As String) As String
    CenterHeaderTdHtml = "<td align='center' style='padding:7px 8px;border-bottom:1px solid #d7e4ef;" & _
                         "color:#385a78;font-weight:bold;text-align:center;'>" & trustedHtml & "</td>"
End Function

Private Function TextTd(ByVal valueText As String, ByVal colorText As String) As String
    TextTd = "<td style='padding:6px 8px;border-bottom:1px solid #e5edf4;color:" & colorText & ";'>" & HtmlEncode(valueText) & "</td>"
End Function

Private Function NumTd(ByVal valueText As String, ByVal colorText As String, ByVal bgColor As String) As String
    NumTd = "<td align='right' style='padding:6px 8px;border-bottom:1px solid #e5edf4;color:" & colorText & ";background:" & bgColor & ";'>" & HtmlEncode(valueText) & "</td>"
End Function

Private Function CenterNumTd(ByVal valueText As String, ByVal colorText As String, _
                             ByVal bgColor As String) As String
    CenterNumTd = "<td align='center' style='padding:6px 8px;border-bottom:1px solid #e5edf4;color:" & _
                  colorText & ";background:" & bgColor & ";text-align:center;'>" & HtmlEncode(valueText) & "</td>"
End Function

Private Function HtmlEncode(ByVal valueText As String) As String
    Dim s As String
    s = valueText
    s = Replace(s, "&", "&amp;")
    s = Replace(s, "<", "&lt;")
    s = Replace(s, ">", "&gt;")
    s = Replace(s, Chr$(34), "&quot;")
    HtmlEncode = s
End Function

Private Sub AppendHtml(ByRef html As String, ByVal fragment As String)
    html = html & fragment
End Sub

Private Function BuildWarningHtml(ByVal warningText As String) As String
    Dim lines As Variant, lineValue As Variant, html As String
    lines = Split(warningText, vbLf)
    For Each lineValue In lines
        If Trim$(CStr(lineValue)) <> "" Then
            html = html & "<div style='font-size:11px;line-height:16px;color:#6b5700;'>&bull; " & HtmlEncode(Trim$(CStr(lineValue))) & "</div>"
        End If
    Next lineValue
    BuildWarningHtml = html
End Function

'------------------------------------------------------------
' GENERAL HELPERS
'------------------------------------------------------------
Private Function FindEmailHeader(ByVal ws As Worksheet, ByVal headerText As String) As Long
    Dim lastCol As Long, c As Long
    lastCol = LastPopulatedEmailColumn(ws, 1)
    For c = 1 To lastCol
        If StrComp(Trim$(CStr(ws.Cells(1, c).value)), headerText, vbTextCompare) = 0 Then
            FindEmailHeader = c
            Exit Function
        End If
    Next c
End Function

Private Function FindEmailHeaderAtRow(ByVal ws As Worksheet, ByVal headerRow As Long, _
                                      ByVal headerText As String) As Long
    Dim lastCol As Long, c As Long
    lastCol = LastPopulatedEmailColumn(ws, headerRow)
    For c = 1 To lastCol
        If StrComp(Trim$(CStr(ws.Cells(headerRow, c).value)), headerText, vbTextCompare) = 0 Then
            FindEmailHeaderAtRow = c
            Exit Function
        End If
    Next c
End Function

Private Function LastPopulatedEmailColumn(ByVal ws As Worksheet, _
                                          ByVal rowNumber As Long) As Long
    Dim lastCell As Range

    ' Range.End can stop before hidden trailing columns. Find searches the
    ' full row and therefore still locates hidden AtRisk headers in O:P.
    On Error Resume Next
    Set lastCell = ws.Rows(rowNumber).Find( _
        What:="*", _
        After:=ws.Cells(rowNumber, 1), _
        LookIn:=xlFormulas, _
        LookAt:=xlPart, _
        SearchOrder:=xlByColumns, _
        SearchDirection:=xlPrevious, _
        MatchCase:=False)
    On Error GoTo 0

    If Not lastCell Is Nothing Then LastPopulatedEmailColumn = lastCell.Column
End Function

Private Function StripEmailGradeSuffix(ByVal headerText As String) As String
    Dim s As String
    s = Trim$(headerText)
    If UCase$(Right$(s, 7)) = "(GRADE)" Then s = Trim$(Left$(s, Len(s) - 7))
    StripEmailGradeSuffix = s
End Function

Private Function IsExcludedEmailSubject(ByVal subjectName As String) As Boolean
    Dim ws As Worksheet
    Dim r As Long
    Dim subjectKey As String, excludedKey As String

    subjectKey = NormalizeEmailSubjectKey(subjectName)
    If subjectKey = "" Then Exit Function

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(SETTINGS_SHEET)
    On Error GoTo 0
    If ws Is Nothing Then Exit Function

    For r = 2 To 100
        excludedKey = NormalizeEmailSubjectKey(CStr(ws.Cells(r, "V").value))
        If excludedKey <> "" And excludedKey = subjectKey Then
            IsExcludedEmailSubject = True
            Exit Function
        End If
    Next r
End Function

Private Function NormalizeEmailSubjectKey(ByVal subjectName As String) As String
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

    NormalizeEmailSubjectKey = s
End Function

Private Function GetEmailSubjectDisplayName(ByVal sourceName As String) As String
    Dim ws As Worksheet, r As Long
    Dim keyText As String, mappedText As String

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(SETTINGS_SHEET)
    On Error GoTo 0

    If Not ws Is Nothing Then
        For r = 2 To 100
            keyText = Trim$(CStr(ws.Cells(r, SUBJECT_MAP_KEY_COL).value))
            mappedText = Trim$(CStr(ws.Cells(r, SUBJECT_MAP_VALUE_COL).value))
            If keyText <> "" And StrComp(keyText, sourceName, vbTextCompare) = 0 Then
                If mappedText <> "" Then
                    GetEmailSubjectDisplayName = mappedText
                    Exit Function
                End If
            End If
        Next r
    End If

    GetEmailSubjectDisplayName = RemoveTrackSuffix(sourceName)
End Function

Private Function RemoveTrackSuffix(ByVal subjectName As String) As String
    Dim s As String, upperText As String, suffix As Variant
    s = Trim$(subjectName)
    For Each suffix In Array(" - G3", " - G2", " - G1", " - O")
        upperText = UCase$(s)
        If Right$(upperText, Len(CStr(suffix))) = CStr(suffix) Then
            s = Trim$(Left$(s, Len(s) - Len(CStr(suffix))))
            Exit For
        End If
    Next suffix
    RemoveTrackSuffix = s
End Function

Private Function GetEmailSetting(ByVal settingKey As String, ByVal defaultValue As String) As String
    Dim ws As Worksheet, r As Long
    Dim keyText As String, valueText As String

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(SETTINGS_SHEET)
    On Error GoTo 0
    If ws Is Nothing Then GetEmailSetting = defaultValue: Exit Function

    For r = 2 To 30
        keyText = Trim$(CStr(ws.Cells(r, EMAIL_KEY_COL).value))
        If StrComp(keyText, settingKey, vbTextCompare) = 0 Then
            valueText = Trim$(CStr(ws.Cells(r, EMAIL_VALUE_COL).value))
            If valueText <> "" Then
                GetEmailSetting = valueText
            Else
                GetEmailSetting = defaultValue
            End If
            Exit Function
        End If
    Next r

    GetEmailSetting = defaultValue
End Function

Private Sub ReadEmailManagementConfig(ByRef config As tEmailManagementConfig)
    Dim ws As Worksheet
    Dim minValue As Double

    config.MinCandidature = 10
    config.ConcernBelowPct = 70#
    config.MonitorBelowPct = 80#
    config.StrongPassAtLeastPct = 95#
    config.StrongDistAtLeastPct = 40#

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(SETTINGS_SHEET)
    On Error GoTo 0
    If ws Is Nothing Then Exit Sub

    EnsureEmailManagementConfigSection ws
    minValue = ReadManagementConfigNumber(ws.Range(MGMT_MIN_N_CELL).value, 10#, False)
    If minValue >= 1# And minValue <= 100000# Then config.MinCandidature = CLng(minValue)
    config.ConcernBelowPct = ReadManagementConfigNumber(ws.Range(MGMT_CONCERN_CELL).value, 70#, True)
    config.MonitorBelowPct = ReadManagementConfigNumber(ws.Range(MGMT_MONITOR_CELL).value, 80#, True)
    config.StrongPassAtLeastPct = ReadManagementConfigNumber(ws.Range(MGMT_STRONG_PASS_CELL).value, 95#, True)
    config.StrongDistAtLeastPct = ReadManagementConfigNumber(ws.Range(MGMT_STRONG_DIST_CELL).value, 40#, True)

    If config.ConcernBelowPct < 0# Or config.ConcernBelowPct > 100# Then config.ConcernBelowPct = 70#
    If config.MonitorBelowPct <= 0# Or config.MonitorBelowPct > 100# Then config.MonitorBelowPct = 80#
    If config.ConcernBelowPct >= config.MonitorBelowPct Then
        config.ConcernBelowPct = 70#
        config.MonitorBelowPct = 80#
    End If
    If config.StrongPassAtLeastPct < 0# Or config.StrongPassAtLeastPct > 100# Then config.StrongPassAtLeastPct = 95#
    If config.StrongDistAtLeastPct < 0# Or config.StrongDistAtLeastPct > 100# Then config.StrongDistAtLeastPct = 40#
End Sub

Private Sub EnsureEmailManagementConfigSection(ByVal ws As Worksheet)
    On Error Resume Next
    If Trim$(CStr(ws.Range("S1").value)) = "" Then ws.Range("S1").value = "Management email configuration"
    If Trim$(CStr(ws.Range(MGMT_MIN_N_CELL).value)) = "" Then ws.Range(MGMT_MIN_N_CELL).value = "MinimumCandidature=10"
    If Trim$(CStr(ws.Range(MGMT_CONCERN_CELL).value)) = "" Then ws.Range(MGMT_CONCERN_CELL).value = "ConcernBelow=70%"
    If Trim$(CStr(ws.Range(MGMT_MONITOR_CELL).value)) = "" Then ws.Range(MGMT_MONITOR_CELL).value = "MonitorBelow=80%"
    If Trim$(CStr(ws.Range(MGMT_STRONG_PASS_CELL).value)) = "" Then ws.Range(MGMT_STRONG_PASS_CELL).value = "StrongPassAtLeast=95%"
    If Trim$(CStr(ws.Range(MGMT_STRONG_DIST_CELL).value)) = "" Then ws.Range(MGMT_STRONG_DIST_CELL).value = "StrongDistinctionAtLeast=40%"
    If Trim$(CStr(ws.Range(MGMT_PRELIM_MODE_CELL).value)) = "" Then ws.Range(MGMT_PRELIM_MODE_CELL).value = "PrelimSummaryMode=ASK"
    ws.Range("S1").Font.Bold = True
    ws.Columns("S").AutoFit
    On Error GoTo 0
End Sub

Private Function ResolvePrelimManagementSummaryMode(ByVal assessmentName As String) As String
    Dim ws As Worksheet
    Dim rawValue As String, modeValue As String
    Dim equalsPos As Long, answer As VbMsgBoxResult

    If CanonicalExamKey(assessmentName) <> "PRELIM" Then
        ResolvePrelimManagementSummaryMode = "LEVEL"
        Exit Function
    End If

    modeValue = "ASK"
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(SETTINGS_SHEET)
    On Error GoTo 0
    If Not ws Is Nothing Then
        EnsureEmailManagementConfigSection ws
        rawValue = Trim$(CStr(ws.Range(MGMT_PRELIM_MODE_CELL).value))
        equalsPos = InStrRev(rawValue, "=")
        If equalsPos > 0 Then rawValue = Trim$(Mid$(rawValue, equalsPos + 1))
        If rawValue <> "" Then modeValue = UCase$(rawValue)
    End If

    Select Case modeValue
        Case "LEVEL"
            ResolvePrelimManagementSummaryMode = "LEVEL"
        Case "STREAM"
            ResolvePrelimManagementSummaryMode = "STREAM"
        Case Else
            answer = MsgBox("Choose the PRELIM management-summary version:" & vbCrLf & vbCrLf & _
                            "Yes  - 4EX / 4NA / 4NT / 5NA breakdown" & vbCrLf & _
                            "No   - existing Sec-level breakdown" & vbCrLf & _
                            "Cancel - stop drafting", _
                            vbYesNoCancel + vbQuestion, "PRELIM Summary Version")
            Select Case answer
                Case vbYes: ResolvePrelimManagementSummaryMode = "STREAM"
                Case vbNo: ResolvePrelimManagementSummaryMode = "LEVEL"
                Case Else: ResolvePrelimManagementSummaryMode = ""
            End Select
    End Select
End Function

Private Function ReadManagementConfigNumber(ByVal rawValue As Variant, _
                                            ByVal defaultValue As Double, _
                                            ByVal isPercentage As Boolean) As Double
    Dim valueText As String, equalsPos As Long
    Dim parsedValue As Double

    If IsError(rawValue) Or IsEmpty(rawValue) Then
        ReadManagementConfigNumber = defaultValue
        Exit Function
    End If

    If IsNumeric(rawValue) Then
        parsedValue = CDbl(rawValue)
        If isPercentage And parsedValue > 0# And parsedValue <= 1# Then parsedValue = parsedValue * 100#
        ReadManagementConfigNumber = parsedValue
        Exit Function
    End If

    valueText = Trim$(CStr(rawValue))
    equalsPos = InStrRev(valueText, "=")
    If equalsPos > 0 Then valueText = Trim$(Mid$(valueText, equalsPos + 1))
    valueText = Replace(valueText, "%", "")
    If IsNumeric(valueText) Then
        ReadManagementConfigNumber = CDbl(valueText)
    Else
        ReadManagementConfigNumber = defaultValue
    End If
End Function

Private Function GetEmailMinN() As Long
    Dim ws As Worksheet, valueData As Variant
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(SETTINGS_SHEET)
    On Error GoTo 0

    If ws Is Nothing Then GetEmailMinN = DEFAULT_MIN_N: Exit Function
    valueData = ws.Range("L6").value
    If IsNumeric(valueData) Then
        GetEmailMinN = CLng(valueData)
        If GetEmailMinN < 1 Then GetEmailMinN = DEFAULT_MIN_N
    Else
        GetEmailMinN = DEFAULT_MIN_N
    End If
End Function

Private Function EmailPct(ByVal numerator As Long, ByVal denominator As Long) As Double
    If denominator > 0 Then EmailPct = (CDbl(numerator) / CDbl(denominator)) * 100#
End Function

Private Sub AppendWarning(ByRef warningText As String, ByVal oneWarning As String)
    If warningText <> "" Then warningText = warningText & vbLf
    warningText = warningText & oneWarning
End Sub

Private Function CountWarnings(ByVal warningText As String) As Long
    If Trim$(warningText) = "" Then
        CountWarnings = 0
    Else
        CountWarnings = UBound(Split(warningText, vbLf)) + 1
    End If
End Function

Private Function RemoveWorkbookExtension(ByVal workbookName As String) As String
    Dim dotPos As Long
    dotPos = InStrRev(workbookName, ".")
    If dotPos > 1 Then
        RemoveWorkbookExtension = Left$(workbookName, dotPos - 1)
    Else
        RemoveWorkbookExtension = workbookName
    End If
End Function
