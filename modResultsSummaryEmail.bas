Attribute VB_Name = "modResultsSummaryEmail"
Option Explicit

'============================================================
' Module: modResultsSummaryEmail
'
' PURPOSE
'   Build a compact HTML results summary from one SEC staging
'   sheet and open it as a Microsoft Outlook draft.
'
' SAFETY
'   This module never calls Send. The message is displayed as
'   a draft so that the user can review recipients and results.
'
' OPTIONAL SETTINGS (Settings!Q2:R30)
'   SchoolName          | School display name
'   PreparedBy          | Name/role shown in footer
'   EmbargoText         | Internal-use banner text
'   EmailTo             | Default To recipients
'   EmailCC             | Default CC recipients
'   EmailSubjectPrefix  | Default: Results Summary
'
' OPTIONAL SUBJECT DISPLAY NAMES (Settings!T2:U100)
'   T = staging subject name/header; U = email display name
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
End Type

Private Const SETTINGS_SHEET As String = "Settings"
Private Const EMAIL_KEY_COL As String = "Q"
Private Const EMAIL_VALUE_COL As String = "R"
Private Const SUBJECT_MAP_KEY_COL As String = "T"
Private Const SUBJECT_MAP_VALUE_COL As String = "U"
Private Const DEFAULT_MIN_N As Long = 10

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

    If matchingSheets.count = 1 Then
        Set SelectEmailSourceByAssessment = matchingSheets(1)
        Exit Function
    End If

    ' More than one level/year has this assessment. Ask using friendly
    ' cohort labels so the user never has to know a staging-sheet name.
    promptText = examLabels(selectedExamIndex) & " is available for:" & vbCrLf & vbCrLf
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
    promptText = promptText & vbCrLf & "Type the number or level (for example, S4):"

    levelAnswer = Trim$(InputBox(promptText, "Select Results Cohort", defaultCohort))
    If levelAnswer = "" Then Exit Function

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
            scheme = DetectEmailScheme(ws, c, lastRow, headerText)
            If scheme <> "" Then
                subjectCount = subjectCount + 1
                ReDim Preserve subjects(1 To subjectCount)

                sourceName = StripEmailGradeSuffix(headerText)
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
                            subjects(subjectCount).AbCount = subjects(subjectCount).AbCount + 1
                        ElseIf gradeText <> "VR" And gradeText <> "-" Then
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
                                  ") contains " & subjects(subjectCount).AbCount & " absence value(s), excluded from subject rates."
                End If
            Else
                AppendWarning warningText, "Skipped unrecognised grade column: " & headerText & "."
            End If
        End If
    Next c
End Sub

Private Function DetectEmailScheme(ByVal ws As Worksheet, ByVal gradeCol As Long, _
                                   ByVal lastRow As Long, ByVal headerText As String) As String
    Dim h As String, g As String
    Dim r As Long

    h = UCase$(Replace(headerText, " ", ""))
    If InStr(1, h, "-G3", vbTextCompare) > 0 Or InStr(1, h, "-O", vbTextCompare) > 0 Then
        DetectEmailScheme = "G3"
        Exit Function
    End If
    If InStr(1, h, "-G2", vbTextCompare) > 0 Then
        DetectEmailScheme = "G2"
        Exit Function
    End If
    If InStr(1, h, "-G1", vbTextCompare) > 0 Then
        DetectEmailScheme = "G1"
        Exit Function
    End If
    If InStr(1, h, "IP", vbTextCompare) > 0 Then Exit Function

    For r = 2 To lastRow
        g = UCase$(Trim$(CStr(ws.Cells(r, gradeCol).value)))
        If g <> "" And g <> "AB" And g <> "VR" And g <> "-" Then
            Select Case g
                Case "A1", "A2", "B3", "B4", "C5", "C6", "D7", "E8", "F9"
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
                Case "B": pointValue = 2: isPass = True: isDist = True
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
    Dim seen As Object

    Set seen = CreateObject("Scripting.Dictionary")
    seen.CompareMode = vbTextCompare

    regCol = FindEmailHeader(ws, "RegNo")
    nameCol = FindEmailHeader(ws, "Name")
    classCol = FindEmailHeader(ws, "Class")
    lastRow = ws.Cells(ws.Rows.count, classCol).End(xlUp).Row

    For r = 2 To lastRow
        If regCol > 0 Then keyText = Trim$(CStr(ws.Cells(r, regCol).value))
        If keyText = "" Then keyText = Trim$(CStr(ws.Cells(r, nameCol).value)) & "|" & Trim$(CStr(ws.Cells(r, classCol).value))
        If Replace(keyText, "|", "") <> "" Then
            If Not seen.Exists(keyText) Then seen.Add keyText, True
        End If
        keyText = ""
    Next r

    CountCandidates = seen.count
End Function

Private Sub CollectEmailStudents(ByVal ws As Worksheet, _
                                 ByRef subjects() As tEmailSubject, ByVal subjectCount As Long, _
                                 ByRef students() As tEmailStudent, ByRef studentCount As Long)
    Dim nameCol As Long, classCol As Long, lastRow As Long
    Dim r As Long, i As Long
    Dim studentName As String, className As String, gradeText As String
    Dim g1Taken As Long, g2Taken As Long, g3Taken As Long
    Dim isValid As Boolean, isPass As Boolean, isDist As Boolean
    Dim pointValue As Double

    nameCol = FindEmailHeader(ws, "Name")
    classCol = FindEmailHeader(ws, "Class")
    lastRow = ws.Cells(ws.Rows.count, classCol).End(xlUp).Row

    For r = 2 To lastRow
        studentName = Trim$(CStr(ws.Cells(r, nameCol).value))
        className = Trim$(CStr(ws.Cells(r, classCol).value))
        If studentName <> "" Then
            studentCount = studentCount + 1
            ReDim Preserve students(1 To studentCount)
            students(studentCount).StudentName = studentName
            students(studentCount).ClassName = className

            g1Taken = 0: g2Taken = 0: g3Taken = 0
            For i = 1 To subjectCount
                gradeText = UCase$(Trim$(CStr(ws.Cells(r, subjects(i).GradeCol).value)))
                If gradeText <> "" And gradeText <> "AB" And gradeText <> "VR" And gradeText <> "-" Then
                    EvaluateEmailGrade subjects(i).Scheme, gradeText, isValid, isPass, isDist, pointValue
                    If isValid Then
                        Select Case subjects(i).Scheme
                            Case "G1": g1Taken = g1Taken + 1
                            Case "G2": g2Taken = g2Taken + 1
                            Case "G3": g3Taken = g3Taken + 1
                        End Select
                        If isPass Then students(studentCount).PassCount = students(studentCount).PassCount + 1
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
        End If
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
    embargoText = GetEmailSetting("EmbargoText", "For Internal Use only. Embargoed until authorised for release.")

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
        Case "G1": criteria = "Distinction: A-B; Pass: A-D; Fail: E."
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
                               ByVal levelText As String)
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

    MsgBox "Outlook draft created. Review the results, recipients and embargo before sending.", vbInformation
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

Private Function PerformanceHeader(ByVal includeMsg As Boolean) As String
    PerformanceHeader = "<tr style='background:#eef5fb;'>" & HeaderTd("Subject") & HeaderTd("N") & HeaderTd("% Dist") & HeaderTd("% Pass") & HeaderTd("% Fail")
    If includeMsg Then PerformanceHeader = PerformanceHeader & HeaderTd("MSG")
    PerformanceHeader = PerformanceHeader & "</tr>"
End Function

Private Function HeaderTd(ByVal valueText As String) As String
    HeaderTd = "<td style='padding:7px 8px;border-bottom:1px solid #d7e4ef;color:#385a78;font-weight:bold;'>" & HtmlEncode(valueText) & "</td>"
End Function

Private Function TextTd(ByVal valueText As String, ByVal colorText As String) As String
    TextTd = "<td style='padding:6px 8px;border-bottom:1px solid #e5edf4;color:" & colorText & ";'>" & HtmlEncode(valueText) & "</td>"
End Function

Private Function NumTd(ByVal valueText As String, ByVal colorText As String, ByVal bgColor As String) As String
    NumTd = "<td align='right' style='padding:6px 8px;border-bottom:1px solid #e5edf4;color:" & colorText & ";background:" & bgColor & ";'>" & HtmlEncode(valueText) & "</td>"
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
    lastCol = ws.Cells(1, ws.Columns.count).End(xlToLeft).Column
    For c = 1 To lastCol
        If StrComp(Trim$(CStr(ws.Cells(1, c).value)), headerText, vbTextCompare) = 0 Then
            FindEmailHeader = c
            Exit Function
        End If
    Next c
End Function

Private Function StripEmailGradeSuffix(ByVal headerText As String) As String
    Dim s As String
    s = Trim$(headerText)
    If UCase$(Right$(s, 7)) = "(GRADE)" Then s = Trim$(Left$(s, Len(s) - 7))
    StripEmailGradeSuffix = s
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
