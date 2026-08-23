Attribute VB_Name = "modResetGeneratedResults"
Option Explicit

'============================================================
' Module: modResetGeneratedResults
'
' PURPOSE
'   Delete generated staging and results worksheets so that a
'   new assessment can be imported into a clean workbook.
'
' SAFETY
'   - Shows a deletion preview before doing anything.
'   - Requires an explicit Yes confirmation.
'   - Preserves Settings, Dashboard, Logs and common workbook
'     configuration/template sheets.
'   - Does not modify VBA modules, files on disk, or source
'     Cockpit workbooks.
'============================================================

'------------------------------------------------------------
' PUBLIC ENTRY POINT - assign a button to this macro
'------------------------------------------------------------
Public Sub DeleteAllGeneratedResultTabs()
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim targets As Collection
    Dim targetName As Variant
    Dim previewText As String
    Dim answer As VbMsgBoxResult
    Dim deletedCount As Long
    Dim oldDisplayAlerts As Boolean
    Dim oldScreenUpdating As Boolean

    Set wb = ThisWorkbook
    Set targets = New Collection

    If wb.ProtectStructure Then
        MsgBox "The workbook structure is protected. Unprotect the workbook structure before deleting generated tabs.", _
               vbExclamation, "Reset Generated Results"
        Exit Sub
    End If

    For Each ws In wb.Worksheets
        If IsGeneratedResultSheet(ws) Then targets.Add ws.Name
    Next ws

    If targets.count = 0 Then
        MsgBox "No generated staging or results tabs were found." & vbCrLf & _
               "Settings, Dashboard, Logs and configuration sheets were not changed.", _
               vbInformation, "Reset Generated Results"
        Exit Sub
    End If

    previewText = BuildDeletionPreview(targets, 24)
    answer = MsgBox( _
        "The following " & targets.count & " generated worksheet(s) will be permanently deleted:" & _
        vbCrLf & vbCrLf & previewText & vbCrLf & _
        "Settings, Dashboard, Logs and configuration/template sheets will be kept." & vbCrLf & vbCrLf & _
        "This cannot be undone in Excel. Continue?", _
        vbYesNo + vbCritical + vbDefaultButton2, _
        "Delete Generated Result Tabs")

    If answer <> vbYes Then Exit Sub

    oldDisplayAlerts = Application.DisplayAlerts
    oldScreenUpdating = Application.ScreenUpdating
    Application.DisplayAlerts = False
    Application.ScreenUpdating = False

    On Error GoTo DeleteFail

    For Each targetName In targets
        wb.Worksheets(CStr(targetName)).Delete
        deletedCount = deletedCount + 1
    Next targetName

CleanExit:
    Application.DisplayAlerts = oldDisplayAlerts
    Application.ScreenUpdating = oldScreenUpdating

    MsgBox deletedCount & " generated worksheet(s) deleted." & vbCrLf & _
           "Settings, Dashboard, Logs and configuration/template sheets were preserved." & vbCrLf & _
           "You can now run ParseCockpitFiles_ToStaging or ParseCockpitFolder_ToStaging.", _
           vbInformation, "Reset Complete"
    Exit Sub

DeleteFail:
    Application.DisplayAlerts = oldDisplayAlerts
    Application.ScreenUpdating = oldScreenUpdating
    MsgBox "Reset stopped after deleting " & deletedCount & " worksheet(s)." & vbCrLf & _
           "Error: " & Err.Description, vbCritical, "Reset Incomplete"
End Sub

'------------------------------------------------------------
' DETECTION
'------------------------------------------------------------
Private Function IsGeneratedResultSheet(ByVal ws As Worksheet) As Boolean
    Dim nm As String, upperName As String

    nm = Trim$(ws.Name)
    upperName = UCase$(nm)

    If IsProtectedWorkbookSheetName(upperName) Then Exit Function

    ' Normalised staging sheets have this exact core header set.
    If HasStagingCoreHeaders(ws) Then
        IsGeneratedResultSheet = True
        Exit Function
    End If

    ' Known generated report families.
    If Left$(upperName, Len("ATRISK_S")) = "ATRISK_S" Then
        IsGeneratedResultSheet = True
    ElseIf Left$(upperName, Len("TOPQUAL_S")) = "TOPQUAL_S" Then
        IsGeneratedResultSheet = True
    ElseIf Left$(upperName, Len("FT_PROGRESS_")) = "FT_PROGRESS_" Then
        IsGeneratedResultSheet = True
    ElseIf Left$(upperName, Len("FT_SUBJECTDELTA_")) = "FT_SUBJECTDELTA_" Then
        IsGeneratedResultSheet = True
    ElseIf Left$(upperName, Len("FP_")) = "FP_" Then
        ' Current form-teacher progress naming format.
        IsGeneratedResultSheet = True
    ElseIf Left$(upperName, Len("SEC_CORREL_")) = "SEC_CORREL_" Then
        IsGeneratedResultSheet = True
    ElseIf upperName = "PHYLOGS" Then
        IsGeneratedResultSheet = True
    ElseIf IsSubjectAnalysisSheetName(upperName) Then
        IsGeneratedResultSheet = True
    End If
End Function

Private Function HasStagingCoreHeaders(ByVal ws As Worksheet) As Boolean
    HasStagingCoreHeaders = _
        HasHeaderOnRowOne(ws, "RegNo") And _
        HasHeaderOnRowOne(ws, "Name") And _
        HasHeaderOnRowOne(ws, "Class") And _
        HasHeaderOnRowOne(ws, "Assessment") And _
        HasHeaderOnRowOne(ws, "Year")
End Function

Private Function HasHeaderOnRowOne(ByVal ws As Worksheet, ByVal headerText As String) As Boolean
    Dim lastCol As Long, c As Long

    lastCol = ws.Cells(1, ws.Columns.count).End(xlToLeft).Column
    For c = 1 To lastCol
        If StrComp(Trim$(CStr(ws.Cells(1, c).value)), headerText, vbTextCompare) = 0 Then
            HasHeaderOnRowOne = True
            Exit Function
        End If
    Next c
End Function

Private Function IsSubjectAnalysisSheetName(ByVal upperName As String) As Boolean
    Dim firstChar As String, levelChar As String

    If Len(upperName) < 4 Then Exit Function
    firstChar = Left$(upperName, 1)
    levelChar = Mid$(upperName, 2, 1)

    If firstChar <> "S" And firstChar <> "Y" Then Exit Function
    If levelChar < "1" Or levelChar > "5" Then Exit Function

    IsSubjectAnalysisSheetName = (InStr(1, upperName, "_SUBJ ANALYSIS_", vbBinaryCompare) > 0)
End Function

Private Function IsProtectedWorkbookSheetName(ByVal upperName As String) As Boolean
    Select Case upperName
        Case "SETTINGS", "DASHBOARD", "LOGS", "MENU", "CONFIG", "CONFIGURATION", _
             "LOOKUP", "LOOKUPS", "TEMPLATE", "INSTRUCTIONS", "README", "READ ME"
            IsProtectedWorkbookSheetName = True
            Exit Function
    End Select

    ' Preserve sheets clearly intended as workbook infrastructure.
    If InStr(1, upperName, "SETTING", vbBinaryCompare) > 0 Then
        IsProtectedWorkbookSheetName = True
    ElseIf InStr(1, upperName, "CONFIG", vbBinaryCompare) > 0 Then
        IsProtectedWorkbookSheetName = True
    ElseIf InStr(1, upperName, "TEMPLATE", vbBinaryCompare) > 0 Then
        IsProtectedWorkbookSheetName = True
    ElseIf InStr(1, upperName, "LOOKUP", vbBinaryCompare) > 0 Then
        IsProtectedWorkbookSheetName = True
    End If
End Function

'------------------------------------------------------------
' CONFIRMATION PREVIEW
'------------------------------------------------------------
Private Function BuildDeletionPreview(ByVal targets As Collection, _
                                      ByVal maximumNames As Long) As String
    Dim i As Long, shown As Long
    Dim resultText As String

    shown = targets.count
    If shown > maximumNames Then shown = maximumNames

    For i = 1 To shown
        resultText = resultText & "  - " & CStr(targets(i)) & vbCrLf
    Next i

    If targets.count > shown Then
        resultText = resultText & "  ... and " & (targets.count - shown) & " more" & vbCrLf
    End If

    BuildDeletionPreview = resultText
End Function
