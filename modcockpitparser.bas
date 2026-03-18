Attribute VB_Name = "modcockpitparser"
Option Explicit
'===========================================================
' Module: modCockpitParser (School Results Analyser ? FSBB-safe)
'
' OUTPUT (per sheet):
'   RegNo | Name | Class | Assessment | Year | <Subject1...> (Score/Grade depending on setting)
'
' SUBJECT HEADERS:
' - We KEEP headers exactly as Cockpit gives them (cleaned):
'       "EL - O"      -> "EL - O"
'       "HCL - G3"    -> "HCL - G3"
'       "Maths - G2"  -> "Maths - G2"
'       "EL IP"       -> "EL IP"
'       "Maths 1"     -> "Maths 1"     (from Settings override)
'
' SUBJECT DETECTION:
' - Primary rule:
'     Detect score/grade pairs by sub-header pattern:
'       score column at c, grade column at c+1 with sub-header "B/G".
' - Fallback (legacy exports):
'     1) Header contains " - " (e.g. "EL - O", "HCL - G3"), OR
'     2) Header ends with " IP" (e.g. "EL IP"), OR
'     3) Header matches extra subject names from Settings!B14.
'
' - We use IsSubjectHeader() to:
'     * Find the FIRST subject column after GEP
'     * Decide which headers in that block are subjects (score+grade pairs)
'
' GRADE STRUCTURE (Cockpit):
' - For each subject header at column c:
'       c   : Score subcolumn (WA1 / WA2 / EYE ...)
'       c+1 : Grade subcolumn (B/G)
'
' TOGGLE: Include Grades?
' - Settings!B15:
'       TRUE  -> Scores + Grades (two columns per subject)
'       FALSE/blank -> Scores only (one column per subject)
'
' SHEET NAMING (per LevelKey + Assessment + Year):
'
'   Let:
'       Prefix   = Settings!B3  (may be blank)
'       LevelKey = from Settings mapping (e.g. 3A?3E ? S3, 3F?3J ? Y3)
'       AssessKey= e.g. "WA1"
'       YearVal  = e.g. "2025"
'
'   If Prefix is NOT blank:
'       SheetName = Prefix & "_" & LevelKey & "_" & AssessKey & "_" & YearVal
'
'   If Prefix IS blank:
'       SheetName = LevelKey & "_" & AssessKey & "_" & YearVal
'
' CLEARING:
'   If Prefix NOT blank:
'       Clear sheets whose name = Prefix or starts with Prefix & "_"
'
'   If Prefix blank:
'       Clear sheets whose names look like:
'           <Letters><Digits>_<token>_<4-digit year>
'
' FOOTERS:
' - Footer rows in source are excluded using prefixes from Settings!B10:B30.
'
' CLASS ? LEVELKEY MAPPING (SEC vs IP, etc):
' - Put patterns in Settings!D2:E50, e.g.:
'       D: ClassPattern (e.g. 3A, 3B, 3F, 3G)
'       E: LevelKey (e.g. S3, Y3, S2, Y2)
'
' EXTRA SUBJECT NAMES (odd cases like "Maths 1", "Maths 2"):
' - Settings!B14 = comma-separated list, e.g.:
'       Maths 1, Maths 2, STRANGE SUBJECT X
'   These are treated as subjects even though they don't match " - " or " IP".
'
' GRADE?POINT MAPPING (for later analysis):
' - Settings!H2:I50, e.g.:
'       A1 -> 1
'       A2 -> 2
'       ...
'       A+ -> 1
'       A  -> 4
'       B  -> 5
'       ...
'
' FAIL GRADE FORMATTING:
' - If grade is one of: D7, E8, F9, D+, D, U (case-insensitive),
'   then BOTH the score cell and grade cell are set to red font.
'
' SUBJECT COLUMN WIDTH:
' - Settings!B16 = desired width for all subject columns (Score / Grade), e.g. 9.
'   If blank/invalid, defaults to 9. Minimum enforced: 5.
'
' FORMATTING:
' - When a staging sheet is first created (no headers yet):
'   - Header row bold, light grey
'   - RegNo col width 10
'   - Name col width 30
'   - Class col width 10
'   - Assessment col width 12
'   - Year col width 8
'   - Subject columns set to width from Settings!B16
'   - Freeze row 1 + columns A:C (freeze pane at D2)
'   - Alternating very light blue shading per subject block,
'     with thin grey borders, applied AFTER data is written.
'
' NOTE: There is NO "Level" column in the output; level is implied from Class + sheet name.
'===========================================================

'========================
' TYPE DEFINITIONS
'========================
Private Type tSettings
    SourceFolder As String
    OutputSheetName As String   ' Base staging name, e.g. "Formatted" or "" for none
    MainHeaderRow As Long
    SubHeaderRow As Long
    YearRow As Long
    AssessmentRow As Long
    firstDataRow As Long
    GepHeaderText As String
    EnableLogging As Boolean
    footerPrefixes As Variant   ' array of UCASE strings
    includeGrades As Boolean    ' TRUE = add grade columns, FALSE = scores only
    SubjectColWidth As Long     ' Width for subject columns (Score/Grade)
End Type

' Class ? LevelKey mapping (from Settings D:E)
' Stored as Collection of "PATTERN<TAB>LEVELKEY"
Private gClassLevelMap As Collection
Private gClassLevelMapLoaded As Boolean

' Extra subject names (from Settings!B14)
' Stored as an array of UCASE cleaned strings
Private gExtraSubjects As Variant     ' may be Empty if none

' Grade?Point map (from Settings!H2:I50)
Private gGradeMap As Collection
Private gGradeMapLoaded As Boolean
Private gLastSettingsError As String

'========================
' PUBLIC ENTRY POINTS
'========================
Public Sub ParseCockpitFolder_ToStaging()
    ParseCockpit_ToStagingCore False
End Sub

Public Sub ParseCockpitFiles_ToStaging()
    ParseCockpit_ToStagingCore True
End Sub

Private Sub ParseCockpit_ToStagingCore(ByVal useFilePicker As Boolean)
    Dim cfg As tSettings
    If Not ReadSettings(cfg) Then
        MsgBox "Settings are incomplete." & vbCrLf & GetSettingsErrorText(), vbExclamation
        Exit Sub
    End If

    ' Always start from clean staging sheets (no append from previous runs)
    ClearStaging False

    ' Speed up
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.DisplayAlerts = False
    Dim oldCalc As XlCalculation
    oldCalc = Application.Calculation
    Application.Calculation = xlCalculationManual

    On Error GoTo CleanFail

    Dim processed As Long, failed As Long
    Dim pickedFiles As Variant
    Dim filePath As String
    processed = 0: failed = 0

    If useFilePicker Then
        pickedFiles = SelectSourceFiles()
        If IsEmpty(pickedFiles) Then
            MsgBox "No source files selected. Operation cancelled.", vbInformation
            GoTo CleanExit
        End If

        Dim v As Variant
        For Each v In pickedFiles
            filePath = CStr(v)
            If Len(filePath) > 0 Then
                On Error GoTo HandleOneFail
                ProcessOneWorkbook filePath, cfg
                processed = processed + 1
            End If
NextPickedFile:
            On Error GoTo CleanFail
        Next v
    Else
        Dim folderPath As String
        folderPath = SelectSourceFolder()
        If Len(folderPath) = 0 Then
            MsgBox "No source folder selected. Operation cancelled.", vbInformation
            GoTo CleanExit
        End If
        folderPath = EnsureTrailingPathSeparator(folderPath)

        Dim f As String
        f = Dir(folderPath & "*.xls*")
        Do While Len(f) > 0
            filePath = folderPath & f
            If Left$(f, 2) <> "~$" Then
                On Error GoTo HandleOneFail
                ProcessOneWorkbook filePath, cfg
                processed = processed + 1
            End If
NextFolderFile:
            On Error GoTo CleanFail
            f = Dir()
        Loop
    End If
    
CleanExit:
    Application.Calculation = oldCalc
    Application.DisplayAlerts = True
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    
    MsgBox "Done. Processed: " & processed & "   Failed: " & failed, vbInformation
    Exit Sub

CleanFail:
    Application.Calculation = oldCalc
    Application.DisplayAlerts = True
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    MsgBox "Unexpected error: " & Err.Description, vbExclamation
    Exit Sub

HandleOneFail:
    failed = failed + 1
    AppendLog "Failed", filePath, Err.Description
    Err.Clear
    If useFilePicker Then
        GoTo NextPickedFile
    Else
        GoTo NextFolderFile
    End If
End Sub

Public Sub ClearStaging(Optional showMessage As Boolean = True)
    Dim cfg As tSettings
    If Not ReadSettings(cfg) Then
        MsgBox "Settings are incomplete." & vbCrLf & GetSettingsErrorText(), vbExclamation
        Exit Sub
    End If
    
    Dim baseName As String
    baseName = Trim$(cfg.OutputSheetName)
    
    Dim ws As Worksheet
    If Len(baseName) > 0 Then
        ' Prefix provided ? clear only sheets with that prefix
        For Each ws In ThisWorkbook.Worksheets
            If ws.Name = baseName Or Left$(ws.Name, Len(baseName) + 1) = baseName & "_" Then
                ws.Cells.Clear
            End If
        Next ws
    Else
        ' No prefix ? clear only sheets that look like <Letters><Digits>_<token>_<4-digit year>
        For Each ws In ThisWorkbook.Worksheets
            If IsLevelAssessYearName(ws.Name) Then
                ws.Cells.Clear
            End If
        Next ws
    End If
    
    If showMessage Then
        MsgBox "Staging sheets cleared."
    End If
End Sub

'========================
' SETTINGS
'========================
Private Function ReadSettings(ByRef cfg As tSettings) As Boolean
    gLastSettingsError = ""
    On Error GoTo ErrRead
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("Settings")
    On Error GoTo ErrRead
    If ws Is Nothing Then
        gLastSettingsError = "Missing worksheet 'Settings'."
        ReadSettings = False
        Exit Function
    End If
    
    ' SourceFolder is now selected via dialog, not read from settings
    cfg.OutputSheetName = NzStr(ws.Range("B3").value, "")     ' can be blank for no prefix
    cfg.MainHeaderRow = NzLng(ws.Range("B4").value, 10)
    cfg.SubHeaderRow = NzLng(ws.Range("B5").value, 11)
    cfg.YearRow = NzLng(ws.Range("B6").value, 3)
    cfg.AssessmentRow = NzLng(ws.Range("B7").value, 6)
    cfg.firstDataRow = NzLng(ws.Range("B8").value, 12)
    cfg.GepHeaderText = NzStr(ws.Range("B9").value, "GEP Indicator")
    If IsError(ws.Range("B13").value) Then
        cfg.EnableLogging = False
    Else
        cfg.EnableLogging = (UCase$(Trim$(CStr(ws.Range("B13").value))) = "TRUE")
    End If
    cfg.footerPrefixes = ReadFooterPrefixes(ws)
    
    ' Extra subject names in Settings!B14 (comma-separated)
    gExtraSubjects = ReadExtraSubjects(ws)
    
    ' IncludeGrades toggle in Settings!B15
    Dim vFlag As Variant, sFlag As String
    vFlag = ws.Range("B15").value
    If VarType(vFlag) = vbBoolean Then
        cfg.includeGrades = CBool(vFlag)
    ElseIf IsError(vFlag) Then
        cfg.includeGrades = False
    Else
        sFlag = UCase$(Trim$(CStr(vFlag)))
        cfg.includeGrades = (sFlag = "TRUE" Or sFlag = "YES" Or sFlag = "Y" Or sFlag = "1")
    End If
    
    ' Subject column width ? Settings!B16 (default 9, min 5)
    cfg.SubjectColWidth = NzLng(ws.Range("B16").value, 9)
    If cfg.SubjectColWidth < 5 Then cfg.SubjectColWidth = 5

    If cfg.MainHeaderRow < 1 Or cfg.SubHeaderRow < 1 Or cfg.firstDataRow < 1 Then
        gLastSettingsError = "Invalid row settings (B4/B5/B8). Rows must be >= 1."
        ReadSettings = False
        Exit Function
    End If
    If cfg.firstDataRow <= cfg.SubHeaderRow Then
        gLastSettingsError = "Invalid row order: B8 (first data row) must be below B5 (sub-header row)."
        ReadSettings = False
        Exit Function
    End If
    
    ' Always refresh class?level mapping from Settings!D2:E50
    InitClassLevelMap ws
    
    ' Grade?Point map (lazy-loaded on first use)
    gGradeMapLoaded = False
    Set gGradeMap = Nothing
    ReadSettings = True
    Exit Function
Bad:
    If gLastSettingsError = "" Then gLastSettingsError = "Invalid Settings values."
    ReadSettings = False
    Exit Function
ErrRead:
    If gLastSettingsError = "" Then
        gLastSettingsError = "Error while reading Settings: " & Err.Description
    End If
    ReadSettings = False
End Function

Private Function GetSettingsErrorText() As String
    If Len(gLastSettingsError) > 0 Then
        GetSettingsErrorText = gLastSettingsError
    Else
        GetSettingsErrorText = "Check worksheet 'Settings' (B3:B16, D2:E50)."
    End If
End Function

Private Function ReadFooterPrefixes(ByVal ws As Worksheet) As Variant
    ' Reads Settings!B10:B30 and returns uppercase prefixes.
    Dim rng As Range, c As Range, col As Collection, s As String
    Set col = New Collection
    
    On Error Resume Next
    Set rng = ws.Range("B10:B30")
    On Error GoTo 0
    
    If Not rng Is Nothing Then
        For Each c In rng.Cells
            If IsError(c.value) Then
                s = ""
            Else
                s = CleanText(CStr(c.value))
            End If
            If Len(s) > 0 Then col.Add UCase$(s)
        Next c
    End If
    
    If col.count = 0 Then
        ' Defaults if user leaves it blank
        Dim defaults As Variant
        defaults = Array("NUMBER OF STUDENT", "NUMBER OF STUDENT PASS", _
                         "PERCENTAGE PASS", "END OF REPORT")
        ReadFooterPrefixes = defaults
    Else
        Dim i As Long
        Dim arr() As String
        ReDim arr(1 To col.count)
        For i = 1 To col.count
            arr(i) = col(i)
        Next i
        ReadFooterPrefixes = arr
    End If
End Function

Private Function ReadExtraSubjects(ByVal ws As Worksheet) As Variant
    ' Settings!B14: "Maths 1, Maths 2, STRANGE SUBJECT X"
    Dim txt As String
    If IsError(ws.Range("B14").value) Then
        txt = ""
    Else
        txt = CStr(ws.Range("B14").value)
    End If
    txt = Replace(txt, vbCr, " ")
    txt = Replace(txt, vbLf, " ")
    txt = Trim$(txt)
    
    If Len(txt) = 0 Then
        ReadExtraSubjects = Empty
        Exit Function
    End If
    
    Dim parts() As String
    parts = Split(txt, ",")
    
    Dim cleaned() As String
    Dim i As Long, s As String, count As Long
    ReDim cleaned(0 To UBound(parts))
    
    count = 0
    For i = LBound(parts) To UBound(parts)
        s = UCase$(CleanText(parts(i)))
        If Len(s) > 0 Then
            cleaned(count) = s
            count = count + 1
        End If
    Next i
    
    If count = 0 Then
        ReadExtraSubjects = Empty
    Else
        ReDim Preserve cleaned(0 To count - 1)
        ReadExtraSubjects = cleaned
    End If
End Function

Private Sub InitClassLevelMap(ByVal ws As Worksheet)
    ' Reads Settings!D2:E50 as ClassPattern -> LevelKey
    Dim rng As Range, r As Range
    Dim pattern As String, lvlKey As String
    
    Set gClassLevelMap = New Collection
    
    On Error Resume Next
    Set rng = ws.Range("D2:E50")
    On Error GoTo 0
    
    If Not rng Is Nothing Then
        For Each r In rng.Rows
            If IsError(r.Cells(1, 1).value) Or IsError(r.Cells(1, 2).value) Then
                pattern = ""
                lvlKey = ""
            Else
                pattern = CleanText(CStr(r.Cells(1, 1).value))  ' Col D
                lvlKey = CleanText(CStr(r.Cells(1, 2).value))   ' Col E
            End If
            If Len(pattern) > 0 And Len(lvlKey) > 0 Then
                pattern = UCase$(Replace(pattern, " ", "")) ' normalise key
                AddClassLevelMapEntry pattern, lvlKey
            End If
        Next r
    End If
    
    gClassLevelMapLoaded = True
End Sub

Private Sub AddClassLevelMapEntry(ByVal pattern As String, ByVal lvlKey As String)
    Dim entry As Variant
    Dim p As Long
    Dim existingPat As String

    If gClassLevelMap Is Nothing Then Exit Sub

    For Each entry In gClassLevelMap
        p = InStr(1, CStr(entry), vbTab, vbBinaryCompare)
        If p > 0 Then
            existingPat = Left$(CStr(entry), p - 1)
            If existingPat = pattern Then Exit Sub
        End If
    Next entry

    gClassLevelMap.Add pattern & vbTab & lvlKey
End Sub

'========================
' CORE PROCESSOR (WIDE OUTPUT, SCORE + optional GRADE, PER LEVEL+ASSESSMENT+YEAR)
'========================
Private Sub ProcessOneWorkbook(ByVal fullPath As String, ByRef cfg As tSettings)
    Dim wb As Workbook, sh As Worksheet
    Set wb = Application.Workbooks.Open(FileName:=fullPath, ReadOnly:=True)
    
    ' Prefer RE_RES_078; else first visible sheet
    On Error Resume Next
    Set sh = wb.Worksheets("RE_RES_078")
    On Error GoTo 0
    If sh Is Nothing Then
        Dim tmp As Worksheet
        For Each tmp In wb.Worksheets
            If tmp.Visible = xlSheetVisible Then
                Set sh = tmp
                Exit For
            End If
        Next tmp
    End If
    If sh Is Nothing Then
        AppendLog "Skipped", fullPath, "No visible sheets"
        wb.Close False
        Exit Sub
    End If
    
    ' Year & Assessment
    Dim yearVal As String, assessToken As String
    yearVal = ParseYearFromLine(Trim$(CStr(sh.Cells(cfg.YearRow, 1).value)))
    assessToken = ParseAssessmentFromLine(Trim$(CStr(sh.Cells(cfg.AssessmentRow, 1).value)))
    
    ' Header anchors in Cockpit sheet
    Dim hdrRow As Long
    hdrRow = cfg.MainHeaderRow
    
    Dim colReg As Long, colName As Long, colClass As Long, colGep As Long
    colReg = FindHeaderCol(sh, hdrRow, "Reg#")
    colName = FindHeaderCol(sh, hdrRow, "Name")
    colClass = FindHeaderCol(sh, hdrRow, "Class")
    colGep = FindHeaderCol(sh, hdrRow, cfg.GepHeaderText)
    If colReg = 0 Or colName = 0 Or colClass = 0 Or colGep = 0 Then
        AppendLog "Skipped", fullPath, "Missing Reg#/Name/Class/GEP Indicator headers"
        wb.Close False
        Exit Sub
    End If
    
    Dim lastCol As Long
    lastCol = sh.Cells(hdrRow, sh.Columns.count).End(xlToLeft).Column
    
    ' Find first subject column robustly (skip any extra non-subject columns after GEP)
    Dim firstSubCol As Long
    firstSubCol = FindFirstSubjectCol(sh, hdrRow, cfg.SubHeaderRow, colGep, lastCol)
    If firstSubCol = 0 Then
        AppendLog "Skipped", fullPath, "No subject header detected after GEP Indicator"
        wb.Close False
        Exit Sub
    End If
    
    ' Build subject list ? ONE score column + ONE grade column per subject (we might ignore grade later)
    Dim subjPairs As Collection
    Set subjPairs = New Collection

    Dim c As Long
    Dim subHdrRow As Long
    subHdrRow = cfg.SubHeaderRow

    c = firstSubCol
    Do While c <= lastCol
        Dim rawSubj As String, subj As String
        rawSubj = CleanText(CStr(sh.Cells(hdrRow, c).value))
        If Not IsSubjectPair(sh, hdrRow, subHdrRow, c, lastCol) Then Exit Do

        subj = NormalizeSubjectName(rawSubj)   ' FSBB-safe: keep "- O", "- G3", "IP", "Maths 1", etc.

        Dim p As cSubjPair
        Set p = New cSubjPair
        p.Subject = subj
        p.scoreCol = c          ' score column; grade at c+1
        p.gradeCol = c + 1
        subjPairs.Add p

        c = c + 2               ' skip grade col in pair
    Loop

    ' Collect additional columns from sub-header row after subject columns
    Dim additionalCols As Collection
    Set additionalCols = New Collection

    ' Continue from current position c (after last subject pair)
    Do While c <= lastCol
        Dim addlHeader As String
        addlHeader = CleanText(CStr(sh.Cells(subHdrRow, c).value))

        ' Include any non-empty headers from sub-header row
        If Len(addlHeader) > 0 Then
            additionalCols.Add Array(addlHeader, c)
        End If

        c = c + 1
    Loop
    
    If subjPairs.count = 0 Then
        AppendLog "Skipped", fullPath, "Subject region empty/invalid after first detected subject header"
        wb.Close False
        Exit Sub
    End If
    
    ' Determine last student row using footer prefixes
    Dim lastDataRow As Long
    lastDataRow = FindLastStudentRow(sh, cfg.firstDataRow, colReg, cfg.footerPrefixes)
    If lastDataRow < cfg.firstDataRow Then
        AppendLog "Skipped", fullPath, "No student rows"
        wb.Close False
        Exit Sub
    End If
    
    ' Determine LEVEL + LEVELKEY from first non-blank Class
    Dim lvlStr As String, r As Long, klass As String
    Dim firstClass As String
    lvlStr = ""
    firstClass = ""
    
    For r = cfg.firstDataRow To lastDataRow
        klass = CleanText(CStr(sh.Cells(r, colClass).value))
        If Len(klass) > 0 Then
            firstClass = klass
            lvlStr = InferLevelFromClass(klass)   ' e.g. "Sec 3"
            If Len(lvlStr) > 0 Then Exit For
        End If
    Next r
    If Len(lvlStr) = 0 Then
        lvlStr = ""   ' MakeLevelKey will handle fallback
    End If
    
    ' Build sheet name using Prefix + LevelKey + AssessKey + Year
    Dim baseName As String, levelKey As String, assessKey As String
    Dim outSheetName As String, yearKey As String
    
    baseName = Trim$(cfg.OutputSheetName)           ' may be blank
    levelKey = GetLevelKeyFromClass(firstClass, lvlStr)   ' uses Settings mapping
    assessKey = MakeAssessKey(assessToken)          ' e.g. WA1 / EYE
    yearKey = IIf(Len(yearVal) > 0, yearVal, "Yr0") ' simple fallback
    
    If Len(baseName) > 0 Then
        outSheetName = baseName & "_" & levelKey & "_" & assessKey & "_" & yearKey
    Else
        outSheetName = levelKey & "_" & assessKey & "_" & yearKey
    End If
    
    ' Prepare output sheet (wide, score + optional grade)
    Dim wsOut As Worksheet
    Set wsOut = GetOrCreateSheet(outSheetName)
    
    Dim map As Collection
    Set map = New Collection   ' Keys: Subject&"|S" or Subject&"|G" -> output column
    
    ' Pass in IncludeGrades + SubjectColWidth + Additional Columns
    SetupHeadersSimple wsOut, subjPairs, additionalCols, map, cfg.includeGrades, cfg.SubjectColWidth
    
    ' Get meta column indices once (by header text)
    Dim colAss As Long, colYear As Long
    colAss = FindHeaderByText(wsOut, "Assessment")
    colYear = FindHeaderByText(wsOut, "Year")
    
    ' Emit one row per student
    For r = cfg.firstDataRow To lastDataRow
        Dim regNo As String, nm As String
        regNo = CleanText(CStr(sh.Cells(r, colReg).value))
        nm = CleanText(CStr(sh.Cells(r, colName).value))
        klass = CleanText(CStr(sh.Cells(r, colClass).value))
        
        If Len(regNo) = 0 And Len(nm) = 0 Then GoTo NextR
        
        Dim outRow As Long
        outRow = wsOut.Cells(wsOut.Rows.count, 1).End(xlUp).Row + 1
        
        wsOut.Cells(outRow, 1).value = regNo
        wsOut.Cells(outRow, 2).value = nm
        wsOut.Cells(outRow, 3).value = klass
        wsOut.Cells(outRow, colAss).value = assessToken
        wsOut.Cells(outRow, colYear).value = yearVal
        ' No Level column written
        
        ' Subjects: write score (and grade if enabled)
        Dim sp As cSubjPair
        Dim outColS As Long, outColG As Long
        Dim gradeVal As String

        For Each sp In subjPairs
            outColS = 0
            outColG = 0

            If MapExists(map, sp.Subject & "|S") Then
                outColS = MapGetLong(map, sp.Subject & "|S")
                wsOut.Cells(outRow, outColS).value = sh.Cells(r, sp.scoreCol).value
            End If

            If cfg.includeGrades Then
                gradeVal = CleanText(CStr(sh.Cells(r, sp.gradeCol).value))
                If Len(gradeVal) > 0 And MapExists(map, sp.Subject & "|G") Then
                    outColG = MapGetLong(map, sp.Subject & "|G")
                    wsOut.Cells(outRow, outColG).value = gradeVal

                    ' If this is a failing grade, make both score + grade red
                    If IsFailGrade(gradeVal) Then
                        wsOut.Cells(outRow, outColG).Font.Color = vbRed
                        If outColS > 0 Then
                            wsOut.Cells(outRow, outColS).Font.Color = vbRed
                        End If
                    End If
                End If
            End If
        Next sp

        ' Additional columns: copy data directly from sub-header columns
        Dim addlCol As Variant
        For Each addlCol In additionalCols
            Dim addlKey As String
            addlKey = "ADDL|" & CStr(addlCol(0))
            If MapExists(map, addlKey) Then
                Dim addlOutCol As Long
                addlOutCol = MapGetLong(map, addlKey)
                wsOut.Cells(outRow, addlOutCol).value = sh.Cells(r, CLng(addlCol(1))).value
            End If
        Next addlCol
NextR:
    Next r
    
    ' Apply shading & borders AFTER data is written
    ApplySubjectShading wsOut, subjPairs, additionalCols, cfg.includeGrades
    
    AppendLog "OK", fullPath, "Imported to " & outSheetName & " (wide, scores" & IIf(cfg.includeGrades, "+grades", "") & ", FSBB-safe headers)"
    wb.Close False
End Sub

'========================
' HEADER ? SCORE + optional GRADE + FORMATTING
'========================
Private Sub SetupHeadersSimple(ByVal wsOut As Worksheet, _
                               ByVal subjPairs As Collection, _
                               ByVal additionalCols As Collection, _
                               ByRef map As Collection, _
                               ByVal includeGrades As Boolean, _
                               ByVal subjWidth As Long)

    Dim headerEmpty As Boolean
    headerEmpty = (Application.WorksheetFunction.CountA(wsOut.Rows(1)) = 0)

    Dim lastCol As Long, c As Long, h As String
    Dim sp As cSubjPair
    Dim col As Long
    Dim sc As Long, w As Long

    w = subjWidth
    If w < 5 Then w = 5   ' safety

    If headerEmpty Then
        '==========================
        ' CREATE HEADER ROW
        '==========================
        wsOut.Cells(1, 1).value = "RegNo"
        wsOut.Cells(1, 2).value = "Name"
        wsOut.Cells(1, 3).value = "Class"
        wsOut.Cells(1, 4).value = "Assessment"
        wsOut.Cells(1, 5).value = "Year"

        col = 6

        ' SUBJECT HEADERS
        For Each sp In subjPairs
            If includeGrades Then
                wsOut.Cells(1, col).value = sp.Subject & " (Score)"
                MapSet map, sp.Subject & "|S", col
                col = col + 1

                wsOut.Cells(1, col).value = sp.Subject & " (Grade)"
                MapSet map, sp.Subject & "|G", col
                col = col + 1
            Else
                wsOut.Cells(1, col).value = sp.Subject
                MapSet map, sp.Subject & "|S", col
                col = col + 1
            End If
        Next sp

        ' ADDITIONAL COLUMN HEADERS (from sub-header row)
        Dim addlCol As Variant
        For Each addlCol In additionalCols
            wsOut.Cells(1, col).value = CStr(addlCol(0))
            MapSet map, "ADDL|" & CStr(addlCol(0)), col
            col = col + 1
        Next addlCol

        '==========================
        ' FORMAT HEADER ROW
        '==========================
        Dim hdr As Range
        Dim lastColF As Long
        lastColF = wsOut.Cells(1, wsOut.Columns.count).End(xlToLeft).Column

        Set hdr = wsOut.Range(wsOut.Cells(1, 1), wsOut.Cells(1, lastColF))
        With hdr
            .Font.Bold = True
            .Interior.Color = RGB(220, 220, 220)   ' light grey header
            .VerticalAlignment = xlVAlignCenter
        End With

        ' Standard useful widths
        wsOut.Columns(1).ColumnWidth = 10
        wsOut.Columns(2).ColumnWidth = 30
        wsOut.Columns(3).ColumnWidth = 10
        wsOut.Columns(4).ColumnWidth = 12
        wsOut.Columns(5).ColumnWidth = 8

        ' AutoFit, then override subject width
        If lastColF >= 6 Then
            wsOut.Range(wsOut.Cells(1, 6), wsOut.Cells(1, lastColF)).EntireColumn.AutoFit

            For sc = 6 To lastColF
                wsOut.Columns(sc).ColumnWidth = w
            Next sc
        End If

        '==========================
        ' FREEZE: Row 1 + Columns A:C
        '==========================
        On Error Resume Next
        wsOut.Activate
        If Not ActiveWindow Is Nothing Then
            ActiveWindow.FreezePanes = False
            wsOut.Range("D2").Select   ' Freeze above row 2, left of column D -> row 1 + cols A:C
            ActiveWindow.FreezePanes = True
        End If
        On Error GoTo 0

        ' NOTE: Shading is now applied AFTER data is written in ProcessOneWorkbook.

    Else
        '==========================
        ' HEADER EXISTS (2nd+ file merged into same sheet)
        ' Map existing headers back to dictionary
        '==========================
        lastCol = wsOut.Cells(1, wsOut.Columns.count).End(xlToLeft).Column
        For c = 1 To lastCol
            h = CleanText(CStr(wsOut.Cells(1, c).value))
            If Len(h) > 0 Then
                Select Case h
                    Case "RegNo", "Name", "Class", "Assessment", "Year"
                        ' ignore
                    Case Else
                        If Right$(h, 7) = "(Score)" Then
                            MapSet map, Trim$(Left$(h, Len(h) - 7)) & "|S", c
                        ElseIf Right$(h, 7) = "(Grade)" Then
                            MapSet map, Trim$(Left$(h, Len(h) - 7)) & "|G", c
                        Else
                            MapSet map, h & "|S", c
                        End If
                End Select
            End If
        Next c

        ' Enforce subject width again
        For sc = 6 To lastCol
            wsOut.Columns(sc).ColumnWidth = w
        Next sc

        ' Map additional columns from existing headers
        For Each addlCol In additionalCols
            Dim addlHeader As String
            addlHeader = CStr(addlCol(0))
            Dim foundAddl As Boolean
            foundAddl = False

            For c = 1 To lastCol
                h = CleanText(CStr(wsOut.Cells(1, c).value))
                If StrComp(h, addlHeader, vbTextCompare) = 0 Then
                    MapSet map, "ADDL|" & addlHeader, c
                    foundAddl = True
                    Exit For
                End If
            Next c

            ' If not found, add new column
            If Not foundAddl Then
                wsOut.Cells(1, lastCol + 1).value = addlHeader
                MapSet map, "ADDL|" & addlHeader, lastCol + 1
                lastCol = lastCol + 1
            End If
        Next addlCol

        ' Shading will be re-applied after data writing in ProcessOneWorkbook.
    End If
End Sub

'==========================
' ALTERNATING SUBJECT SHADING (Optimised)
'==========================
Private Sub ApplySubjectShading(ByVal wsOut As Worksheet, _
                                ByVal subjPairs As Collection, _
                                ByVal additionalCols As Collection, _
                                ByVal includeGrades As Boolean)

    Dim lastRow As Long
    Dim startCol As Long, endCol As Long
    Dim block As Long
    Dim sp As cSubjPair
    Dim col As Long

    ' Determine last used row using column A (RegNo)
    lastRow = wsOut.Cells(wsOut.Rows.count, 1).End(xlUp).Row
    If lastRow < 2 Then lastRow = 2   ' safety

    block = 0
    col = 6   ' first subject column

    For Each sp In subjPairs
        block = block + 1

        If includeGrades Then
            startCol = col
            endCol = col + 1
            col = col + 2
        Else
            startCol = col
            endCol = col
            col = col + 1
        End If

        '===============================
        ' ALTERNATING LIGHT-BLUE SHADING
        '===============================
        If block Mod 2 = 1 Then
            With wsOut.Range(wsOut.Cells(1, startCol), wsOut.Cells(lastRow, endCol))
                .Interior.Color = RGB(240, 248, 255)   ' very light blue
            End With
        Else
            ' Ensure even blocks have NO shading
            With wsOut.Range(wsOut.Cells(1, startCol), wsOut.Cells(lastRow, endCol))
                .Interior.ColorIndex = xlNone
            End With
        End If

        '===================================
        ' THIN GREY BORDER AROUND THIS BLOCK
        '===================================
        With wsOut.Range(wsOut.Cells(1, startCol), wsOut.Cells(lastRow, endCol)).Borders
            .LineStyle = xlContinuous
            .Color = RGB(200, 200, 200) ' light grey
            .Weight = xlThin
        End With

    Next sp
End Sub

'========================
' SUBJECT HEADER DETECTION
'========================
Private Function IsSubjectHeader(ByVal raw As String) As Boolean
    Dim s As String
    s = CleanText(raw)
    If Len(s) = 0 Then Exit Function
    
    ' Typical forms:
    '   "EL - O", "HCL - G3", "Maths - G2"
    '   "EL IP", "HCL IP"
    If InStr(1, s, " - ", vbTextCompare) > 0 Then
        IsSubjectHeader = True
        Exit Function
    End If
    
    If Right$(s, 3) = " IP" Then
        IsSubjectHeader = True
        Exit Function
    End If
    
    ' Extra subject names from Settings!B14
    If IsExtraSubject(s) Then
        IsSubjectHeader = True
        Exit Function
    End If
End Function

Private Function IsSubjectPair(ByVal ws As Worksheet, _
                               ByVal mainHdrRow As Long, _
                               ByVal subHdrRow As Long, _
                               ByVal scoreCol As Long, _
                               ByVal lastCol As Long) As Boolean
    Dim mainHdr As String
    Dim gradeSubHdr As String

    If scoreCol <= 0 Or scoreCol + 1 > lastCol Then Exit Function

    mainHdr = CleanText(CStr(ws.Cells(mainHdrRow, scoreCol).value))
    If Len(mainHdr) = 0 Then Exit Function

    gradeSubHdr = UCase$(CleanText(CStr(ws.Cells(subHdrRow, scoreCol + 1).value)))

    ' Cockpit subject pairs always use score column + B/G column.
    If gradeSubHdr = "B/G" Then
        IsSubjectPair = True
        Exit Function
    End If

    ' Fallback for legacy exports without clean sub-headers.
    IsSubjectPair = IsSubjectHeader(mainHdr)
End Function

Private Function IsExtraSubject(ByVal s As String) As Boolean
    Dim arr As Variant, i As Long
    IsExtraSubject = False
    arr = gExtraSubjects
    If IsEmpty(arr) Then Exit Function
    
    For i = LBound(arr) To UBound(arr)
        If s = arr(i) Then
            IsExtraSubject = True
            Exit Function
        End If
    Next i
End Function

Private Function FindFirstSubjectCol(ByVal ws As Worksheet, ByVal hdrRow As Long, _
                                     ByVal subHdrRow As Long, ByVal colGep As Long, _
                                     ByVal lastCol As Long) As Long
    Dim c As Long
    For c = colGep + 1 To lastCol
        If IsSubjectPair(ws, hdrRow, subHdrRow, c, lastCol) Then
            FindFirstSubjectCol = c
            Exit Function
        End If
    Next c
    FindFirstSubjectCol = 0
End Function

'========================
' FAIL GRADE DETECTION
'========================
Private Function IsFailGrade(ByVal gradeVal As String) As Boolean
    Dim g As String
    g = UCase$(CleanText(gradeVal))
    
    ' O-Level style fails
    If g = "D7" Or g = "E8" Or g = "F9" Then
        IsFailGrade = True
        Exit Function
    End If

    ' Some exports encode F9 as "9".
    If g = "9" Then
        IsFailGrade = True
        Exit Function
    End If
    
    ' FSBB style fails
    If g = "D+" Or g = "D" Or g = "U" Then
        IsFailGrade = True
        Exit Function
    End If

    ' G2 / G1 fail bands
    If g = "6" Or g = "E" Then
        IsFailGrade = True
        Exit Function
    End If
    
    IsFailGrade = False
End Function

'========================
' LAST STUDENT ROW via Settings prefixes
'========================
Private Function FindLastStudentRow(ByVal ws As Worksheet, ByVal firstDataRow As Long, _
                                    ByVal colReg As Long, ByVal footerPrefixes As Variant) As Long
    Dim r As Long, val As String, up As String
    Dim lastNonBlank As Long
    lastNonBlank = 0
    
    For r = firstDataRow To ws.Rows.count
        val = CleanText(CStr(ws.Cells(r, colReg).value))
        up = UCase$(val)
        
        If Len(val) > 0 Then
            If StartsWithAny(up, footerPrefixes) Then
                FindLastStudentRow = IIf(lastNonBlank > 0, lastNonBlank, r - 1)
                Exit Function
            Else
                lastNonBlank = r
            End If
        Else
            If r - lastNonBlank > 200 Then Exit For
        End If
        
        If r > ws.UsedRange.Row + ws.UsedRange.Rows.count + 50 Then Exit For
    Next r
    
    FindLastStudentRow = lastNonBlank
End Function

Private Function StartsWithAny(ByVal upValue As String, ByVal prefixes As Variant) As Boolean
    Dim i As Long, p As String
    StartsWithAny = False
    If IsArray(prefixes) Then
        For i = LBound(prefixes) To UBound(prefixes)
            p = CStr(prefixes(i))
            If Len(p) > 0 Then
                If Left$(upValue, Len(p)) = p Then
                    StartsWithAny = True
                    Exit Function
                End If
            End If
        Next i
    End If
End Function

'------------------------
' Collection key-value helpers (cross-platform replacement for Dictionary)
'------------------------
Private Sub MapSet(ByRef map As Collection, ByVal key As String, ByVal value As Variant)
    On Error Resume Next
    map.Remove key
    Err.Clear
    On Error GoTo 0
    map.Add value, key
End Sub

Private Function MapExists(ByVal map As Collection, ByVal key As String) As Boolean
    On Error GoTo Nope
    Dim tmp As Variant
    tmp = map(key)
    MapExists = True
    Exit Function
Nope:
    MapExists = False
End Function

Private Function MapGetLong(ByVal map As Collection, ByVal key As String, Optional ByVal fallback As Long = 0) As Long
    On Error GoTo Nope
    MapGetLong = CLng(map(key))
    Exit Function
Nope:
    MapGetLong = fallback
End Function

Private Function MapGetVariant(ByVal map As Collection, ByVal key As String) As Variant
    On Error GoTo Nope
    MapGetVariant = map(key)
    Exit Function
Nope:
    MapGetVariant = Empty
End Function

'========================
' UTILITIES
'========================
Private Function GetOrCreateSheet(ByVal sheetName As String) As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.count))
        ws.Name = sheetName
    End If
    Set GetOrCreateSheet = ws
End Function

Private Function SelectSourceFolder() As String
    ' Shows a folder picker dialog that works on both Mac and Windows
    Dim folderPath As String
    folderPath = ""

    On Error GoTo ErrorHandler

    With Application.FileDialog(msoFileDialogFolderPicker)
        .Title = "Select Source Folder Containing Cockpit Excel Files"
        .AllowMultiSelect = False
        .InitialFileName = ""  ' Start from default location

        If .Show = -1 Then  ' User clicked OK
            folderPath = .SelectedItems(1)
        Else
            ' User cancelled - return empty string
            folderPath = ""
        End If
    End With

    SelectSourceFolder = folderPath
    Exit Function

ErrorHandler:
    ' Fallback for systems where FileDialog doesn't work
    MsgBox "Folder picker dialog failed. Please ensure you have the necessary permissions.", vbExclamation, "Error"
    SelectSourceFolder = ""
End Function

Private Function SelectSourceFiles() As Variant
    ' Returns a 1-based array of full file paths, or Empty when cancelled.
    Dim arr() As String
    Dim i As Long

    On Error GoTo FallbackFolderMode

    With Application.FileDialog(msoFileDialogFilePicker)
        .Title = "Select Cockpit Excel File(s)"
        .AllowMultiSelect = True
        On Error Resume Next
        .Filters.Clear
        .Filters.Add "Excel Files", "*.xls;*.xlsx;*.xlsm;*.xlsb"
        On Error GoTo FallbackFolderMode

        If .Show = -1 Then
            If .SelectedItems.count = 0 Then
                SelectSourceFiles = Empty
                Exit Function
            End If

            ReDim arr(1 To .SelectedItems.count)
            For i = 1 To .SelectedItems.count
                arr(i) = CStr(.SelectedItems(i))
            Next i
            SelectSourceFiles = arr
        Else
            SelectSourceFiles = Empty
        End If
    End With
    Exit Function

FallbackFolderMode:
    On Error GoTo FallbackFailed
    Dim folderPath As String
    Dim f As String
    Dim count As Long

    folderPath = SelectSourceFolder()
    If Len(folderPath) = 0 Then
        SelectSourceFiles = Empty
        Exit Function
    End If

    folderPath = EnsureTrailingPathSeparator(folderPath)
    f = Dir(folderPath & "*.xls*")
    count = 0

    Do While Len(f) > 0
        If Left$(f, 2) <> "~$" Then
            count = count + 1
            If count = 1 Then
                ReDim arr(1 To 1)
            Else
                ReDim Preserve arr(1 To count)
            End If
            arr(count) = folderPath & f
        End If
        f = Dir()
    Loop

    If count = 0 Then
        MsgBox "No Excel files found in selected folder.", vbInformation
        SelectSourceFiles = Empty
    Else
        SelectSourceFiles = arr
    End If
    Exit Function

FallbackFailed:
    On Error GoTo 0
    MsgBox "File picker is not available in this Excel environment.", vbExclamation
    SelectSourceFiles = Empty
End Function

Private Function FindHeaderCol(ByVal ws As Worksheet, ByVal headerRow As Long, ByVal headerText As String) As Long
    Dim lastCol As Long, c As Long, v As String
    lastCol = ws.Cells(headerRow, ws.Columns.count).End(xlToLeft).Column
    For c = 1 To lastCol
        v = CleanText(CStr(ws.Cells(headerRow, c).value))
        If StrComp(v, CleanText(headerText), vbTextCompare) = 0 Then
            FindHeaderCol = c
            Exit Function
        End If
    Next c
    FindHeaderCol = 0
End Function

Private Function FindHeaderByText(ByVal ws As Worksheet, ByVal txt As String) As Long
    Dim lastCol As Long, c As Long
    lastCol = ws.Cells(1, ws.Columns.count).End(xlToLeft).Column
    For c = 1 To lastCol
        If StrComp(CleanText(CStr(ws.Cells(1, c).value)), txt, vbTextCompare) = 0 Then
            FindHeaderByText = c
            Exit Function
        End If
    Next c
    FindHeaderByText = 0
End Function

Private Function CleanText(ByVal s As String) As String
    Dim t As String
    t = Trim$(s)
    t = Replace(t, Chr$(160), " ")
    t = Replace(t, vbTab, " ")
    t = WorksheetFunction.Trim(t)
    CleanText = t
End Function

Private Function NormalizeSubjectName(ByVal raw As String) As String
    ' FSBB-safe: do NOT strip "- O", "- G1/G2/G3", "IP", etc.
    NormalizeSubjectName = CleanText(raw)
End Function

Private Function MakeLevelKey(ByVal lvlStr As String) As String
    ' Fallback: turn "Sec 3" -> "S3", else letters+digits
    Dim s As String
    s = UCase$(CleanText(lvlStr))   ' e.g. "SEC 3"
    
    Dim i As Long, ch As String, numPart As String
    numPart = ""
    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        If ch >= "0" And ch <= "9" Then
            numPart = numPart & ch
        End If
    Next i
    
    If Len(numPart) > 0 Then
        MakeLevelKey = "S" & numPart   ' "Sec 3" -> "S3"
    ElseIf Len(s) > 0 Then
        Dim letters As String
        letters = ""
        For i = 1 To Len(s)
            ch = Mid$(s, i, 1)
            If (ch >= "A" And ch <= "Z") Then letters = letters & ch
        Next i
        If Len(letters) = 0 Then
            MakeLevelKey = "Lv0"
        Else
            MakeLevelKey = letters
        End If
    Else
        MakeLevelKey = "Lv0"
    End If
End Function

Private Function GetLevelKeyFromClass(ByVal klass As String, ByVal lvlStr As String) As String
    ' 1) Try Settings mapping (ClassPattern -> LevelKey)
    ' 2) If no match, fallback to MakeLevelKey(lvlStr)
    Dim clsKey As String
    Dim entry As Variant
    Dim p As Long
    Dim pat As String
    Dim lvl As String
    
    clsKey = UCase$(Replace(CleanText(klass), " ", ""))
    
    If Not gClassLevelMap Is Nothing Then
        For Each entry In gClassLevelMap
            p = InStr(1, CStr(entry), vbTab, vbBinaryCompare)
            If p > 0 Then
                pat = Left$(CStr(entry), p - 1)
                lvl = Mid$(CStr(entry), p + 1)
            Else
                pat = CStr(entry)
                lvl = ""
            End If
            If Len(pat) > 0 And Left$(clsKey, Len(pat)) = pat Then
                GetLevelKeyFromClass = lvl
                Exit Function
            End If
        Next entry
    End If
    
    ' Fallback
    GetLevelKeyFromClass = MakeLevelKey(lvlStr)
End Function

Private Function MakeAssessKey(ByVal assessToken As String) As String
    Dim s As String
    s = UCase$(CleanText(assessToken))
    If Len(s) = 0 Then
        MakeAssessKey = "Unknown"
        Exit Function
    End If
    
    Dim i As Long, ch As String, out As String
    out = ""
    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        If (ch >= "A" And ch <= "Z") Or (ch >= "0" And ch <= "9") Then
            out = out & ch
        End If
    Next i
    If Len(out) = 0 Then
        MakeAssessKey = "Assess"
    Else
        MakeAssessKey = out
    End If
End Function

Private Function IsLevelAssessYearName(ByVal nm As String) As Boolean
    ' Pattern for prefix-blank mode:
    '   <Letters><Digits>_<token>_<4-digit year>
    ' e.g. "S3_WA1_2025", "Y3_WA2_2025"
    Dim parts() As String
    If InStr(nm, "_") = 0 Then Exit Function
    
    parts = Split(nm, "_")
    If UBound(parts) <> 2 Then Exit Function   ' need exactly 3 parts
    
    Dim p0 As String, p2 As String
    Dim i As Long, ch As String
    
    p0 = parts(0)
    p2 = parts(2)
    
    ' First part: letters + digits (first char letter, rest digits)
    If Len(p0) < 2 Then Exit Function
    ch = Mid$(p0, 1, 1)
    If ch < "A" Or ch > "Z" Then Exit Function
    For i = 2 To Len(p0)
        ch = Mid$(p0, i, 1)
        If ch < "0" Or ch > "9" Then Exit Function
    Next i
    
    ' Last part: 4-digit year
    If Len(p2) <> 4 Then Exit Function
    For i = 1 To 4
        ch = Mid$(p2, i, 1)
        If ch < "0" Or ch > "9" Then Exit Function
    Next i
    
    IsLevelAssessYearName = True
End Function

Private Function NzStr(ByVal v As Variant, ByVal fallback As String) As String
    If IsError(v) Then
        NzStr = fallback
    ElseIf Len(Trim$(CStr(v))) = 0 Then
        NzStr = fallback
    Else
        NzStr = CStr(v)
    End If
End Function

Private Function NzLng(ByVal v As Variant, ByVal fallback As Long) As Long
    On Error GoTo Bad
    If Len(Trim$(CStr(v))) = 0 Then
        NzLng = fallback
    Else
        NzLng = CLng(v)
    End If
    Exit Function
Bad:
    NzLng = fallback
End Function

Private Function GetFileNameFromPath(ByVal fullPath As String) As String
    Dim pSlash As Long, pBack As Long, pColon As Long, p As Long
    pSlash = InStrRev(fullPath, "/")
    pBack = InStrRev(fullPath, "\")
    pColon = InStrRev(fullPath, ":")

    p = pSlash
    If pBack > p Then p = pBack
    If pColon > p Then p = pColon

    If p > 0 Then
        GetFileNameFromPath = Mid$(fullPath, p + 1)
    Else
        GetFileNameFromPath = fullPath
    End If
End Function

Private Function EnsureTrailingPathSeparator(ByVal folderPath As String) As String
    Dim s As String
    s = Trim$(folderPath)
    If Len(s) = 0 Then
        EnsureTrailingPathSeparator = ""
        Exit Function
    End If

    Dim lastCh As String
    lastCh = Right$(s, 1)
    If lastCh = "/" Or lastCh = "\" Or lastCh = ":" Then
        EnsureTrailingPathSeparator = s
        Exit Function
    End If

    If InStr(1, s, "/", vbBinaryCompare) > 0 Then
        EnsureTrailingPathSeparator = s & "/"
    ElseIf InStr(1, s, ":", vbBinaryCompare) > 0 Then
        EnsureTrailingPathSeparator = s & ":"
    Else
        EnsureTrailingPathSeparator = s & "\"
    End If
End Function

'========================
' PARSERS
'========================
Private Function ParseYearFromLine(ByVal line As String) As String
    Dim s As String
    Dim i As Long
    Dim cand As String
    s = CleanText(line)

    For i = 1 To Len(s) - 3
        cand = Mid$(s, i, 4)
        If IsNumeric(cand) Then
            If Left$(cand, 2) = "20" Then
                ParseYearFromLine = cand
                Exit Function
            End If
        End If
    Next i
End Function

Private Function ParseAssessmentFromLine(ByVal line As String) As String
    Dim p As Long, afterColon As String
    p = InStr(1, line, ":")
    If p > 0 Then
        afterColon = CleanText(Mid$(line, p + 1))
        If Len(afterColon) > 0 Then
            ParseAssessmentFromLine = afterColon
            Exit Function
        End If
    End If
    
    Dim u As String
    u = UCase$(CleanText(line))

    If InStr(1, u, "WA1", vbBinaryCompare) > 0 Then
        ParseAssessmentFromLine = "WA1"
    ElseIf InStr(1, u, "WA2", vbBinaryCompare) > 0 Then
        ParseAssessmentFromLine = "WA2"
    ElseIf InStr(1, u, "WA3", vbBinaryCompare) > 0 Then
        ParseAssessmentFromLine = "WA3"
    ElseIf InStr(1, u, "EYE", vbBinaryCompare) > 0 Then
        ParseAssessmentFromLine = "EYE"
    ElseIf InStr(1, u, "MID YEAR", vbBinaryCompare) > 0 Or InStr(1, u, "MID-YEAR", vbBinaryCompare) > 0 Then
        ParseAssessmentFromLine = "Mid Year"
    ElseIf InStr(1, u, "PRELIM", vbBinaryCompare) > 0 Then
        ParseAssessmentFromLine = "Prelim"
    End If
End Function

Private Function InferLevelFromClass(ByVal klass As String) As String
    Dim s As String
    Dim i As Long
    Dim ch As String

    s = UCase$(CleanText(klass))
    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        If ch >= "1" And ch <= "5" Then
            InferLevelFromClass = "Sec " & ch
            Exit Function
        End If
    Next i
End Function

Private Function InferStreamFromClass(ByVal klass As String) As String
    ' Currently unused, kept for future if needed
    Dim u As String
    u = UCase$(klass)
    If InStr(u, " IP") > 0 Or Right$(u, 2) = "IP" Or InStr(u, "IP") > 0 Then
        InferStreamFromClass = "IP"
    ElseIf InStr(u, "SEC") > 0 Then
        InferStreamFromClass = "SEC"
    End If
End Function

'========================
' GRADE ? POINT (Helper for future analysis, uses Settings!H2:I50)
'========================
Private Sub EnsureGradeMapLoaded()
    If gGradeMapLoaded Then Exit Sub
    
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("Settings")
    On Error GoTo 0
    If ws Is Nothing Then
        Set gGradeMap = Nothing
        gGradeMapLoaded = True
        Exit Sub
    End If
    
    Dim rng As Range, r As Range
    On Error Resume Next
    Set rng = ws.Range("H2:I50")
    On Error GoTo 0
    
    Set gGradeMap = New Collection
    
    If Not rng Is Nothing Then
        For Each r In rng.Rows
            Dim g As String, pt As Variant
            g = UCase$(CleanText(CStr(r.Cells(1, 1).value)))   ' H: Grade
            pt = r.Cells(1, 2).value                           ' I: Point
            If Len(g) > 0 And Not IsEmpty(pt) And Not IsNull(pt) Then
                If Not MapExists(gGradeMap, g) Then MapSet gGradeMap, g, pt
            End If
        Next r
    End If
    
    gGradeMapLoaded = True
End Sub

Public Function GradeToPoint(ByVal grade As String) As Variant
    ' Returns the mapped point, or Empty if not found.
    EnsureGradeMapLoaded
    If gGradeMap Is Nothing Then
        GradeToPoint = Empty
        Exit Function
    End If
    
    Dim g As String
    g = UCase$(CleanText(grade))
    
    If MapExists(gGradeMap, g) Then
        GradeToPoint = MapGetVariant(gGradeMap, g)
    Else
        GradeToPoint = Empty
    End If
End Function

'========================
' LOGGING (optional)
'========================
Private Sub AppendLog(ByVal statusText As String, ByVal filePath As String, ByVal notes As String)
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("Logs")
    On Error GoTo 0
    If ws Is Nothing Then Exit Sub
    
    Dim nr As Long
    nr = ws.Cells(ws.Rows.count, 1).End(xlUp).Row + 1
    ws.Cells(nr, 1).value = Now
    ws.Cells(nr, 2).value = GetFileNameFromPath(filePath)
    ws.Cells(nr, 3).value = statusText
    ws.Cells(nr, 4).value = notes
End Sub


