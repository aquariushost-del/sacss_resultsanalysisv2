Attribute VB_Name = "modSubjectNavigation"
Option Explicit

'=========================================================
' Module: modSubjectNavigation
'
' PURPOSE:
'   Build a navigation panel on "Dashboard" with rounded-
'   rectangle buttons linking to each SEC Subject Analysis
'   sheet, ordered S1 � S4.
'
' ASSUMPTIONS:
'   - Analysis sheet names start with:
'       S1_Subj Analysis_...
'       S2_Subj Analysis_...
'       S3_Subj Analysis_...
'       S4_Subj Analysis_...
'   - There is a sheet called "Dashboard".
'   - Navigation starts from Dashboard!G3.
'
' OUTPUT (example):
'   S1 Subject Analysis
'     [S1_Subj Analysis_S1_WA1_2025]
'     [S1_Subj Analysis_S1_WA2_2025]
'
'   S2 Subject Analysis
'     [S2_Subj Analysis_S2_WA1_2025]
'     ...
'=========================================================

Private Const NAV_SHEET_NAME As String = "Dashboard"
Private Const NAV_START_CELL As String = "G3"
Private Const NAV_BTN_PREFIX As String = "Nav_Subj_"
Private Const SHAPE_ROUNDED_RECTANGLE As Long = 5   ' msoShapeRoundedRectangle

'---------------------------------------------------------
' PUBLIC ENTRY POINT
'---------------------------------------------------------
Public Sub BuildSubjectAnalysisNavigation()
    Dim wb As Workbook
    Dim wsNav As Worksheet
    Dim startCell As Range
    Dim startRow As Long, startCol As Long
    
    Dim levelDict As Object
    Dim lvl As Variant
    Dim ws As Worksheet
    
    Dim arrNames() As String
    Dim i As Long
    
    Set wb = ThisWorkbook
    
    '--- Get navigation sheet
    On Error Resume Next
    Set wsNav = wb.Worksheets(NAV_SHEET_NAME)
    On Error GoTo 0
    
    If wsNav Is Nothing Then
        MsgBox "Navigation sheet '" & NAV_SHEET_NAME & "' not found.", vbExclamation
        Exit Sub
    End If
    
    Set startCell = wsNav.Range(NAV_START_CELL)
    startRow = startCell.Row
    startCol = startCell.Column
    
    '--- Collect analysis sheets by level (S1..S5)
    Set levelDict = CreateObject("Scripting.Dictionary")
    
    For Each lvl In Array("S1", "S2", "S3", "S4", "S5")
        Dim tmpCol As Collection
        Set tmpCol = New Collection
        levelDict.Add CStr(lvl), tmpCol
    Next lvl
    
    For Each ws In wb.Worksheets
        Dim levelTag As String
        levelTag = Left$(ws.Name, 2)  ' "S1", "S2", etc.
        
        If levelDict.Exists(levelTag) Then
            ' Check that this is a Subject Analysis sheet
            If InStr(1, ws.Name, "_Subj Analysis_", vbTextCompare) > 0 Then
                levelDict(levelTag).Add ws.Name
            End If
        End If
    Next ws
    
    '--- Clear previous nav area & old buttons
    ClearOldNavigation wsNav, startRow, startCol
    
    '--- Build navigation by level, always in order S1�S4
    Dim curRow As Long
    curRow = startRow
    
    For Each lvl In Array("S1", "S2", "S3", "S4", "S5")
        Dim colNames As Collection
        Set colNames = levelDict(CStr(lvl))
        
        ' Header row for this level
        wsNav.Cells(curRow, startCol).value = lvl & " Subject Analysis"
        With wsNav.Cells(curRow, startCol)
            .Font.Bold = True
            .Font.Size = 12
        End With
        curRow = curRow + 1
        
        If colNames.count = 0 Then
            ' No analysis yet � optional message
            wsNav.Cells(curRow, startCol).value = "(No subject analysis sheets found.)"
            wsNav.Cells(curRow, startCol).Font.Italic = True
            curRow = curRow + 2
        Else
            ' Copy collection to array for sorting
            ReDim arrNames(1 To colNames.count)
            For i = 1 To colNames.count
                arrNames(i) = CStr(colNames(i))
            Next i
            SortStringArray arrNames
            
            ' Create one button per sheet
            For i = LBound(arrNames) To UBound(arrNames)
                CreateNavButton wsNav, arrNames(i), curRow, startCol
                curRow = curRow + 2   ' add space between buttons
            Next i
            
            curRow = curRow + 1      ' extra blank line between levels
        End If
    Next lvl
    
    ' Build / refresh HOME buttons on all Subject Analysis sheets
    BuildAllHomeButtons
    
    wsNav.Activate
    wsNav.Range(NAV_START_CELL).Select
    'MsgBox "Subject Analysis navigation updated on '" & NAV_SHEET_NAME & "'.", vbInformation
End Sub

'---------------------------------------------------------
' Clear old nav shapes and text around start cell
'---------------------------------------------------------
Private Sub ClearOldNavigation(ByVal wsNav As Worksheet, _
                               ByVal startRow As Long, _
                               ByVal startCol As Long)
    Dim shp As Shape
    Dim lastRow As Long, lastCol As Long
    Dim k As Long
    
    ' Define a reasonable area to clear (e.g. 200 rows, 6 columns)
    lastRow = startRow + 200
    lastCol = startCol + 5
    
    With wsNav.Range(wsNav.Cells(startRow, startCol), wsNav.Cells(lastRow, lastCol))
        .ClearContents
        .ClearFormats
    End With
    
    ' Delete previous navigation buttons (by name prefix)
    For k = wsNav.Shapes.count To 1 Step -1
        Set shp = wsNav.Shapes(k)
        If Left$(shp.Name, Len(NAV_BTN_PREFIX)) = NAV_BTN_PREFIX Then
            shp.Delete
        End If
    Next k
End Sub

'---------------------------------------------------------
' Create a rounded-rectangle button linking to a sheet
' rowNum / firstCol control its position
'---------------------------------------------------------
Private Sub CreateNavButton(ByVal wsNav As Worksheet, _
                            ByVal sheetName As String, _
                            ByVal rowNum As Long, _
                            ByVal firstCol As Long)
    Dim wb As Workbook
    Dim shp As Shape
    Dim leftPos As Double, topPos As Double
    Dim btnWidth As Double, btnHeight As Double
    Dim displayText As String
    
    Set wb = ThisWorkbook
    
    ' Position: based on cell(rowNum, firstCol)
    leftPos = wsNav.Cells(rowNum, firstCol).Left
    topPos = wsNav.Cells(rowNum, firstCol).Top
    
    ' Button spans 5 columns, slightly taller than the row
    btnWidth = wsNav.Columns(firstCol).Resize(, 5).Width
    btnHeight = wsNav.Rows(rowNum).Height * 1.3
    
    ' Text on button � currently full sheet name
    displayText = sheetName
    ' If you prefer shorter labels, you can tweak, e.g.:
    ' displayText = Replace(sheetName, "_Subj Analysis_", " � ")
    
    Set shp = wsNav.Shapes.AddShape( _
        Type:=SHAPE_ROUNDED_RECTANGLE, _
        Left:=leftPos, _
        Top:=topPos, _
        Width:=btnWidth, _
        Height:=btnHeight)
    
    With shp
        .Name = NAV_BTN_PREFIX & sheetName
        
        ' Fill & border
        .Fill.ForeColor.RGB = RGB(79, 129, 189)   ' soft blue
        .Fill.Transparency = 0#
        .line.ForeColor.RGB = RGB(55, 86, 130)
        .line.Weight = 1.5
        
        ' Text formatting
        With .TextFrame2
            .TextRange.text = displayText
            .TextRange.Font.Name = "Calibri"
            .TextRange.Font.Size = 10.5
            .TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
            .TextRange.ParagraphFormat.Alignment = msoAlignCenter
            .VerticalAnchor = msoAnchorMiddle
            .MarginLeft = 6
            .MarginRight = 6
            .MarginTop = 3
            .MarginBottom = 3
        End With
    End With
    
    ' Add hyperlink to the sheet (internal link) via Hyperlinks.Add
    wsNav.Hyperlinks.Add Anchor:=shp, _
                          Address:="", _
                          SubAddress:="'" & sheetName & "'!A1"
End Sub

'---------------------------------------------------------
' Simple alphabetical sort for String array (1-based)
'---------------------------------------------------------
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

'=========================================================
' HOME BUTTON ENGINE � adds ?? Home button to all
' S1/S2/S3/S4/S5 Subject Analysis sheets
'=========================================================
Public Sub BuildAllHomeButtons()
    Dim wb As Workbook
    Dim ws As Worksheet
    
    Set wb = ThisWorkbook
    
    For Each ws In wb.Worksheets
        If IsSubjectAnalysisSheet(ws.Name) Then
            AddHomeButton ws
        End If
    Next ws
End Sub

'---------------------------------------------------------
' Detect Subject Analysis sheets
'---------------------------------------------------------
Private Function IsSubjectAnalysisSheet(ByVal nm As String) As Boolean
    Dim lvl As String
    lvl = Left$(nm, 2)   ' S1, S2, S3, S4
    
    If (lvl = "S1" Or lvl = "S2" Or lvl = "S3" Or lvl = "S4" Or lvl = "S5") _
       And InStr(1, nm, "_Subj Analysis_", vbTextCompare) > 0 Then
        IsSubjectAnalysisSheet = True
    Else
        IsSubjectAnalysisSheet = False
    End If
End Function

'---------------------------------------------------------
' Create HOME button at N1 on a specific sheet
'---------------------------------------------------------
Private Sub AddHomeButton(ws As Worksheet)
    Dim shp As Shape
    Dim tgtCell As Range
    Dim leftPos As Double, topPos As Double
    Dim btnWidth As Double, btnHeight As Double
    
    ' Target location: N1
    Set tgtCell = ws.Range("N1")
    leftPos = tgtCell.Left
    topPos = tgtCell.Top
    btnWidth = tgtCell.Width * 1.2
    btnHeight = tgtCell.Height * 1.2
    
    ' Remove existing Home button if any
    On Error Resume Next
    ws.Shapes("HomeBtn").Delete
    On Error GoTo 0
    
    ' Build shape
    Set shp = ws.Shapes.AddShape( _
        Type:=SHAPE_ROUNDED_RECTANGLE, _
        Left:=leftPos, _
        Top:=topPos, _
        Width:=btnWidth, _
        Height:=btnHeight)
    
    With shp
        .Name = "HomeBtn"
        
        ' Visual style (same palette as nav buttons)
        .Fill.ForeColor.RGB = RGB(79, 129, 189)
        .line.ForeColor.RGB = RGB(55, 86, 130)
        .line.Weight = 1.5
        
        With .TextFrame2
            .TextRange.text = "Home"
            .TextRange.Font.Name = "Calibri"
            .TextRange.Font.Size = 11
            .TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
            .VerticalAnchor = msoAnchorMiddle
            .TextRange.ParagraphFormat.Alignment = msoAlignCenter
            .MarginLeft = 4
            .MarginRight = 4
        End With
    End With
    
    ' Hyperlink to Dashboard
    ws.Hyperlinks.Add Anchor:=shp, _
        Address:="", _
        SubAddress:="'Dashboard'!A1"
End Sub


