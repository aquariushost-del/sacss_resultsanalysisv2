Attribute VB_Name = "modIpSubjectNavigation"
Option Explicit

'=========================================================
' Module: modIpSubjectNavigation
'
' PURPOSE:
'   Build an IP navigation panel on "Dashboard" with
'   rounded-rectangle buttons linking to each IP Subject
'   Analysis sheet, ordered Y1 Ð Y4.
'
' ASSUMPTIONS:
'   - IP analysis sheet names start with:
'       Y1_Subj Analysis_...
'       Y2_Subj Analysis_...
'       Y3_Subj Analysis_...
'       Y4_Subj Analysis_...
'
'   - There is a sheet called "Dashboard".
'   - Navigation starts from Dashboard!M3.
'
' OUTPUT (example):
'   Y1 Subject Analysis (IP)
'     [Y1_Subj Analysis_Y1_WA1_2025]
'     [Y1_Subj Analysis_Y1_WA2_2025]
'
'   Y2 Subject Analysis (IP)
'     ...
'
' HOME BUTTONS (IP only):
'   - Each IP Subject Analysis sheet (Y1ÐY4) gets a small
'     rounded "Home" button at P1 linking back to Dashboard.
'=========================================================

Private Const NAV_SHEET_NAME As String = "Dashboard"
Private Const IP_NAV_START_CELL As String = "M3"
Private Const IP_NAV_BTN_PREFIX As String = "Nav_IP_"

Private Const SHAPE_ROUNDED_RECTANGLE As Long = 5   ' msoShapeRoundedRectangle

'---------------------------------------------------------
' PUBLIC ENTRY POINT
'   Run this to build IP navigation & IP home buttons
'---------------------------------------------------------
Public Sub BuildIpSubjectAnalysisNavigation()
    Dim wb As Workbook
    Dim wsNav As Worksheet
    
    Set wb = ThisWorkbook
    
    '--- Get navigation sheet
    On Error Resume Next
    Set wsNav = wb.Worksheets(NAV_SHEET_NAME)
    On Error GoTo 0
    
    If wsNav Is Nothing Then
        MsgBox "Navigation sheet '" & NAV_SHEET_NAME & "' not found.", vbExclamation
        Exit Sub
    End If
    
    '--- Build IP navigation (Y1ÐY4) in purple, from M3
    BuildIpNavigation wsNav
    
    '--- Build / refresh HOME buttons only on IP Subject Analysis sheets
    BuildAllIpHomeButtons
    
    wsNav.Activate
    wsNav.Range(IP_NAV_START_CELL).Select
    'MsgBox "IP Subject Analysis navigation updated on '" & NAV_SHEET_NAME & "'.", vbInformation
End Sub

'=========================================================
' IP NAVIGATION (Y1ÐY4, purple, starts at M3)
'=========================================================
Private Sub BuildIpNavigation(ByVal wsNav As Worksheet)
    Dim wb As Workbook
    Dim startCell As Range
    Dim startRow As Long, startCol As Long
    
    Dim levelDict As Object
    Dim lvl As Variant
    Dim ws As Worksheet
    
    Dim arrNames() As String
    Dim i As Long
    
    Set wb = ThisWorkbook
    Set startCell = wsNav.Range(IP_NAV_START_CELL)
    startRow = startCell.Row
    startCol = startCell.Column
    
    '--- Collect IP analysis sheets by level (Y1..Y4)
    Set levelDict = CreateObject("Scripting.Dictionary")
    
    For Each lvl In Array("Y1", "Y2", "Y3", "Y4")
        Dim tmpCol As Collection
        Set tmpCol = New Collection
        levelDict.Add CStr(lvl), tmpCol
    Next lvl
    
    For Each ws In wb.Worksheets
        Dim levelTag As String
        levelTag = Left$(ws.Name, 2)  ' "Y1", "Y2", etc.
        
        If levelDict.Exists(levelTag) Then
            ' Check that this is a Subject Analysis sheet
            If InStr(1, ws.Name, "_Subj Analysis_", vbTextCompare) > 0 Then
                levelDict(levelTag).Add ws.Name
            End If
        End If
    Next ws
    
    '--- Clear previous IP nav area & old IP buttons
    ClearOldIpNavigation wsNav, startRow, startCol, IP_NAV_BTN_PREFIX, 6
    
    '--- Build IP navigation by level, always in order Y1ÐY4
    Dim curRow As Long
    curRow = startRow
    
    For Each lvl In Array("Y1", "Y2", "Y3", "Y4")
        Dim colNames As Collection
        Set colNames = levelDict(CStr(lvl))
        
        ' Header row for this IP level
        wsNav.Cells(curRow, startCol).value = lvl & " Subject Analysis (IP)"
        With wsNav.Cells(curRow, startCol)
            .Font.Bold = True
            .Font.Size = 12
        End With
        curRow = curRow + 1
        
        If colNames.count = 0 Then
            ' No analysis yet Ð optional message
            wsNav.Cells(curRow, startCol).value = "(No IP subject analysis sheets found.)"
            wsNav.Cells(curRow, startCol).Font.Italic = True
            curRow = curRow + 2
        Else
            ' Copy collection to array for sorting
            ReDim arrNames(1 To colNames.count)
            For i = 1 To colNames.count
                arrNames(i) = CStr(colNames(i))
            Next i
            SortStringArray arrNames
            
            ' Create one purple button per IP sheet
            For i = LBound(arrNames) To UBound(arrNames)
                CreateIpNavButton wsNav, _
                                  arrNames(i), _
                                  curRow, _
                                  startCol, _
                                  IP_NAV_BTN_PREFIX, _
                                  RGB(112, 48, 160), _
                                  RGB(74, 38, 115)
                curRow = curRow + 2   ' add space between buttons
            Next i
            
            curRow = curRow + 1      ' extra blank line between levels
        End If
    Next lvl
End Sub

'=========================================================
' CORE HELPERS (IP)
'=========================================================

'---------------------------------------------------------
' Clear old IP nav shapes and text around start cell
'   - widthCols: how many columns to clear from startCol
'---------------------------------------------------------
Private Sub ClearOldIpNavigation(ByVal wsNav As Worksheet, _
                                 ByVal startRow As Long, _
                                 ByVal startCol As Long, _
                                 ByVal shapePrefix As String, _
                                 ByVal widthCols As Long)
    Dim shp As Shape
    Dim lastRow As Long, lastCol As Long
    Dim k As Long
    
    ' Define a reasonable area to clear (e.g. 200 rows)
    lastRow = startRow + 200
    lastCol = startCol + widthCols - 1
    
    With wsNav.Range(wsNav.Cells(startRow, startCol), wsNav.Cells(lastRow, lastCol))
        .ClearContents
        .ClearFormats
    End With
    
    ' Delete previous IP navigation buttons for this block (by name prefix)
    For k = wsNav.Shapes.count To 1 Step -1
        Set shp = wsNav.Shapes(k)
        If Left$(shp.Name, Len(shapePrefix)) = shapePrefix Then
            shp.Delete
        End If
    Next k
End Sub

'---------------------------------------------------------
' Create a rounded-rectangle IP button linking to a sheet
' rowNum / firstCol control its position
'
' colourFillRGB / colourLineRGB control the palette
'---------------------------------------------------------
Private Sub CreateIpNavButton(ByVal wsNav As Worksheet, _
                              ByVal sheetName As String, _
                              ByVal rowNum As Long, _
                              ByVal firstCol As Long, _
                              ByVal shapePrefix As String, _
                              ByVal colourFillRGB As Long, _
                              ByVal colourLineRGB As Long)
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
    
    ' Text on button Ð currently full sheet name
    displayText = sheetName
    ' If you prefer shorter labels, you can tweak, e.g.:
    ' displayText = Replace(sheetName, "_Subj Analysis_", " Ð ")
    
    Set shp = wsNav.Shapes.AddShape( _
        Type:=SHAPE_ROUNDED_RECTANGLE, _
        Left:=leftPos, _
        Top:=topPos, _
        Width:=btnWidth, _
        Height:=btnHeight)
    
    With shp
        .Name = shapePrefix & sheetName
        
        ' Fill & border (purple theme)
        .Fill.ForeColor.RGB = colourFillRGB
        .Fill.Transparency = 0#
        .line.ForeColor.RGB = colourLineRGB
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
' IP HOME BUTTON ENGINE Ð adds Home button to all
' Y1/Y2/Y3/Y4 Subject Analysis sheets
'=========================================================
Public Sub BuildAllIpHomeButtons()
    Dim wb As Workbook
    Dim ws As Worksheet
    
    Set wb = ThisWorkbook
    
    For Each ws In wb.Worksheets
        If IsIpSubjectAnalysisSheet(ws.Name) Then
            AddIpHomeButton ws
        End If
    Next ws
End Sub

'---------------------------------------------------------
' Detect IP Subject Analysis sheets only
'   - Y1..Y4_Subj Analysis_...
'---------------------------------------------------------
Private Function IsIpSubjectAnalysisSheet(ByVal nm As String) As Boolean
    Dim lvl As String
    lvl = Left$(nm, 2)   ' Y1, Y2, Y3, Y4
    
    If (lvl = "Y1" Or lvl = "Y2" Or lvl = "Y3" Or lvl = "Y4") _
       And InStr(1, nm, "_Subj Analysis_", vbTextCompare) > 0 Then
        IsIpSubjectAnalysisSheet = True
    Else
        IsIpSubjectAnalysisSheet = False
    End If
End Function

'---------------------------------------------------------
' Create HOME button at P1 on a specific IP sheet
'---------------------------------------------------------
Private Sub AddIpHomeButton(ws As Worksheet)
    Dim shp As Shape
    Dim tgtCell As Range
    Dim leftPos As Double, topPos As Double
    Dim btnWidth As Double, btnHeight As Double
    
    ' Target location: P1
    Set tgtCell = ws.Range("P1")
    leftPos = tgtCell.Left
    topPos = tgtCell.Top
    btnWidth = tgtCell.Width * 1.2
    btnHeight = tgtCell.Height * 1.2
    
    ' Remove existing Home button if any (IP)
    On Error Resume Next
    ws.Shapes("HomeBtn_IP").Delete
    On Error GoTo 0
    
    ' Build shape
    Set shp = ws.Shapes.AddShape( _
        Type:=SHAPE_ROUNDED_RECTANGLE, _
        Left:=leftPos, _
        Top:=topPos, _
        Width:=btnWidth, _
        Height:=btnHeight)
    
    With shp
        .Name = "HomeBtn_IP"
        
        ' Visual style (purple to match IP nav)
        .Fill.ForeColor.RGB = RGB(112, 48, 160)
        .line.ForeColor.RGB = RGB(74, 38, 115)
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


