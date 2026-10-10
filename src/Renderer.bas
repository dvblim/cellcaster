Attribute VB_Name = "Renderer"
'==============================================================================
' Renderer: the worksheet is the display.
'
' Colouring 9,000 cells one by one every frame is far too slow. Instead:
'   1. one conditional-formatting rule per palette index is created once;
'   2. each frame is a single Range.Value2 write of a 2-D array of indices.
' Excel then repaints the fills itself, which is the fastest path available.
'==============================================================================
Option Explicit

Private mSheet As Worksheet
Private mScreen As Range
Private mHud As Range

' Creates (or reuses) the game sheet and turns a block of cells into a screen.
Public Sub Init()
    Dim updating As Boolean

    updating = Application.ScreenUpdating
    Application.ScreenUpdating = False

    Set mSheet = EnsureSheet()
    ThisWorkbook.Activate
    mSheet.Activate

    Set mScreen = mSheet.Cells(SCREEN_TOP, SCREEN_LEFT).Resize(SCREEN_H, SCREEN_W)
    Set mHud = mSheet.Cells(SCREEN_TOP - 1, SCREEN_LEFT)

    mSheet.Cells.Interior.Color = RGB(0, 0, 0)
    mSheet.Cells.Font.Color = RGB(230, 230, 230)
    mSheet.Cells.Font.Name = "Consolas"
    mSheet.Cells.Font.Size = 10

    SquareCells
    BuildPalette

    ' ";;;" hides the numbers, so only the fills are visible.
    mScreen.NumberFormat = ";;;"
    mScreen.Value2 = 0

    FitWindow
    Application.ScreenUpdating = updating
End Sub

Private Function EnsureSheet() As Worksheet
    Dim ws As Worksheet

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(SHEET_NAME)
    On Error GoTo 0

    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = SHEET_NAME
    End If
    Set EnsureSheet = ws
End Function

' Row height is set in points, but column width is measured in characters of
' the default font, so the width is corrected until the cells come out square.
Private Sub SquareCells()
    Dim width As Double
    Dim actual As Double
    Dim i As Long

    mSheet.Columns(1).ColumnWidth = 1
    mSheet.Rows(SCREEN_TOP - 1).RowHeight = 15
    mScreen.RowHeight = CELL_SIZE_PT

    width = 0.5
    For i = 1 To 6
        mScreen.ColumnWidth = width
        actual = mScreen.Columns(1).width
        If actual <= 0 Then
            width = width * 2
        ElseIf Abs(actual - CELL_SIZE_PT) < 0.05 Then
            Exit For
        Else
            width = width * CELL_SIZE_PT / actual
        End If
    Next i
End Sub

' One rule per palette index. The formulas are whole numbers on purpose:
' they read the same under any regional decimal separator.
Private Sub BuildPalette()
    Dim i As Long
    Dim rule As FormatCondition

    mScreen.FormatConditions.Delete
    For i = 1 To PALETTE_SIZE
        Set rule = mScreen.FormatConditions.Add(Type:=xlCellValue, Operator:=xlEqual, Formula1:="=" & i)
        rule.Interior.Color = PaletteColor(i)
        rule.StopIfTrue = True
    Next i
End Sub

' Zooms so the screen fills the window, and hides the spreadsheet chrome.
Private Sub FitWindow()
    On Error Resume Next
    ActiveWindow.DisplayGridlines = False
    ActiveWindow.DisplayHeadings = False
    mSheet.Range(mHud, mScreen).Select
    ActiveWindow.Zoom = True
    mSheet.Cells(1, 1).Select
    ActiveWindow.ScrollRow = 1
    ActiveWindow.ScrollColumn = 1
    On Error GoTo 0
End Sub

' Shows a frame. Errors are ignored: if the player double-clicks a cell,
' Excel enters edit mode and rejects writes until they leave it.
Public Sub Blit(ByRef frame() As Variant)
    On Error Resume Next
    mScreen.Value2 = frame
    On Error GoTo 0
End Sub

Public Sub Hud(ByVal text As String)
    On Error Resume Next
    mHud.Value2 = text
    On Error GoTo 0
End Sub

' False once the player switches sheet or closes the workbook.
Public Function IsActive() As Boolean
    On Error Resume Next
    IsActive = (ActiveSheet Is mSheet)
    On Error GoTo 0
End Function
