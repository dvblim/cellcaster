Attribute VB_Name = "Game"
'==============================================================================
' Game: entry points and the main loop.
'
'   StartGame    - play (also bound to the JOGAR button on the Menu sheet)
'   InstallMenu  - builds the Menu sheet; run once after importing the modules
'==============================================================================
Option Explicit

Private mRunning As Boolean
Private mShowMap As Boolean
Private mMapKeyWasDown As Boolean

Public Sub StartGame()
    Dim frame() As Variant
    Dim prevCalculation As Long
    Dim prevEvents As Boolean
    Dim haveState As Boolean
    Dim tStart As Double
    Dim tPrev As Double
    Dim tNow As Double
    Dim dt As Double
    Dim spare As Double
    Dim hudTime As Double
    Dim hudFrames As Long
    Dim fps As Double
    Dim tMark As Double
    Dim msDraw As Double
    Dim msWrite As Double
    Dim msPaint As Double
    Dim won As Boolean
    Dim failure As String

    #If Mac Then
        MsgBox "O CellCaster usa a API do Windows e funciona apenas no Excel para Windows.", vbInformation, "CellCaster"
        Exit Sub
    #End If

    If mRunning Then Exit Sub
    mRunning = True
    On Error GoTo Fail

    prevCalculation = Application.Calculation
    prevEvents = Application.EnableEvents
    haveState = True
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    Application.EnableCancelKey = xlErrorHandler   ' Esc raises error 18 instead of a dialog

    World.LoadLevel
    Textures.Build
    Raycaster.Init
    Renderer.Init
    ReDim frame(1 To SCREEN_H, 1 To SCREEN_W)

    mShowMap = True
    mMapKeyWasDown = False
    Platform.CaptureKeys True
    On Error Resume Next
    AppActivate Application.Caption     ' take focus when started from the VBA editor
    On Error GoTo Fail

    tStart = Platform.Clock()
    tPrev = tStart
    hudTime = tStart

    Do
        tNow = Platform.Clock()
        dt = tNow - tPrev
        tPrev = tNow
        If dt > 0.1 Then dt = 0.1      ' a long hitch must not teleport the player

        If Not Renderer.IsActive() Then Exit Do

        If Platform.HasFocus() Then
            If Platform.KeyDown(vbKeyEscape) Then Exit Do
            HandleInput dt
        End If

        If World.CollectGems() > 0 Then Beep

        ' The three stages are timed separately: the HUD shows where a frame goes.
        tMark = Platform.Clock()
        Raycaster.Render frame, tNow - tStart
        If mShowMap Then Raycaster.DrawMinimap frame
        msDraw = msDraw + (Platform.Clock() - tMark) * 1000#

        tMark = Platform.Clock()
        Renderer.Blit frame
        msWrite = msWrite + (Platform.Clock() - tMark) * 1000#

        tMark = Platform.Clock()
        DoEvents                        ' lets Excel repaint and stay responsive
        msPaint = msPaint + (Platform.Clock() - tMark) * 1000#

        hudFrames = hudFrames + 1
        If tNow - hudTime >= 0.5 Then
            fps = hudFrames / (tNow - hudTime)
            Renderer.Hud "CELLCASTER   gemas " & (World.GemCount - World.GemsLeft) & "/" & World.GemCount & _
                         "   tempo " & Format$(tNow - tStart, "0.0") & "s" & _
                         "   " & Format$(fps, "0") & " fps (desenho " & Format$(msDraw / hudFrames, "0") & _
                         " + escrita " & Format$(msWrite / hudFrames, "0") & _
                         " + pintura " & Format$(msPaint / hudFrames, "0") & " ms)" & _
                         "   |   WASD/setas: mover   M: mapa   Esc: sair"
            hudFrames = 0
            hudTime = tNow
            msDraw = 0#: msWrite = 0#: msPaint = 0#
        End If

        If World.GemsLeft = 0 Then
            won = True
            Exit Do
        End If

        ' Give the CPU back when the frame finished early.
        spare = 1# / TARGET_FPS - (Platform.Clock() - tNow)
        If spare > 0.002 Then Platform.Nap CLng(spare * 1000#)
    Loop

Cleanup:
    On Error Resume Next
    Platform.CaptureKeys False
    If haveState Then
        Application.Calculation = prevCalculation
        Application.EnableEvents = prevEvents
    End If
    Application.EnableCancelKey = xlInterrupt
    mRunning = False

    If Len(failure) > 0 Then
        MsgBox "O CellCaster parou: " & failure, vbExclamation, "CellCaster"
    ElseIf won Then
        Renderer.Hud "CELLCASTER   todas as " & World.GemCount & " gemas em " & Format$(tNow - tStart, "0.0") & "s - execute StartGame para jogar de novo"
        MsgBox "Todas as " & World.GemCount & " gemas coletadas em " & Format$(tNow - tStart, "0.0") & " segundos!", vbInformation, "CellCaster"
    Else
        Renderer.Hud "CELLCASTER   parado - execute StartGame para jogar de novo"
    End If
    Exit Sub

Fail:
    ' Error 18 is "user interrupt" (Esc or Ctrl+Break): a normal way to quit.
    If Err.Number <> 18 Then failure = Err.Description & " (erro " & Err.Number & ")"
    Resume Cleanup
End Sub

Private Sub HandleInput(ByVal dt As Double)
    Dim forward As Double
    Dim strafe As Double
    Dim turn As Double
    Dim dx As Double
    Dim dy As Double
    Dim length As Double
    Dim mapKeyDown As Boolean

    If Platform.KeyDown(vbKeyW) Or Platform.KeyDown(vbKeyUp) Then forward = forward + 1#
    If Platform.KeyDown(vbKeyS) Or Platform.KeyDown(vbKeyDown) Then forward = forward - 1#
    If Platform.KeyDown(vbKeyD) Then strafe = strafe + 1#
    If Platform.KeyDown(vbKeyA) Then strafe = strafe - 1#
    If Platform.KeyDown(vbKeyRight) Then turn = turn + 1#
    If Platform.KeyDown(vbKeyLeft) Then turn = turn - 1#

    If turn <> 0# Then World.Rotate turn * TURN_SPEED * dt

    ' The right-hand vector of the view is (-DirY, DirX).
    dx = World.DirX * forward - World.DirY * strafe
    dy = World.DirY * forward + World.DirX * strafe
    length = Sqr(dx * dx + dy * dy)
    If length > 0# Then
        World.TryMove dx / length * MOVE_SPEED * dt, dy / length * MOVE_SPEED * dt
    End If

    ' Toggle the minimap on the key press, not while it is held.
    mapKeyDown = Platform.KeyDown(vbKeyM)
    If mapKeyDown And Not mMapKeyWasDown Then mShowMap = Not mShowMap
    mMapKeyWasDown = mapKeyDown
End Sub

' Builds a "Menu" sheet with the instructions and a JOGAR (play) button.
Public Sub InstallMenu()
    Dim ws As Worksheet
    Dim button As Shape
    Dim i As Long

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("Menu")
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(Before:=ThisWorkbook.Worksheets(1))
        ws.Name = "Menu"
    End If

    For i = ws.Shapes.Count To 1 Step -1
        ws.Shapes(i).Delete
    Next i
    ws.Cells.Clear
    ws.Cells.Interior.Color = RGB(18, 18, 24)
    ws.Cells.Font.Color = RGB(230, 230, 230)
    ws.Cells.Font.Name = "Consolas"
    ws.Columns(1).ColumnWidth = 3

    With ws.Range("B2")
        .Value2 = "CELLCASTER"
        .Font.Size = 32
        .Font.Bold = True
        .Font.Color = RGB(250, 205, 60)
    End With
    ws.Range("B3").Value2 = "Um motor de raycasting com texturas, desenhado na planilha. VBA puro."
    ws.Range("B9").Value2 = "W / S ou setas cima / baixo    andar"
    ws.Range("B10").Value2 = "A / D                          andar de lado"
    ws.Range("B11").Value2 = "Setas esquerda / direita       girar"
    ws.Range("B12").Value2 = "M                              mostrar ou ocultar o minimapa"
    ws.Range("B13").Value2 = "Esc                            sair (segure por um segundo)"
    ws.Range("B15").Value2 = "Objetivo: colete todas as gemas no menor tempo."

    Set button = ws.Shapes.AddShape(5, ws.Range("B5").Left, ws.Range("B5").Top, 180, 44)   ' 5 = rounded rectangle
    button.Name = "PlayButton"
    button.Fill.ForeColor.RGB = RGB(190, 70, 55)
    button.Line.Visible = False
    button.TextFrame.Characters.text = "JOGAR"
    button.TextFrame.Characters.Font.Size = 20
    button.TextFrame.Characters.Font.Bold = True
    button.TextFrame.Characters.Font.Color = RGB(255, 255, 255)
    button.TextFrame.HorizontalAlignment = xlHAlignCenter
    button.TextFrame.VerticalAlignment = xlVAlignCenter
    button.OnAction = "StartGame"

    On Error Resume Next
    ws.Activate
    ActiveWindow.DisplayGridlines = False
    On Error GoTo 0
End Sub

' Times each stage of a frame and reports milliseconds. Run it from Alt+F8.
Public Sub Benchmark()
    Const N As Long = 20
    Dim frame() As Variant
    Dim i As Long
    Dim t0 As Double
    Dim msRender As Double
    Dim msWrite As Double
    Dim msPaint As Double

    World.LoadLevel
    Textures.Build
    Raycaster.Init
    Renderer.Init
    ReDim frame(1 To SCREEN_H, 1 To SCREEN_W)

    t0 = Platform.Clock()
    For i = 1 To N
        World.Rotate 0.05
        Raycaster.Render frame, 0#
    Next i
    msRender = (Platform.Clock() - t0) / N * 1000#

    t0 = Platform.Clock()
    For i = 1 To N
        World.Rotate 0.05
        Raycaster.Render frame, 0#
        Renderer.Blit frame
    Next i
    msWrite = (Platform.Clock() - t0) / N * 1000# - msRender

    t0 = Platform.Clock()
    For i = 1 To N
        World.Rotate 0.05
        Raycaster.Render frame, 0#
        Renderer.Blit frame
        DoEvents
    Next i
    msPaint = (Platform.Clock() - t0) / N * 1000# - msRender - msWrite

    MsgBox "Per frame, in milliseconds:" & vbCrLf & _
           "  raycasting (VBA maths):  " & Format$(msRender, "0.0") & vbCrLf & _
           "  writing the cells:       " & Format$(msWrite, "0.0") & vbCrLf & _
           "  repaint (DoEvents):      " & Format$(msPaint, "0.0"), vbInformation, "CellCaster benchmark"
End Sub
