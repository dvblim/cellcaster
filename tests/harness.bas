Option VBASupport 1
Option Explicit
' Test harness. Runs the real engine modules inside LibreOffice Basic (in VBA
' compatibility mode) and dumps frames to the file named by CELLCASTER_OUT.
' tests/run_vba_tests.py compares that dump with the Python reference port.

Private Sub Dump(ByVal f As Integer, ByRef frame() As Variant)
    Dim x As Long, y As Long, s As String
    For y = 1 To SCREEN_H
        s = ""
        For x = 1 To SCREEN_W
            s = s & frame(y, x) & " "
        Next x
        Print #f, s
    Next y
End Sub

Sub Main()
    Dim frame() As Variant
    Dim f As Integer, i As Long, n As Long, s As String

    f = FreeFile
    Open Environ("CELLCASTER_OUT") For Output As #f
    On Error GoTo Oops
    ReDim frame(1 To SCREEN_H, 1 To SCREEN_W)

    s = ""
    For i = 1 To PALETTE_SIZE
        s = s & PaletteColor(i) & " "
    Next i
    Print #f, s

    World.LoadLevel: Textures.Build: Raycaster.Init
    PosX = 3.5: PosY = 4.5: World.Rotate 0.35
    Raycaster.Render frame, 0.4: Raycaster.DrawMinimap frame: Dump f, frame

    World.LoadLevel
    PosX = 16.6: PosY = 17.5: World.Rotate -0.25
    Raycaster.Render frame, 1.1: Dump f, frame

    World.LoadLevel
    PosX = 19.5: PosY = 14.5: World.Rotate -1.5707963
    Raycaster.Render frame, 2.3: Dump f, frame

    World.LoadLevel
    World.Rotate 2.9
    Raycaster.Render frame, 0#: Dump f, frame

    ' A long walk into walls: collisions and sliding.
    World.LoadLevel
    For i = 1 To 400
        World.Rotate 0.011
        World.TryMove DirX * 0.09 + 0.01, DirY * 0.09 + 0.03
    Next i
    Print #f, CLng(PosX * 1000000#) & " " & CLng(PosY * 1000000#)
    Raycaster.Render frame, 5#: Raycaster.DrawMinimap frame: Dump f, frame

    ' Stand on every gem in turn: pickups.
    World.LoadLevel
    For i = 1 To GemCount
        PosX = Gems(i).x + 0.1: PosY = Gems(i).y - 0.1
        n = n + World.CollectGems()
    Next i
    Print #f, n & " " & GemsLeft
    Raycaster.Render frame, 0.7: Raycaster.DrawMinimap frame: Dump f, frame

    Print #f, "DONE"
    Close #f
    Exit Sub
Oops:
    Print #f, "ERROR " & Err.Number & " " & Err.Description
    Close #f
End Sub
