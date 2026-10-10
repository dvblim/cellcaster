Attribute VB_Name = "World"
'==============================================================================
' World: the level, the player and the collectable gems.
'==============================================================================
Option Explicit

Public Const MAP_W As Long = 24
Public Const MAP_H As Long = 24

Public Type Sprite
    x As Double
    y As Double
    alive As Boolean
End Type

Public Map() As Long            ' Map(x, y): 0 = empty, 1..NUM_TEXTURES = wall
Public Gems() As Sprite         ' 1..GemCount
Public GemCount As Long
Public GemsLeft As Long

Public PosX As Double           ' player position, in tiles
Public PosY As Double
Public DirX As Double           ' unit view direction
Public DirY As Double
Public PlaneX As Double         ' camera plane, perpendicular to the direction
Public PlaneY As Double

Private mRow As Long

' Level legend:  1-5 wall texture   . floor   P player start   * gem
' The outer ring must be solid: the ray loop relies on it to terminate.
Public Sub LoadLevel()
    ReDim Map(0 To MAP_W - 1, 0 To MAP_H - 1)
    ReDim Gems(1 To MAP_W * MAP_H)
    GemCount = 0
    mRow = 0

    AddRow "111111111111111111111111"
    AddRow "1P.....2.......3.......1"
    AddRow "1......2...*...3.......1"
    AddRow "1......2.......3...*...1"
    AddRow "1..........33333.......1"
    AddRow "1......2...............1"
    AddRow "12.22222.......3.......1"
    AddRow "1......2.......3333.3331"
    AddRow "1......2...4...........1"
    AddRow "1..*...2...4...........1"
    AddRow "1......2...4444...555..1"
    AddRow "1..........4......5.5..1"
    AddRow "1......2...4......5*5..1"
    AddRow "12222222..........5.5..1"
    AddRow "1.............4........1"
    AddRow "1...5.5.......4444.44441"
    AddRow "1.............4........1"
    AddRow "1...5.5...3...4...*....1"
    AddRow "1.............4........1"
    AddRow "1...5.5...3...4...22...1"
    AddRow "1.........3........2...1"
    AddRow "1....*....3........2...1"
    AddRow "1.........3............1"
    AddRow "111111111111111111111111"

    GemsLeft = GemCount

    ' Start facing east. With map Y growing downwards, the right-hand side of
    ' the view is +Y, so the camera plane is (0, +FOV_PLANE).
    DirX = 1#: DirY = 0#
    PlaneX = 0#: PlaneY = FOV_PLANE
End Sub

Private Sub AddRow(ByVal text As String)
    Dim x As Long
    Dim ch As String

    For x = 0 To MAP_W - 1
        ch = Mid$(text, x + 1, 1)
        Select Case ch
            Case "1" To "9"
                Map(x, mRow) = CLng(ch)
            Case "P"
                PosX = x + 0.5
                PosY = mRow + 0.5
            Case "*"
                GemCount = GemCount + 1
                Gems(GemCount).x = x + 0.5
                Gems(GemCount).y = mRow + 0.5
                Gems(GemCount).alive = True
        End Select
    Next x
    mRow = mRow + 1
End Sub

Public Function IsWall(ByVal x As Double, ByVal y As Double) As Boolean
    Dim mx As Long
    Dim my As Long

    mx = Int(x)
    my = Int(y)
    If mx < 0 Or my < 0 Or mx >= MAP_W Or my >= MAP_H Then
        IsWall = True
    Else
        IsWall = (Map(mx, my) <> 0)
    End If
End Function

' Moves the player by (dx, dy), sliding along walls: each axis is tested on
' its own, so walking diagonally into a wall still moves along it.
Public Sub TryMove(ByVal dx As Double, ByVal dy As Double)
    Dim nx As Double
    Dim ny As Double

    nx = PosX + dx
    If Not IsWall(nx + Sgn(dx) * PLAYER_RADIUS, PosY - PLAYER_RADIUS) Then
        If Not IsWall(nx + Sgn(dx) * PLAYER_RADIUS, PosY + PLAYER_RADIUS) Then PosX = nx
    End If

    ny = PosY + dy
    If Not IsWall(PosX - PLAYER_RADIUS, ny + Sgn(dy) * PLAYER_RADIUS) Then
        If Not IsWall(PosX + PLAYER_RADIUS, ny + Sgn(dy) * PLAYER_RADIUS) Then PosY = ny
    End If
End Sub

' Rotates the view by angle radians (positive = turn right).
Public Sub Rotate(ByVal angle As Double)
    Dim c As Double
    Dim s As Double
    Dim t As Double

    c = Cos(angle)
    s = Sin(angle)

    t = DirX
    DirX = t * c - DirY * s
    DirY = t * s + DirY * c

    t = PlaneX
    PlaneX = t * c - PlaneY * s
    PlaneY = t * s + PlaneY * c
End Sub

' Collects any gem the player is touching. Returns how many were picked up.
Public Function CollectGems() As Long
    Dim i As Long
    Dim dx As Double
    Dim dy As Double

    For i = 1 To GemCount
        If Gems(i).alive Then
            dx = Gems(i).x - PosX
            dy = Gems(i).y - PosY
            If dx * dx + dy * dy < PICKUP_RADIUS * PICKUP_RADIUS Then
                Gems(i).alive = False
                GemsLeft = GemsLeft - 1
                CollectGems = CollectGems + 1
            End If
        End If
    Next i
End Function
