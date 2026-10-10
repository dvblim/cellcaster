Attribute VB_Name = "Config"
'==============================================================================
' CellCaster - a textured raycasting engine that renders into worksheet cells.
' Config: every tunable constant and the colour palette live here.
'==============================================================================
Option Explicit

' --- Screen -------------------------------------------------------------------
Public Const SCREEN_W As Long = 128          ' width in cells ("pixels")
Public Const SCREEN_H As Long = 72           ' height in cells (keep it even)
Public Const CELL_SIZE_PT As Double = 3.75   ' cell edge in points (3.75pt = 5px)
Public Const SCREEN_TOP As Long = 2          ' first worksheet row of the screen
Public Const SCREEN_LEFT As Long = 2         ' first worksheet column of the screen
Public Const SHEET_NAME As String = "CellCaster"

' --- Camera and movement ------------------------------------------------------
Public Const FOV_PLANE As Double = 0.8       ' camera plane half-length (wider = more FOV)
Public Const MOVE_SPEED As Double = 3#       ' tiles per second
Public Const TURN_SPEED As Double = 2.4      ' radians per second
Public Const PLAYER_RADIUS As Double = 0.2   ' collision margin in tiles
Public Const TARGET_FPS As Long = 30

' --- Look ---------------------------------------------------------------------
Public Const FOG_DISTANCE As Double = 14#    ' distance at which everything is black
Public Const SIDE_SHADE As Double = 0.75     ' brightness of north/south wall faces
Public Const NUM_SHADES As Long = 4          ' lit levels per colour (plus black)
Public Const TEX_SIZE As Long = 16           ' textures are TEX_SIZE x TEX_SIZE
Public Const NUM_TEXTURES As Long = 5
Public Const GEM_SCALE As Double = 0.45      ' gem size relative to a wall tile
Public Const PICKUP_RADIUS As Double = 0.45

' --- Palette ------------------------------------------------------------------
' A cell holds a palette index. 0 is black; every other index is
'   colour * NUM_SHADES + level      (level = 1..NUM_SHADES)
' One conditional-formatting rule per index turns the number into a fill.
Public Enum ColorId
    colBrick = 0
    colMortar = 1
    colStone = 2
    colStoneDark = 3
    colWood = 4
    colWoodDark = 5
    colMetal = 6
    colMoss = 7
    colGold = 8
    colWhite = 9
    colCeiling = 10
    colFloor = 11
End Enum

Public Const NUM_COLORS As Long = 12
Public Const PALETTE_SIZE As Long = 48       ' NUM_COLORS * NUM_SHADES

' Full-brightness RGB of a base colour.
Public Function BaseColor(ByVal id As Long) As Long
    Select Case id
        Case colBrick:     BaseColor = RGB(190, 70, 55)
        Case colMortar:    BaseColor = RGB(215, 205, 185)
        Case colStone:     BaseColor = RGB(150, 150, 160)
        Case colStoneDark: BaseColor = RGB(95, 95, 110)
        Case colWood:      BaseColor = RGB(170, 115, 60)
        Case colWoodDark:  BaseColor = RGB(115, 72, 35)
        Case colMetal:     BaseColor = RGB(80, 125, 200)
        Case colMoss:      BaseColor = RGB(85, 160, 80)
        Case colGold:      BaseColor = RGB(250, 205, 60)
        Case colWhite:     BaseColor = RGB(245, 245, 245)
        Case colCeiling:   BaseColor = RGB(70, 80, 125)
        Case colFloor:     BaseColor = RGB(135, 120, 105)
        Case Else:         BaseColor = RGB(255, 0, 255)
    End Select
End Function

' RGB of a palette index (1..PALETTE_SIZE).
Public Function PaletteColor(ByVal index As Long) As Long
    Dim baseRgb As Long
    Dim level As Long
    Dim f As Double

    baseRgb = BaseColor((index - 1) \ NUM_SHADES)
    level = ((index - 1) Mod NUM_SHADES) + 1
    f = 0.15 + 0.85 * level / NUM_SHADES

    PaletteColor = RGB(Int((baseRgb And 255) * f), _
                       Int(((baseRgb \ 256) And 255) * f), _
                       Int(((baseRgb \ 65536) And 255) * f))
End Function
