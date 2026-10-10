Attribute VB_Name = "Raycaster"
'==============================================================================
' Raycaster: turns the world into a frame of palette indices.
' Pure maths - no Excel objects and no Windows API - so it is easy to test.
'
' Frame layout: frame(1..SCREEN_H, 1..SCREEN_W), row first, exactly the shape
' Range.Value2 expects, so the whole frame is written in one call.
'==============================================================================
Option Explicit

Private mBayer(0 To 3, 0 To 3) As Double   ' ordered-dither thresholds, 0..1
Private mFloorLevel() As Long              ' (row below horizon, x And 3) -> shade level
Private mCeilIndex() As Long               ' (row below horizon, x And 3) -> palette index
Private mZBuffer() As Double               ' wall distance per screen column
Private mWallScale As Double               ' wall height in cells at distance 1
Private mReady As Boolean

Public Sub Init()
    Dim half As Long
    Dim r As Long
    Dim bx As Long
    Dim y As Long
    Dim yCeil As Long
    Dim dist As Double
    Dim level As Long

    FillBayer

    ' Cells are square, so a unit cube looks square when its on-screen height
    ' equals its on-screen width: SCREEN_W / (2 * FOV_PLANE) cells at distance 1.
    mWallScale = SCREEN_W / (2# * FOV_PLANE)

    half = SCREEN_H \ 2
    ReDim mFloorLevel(0 To half - 1, 0 To 3)
    ReDim mCeilIndex(0 To half - 1, 0 To 3)
    ReDim mZBuffer(0 To SCREEN_W - 1)

    ' The distance seen through a screen row never changes, so the fog level
    ' of floor and ceiling is precomputed per row and per dither column.
    For r = 0 To half - 1
        y = half + r
        yCeil = half - 1 - r
        dist = (mWallScale * 0.5) / (r + 0.5)
        For bx = 0 To 3
            mFloorLevel(r, bx) = ShadeLevel(Brightness(dist), bx, y And 3)
            level = ShadeLevel(Brightness(dist) * 0.9, bx, yCeil And 3)
            mCeilIndex(r, bx) = PaletteIndex(colCeiling, level)
        Next bx
    Next r

    mReady = True
End Sub

Private Sub FillBayer()
    Dim v As Variant
    Dim i As Long

    v = Array(0, 8, 2, 10, 12, 4, 14, 6, 3, 11, 1, 9, 15, 7, 13, 5)
    For i = 0 To 15
        mBayer(i And 3, i \ 4) = (v(i) + 0.5) / 16#
    Next i
End Sub

' Linear fog: 1 next to the camera, 0 at FOG_DISTANCE and beyond.
Private Function Brightness(ByVal dist As Double) As Double
    Brightness = 1# - dist / FOG_DISTANCE
    If Brightness < 0# Then Brightness = 0#
End Function

' Quantises a brightness to 0..NUM_SHADES. The Bayer threshold decides which
' way the fraction rounds, which fakes the in-between shades with a pattern.
Private Function ShadeLevel(ByVal bright As Double, ByVal bx As Long, ByVal by As Long) As Long
    ShadeLevel = Int(bright * NUM_SHADES + mBayer(bx, by))
    If ShadeLevel > NUM_SHADES Then ShadeLevel = NUM_SHADES
End Function

Private Function PaletteIndex(ByVal color As Long, ByVal level As Long) As Long
    If level > 0 Then PaletteIndex = color * NUM_SHADES + level
End Function

' Renders one frame. t is the elapsed time in seconds (used to animate gems).
Public Sub Render(ByRef frame() As Variant, ByVal t As Double)
    If Not mReady Then Init
    DrawFloorAndCeiling frame
    DrawWalls frame
    DrawGems frame, t
End Sub

' Floor casting: each row below the horizon is a line across the floor at a
' known distance, so the world position just steps linearly from left to right.
Private Sub DrawFloorAndCeiling(ByRef frame() As Variant)
    Dim half As Long
    Dim r As Long
    Dim x As Long
    Dim bx As Long
    Dim rowFloor As Long
    Dim rowCeil As Long
    Dim dist As Double
    Dim fx As Double
    Dim fy As Double
    Dim stepX As Double
    Dim stepY As Double
    Dim level As Long

    half = SCREEN_H \ 2
    For r = 0 To half - 1
        rowFloor = half + r + 1
        rowCeil = half - r
        dist = (mWallScale * 0.5) / (r + 0.5)

        stepX = dist * 2# * PlaneX / SCREEN_W
        stepY = dist * 2# * PlaneY / SCREEN_W
        fx = PosX + dist * (DirX - PlaneX) + stepX * 0.5
        fy = PosY + dist * (DirY - PlaneY) + stepY * 0.5

        For x = 0 To SCREEN_W - 1
            bx = x And 3
            level = mFloorLevel(r, bx)
            If level = 0 Then
                frame(rowFloor, x + 1) = 0
            ElseIf ((Int(fx) + Int(fy)) And 1) = 0 Then
                frame(rowFloor, x + 1) = colFloor * NUM_SHADES + level
            Else
                frame(rowFloor, x + 1) = colStoneDark * NUM_SHADES + level
            End If
            frame(rowCeil, x + 1) = mCeilIndex(r, bx)
            fx = fx + stepX
            fy = fy + stepY
        Next x
    Next r
End Sub

' One ray per screen column, walked through the grid with DDA.
Private Sub DrawWalls(ByRef frame() As Variant)
    Dim x As Long
    Dim y As Long
    Dim k As Long
    Dim cameraX As Double
    Dim rayX As Double
    Dim rayY As Double
    Dim mapX As Long
    Dim mapY As Long
    Dim deltaX As Double
    Dim deltaY As Double
    Dim sideX As Double
    Dim sideY As Double
    Dim stepX As Long
    Dim stepY As Long
    Dim side As Long
    Dim dist As Double
    Dim lineH As Double
    Dim yTop As Double
    Dim drawStart As Long
    Dim drawEnd As Long
    Dim wallX As Double
    Dim texId As Long
    Dim texX As Long
    Dim texY As Long
    Dim texPos As Double
    Dim texStep As Double
    Dim bright As Double
    Dim level As Long
    Dim levels(0 To 3) As Long

    For x = 0 To SCREEN_W - 1
        cameraX = 2# * (x + 0.5) / SCREEN_W - 1#
        rayX = DirX + PlaneX * cameraX
        rayY = DirY + PlaneY * cameraX

        mapX = Int(PosX)
        mapY = Int(PosY)

        If rayX = 0# Then deltaX = 1E+30 Else deltaX = Abs(1# / rayX)
        If rayY = 0# Then deltaY = 1E+30 Else deltaY = Abs(1# / rayY)

        If rayX < 0# Then
            stepX = -1
            sideX = (PosX - mapX) * deltaX
        Else
            stepX = 1
            sideX = (mapX + 1# - PosX) * deltaX
        End If
        If rayY < 0# Then
            stepY = -1
            sideY = (PosY - mapY) * deltaY
        Else
            stepY = 1
            sideY = (mapY + 1# - PosY) * deltaY
        End If

        ' Step to the next grid line until a wall is hit. The solid outer ring
        ' of the level guarantees a hit; the bounds test is only a safety net.
        Do
            If sideX < sideY Then
                sideX = sideX + deltaX
                mapX = mapX + stepX
                side = 0
            Else
                sideY = sideY + deltaY
                mapY = mapY + stepY
                side = 1
            End If
            If mapX < 0 Or mapY < 0 Or mapX >= MAP_W Or mapY >= MAP_H Then Exit Do
        Loop While Map(mapX, mapY) = 0

        ' Perpendicular distance (not Euclidean) avoids the fisheye effect.
        If side = 0 Then dist = sideX - deltaX Else dist = sideY - deltaY
        If dist < 0.05 Then dist = 0.05
        mZBuffer(x) = dist

        If mapX >= 0 And mapY >= 0 And mapX < MAP_W And mapY < MAP_H Then
            texId = Map(mapX, mapY)
        Else
            texId = 1
        End If

        ' Where along the wall the ray landed picks the texture column.
        If side = 0 Then wallX = PosY + dist * rayY Else wallX = PosX + dist * rayX
        wallX = wallX - Int(wallX)
        texX = Int(wallX * TEX_SIZE)
        If texX > TEX_SIZE - 1 Then texX = TEX_SIZE - 1
        If side = 0 And rayX < 0# Then texX = TEX_SIZE - 1 - texX
        If side = 1 And rayY > 0# Then texX = TEX_SIZE - 1 - texX

        ' A cell is painted when its centre falls inside the wall slice.
        lineH = mWallScale / dist
        yTop = (SCREEN_H - lineH) * 0.5
        drawStart = Int(yTop + 0.5)
        drawEnd = Int(yTop + lineH - 0.5)
        If drawStart < 0 Then drawStart = 0
        If drawEnd > SCREEN_H - 1 Then drawEnd = SCREEN_H - 1

        ' The whole column shares one distance, so only 4 dither results exist.
        bright = Brightness(dist)
        If side = 1 Then bright = bright * SIDE_SHADE
        For k = 0 To 3
            levels(k) = ShadeLevel(bright, x And 3, k)
        Next k

        texStep = TEX_SIZE / lineH
        texPos = (drawStart + 0.5 - yTop) * texStep
        For y = drawStart To drawEnd
            level = levels(y And 3)
            If level = 0 Then
                frame(y + 1, x + 1) = 0
            Else
                texY = Int(texPos)
                If texY > TEX_SIZE - 1 Then texY = TEX_SIZE - 1
                frame(y + 1, x + 1) = Tex(texId, texX, texY) * NUM_SHADES + level
            End If
            texPos = texPos + texStep
        Next y
    Next x
End Sub

' Gems are billboards: flat images that always face the camera. They are
' drawn far to near and tested against the wall z-buffer column by column.
Private Sub DrawGems(ByRef frame() As Variant, ByVal t As Double)
    Dim order() As Long
    Dim depth() As Double
    Dim n As Long
    Dim i As Long
    Dim j As Long
    Dim tmpL As Long
    Dim tmpD As Double
    Dim invDet As Double
    Dim relX As Double
    Dim relY As Double
    Dim camX As Double
    Dim camZ As Double
    Dim size As Double
    Dim centreX As Double
    Dim centreY As Double
    Dim x0 As Long
    Dim x1 As Long
    Dim y0 As Long
    Dim y1 As Long
    Dim sx As Long
    Dim sy As Long
    Dim u As Long
    Dim v As Long
    Dim texel As Long
    Dim level As Long
    Dim bright As Double

    If GemCount = 0 Then Exit Sub
    ReDim order(1 To GemCount)
    ReDim depth(1 To GemCount)

    invDet = 1# / (PlaneX * DirY - DirX * PlaneY)

    ' Camera-space depth of every living gem in front of the player.
    For i = 1 To GemCount
        If Gems(i).alive Then
            relX = Gems(i).x - PosX
            relY = Gems(i).y - PosY
            camZ = invDet * (-PlaneY * relX + PlaneX * relY)
            If camZ > 0.2 Then
                n = n + 1
                order(n) = i
                depth(n) = camZ
            End If
        End If
    Next i

    ' Insertion sort, farthest first. There are only a handful of gems.
    For i = 2 To n
        tmpL = order(i)
        tmpD = depth(i)
        j = i - 1
        Do While j >= 1
            If depth(j) >= tmpD Then Exit Do
            order(j + 1) = order(j)
            depth(j + 1) = depth(j)
            j = j - 1
        Loop
        order(j + 1) = tmpL
        depth(j + 1) = tmpD
    Next i

    For i = 1 To n
        relX = Gems(order(i)).x - PosX
        relY = Gems(order(i)).y - PosY
        camX = invDet * (DirY * relX - DirX * relY)
        camZ = depth(i)

        size = mWallScale / camZ * GEM_SCALE
        If size >= 1# Then
            centreX = (SCREEN_W * 0.5) * (1# + camX / camZ)
            centreY = SCREEN_H * 0.5 + Sin(t * 3# + order(i)) * size * 0.15

            x0 = Int(centreX - size * 0.5 + 0.5)
            x1 = Int(centreX + size * 0.5 - 0.5)
            y0 = Int(centreY - size * 0.5 + 0.5)
            y1 = Int(centreY + size * 0.5 - 0.5)

            ' Gems glow: they never fade below half brightness.
            bright = Brightness(camZ)
            If bright < 0.5 Then bright = 0.5

            For sx = x0 To x1
                If sx >= 0 And sx < SCREEN_W Then
                    If camZ < mZBuffer(sx) Then
                        u = Int((sx + 0.5 - (centreX - size * 0.5)) * TEX_SIZE / size)
                        If u < 0 Then u = 0
                        If u > TEX_SIZE - 1 Then u = TEX_SIZE - 1
                        For sy = y0 To y1
                            If sy >= 0 And sy < SCREEN_H Then
                                v = Int((sy + 0.5 - (centreY - size * 0.5)) * TEX_SIZE / size)
                                If v < 0 Then v = 0
                                If v > TEX_SIZE - 1 Then v = TEX_SIZE - 1
                                texel = Gem(u, v)
                                If texel >= 0 Then
                                    level = ShadeLevel(bright, sx And 3, sy And 3)
                                    frame(sy + 1, sx + 1) = texel * NUM_SHADES + level
                                End If
                            End If
                        Next sy
                    End If
                End If
            Next sx
        End If
    Next i
End Sub

' Top-left overlay: one cell per map tile, plus the player and a heading tick.
Public Sub DrawMinimap(ByRef frame() As Variant)
    Dim x As Long
    Dim y As Long
    Dim i As Long
    Dim hx As Long
    Dim hy As Long

    If MAP_W + 2 > SCREEN_W Or MAP_H + 2 > SCREEN_H Then Exit Sub

    For y = 0 To MAP_H - 1
        For x = 0 To MAP_W - 1
            If Map(x, y) <> 0 Then
                frame(y + 2, x + 2) = colStone * NUM_SHADES + 2
            Else
                frame(y + 2, x + 2) = 0
            End If
        Next x
    Next y

    For i = 1 To GemCount
        If Gems(i).alive Then
            frame(Int(Gems(i).y) + 2, Int(Gems(i).x) + 2) = colGold * NUM_SHADES + NUM_SHADES
        End If
    Next i

    hx = Int(PosX + DirX * 1.5)
    hy = Int(PosY + DirY * 1.5)
    If hx >= 0 And hy >= 0 And hx < MAP_W And hy < MAP_H Then
        If Map(hx, hy) = 0 Then frame(hy + 2, hx + 2) = colWhite * NUM_SHADES + 2
    End If
    frame(Int(PosY) + 2, Int(PosX) + 2) = colWhite * NUM_SHADES + NUM_SHADES
End Sub
