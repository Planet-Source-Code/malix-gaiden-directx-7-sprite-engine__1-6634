Attribute VB_Name = "SpriteEng"
'Ack TempTextWindow is switched usage with tempground
'be careful! Debugged but still uncertain
Option Explicit

Dim dx As New DirectX7
Dim dd As DirectDraw7

Global LightBased As Boolean
Global Running As Boolean
Global Const InitResX = 800 'res that the program was
Global Const InitResY = 600  'developed under
Global Const Longest = 256
Global Const Widest = 256

Type SpriteBase
    xpos As Integer
    ypos As Integer
    width As Integer
    height As Integer
    Frame As Integer
    Orientation As Integer
    TransparentColor As Long
    visible As Boolean
    effect As Integer 'number of the effect
    Effected As Boolean 'is it effected
    SpriteName As String 'Filename
    SetGround As Byte 'What background is it on
    BoxSize As RECT 'size of blt from dest
    DestBoxSize As RECT 'size of dest blt
    BoundingBox As RECT
    scalewidth As Integer
    ScaleHeigth As Integer
End Type

Type Overlay
    ViewPortX As Integer 'Width of viewing area
    ViewPortY As Integer 'Height of viewing area
    ToScreen As RECT 'Where on the screen to blt
    CornerPosX As Integer
    CornerPosY As Integer
    width As Integer
    height As Integer
    visible As Boolean
    TileWidth As Integer
    TileHeight As Integer
    ViewingArea As RECT 'use viewport and cornerpos to make box to blt toscreen
End Type

Type LightEffectBase
    color As Long
    Radius As Integer
End Type

Dim Primary As DirectDrawSurface7
Dim BackBuffer As DirectDrawSurface7
Dim ddsdPrimary As DDSURFACEDESC2


Dim BackGroundOverlays(1 To 3) As Overlay
Dim BackGrounds(1 To 3) As DirectDrawSurface7
Dim TempGround As DirectDrawSurface7
Dim ddsdTempGround As DDSURFACEDESC2

Dim TextGround As DirectDrawSurface7
Dim ddsdTextGround As DDSURFACEDESC2
Dim FontBase As DirectDrawSurface7
Dim ddsdFontBase As DDSURFACEDESC2
Dim fonth As Integer
Dim fontw As Integer
Dim TextGroundVisible As Boolean
Dim TextWindowWidth As Integer
Dim TextWindowHeight As Integer

Dim Sprites(1 To 40) As SpriteBase
Dim SpritesSurface(1 To 40) As DirectDrawSurface7
Dim ddsdSprite As DDSURFACEDESC2

Global CurResX As Integer
Global CurResY As Integer

Dim LightEffects(1 To 10) As LightEffectBase

Dim ddsdTextWindow As DDSURFACEDESC2
Dim TextWindow As DirectDrawSurface7
Dim TextWindowVisibile As Boolean
Dim TempTextWindow As DirectDrawSurface7

Dim Tiles(1 To 40) As DirectDrawSurface7
Dim TileOverLays(1 To 3, 1 To Longest, 1 To Widest) As Byte
Dim TileWidthSize As Integer
Dim TileHeightSize As Integer

Dim ScaleXConst As Integer
Dim scaleYconst As Integer

'Orientation Consts
Global Const Normal = 0
Global Const FlipX = 1
Global Const FlipY = 2
Global Const FlipXY = 3

'Effect Const
Global Const FadeOutDown = 1
Global Const FadeInDown = 2
Global Const FadeOutup = 3
Global Const FadeInup = 4

'Ground Consts
Global Const BackGround = 1
Global Const MidGround = 2
Global Const ForeGround = 3
Global Const View = 1
Global Const Corner = 2

Sub InitializeSystem(screenx As Integer, screeny As Integer)
Dim i As Integer
Dim coulor As DDCOLORKEY
On Local Error GoTo errout

Set dd = dx.DirectDrawCreate("")

Call dd.SetCooperativeLevel(VisibleForm.hWnd, DDSCL_FULLSCREEN Or DDSCL_EXCLUSIVE Or DDSCL_ALLOWREBOOT)
Call dd.SetDisplayMode(screenx, screeny, 16, 0, DDSDM_DEFAULT)

LightBased = False

CurResX = screenx
CurResY = screeny

ScaleXConst = CurResX / InitResX
scaleYconst = CurResY / InitResY

'Primary initialize
ddsdPrimary.lFlags = DDSD_CAPS Or DDSD_BACKBUFFERCOUNT
ddsdPrimary.ddsCaps.lCaps = DDSCAPS_PRIMARYSURFACE Or DDSCAPS_FLIP Or DDSCAPS_COMPLEX
ddsdPrimary.lBackBufferCount = 1
Set Primary = dd.CreateSurface(ddsdPrimary)

'BackBuffer
Dim caps As DDSCAPS2
caps.lCaps = DDSCAPS_BACKBUFFER
Set BackBuffer = Primary.GetAttachedSurface(caps)

coulor.high = vbBlack
coulor.low = vbBlack
BackBuffer.SetColorKey DDCKEY_SRCBLT, coulor

'Set overlays to nothing
Set BackGrounds(1) = Nothing
Set BackGrounds(2) = Nothing
Set BackGrounds(3) = Nothing

'Hide Backgrounds/overlays
For i = 1 To 3
    BackGroundOverlays(i).visible = False
Next i

'Set TempGround
ddsdTempGround.lFlags = DDSD_CAPS Or DDSD_WIDTH Or DDSD_HEIGHT
ddsdTempGround.ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN
ddsdTempGround.lWidth = screenx
ddsdTempGround.lHeight = screeny
Set TempGround = dd.CreateSurface(ddsdTempGround)

'Set TextGrounds

ddsdTextGround.lFlags = DDSD_CAPS Or DDSD_WIDTH Or DDSD_HEIGHT
ddsdTextGround.ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN
ddsdTextGround.lWidth = screenx
ddsdTextGround.lHeight = screeny
Set TextGround = dd.CreateSurface(ddsdTextGround)
Set TempTextWindow = dd.CreateSurface(ddsdTextGround)
Set TextWindow = dd.CreateSurface(ddsdTextGround)
coulor.high = vbBlack
coulor.low = vbBlack

TempTextWindow.SetColorKey DDCKEY_SRCBLT, coulor
TextGround.SetColorKey DDCKEY_SRCBLT, coulor
TextWindow.SetColorKey DDCKEY_SRCBLT, coulor
TempGround.SetColorKey DDCKEY_SRCBLT, coulor
TempTextWindow.SetFillColor vbBlack
TextWindow.SetFillColor vbBlack
TextGround.SetFillColor vbBlack
TempGround.SetFillColor vbBlack
Dim fullscreenblt As RECT

With fullscreenblt
    .Top = 0
    .Left = 0
    .Bottom = CurResY
    .Right = CurResX
End With

TempTextWindow.BltColorFill fullscreenblt, vbBlack
TextWindow.BltColorFill fullscreenblt, vbBlack
TextGround.BltColorFill fullscreenblt, vbBlack
TempGround.BltColorFill fullscreenblt, vbBlack

TextWindowVisibile = False

'Set Sprites to NonVisible
For i = 1 To 40
    Sprites(i).visible = False
Next i

Exit Sub
errout:
Running = False
End Sub

Sub TerminateSystem()
Dim i As Integer

Set Primary = Nothing
Set BackBuffer = Nothing
Set TempGround = Nothing
Set TextGround = Nothing
Set TempTextWindow = Nothing
Set BackGrounds(1) = Nothing
Set BackGrounds(2) = Nothing
Set BackGrounds(3) = Nothing
Set TextWindow = Nothing


For i = 1 To 40
    Set SpritesSurface(i) = Nothing
Next i

Call dd.RestoreDisplayMode
Call dd.SetCooperativeLevel(VisibleForm.hWnd, DDSCL_NORMAL)

End Sub

Sub ClearBackgrounds()
ClearBackground 1
ClearBackground 2
ClearBackground 3

End Sub

Sub ClearBackground(backgroundnumber As Byte)
Dim temp As RECT

With temp
    .Top = 0
    .Left = 0
    .Bottom = BackGroundOverlays(backgroundnumber).height
    .Right = BackGroundOverlays(backgroundnumber).width
End With

BackGrounds(backgroundnumber).BltColorFill temp, vbBlack
End Sub

Sub ClearTempGround()
Dim temp As RECT

With temp
    .Top = 0
    .Left = 0
    .Bottom = CurResY
    .Right = CurResX
End With
TempGround.BltColorFill temp, vbBlack
End Sub

Sub ClearBackBuffer()
Dim temp As RECT

With temp
 .Top = 0
 .Left = 0
 .Bottom = CurResY
 .Right = CurResX
End With
BackBuffer.BltColorFill temp, vbBlack
End Sub

Sub ClearTextGround()
Dim temp As RECT

With temp
    .Top = 0
    .Left = 0
    .Bottom = CurResY
    .Right = CurResX
End With
TextGround.BltColorFill temp, vbBlack
End Sub

Sub SetBackGroundByFileName(backgroundnumber As Byte, BackGround As String, xwidth As Integer, ywidth As Integer)
Dim ddsdtempfile As DDSURFACEDESC2
Dim coulor As DDCOLORKEY
coulor.high = vbBlack
coulor.low = vbBlack
ddsdtempfile.lFlags = DDSD_CAPS Or DDSD_WIDTH Or DDSD_HEIGHT
ddsdtempfile.ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN
ddsdtempfile.lWidth = xwidth
ddsdtempfile.lHeight = ywidth
Set BackGrounds(backgroundnumber) = dd.CreateSurfaceFromFile(BackGround & ".bmp", ddsdtempfile)
BackGroundOverlays(backgroundnumber).height = ywidth
BackGroundOverlays(backgroundnumber).width = xwidth
BackGrounds(backgroundnumber).SetColorKey DDCKEY_SRCBLT, coulor
End Sub

Sub SetPortionBackGround(backgroundnumber As Integer, newX As Integer, newy As Integer, vieworcorner As Byte)
'change view port
If vieworcorner = View Then
    BackGroundOverlays(backgroundnumber).ViewPortX = newX
    BackGroundOverlays(backgroundnumber).ViewPortY = newy
Else
    BackGroundOverlays(backgroundnumber).CornerPosX = newX
    BackGroundOverlays(backgroundnumber).CornerPosY = newy
End If

'Recalculate Viewing Area
With BackGroundOverlays(backgroundnumber).ViewingArea
    .Top = BackGroundOverlays(backgroundnumber).CornerPosY
    .Bottom = BackGroundOverlays(backgroundnumber).CornerPosY + BackGroundOverlays(backgroundnumber).ViewPortY
    .Left = BackGroundOverlays(backgroundnumber).CornerPosX
    .Right = BackGroundOverlays(backgroundnumber).CornerPosX + BackGroundOverlays(backgroundnumber).ViewPortX
End With
End Sub

Sub SetBackgroundToScreen(backgroundnumber As Byte, newX As Integer, newy As Integer, newwidth As Integer, newheight As Integer)
With BackGroundOverlays(backgroundnumber).ToScreen
    .Top = newy * ScaleXConst
    .Left = newX * scaleYconst
    .Bottom = .Top + newheight * scaleYconst
    .Right = .Left + newwidth * ScaleXConst
End With
End Sub

Sub SetBackGroundVisibility(backgroundnumber As Byte, visible As Boolean)
BackGroundOverlays(backgroundnumber).visible = visible
End Sub

Sub FadeEffect(effect As Integer, color As Long, width As Integer)
Dim i As Integer
Dim temprect As RECT

If effect = FadeOutDown Then
'Fadeout
With temprect
    .Top = 0
    .Left = 0
    .Right = CurResX
End With
 
For i = 1 To (CurResY + width * scaleYconst) Step (width * scaleYconst)
    temprect.Bottom = i
    Primary.BltColorFill temprect, color
    delay 10
Next i
 temprect.Bottom = CurResY
 Primary.BltColorFill temprect, color
End If

If effect = FadeOutup Then
'Fadeout
With temprect
    .Top = CurResY - width * scaleYconst
    .Left = 0
    .Right = CurResX
    .Bottom = CurResY
End With
 
For i = (CurResY + width * scaleYconst) To 1 Step -(width * scaleYconst)
    temprect.Top = i
    Primary.BltColorFill temprect, color
    delay 10
Next i
 temprect.Top = 0
 Primary.BltColorFill temprect, color

End If

If effect = FadeInDown Then

With temprect
    .Top = 0
    .Left = 0
    .Right = CurResX
End With
 
For i = 1 To (CurResY + width * scaleYconst) Step (width * scaleYconst)
    temprect.Bottom = i
    Primary.Blt temprect, BackBuffer, temprect, DDBLT_KEYSRC Or DDBLT_WAIT
    delay 10
Next i
 temprect.Bottom = CurResY
 Primary.Blt temprect, BackBuffer, temprect, DDBLT_KEYSRC Or DDBLT_WAIT

End If

If effect = FadeInup Then

With temprect
    .Top = CurResY - width * scaleYconst
    .Left = 0
    .Right = CurResX
    .Bottom = CurResY
End With
 
For i = (CurResY + width * scaleYconst) To 1 Step -(width * scaleYconst)
    temprect.Top = i
    Primary.Blt temprect, BackBuffer, temprect, DDBLT_KEYSRC Or DDBLT_WAIT
    delay 10
Next i
 temprect.Top = 0
 Primary.Blt temprect, BackBuffer, temprect, DDBLT_KEYSRC Or DDBLT_WAIT

End If

End Sub

Sub DisplayScreen()
Dim i As Integer
Dim gr As Integer
Dim tempa As RECT
Dim dbltfx As DDBLTFX
Dim temp As RECT

With temp
    .Top = 0
    .Left = 0
    .Bottom = CurResY
    .Right = CurResX
End With

BackBuffer.BltColorFill temp, vbBlack

For gr = 1 To 3
'Copy Ground background
    If BackGroundOverlays(gr).visible Then
        BackBuffer.Blt BackGroundOverlays(gr).ToScreen, BackGrounds(gr), BackGroundOverlays(gr).ViewingArea, DDBLT_WAIT Or DDBLT_KEYSRC
    End If
'copy ground sprite
For i = 1 To 40
    If (Sprites(i).visible = True) Then
        If (Sprites(i).SetGround = gr) Then
        If Sprites(i).Orientation = Normal Then
            BackBuffer.Blt Sprites(i).DestBoxSize, SpritesSurface(i), Sprites(i).BoxSize, DDBLT_KEYSRC Or DDBLT_WAIT
        Else
            If Sprites(i).Orientation = FlipX Then dbltfx.lDDFX = DDBLTFX_NOTEARING Or DDBLTFX_MIRRORLEFTRIGHT
            If Sprites(i).Orientation = FlipY Then dbltfx.lDDFX = DDBLTFX_NOTEARING Or DDBLTFX_MIRRORUPDOWN
            If Sprites(i).Orientation = FlipXY Then dbltfx.lDDFX = DDBLTFX_NOTEARING Or DDBLTFX_ROTATE270
            BackBuffer.BltFx Sprites(i).DestBoxSize, SpritesSurface(i), Sprites(i).BoxSize, DDBLT_KEYSRC Or DDBLT_WAIT Or DDBLT_DDFX, dbltfx
       End If
    End If
    End If
Next i
Next gr
'copy text ground
Dim FullScreenRect As RECT
With FullScreenRect
    .Top = 0
    .Left = 0
    .Right = CurResX
    .Bottom = CurResY
End With

If TextWindowVisibile = True Then BackBuffer.Blt FullScreenRect, TextWindow, FullScreenRect, DDBLT_KEYSRC Or DDBLT_WAIT
If TextGroundVisible = True Then BackBuffer.Blt FullScreenRect, TextGround, FullScreenRect, DDBLT_KEYSRC Or DDBLT_WAIT

With tempa
    .Top = 0
    .Left = 0
    .Bottom = CurResY
    .Right = CurResX
End With

'BackBuffer.BltFast 0, 0, TempGround, tempa, DDBLTFAST_WAIT
End Sub

Sub SetSpriteOrientation(spritenumber As Byte, Orient As Byte)
Sprites(spritenumber).Orientation = Orient
End Sub

Sub SetSpriteFrame(spritenumber As Byte, NewFrame As Byte)
Dim tempsurf As DDSURFACEDESC2
Dim color As DDCOLORKEY

color.high = Sprites(spritenumber).TransparentColor
color.low = Sprites(spritenumber).TransparentColor
Sprites(spritenumber).Frame = NewFrame

tempsurf.lFlags = DDSD_CAPS Or DDSD_WIDTH Or DDSD_HEIGHT
tempsurf.ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN
tempsurf.lWidth = Sprites(spritenumber).width
tempsurf.lHeight = Sprites(spritenumber).height
If NewFrame = 1 Then
 Set SpritesSurface(spritenumber) = dd.CreateSurfaceFromFile(Sprites(spritenumber).SpriteName + ".bmp", tempsurf)
Else
 Set SpritesSurface(spritenumber) = dd.CreateSurfaceFromFile(Sprites(spritenumber).SpriteName & Right$(Str$(NewFrame), 1) + ".bmp", tempsurf)
End If

SpritesSurface(spritenumber).SetColorKey DDCKEY_SRCBLT, color
End Sub

Sub ReloadSprite(spritenumber As Integer)
Dim tempsurf As DDSURFACEDESC2
Dim color As DDCOLORKEY
Dim tempstring As String

tempstring = Sprites(spritenumber).SpriteName
If Sprites(spritenumber).Frame > 1 Then
 tempstring = tempstring + Str$(Sprites(spritenumber).Frame)
End If

tempstring = tempstring + ".bmp"

tempsurf.lFlags = DDSD_CAPS Or DDSD_WIDTH Or DDSD_HEIGHT
tempsurf.ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN
tempsurf.lWidth = Sprites(spritenumber).width
tempsurf.lHeight = Sprites(spritenumber).height

Set SpritesSurface(spritenumber) = dd.CreateSurfaceFromFile(tempstring, tempsurf)

color.high = Sprites(spritenumber).TransparentColor
color.low = Sprites(spritenumber).TransparentColor
End Sub

Sub SetSprite(spritenumber As Integer, Effected As Boolean, EffectNumber As Byte, SpriteName As String, SpriteWidth As Integer, SpriteHeight As Integer, SpriteGround As Byte, InvisibleColor As Long)
Dim tempsurf As DDSURFACEDESC2
Dim color As DDCOLORKEY

Sprites(spritenumber).Effected = Effected
Sprites(spritenumber).effect = EffectNumber
Sprites(spritenumber).width = SpriteWidth
Sprites(spritenumber).height = SpriteHeight

Sprites(spritenumber).BoxSize.Top = 0
Sprites(spritenumber).BoxSize.Left = 0
Sprites(spritenumber).BoxSize.Bottom = SpriteHeight
Sprites(spritenumber).BoxSize.Right = SpriteWidth

Sprites(spritenumber).scalewidth = SpriteWidth
Sprites(spritenumber).ScaleHeigth = SpriteHeight

Sprites(spritenumber).Frame = 0
Sprites(spritenumber).Orientation = Normal
Sprites(spritenumber).SetGround = SpriteGround
Sprites(spritenumber).TransparentColor = InvisibleColor

Sprites(spritenumber).visible = True
Sprites(spritenumber).xpos = 0
Sprites(spritenumber).ypos = 0

Sprites(spritenumber).SpriteName = SpriteName

tempsurf.lFlags = DDSD_CAPS Or DDSD_WIDTH Or DDSD_HEIGHT
tempsurf.ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN
tempsurf.lWidth = SpriteWidth
tempsurf.lHeight = SpriteHeight


Set SpritesSurface(spritenumber) = dd.CreateSurfaceFromFile(SpriteName + ".bmp", tempsurf)

color.high = InvisibleColor
color.low = InvisibleColor

SpritesSurface(spritenumber).SetColorKey DDCKEY_SRCBLT, color
End Sub

Sub MoveSprite(spritenumber As Byte, spriteX As Integer, spriteY As Integer)
Sprites(spritenumber).xpos = spriteX * ScaleXConst
Sprites(spritenumber).ypos = spriteY * scaleYconst
With Sprites(spritenumber).DestBoxSize
    .Top = spriteY * scaleYconst
    .Left = spriteX * ScaleXConst
    .Bottom = spriteY * scaleYconst + Sprites(spritenumber).ScaleHeigth * scaleYconst
    .Right = spriteX * ScaleXConst + Sprites(spritenumber).scalewidth * ScaleXConst
End With
End Sub

Sub SetSpriteVisibility(spritenumber As Integer, visible As Boolean)
Sprites(spritenumber).visible = visible
End Sub

Sub SetSpriteScale(spritenumber As Byte, scalex As Integer, scaley As Integer)
Sprites(spritenumber).ScaleHeigth = scaley
Sprites(spritenumber).scalewidth = scalex
MoveSprite spritenumber, Sprites(spritenumber).xpos, Sprites(spritenumber).ypos
End Sub

Sub SetTextColor(color As Long)
TempGround.SetForeColor color
End Sub

Sub ClearTempWindow()
Dim temp As RECT
With temp
 .Left = 0
 .Top = 0
 .Bottom = CurResY
 .Right = CurResX
End With
TempTextWindow.BltColorFill temp, vbBlack
End Sub

Sub DisplayText(x As Integer, y As Integer, strs As String, ScaleXt As Integer, ScaleYT As Integer)

ClearTempGround

TempGround.DrawText 0, 0, strs, False

Dim temp As RECT
temp.Left = 0
temp.Top = 0
temp.Right = Len(strs) * 8

temp.Bottom = 16


Dim tempb As RECT
tempb.Top = y * scaleYconst
tempb.Left = x * ScaleXConst
tempb.Right = (ScaleXt + x) * ScaleXConst
tempb.Bottom = (ScaleYT + y) * scaleYconst

TextGround.Blt tempb, TempGround, temp, DDBLT_KEYSRC Or DDBLT_WAIT
End Sub

Sub SetFont(fontname As String, isbold As Boolean, isunderlined As Boolean, isstrikethrough As Boolean, isitalics As Boolean)
Dim newfont As New StdFont

newfont.Size = 8
newfont.Bold = isbold
newfont.Underline = isunderlined
newfont.Strikethrough = isstrikethrough
newfont.Italic = isitalics
newfont.Name = fontname
TempGround.SetFont newfont
End Sub
Sub SetTextBkgColor(color As Long)
Dim temp As RECT
Dim coulor As DDCOLORKEY

With temp
    .Top = 0
    .Left = 0
    .Bottom = CurResY
    .Right = CurResX
End With
coulor.high = color
coulor.low = color

TempGround.SetColorKey DDCKEY_SRCBLT, coulor
TempGround.BltColorFill temp, color

End Sub


Sub SetTextVisibility(Visibility As Boolean)
TextGroundVisible = Visibility
End Sub

Sub Flip()
    Primary.Flip Nothing, DDFLIP_WAIT
End Sub

Sub SetTextWindowByFileName(WindowFile As String, xwidth As Integer, ywidth As Integer)
Dim ddsdtempfile As DDSURFACEDESC2
Dim coulor As DDCOLORKEY

coulor.high = vbBlack
coulor.low = vbBlack

ddsdtempfile.lFlags = DDSD_CAPS Or DDSD_WIDTH Or DDSD_HEIGHT
ddsdtempfile.ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN
ddsdtempfile.lWidth = xwidth
ddsdtempfile.lHeight = ywidth
Set TempTextWindow = dd.CreateSurfaceFromFile(WindowFile, ddsdtempfile)

TempTextWindow.SetColorKey DDCKEY_SRCBLT, coulor

TextWindowWidth = xwidth
TextWindowHeight = ywidth
End Sub

Sub DisplayAbox(x As Integer, y As Integer, width As Integer, height As Integer)
Dim tempa As RECT
Dim tempb As RECT

tempa.Top = y * scaleYconst
tempa.Left = x * ScaleXConst
tempa.Bottom = (height + y) * scaleYconst
tempa.Right = (width + x) * ScaleXConst


tempb.Top = 0
tempb.Left = 0
tempb.Right = TextWindowWidth
tempb.Bottom = TextWindowHeight

TextWindow.Blt tempa, TempTextWindow, tempb, DDBLT_KEYSRC Or DDBLT_WAIT
End Sub
Sub SetTextWindowVisibility(visible As Boolean)
TextWindowVisibile = visible
End Sub

Sub SetBackGroundTileSize(newwidth As Integer, newheight As Integer)
BackGroundTileWidth = newwidth * ScaleXConst
BackGroundTileHeight = newheight * scaleYconst
End Sub

Sub AutoSetWindow(x As Integer, y As Integer, width As Integer, strs As String, valx As Integer, valy As Integer)
DisplayAbox x - width, y, valx + width, valy
DisplayText x, y, strs, valx, valy

End Sub

Sub SetSpriteNewGround(spritenumber As Integer, newground As Byte)
Sprites(spritenumber).SetGround = newground
End Sub

Sub delay(val As Integer)
Dim i As Integer
Dim b As Integer
For i = 1 To val
 For b = 1 To 10000
 Next b
Next i
End Sub
  
Function SpriteFrame(spritenumber As Integer) As Integer
SpriteFrame = Sprites(spritenumber).Frame
End Function

Sub Loadtile(tilenumber As Byte, Filename As String)
Dim ddsdTempSurf As DDSURFACEDESC2
Dim coulor As DDCOLORKEY

ddsdTempSurf.lFlags = DDSD_CAPS Or DDSD_WIDTH Or DDSD_HEIGHT
ddsdTempSurf.ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN
ddsdTempSurf.lWidth = TileWidthSize
ddsdTempSurf.lHeight = TileHeightSize

Set Tiles(tilenumber) = dd.CreateSurfaceFromFile(Filename, ddsdTempSurf)

coulor.low = vbBlack
coulor.high = vbBlack
Tiles(tilenumber).SetColorKey DDCKEY_SRCBLT, coulor

End Sub

Sub SetTileDimensions(width As Integer, height As Integer)
TileWidthSize = width
TileHeightSize = height
End Sub

Sub SetTile(Ground As Byte, x As Integer, y As Integer, tilenumber As Byte)
TileOverLays(Ground, x, y) = tilenumber
End Sub

Sub InitializeTileGround(GroundNumber As Byte, TileWidth As Integer, TileHeight As Integer)
Dim ddsdTempSurf As DDSURFACEDESC2
Dim coulor As DDCOLORKEY

ddsdTempSurf.lFlags = DDSD_CAPS Or DDSD_WIDTH Or DDSD_HEIGHT
ddsdTempSurf.ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN
ddsdTempSurf.lWidth = TileWidth * TileWidthSize
ddsdTempSurf.lHeight = TileHeight * TileHeightSize

Set BackGrounds(GroundNumber) = dd.CreateSurface(ddsdTempSurf)

coulor.high = vbBlack
coulor.low = vbBlack
BackGrounds(GroundNumber).SetColorKey DDCKEY_SRCBLT, coulor

BackGroundOverlays(GroundNumber).TileHeight = TileHeight
BackGroundOverlays(GroundNumber).TileWidth = TileWidth
BackGroundOverlays(GroundNumber).width = TileWidth * TileWidthSize
BackGroundOverlays(GroundNumber).height = TileHeight * TileHeightSize
End Sub
 
Sub DisplayTileBackground(whichground As Byte, clearfirst As Boolean)
Dim xa As Integer
Dim ya As Integer
Dim tempa As RECT
Dim tempb As RECT

tempa.Left = 0
tempa.Top = 0
tempa.Right = TileWidthSize - 1
tempa.Bottom = TileHeightSize - 1

If clearfirst Then ClearBackground whichground

For xa = 1 To BackGroundOverlays(whichground).TileWidth
    For ya = 1 To BackGroundOverlays(whichground).TileHeight
        tempb.Top = (ya - 1) * (TileHeightSize - 1) * scaleYconst
        tempb.Bottom = tempb.Top + (TileHeightSize - 1) * scaleYconst
        tempb.Left = (xa - 1) * (TileWidthSize - 1) * ScaleXConst
        tempb.Right = tempb.Left + (TileWidthSize - 1) * ScaleXConst
        BackGrounds(whichground).Blt tempb, Tiles(TileOverLays(whichground, xa, ya)), tempa, DDBLT_KEYSRC
    Next ya
Next xa

End Sub
