Attribute VB_Name = "Mod_TileEngine"
'Argentum Online 0.9.0.9
'
'Copyright (C) 2002 Márquez Pablo Ignacio
'Copyright (C) 2002 Otto Perez
'Copyright (C) 2002 Aaron Perkins
'Copyright (C) 2002 Matías Fernando Pequeño
'
'This program is free software; you can redistribute it and/or modify
'it under the terms of the GNU General Public License as published by
'the Free Software Foundation; either version 2 of the License, or
'any later version.
'
'This program is distributed in the hope that it will be useful,
'but WITHOUT ANY WARRANTY; without even the implied warranty of
'MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'GNU General Public License for more details.
'
'You should have received a copy of the GNU General Public License
'along with this program; if not, write to the Free Software
'Foundation, Inc., 59 Temple Place, Suite 330, Boston, MA  02111-1307  USA
'
'Argentum Online is based on Baronsoft's VB6 Online RPG
'You can contact the original creator of ORE at aaron@baronsoft.com
'for more information about ORE please visit http://www.baronsoft.com/
'
'
'You can contact me at:
'morgolock@speedy.com.ar
'www.geocities.com/gmorgolock
'Calle 3 número 983 piso 7 dto A
'La Plata - Pcia, Buenos Aires - Republica Argentina
'Código Postal 1900
'Pablo Ignacio Márquez



Option Explicit

'¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?
'¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?
'    C       O       N       S      T
'¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?
'¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?
'Map sizes in tiles
Public Const XMaxMapSize = 100
Public Const XMinMapSize = 1
Public Const YMaxMapSize = 100
Public Const YMinMapSize = 1

Public Const GrhFogata = 1521

'bltbit constant
Public Const SRCCOPY = &HCC0020 ' (DWORD) dest = source
Public Const COLOR_INVI As Long = &H5E1FA

'¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?
'¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?
'    T       I      P      O      S
'¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?
'¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?

'Encabezado bmp
Private Type BITMAPFILEHEADER
        bfType As Integer
        bfSize As Long
        bfReserved1 As Integer
        bfReserved2 As Integer
        bfOffBits As Long
End Type

Public Type ErroresGrh
    colores(0 To 9) As Long
    ErrorCritico As Boolean
    EsAnimacion As Boolean
End Type


'Info del encabezado del bmp
Private Type BITMAPINFOHEADER
        biSize As Long
        biWidth As Long
        biHeight As Long
        biPlanes As Integer
        biBitCount As Integer
        biCompression As Long
        biSizeImage As Long
        biXPelsPerMeter As Long
        biYPelsPerMeter As Long
        biClrUsed As Long
        biClrImportant As Long
End Type

'Posicion en un mapa
Public Type Position
    x As Integer
    y As Integer
End Type

'Posicion en el Mundo
Public Type WorldPos
    Map As Integer
    x As Integer
    y As Integer
End Type

'Contiene info acerca de donde se puede encontrar un grh
'tamaño y animacion
Public Type Grhdata
    sX As Integer
    sY As Integer
    FileNum As Long
    pixelWidth As Integer
    pixelHeight As Integer
    TileWidth As Single
    TileHeight As Single
    
    NumFrames As Integer
    Frames(1 To 45) As Long
    Speed As Single
End Type

'apunta a una estructura grhdata y mantiene la animacion
Public Type Grh
    GrhIndex As Long
    FrameCounter As Single
    Speed As Single
    Started As Byte
    Loops As Integer
    angle As Single
End Type

'Lista de cuerpos
Public Type BodyData
    Walk(1 To 4) As Grh
    HeadOffset As Position
End Type

'Lista de cabezas
Public Type HeadData
    Head(1 To 4) As Grh
End Type

'Lista de las animaciones de las armas
Type WeaponAnimData
    WeaponWalk(1 To 4) As Grh
    '[ANIM ATAK]
    WeaponAttack As Byte
End Type

Public Type tSupData
    GrhIndex As Long
    Ancho As Byte
    Alto As Byte
    Nombre As String
    Capa As Byte
End Type

'Lista de las animaciones de los escudos
Type ShieldAnimData
    ShieldWalk(1 To 4) As Grh
End Type


'Lista de cuerpos
Public Type FxData
    FX As Grh
    offsetx As Long
    offsety As Long
End Type

'Apariencia del personaje
Public Type Char
    Active As Byte
    Heading As Byte ' As E_Heading ?
    Pos As Position
    
    iHead As Integer
    Ibody As Integer
    Body As BodyData
    Head As HeadData
    Casco As HeadData
    Espalda As HeadData
    Botas As HeadData
    
    Arma As WeaponAnimData
    Escudo As ShieldAnimData
    UsandoArma As Boolean
    FX As Integer
    FxLoopTimes As Integer
    Criminal As Byte
    
    Nombre As String
    
    Moving As Byte
    MoveOffset As Position
    ServerIndex As Integer
    
    pie As Boolean
    muerto As Boolean
    Invisible As Boolean
    priv As Byte
    ClanID As Integer
    ClanName As String
    ColorName As Long
End Type

'Info de un objeto
Public Type Obj
    OBJIndex As Integer
    Amount As Integer
End Type

'Tipo de las celdas del mapa
Public Type MapBlock
    Graphic(1 To 4) As Grh
    CharIndex As Integer
    ObjGrh As Grh
    
    npcIndex As Integer
    OBJInfo As Obj
    TileExit As WorldPos
    Blocked As Byte
    
    Trigger As Integer
End Type

'Info de cada mapa
Public Type MapInfo
    Music As String
    name As String
    StartPos As WorldPos
    MapVersion As Integer
    
    'ME Only
    Changed As Byte
End Type

Public Type tIndiceCabeza
    Head(1 To 4) As Integer
End Type

Public Type tIndiceCabezaLong
    Head(1 To 4) As Long
End Type

Public configdirinits As String
Public MapPath As String


'Bordes del mapa
Public MinXBorder As Byte
Public MaxXBorder As Byte
Public MinYBorder As Byte
Public MaxYBorder As Byte

'Status del user
Public CurMap As Integer 'Mapa actual
Public UserIndex As Integer
Public MiClanID As Integer
Public UserMoving As Byte
Public UserBody As Integer
Public UserHead As Integer
Public UserPos As Position 'Posicion
Public AddtoUserPos As Position 'Si se mueve
Public UserCharIndex As Integer

Public UserMaxAGU As Integer
Public UserMinAGU As Integer
Public UserMaxHAM As Integer
Public UserMinHAM As Integer

Public EngineRun As Boolean
Public FramesPerSec As Integer
Public FramesPerSecCounter As Long

'Tamaño del la vista en Tiles
Public WindowTileWidth As Integer
Public WindowTileHeight As Integer

'Offset del desde 0,0 del main view
Public MainViewTop As Integer
Public MainViewLeft As Integer

'Cuantos tiles el engine mete en el BUFFER cuando
'dibuja el mapa. Ojo un tamaño muy grande puede
'volver el engine muy lento
Public TileBufferSize As Integer

'Handle to where all the drawing is going to take place
Public DisplayFormhWnd As Long

'Tamaño de los tiles en pixels
Public TilePixelHeight As Integer
Public TilePixelWidth As Integer

'?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿Totales?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?

Public NumBodies As Integer
Public numfxs As Integer

Public NumChars As Integer
Public LastChar As Integer
Public NumWeaponAnims As Integer
Public NumShieldAnims As Integer

'¿?¿?¿?¿?¿?¿?¿?¿?¿?¿Graficos¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?

Public LastTime As Long 'Para controlar la velocidad


'[CODE]:MatuX'
Public MainDestRect   As RECT
'[END]'
Public MainViewRect   As RECT
Public BackBufferRect As RECT

Public MainViewWidth As Integer
Public MainViewHeight As Integer



Public NumSuperficies As Integer
'¿?¿?¿?¿?¿?¿?¿?¿?¿?¿Graficos¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?
Public Grhdata() As Grhdata 'Guarda todos los grh
Public BodyData() As BodyData
Public HeadData() As HeadData
Public FxData() As FxData
Public WeaponAnimData() As WeaponAnimData
Public ShieldAnimData() As ShieldAnimData
Public CascoAnimData() As HeadData
Public SupData() As tSupData
Public Grh() As Grh 'Animaciones publicas
'¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?

'¿?¿?¿?¿?¿?¿?¿?¿?¿?¿Mapa?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?
Public MapData() As MapBlock ' Mapa
Public MapInfo As MapInfo ' Info acerca del mapa en uso
'¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?

Public PuedeVerClan As Boolean
'¿?¿?¿?¿?¿?¿?¿?¿?¿?¿Usuarios?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?
'
'epa ;)
'¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?

'¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿API?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?
'Blt
Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
'¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?


'¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?
'       [CODE 000]: MatuX
'
Public bRain        As Boolean 'está raineando?
Public bTecho       As Boolean 'hay techo?
Public brstTick     As Long

Private RLluvia(7)  As RECT  'RECT de la lluvia
Private iFrameIndex As Byte  'Frame actual de la LL
Private llTick      As Long  'Contador
Private LTLluvia(4) As Integer

Public charlist(1 To 10000) As Char

#If SeguridadAlkon Then

Public MI(1 To 1233) As clsManagerInvisibles
Public CualMI As Integer

#End If

'estados internos del surface (read only)
Public Enum TextureStatus
    tsOriginal = 0
    tsNight = 1
    tsFog = 2
End Enum

'[CODE 001]:MatuX
Public Enum PlayLoop
    plNone = 0
    plLluviain = 1
    plLluviaout = 2
End Enum
'[END]'
'
'       [END]
'¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?


Private Declare Function BltAlphaFast Lib "vbabdx" (ByRef lpDDSDest As Any, ByRef lpDDSSource As Any, ByVal iWidth As Long, ByVal iHeight As Long, _
        ByVal pitchSrc As Long, ByVal pitchDst As Long, ByVal dwMode As Long) As Long
Private Declare Function BltEfectoNoche Lib "vbabdx" (ByRef lpDDSDest As Any, ByVal iWidth As Long, ByVal iHeight As Long, _
        ByVal pitchDst As Long, ByVal dwMode As Long) As Long

Private Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long


Sub CargarTips()
On Error Resume Next
Dim N As Integer, i As Integer
Dim NumTips As Integer

N = FreeFile
Open ConfigDir.Inits & "\Tips.ayu" For Binary Access Read As #N

'cabecera
Get #N, , MiCabecera

'num de cabezas
Get #N, , NumTips

'Resize array
If NumTips = 0 Then Close #N: Exit Sub
ReDim Tips(1 To NumTips) As String * 255

For i = 1 To NumTips
    Get #N, , Tips(i)
Next i

Close #N

End Sub

Sub CargarArrayLluvia()
On Error Resume Next
Dim N As Integer, i As Integer
Dim Nu As Integer

N = FreeFile
If Not FileExist(ConfigDir.Inits & "\fk.ind", vbNormal) Then Exit Sub

Open ConfigDir.Inits & "\fk.ind" For Binary Access Read As #N

'cabecera
Get #N, , MiCabecera

'num de cabezas
Get #N, , Nu

'Resize array
ReDim bLluvia(1 To Nu) As Byte

For i = 1 To Nu
    Get #N, , bLluvia(i)
Next i

Close #N

End Sub
Sub ConvertCPtoTP(StartPixelLeft As Integer, StartPixelTop As Integer, ByVal CX As Single, ByVal CY As Single, tX As Integer, tY As Integer)
'******************************************
'Converts where the user clicks in the main window
'to a tile position
'******************************************
Dim HWindowX As Integer
Dim HWindowY As Integer

CX = CX - StartPixelLeft
CY = CY - StartPixelTop

HWindowX = (WindowTileWidth \ 2)
HWindowY = (WindowTileHeight \ 2)

'Figure out X and Y tiles
CX = (CX \ TilePixelWidth)
CY = (CY \ TilePixelHeight)

If CX > HWindowX Then
    CX = (CX - HWindowX)

Else
    If CX < HWindowX Then
        CX = (0 - (HWindowX - CX))
    Else
        CX = 0
    End If
End If

If CY > HWindowY Then
    CY = (0 - (HWindowY - CY))
Else
    If CY < HWindowY Then
        CY = (CY - HWindowY)
    Else
        CY = 0
    End If
End If

tX = UserPos.x + CX
tY = UserPos.y + CY

End Sub






Sub ResetCharInfo(ByVal CharIndex As Integer)

    charlist(CharIndex).Active = 0
    charlist(CharIndex).Criminal = 0
    charlist(CharIndex).FX = 0
    charlist(CharIndex).FxLoopTimes = 0
    charlist(CharIndex).Invisible = False

#If SeguridadAlkon Then
    Call MI(CualMI).ResetInvisible(CharIndex)
#End If

    charlist(CharIndex).Moving = 0
    charlist(CharIndex).muerto = False
    charlist(CharIndex).Nombre = ""
    charlist(CharIndex).ClanName = ""
    charlist(CharIndex).ClanID = 0
    charlist(CharIndex).pie = False
    charlist(CharIndex).Pos.x = 0
    charlist(CharIndex).Pos.y = 0
    charlist(CharIndex).UsandoArma = False

End Sub


Sub EraseChar(ByVal CharIndex As Integer)
On Error Resume Next

'*****************************************************************
'Erases a character from CharList and map
'*****************************************************************

charlist(CharIndex).Active = 0

'Update lastchar
If CharIndex = LastChar Then
    Do Until charlist(LastChar).Active = 1
        LastChar = LastChar - 1
        If LastChar = 0 Then Exit Do
    Loop
End If


MapData(charlist(CharIndex).Pos.x, charlist(CharIndex).Pos.y).CharIndex = 0

Call ResetCharInfo(CharIndex)

'Update NumChars
NumChars = NumChars - 1

End Sub

Sub InitGrh(ByRef Grh As Grh, ByVal GrhIndex As Long, Optional Started As Byte = 2)
'*****************************************************************
'Sets up a grh. MUST be done before rendering
'*****************************************************************
If GrhIndex = 0 Then Exit Sub
Grh.GrhIndex = GrhIndex

If Started = 2 Then
    If Grhdata(Grh.GrhIndex).NumFrames > 1 Then
        Grh.Started = 1
    Else
        Grh.Started = 0
    End If
Else
    If Grhdata(Grh.GrhIndex).NumFrames = 1 Then Started = 0
    Grh.Started = Started
End If

Grh.FrameCounter = 1

If Grh.Started Then

Else

End If
'[CODE 000]:MatuX
'
'  La linea generaba un error en la IDE, (no ocurría debido al
' on error)
'
'   Grh.SpeedCounter = GrhData(Grh.GrhIndex).Speed
'
If Grh.GrhIndex <> 0 Then Grh.Speed = Grhdata(Grh.GrhIndex).Speed
'
'[END]'

End Sub


Sub DDrawTransGrhIndextoSurface(Surface As DirectDrawSurface7, Grh As Long, ByVal x As Integer, ByVal y As Integer, center As Byte, Animate As Byte)
Dim CurrentGrh As Grh
Dim destRect As RECT
Dim sourceRect As RECT
Dim SurfaceDesc As DDSURFACEDESC2

With destRect
    .Left = x
    .Top = y
    .Right = .Left + Grhdata(Grh).pixelWidth
    .Bottom = .Top + Grhdata(Grh).pixelHeight
End With

Surface.GetSurfaceDesc SurfaceDesc

'Draw
If destRect.Left >= 0 And destRect.Top >= 0 And destRect.Right <= SurfaceDesc.lWidth And destRect.Bottom <= SurfaceDesc.lHeight Then
    With sourceRect
        .Left = Grhdata(Grh).sX
        .Top = Grhdata(Grh).sY
        .Right = .Left + Grhdata(Grh).pixelWidth
        .Bottom = .Top + Grhdata(Grh).pixelHeight
    End With
    
    Surface.BltFast destRect.Left, destRect.Top, SurfaceDB.Surface(Grhdata(Grh).FileNum), sourceRect, DDBLTFAST_SRCCOLORKEY Or DDBLTFAST_WAIT
   
End If

End Sub


Sub DrawBackBufferSurface()
    PrimarySurface.Blt MainViewRect, BackBufferSurface, MainDestRect, DDBLT_WAIT
End Sub

Function GetBitmapDimensions(BmpFile As String, ByRef bmWidth As Long, ByRef bmHeight As Long)
'*****************************************************************
'Gets the dimensions of a bmp
'*****************************************************************
Dim BMHeader As BITMAPFILEHEADER
Dim BINFOHeader As BITMAPINFOHEADER

Open BmpFile For Binary Access Read As #1
Get #1, , BMHeader
Get #1, , BINFOHeader
Close #1
bmWidth = BINFOHeader.biWidth
bmHeight = BINFOHeader.biHeight
End Function

Sub DrawGrhtoHdc(hWnd As Long, hdc As Long, FileNum As Integer, sourceRect As RECT, destRect As RECT)
    If FileNum <= 0 Then Exit Sub
    On Error Resume Next
    SecundaryClipper.SetHWnd hWnd
    
    SurfaceDB.Surface(FileNum).BltToDC hdc, sourceRect, destRect
    
End Sub


Sub LoadGraphics()
'**************************************************************
'Author: Juan Martín Sotuyo Dodero - complete rewrite
'Last Modify Date: 11/03/2006
'Initializes the SurfaceDB and sets up the rain rects
'**************************************************************
    'New surface manager :D
    Call SurfaceDB.Initialize(DirectDraw, True, ConfigDir.Graficos & "\")
          
    'Set up te rain rects
    RLluvia(0).Top = 0:      RLluvia(1).Top = 0:      RLluvia(2).Top = 0:      RLluvia(3).Top = 0
    RLluvia(0).Left = 0:     RLluvia(1).Left = 128:   RLluvia(2).Left = 256:   RLluvia(3).Left = 384
    RLluvia(0).Right = 128:  RLluvia(1).Right = 256:  RLluvia(2).Right = 384:  RLluvia(3).Right = 512
    RLluvia(0).Bottom = 128: RLluvia(1).Bottom = 128: RLluvia(2).Bottom = 128: RLluvia(3).Bottom = 128

    RLluvia(4).Top = 128:    RLluvia(5).Top = 128:    RLluvia(6).Top = 128:    RLluvia(7).Top = 128
    RLluvia(4).Left = 0:     RLluvia(5).Left = 128:   RLluvia(6).Left = 256:   RLluvia(7).Left = 384
    RLluvia(4).Right = 128:  RLluvia(5).Right = 256:  RLluvia(6).Right = 384:  RLluvia(7).Right = 512
    RLluvia(4).Bottom = 256: RLluvia(5).Bottom = 256: RLluvia(6).Bottom = 256: RLluvia(7).Bottom = 256
    
    'We are done!
End Sub

'[END]'
Function InitTileEngine(ByRef setDisplayFormhWnd As Long, setMainViewTop As Integer, setMainViewLeft As Integer, setTilePixelHeight As Integer, setTilePixelWidth As Integer, setWindowTileHeight As Integer, setWindowTileWidth As Integer, setTileBufferSize As Integer) As Boolean
'*****************************************************************
'InitEngine
'*****************************************************************
Dim SurfaceDesc As DDSURFACEDESC2
Dim ddck As DDCOLORKEY

ConfigDir.Inits = ConfigDir.Inits & "\"

'Set intial user position
UserPos.x = MinXBorder
UserPos.y = MinYBorder

'Fill startup variables

DisplayFormhWnd = setDisplayFormhWnd
MainViewTop = setMainViewTop
MainViewLeft = setMainViewLeft
TilePixelWidth = setTilePixelWidth
TilePixelHeight = setTilePixelHeight
WindowTileHeight = setWindowTileHeight
WindowTileWidth = setWindowTileWidth
TileBufferSize = setTileBufferSize

MinXBorder = XMinMapSize + (WindowTileWidth \ 2)
MaxXBorder = XMaxMapSize - (WindowTileWidth \ 2)
MinYBorder = YMinMapSize + (WindowTileHeight \ 2)
MaxYBorder = YMaxMapSize - (WindowTileHeight \ 2)

MainViewWidth = (TilePixelWidth * WindowTileWidth)
MainViewHeight = (TilePixelHeight * WindowTileHeight)


ReDim MapData(XMinMapSize To XMaxMapSize, YMinMapSize To YMaxMapSize) As MapBlock


DirectDraw.SetCooperativeLevel DisplayFormhWnd, DDSCL_NORMAL

'Primary Surface
' Fill the surface description structure
With SurfaceDesc
    .lFlags = DDSD_CAPS
    .ddsCaps.lCaps = DDSCAPS_PRIMARYSURFACE
End With


Set PrimarySurface = DirectDraw.CreateSurface(SurfaceDesc)

Set PrimaryClipper = DirectDraw.CreateClipper(0)
PrimaryClipper.SetHWnd frmMain.hWnd
PrimarySurface.SetClipper PrimaryClipper

Set SecundaryClipper = DirectDraw.CreateClipper(0)

With BackBufferRect
    .Left = 0
    .Top = 0
    .Right = TilePixelWidth * (WindowTileWidth + 2 * TileBufferSize)
    .Bottom = TilePixelHeight * (WindowTileHeight + 2 * TileBufferSize)
End With

With SurfaceDesc
    .lFlags = DDSD_CAPS Or DDSD_HEIGHT Or DDSD_WIDTH
    If True Then
        .ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN
    Else
        .ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN Or DDSCAPS_SYSTEMMEMORY
    End If
    .lHeight = BackBufferRect.Bottom
    .lWidth = BackBufferRect.Right
End With

Set BackBufferSurface = DirectDraw.CreateSurface(SurfaceDesc)

ddck.low = 0
ddck.high = 0
BackBufferSurface.SetColorKey DDCKEY_SRCBLT, ddck
Call CargarConfig

frmCargando.Show
DoEvents
Dim i As Long
For i = 1 To 5000
    DoEvents
Next i

frmCargando.lblEstado.Caption = "Cargando Grh..."

DoEvents
Call LoadGrhData

DoEvents
frmCargando.lblEstado.Caption = "Cargando Bodys..."

DoEvents
Call CargarCuerpos

frmCargando.lblEstado.Caption = "Cargando Cabezas..."

DoEvents
Call CargarCabezas


Call CargarCascos

frmCargando.lblEstado.Caption = "Cargando Superficies..."

DoEvents
Call CargarSuperficies
'Call CargarEspalda
'Call CargarBotas
frmCargando.lblEstado.Caption = "Cargando FXs..."

DoEvents
Call CargarFxs

DoEvents
'Call InitializeDats

LTLluvia(0) = 224
LTLluvia(1) = 352
LTLluvia(2) = 480
LTLluvia(3) = 608
LTLluvia(4) = 736

Call LoadGraphics

InitTileEngine = True

End Function


Sub CrearGrh(GrhIndex As Long, Index As Long)
ReDim Preserve Grh(1 To Index) As Grh
Grh(Index).FrameCounter = 1
Grh(Index).GrhIndex = GrhIndex
Grh(Index).Speed = Grhdata(GrhIndex).Speed
Grh(Index).Started = 1
End Sub

Sub CargarAnimsExtra()
'Call CrearGrh(6580, 1) 'Anim Invent
'Call CrearGrh(534, 2) 'Animacion de teleport
End Sub

Function ControlVelocidad(ByVal LastTime As Long) As Boolean
ControlVelocidad = (GetTickCount - LastTime > 20)
End Function


Sub dibujapj(Surface As DirectDrawSurface7, Grh As Grhdata)
On Error Resume Next
Dim r2 As RECT, auxr As RECT, auxr2 As RECT
Dim iGrhIndex As Long
Dim SurfaceDesc As DDSURFACEDESC2


If Grh.FileNum <= 0 Then Exit Sub

SurfaceDB.Surface(Grh.FileNum).GetSurfaceDesc SurfaceDesc

With r2
   .Left = Grh.sX
   .Top = Grh.sY
   If .Left + Grh.pixelWidth <= SurfaceDesc.lWidth Then
        .Right = .Left + Grh.pixelWidth
    Else
        .Right = SurfaceDesc.lWidth
    End If
    If .Top + Grh.pixelHeight <= SurfaceDesc.lHeight Then
        .Bottom = .Top + Grh.pixelHeight
    Else
        .Bottom = SurfaceDesc.lHeight
    End If
   If .Bottom > 990 Then .Bottom = 990
End With

With auxr
    .Left = 0
    .Top = 0
    .Right = Grh.pixelWidth
    .Bottom = Grh.pixelHeight
End With

If auxr.Bottom > 990 Then auxr.Bottom = 990
If auxr.Right > 1024 Then auxr.Right = 1024
auxr2 = auxr



Surface.BltFast 0, 0, SurfaceDB.Surface(Grh.FileNum), r2, DDBLTFAST_SRCCOLORKEY Or DDBLTFAST_WAIT
Surface.BltToDC frmMain.Visor.hdc, auxr2, auxr

End Sub

Sub dibujapjESpecial(Surface As DirectDrawSurface7, Grh As Grhdata, ByVal x As Integer, ByVal y As Integer)
On Error Resume Next
Dim r2 As RECT, auxr As RECT, auxr2 As RECT
Dim iGrhIndex As Long
Dim SurfaceDesc As DDSURFACEDESC2


If Grh.FileNum <= 0 Then Exit Sub

SurfaceDB.Surface(Grh.FileNum).GetSurfaceDesc SurfaceDesc

With r2
   .Left = Grh.sX
   .Top = Grh.sY
   If .Left + Grh.pixelWidth <= SurfaceDesc.lWidth Then
        .Right = .Left + Grh.pixelWidth
    Else
        .Right = SurfaceDesc.lWidth
    End If
    If .Top + Grh.pixelHeight <= SurfaceDesc.lHeight Then
        .Bottom = .Top + Grh.pixelHeight
    Else
        .Bottom = SurfaceDesc.lHeight
    End If
    If .Bottom > 990 Then .Bottom = 990
End With

With auxr
    .Left = 0
    .Top = 0
    .Right = Grh.pixelWidth
    .Bottom = Grh.pixelHeight
End With
auxr2 = auxr



Surface.BltFast x, y, SurfaceDB.Surface(Grh.FileNum), r2, DDBLTFAST_SRCCOLORKEY Or DDBLTFAST_WAIT
'Surface.BltToDC frmMain.Visor.hDC, auxr2, auxr
End Sub
Sub dibujaBMP(Surface As DirectDrawSurface7, FileNum As Integer)
On Error Resume Next
Dim r2 As RECT, auxr As RECT, auxr2 As RECT
Dim r As RECT
Dim iGrhIndex As Long
Dim SurfaceDesc As DDSURFACEDESC2
Dim ddsd As DDSURFACEDESC2
Dim ddck As DDCOLORKEY
Dim surfacecuadro  As DirectDrawSurface7
Dim ii As Long





If FileNum <= 0 Then Exit Sub

SurfaceDB.Surface(FileNum).GetSurfaceDesc SurfaceDesc

With r2
   .Left = 0
   .Top = 0
    .Right = SurfaceDesc.lWidth
    .Bottom = SurfaceDesc.lHeight
   If .Bottom = 1024 Then .Bottom = 990
End With
Debug.Print r2.Bottom
auxr = r2
auxr2 = auxr



Surface.BltFast 0, 0, SurfaceDB.Surface(FileNum), r2, DDBLTFAST_SRCCOLORKEY Or DDBLTFAST_WAIT

If DibujarIndexaciones.activo Then
    'ddsd.lFlags = DDSD_CAPS Or DDSD_HEIGHT Or DDSD_WIDTH
    'ddsd.ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN
    'ddsd.lWidth = DibujarIndexaciones.Ancho
    'ddsd.lHeight = DibujarIndexaciones.Alto
    'Set surfacecuadro = DirectDraw.CreateSurface(ddsd)
    'Call surfacecuadro.BltColorFill(r, vbBlack)
    'surfacecuadro.SetForeColor vbGreen
    'surfacecuadro.SetFillColor vbBlack
    'Call surfacecuadro.DrawBox(0, 0, DibujarIndexaciones.Ancho, DibujarIndexaciones.Alto)
    
    'ddck.high = 0
    'ddck.low = 0
    'Call surfacecuadro.SetColorKey(DDCKEY_SRCBLT, ddck)
    'Call surfacecuadro.GetSurfaceDesc(ddsd)
End If

If DibujarIndexaciones.activo Then
    Surface.SetForeColor vbGreen
    Surface.setDrawStyle DrawStyleConstants.vbDot
    For ii = 1 To DibujarIndexaciones.Total
    
    

        Surface.DrawBox DibujarIndexaciones.Inicios(ii).x, DibujarIndexaciones.Inicios(ii).y, DibujarIndexaciones.Inicios(ii).x + DibujarIndexaciones.Ancho, DibujarIndexaciones.Inicios(ii).y + DibujarIndexaciones.Alto
        
        'Surface.BltFast DibujarIndexaciones.Inicios(ii).X, DibujarIndexaciones.Inicios(ii).Y, surfacecuadro, r2, DDBLTFAST_SRCCOLORKEY Or DDBLTFAST_WAIT
        'Debug.Print DibujarIndexaciones.Inicios(ii).X & "  -  " & DibujarIndexaciones.Inicios(ii).Y
    Next ii
End If

Surface.BltToDC frmMain.Visor.hdc, auxr2, auxr

End Sub
Sub dibujarGrh2(ByRef Grh As Grhdata)
On Error Resume Next
Dim r As RECT

If DibujarFondo Then
    BackBufferSurface.BltColorFill r, ColorFondo
Else
    BackBufferSurface.BltColorFill r, 0
End If
Call dibujapj(BackBufferSurface, Grh)
'*************** *********************************** ***************

End Sub
Sub dibujarBMP2(ByRef Grh As Integer)
On Error Resume Next
Dim r As RECT

If DibujarFondo Then
    BackBufferSurface.BltColorFill r, ColorFondo
Else
    BackBufferSurface.BltColorFill r, 0
End If

Call dibujaBMP(BackBufferSurface, Grh)
'*************** *********************************** ***************
'If dibujarindexacion Then
    
'End If

End Sub

Sub DrawGrh()

End Sub

Sub dibujarGrh4(ByRef Grh As Grhdata, Optional ByVal ofx As Integer, Optional ByVal ofy As Integer)
On Error Resume Next
Dim r As RECT

If DibujarFondo Then
    BackBufferSurface.BltColorFill r, ColorFondo
Else
    BackBufferSurface.BltColorFill r, 0
End If
Call dibujapj1(BackBufferSurface, Grh, ofx, ofy)
'*************** *********************************** ***************

End Sub


Sub dibujapj1(Surface As DirectDrawSurface7, Grh As Grhdata, Optional ByVal ofx As Integer, Optional ByVal ofy As Integer)
On Error Resume Next
Dim r2 As RECT, auxr As RECT, auxr2 As RECT
Dim iGrhIndex As Long
Dim SurfaceDesc As DDSURFACEDESC2


If Grh.FileNum <= 0 Then Exit Sub

SurfaceDB.Surface(Grh.FileNum).GetSurfaceDesc SurfaceDesc

With r2
   .Left = Grh.sX
   .Top = Grh.sY
   If .Left + Grh.pixelWidth <= SurfaceDesc.lWidth Then
        .Right = .Left + Grh.pixelWidth
    Else
        .Right = SurfaceDesc.lWidth
    End If
    If .Top + Grh.pixelHeight <= SurfaceDesc.lHeight Then
        .Bottom = .Top + Grh.pixelHeight
    Else
        .Bottom = SurfaceDesc.lHeight
    End If
   If .Bottom > 990 Then .Bottom = 990
End With

With auxr
    .Left = 0
    .Top = 0
    .Right = Grh.pixelWidth + ofx
    .Bottom = Grh.pixelHeight + ofy
End With

If auxr.Bottom > 990 Then auxr.Bottom = 990
If auxr.Right > 1024 Then auxr.Right = 1024
auxr2 = auxr



Surface.BltFast ofx, ofy, SurfaceDB.Surface(Grh.FileNum), r2, DDBLTFAST_SRCCOLORKEY Or DDBLTFAST_WAIT
Surface.BltToDC frmDats.picGrh.hdc, auxr2, auxr

End Sub

Sub GetSlotOffset(ByVal Slot As Byte, ByRef offsetx As Integer, ByRef offsety As Integer)
    If Slot <= 5 Then
        offsetx = 32 * (Slot - 1)
        offsety = 0
    ElseIf Slot > 5 And Slot <= 10 Then
        offsetx = (32 * ((Slot - 5) - 1))
        offsety = 32
        
    ElseIf Slot > 10 And Slot <= 15 Then
        offsetx = (32 * ((Slot - 10) - 1))
        offsety = 64
    ElseIf Slot > 15 And Slot <= 20 Then
        offsetx = (32 * ((Slot - 15) - 1))
        offsety = 32 + 64
    ElseIf Slot > 20 And Slot <= 25 Then
        offsetx = (32 * ((Slot - 20) - 1))
        offsety = 128
    ElseIf Slot > 25 And Slot <= 30 Then
        offsetx = (32 * ((Slot - 25) - 1))
        offsety = 128 + 32
    End If
End Sub

Sub dibujaSlot(Surface As DirectDrawSurface7, Grh As Grhdata, Slot As Byte, cantidad As Integer)
On Error Resume Next
Dim r2 As RECT, auxr As RECT, auxr2 As RECT
Dim iGrhIndex As Long
Dim SurfaceDesc As DDSURFACEDESC2
Dim offsetx As Integer, offsety As Integer

If Grh.FileNum <= 0 Then Exit Sub

SurfaceDB.Surface(Grh.FileNum).GetSurfaceDesc SurfaceDesc

Call GetSlotOffset(Slot, offsetx, offsety)

With r2
   .Left = Grh.sX
   .Top = Grh.sY
   If .Left + Grh.pixelWidth <= SurfaceDesc.lWidth Then
        .Right = .Left + Grh.pixelWidth
    Else
        .Right = SurfaceDesc.lWidth
    End If
    If .Top + Grh.pixelHeight <= SurfaceDesc.lHeight Then
        .Bottom = .Top + Grh.pixelHeight
    Else
        .Bottom = SurfaceDesc.lHeight
    End If
   If .Bottom > 990 Then .Bottom = 990
End With

With auxr
    .Left = 0
    .Top = 0
    .Right = Grh.pixelWidth
    .Bottom = Grh.pixelHeight
End With

If auxr.Bottom > 990 Then auxr.Bottom = 990
If auxr.Right > 1024 Then auxr.Right = 1024
auxr2 = auxr


Surface.SetForeColor vbWhite

Surface.BltFast offsetx, offsety, SurfaceDB.Surface(Grh.FileNum), r2, DDBLTFAST_SRCCOLORKEY Or DDBLTFAST_WAIT
'Surface.SetFontBackColor vbWhite

'dib cantidad
Surface.DrawText offsetx, offsety, cantidad, False


End Sub



Sub dibujagrh(Surface As DirectDrawSurface7, Grh As Grhdata, x As Integer, y As Integer)
On Error Resume Next
Dim r2 As RECT, auxr As RECT, auxr2 As RECT
Dim iGrhIndex As Long
Dim SurfaceDesc As DDSURFACEDESC2

If Grh.FileNum <= 0 Then Exit Sub

SurfaceDB.Surface(Grh.FileNum).GetSurfaceDesc SurfaceDesc

With r2
   .Left = Grh.sX
   .Top = Grh.sY
   If .Left + Grh.pixelWidth <= SurfaceDesc.lWidth Then
        .Right = .Left + Grh.pixelWidth
    Else
        .Right = SurfaceDesc.lWidth
    End If
    If .Top + Grh.pixelHeight <= SurfaceDesc.lHeight Then
        .Bottom = .Top + Grh.pixelHeight
    Else
        .Bottom = SurfaceDesc.lHeight
    End If
   If .Bottom > 990 Then .Bottom = 990
End With

With auxr
    .Left = 0
    .Top = 0
    .Right = Grh.pixelWidth
    .Bottom = Grh.pixelHeight
End With

If auxr.Bottom > 990 Then auxr.Bottom = 990
If auxr.Right > 1024 Then auxr.Right = 1024
auxr2 = auxr


Surface.SetForeColor vbWhite

Surface.BltFast x, y, SurfaceDB.Surface(Grh.FileNum), r2, DDBLTFAST_SRCCOLORKEY Or DDBLTFAST_WAIT
'Surface.SetFontBackColor vbWhite

'dib cantidad
'Surface.DrawText offsetX, offsetY, cantidad, False


End Sub

