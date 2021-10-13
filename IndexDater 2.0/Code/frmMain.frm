VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmMain 
   Caption         =   "Indexador"
   ClientHeight    =   10575
   ClientLeft      =   165
   ClientTop       =   -4095
   ClientWidth     =   13755
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   10575
   ScaleWidth      =   13755
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "Abrir menu DATS"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   120
      TabIndex        =   45
      Top             =   120
      Width           =   4575
   End
   Begin VB.CommandButton BotonI 
      BackColor       =   &H000080FF&
      Caption         =   "Superficies"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   8
      Left            =   0
      TabIndex        =   44
      Top             =   5160
      Width           =   1215
   End
   Begin VB.PictureBox ProgressBar1 
      Height          =   735
      Left            =   4800
      ScaleHeight     =   675
      ScaleWidth      =   3795
      TabIndex        =   42
      Top             =   12720
      Visible         =   0   'False
      Width           =   3855
   End
   Begin VB.CommandButton CmbResourceFile 
      Caption         =   "Limpiar Memoria"
      Height          =   255
      Left            =   11040
      TabIndex        =   41
      Top             =   7320
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.CheckBox Checkcabeza 
      Caption         =   "cabeza"
      Height          =   195
      Left            =   1800
      TabIndex        =   40
      Top             =   8400
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Fondo verde"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1219
      Left            =   3000
      TabIndex        =   39
      Top             =   9120
      Width           =   1215
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   240
      Top             =   11880
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton BotonI 
      Caption         =   "Fx"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   6
      Left            =   0
      TabIndex        =   36
      Top             =   4680
      Width           =   1215
   End
   Begin VB.ComboBox CDibujarWalk 
      Height          =   315
      ItemData        =   "frmMain.frx":0CCE
      Left            =   1800
      List            =   "frmMain.frx":0CE1
      TabIndex        =   35
      Top             =   5880
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton BotonI 
      Caption         =   "Armas"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   5
      Left            =   0
      TabIndex        =   34
      Top             =   4200
      Width           =   1215
   End
   Begin VB.CommandButton BotonI 
      Caption         =   "Escudos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   4
      Left            =   0
      TabIndex        =   33
      Top             =   3720
      Width           =   1215
   End
   Begin VB.CommandButton BotonI 
      Caption         =   "Cascos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   3
      Left            =   0
      TabIndex        =   32
      Top             =   3240
      Width           =   1215
   End
   Begin VB.CommandButton BotonI 
      Caption         =   "Cabezas"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   2
      Left            =   0
      TabIndex        =   31
      Top             =   2760
      Width           =   1215
   End
   Begin VB.CommandButton BotonI 
      Caption         =   "Bodys"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   1
      Left            =   0
      TabIndex        =   30
      Top             =   2280
      Width           =   1215
   End
   Begin VB.CommandButton BotonI 
      Caption         =   "Grhs"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   0
      Left            =   0
      TabIndex        =   29
      Top             =   1800
      Width           =   1215
   End
   Begin VB.CommandButton NuevoGhr 
      Caption         =   "Nuevo/buscar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1200
      TabIndex        =   24
      Top             =   840
      Width           =   3495
   End
   Begin VB.CommandButton Command10 
      Caption         =   "Guardar Graficos.ind"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   480
      TabIndex        =   23
      Top             =   9480
      Width           =   2535
   End
   Begin VB.Timer Dibujado 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   0
      Top             =   6480
   End
   Begin VB.TextBox TextDatos 
      Enabled         =   0   'False
      Height          =   285
      Index           =   9
      Left            =   1800
      TabIndex        =   11
      Top             =   9480
      Width           =   1095
   End
   Begin VB.TextBox TextDatos 
      Enabled         =   0   'False
      Height          =   285
      Index           =   8
      Left            =   1800
      TabIndex        =   10
      Top             =   9120
      Width           =   1095
   End
   Begin VB.TextBox TextDatos 
      Height          =   285
      Index           =   7
      Left            =   1800
      TabIndex        =   9
      Top             =   8760
      Width           =   1095
   End
   Begin VB.TextBox TextDatos 
      Height          =   285
      Index           =   6
      Left            =   1800
      TabIndex        =   8
      Top             =   8400
      Width           =   1095
   End
   Begin VB.TextBox TextDatos 
      Height          =   285
      Index           =   5
      Left            =   1800
      TabIndex        =   7
      Top             =   8040
      Width           =   1095
   End
   Begin VB.TextBox TextDatos 
      Height          =   285
      Index           =   4
      Left            =   1800
      TabIndex        =   6
      Top             =   7680
      Width           =   1095
   End
   Begin VB.TextBox TextDatos 
      Height          =   285
      Index           =   3
      Left            =   1800
      TabIndex        =   5
      Top             =   7320
      Width           =   1095
   End
   Begin VB.TextBox TextDatos 
      Height          =   285
      Index           =   2
      Left            =   1800
      TabIndex        =   4
      Top             =   6960
      Width           =   1095
   End
   Begin VB.TextBox TextDatos 
      Height          =   285
      Index           =   1
      Left            =   1200
      ScrollBars      =   1  'Horizontal
      TabIndex        =   3
      Top             =   5880
      Width           =   3015
   End
   Begin VB.TextBox TextDatos 
      Height          =   285
      Index           =   0
      Left            =   1800
      TabIndex        =   2
      Top             =   6600
      Width           =   1095
   End
   Begin VB.ListBox Lista 
      Height          =   4350
      ItemData        =   "frmMain.frx":0D05
      Left            =   1200
      List            =   "frmMain.frx":0D07
      TabIndex        =   1
      Top             =   1320
      Width           =   3495
   End
   Begin VB.PictureBox Visor 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   10335
      Left            =   5040
      ScaleHeight     =   10305
      ScaleWidth      =   14385
      TabIndex        =   0
      Top             =   120
      Width           =   14415
   End
   Begin VB.CommandButton BotonBorrrar 
      Caption         =   "Borrar"
      Height          =   375
      Left            =   2280
      TabIndex        =   27
      Top             =   9120
      Width           =   735
   End
   Begin VB.CommandButton BotonGuardar 
      Caption         =   "Guardar"
      Height          =   375
      Left            =   480
      TabIndex        =   22
      Top             =   9120
      Width           =   1815
   End
   Begin VB.CommandButton BotonI 
      Caption         =   "Graficos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   7
      Left            =   0
      TabIndex        =   37
      Top             =   1320
      Width           =   1215
   End
   Begin VB.Label Progress 
      Alignment       =   2  'Center
      Caption         =   "0%"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4920
      TabIndex        =   43
      Top             =   12360
      Visible         =   0   'False
      Width           =   3855
   End
   Begin VB.Label DescripcionAyuda 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Height          =   735
      Left            =   360
      TabIndex        =   38
      Top             =   8400
      Width           =   2655
   End
   Begin VB.Label LUlitError 
      BackColor       =   &H80000004&
      Caption         =   " AACA"
      Height          =   495
      Left            =   600
      TabIndex        =   28
      Top             =   11160
      Width           =   10935
   End
   Begin VB.Label LGHRnumeroA 
      Caption         =   "0"
      Height          =   255
      Left            =   1200
      TabIndex        =   26
      Top             =   6240
      Width           =   1095
   End
   Begin VB.Label LNumActual 
      Caption         =   "Ghr:"
      Height          =   255
      Left            =   600
      TabIndex        =   25
      Top             =   6240
      Width           =   375
   End
   Begin VB.Label LTexto 
      Caption         =   "Ancho Titles:"
      Height          =   255
      Index           =   9
      Left            =   480
      TabIndex        =   21
      Top             =   9480
      Width           =   1215
   End
   Begin VB.Label LTexto 
      Caption         =   "Alto Titles:"
      Height          =   255
      Index           =   8
      Left            =   480
      TabIndex        =   20
      Top             =   9120
      Width           =   1215
   End
   Begin VB.Label LTexto 
      Caption         =   "PosicionY:"
      Height          =   255
      Index           =   7
      Left            =   480
      TabIndex        =   19
      Top             =   8760
      Width           =   1215
   End
   Begin VB.Label LTexto 
      Caption         =   "PosicionX:"
      Height          =   255
      Index           =   6
      Left            =   480
      TabIndex        =   18
      Top             =   8400
      Width           =   1215
   End
   Begin VB.Label LTexto 
      Caption         =   "Velocidad:"
      Height          =   255
      Index           =   5
      Left            =   480
      TabIndex        =   17
      Top             =   8040
      Width           =   1215
   End
   Begin VB.Label LTexto 
      Caption         =   "Ancho:"
      Height          =   255
      Index           =   4
      Left            =   480
      TabIndex        =   16
      Top             =   7680
      Width           =   1215
   End
   Begin VB.Label LTexto 
      Caption         =   "Alto:"
      Height          =   255
      Index           =   3
      Left            =   480
      TabIndex        =   15
      Top             =   7320
      Width           =   1215
   End
   Begin VB.Label LTexto 
      Caption         =   "Numero Frames:"
      Height          =   255
      Index           =   2
      Left            =   480
      TabIndex        =   14
      Top             =   6960
      Width           =   1215
   End
   Begin VB.Label LTexto 
      Caption         =   "Frames:"
      Height          =   255
      Index           =   1
      Left            =   480
      TabIndex        =   13
      Top             =   5880
      Width           =   735
   End
   Begin VB.Label LTexto 
      Caption         =   "Numero BMP:"
      Height          =   255
      Index           =   0
      Left            =   480
      TabIndex        =   12
      Top             =   6600
      Width           =   1215
   End
   Begin VB.Menu MenuArchivo 
      Caption         =   "Archivo"
      Begin VB.Menu MenuArchivoGuardar 
         Caption         =   "Guardar"
         Shortcut        =   ^S
      End
      Begin VB.Menu MenuBotonGuardarP 
         Caption         =   "Guardar..."
         Visible         =   0   'False
      End
      Begin VB.Menu mnuGuardarTodo 
         Caption         =   "Guardar todo"
      End
      Begin VB.Menu mnuConfig 
         Caption         =   "Reconfigurar directorios"
      End
      Begin VB.Menu MenuArchivoCargar 
         Caption         =   "Cargar"
         Shortcut        =   ^O
         Visible         =   0   'False
      End
      Begin VB.Menu MenuBotonCargarP 
         Caption         =   "Cargar..."
         Visible         =   0   'False
      End
   End
   Begin VB.Menu Medicion 
      Caption         =   "Edicion"
      Visible         =   0   'False
      Begin VB.Menu MenuEdicionNuevo 
         Caption         =   "Nuevo/Ir A"
         Shortcut        =   ^F
      End
      Begin VB.Menu menuEdicionMover 
         Caption         =   "Mover"
      End
      Begin VB.Menu MenuEdicionCopiar 
         Caption         =   "Copiar"
      End
      Begin VB.Menu MenuEdicionBorrar 
         Caption         =   "Borrar"
         Shortcut        =   {DEL}
      End
      Begin VB.Menu MenuEdicionClonar 
         Caption         =   "Clonar..."
      End
      Begin VB.Menu menuEdicionColor 
         Caption         =   "Color de fondo..."
      End
   End
   Begin VB.Menu MenuHerramientas 
      Caption         =   "Herramientas"
      Begin VB.Menu MenuHerramientasBG 
         Caption         =   "Buscar Grh Con bmp..."
      End
      Begin VB.Menu MenuHerramientasNI 
         Caption         =   "Buscar Bmps sin indexar"
      End
      Begin VB.Menu MenuHerramientasBN 
         Caption         =   "Buscar siguiente"
         Shortcut        =   {F3}
         Visible         =   0   'False
      End
      Begin VB.Menu MenuHerramientasAAnim 
         Caption         =   "Autoindexador"
      End
      Begin VB.Menu MenuHerramientasBR 
         Caption         =   "Buscar Grh Repetidos"
      End
      Begin VB.Menu mnuRecargarResource 
         Caption         =   "Recargar lista de graficos .BMP"
      End
   End
   Begin VB.Menu MenuAyuda 
      Caption         =   "Ayuda"
      Begin VB.Menu MenuAcercaDe 
         Caption         =   "Acerca de..."
      End
   End
   Begin VB.Menu mnuautoI 
      Caption         =   "Indexar como..."
      Visible         =   0   'False
      Begin VB.Menu IAnim 
         Caption         =   "Indexar como animacion/Fxs"
      End
      Begin VB.Menu mnIgeneral 
         Caption         =   "Indexar como superficies, adornos y graficos unicos"
      End
      Begin VB.Menu mnuibody 
         Caption         =   "Indexar como espada/escudo/body/npc"
      End
      Begin VB.Menu mnuInvObj 
         Caption         =   "Indexar como objeto de inventario"
      End
   End
   Begin VB.Menu mnuDatBodys 
      Caption         =   "Datear body como..."
      Visible         =   0   'False
      Begin VB.Menu mnuDatRopa 
         Caption         =   "Datear Armadura/Tunica/Ropaje"
      End
      Begin VB.Menu mnuDatNPC 
         Caption         =   "Datear NPC"
      End
   End
   Begin VB.Menu mnuDat 
      Caption         =   "Datear..."
      Visible         =   0   'False
      Begin VB.Menu mnuDatObjeto 
         Caption         =   "Datear objeto"
      End
   End
   Begin VB.Menu mnuEspada 
      Caption         =   "Datear"
      Visible         =   0   'False
      Begin VB.Menu mnuDatEspadas 
         Caption         =   "Datear  espada"
      End
   End
   Begin VB.Menu mnuEscudo 
      Caption         =   "Datear"
      Visible         =   0   'False
      Begin VB.Menu mnuDatEscudo 
         Caption         =   "Datear Escudo"
      End
   End
   Begin VB.Menu mnuCasco 
      Caption         =   "Datear"
      Visible         =   0   'False
      Begin VB.Menu mnuDatCasco 
         Caption         =   "Datear Casco"
      End
   End
   Begin VB.Menu mnuHechizo 
      Caption         =   "Datear"
      Visible         =   0   'False
      Begin VB.Menu mnuDatHechizo 
         Caption         =   "Datear hechizo"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub DibujarGHRVisor(ByVal GhrIndex As Integer)
On Error Resume Next
If Not GrHCambiando Then
    If GhrIndex <= 0 Then Exit Sub
    Call dibujarGrh2(Grhdata(GhrIndex))
    frmMain.Visor.Refresh
Else
    Call dibujarGrh2(TempGrh)
    frmMain.Visor.Refresh
End If
End Sub
Private Sub DibujarBMPVisor(ByVal GhrIndex As Integer)
Dim SR As RECT, DR As RECT
Dim Alto As Long
Dim Ancho As Long
frmMain.Visor.Cls
Dim dummy As Integer
    If GhrIndex <= 0 Then Exit Sub
    
'    Call GetTamañoBMP(GhrIndex, Alto, Ancho, dummy)
'    SR.Left = 0
'    SR.Top = 0
'    SR.Right = Ancho
'    SR.Bottom = Alto

'    DR.Left = 0
'    DR.Top = 0
'    DR.Right = SR.Right
'    DR.Bottom = SR.Bottom
'    Call DrawGrhtoHdc(frmMain.Visor.hWnd, frmMain.Visor.hDC, GhrIndex, SR, DR)
    Call dibujarBMP2(GhrIndex)
    frmMain.Visor.Refresh
End Sub
Private Sub DibujarDataIndex(ByRef DataIndex As BodyData, Optional ByVal frame As Integer = 1, Optional ByVal Index As Byte = 0)
On Error Resume Next
Dim SR As RECT, DR As RECT
Dim r As RECT
Dim sourceRect As RECT, destRect As RECT

Dim i As Long
Dim curx As Long
Dim cury As Long
Dim GhrIndex(1 To 4) As Grh
Dim Posiciones(1 To 4) As Position
Dim tGrhIndex As Long
curx = 0
cury = 0
If EstadoIndexador = e_EstadoIndexador.FX Then
    Index = 1
End If
With sourceRect
    .Bottom = 500
    .Left = 0
    .Right = 500
    .Top = 0
End With



If (Index > 0 And Index < 5) Or EstadoIndexador = e_EstadoIndexador.FX Then
        If DataIndex.Walk(Index).GrhIndex <= 0 Then DataIndex.Walk(Index).GrhIndex = 1
        If Grhdata(DataIndex.Walk(Index).GrhIndex).NumFrames > 1 Then
            tGrhIndex = Grhdata(DataIndex.Walk(Index).GrhIndex).Frames(frame)
        Else
            tGrhIndex = DataIndex.Walk(Index).GrhIndex
        End If
        If tGrhIndex <= 0 Then Exit Sub
        
'        SR.Left = Grhdata(tGrhIndex).sX
'        SR.Top = Grhdata(tGrhIndex).sY
'        SR.Right = Grhdata(tGrhIndex).sX + Grhdata(tGrhIndex).pixelWidth
'        SR.Bottom = Grhdata(tGrhIndex).sY + Grhdata(tGrhIndex).pixelHeight
'
'        DR.Left = CurX
'        DR.Top = CurY
'        DR.Right = CurX + Grhdata(tGrhIndex).pixelWidth
'        DR.Bottom = CurY + Grhdata(tGrhIndex).pixelHeight
        Call dibujarGrh2(Grhdata(tGrhIndex))
        'Call DrawGrhtoHdc(frmMain.Visor.hWnd, frmMain.Visor.hDC, Grhdata(tGrhIndex).FileNum, SR, DR)
        frmMain.Visor.Refresh
Else
    If DibujarFondo Then
        BackBufferSurface.BltColorFill r, ColorFondo
    Else
        BackBufferSurface.BltColorFill r, 0
    End If
    Call CalcularPosiciones(DataIndex, Posiciones)
    For i = 1 To 4
        If DataIndex.Walk(i).GrhIndex <= 0 Then DataIndex.Walk(i).GrhIndex = 1
        If Grhdata(DataIndex.Walk(i).GrhIndex).NumFrames > 1 Then
            tGrhIndex = Grhdata(DataIndex.Walk(i).GrhIndex).Frames(frame)
        Else
            tGrhIndex = DataIndex.Walk(i).GrhIndex
        End If
        If tGrhIndex <= 0 Then Exit Sub
        
        SR.Left = Grhdata(tGrhIndex).sX
        SR.Top = Grhdata(tGrhIndex).sY
        SR.Right = Grhdata(tGrhIndex).sX + Grhdata(tGrhIndex).pixelWidth
        SR.Bottom = Grhdata(tGrhIndex).sY + Grhdata(tGrhIndex).pixelHeight
        
        DR.Left = Posiciones(i).x
        DR.Top = Posiciones(i).y
        DR.Right = Posiciones(i).x + Grhdata(tGrhIndex).pixelWidth
        DR.Bottom = Posiciones(i).y + Grhdata(tGrhIndex).pixelHeight
        

        Call dibujapjESpecial(BackBufferSurface, Grhdata(tGrhIndex), DR.Left, DR.Top)
        
    Next i
    If EstadoIndexador = e_EstadoIndexador.Body And cabezaActual <> 0 Then
        If cabezaActual > 0 And cabezaActual <= MAXGrH Then
            Call dibujapjESpecial(BackBufferSurface, Grhdata(cabezaActual), Posiciones(3).x + (Grhdata(Grhdata(DataIndex.Walk(3).GrhIndex).Frames(frame)).pixelWidth / 2) - (Grhdata(cabezaActual).pixelWidth / 2) + DataIndex.HeadOffset.x, Posiciones(3).y + Grhdata(Grhdata(DataIndex.Walk(3).GrhIndex).Frames(frame)).pixelHeight - Grhdata(cabezaActual).pixelHeight + DataIndex.HeadOffset.y - 1)
        End If
    End If
    If EstadoIndexador = e_EstadoIndexador.Cabezas Then
        If frmMain.Checkcabeza.Value = vbChecked Then
            cabezaActual = DataIndex.Walk(3).GrhIndex
        End If
    End If
    'SecundaryClipper.SetHWnd frmMain.Visor.hWnd
    If DataIndex.Walk(4).GrhIndex > 0 Then
        sourceRect.Right = Posiciones(2).x + Grhdata(DataIndex.Walk(4).GrhIndex).pixelWidth
    Else
         sourceRect.Right = Posiciones(2).x * 2
    End If
    If DataIndex.Walk(3).GrhIndex > 0 Then
        sourceRect.Bottom = Posiciones(3).y + Grhdata(DataIndex.Walk(3).GrhIndex).pixelHeight
    Else
        sourceRect.Bottom = Posiciones(3).y * 2
    End If
    If sourceRect.Bottom > 990 Then sourceRect.Bottom = 990
    destRect = sourceRect
    BackBufferSurface.BltToDC frmMain.Visor.hdc, sourceRect, destRect
    
    frmMain.Visor.Refresh
End If
End Sub
Private Sub DibujarTempGHRVisor(ByVal loopAnim As Integer)
On Error Resume Next
Dim SR As RECT, DR As RECT
Dim GhrIndex As Integer
GhrIndex = loopAnim
    If GhrIndex <= 0 Then Exit Sub
    
    SR.Left = Grhdata(GhrIndex).sX
    SR.Top = Grhdata(GhrIndex).sY
    SR.Right = Grhdata(GhrIndex).sX + Grhdata(GhrIndex).pixelWidth
    SR.Bottom = Grhdata(GhrIndex).sY + Grhdata(GhrIndex).pixelHeight
    
    DR.Left = 0
    DR.Top = 0
    DR.Right = Grhdata(GhrIndex).pixelWidth
    DR.Bottom = Grhdata(GhrIndex).pixelHeight
    Call DrawGrhtoHdc(frmMain.Visor.hWnd, frmMain.Visor.hdc, CInt(Grhdata(GhrIndex).FileNum), SR, DR)
    frmMain.Visor.Refresh
End Sub


Private Sub GetInfoGHR(ByVal GrhIndex As Long)
If GrhIndex <= 0 Then Exit Sub
LoadingNew = True
Dim i As Long
Dim Ancho As Long
Dim Alto As Long
Dim PrimerAlto As Long
Dim PrimerAncho As Long
Dim dumYin As Integer


TextDatos(0).Text = Grhdata(GrhIndex).FileNum
TextDatos(1).Text = ""
TextDatos(2).Text = Grhdata(GrhIndex).NumFrames

TextDatos(3).Text = Grhdata(GrhIndex).pixelHeight
TextDatos(4).Text = Grhdata(GrhIndex).pixelWidth
TextDatos(5).Text = Grhdata(GrhIndex).Speed
TextDatos(6).Text = Grhdata(GrhIndex).sX
TextDatos(7).Text = Grhdata(GrhIndex).sY
TextDatos(8).Text = Grhdata(GrhIndex).TileHeight
TextDatos(9).Text = Grhdata(GrhIndex).TileWidth
LUlitError.Caption = ""
If Grhdata(GrhIndex).NumFrames = 1 Then
    TextDatos(1).BackColor = vbWhite
    TextDatos(1).Text = Grhdata(GrhIndex).Frames(1)
    Call GetTamañoBMP(Grhdata(GrhIndex).FileNum, Alto, Ancho, dumYin)
    frmMain.Dibujado.Enabled = False
    TextDatos(1).Enabled = False
    For i = 3 To 4
        TextDatos(i).Enabled = True
    Next i
    TextDatos(5).Enabled = False
    For i = 6 To 7
        TextDatos(i).Enabled = True
    Next i
Else
    TextDatos(1).BackColor = vbWhite
    For i = 1 To Grhdata(GrhIndex).NumFrames
        If i = 1 Then
            TextDatos(1).Text = Grhdata(GrhIndex).Frames(i)
        Else
            TextDatos(1).Text = TextDatos(1).Text & "-" & Grhdata(GrhIndex).Frames(i)
        End If

    Next i
    If Grhdata(GrhIndex).Speed > 0 Then ' pervenimos division por 0
        frmMain.Dibujado.Interval = 50 '* Grhdata(GrhIndex).Speed
    Else
        frmMain.Dibujado.Interval = 100
    End If
    frmMain.Dibujado.Enabled = True
    TextDatos(1).Enabled = True
    For i = 3 To 4
        TextDatos(i).Enabled = False
        TextDatos(i).BackColor = vbWhite
    Next i
    TextDatos(5).Enabled = True
    For i = 6 To 7
        TextDatos(i).Enabled = False
        TextDatos(i).BackColor = vbWhite
    Next i

End If

    GrHCambiando = False
    LNumActual.Caption = "Ghr:"
    BotonGuardar.Visible = False
    LoadingNew = False
End Sub

Private Sub GetInfoBmp(ByVal GrhIndex As Long)
If GrhIndex <= 0 Then Exit Sub
Dim i As Long
Dim Ancho As Long
Dim Alto As Long
Dim PrimerAlto As Long
Dim PrimerAncho As Long
Dim BitCount As Integer
Dim existenciaBMP As Byte
Dim ResourceS As String

existenciaBMP = ExisteBMP(GrhIndex)
If existenciaBMP = 0 Then Exit Sub
If existenciaBMP = 1 And ResourceFile = 3 Then
    If GrhIndex > 0 And GrhIndex <= UBound(ResourceF.Graficos) Then
        If ResourceF.Graficos(GrhIndex).tamaño > 0 Then ResourceS = "+ ResF"
    End If
End If

If existenciaBMP = 2 Then TextDatos(0).Text = ResourceF.Graficos(GrhIndex).tamaño
TextDatos(1).Text = ""
TextDatos(2).Text = Alto

TextDatos(3).Text = Ancho
TextDatos(4).Text = BitCount
TextDatos(5).Text = StringRecurso(existenciaBMP)
If ResourceS <> vbNullString Then TextDatos(5).Text = TextDatos(5).Text & ResourceS

    LNumActual.Caption = "BMP:"
    BotonGuardar.Visible = False
End Sub


Private Sub GetInfoDataIndex(ByVal DataIndex As Integer)
If DataIndex <= 0 Then Exit Sub
Dim i As Long
Dim Ancho As Long
Dim Alto As Long
Dim PrimerAlto As Long
Dim PrimerAncho As Long

LoadingNew = True
Dim GhrIndex(1 To 4) As Integer
Dim tGrhIndex As Long
TextDatos(5).Visible = False
LTexto(5).Visible = False
TextDatos(5).Text = ""
LUlitError.Caption = ""
For i = 1 To 4
    If EstadoIndexador = e_EstadoIndexador.Body Then
        GhrIndex(i) = BodyData(DataIndex).Walk(i).GrhIndex
        tempDataIndex.Walk(i).GrhIndex = BodyData(DataIndex).Walk(i).GrhIndex
    ElseIf EstadoIndexador = e_EstadoIndexador.Cabezas Then
        GhrIndex(i) = HeadData(DataIndex).Head(i).GrhIndex
        tempDataIndex.Walk(i).GrhIndex = HeadData(DataIndex).Head(i).GrhIndex
    ElseIf EstadoIndexador = e_EstadoIndexador.Cascos Then
        GhrIndex(i) = CascoAnimData(DataIndex).Head(i).GrhIndex
        tempDataIndex.Walk(i).GrhIndex = CascoAnimData(DataIndex).Head(i).GrhIndex
    ElseIf EstadoIndexador = e_EstadoIndexador.Escudos Then
        GhrIndex(i) = ShieldAnimData(DataIndex).ShieldWalk(i).GrhIndex
        tempDataIndex.Walk(i).GrhIndex = ShieldAnimData(DataIndex).ShieldWalk(i).GrhIndex
    ElseIf EstadoIndexador = e_EstadoIndexador.Armas Then
        GhrIndex(i) = WeaponAnimData(DataIndex).WeaponWalk(i).GrhIndex
        tempDataIndex.Walk(i).GrhIndex = WeaponAnimData(DataIndex).WeaponWalk(i).GrhIndex
    ElseIf EstadoIndexador = e_EstadoIndexador.FX Then
        GhrIndex(i) = FxData(DataIndex).FX.GrhIndex
        tempDataIndex.Walk(i).GrhIndex = FxData(DataIndex).FX.GrhIndex
    End If
Next i

TextDatos(0).Text = GhrIndex(1)
TextDatos(2).Text = GhrIndex(2)

TextDatos(3).Text = GhrIndex(3)
TextDatos(4).Text = GhrIndex(4)
If EstadoIndexador = e_EstadoIndexador.Body Then
    TextDatos(5).Text = BodyData(DataIndex).HeadOffset.y & "º" & BodyData(DataIndex).HeadOffset.x
    tempDataIndex.HeadOffset.x = BodyData(DataIndex).HeadOffset.x
    tempDataIndex.HeadOffset.y = BodyData(DataIndex).HeadOffset.y
    TextDatos(5).Visible = True
    LTexto(5).Visible = True
ElseIf EstadoIndexador = e_EstadoIndexador.FX Then
    TextDatos(2).Text = FxData(DataIndex).offsety & "º" & FxData(DataIndex).offsetx
    tempDataIndex.HeadOffset.x = FxData(DataIndex).offsetx
    tempDataIndex.HeadOffset.y = FxData(DataIndex).offsety
    TextDatos(2).Visible = True
    LTexto(2).Visible = True
End If
    GrHCambiando = False
    BotonGuardar.Visible = False
    
LoadingNew = False
End Sub



Private Sub BotonBorrrar_Click()
Call SBotonBorrrar
End Sub

Public Sub CambiarEstado(ByVal Index As Integer)
' Cambia el estado del indexador entre las distintas secciones. Oculta/cambia labels
On Error Resume Next
Dim i As Long
    EstadoIndexador = Index
    Dibujado.Enabled = False
    Visor.Cls
    Lista.Clear
    GrHCambiando = False
    CDibujarWalk.Visible = False
    LUlitError.Caption = ""
    MenuEdicionClonar.Visible = False
    MenuHerramientas.Visible = False
    Command10.Visible = True
    BotonBorrrar.Visible = True
    DescripcionAyuda.Visible = False
    Checkcabeza.Visible = False
    Select Case EstadoIndexador
        Case e_EstadoIndexador.Grh
            Call RenuevaListaGrH   'mostramos lista de grhs
            For i = 0 To 9
                TextDatos(i).Visible = True
                TextDatos(i).BackColor = vbWhite
                LTexto(i).Visible = True
            Next i
            
            MenuHerramientas.Visible = True
            MenuHerramientasBN.Visible = False
            MenuEdicionClonar.Visible = True
            LNumActual.Caption = "Grh: "
            LTexto(0).Caption = "Numero BMP:"
            LTexto(1).Caption = "Frames:"
            LTexto(2).Caption = "Numero Frames:"
            LTexto(3).Caption = "Alto:"
            LTexto(4).Caption = "Ancho:"
            LTexto(5).Caption = "Velocidad:"
            LTexto(6).Caption = "PosicionX:"
            LTexto(7).Caption = "PosicionY:"
            LTexto(8).Caption = "Alto Titles:"
            LTexto(9).Caption = "Ancho Titles:"
            LUlitError.Caption = ""
        Case e_EstadoIndexador.Body
            Checkcabeza.Visible = True
            CDibujarWalk.Visible = True
            CDibujarWalk.listIndex = 0
            Call RenuevaListaBodys
            For i = 0 To 5
                TextDatos(i).Visible = True
                TextDatos(i).BackColor = vbWhite
                TextDatos(i).Enabled = True
            Next i
            LTexto(1).Visible = False
            TextDatos(1).Visible = False
            For i = 6 To 9
                TextDatos(i).Visible = False
                TextDatos(i).BackColor = vbWhite
            Next i
            LNumActual.Caption = "Body: "
            LTexto(0).Caption = "Arriba:"
            LTexto(1).Caption = ""
            LTexto(2).Caption = "Derecha:"
            LTexto(3).Caption = "Abajo:"
            LTexto(4).Caption = "Izquierda:"
            LTexto(5).Caption = "Offset"
            LTexto(6).Caption = ""
            LTexto(7).Caption = ""
            LTexto(8).Caption = ""
            LTexto(9).Caption = ""
            LUlitError.Caption = ""
        Case e_EstadoIndexador.Cabezas
            CDibujarWalk.Visible = True
            CDibujarWalk.listIndex = 0
            Call RenuevaListaCabezas
            For i = 0 To 4
                TextDatos(i).Visible = True
                TextDatos(i).BackColor = vbWhite
                TextDatos(i).Enabled = True
            Next i
            LTexto(1).Visible = False
            TextDatos(1).Visible = False
            For i = 5 To 9
                TextDatos(i).Visible = False
                TextDatos(i).BackColor = vbWhite
            Next i
            LNumActual.Caption = "Cabeza: "
            LTexto(0).Caption = "Arriba:"
            LTexto(1).Caption = ""
            LTexto(2).Caption = "Derecha:"
            LTexto(3).Caption = "Abajo:"
            LTexto(4).Caption = "Izquierda:"
            LTexto(5).Caption = ""
            LTexto(6).Caption = ""
            LTexto(7).Caption = ""
            LTexto(8).Caption = ""
            LTexto(9).Caption = ""
            LUlitError.Caption = ""
        Case e_EstadoIndexador.Cascos
            CDibujarWalk.Visible = True
            CDibujarWalk.listIndex = 0
            Call RenuevaListaCascos
            For i = 0 To 4
                TextDatos(i).Visible = True
                TextDatos(i).BackColor = vbWhite
                TextDatos(i).Enabled = True
            Next i
            LTexto(1).Visible = False
            TextDatos(1).Visible = False
            For i = 5 To 9
                TextDatos(i).Visible = False
                TextDatos(i).BackColor = vbWhite
            Next i
            LNumActual.Caption = "Casco: "
            LTexto(0).Caption = "Arriba:"
            LTexto(1).Caption = ""
            LTexto(2).Caption = "Derecha:"
            LTexto(3).Caption = "Abajo:"
            LTexto(4).Caption = "Izquierda:"
            LTexto(5).Caption = ""
            LTexto(6).Caption = ""
            LTexto(7).Caption = ""
            LTexto(8).Caption = ""
            LTexto(9).Caption = ""
            LUlitError.Caption = ""
        Case e_EstadoIndexador.Escudos
            CDibujarWalk.Visible = True
            CDibujarWalk.listIndex = 0
             Call RenuevaListaEscudos
             For i = 0 To 4
                TextDatos(i).Visible = True
                TextDatos(i).BackColor = vbWhite
                TextDatos(i).Enabled = True
            Next i
            LTexto(1).Visible = False
            TextDatos(1).Visible = False
            For i = 5 To 9
                TextDatos(i).Visible = False
                TextDatos(i).BackColor = vbWhite
            Next i
            LNumActual.Caption = "Escudo: "
            LTexto(0).Caption = "Arriba:"
            LTexto(1).Caption = ""
            LTexto(2).Caption = "Derecha:"
            LTexto(3).Caption = "Abajo:"
            LTexto(4).Caption = "Izquierda:"
            LTexto(5).Caption = ""
            LTexto(6).Caption = ""
            LTexto(7).Caption = ""
            LTexto(8).Caption = ""
            LTexto(9).Caption = ""
            LUlitError.Caption = ""
        Case e_EstadoIndexador.Armas
            CDibujarWalk.Visible = True
            CDibujarWalk.listIndex = 0
            Call RenuevaListaArmas
            For i = 0 To 4
                TextDatos(i).Visible = True
                TextDatos(i).BackColor = vbWhite
                TextDatos(i).Enabled = True
            Next i
            LTexto(1).Visible = False
            TextDatos(1).Visible = False
            For i = 5 To 9
                TextDatos(i).Visible = False
                TextDatos(i).BackColor = vbWhite
            Next i
            LNumActual.Caption = "Armas: "
            LTexto(0).Caption = "Arriba:"
            LTexto(1).Caption = ""
            LTexto(2).Caption = "Derecha:"
            LTexto(3).Caption = "Abajo:"
            LTexto(4).Caption = "Izquierda:"
            LTexto(5).Caption = ""
            LTexto(6).Caption = ""
            LTexto(7).Caption = ""
            LTexto(8).Caption = ""
            LTexto(9).Caption = ""
            LUlitError.Caption = ""

    

    Case e_EstadoIndexador.FX
            Call RenuevaListaFX
            For i = 0 To 2
                TextDatos(i).Visible = True
                TextDatos(i).BackColor = vbWhite
                TextDatos(i).Enabled = True
            Next i
            LTexto(1).Visible = False
            TextDatos(1).Visible = False
            For i = 3 To 9
                TextDatos(i).Visible = False
                TextDatos(i).BackColor = vbWhite
            Next i
            LNumActual.Caption = "Fx: "
            LTexto(0).Caption = "NumGrh:"
            LTexto(1).Caption = ""
            LTexto(2).Caption = "Offset:"
            LTexto(3).Caption = ""
            LTexto(4).Caption = ""
            LTexto(5).Caption = ""
            LTexto(6).Caption = ""
            LTexto(7).Caption = ""
            LTexto(8).Caption = ""
            LTexto(9).Caption = ""
            LUlitError.Caption = ""
    Case e_EstadoIndexador.Resource
            Call RenuevaListaResource
            For i = 0 To 5
                TextDatos(i).Visible = True
                TextDatos(i).BackColor = vbWhite
                TextDatos(i).Enabled = True
            Next i
            LTexto(1).Visible = False
            TextDatos(1).Visible = False
            For i = 6 To 9
                TextDatos(i).Visible = False
                TextDatos(i).BackColor = vbWhite
            Next i
            DibujarIndexaciones.activo = False
            LNumActual.Caption = "Crypt: "
            LTexto(0).Caption = "Tamaño:"
            LTexto(1).Caption = ""
            LTexto(2).Caption = "Alto:"
            LTexto(3).Caption = "Ancho:"
            LTexto(4).Caption = "Bits:"
            LTexto(5).Caption = "Situacion:"
            LTexto(6).Caption = ""
            LTexto(7).Caption = ""
            LTexto(8).Caption = ""
            LTexto(9).Caption = ""
            LUlitError.Caption = ""
            Command10.Visible = False
            BotonBorrrar.Visible = False
            DescripcionAyuda.Visible = True
            DescripcionAyuda.Caption = "Click derecho en la lista de graficos para indexar"
            
        Case e_EstadoIndexador.Superficies
            Call RenuevaListaSuperficie
            For i = 0 To 3
                TextDatos(i).Visible = True
                TextDatos(i).BackColor = vbWhite
                TextDatos(i).Enabled = True
            Next i
            'LTexto(1).Visible = False
            'TextDatos(1).Visible = False
            For i = 4 To 9
                TextDatos(i).Visible = False
                TextDatos(i).BackColor = vbWhite
            Next i
            
            LNumActual.Caption = "Sup: "
            LTexto(0).Caption = "Alto:"
            LTexto(1).Caption = "Nombre:"
            LTexto(1).Visible = True
            LTexto(2).Caption = "Ancho:"
            LTexto(3).Caption = "GrhInicio:"
            LTexto(4).Caption = "Capa:"
            LTexto(5).Caption = ""
            LTexto(6).Caption = ""
            LTexto(7).Caption = ""
            LTexto(8).Caption = ""
            LTexto(9).Caption = ""
            LUlitError.Caption = ""
            Command10.Caption = "Guardar indices.ini"
            
    End Select
    Call CambiarcaptionCommand10
End Sub
Private Sub MoverGrh(ByVal numGRH As Integer, ByVal OrigenGRH As Integer, ByVal BorrarOriginal As Boolean)
Dim TempLong As Long
Dim cadena As String
Dim respuesta As Byte
Dim GrhVacio As Grhdata
Dim looPero As Long

TempLong = ListaindexGrH(OrigenGRH)
If TempLong <= 0 Then
    LUlitError.Caption = "grafico incorrecto"
    Exit Sub
End If
TempLong = ListaindexGrH(numGRH)
If TempLong > 0 Then
    respuesta = MsgBox("El grafico " & numGRH & " ya existe, ¿deseas sobreescribirlo?", 4, "Aviso")
    If respuesta = vbYes Then
        Grhdata(numGRH) = Grhdata(OrigenGRH)
        If BorrarOriginal Then
            Grhdata(OrigenGRH) = GrhVacio
        End If
        GRHActual = Val(numGRH)
        LOOPActual = 1
        'frmMain.Visor.Cls
        'Call DibujarGHRVisor(GRHActual)
        'Call GetInfoGHR(GRHActual)
        LGHRnumeroA.Caption = GRHActual
        TempLong = ListaindexGrH(GRHActual)
        frmMain.Lista.listIndex = TempLong
         EstadoNoGuardado(e_EstadoIndexador.Grh) = True
    End If
Else
    Grhdata(numGRH) = Grhdata(OrigenGRH)
    If BorrarOriginal Then
        Grhdata(OrigenGRH) = GrhVacio
    End If
    GRHActual = numGRH
    LOOPActual = 1
    'frmMain.Visor.Cls
    'Call DibujarGHRVisor(GRHActual)
    'Call GetInfoGHR(GRHActual)
    LGHRnumeroA.Caption = GRHActual
    TempLong = ListaindexGrH(GRHActual)
    frmMain.Lista.listIndex = TempLong
     EstadoNoGuardado(e_EstadoIndexador.Grh) = True
End If
    
End Sub

Private Sub SBotonMover(ByVal BorrarOriginal As Boolean, Optional ByVal CantidadM As Integer = 1)
Dim TempLong As Long
Dim cadena As String
Dim respuesta As Byte
Dim GrhVacio As Grhdata
Dim LooPer As Long

Select Case EstadoIndexador
    Case e_EstadoIndexador.Grh
            If GrHCambiando Then
                GrHCambiando = False
                LNumActual.Caption = "Ghr:"
                BotonGuardar.Visible = False
            End If
        cadena = InputBox("Introduzca número de GHR al que quieres mover el grafico " & GRHActual, "Mover Grafico")
        If IsNumeric(cadena) Then
            If Val(cadena) > 0 And Val(cadena) < MAXGrH Then
                Call MoverGrh(Val(cadena), GRHActual, BorrarOriginal)
                Call RenuevaListaGrH
                TempLong = ListaindexGrH(Val(cadena))
                frmMain.Lista.listIndex = TempLong
            Else
                LUlitError.Caption = "introduzca un numero correcto"
            End If
        Else
            LUlitError.Caption = "introduzca un numero"
        End If
    Case Else
        Dim StringCaso As String
        If EstadoIndexador = e_EstadoIndexador.Body Then
            StringCaso = "Body"
        ElseIf EstadoIndexador = e_EstadoIndexador.Cabezas Then
            StringCaso = "Cabeza"
        ElseIf EstadoIndexador = e_EstadoIndexador.Cascos Then
            StringCaso = "Casco"
        ElseIf EstadoIndexador = e_EstadoIndexador.Escudos Then
            StringCaso = "Escudo"
        ElseIf EstadoIndexador = e_EstadoIndexador.Armas Then
            StringCaso = "Arma"
        ElseIf EstadoIndexador = e_EstadoIndexador.FX Then
            StringCaso = "Fx"
        ElseIf EstadoIndexador = e_EstadoIndexador.Resource Then
            Exit Sub
        End If
        cadena = InputBox("Introduzca numero de " & StringCaso & " al que quieres mover", "Mover " & StringCaso)
        If IsNumeric(cadena) And (Val(cadena) < 31000) Then
            If EstadoIndexador = e_EstadoIndexador.Body Then
                Call mueveBody(Val(cadena), DataIndexActual, BorrarOriginal)
            ElseIf EstadoIndexador = e_EstadoIndexador.Cabezas Then
                Call MueveCabeza(Val(cadena), DataIndexActual, BorrarOriginal)
            ElseIf EstadoIndexador = e_EstadoIndexador.Cascos Then
                Call MueveCasco(Val(cadena), DataIndexActual, BorrarOriginal)
            ElseIf EstadoIndexador = e_EstadoIndexador.Escudos Then
                Call MueveEscudo(Val(cadena), DataIndexActual, BorrarOriginal)
            ElseIf EstadoIndexador = e_EstadoIndexador.Armas Then
                Call MueveArma(Val(cadena), DataIndexActual, BorrarOriginal)
            ElseIf EstadoIndexador = e_EstadoIndexador.FX Then
                Call MueveFX(Val(cadena), DataIndexActual, BorrarOriginal)
            End If
                DataIndexActual = Val(cadena)
                LOOPActual = 1
                frmMain.Visor.Cls
                Call GetInfoDataIndex(DataIndexActual)
                Dibujado.Interval = 100
                Dibujado.Enabled = True
                LGHRnumeroA.Caption = DataIndexActual
                TempLong = ListaindexGrH(DataIndexActual)
                frmMain.Lista.listIndex = TempLong
                 EstadoNoGuardado(EstadoIndexador) = True
        Else
            LUlitError.Caption = "introduzca un numero valido"
        End If
End Select
End Sub



Private Sub BotonI_Click(Index As Integer)
If EstadoIndexador <> Index Then Call CambiarEstado(Index)
Call ComprobarIndexLista

End Sub

Private Sub CDibujarWalk_Click()
    DibujarWalk = CDibujarWalk.listIndex
    Visor.Cls
End Sub

Private Sub Checkcabeza_Click()

If Checkcabeza.Value = vbChecked Then
    cabezaActual = 3008
Else
    cabezaActual = 0
End If
End Sub

Private Sub CmbResourceFile_Click()
'Borramos todas las surfaces de la memoria. Sirve por si se hacen cambios en los BMPs y se necesita obligar a recargarlos

    
    If Not IniciadoTodo Then
        Call SurfaceDB.BorrarTodo
    Else
        IniciadoTodo = False
    End If
    
    Call CambiarEstado(EstadoIndexador)
    
    
    Call ComprobarIndexLista
End Sub


Private Sub SBotonBorrrar()
Dim respuesta As Byte
Dim TempLong As Long

TempLong = frmMain.Lista.listIndex

Select Case EstadoIndexador
    Case e_EstadoIndexador.Superficies
        TempLong = Val(ReadField(1, Lista.List(Lista.listIndex), Asc(" ")))
        respuesta = MsgBox("ATENCION ¿Estas segudo de borrar la superficie N° " & TempLong & " - " & SupData(TempLong).Nombre & " ?", 4, "¡¡ADVERTENCIA!!")
        
        If respuesta = vbYes Then
            SupData(TempLong).GrhIndex = 0
            'Call RenuevaListaGrH
            frmMain.Lista.RemoveItem Lista.listIndex
        End If
        'EstadoNoGuardado(e_EstadoIndexador.Superficies) = True
        
    Case e_EstadoIndexador.Grh
        If GRHActual = 0 Then Exit Sub
        respuesta = MsgBox("ATENCION ¿Estas segudo de borrar el Grh " & GRHActual & " ?", 4, "¡¡ADVERTENCIA!!")
        If respuesta = vbYes Then
            Grhdata(GRHActual).NumFrames = 0
            'Call RenuevaListaGrH
            frmMain.Lista.RemoveItem TempLong
        End If
    Case e_EstadoIndexador.Body
        If DataIndexActual = 0 Then Exit Sub
        respuesta = MsgBox("ATENCION ¿Estas segudo de borrar el body " & DataIndexActual & " ?", 4, "¡¡ADVERTENCIA!!")
        If respuesta = vbYes Then
            BodyData(DataIndexActual).Walk(1).GrhIndex = 0
            BodyData(DataIndexActual).Walk(2).GrhIndex = 0
            BodyData(DataIndexActual).Walk(3).GrhIndex = 0
            BodyData(DataIndexActual).Walk(4).GrhIndex = 0
            'Call RenuevaListaBodys
            frmMain.Lista.RemoveItem TempLong
        End If
    Case e_EstadoIndexador.Cabezas
        If DataIndexActual = 0 Then Exit Sub
        respuesta = MsgBox("ATENCION ¿Estas segudo de borrar la Cabeza " & DataIndexActual & " ?", 4, "¡¡ADVERTENCIA!!")
        If respuesta = vbYes Then
            HeadData(DataIndexActual).Head(1).GrhIndex = 0
            HeadData(DataIndexActual).Head(2).GrhIndex = 0
            HeadData(DataIndexActual).Head(3).GrhIndex = 0
            HeadData(DataIndexActual).Head(4).GrhIndex = 0
            'Call RenuevaListaCabezas
            frmMain.Lista.RemoveItem TempLong
        End If
    Case e_EstadoIndexador.Cascos
        If DataIndexActual = 0 Then Exit Sub
        respuesta = MsgBox("ATENCION ¿Estas segudo de borrar el casco " & DataIndexActual & " ?", 4, "¡¡ADVERTENCIA!!")
        If respuesta = vbYes Then
            CascoAnimData(DataIndexActual).Head(1).GrhIndex = 0
            CascoAnimData(DataIndexActual).Head(2).GrhIndex = 0
            CascoAnimData(DataIndexActual).Head(3).GrhIndex = 0
            CascoAnimData(DataIndexActual).Head(4).GrhIndex = 0
            'Call RenuevaListaCascos
            frmMain.Lista.RemoveItem TempLong
        End If
    Case e_EstadoIndexador.Escudos
        If DataIndexActual = 0 Then Exit Sub
        respuesta = MsgBox("ATENCION ¿Estas segudo de borrar el escudo " & DataIndexActual & " ?", 4, "¡¡ADVERTENCIA!!")
        If respuesta = vbYes Then
            ShieldAnimData(DataIndexActual).ShieldWalk(1).GrhIndex = 0
            ShieldAnimData(DataIndexActual).ShieldWalk(2).GrhIndex = 0
            ShieldAnimData(DataIndexActual).ShieldWalk(3).GrhIndex = 0
            ShieldAnimData(DataIndexActual).ShieldWalk(4).GrhIndex = 0
            frmMain.Lista.RemoveItem TempLong
            'Call RenuevaListaEscudos
        End If
    Case e_EstadoIndexador.Armas
        If DataIndexActual = 0 Then Exit Sub
        respuesta = MsgBox("ATENCION ¿Estas segudo de borrar el arma " & DataIndexActual & " ?", 4, "¡¡ADVERTENCIA!!")
        If respuesta = vbYes Then
            WeaponAnimData(DataIndexActual).WeaponWalk(1).GrhIndex = 0
            WeaponAnimData(DataIndexActual).WeaponWalk(2).GrhIndex = 0
            WeaponAnimData(DataIndexActual).WeaponWalk(3).GrhIndex = 0
            WeaponAnimData(DataIndexActual).WeaponWalk(4).GrhIndex = 0
            'Call RenuevaListaArmas
            frmMain.Lista.RemoveItem TempLong
        End If

    Case e_EstadoIndexador.FX
        If DataIndexActual = 0 Then Exit Sub
        respuesta = MsgBox("ATENCION ¿Estas segudo de borrar el FX " & DataIndexActual & " ?", 4, "¡¡ADVERTENCIA!!")
        If respuesta = vbYes Then
            FxData(DataIndexActual).FX.GrhIndex = 0
            FxData(DataIndexActual).offsetx = 0
            FxData(DataIndexActual).offsety = 0
            'Call RenuevaListaFX
            frmMain.Lista.RemoveItem TempLong
        End If
End Select

If TempLong < frmMain.Lista.ListCount Then
    frmMain.Lista.listIndex = TempLong
Else
    frmMain.Lista.listIndex = frmMain.Lista.ListCount - 1
End If
End Sub
Public Function StringGuardadoActual(PEstado As e_EstadoIndexador) As String
Dim elq As String
Select Case PEstado
    Case e_EstadoIndexador.Grh
        If SavePath = 0 Then
            elq = "Graficos"
        Else
            elq = "Graficos" & SavePath
        End If
    Case e_EstadoIndexador.Body
        If SavePath = 0 Then
            elq = "personajes"
        Else
            elq = "personajes" & SavePath
        End If
    Case e_EstadoIndexador.Cabezas
        If SavePath = 0 Then
            elq = "cabezas"
        Else
            elq = "cabezas" & SavePath
        End If
    Case e_EstadoIndexador.Cascos
        If SavePath = 0 Then
            elq = "cascos"
        Else
            elq = "cascos" & SavePath
        End If
    Case e_EstadoIndexador.Escudos
        If SavePath = 0 Then
            elq = "escudos"
        Else
            elq = "escudos" & SavePath
        End If
    Case e_EstadoIndexador.Armas
        If SavePath = 0 Then
            elq = "armas"
        Else
            elq = "armas" & SavePath
        End If

    Case e_EstadoIndexador.FX
        If SavePath = 0 Then
            elq = "fxs"
        Else
            elq = "fxs" & SavePath
        End If
    Case e_EstadoIndexador.Resource
        elq = ""
        
    Case e_EstadoIndexador.Superficies
        elq = "Superficies"
End Select
StringGuardadoActual = elq
End Function
Private Sub CambiarcaptionCommand10()
Command10.Caption = "Guardar " & StringGuardadoActual(EstadoIndexador)
MenuArchivoGuardar.Caption = "Guardar " & StringGuardadoActual(EstadoIndexador)
MenuArchivoCargar.Caption = "Cargar " & StringGuardadoActual(EstadoIndexador)

If EstadoIndexador = e_EstadoIndexador.Escudos Or EstadoIndexador = e_EstadoIndexador.Armas Then
    Command10.Caption = Command10.Caption & ".dat"
    MenuArchivoGuardar.Caption = MenuArchivoGuardar.Caption & ".dat"
    MenuArchivoCargar.Caption = MenuArchivoCargar.Caption & ".dat"
Else
    Command10.Caption = Command10.Caption & ".ind"
    MenuArchivoGuardar.Caption = MenuArchivoGuardar.Caption & ".ind"
    MenuArchivoCargar.Caption = MenuArchivoCargar.Caption & ".ind"
End If
End Sub

Private Sub Command1_Click()
DibujarFondo = Not DibujarFondo
Call ClickEnLista
End Sub

Private Sub Command10_Click()
'Boton de guardado en disco
Call BotonGuardado
End Sub


Private Sub Command4_Click()
On Error GoTo errhandler
Call CargarTips

Dim N As Integer, i As Integer
N = FreeFile

Open ConfigDir.Inits & "\Tips.ayu" For Binary As #N
'Escribimos la cabecera
Put #N, , MiCabecera
'Guardamos las cabezas
Put #N, , NumTips

For i = 1 To NumTips
    Put #N, , Tips(i)
Next i

Close #N
Call MsgBox("Listo, encode ok!!")

Exit Sub
errhandler:
Call MsgBox("Error en tip " & i)

End Sub







Private Sub BotonGuardar_Click()
'boton de guardado en memoria
    Dim i As Long
Select Case EstadoIndexador
    Case e_EstadoIndexador.Grh
        'guardando un grafico
        
        If GRHActual = 0 Then Exit Sub
    
        If Val(TextDatos(2).Text) <= 0 Then ' numframes = 0
            MsgBox "numero de frames incorrecto"
            Exit Sub
        End If
    
        If Val(TextDatos(2).Text) = 1 Then ' si no es animacion se comprueba si existe el BMP
            If ExisteBMP(Val(TextDatos(0).Text)) = ResourceFile Or (ResourceFile And ExisteBMP(Val(TextDatos(0).Text)) > 0) Then
            Else
                LUlitError.Caption = "No existe el archivo del grafico"
                Exit Sub
            End If
        End If
        
        Grhdata(GRHActual).FileNum = Val(TextDatos(0).Text)
        Grhdata(GRHActual).NumFrames = Val(TextDatos(2).Text)
        If Grhdata(GRHActual).NumFrames = 1 Then
            Grhdata(GRHActual).Frames(1) = GRHActual
            Grhdata(GRHActual).pixelHeight = Val(TextDatos(3).Text)
            Grhdata(GRHActual).pixelWidth = Val(TextDatos(4).Text)
            Grhdata(GRHActual).Speed = Val(TextDatos(5).Text)
            Grhdata(GRHActual).sX = Val(TextDatos(6).Text)
            Grhdata(GRHActual).sY = Val(TextDatos(7).Text)
        
            Grhdata(GRHActual).TileHeight = Grhdata(GRHActual).pixelHeight / TilePixelHeight
            Grhdata(GRHActual).TileWidth = Grhdata(GRHActual).pixelWidth / TilePixelWidth
        Else
            For i = 1 To Grhdata(GRHActual).NumFrames
                If Val(ReadField(i, TextDatos(1).Text, Asc("-"))) < 32000 Then
                    Grhdata(GRHActual).Frames(i) = Val(ReadField(i, TextDatos(1).Text, Asc("-")))
                End If
            Next i
            Grhdata(GRHActual).Speed = Val(TextDatos(5).Text)
            If Grhdata(GRHActual).Frames(1) > 0 Then
                Grhdata(GRHActual).pixelHeight = Grhdata(Grhdata(GRHActual).Frames(1)).pixelHeight
                Grhdata(GRHActual).pixelWidth = Grhdata(Grhdata(GRHActual).Frames(1)).pixelWidth
                Grhdata(GRHActual).sX = Grhdata(Grhdata(GRHActual).Frames(1)).sX
                Grhdata(GRHActual).sY = Grhdata(Grhdata(GRHActual).Frames(1)).sY
                Grhdata(GRHActual).TileHeight = Grhdata(Grhdata(GRHActual).Frames(1)).TileHeight
                Grhdata(GRHActual).TileWidth = Grhdata(Grhdata(GRHActual).Frames(1)).TileWidth
            Else
                Grhdata(GRHActual).pixelHeight = Val(TextDatos(3).Text)
                Grhdata(GRHActual).pixelWidth = Val(TextDatos(4).Text)
                Grhdata(GRHActual).sX = Val(TextDatos(6).Text)
                Grhdata(GRHActual).sY = Val(TextDatos(7).Text)
                Grhdata(GRHActual).TileHeight = Grhdata(GRHActual).pixelHeight / TilePixelHeight
                Grhdata(GRHActual).TileWidth = Grhdata(GRHActual).pixelWidth / TilePixelWidth
            End If

        End If
        
        Call GetInfoGHR(GRHActual)
        frmMain.Visor.Cls
        Call DibujarGHRVisor(GRHActual)
     Case e_EstadoIndexador.Body
        BodyData(DataIndexActual).HeadOffset.y = Val(ReadField(1, TextDatos(5).Text, Asc("º")))
        BodyData(DataIndexActual).HeadOffset.x = Val(ReadField(2, TextDatos(5).Text, Asc("º")))
        BodyData(DataIndexActual).Walk(1).GrhIndex = Val(TextDatos(0).Text)
        BodyData(DataIndexActual).Walk(2).GrhIndex = Val(TextDatos(2).Text)
        BodyData(DataIndexActual).Walk(3).GrhIndex = Val(TextDatos(3).Text)
        BodyData(DataIndexActual).Walk(4).GrhIndex = Val(TextDatos(4).Text)
     Case e_EstadoIndexador.Cabezas
        HeadData(DataIndexActual).Head(1).GrhIndex = Val(TextDatos(0).Text)
        HeadData(DataIndexActual).Head(2).GrhIndex = Val(TextDatos(2).Text)
        HeadData(DataIndexActual).Head(3).GrhIndex = Val(TextDatos(3).Text)
        HeadData(DataIndexActual).Head(4).GrhIndex = Val(TextDatos(4).Text)
     Case e_EstadoIndexador.Cascos
        CascoAnimData(DataIndexActual).Head(1).GrhIndex = Val(TextDatos(0).Text)
        CascoAnimData(DataIndexActual).Head(2).GrhIndex = Val(TextDatos(2).Text)
        CascoAnimData(DataIndexActual).Head(3).GrhIndex = Val(TextDatos(3).Text)
        CascoAnimData(DataIndexActual).Head(4).GrhIndex = Val(TextDatos(4).Text)
    Case e_EstadoIndexador.Armas
        WeaponAnimData(DataIndexActual).WeaponWalk(1).GrhIndex = Val(TextDatos(0).Text)
        WeaponAnimData(DataIndexActual).WeaponWalk(2).GrhIndex = Val(TextDatos(2).Text)
        WeaponAnimData(DataIndexActual).WeaponWalk(3).GrhIndex = Val(TextDatos(3).Text)
        WeaponAnimData(DataIndexActual).WeaponWalk(4).GrhIndex = Val(TextDatos(4).Text)
     Case e_EstadoIndexador.Escudos
        ShieldAnimData(DataIndexActual).ShieldWalk(1).GrhIndex = Val(TextDatos(0).Text)
        ShieldAnimData(DataIndexActual).ShieldWalk(2).GrhIndex = Val(TextDatos(2).Text)
        ShieldAnimData(DataIndexActual).ShieldWalk(3).GrhIndex = Val(TextDatos(3).Text)
        ShieldAnimData(DataIndexActual).ShieldWalk(4).GrhIndex = Val(TextDatos(4).Text)
    Case e_EstadoIndexador.FX
        FxData(DataIndexActual).FX.GrhIndex = Val(TextDatos(0).Text)
        FxData(DataIndexActual).offsetx = Val(ReadField(2, TextDatos(2).Text, Asc("º")))
        FxData(DataIndexActual).offsety = Val(ReadField(1, TextDatos(2).Text, Asc("º")))
End Select
If EstadoIndexador <> e_EstadoIndexador.Grh Then
    Call GetInfoDataIndex(DataIndexActual)
End If
EstadoNoGuardado(EstadoIndexador) = True
End Sub



Private Sub Command2_Click()

    frmDats.Show
End Sub

Private Sub Command3_Click()
    frmSelectBMP.Show , Me
End Sub

Private Sub Dibujado_Timer()
On Error Resume Next
If Me.WindowState = vbMinimized Then Exit Sub
Select Case EstadoIndexador
    Case e_EstadoIndexador.Grh
        If Not GrHCambiando Then
             If GRHActual <= 0 Then Exit Sub
             If LOOPActual > Grhdata(GRHActual).NumFrames Then LOOPActual = 1
             Call DibujarGHRVisor(Grhdata(GRHActual).Frames(LOOPActual))
             LOOPActual = LOOPActual + 1
         Else
             If LOOPActual > TempGrh.NumFrames Then LOOPActual = 1
             Call DibujarTempGHRVisor(TempGrh.Frames(LOOPActual))
             LOOPActual = LOOPActual + 1
         End If
    Case e_EstadoIndexador.Resource
    
    Case Else
             If DataIndexActual <= 0 Then Exit Sub
             If tempDataIndex.Walk(1).GrhIndex = 0 Then Exit Sub
             If LOOPActual > Grhdata(tempDataIndex.Walk(1).GrhIndex).NumFrames Then LOOPActual = 1
             Call DibujarDataIndex(tempDataIndex, LOOPActual, DibujarWalk)
             LOOPActual = LOOPActual + 1
End Select
End Sub
Private Sub Form_close()
    End
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Dim i As Long
Dim HayCambios As Boolean
Dim respuesta As Byte
Dim tStr As String

Inicio:
HayCambios = False
For i = e_EstadoIndexador.Grh To Resource
    If EstadoNoGuardado(i) Then
        HayCambios = True
        tStr = tStr & StringGuardadoActual(i) & vbCrLf
    End If
Next i
For i = 1 To 3
    If DatNoGuardado(i) = True Then
        HayCambios = True
        Select Case i
            Case eModo.Objetos
                tStr = tStr & "Objetos dats" & vbCrLf
            Case eModo.Npc
                tStr = tStr & "NPCs dats" & vbCrLf
            Case eModo.Hechizo
                tStr = tStr & "Hechizos dats" & vbCrLf
        End Select
    End If
Next i

    If HayCambios Then
        respuesta = MsgBox("Hay cambios sin Guardar en :" & vbCrLf & vbCrLf & tStr & vbCrLf & "¿Quieres GUARDAR los cambios antes de salir?" & vbCrLf & "(Si pulsas NO se perderan estos cambios)" & vbCrLf, 3, "Aviso")
        If respuesta = vbCancel Then
            Cancel = 1 ' cancelamos la salida
            Exit Sub
        ElseIf respuesta = vbYes Then
            For i = e_EstadoIndexador.Grh To e_EstadoIndexador.Superficies ' guardamos
                If EstadoNoGuardado(i) Then
                    EstadoIndexador = i
                    Call BotonGuardado
                End If
            Next i
            For i = 1 To 3
                If DatNoGuardado(i) = True Then
                    estadoDat = i
                    frmDats.GuardarEstadoActual
                End If
            Next i
            tStr = vbNullString
            GoTo Inicio ' weno el goto es el alien d la programacion estructurada pero paso d romperme la cabeza xD asi se ve mejor
            ' volvemos a comprobar si algo no se guardo
        End If
        
    End If

End Sub

Private Sub Form_resize()
    Visor.Height = Abs(frmMain.Height - Visor.Top - LUlitError.Height - 950)
    Visor.Width = Abs(frmMain.Width - Visor.Left - 120)
    LUlitError.Top = Abs(frmMain.Height - LUlitError.Height - 700)
    LUlitError.Width = Abs(frmMain.Width - 155)
    Call ClickEnLista
End Sub

Private Sub Form_Load()

    'configuracion inicial:
    SavePath = 0
    LoadingNew = False ' variable que evita redibujado excesibo
    IniciadoTodo = True
    ColorFondo = vbGreen
    
    modProgressBar.Init 0
    ResourceFile = 1 ' siempre cargamos lo bmps, esta deshabilitado el archivo de recursos.
    
    If ResourceFile <= 0 Then ResourceFile = 1
 
    Call IniciarCabecera(MiCabecera)
    
    Call IniciarObjetosDirectX
    Set SurfaceDB = New clsSurfaceManDyn
    Call InitTileEngine(frmMain.hWnd, 155, 16, 32, 32, 13, 17, 9)
      
    'Prepare surfaces for text rendering
    BackBufferSurface.SetFontTransparency True
'TODO : Fonts should be in a separate class / collection
    Dim font As New StdFont
    Dim Ifnt As IFont
    
    font.name = "Verdana"
    font.Bold = True
    font.Italic = False
    font.Size = 6
    font.Underline = False
    font.Strikethrough = False
    
    
    Set Ifnt = font
    
    BackBufferSurface.SetFont Ifnt    'If cargado = True Then Exit Sub
    
    Call CargarAnimsExtra
    Call CargarTips
    
    Call CargarArrayLluvia
    Call CargarAnimArmas
    Call CargarAnimEscudos

    EstadoIndexador = e_EstadoIndexador.Grh
    'Dim Lister As Long
    'For Lister = 0 To 7
    '    MenuIndiceGuardado(Lister).Checked = False
   ' Next Lister
    'MenuIndiceGuardado(0).Checked = True
    Call CambiarcaptionCommand10
    If reConfigurarPath = True Then
        frmConfig.Show vbModal
        Unload Me
    End If

End Sub



Private Sub CargarMapas()
Dim loopC As Integer

NumMapas = Val(GetVar(App.Path & "\encode\mapas.dat", "INIT", "NumMaps"))

ReDim Mapas(0 To NumMapas + 1) As Byte

For loopC = 1 To NumMapas
    Mapas(loopC) = Val(GetVar(App.Path & "\encode\mapas.dat", "Map" & loopC, "Lluvia"))
Next loopC

End Sub




Private Sub CargarTips()
Dim loopC As Integer
NumTips = Val(GetVar(App.Path & "\encode\tips.dat", "INIT", "Tips"))

ReDim Tips(0 To NumTips + 1) As String * 255

For loopC = 1 To NumTips
    Tips(loopC) = GetVar(App.Path & "\encode\tips.dat", "Tip" & loopC, "Tip")
Next loopC

End Sub




Private Sub Form_Unload(Cancel As Integer)
    End
End Sub

Private Sub IAnim_Click()
Dim textoActual As String
Dim bmpstring As String
Dim BMPBuscado As Long


bmpstring = ReadField(1, Lista.List(Lista.listIndex), Asc(" "))

If IsNumeric(bmpstring) Then
    BMPBuscado = Val(bmpstring)
    If BMPBuscado > 0 And BMPBuscado <= 32000 Then
        If FormAuto.Visible Then
            FormAuto.SetFocus
        Else
            FormAuto.Show , frmMain
            
        End If
        
        FormAuto.FrameAnim(0).Visible = True
        FormAuto.FrameAnim(1).Visible = False
        FormAuto.FrameAnim(2).Visible = False
        FormAuto.loadTabStrip
        FormAuto.TextDatos(4).Text = BMPBuscado
    End If
End If
End Sub


Private Sub lblEstado_Click()

End Sub

Private Sub Lista_Click()
    Call ClickEnLista
End Sub
Public Sub SuperficieClick(ByVal Index As Integer)

    With SupData(Index)
            TextDatos(1).Text = .Nombre
            TextDatos(0).Text = .Ancho 'ancho
            TextDatos(2).Text = .Alto  'alto
            TextDatos(3).Text = .GrhIndex 'gthindex
            TextDatos(4).Text = .Capa
            If .Alto = 0 And .Ancho = 0 Then
                Call DibujarGHRVisor(.GrhIndex)
                
            Else
                Dim r As RECT
    
                If DibujarFondo Then
                    BackBufferSurface.BltColorFill r, ColorFondo
                Else
                    BackBufferSurface.BltColorFill r, 0
                End If
                Dim x As Long, y As Long, nINdex As Long
                For y = 1 To .Alto
                    For x = 1 To .Ancho
                        nINdex = .GrhIndex + (x - 1) + ((y - 1) * (.Ancho))
                        'Debug.Print (x - 1) + ((Y - 1) * (.Ancho))
                        Call DDrawTransGrhIndextoSurface(BackBufferSurface, nINdex, (x - 1) * Grhdata(nINdex).pixelWidth, (y - 1) * Grhdata(nINdex).pixelWidth, 0, 0)
                    Next x
                Next y
                Dim auxr As RECT
                auxr.Left = 0
                auxr.Top = 0
                auxr.Bottom = .Alto * Grhdata(.GrhIndex).pixelHeight
                auxr.Right = .Ancho * Grhdata(.GrhIndex).pixelWidth
                r = auxr
                
                
                Dim FramesY As Long, FramesX As Long, curx As Byte, cury As Byte, actualFrame As Byte
                Dim FramesTotales As Integer
                Dim FramesAncho As Integer, FramesAlto As Integer, Alto As Long, Ancho As Long
                FramesTotales = .Ancho * .Alto
                Alto = Grhdata(.GrhIndex).pixelHeight
                Ancho = Grhdata(.GrhIndex).pixelWidth
                FramesAncho = .Ancho
                FramesAlto = .Alto
                DibujarIndexaciones.Ancho = Ancho
                DibujarIndexaciones.Alto = Alto
                DibujarIndexaciones.activo = True
                For FramesY = 1 To FramesAlto
                    For FramesX = 1 To FramesAncho
                        DibujarIndexaciones.Inicios(actualFrame + 1).x = Ancho * curx
                        DibujarIndexaciones.Inicios(actualFrame + 1).y = Alto * cury
                        curx = curx + 1
                        actualFrame = actualFrame + 1
                    Next FramesX
                    curx = 0
                    cury = cury + 1
                Next FramesY
                DibujarIndexaciones.Total = FramesTotales
                BackBufferSurface.SetForeColor vbGreen
                BackBufferSurface.setDrawStyle DrawStyleConstants.vbDot
                Dim ii As Long
                For ii = 1 To DibujarIndexaciones.Total
                    BackBufferSurface.DrawBox DibujarIndexaciones.Inicios(ii).x, DibujarIndexaciones.Inicios(ii).y, DibujarIndexaciones.Inicios(ii).x + DibujarIndexaciones.Ancho, DibujarIndexaciones.Inicios(ii).y + DibujarIndexaciones.Alto
                Next ii
                
                BackBufferSurface.BltToDC frmMain.Visor.hdc, r, auxr
                frmMain.Visor.Refresh
            End If
            
        End With
End Sub
Public Sub ClickEnLista()

On Error Resume Next

Select Case EstadoIndexador
    Case e_EstadoIndexador.Grh
        GRHActual = Val(ReadField(1, Lista.List(Lista.listIndex), Asc(" ")))
        LOOPActual = 1
        Call GetInfoGHR(GRHActual)
        Call DibujarGHRVisor(GRHActual)
        LGHRnumeroA.Caption = GRHActual
    Case e_EstadoIndexador.Resource
        GRHActual = Val(Lista.List(Lista.listIndex))
        frmMain.Visor.Cls
        If ExisteBMP(GRHActual) = 0 Then Exit Sub
        Call GetInfoBmp(GRHActual)
        Call DibujarBMPVisor(GRHActual)
        LGHRnumeroA.Caption = GRHActual
        If DibujarIndexaciones.activo = True Then
            FormAuto.TextDatos(4).Text = Val(GRHActual)
            FormAuto.TextDatos2(4).Text = Val(GRHActual)
           FormAuto.TextDatos3(4).Text = Val(GRHActual)
       End If
    Case e_EstadoIndexador.Superficies
        GRHActual = Val(ReadField(1, Lista.List(Lista.listIndex), Asc(" ")))
        frmMain.Visor.Cls
        Call SuperficieClick(GRHActual)
    
    Case Else
        frmMain.Visor.Cls
        DataIndexActual = Val(Lista.List(Lista.listIndex))
        LOOPActual = 1
        Call GetInfoDataIndex(DataIndexActual)
        
        LGHRnumeroA.Caption = DataIndexActual
End Select
UltimoindexE(EstadoIndexador) = Lista.listIndex

End Sub


Private Sub Lista_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = vbRightButton And EstadoIndexador = Resource Then
        Call Me.PopupMenu(Me.mnuautoI)
    End If
    If Button = vbRightButton And EstadoIndexador <> Resource Then
        Select Case EstadoIndexador
            Case e_EstadoIndexador.Body
                Call Me.PopupMenu(Me.mnuDatBodys)
                            
            Case e_EstadoIndexador.Grh
                
                Me.PopupMenu Me.mnuDat
                
            Case e_EstadoIndexador.Armas
                
                Me.PopupMenu Me.mnuEspada
                
            Case e_EstadoIndexador.Cascos

                Me.PopupMenu Me.mnuCasco
                
            Case e_EstadoIndexador.Escudos
                
                Me.PopupMenu Me.mnuEscudo
                
            Case e_EstadoIndexador.FX

                Me.PopupMenu Me.mnuHechizo
                
        End Select
    End If
End Sub

Private Sub MenuAcercaDe_Click()
MsgBox "IndexDater v1.0 by El_Santo43" & vbCrLf & vbCrLf & "Email: ladeuixsanti@gmail.com" & vbCrLf & vbCrLf & "Version: " & VERSION_ACTUAL
End Sub

Private Sub MenuArchivoCargar_Click()
    Call BotonCargado
End Sub

Private Sub MenuArchivoGuardar_Click()
    Call BotonGuardado
End Sub

Private Sub MenuBotonCargarP_Click()
On Error GoTo Cancelar
With CommonDialog1
    .Filter = "Binario(.ind)|*.ind|Archivo de texto DAT(*.dat)|*.dat"
    If EstadoIndexador = e_EstadoIndexador.Armas Or EstadoIndexador = e_EstadoIndexador.Escudos Then
        .Filter = "Archivo de texto DAT(*.dat)|*.dat"
    End If
    .CancelError = True
    .flags = cdlOFNFileMustExist
    .FileName = StringGuardadoActual(EstadoIndexador)
    .ShowOpen
End With
 
 
Select Case UCase(Right(CommonDialog1.FileName, 3))
    Case "DAT"
        Call BotonCargadoDat(CommonDialog1.FileName)
    Case "IND"
        Call BotonCargado(CommonDialog1.FileName)
    Case Else
       Exit Sub
End Select
Exit Sub
Cancelar:
End Sub

Private Sub MenuBotonGuardarP_Click()
On Error GoTo Cancelar
With CommonDialog1
    .Filter = "Binario(.ind)|*.ind|Archivo de texto DAT(*.dat)|*.dat"
    If EstadoIndexador = e_EstadoIndexador.Armas Or EstadoIndexador = e_EstadoIndexador.Escudos Then
        .Filter = "Archivo de texto DAT(*.dat)|*.dat"
    End If
    .CancelError = True
    .flags = cdlOFNOverwritePrompt
    .FileName = StringGuardadoActual(EstadoIndexador)
    .ShowSave
End With
 
 
Select Case UCase(Right(CommonDialog1.FileName, 3))
    Case "DAT"
        Call BotonGuardadoDat(CommonDialog1.FileName)
    Case "IND"
        Call BotonGuardado(CommonDialog1.FileName)
    Case Else
       Exit Sub
End Select

Exit Sub
Cancelar:
End Sub

Private Sub MenuEdicionBorrar_Click()
    Call SBotonBorrrar
End Sub

Private Sub MenuEdicionClonar_Click()
    Call SbotonClonar
End Sub

Private Sub menuEdicionColor_Click()
With CommonDialog1
    .DialogTitle = "Seleccionar color para el fondo"
    .ShowColor
End With

ColorFondo = CommonDialog1.Color
Call ClickEnLista
End Sub

Private Sub MenuEdicionCopiar_Click()
    Call SBotonMover(False)
End Sub

Private Sub menuEdicionMover_Click()
    Call SBotonMover(True)
End Sub
Public Sub SbotonClonar()
Dim TempLong As Long
Dim cadena As String
Dim respuesta As Byte
Dim LooPer As Long
Dim Inicial As Long
Dim Final As Long
Dim Origen As Long

Select Case EstadoIndexador
    Case e_EstadoIndexador.Grh
            If GrHCambiando Then
                GrHCambiando = False
                LNumActual.Caption = "Ghr:"
                BotonGuardar.Visible = False
            End If
        If GRHActual < 0 Or GRHActual > MAXGrH Then Exit Sub
        cadena = InputBox("Introduzca el primer Numero de GHR al que quieres mover el grafico " & GRHActual, "Clonar Grafico")
        If IsNumeric(cadena) Then
            Inicial = Val(cadena)
            If Inicial > 0 And Inicial < MAXGrH Then
                cadena = InputBox("Introduzca Cantidad de veces que quieres clonar el grafico " & GRHActual & " a partir de la posicion: " & Inicial, "Clonar Grafico")
                If IsNumeric(cadena) Then
                    Final = Val(cadena) + Inicial
                    If Final > 0 And Final < MAXGrH Then
                        Origen = GRHActual
                        For LooPer = Inicial To Final
                            Call MoverGrh(LooPer, Origen, False)
                        Next LooPer
                        Call RenuevaListaGrH
                        TempLong = ListaindexGrH(Inicial)
                        frmMain.Lista.listIndex = TempLong
                         EstadoNoGuardado(e_EstadoIndexador.Grh) = True
                    Else
                        MsgBox "Fuera de los limites"
                    End If
                End If
            Else
                MsgBox "numero incorrecto"
            End If
        End If
'    Case Else
'        Dim StringCaso As String
'        If EstadoIndexador = Body Then
'            StringCaso = "Body"
'        ElseIf EstadoIndexador = e_EstadoIndexador.Cabezas Then
'            StringCaso = "Cabeza"
'        ElseIf EstadoIndexador = e_EstadoIndexador.Cascos Then
'            StringCaso = "Casco"
'        ElseIf EstadoIndexador = e_EstadoIndexador.Escudos Then
'            StringCaso = "Escudo"
'        ElseIf EstadoIndexador = e_EstadoIndexador.Armas Then
'            StringCaso = "Arma"
'        ElseIf EstadoIndexador = e_EstadoIndexador.Botas Then
'            StringCaso = "Bota"
'        ElseIf EstadoIndexador = e_EstadoIndexador.Capas Then
'            StringCaso = "Capa"
'        ElseIf EstadoIndexador = e_EstadoIndexador.Fx Then
'            StringCaso = "Fx"
'        End If
'        cadena = InputBox("Introduzca numero de " & StringCaso & " al que quieres mover", "Mover " & StringCaso)
'        If IsNumeric(cadena) And (Val(cadena) < 31000) Then
'            If EstadoIndexador = e_EstadoIndexador.Body Then
'                Call mueveBody(Val(cadena), DataIndexActual, False)
'            ElseIf EstadoIndexador = e_EstadoIndexador.Cabezas Then
'                Call MueveCabeza(Val(cadena), DataIndexActual, False)
'            ElseIf EstadoIndexador = e_EstadoIndexador.Cascos Then
'                Call MueveCasco(Val(cadena), DataIndexActual, False)
'            ElseIf EstadoIndexador = e_EstadoIndexador.Escudos Then
'                Call MueveEscudo(Val(cadena), DataIndexActual, False)
'            ElseIf EstadoIndexador = e_EstadoIndexador.Armas Then
'                Call MueveArma(Val(cadena), DataIndexActual, False)
'            ElseIf EstadoIndexador = e_EstadoIndexador.Botas Then
'                Call MueveBota(Val(cadena), DataIndexActual, False)
'            ElseIf EstadoIndexador = e_EstadoIndexador.Capas Then
'                Call MueveCapa(Val(cadena), DataIndexActual, False)
'            ElseIf EstadoIndexador = e_EstadoIndexador.Fx Then
'                Call MueveFX(Val(cadena), DataIndexActual, False)
'            End If
'                DataIndexActual = Val(cadena)
'                LOOPActual = 1
'                frmMain.Visor.Cls
'                Call GetInfoDataIndex(DataIndexActual)
'                Dibujado.Interval = 200
'                Dibujado.Enabled = True
'                LGHRnumeroA.Caption = DataIndexActual
'                templong = ListaindexGrH(DataIndexActual)
'                frmMain.Lista.listIndex = templong
'        Else
'            MsgBox "introduzca un numero valido"
'        End If
End Select
End Sub
Private Sub MenuEdicionNuevo_Click()
Call BotonNuevoGRH
End Sub



Private Sub MenuHerramientasAAnim_Click()
    If FormAuto.Visible Then
        FormAuto.SetFocus
    Else
        FormAuto.Show , frmMain
    End If
    'loadTabStrip
End Sub

Private Sub MenuHerramientasBG_Click()
Dim i As Long
Dim cadena As String

LastFound = 0
cadena = InputBox("Introduzca número de Bmp a buscar", "Nuevo Grafico")
If Val(cadena) > 0 And Val(cadena) <= MAXGrH Then
    BMPBuscado = Val(cadena)
    For i = 1 To MAXGrH
        If Grhdata(i).FileNum = BMPBuscado Then
            Call BuscarNuevoF(i)
            LastFound = i
            MenuHerramientasBN.Visible = True
            LUlitError.Caption = "F3 para continuar la busqueda"
            Exit Sub
        End If
    Next i
    LUlitError.Caption = "BMP no encontrado"
    MenuHerramientasBN.Visible = False
End If
End Sub

Private Sub MenuHerramientasBN_Click()
Dim i As Long

If LastFound = 0 Or BMPBuscado = 0 Then Exit Sub
For i = LastFound + 1 To MAXGrH
    If Grhdata(i).FileNum = BMPBuscado Then
        Call BuscarNuevoF(i)
        LastFound = i
        LUlitError.Caption = "F3 para continuar la busqueda"
        Exit Sub
    End If
Next i
LUlitError.Caption = " Se termino la busqueda"
MenuHerramientasBN.Visible = False
LastFound = 0
BMPBuscado = 0
End Sub

Private Sub MenuHerramientasBR_Click()
If FrmSearch.Visible Then
    FrmSearch.SetFocus
Else
    FrmSearch.Show , frmMain
End If
Call FrmSearch.HacerBusquedaR

End Sub

Private Sub MenuHerramientasNI_Click()
If FrmSearch.Visible Then
    FrmSearch.SetFocus
Else
    FrmSearch.Show , frmMain
End If
Call FrmSearch.HacerBusquedaNI
End Sub

Private Sub mnuinvobj_click()
    Dim textoActual As String
    Dim bmpstring As String
    Dim BMPBuscado As Long
    Dim nGrh As Long
    
    bmpstring = ReadField(1, Lista.List(Lista.listIndex), Asc(" "))
    
    If IsNumeric(bmpstring) Then
        BMPBuscado = Val(bmpstring)
        If BMPBuscado > 0 And BMPBuscado <= 32000 Then
            nGrh = BuscarGrHlibres(1)
            With Grhdata(nGrh)
                .FileNum = BMPBuscado
                .Frames(1) = nGrh
                .NumFrames = 1
                .pixelHeight = 32
                .pixelWidth = 32
                .TileHeight = 1
                .TileWidth = 1
                .sX = 0
                .sY = 0
                
                If EstadoIndexador <> e_EstadoIndexador.Grh Then CambiarEstado (e_EstadoIndexador.Grh)
                DoEvents
                EstadoNoGuardado(e_EstadoIndexador.Grh) = True
                Call BuscarNuevoF(nGrh)
            End With
        End If
    End If
End Sub


Private Sub mnIgeneral_Click()
Dim textoActual As String
Dim bmpstring As String
Dim BMPBuscado As Long


bmpstring = ReadField(1, Lista.List(Lista.listIndex), Asc(" "))

If IsNumeric(bmpstring) Then
    BMPBuscado = Val(bmpstring)
    If BMPBuscado > 0 And BMPBuscado <= 32000 Then
        If FormAuto.Visible Then
            FormAuto.SetFocus
        Else
            FormAuto.Show , frmMain
        End If
        
        FormAuto.FrameAnim(0).Visible = False
        FormAuto.FrameAnim(1).Visible = False
        FormAuto.FrameAnim(2).Visible = True
        FormAuto.loadTabStrip
        FormAuto.TextDatos3(4).Text = BMPBuscado
        FormAuto.TextDatos3(5).Text = BuscarGrHlibres(1)
    End If
End If

End Sub

Private Sub mnuConfig_Click()
    Call frmConfig.Show
End Sub

Private Sub mnuDatCasco_Click()
    frmDats.DatearClickDerecho eModo.Objetos, e_EstadoIndexador.Cascos
End Sub

Private Sub mnuDatEscudo_Click()
    frmDats.DatearClickDerecho eModo.Objetos, e_EstadoIndexador.Escudos
End Sub

Private Sub mnuDatEspadas_Click()
    frmDats.DatearClickDerecho eModo.Objetos, e_EstadoIndexador.Armas
End Sub

Private Sub mnuDatHechizo_Click()
    frmDats.DatearClickDerecho eModo.Hechizo, e_EstadoIndexador.FX
End Sub

Private Sub mnuDatNPC_Click()
    frmDats.DatearClickDerecho eModo.Npc, e_EstadoIndexador.Body
End Sub

Private Sub mnuDatObjeto_Click()
    frmDats.DatearClickDerecho eModo.Objetos, e_EstadoIndexador.Grh
End Sub

Private Sub mnuDatRopa_Click()
    frmDats.DatearClickDerecho eModo.Objetos, e_EstadoIndexador.Body
End Sub

Private Sub mnuGuardarTodo_Click()

    If EstadoNoGuardado(e_EstadoIndexador.Grh) = True Then
        Call SaveGrhData
    End If
    
    If EstadoNoGuardado(e_EstadoIndexador.Body) = True Then
        Call GuardarBodys
    End If
    
    If EstadoNoGuardado(e_EstadoIndexador.Cabezas) = True Then
        Call GuardarCabezas
    End If
    
    If EstadoNoGuardado(e_EstadoIndexador.Cascos) = True Then
        Call GuardarCascos
    End If
    
    If EstadoNoGuardado(e_EstadoIndexador.Escudos) Then
        Call GuardarEscudos
    End If
    
    If EstadoNoGuardado(e_EstadoIndexador.Armas) = True Then
        Call GuardarArmas
    End If
    
    If EstadoNoGuardado(e_EstadoIndexador.FX) = True Then
        Call GuardarFxs
    End If
    
    If EstadoNoGuardado(e_EstadoIndexador.Superficies) = True Then
        Call GuardarSuperficies
    End If
    
End Sub

Private Sub mnuibody_Click()
Dim textoActual As String
Dim bmpstring As String
Dim BMPBuscado As Long


bmpstring = ReadField(1, Lista.List(Lista.listIndex), Asc(" "))

If IsNumeric(bmpstring) Then
    BMPBuscado = Val(bmpstring)
    If BMPBuscado > 0 And BMPBuscado <= 32000 Then
        If FormAuto.Visible Then
            FormAuto.SetFocus
        Else
            FormAuto.Show , frmMain
        End If
        
        FormAuto.FrameAnim(1).Visible = True
        FormAuto.FrameAnim(0).Visible = False
        FormAuto.FrameAnim(2).Visible = False
        FormAuto.loadTabStrip
        FormAuto.Combo2.Visible = False
        FormAuto.Labelbody.Visible = False
        FormAuto.Labelbody1.Visible = False
        FormAuto.Labelbody2.Visible = False
        FormAuto.Loff.Visible = False
        FormAuto.Loffx.Visible = False
        FormAuto.Loffy.Visible = False
        FormAuto.TextDatos2(7).Visible = False
        FormAuto.TextDatos2(8).Visible = False
        FormAuto.TextDatos2(0).Enabled = False
        FormAuto.TextDatos2(1).Enabled = False
        FormAuto.TextDatos2(6).Enabled = False
        FormAuto.Text1.Visible = False
        FormAuto.Text2.Visible = False
        FormAuto.CheckAuto.Visible = False
        FormAuto.optionDimension(0).Visible = False
        FormAuto.optionDimension(1).Visible = False
        FormAuto.optionDimension(2).Visible = False
        FormAuto.Label5.Visible = False
        FormAuto.Label6.Visible = False
        FormAuto.TextDatos2(0).Text = 6
        FormAuto.TextDatos2(1).Text = 4
        FormAuto.TextDatos2(4).Text = BMPBuscado
        FormAuto.TextDatos2(6).Text = 22
        FormAuto.TextDatos2(2).Text = 46
        FormAuto.TextDatos2(3).Text = 26
        FormAuto.optionDimension(0).Value = True
        FormAuto.Labelbody.Visible = True
        FormAuto.Labelbody1.Visible = True
        FormAuto.Labelbody2.Visible = True
        FormAuto.Text1.Visible = True
        FormAuto.Text1.Enabled = False
        FormAuto.Text2.Visible = True
        FormAuto.Text2.Enabled = False
        FormAuto.CheckAuto.Visible = True
        FormAuto.CheckAuto.Value = vbUnchecked
        FormAuto.Text1.Text = UBound(BodyData) + 1
        FormAuto.Text2.Text = "-38º0"
        FormAuto.Combo2.Visible = True
        FormAuto.Combo2.listIndex = 0
        FormAuto.optionDimension(0).Visible = True
        FormAuto.optionDimension(1).Visible = True
        FormAuto.optionDimension(2).Visible = True
        FormAuto.Label5.Visible = True
        FormAuto.Label6.Visible = True
    End If
End If
 


End Sub

Private Sub mnuRecargarResource_Click()
    YaCargo = False
    If EstadoIndexador <> e_EstadoIndexador.Resource Then
    Call CambiarEstado(e_EstadoIndexador.Resource)
    Else
    Call RenuevaListaResource
    End If

End Sub

Private Sub NuevoGhr_Click()
Call BotonNuevoGRH
End Sub
Public Sub BotonNuevoGRH()
Dim cadena As String
Dim respuesta As Byte
Dim TempLong As Long
Select Case EstadoIndexador
    Case e_EstadoIndexador.Grh
            If GrHCambiando Then
                GrHCambiando = False
                LNumActual.Caption = "Ghr:"
                BotonGuardar.Visible = False
            End If
        cadena = InputBox("Introduzca el número de GHR (0 Para encontrar un grh que no este siendo utilizado)", "Nuevo Grafico")
        If IsNumeric(cadena) Then
            Call BuscarNuevoF(Val(cadena))
        Else
            LUlitError.Caption = "introduzca un numero"
        End If
    Case e_EstadoIndexador.Resource
        cadena = InputBox("Introduzca el número de BMP", "Nuevo Grafico")
        If IsNumeric(cadena) Then
            Call BuscarNuevoF(Val(cadena))
        Else
            LUlitError.Caption = "introduzca un numero"
        End If
        Exit Sub
    Case e_EstadoIndexador.Superficies
        cadena = InputBox("Introduzca el número de Superficie", "Nuevo Grafico")
        If IsNumeric(cadena) Then
            Call BuscarNuevoF(Val(cadena))
        Else
            LUlitError.Caption = "introduzca un numero"
        End If
        Exit Sub
        
    Case Else
        Dim StringCaso As String
        If EstadoIndexador = e_EstadoIndexador.Body Then
            StringCaso = "Body"
        ElseIf EstadoIndexador = e_EstadoIndexador.Cabezas Then
            StringCaso = "Cabeza"
        ElseIf EstadoIndexador = e_EstadoIndexador.Cascos Then
            StringCaso = "Casco"
        ElseIf EstadoIndexador = e_EstadoIndexador.Escudos Then
            StringCaso = "Escudo"
        ElseIf EstadoIndexador = e_EstadoIndexador.Armas Then
            StringCaso = "Arma"
        ElseIf EstadoIndexador = e_EstadoIndexador.FX Then
            StringCaso = "Fx"
        End If
        
        cadena = InputBox("Introduzca " & StringCaso & " (0 Para encontrar un " & StringCaso & " libre)", "Nuevo " & StringCaso & "/buscar")
        If IsNumeric(cadena) Then
            Call BuscarNuevoF(Val(cadena))
        Else
            LUlitError.Caption = "introduzca un numero"
        End If
End Select
End Sub
Public Sub BuscarNuevoF(ByVal Index As Long)
Dim cadena As String
Dim respuesta As Byte
Dim TempLong As Long

        If Index > 0 And Index <= 60000 Then
            TempLong = ListaindexGrH(Index)
            If TempLong >= 0 Then
                GRHActual = Index
                LOOPActual = 1
                frmMain.Visor.Cls
                'Call DibujarBMPVisor(GRHActual)
                Call GetInfoBmp(GRHActual)
                Call DibujarBMPVisor(GRHActual)
                LGHRnumeroA.Caption = GRHActual
                frmMain.Lista.listIndex = TempLong
            Else
                LUlitError.Caption = "Bmp no existe"
            End If
        Else
            LUlitError.Caption = "Valor no valido"
        End If
        Exit Sub
        
End Sub


Private Sub TextDatos_DblClick(Index As Integer)
If EstadoIndexador = e_EstadoIndexador.Grh Or Index > 4 Or _
(EstadoIndexador = e_EstadoIndexador.FX) And Index > 0 Then Exit Sub
If Val(TextDatos(Index).Text) > 0 And Val(TextDatos(Index).Text) < MAXGrH Then

    If EstadoIndexador <> e_EstadoIndexador.Grh Then Call CambiarEstado(e_EstadoIndexador.Grh)
    Call BuscarNuevoF(TextDatos(Index).Text)
End If
End Sub
Private Sub TextDatos_Change(Index As Integer)
'Comprueba que los datos introducidos son correctos

Dim Ancho As Long
Dim Alto As Long
Dim PrimerAncho As Long
Dim PrimerAlto As Long
Dim i As Long
Dim Algun_Error As Boolean
Dim ErroresGrh As ErroresGrh
Dim tdouble1 As Double, tdouble2 As Double



If EstadoIndexador = e_EstadoIndexador.Resource Then Exit Sub

2 For i = 0 To 7
    If i <> 1 And ((i <> 5) Or EstadoIndexador <> e_EstadoIndexador.Body) And ((i <> 2) Or EstadoIndexador <> FX) Then ' el 1 son los frames y el 5 se usa para offset
        If Val(TextDatos(i).Text) > MAXGrH Then
            TextDatos(i).Text = MAXGrH
        End If
    ElseIf ((i = 5) And EstadoIndexador = e_EstadoIndexador.Body) Or ((i = 2) And EstadoIndexador = FX) Then
        tdouble1 = Val(ReadField(1, TextDatos(i).Text, Asc("º")))
        tdouble2 = Val(ReadField(2, TextDatos(i).Text, Asc("º")))
        If tdouble1 < -32000 Or tdouble1 > 32000 Then
            TextDatos(i).Text = "0º" & tdouble2
            tdouble1 = 0
        End If
        
        If tdouble2 < -32000 Or tdouble2 > 32000 Then
            TextDatos(i).Text = tdouble1 & "º0"
        End If

    End If
    ErroresGrh.colores(i) = vbWhite
Next i

ErroresGrh.colores(8) = vbWhite
ErroresGrh.colores(9) = vbWhite


LUlitError.Caption = ""
Dim resul As Long
Dim MensageError As String

Select Case EstadoIndexador
    Case e_EstadoIndexador.Superficies
        If Index = 1 Then
            If Lista.listIndex <= UBound(SupData) Then
                'Lista.List(Lista.listIndex) = Lista.listIndex & " - " & TextDatos(1)
                'guardarSuperficie Lista.listIndex
                'SupData(Lista.listIndex).Nombre = TextDatos(1).Text
            End If
        End If
    Case e_EstadoIndexador.Grh
        If Not GrHCambiando Then
            GrHCambiando = True
            TempGrh = Grhdata(GRHActual)
            LNumActual.Caption = "**Ghr:"
            BotonGuardar.Visible = True
        End If
            If Val(TextDatos(5).Text) > MAXGrH Then
                TextDatos(5).Text = MAXGrH
            End If
            If Val(TextDatos(2).Text) > 25 Then 'numframes > 25
                TextDatos(2).Text = 25
            ElseIf Val(TextDatos(2).Text) < 1 Then 'numframes < 1
                TextDatos(2).Text = 1
            End If
            
            If Val(TextDatos(2).Text) = 1 Then ' Es grh normal
                TextDatos(1).Enabled = False
                For i = 3 To 4
                    TextDatos(i).Enabled = True
                Next i
                TextDatos(5).Enabled = False
                For i = 6 To 7
                    TextDatos(i).Enabled = True
                Next i
            ElseIf Val(TextDatos(2).Text) > 1 Then ' es animacion
                TextDatos(1).Enabled = True
                For i = 3 To 4
                    TextDatos(i).Enabled = False
                Next i
                TextDatos(5).Enabled = True
                For i = 6 To 7
                    TextDatos(i).Enabled = False
                Next i
            End If
            

            TempGrh.FileNum = Val(TextDatos(0).Text)
            TempGrh.NumFrames = Val(TextDatos(2).Text)
            If TempGrh.NumFrames = 1 Then
                TempGrh.Frames(1) = Val(ReadField(1, TextDatos(1).Text, Asc("-")))
            Else
                For i = 1 To TempGrh.NumFrames
                    If Val(ReadField(i, TextDatos(1).Text, Asc("-"))) < 32000 Then
                        TempGrh.Frames(i) = Val(ReadField(i, TextDatos(1).Text, Asc("-")))
                    End If
                Next i
            End If
            TempGrh.pixelHeight = Val(TextDatos(3).Text)
            TempGrh.pixelWidth = Val(TextDatos(4).Text)
            TempGrh.Speed = Val(TextDatos(5).Text)
            TempGrh.sX = Val(TextDatos(6).Text)
            TempGrh.sY = Val(TextDatos(7).Text)
            TempGrh.TileHeight = Grhdata(GRHActual).pixelHeight / TilePixelHeight
            TextDatos(8).Text = TempGrh.TileHeight
            TempGrh.TileWidth = Grhdata(GRHActual).pixelWidth / TilePixelWidth
            TextDatos(9).Text = TempGrh.TileWidth
            

            resul = GrhCorrecto(TempGrh, MensageError, ErroresGrh)
            LUlitError.Caption = MensageError
            
            For i = 0 To 9
                TextDatos(i).BackColor = ErroresGrh.colores(i)
            Next i
            
            If ErroresGrh.ErrorCritico Then
                BotonGuardar.Visible = False
                Exit Sub
            Else
                BotonGuardar.Visible = True
            End If
            
            frmMain.Visor.Cls
            If Not LoadingNew Then Call DibujarGHRVisor(GRHActual)
            If TempGrh.NumFrames = 1 Then
                frmMain.Dibujado.Enabled = False
            ElseIf TempGrh.NumFrames > 1 Then
                If TempGrh.Speed > 0 Then ' pervenimos division por 0
                    frmMain.Dibujado.Interval = 150 '(TempGrh.Speed / 5)
                Else
                    frmMain.Dibujado.Interval = 100
                End If
                frmMain.Dibujado.Enabled = True
            Else
                frmMain.Dibujado.Enabled = False
            End If
    
    Case Else
        If Not GrHCambiando Then
            GrHCambiando = True
            BotonGuardar.Visible = True
        End If
        
            If Not LoadingNew Then frmMain.Visor.Cls ' Si no estamos cargando limpiamos
            
            Dibujado.Interval = 100
            Dibujado.Enabled = True
            If EstadoIndexador = e_EstadoIndexador.Body Then
                tempDataIndex.HeadOffset.y = Val(ReadField(1, TextDatos(5).Text, Asc("º")))
                tempDataIndex.HeadOffset.x = Val(ReadField(2, TextDatos(5).Text, Asc("º")))
            End If
            Dim III As Long
            Dim tStr As String
            Algun_Error = False
            For i = 1 To 4
                If i = 1 Then
                    III = 0
                Else
                    III = i
                End If
                If i = 1 Then
                    If EstadoIndexador = e_EstadoIndexador.FX Then
                        tStr = "FX"
                    Else
                        tStr = "Arriba"
                    End If
                ElseIf i = 2 Then
                    tStr = "Derecha"
                ElseIf i = 3 Then
                    tStr = "Abajo"
                ElseIf i = 4 Then
                    tStr = "Izquierda"
                End If
                If i = 1 Or EstadoIndexador <> e_EstadoIndexador.FX Then
                tempDataIndex.Walk(i).GrhIndex = Val(TextDatos(III).Text)
                If tempDataIndex.Walk(i).GrhIndex > 1 Then
                    MensageError = ""
                    resul = GrhCorrecto(Grhdata(tempDataIndex.Walk(i).GrhIndex), MensageError, ErroresGrh)
                    If ErroresGrh.ErrorCritico Then
                        Algun_Error = True
                        TextDatos(III).BackColor = vbRed
                        LUlitError.Caption = LUlitError.Caption & "(" & tStr & ") " & MensageError & vbCrLf
                    Else
                        If EstadoIndexador = e_EstadoIndexador.Cabezas Or EstadoIndexador = e_EstadoIndexador.Cascos Then
                            If ErroresGrh.EsAnimacion Then
                                TextDatos(III).BackColor = vbYellow
                                LUlitError.Caption = LUlitError.Caption & "(" & tStr & ") Es una animacion" & vbCrLf
                            Else
                                TextDatos(III).BackColor = vbWhite
                            End If
                        Else
                            If Not ErroresGrh.EsAnimacion Then
                                TextDatos(III).BackColor = vbYellow
                                LUlitError.Caption = LUlitError.Caption & "(" & tStr & ") No es una animacion" & vbCrLf
                            Else
                                TextDatos(III).BackColor = vbWhite
                            End If
                        End If
                    End If
                End If
                End If
            Next i
            If Algun_Error Then
                BotonGuardar.Visible = False
            Else
                BotonGuardar.Visible = True
            End If
        
End Select
End Sub

