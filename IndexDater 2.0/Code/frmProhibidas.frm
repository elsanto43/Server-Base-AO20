VERSION 5.00
Begin VB.Form frmSelectBMP 
   BackColor       =   &H00000000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Seleccionar BMP's"
   ClientHeight    =   9630
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   13530
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9630
   ScaleWidth      =   13530
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command8 
      Caption         =   "Buscar GRH libre"
      Height          =   320
      Left            =   2400
      TabIndex        =   16
      Top             =   6680
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   1200
      TabIndex        =   14
      Top             =   6680
      Width           =   1215
   End
   Begin VB.CheckBox Check1 
      BackColor       =   &H00000000&
      Caption         =   "Guardar en el indice de FX"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   13
      Top             =   7800
      Width           =   3255
   End
   Begin VB.CommandButton Command7 
      Caption         =   "UP"
      Height          =   495
      Left            =   4920
      TabIndex        =   12
      Top             =   2640
      Width           =   735
   End
   Begin VB.CommandButton Command6 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "Simular indexacion"
      Height          =   675
      Left            =   3600
      MaskColor       =   &H00808080&
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   8640
      Width           =   9735
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   80
      Left            =   4200
      Top             =   6240
   End
   Begin VB.PictureBox picAnim 
      AutoRedraw      =   -1  'True
      Height          =   8295
      Left            =   5880
      ScaleHeight     =   8235
      ScaleWidth      =   7395
      TabIndex        =   10
      Top             =   240
      Width           =   7455
   End
   Begin VB.CommandButton Command5 
      Caption         =   "DOWN"
      Height          =   495
      Left            =   4920
      TabIndex        =   8
      Top             =   3120
      Width           =   735
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Quitar"
      Height          =   495
      Left            =   2640
      TabIndex        =   7
      Top             =   5160
      Width           =   2295
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Cancelar"
      Height          =   495
      Left            =   120
      TabIndex        =   6
      Top             =   8760
      Width           =   3255
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Guardar indexacion"
      Height          =   495
      Left            =   120
      TabIndex        =   5
      Top             =   8160
      Width           =   3255
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Agregar"
      Height          =   495
      Left            =   120
      TabIndex        =   4
      Top             =   5760
      Width           =   2295
   End
   Begin VB.ListBox List2 
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   3960
      Left            =   2640
      TabIndex        =   1
      Top             =   1080
      Width           =   2295
   End
   Begin VB.ListBox List1 
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   5325
      Left            =   120
      MultiSelect     =   2  'Extended
      TabIndex        =   0
      Top             =   360
      Width           =   2295
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Primer GRH"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Index           =   2
      Left            =   120
      TabIndex        =   15
      Top             =   6720
      Width           =   1215
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Para que la animacion sea guardada correctamente, los BMPS deben estar en orden"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   735
      Index           =   1
      Left            =   120
      TabIndex        =   9
      Top             =   7080
      Width           =   3135
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "BMPs de la animacion"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   0
      Left            =   2640
      TabIndex        =   3
      Top             =   720
      Width           =   2295
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Todos los BMPS"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   2295
   End
End
Attribute VB_Name = "frmSelectBMP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()

    Dim k As Long
    Dim s As String
    Dim cant As Integer

    For k = 0 To List1.ListCount - 1
        If List1.Selected(k) Then
            If List2.ListCount < 25 Then
                s = List1.List(k)
                List2.AddItem s 'List1.List(List1.listIndex)
                List2.listIndex = List2.ListCount - 1
                cant = cant + 1
            Else
                MsgBox "El sistema de GRH de la version FENIX Y 0.9.9z tiene un maximo de 25 frames"
                Exit For
                
            End If
        End If
    Next k
End Sub

Private Sub Command2_Click()
    If List2.listIndex < 0 Then Exit Sub

    List2.RemoveItem List2.listIndex
End Sub



Private Sub Command3_Click()
Dim FramesTotales As Integer
Dim NumeroBMP As Integer, PrimerIndice As Long
Dim existenciaBMP As Integer
Dim Alto As Long, Ancho As Long
Dim AltoBMP As Long, AnchoBMP As Long
Dim BitCount As Integer
Dim ii As Integer
Dim FramesX As Long, FramesY As Long
Dim actualFrame As Integer
Dim curx As Long, cury As Long
    
    PrimerIndice = Val(Me.Text1.Text)
    FramesTotales = List2.ListCount

    If (Not hayGrHlibres(PrimerIndice, FramesTotales + 1)) Or PrimerIndice <= 0 Or PrimerIndice > MAXGrH Then
        MsgBox "No hay sitio para la animacion" & vbCrLf & "Sobreescribir x implementar"
    Exit Sub
    End If

    For FramesY = 1 To FramesTotales
        NumeroBMP = Val(List2.List(FramesY - 1))
        Call GetTamañoBMP(NumeroBMP, AltoBMP, AnchoBMP, BitCount)
        
        Grhdata(PrimerIndice + actualFrame).FileNum = NumeroBMP
        Grhdata(PrimerIndice + actualFrame).Frames(1) = PrimerIndice + actualFrame
        Grhdata(PrimerIndice + actualFrame).NumFrames = 1
        Grhdata(PrimerIndice + actualFrame).pixelHeight = AltoBMP
        Grhdata(PrimerIndice + actualFrame).pixelWidth = AnchoBMP
        Grhdata(PrimerIndice + actualFrame).sX = 0
        Grhdata(PrimerIndice + actualFrame).sY = 0
        Grhdata(PrimerIndice + actualFrame).TileHeight = AltoBMP / TilePixelHeight
        Grhdata(PrimerIndice + actualFrame).TileWidth = AnchoBMP / TilePixelWidth
        actualFrame = actualFrame + 1
        If actualFrame >= FramesTotales Then GoTo TerminarAnim
    Next FramesY
    
    
    
    
TerminarAnim:

EstadoNoGuardado(e_EstadoIndexador.Grh) = True
Grhdata(PrimerIndice + FramesTotales).FileNum = NumeroBMP
Grhdata(PrimerIndice + FramesTotales).NumFrames = FramesTotales

Grhdata(PrimerIndice + FramesTotales).pixelHeight = Grhdata(PrimerIndice).pixelHeight
Grhdata(PrimerIndice + FramesTotales).pixelWidth = Grhdata(PrimerIndice).pixelWidth
Grhdata(PrimerIndice + FramesTotales).sX = Grhdata(PrimerIndice).sX
Grhdata(PrimerIndice + FramesTotales).sY = Grhdata(PrimerIndice).sY
Grhdata(PrimerIndice + FramesTotales).TileHeight = Grhdata(PrimerIndice).TileHeight
Grhdata(PrimerIndice + FramesTotales).TileWidth = Grhdata(PrimerIndice).TileWidth
            
Dim tS As String

tS = Round(FramesTotales / 2)
For ii = 1 To FramesTotales
    Grhdata(PrimerIndice + FramesTotales).Frames(ii) = PrimerIndice + ii - 1
Next ii
Grhdata(PrimerIndice + FramesTotales).Speed = CSng(tS & tS & tS & "." & tS & tS & tS)

If Check1.Value = 1 Then
    Dim nINdex As Integer
    nINdex = UBound(FxData) + 1
    Call AgregaFx(nINdex)
    FxData(nINdex).FX.GrhIndex = PrimerIndice + FramesTotales
    FxData(nINdex).FX.Speed = CSng(tS & tS & tS & "." & tS & tS & tS)
    EstadoNoGuardado(e_EstadoIndexador.FX) = True
    Call frmMain.CambiarEstado(e_EstadoIndexador.FX)
    Call frmMain.BuscarNuevoF(nINdex)
Else
    Call frmMain.CambiarEstado(e_EstadoIndexador.Grh)
    Call frmMain.BuscarNuevoF(PrimerIndice)
End If

Unload Me

End Sub

Private Sub Command4_Click()
    Unload Me
End Sub

Private Sub Command5_Click()
    Dim nStore As Long
    If List2.ListCount < 2 Then Exit Sub
    If List2.listIndex = (List2.ListCount - 1) Then Exit Sub
    
    
    nStore = Val(List2.List(List2.listIndex))
    List2.List(List2.listIndex) = List2.List(List2.listIndex + 1)
    List2.List(List2.listIndex + 1) = nStore
    
    List2.listIndex = List2.listIndex + 1
End Sub

Private Sub Command6_Click()
    Timer1.Enabled = Not Timer1.Enabled
    If Timer1.Enabled = False Then
        Command6.BackColor = vbRed
    Else
        Command6.BackColor = vbGreen
    End If
End Sub

Private Sub Command7_Click()
    Dim nStore As Long
    If List2.ListCount < 2 Then Exit Sub
    If List2.listIndex <= 0 Then Exit Sub
    nStore = Val(List2.List(List2.listIndex))
    List2.List(List2.listIndex) = List2.List(List2.listIndex - 1)
    List2.List(List2.listIndex - 1) = nStore
    List2.listIndex = List2.listIndex - 1
End Sub

Private Sub Command8_Click()
    Text1.Text = BuscarGrHlibres(List2.ListCount + 1)
End Sub

Private Sub Form_Load()
    Dim i As Long
    If EstadoIndexador <> e_EstadoIndexador.Resource Then Call frmMain.CambiarEstado(e_EstadoIndexador.Resource)
    
    For i = 1 To frmMain.Lista.ListCount
        List1.AddItem frmMain.Lista.List(i)
        
    Next i
    Command6.BackColor = vbRed
    If frmMain.Lista.listIndex < 0 Then Exit Sub
    
    List1.listIndex = frmMain.Lista.listIndex
    List1.Selected(List1.listIndex) = True
End Sub

Private Sub List1_Click()
    If List1.listIndex < 0 Then Exit Sub
    Call DrawBMP(BackBufferSurface, Val(List1.List(List1.listIndex)))
End Sub

Private Sub List2_Click()
    If List2.listIndex < 0 Then Exit Sub
    Call DrawBMP(BackBufferSurface, Val(List2.List(List2.listIndex)))
End Sub

Private Sub Text1_Change()
If Not hayGrHlibres(Val(Text1.Text), List2.ListCount + 1) Then
    Text1.BackColor = vbRed
Else
    Text1.BackColor = RGB(180, 255, 220)
End If

End Sub

Private Sub Timer1_Timer()
    If List2.ListCount <= 0 Then Exit Sub
    Static aGrh As Byte
    
    If Me.WindowState = vbMinimized Then Exit Sub
    Call DrawBMP(BackBufferSurface, Val(List2.List(aGrh)))
    aGrh = aGrh + 1
    If aGrh >= List2.ListCount Then aGrh = 0
End Sub






Sub DrawBMP(Surface As DirectDrawSurface7, FileNum As Integer)
On Error Resume Next
Dim r2 As RECT, auxr As RECT, auxr2 As RECT
Dim r As RECT
Dim iGrhIndex As Long
Dim SurfaceDesc As DDSURFACEDESC2
Dim ddsd As DDSURFACEDESC2
Dim ddck As DDCOLORKEY
Dim surfacecuadro  As DirectDrawSurface7
Dim ii As Long

picAnim.Cls


If DibujarFondo Then
    BackBufferSurface.BltColorFill r, ColorFondo
Else
    BackBufferSurface.BltColorFill r, 0
End If



If FileNum <= 0 Then Exit Sub

SurfaceDB.Surface(FileNum).GetSurfaceDesc SurfaceDesc

With r2
   .Left = 0
   .Top = 0
    .Right = SurfaceDesc.lWidth
    .Bottom = SurfaceDesc.lHeight
   If .Bottom = 1024 Then .Bottom = 990
End With

auxr = r2
auxr2 = auxr


Surface.BltFast 0, 0, SurfaceDB.Surface(FileNum), r2, DDBLTFAST_SRCCOLORKEY Or DDBLTFAST_WAIT


Surface.BltToDC picAnim.hdc, auxr2, auxr



picAnim.Refresh
End Sub





