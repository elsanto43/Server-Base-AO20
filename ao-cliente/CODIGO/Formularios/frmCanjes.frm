VERSION 5.00
Begin VB.Form frmCanjes 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Canjes"
   ClientHeight    =   4035
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5820
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4035
   ScaleWidth      =   5820
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin AOLibre.uAOButton uAOButton1 
      Height          =   615
      Left            =   120
      TabIndex        =   2
      Top             =   3240
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   1085
      TX              =   ""
      ENAB            =   -1  'True
      FCOL            =   16777215
      OCOL            =   16777215
      PICE            =   "frmCanjes.frx":0000
      PICF            =   "frmCanjes.frx":001C
      PICH            =   "frmCanjes.frx":0038
      PICV            =   "frmCanjes.frx":0054
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Height          =   2955
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   3015
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      DrawStyle       =   3  'Dash-Dot
      ForeColor       =   &H80000008&
      Height          =   540
      Left            =   4200
      ScaleHeight     =   34
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   34
      TabIndex        =   0
      Top             =   240
      Width           =   540
   End
   Begin AOLibre.uAOButton uAOButton2 
      Height          =   615
      Left            =   3000
      TabIndex        =   3
      Top             =   3240
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   1085
      TX              =   "Canjear objeto"
      ENAB            =   -1  'True
      FCOL            =   7314354
      OCOL            =   16777215
      PICE            =   "frmCanjes.frx":0070
      PICF            =   "frmCanjes.frx":008C
      PICH            =   "frmCanjes.frx":00A8
      PICV            =   "frmCanjes.frx":00C4
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label lblCerrar 
      BackStyle       =   0  'Transparent
      Caption         =   "x"
      BeginProperty Font 
         Name            =   "Myriad Hebrew"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   5520
      MousePointer    =   99  'Custom
      TabIndex        =   8
      Top             =   0
      Width           =   255
   End
   Begin VB.Label bronces 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Trofeos de bronce: "
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
      Left            =   3240
      TabIndex        =   7
      Top             =   2400
      Width           =   2535
   End
   Begin VB.Label platas 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Trofeos de plata: "
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
      Left            =   3240
      TabIndex        =   6
      Top             =   2040
      Width           =   2535
   End
   Begin VB.Label oros 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Trofeos de oro: "
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
      Left            =   3240
      TabIndex        =   5
      Top             =   1680
      Width           =   2535
   End
   Begin VB.Label lblname 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Armadura de pieles de oso pardo"
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
      Left            =   3240
      TabIndex        =   4
      Top             =   960
      Width           =   2535
   End
End
Attribute VB_Name = "frmCanjes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Label1_Click()

End Sub

Private Sub Label3_Click()

End Sub

Private Sub lblCerrar_Click()
    Unload Me
End Sub

Private Sub list1_Click()
    If List1.ListIndex >= 0 Then
        
        Dim iCanje As Integer
        iCanje = List1.ListIndex + 1
        lblName.Caption = canjes(iCanje).Nombre
        
        oros.Caption = "Trofeos de oro: " & canjes(iCanje).oros
        platas.Caption = "Trofeos de plata: " & canjes(iCanje).platas
        bronces.Caption = "Trofeos de bronce: " & canjes(iCanje).bronces
        
        Call RenderItem(Picture1, canjes(iCanje).GrhIndex)
    End If
End Sub

Private Sub uAOButton1_Click()
    If List1.ListIndex >= 0 Then
        If MsgBox("Seguro que deseas canjear '" & canjes(List1.ListIndex + 1).Nombre & "' ?", vbYesNo) = vbYes Then _
            Call writeCanjearItem(List1.ListIndex + 1)
    End If
End Sub

