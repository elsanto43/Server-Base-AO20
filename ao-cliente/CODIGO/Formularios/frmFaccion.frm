VERSION 5.00
Begin VB.Form frmFaccion 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Elegir faccion"
   ClientHeight    =   1305
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   5910
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1305
   ScaleWidth      =   5910
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin AOLibre.uAOButton bFacc 
      Height          =   855
      Index           =   0
      Left            =   2040
      TabIndex        =   0
      Top             =   240
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   1508
      TX              =   "Neutral"
      ENAB            =   -1  'True
      FCOL            =   14737632
      OCOL            =   16777215
      PICE            =   "frmFaccion.frx":0000
      PICF            =   "frmFaccion.frx":001C
      PICH            =   "frmFaccion.frx":0038
      PICV            =   "frmFaccion.frx":0054
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Papyrus"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin AOLibre.uAOButton bFacc 
      Height          =   855
      Index           =   1
      Left            =   120
      TabIndex        =   1
      Top             =   240
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   1508
      TX              =   "Armada "
      ENAB            =   -1  'True
      FCOL            =   16776960
      OCOL            =   16777215
      PICE            =   "frmFaccion.frx":0070
      PICF            =   "frmFaccion.frx":008C
      PICH            =   "frmFaccion.frx":00A8
      PICV            =   "frmFaccion.frx":00C4
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Papyrus"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin AOLibre.uAOButton bFacc 
      Height          =   855
      Index           =   2
      Left            =   3960
      TabIndex        =   2
      Top             =   240
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   1508
      TX              =   "Legion"
      ENAB            =   -1  'True
      FCOL            =   8421631
      OCOL            =   16777215
      PICE            =   "frmFaccion.frx":00E0
      PICF            =   "frmFaccion.frx":00FC
      PICH            =   "frmFaccion.frx":0118
      PICV            =   "frmFaccion.frx":0134
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Papyrus"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "frmFaccion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub bFacc_Click(Index As Integer)
    Dim t As String
    If Index = 1 Then t = "la Armada Imperial"
    If Index = 2 Then t = "la Horda Infernal"
    
    If MsgBox("Estas seguro que deseas pertenecer a " & t & "?", vbYesNo) = vbYes Then _
        Call WriteElegirFaccion(CByte(Index))
End Sub
