VERSION 5.00
Begin VB.Form frmCargando 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   3345
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7350
   LinkTopic       =   "Form1"
   ScaleHeight     =   3345
   ScaleWidth      =   7350
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   3375
      Left            =   0
      ScaleHeight     =   3345
      ScaleWidth      =   7305
      TabIndex        =   0
      Top             =   0
      Width           =   7335
      Begin VB.PictureBox Picture2 
         BorderStyle     =   0  'None
         Height          =   735
         Left            =   120
         ScaleHeight     =   49
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   465
         TabIndex        =   3
         Top             =   2280
         Width           =   6975
         Begin VB.Shape shpProgress 
            BackColor       =   &H00FF8080&
            BackStyle       =   1  'Opaque
            Height          =   735
            Left            =   0
            Top             =   0
            Width           =   6975
         End
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Index-Dater 1.0"
         BeginProperty Font 
            Name            =   "MS UI Gothic"
            Size            =   48
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080FF80&
         Height          =   1335
         Left            =   120
         TabIndex        =   2
         Top             =   120
         Width           =   7095
      End
      Begin VB.Label lblEstado 
         Alignment       =   2  'Center
         BackColor       =   &H0000FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Cargando Obj.dat"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   22.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   735
         Left            =   600
         TabIndex        =   1
         Top             =   1560
         Width           =   6255
      End
   End
End
Attribute VB_Name = "frmCargando"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
