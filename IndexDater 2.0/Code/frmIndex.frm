VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmIndex 
   Caption         =   "Indexar grafico"
   ClientHeight    =   6570
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   7095
   LinkTopic       =   "Form1"
   ScaleHeight     =   6570
   ScaleWidth      =   7095
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      BorderStyle     =   0  'None
      Caption         =   "Superficies"
      ClipControls    =   0   'False
      Height          =   5895
      Left            =   7200
      TabIndex        =   3
      Top             =   1560
      Width           =   6495
      Begin VB.TextBox Text2 
         Height          =   375
         Left            =   2280
         TabIndex        =   4
         Text            =   "Text1"
         Top             =   2760
         Width           =   2535
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Animaciones y npcs"
      Height          =   5895
      Left            =   240
      TabIndex        =   1
      Top             =   480
      Width           =   6495
      Begin VB.TextBox Text1 
         Height          =   375
         Left            =   600
         TabIndex        =   2
         Text            =   "Text1"
         Top             =   480
         Width           =   2535
      End
   End
   Begin MSComctlLib.TabStrip TabStrip1 
      Height          =   6375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6855
      _ExtentX        =   12091
      _ExtentY        =   11245
      MultiRow        =   -1  'True
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   3
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Animaciones y npcs"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Superficies"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Bodys/armas/escudos"
            ImageVarType    =   2
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "frmIndex"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub TabStrip1_Click()
Dim i As Integer

i = TabStrip1.SelectedItem.Index

Frame1(i - 1).ZOrder
End Sub
