VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmConfig 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Configurar carpetas"
   ClientHeight    =   4545
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   8895
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4545
   ScaleWidth      =   8895
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   3480
      Top             =   1680
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Cerrar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   15
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   4560
      TabIndex        =   13
      Top             =   3600
      Width           =   4215
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Guardar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   15
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   240
      TabIndex        =   12
      Top             =   3600
      Width           =   4215
   End
   Begin VB.CommandButton cmdCarpeta 
      Caption         =   "Cargar..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   3
      Left            =   7080
      TabIndex        =   10
      Top             =   3000
      Width           =   1695
   End
   Begin VB.TextBox txtDir 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   15
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Index           =   3
      Left            =   120
      TabIndex        =   9
      Top             =   3000
      Width           =   6855
   End
   Begin VB.CommandButton cmdCarpeta 
      Caption         =   "Cargar..."
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   2
      Left            =   7080
      TabIndex        =   7
      Top             =   2160
      Width           =   1695
   End
   Begin VB.TextBox txtDir 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   15
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Index           =   2
      Left            =   120
      TabIndex        =   6
      Top             =   2160
      Width           =   6855
   End
   Begin VB.CommandButton cmdCarpeta 
      Caption         =   "Cargar..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   1
      Left            =   7080
      TabIndex        =   3
      Top             =   1320
      Width           =   1695
   End
   Begin VB.CommandButton cmdCarpeta 
      Caption         =   "Cargar..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   0
      Left            =   7080
      TabIndex        =   2
      Top             =   480
      Width           =   1695
   End
   Begin VB.TextBox txtDir 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   15
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Index           =   1
      Left            =   120
      TabIndex        =   1
      Top             =   1320
      Width           =   6855
   End
   Begin VB.TextBox txtDir 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   15
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   6855
   End
   Begin VB.Label Label4 
      Caption         =   "Carpeta donde esta indices.ini (WE)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   11
      Top             =   2760
      Width           =   3255
   End
   Begin VB.Label Label3 
      Caption         =   "Carpeta de dats"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   1920
      Width           =   3255
   End
   Begin VB.Label Label2 
      Caption         =   "Carpeta de graficos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   1080
      Width           =   3255
   End
   Begin VB.Label Label1 
      Caption         =   "Carpeta de inits"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   240
      Width           =   3255
   End
End
Attribute VB_Name = "frmConfig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Enum eDirs
    Inits = 0
    grafs = 1
    Dats = 2
    InitWE = 3
End Enum

Private Sub cmdCarpeta_Click(Index As Integer)
    Static lasti As Integer
    
txtDir(Index) = BuscarCArpeta(, txtDir(lasti).Text)
lasti = Index
End Sub

Private Sub Dir1_Change()

End Sub


Private Sub Dir1_Validate(Cancel As Boolean)

End Sub


Private Sub Command5_Click()
    ConfigDir.Inits = txtDir(eDirs.Inits)
    ConfigDir.Graficos = txtDir(eDirs.grafs)
    ConfigDir.Dats = txtDir(eDirs.Dats)
    ConfigDir.InitWE = txtDir(eDirs.InitWE)
    guardarConfig
    reConfigurarPath = False
    
    frmMain.Lista.Clear
    
    'If MsgBox("Desea reiniciar el indexdater para ver los cambios?", vbYesNo, "Cambios en directorios") = vbYes Then
        
   ' End If
    
    Call CargarCuerpos
    Call CargarCabezas
    Call CargarCascos
   ' Call CargarSuperficies
    'Call CargarEspalda
    'Call CargarBotas
    Call CargarFxs
    Call CargarAnimsExtra
    Call CargarTips

    Call CargarArrayLluvia
    Call CargarAnimArmas
    Call CargarAnimEscudos
    
    Call LoadGrhData

'    Call InitializeDats
    If Not frmMain.Visible Then frmMain.Show vbModal
        Unload Me

End Sub

Private Sub Command6_Click()
End
End Sub

Private Sub Form_Load()
    txtDir(0).Text = ConfigDir.Inits
    txtDir(1).Text = ConfigDir.Graficos
    txtDir(2).Text = ConfigDir.Dats
    txtDir(3).Text = ConfigDir.InitWE
    
End Sub










