VERSION 5.00
Begin VB.Form FormAuto 
   Caption         =   "Indexador automatico"
   ClientHeight    =   7500
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6390
   ScaleHeight     =   7500
   ScaleWidth      =   6390
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command13 
      Appearance      =   0  'Flat
      Caption         =   "Superficies, decoraciones, arboles, techos, casas, etc"
      Height          =   615
      Index           =   2
      Left            =   4200
      Style           =   1  'Graphical
      TabIndex        =   104
      Top             =   0
      Width           =   2055
   End
   Begin VB.CommandButton Command13 
      Appearance      =   0  'Flat
      Caption         =   "NPC's y chars(Body, escudo, arma, cascos, cabezas)"
      Height          =   615
      Index           =   1
      Left            =   2040
      Style           =   1  'Graphical
      TabIndex        =   103
      Top             =   0
      Width           =   2175
   End
   Begin VB.CommandButton Command13 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000A&
      Caption         =   "Hechizos y FX's"
      Height          =   615
      Index           =   0
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   102
      Top             =   0
      Width           =   2055
   End
   Begin VB.Frame FrameAnim 
      Caption         =   "Superficies y objetos para mapas"
      Height          =   6615
      Index           =   2
      Left            =   0
      TabIndex        =   0
      Top             =   720
      Visible         =   0   'False
      Width           =   6255
      Begin VB.ComboBox cmbCapa 
         Height          =   315
         Left            =   4800
         TabIndex        =   108
         Top             =   2400
         Width           =   1095
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Autocalcular por tamaño de imagen"
         Height          =   495
         Left            =   4080
         TabIndex        =   101
         Top             =   3000
         Width           =   1575
      End
      Begin VB.CheckBox Check2 
         Caption         =   "Indexar como superficie en el indice del worldeditor"
         Height          =   375
         Left            =   240
         TabIndex        =   99
         Top             =   4440
         Width           =   4335
      End
      Begin VB.TextBox txtNombre 
         Height          =   285
         Left            =   1920
         TabIndex        =   57
         Text            =   "Nueva superficie"
         Top             =   960
         Width           =   3135
      End
      Begin VB.CommandButton Command12 
         Caption         =   "-32"
         Height          =   255
         Left            =   1680
         TabIndex        =   52
         Top             =   2520
         Width           =   495
      End
      Begin VB.CommandButton Command11 
         Caption         =   "-32"
         Height          =   255
         Left            =   1680
         TabIndex        =   51
         Top             =   2160
         Width           =   495
      End
      Begin VB.CommandButton Command10 
         Caption         =   "+32"
         Height          =   255
         Left            =   3840
         TabIndex        =   50
         Top             =   2640
         Width           =   495
      End
      Begin VB.CommandButton Command9 
         Caption         =   "+32"
         Height          =   255
         Left            =   3840
         TabIndex        =   49
         Top             =   2160
         Width           =   495
      End
      Begin VB.TextBox TextDatos3 
         Height          =   285
         Index           =   8
         Left            =   2160
         TabIndex        =   46
         Top             =   2520
         Width           =   1695
      End
      Begin VB.TextBox TextDatos3 
         Height          =   285
         Index           =   7
         Left            =   2160
         TabIndex        =   45
         Top             =   2160
         Width           =   1695
      End
      Begin VB.TextBox TextDatos3 
         Enabled         =   0   'False
         Height          =   285
         Index           =   0
         Left            =   2160
         TabIndex        =   38
         Text            =   "1"
         Top             =   1680
         Width           =   615
      End
      Begin VB.TextBox TextDatos3 
         Enabled         =   0   'False
         Height          =   285
         Index           =   1
         Left            =   3120
         TabIndex        =   37
         Text            =   "1"
         Top             =   1680
         Width           =   735
      End
      Begin VB.TextBox TextDatos3 
         Enabled         =   0   'False
         Height          =   285
         Index           =   6
         Left            =   4080
         TabIndex        =   36
         Text            =   "1"
         Top             =   1680
         Width           =   735
      End
      Begin VB.TextBox TextDatos3 
         Height          =   285
         Index           =   3
         Left            =   2160
         TabIndex        =   35
         Top             =   3240
         Width           =   1695
      End
      Begin VB.TextBox TextDatos3 
         Height          =   285
         Index           =   2
         Left            =   2160
         TabIndex        =   34
         Top             =   2880
         Width           =   1695
      End
      Begin VB.CommandButton Command7 
         Caption         =   "?"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   33
         Top             =   480
         Width           =   255
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Puertas, entradas, adornos, arboles, graficos unicos"
         Height          =   255
         Index           =   1
         Left            =   480
         TabIndex        =   31
         Top             =   600
         Width           =   4455
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Pisos, costas, techos, paredes, muros(Texturas continuas)"
         Height          =   315
         Index           =   0
         Left            =   480
         TabIndex        =   30
         Top             =   240
         Value           =   -1  'True
         Width           =   4575
      End
      Begin VB.CommandButton Command6 
         Caption         =   "Crear indexacion"
         Height          =   495
         Left            =   1920
         TabIndex        =   6
         Top             =   4680
         Width           =   2415
      End
      Begin VB.CommandButton Command5 
         Caption         =   "buscar grh disponible"
         Height          =   255
         Left            =   3960
         TabIndex        =   5
         Top             =   3960
         Width           =   1815
      End
      Begin VB.TextBox TextDatos3 
         Height          =   285
         Index           =   5
         Left            =   2160
         TabIndex        =   3
         Top             =   3960
         Width           =   1695
      End
      Begin VB.TextBox TextDatos3 
         Height          =   285
         Index           =   4
         Left            =   2160
         TabIndex        =   2
         Top             =   3600
         Width           =   1695
      End
      Begin VB.Label Label11 
         Caption         =   "Capa:"
         Height          =   255
         Index           =   1
         Left            =   4800
         TabIndex        =   107
         Top             =   2160
         Width           =   975
      End
      Begin VB.Line Line4 
         BorderWidth     =   3
         X1              =   3720
         X2              =   4200
         Y1              =   3480
         Y2              =   3240
      End
      Begin VB.Line Line3 
         BorderWidth     =   3
         X1              =   3840
         X2              =   4320
         Y1              =   3000
         Y2              =   3120
      End
      Begin VB.Label Label17 
         Caption         =   $"FormAuto.frx":0000
         Height          =   855
         Left            =   120
         TabIndex        =   59
         Top             =   5640
         Width           =   6015
      End
      Begin VB.Label Label15 
         Caption         =   "Label15"
         Height          =   15
         Left            =   360
         TabIndex        =   58
         Top             =   5880
         Width           =   1095
      End
      Begin VB.Label Label14 
         Caption         =   "Nombre de la sup"
         Height          =   255
         Left            =   600
         TabIndex        =   56
         Top             =   960
         Width           =   1695
      End
      Begin VB.Label Label13 
         Caption         =   "Offset X:"
         Height          =   255
         Left            =   600
         TabIndex        =   48
         Top             =   2160
         Width           =   975
      End
      Begin VB.Label Label12 
         Caption         =   "Offset Y:"
         Height          =   255
         Left            =   600
         TabIndex        =   47
         Top             =   2520
         Width           =   975
      End
      Begin VB.Label Label11 
         Caption         =   "Ancho:"
         Height          =   255
         Index           =   0
         Left            =   600
         TabIndex        =   44
         Top             =   3240
         Width           =   975
      End
      Begin VB.Label Label10 
         Caption         =   "Alto:"
         Height          =   255
         Left            =   600
         TabIndex        =   43
         Top             =   2880
         Width           =   975
      End
      Begin VB.Label Ltext 
         Caption         =   "nºframes:"
         Height          =   255
         Index           =   10
         Left            =   600
         TabIndex        =   42
         Top             =   1680
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "A lo ancho"
         Height          =   255
         Index           =   2
         Left            =   2040
         TabIndex        =   41
         Top             =   1320
         Width           =   855
      End
      Begin VB.Label Label2 
         Caption         =   "A lo alto"
         Height          =   255
         Index           =   2
         Left            =   3120
         TabIndex        =   40
         Top             =   1320
         Width           =   855
      End
      Begin VB.Label Label3 
         Caption         =   "Totales"
         Height          =   255
         Index           =   2
         Left            =   4080
         TabIndex        =   39
         Top             =   1320
         Width           =   735
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         Caption         =   "Las superficies generalmente se separan en indexaciones de 128x128 para que a la hora de usarlas en el worldeditor sea mas facil."
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   120
         TabIndex        =   32
         Top             =   5160
         Width           =   6015
      End
      Begin VB.Label Label8 
         Caption         =   "Primer grh:"
         Height          =   255
         Left            =   600
         TabIndex        =   4
         Top             =   3960
         Width           =   975
      End
      Begin VB.Label Label7 
         Caption         =   "Bmp:"
         Height          =   255
         Left            =   600
         TabIndex        =   1
         Top             =   3600
         Width           =   735
      End
   End
   Begin VB.Frame FrameAnim 
      Caption         =   "Hechizos y fx's"
      Height          =   4695
      Index           =   0
      Left            =   0
      TabIndex        =   7
      Top             =   720
      Width           =   6255
      Begin VB.CommandButton Command3 
         Caption         =   "Indexr animacion en varios archivos(Ej: Hechizos, apocalipsis, etc)"
         Height          =   375
         Left            =   600
         TabIndex        =   100
         Top             =   240
         Width           =   5295
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Autoindexar como FX"
         Height          =   375
         Left            =   2040
         TabIndex        =   98
         Top             =   3120
         Width           =   2415
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Option1"
         Height          =   195
         Index           =   1
         Left            =   3000
         TabIndex        =   19
         Top             =   240
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Option1"
         Height          =   195
         Index           =   0
         Left            =   720
         TabIndex        =   18
         Top             =   240
         Value           =   -1  'True
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.CommandButton CommandCalu 
         Caption         =   "AutoCalcular con tamaño de imagen"
         Enabled         =   0   'False
         Height          =   495
         Left            =   3600
         TabIndex        =   17
         Top             =   1920
         Width           =   2415
      End
      Begin VB.CommandButton CommandBuscar 
         Caption         =   "Buscar grh's disponibles"
         Enabled         =   0   'False
         Height          =   255
         Left            =   3000
         TabIndex        =   16
         Top             =   1320
         Width           =   1815
      End
      Begin VB.TextBox TextDatos 
         Height          =   285
         Index           =   6
         Left            =   3840
         TabIndex        =   15
         Top             =   840
         Width           =   735
      End
      Begin VB.TextBox TextDatos 
         Height          =   285
         Index           =   1
         Left            =   2880
         TabIndex        =   14
         Text            =   "1"
         Top             =   840
         Width           =   735
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Crear"
         Height          =   495
         Left            =   960
         TabIndex        =   13
         Top             =   3600
         Width           =   4455
      End
      Begin VB.TextBox TextDatos 
         Height          =   285
         Index           =   4
         Left            =   1920
         TabIndex        =   12
         Top             =   2760
         Width           =   1095
      End
      Begin VB.TextBox TextDatos 
         Height          =   285
         Index           =   3
         Left            =   1920
         TabIndex        =   11
         Top             =   2280
         Width           =   1095
      End
      Begin VB.TextBox TextDatos 
         Height          =   285
         Index           =   2
         Left            =   1920
         TabIndex        =   10
         Top             =   1800
         Width           =   1095
      End
      Begin VB.TextBox TextDatos 
         Height          =   285
         Index           =   5
         Left            =   1920
         TabIndex        =   9
         Top             =   1320
         Width           =   1095
      End
      Begin VB.TextBox TextDatos 
         Height          =   285
         Index           =   0
         Left            =   1920
         TabIndex        =   8
         Text            =   "1"
         Top             =   840
         Width           =   615
      End
      Begin VB.Label Label4 
         Caption         =   "BMP consecutivos"
         Height          =   255
         Index           =   1
         Left            =   3360
         TabIndex        =   29
         Top             =   240
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.Label Label4 
         Caption         =   "Mismo BMP"
         Height          =   255
         Index           =   0
         Left            =   1080
         TabIndex        =   28
         Top             =   240
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.Line Line2 
         Index           =   0
         X1              =   3120
         X2              =   3480
         Y1              =   2400
         Y2              =   2160
      End
      Begin VB.Line Line1 
         Index           =   0
         X1              =   3120
         X2              =   3480
         Y1              =   1920
         Y2              =   2160
      End
      Begin VB.Label Label3 
         Caption         =   "Totales"
         Height          =   255
         Index           =   0
         Left            =   3840
         TabIndex        =   27
         Top             =   600
         Width           =   735
      End
      Begin VB.Label Label2 
         Caption         =   "A lo alto"
         Height          =   255
         Index           =   0
         Left            =   3000
         TabIndex        =   26
         Top             =   600
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "A lo ancho"
         Height          =   255
         Index           =   0
         Left            =   1800
         TabIndex        =   25
         Top             =   600
         Width           =   855
      End
      Begin VB.Label Ltext 
         Caption         =   "BMP:"
         Height          =   255
         Index           =   4
         Left            =   840
         TabIndex        =   24
         Top             =   2760
         Width           =   735
      End
      Begin VB.Label Ltext 
         Caption         =   "Ancho:"
         Height          =   255
         Index           =   3
         Left            =   840
         TabIndex        =   23
         Top             =   2280
         Width           =   735
      End
      Begin VB.Label Ltext 
         Caption         =   "Alto:"
         Height          =   255
         Index           =   2
         Left            =   840
         TabIndex        =   22
         Top             =   1800
         Width           =   735
      End
      Begin VB.Label Ltext 
         Caption         =   "Primer grh: "
         Height          =   255
         Index           =   1
         Left            =   840
         TabIndex        =   21
         Top             =   1320
         Width           =   1215
      End
      Begin VB.Label Ltext 
         Caption         =   "nºframes:"
         Height          =   255
         Index           =   0
         Left            =   840
         TabIndex        =   20
         Top             =   840
         Width           =   735
      End
   End
   Begin VB.CommandButton Command8 
      Caption         =   "?"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2160
      TabIndex        =   55
      Top             =   8760
      Width           =   255
   End
   Begin VB.OptionButton Option3 
      Caption         =   "Indexacion de imagen completa"
      Height          =   255
      Index           =   0
      Left            =   840
      TabIndex        =   54
      Top             =   10920
      Width           =   2655
   End
   Begin VB.OptionButton Option3 
      Caption         =   "Indexacion dentro de tileset"
      Height          =   495
      Index           =   1
      Left            =   840
      TabIndex        =   53
      Top             =   11040
      Value           =   -1  'True
      Width           =   3015
   End
   Begin VB.Frame FrameAnim 
      Caption         =   "Cuerpos, npcs y equipamientos"
      Height          =   5415
      Index           =   1
      Left            =   120
      TabIndex        =   60
      Top             =   720
      Width           =   6255
      Begin VB.OptionButton optionDimension 
         Caption         =   "Barca 2"
         Height          =   255
         Index           =   4
         Left            =   4560
         TabIndex        =   106
         Top             =   3360
         Width           =   1095
      End
      Begin VB.OptionButton optionDimension 
         Caption         =   "Barca 1"
         Height          =   255
         Index           =   3
         Left            =   4560
         TabIndex        =   105
         Top             =   3120
         Width           =   1095
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         ItemData        =   "FormAuto.frx":013F
         Left            =   600
         List            =   "FormAuto.frx":0149
         TabIndex        =   97
         Text            =   "Combo2"
         Top             =   4920
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.TextBox TextDatos2 
         Height          =   285
         Index           =   0
         Left            =   1920
         TabIndex        =   80
         Text            =   "6"
         Top             =   1560
         Width           =   615
      End
      Begin VB.TextBox TextDatos2 
         Height          =   285
         Index           =   5
         Left            =   1920
         TabIndex        =   79
         Top             =   2040
         Width           =   1095
      End
      Begin VB.TextBox TextDatos2 
         Height          =   285
         Index           =   2
         Left            =   1920
         TabIndex        =   78
         Top             =   2520
         Width           =   1095
      End
      Begin VB.TextBox TextDatos2 
         Height          =   285
         Index           =   3
         Left            =   1920
         TabIndex        =   77
         Top             =   3000
         Width           =   1095
      End
      Begin VB.TextBox TextDatos2 
         Height          =   285
         Index           =   4
         Left            =   1920
         TabIndex        =   76
         Top             =   3480
         Width           =   1095
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Crear indexacion"
         Height          =   495
         Left            =   1920
         TabIndex        =   75
         Top             =   4440
         Width           =   2295
      End
      Begin VB.TextBox TextDatos2 
         Height          =   285
         Index           =   1
         Left            =   2880
         TabIndex        =   74
         Text            =   "4"
         Top             =   1560
         Width           =   735
      End
      Begin VB.TextBox TextDatos2 
         Height          =   285
         Index           =   6
         Left            =   3840
         TabIndex        =   73
         Text            =   "22"
         Top             =   1560
         Width           =   735
      End
      Begin VB.CommandButton CommandBuscar2 
         Caption         =   "Buscar"
         Enabled         =   0   'False
         Height          =   255
         Left            =   3240
         TabIndex        =   72
         Top             =   2040
         Width           =   735
      End
      Begin VB.CommandButton CommandCalu2 
         Caption         =   "AutoCalcular"
         Enabled         =   0   'False
         Height          =   255
         Left            =   3360
         TabIndex        =   71
         Top             =   2880
         Width           =   1095
      End
      Begin VB.ComboBox ComboTipoAnim 
         Height          =   315
         ItemData        =   "FormAuto.frx":015C
         Left            =   840
         List            =   "FormAuto.frx":0166
         TabIndex        =   70
         Top             =   480
         Width           =   2295
      End
      Begin VB.TextBox TextDatos2 
         Height          =   285
         Index           =   7
         Left            =   3840
         TabIndex        =   69
         Top             =   4080
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.TextBox TextDatos2 
         Height          =   285
         Index           =   8
         Left            =   3840
         TabIndex        =   68
         Top             =   4080
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.OptionButton optionDimension 
         Caption         =   "Option2"
         Height          =   255
         Index           =   0
         Left            =   4560
         TabIndex        =   67
         Top             =   2520
         Width           =   255
      End
      Begin VB.OptionButton optionDimension 
         Caption         =   "Option3"
         Height          =   255
         Index           =   1
         Left            =   4560
         TabIndex        =   66
         Top             =   2880
         Width           =   255
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   2160
         TabIndex        =   65
         Text            =   "1"
         Top             =   4080
         Width           =   855
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   3840
         TabIndex        =   64
         Top             =   4080
         Width           =   735
      End
      Begin VB.CheckBox CheckAuto 
         Caption         =   "Check1"
         Height          =   255
         Left            =   480
         TabIndex        =   63
         Top             =   4080
         Width           =   255
      End
      Begin VB.ComboBox Combo2 
         Height          =   315
         ItemData        =   "FormAuto.frx":018C
         Left            =   600
         List            =   "FormAuto.frx":0199
         TabIndex        =   62
         Text            =   "Combo2"
         Top             =   4440
         Width           =   1215
      End
      Begin VB.OptionButton optionDimension 
         Caption         =   " (pj) alkon"
         Height          =   255
         Index           =   2
         Left            =   4560
         TabIndex        =   61
         Top             =   2200
         Width           =   1695
      End
      Begin VB.Label Ltext 
         Caption         =   "nºframes:"
         Height          =   255
         Index           =   5
         Left            =   840
         TabIndex        =   96
         Top             =   1560
         Width           =   735
      End
      Begin VB.Label Ltext 
         Caption         =   "Primer indice:"
         Height          =   255
         Index           =   6
         Left            =   840
         TabIndex        =   95
         Top             =   2040
         Width           =   1215
      End
      Begin VB.Label Ltext 
         Caption         =   "Alto:"
         Height          =   255
         Index           =   7
         Left            =   840
         TabIndex        =   94
         Top             =   2520
         Width           =   735
      End
      Begin VB.Label Ltext 
         Caption         =   "Ancho:"
         Height          =   255
         Index           =   8
         Left            =   840
         TabIndex        =   93
         Top             =   3000
         Width           =   735
      End
      Begin VB.Label Ltext 
         Caption         =   "BMP:"
         Height          =   255
         Index           =   9
         Left            =   840
         TabIndex        =   92
         Top             =   3480
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "A lo ancho"
         Height          =   255
         Index           =   1
         Left            =   1800
         TabIndex        =   91
         Top             =   1200
         Width           =   855
      End
      Begin VB.Label Label2 
         Caption         =   "A lo alto"
         Height          =   255
         Index           =   1
         Left            =   2880
         TabIndex        =   90
         Top             =   1200
         Width           =   855
      End
      Begin VB.Label Label3 
         Caption         =   "Totales"
         Height          =   255
         Index           =   1
         Left            =   3840
         TabIndex        =   89
         Top             =   1200
         Width           =   735
      End
      Begin VB.Line Line1 
         Index           =   1
         X1              =   3000
         X2              =   3360
         Y1              =   2760
         Y2              =   3000
      End
      Begin VB.Line Line2 
         Index           =   1
         X1              =   3000
         X2              =   3360
         Y1              =   3240
         Y2              =   3000
      End
      Begin VB.Label Loff 
         Caption         =   "Offset:"
         Height          =   255
         Left            =   840
         TabIndex        =   88
         Top             =   3840
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.Label Loffx 
         Caption         =   "x"
         Height          =   255
         Left            =   1560
         TabIndex        =   87
         Top             =   3840
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Label Loffy 
         Caption         =   "y"
         Height          =   255
         Left            =   3360
         TabIndex        =   86
         Top             =   3840
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Label Label5 
         Caption         =   "(pj) clasico"
         Height          =   255
         Left            =   4800
         TabIndex        =   85
         Top             =   2520
         Width           =   975
      End
      Begin VB.Label Label6 
         Caption         =   "Normal(NPC)"
         Height          =   255
         Left            =   4800
         TabIndex        =   84
         Top             =   2880
         Width           =   1095
      End
      Begin VB.Label Labelbody 
         Caption         =   "Autoindexar"
         Height          =   255
         Left            =   720
         TabIndex        =   83
         Top             =   4080
         Width           =   855
      End
      Begin VB.Label Labelbody1 
         Caption         =   "i:"
         Height          =   255
         Left            =   1920
         TabIndex        =   82
         Top             =   4080
         Width           =   135
      End
      Begin VB.Label Labelbody2 
         Caption         =   "offset:"
         Height          =   255
         Left            =   3240
         TabIndex        =   81
         Top             =   4080
         Width           =   615
      End
   End
End
Attribute VB_Name = "FormAuto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private PosicionNormales(1 To 22) As Position
Private PosicionNormales2(1 To 22) As Position
Dim tFramesTotales As Integer
Dim tFramesAncho As Integer, tFramesAlto As Integer
Dim tAlto As Long, tAncho As Long

Private Sub CheckAuto_Click()
    If CheckAuto.Value = vbChecked Then
        Text1.Enabled = True
        Text2.Enabled = True
    Else
        Text1.Enabled = False
        Text2.Enabled = False
    End If
End Sub



Private Sub ComboTipoAnim_Click()
        Combo2.Visible = False
        Labelbody.Visible = False
        Labelbody1.Visible = False
        Labelbody2.Visible = False
        Loff.Visible = False
        Loffx.Visible = False
        Loffy.Visible = False
        FormAuto.TextDatos2(7).Visible = False
        FormAuto.TextDatos2(8).Visible = False
        FormAuto.TextDatos2(0).Enabled = False
        FormAuto.TextDatos2(1).Enabled = False
        FormAuto.TextDatos2(6).Enabled = False
        Text1.Visible = False
        Text2.Visible = False
        CheckAuto.Visible = False
        optionDimension(0).Visible = False
        optionDimension(1).Visible = False
        optionDimension(2).Visible = False
        Label5.Visible = False
        Label6.Visible = False
        Combo1.Visible = False
        CommandCalu2.Visible = False
        Line1(1).Visible = False
        Line2(1).Visible = False
Select Case ComboTipoAnim.listIndex
    Case 0
        Line1(1).Visible = True
        Line2(1).Visible = True
        CommandCalu2.Visible = True
        FormAuto.TextDatos2(0).Text = 6
        FormAuto.TextDatos2(1).Text = 4
        FormAuto.TextDatos2(6).Text = 22
        FormAuto.TextDatos2(2).Text = 46
        FormAuto.TextDatos2(3).Text = 26
        Labelbody.Visible = True
        Labelbody1.Visible = True
        Labelbody2.Visible = True
        Text1.Visible = True
        Text1.Enabled = False
        Text2.Visible = True
        Text2.Enabled = False
        CheckAuto.Visible = True
        CheckAuto.Value = vbUnchecked
        Text1.Text = UBound(BodyData) + 1
        Text2.Text = "-38º0"
        Combo2.Visible = True
        Combo2.listIndex = 0
        optionDimension(0).Visible = True
        optionDimension(1).Visible = True
        optionDimension(2).Visible = True
        optionDimension(0).Value = True
        Label5.Visible = True
        Label6.Visible = True
        FormAuto.TextDatos2(2).Enabled = False
        FormAuto.TextDatos2(3).Enabled = False
        CommandCalu2.Enabled = True
    Case 1
        'Loff.Visible = True
        'Loffx.Visible = True
        'Loffy.Visible = True
        'FormAuto.TextDatos2(7).Visible = True
        'FormAuto.TextDatos2(8).Visible = True
        'FormAuto.TextDatos2(0).Enabled = True
        'FormAuto.TextDatos2(1).Enabled = True
        CommandCalu2.Visible = True
        FormAuto.TextDatos2(6).Enabled = False
        FormAuto.TextDatos2(2).Enabled = True
        FormAuto.TextDatos2(3).Enabled = True
        Combo1.Visible = True
        Combo1.listIndex = 0
        'FormAuto.CommandCalu2.Enabled = True
        TextDatos2(0).Text = "4"
        TextDatos2(1).Text = "1"
        TextDatos2(6).Text = "4"
        TextDatos2(2).Text = "16"
        TextDatos2(3).Text = "16"
        CommandBuscar2_Click
End Select
End Sub


Private Sub Command1_Click()
Dim FramesTotales As Integer
Dim FramesAncho As Integer, FramesAlto As Integer
Dim NumeroBMP As Integer, PrimerIndice As Long
Dim existenciaBMP As Integer
Dim Alto As Long, Ancho As Long
Dim AltoBMP As Long, AnchoBMP As Long
Dim BitCount As Integer
Dim ii As Integer
Dim FramesX As Long, FramesY As Long
Dim actualFrame As Integer
Dim curx As Long, cury As Long

On Error GoTo errh

For ii = 1 To 6
    If Val(FormAuto.TextDatos(ii).Text) <= 0 Then
        FormAuto.TextDatos(ii).Text = 0
    End If
Next ii

FramesTotales = Val(FormAuto.TextDatos(6).Text)
PrimerIndice = Val(FormAuto.TextDatos(5).Text)
NumeroBMP = Val(FormAuto.TextDatos(4).Text)
Alto = Val(FormAuto.TextDatos(2).Text)
Ancho = Val(FormAuto.TextDatos(3).Text)
FramesAncho = Val(FormAuto.TextDatos(0).Text)
FramesAlto = Val(FormAuto.TextDatos(1).Text)


If (Not hayGrHlibres(PrimerIndice, FramesTotales + 1)) Or PrimerIndice <= 0 Or PrimerIndice > MAXGrH Then
    MsgBox "No hay sitio para la animacion" & vbCrLf & "Sobreescribir x implementar"
Exit Sub
End If

actualFrame = 0
curx = 0
cury = 0
If Option1(0).Value Then
    ' Frames en el mismo BMP
    For FramesY = 1 To FramesAlto
        For FramesX = 1 To FramesAncho
            Grhdata(PrimerIndice + actualFrame).FileNum = NumeroBMP
            Grhdata(PrimerIndice + actualFrame).Frames(1) = PrimerIndice + actualFrame
            Grhdata(PrimerIndice + actualFrame).NumFrames = 1
            Grhdata(PrimerIndice + actualFrame).pixelHeight = Alto
            Grhdata(PrimerIndice + actualFrame).pixelWidth = Ancho
            Grhdata(PrimerIndice + actualFrame).sX = Ancho * curx
            Grhdata(PrimerIndice + actualFrame).sY = Alto * cury
            Grhdata(PrimerIndice + actualFrame).TileHeight = Grhdata(PrimerIndice + actualFrame).pixelHeight / TilePixelHeight
            Grhdata(PrimerIndice + actualFrame).TileWidth = Grhdata(PrimerIndice + actualFrame).pixelWidth / TilePixelWidth
            curx = curx + 1
            actualFrame = actualFrame + 1
            If actualFrame >= FramesTotales Then GoTo TerminarAnim
        Next FramesX
        curx = 0
        cury = cury + 1
    Next FramesY
    
Else
    For FramesY = 1 To FramesTotales
            Grhdata(PrimerIndice + actualFrame).FileNum = NumeroBMP + actualFrame
            Grhdata(PrimerIndice + actualFrame).Frames(1) = PrimerIndice + actualFrame
            Grhdata(PrimerIndice + actualFrame).NumFrames = 1
            Grhdata(PrimerIndice + actualFrame).pixelHeight = Alto
            Grhdata(PrimerIndice + actualFrame).pixelWidth = Ancho
            Grhdata(PrimerIndice + actualFrame).sX = 0
            Grhdata(PrimerIndice + actualFrame).sY = 0
            Grhdata(PrimerIndice + actualFrame).TileHeight = Grhdata(PrimerIndice + actualFrame).pixelHeight / TilePixelHeight
            Grhdata(PrimerIndice + actualFrame).TileWidth = Grhdata(PrimerIndice + actualFrame).pixelWidth / TilePixelWidth
            actualFrame = actualFrame + 1
            If actualFrame >= FramesTotales Then GoTo TerminarAnim
    Next FramesY
End If


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
            

For ii = 1 To FramesTotales
    Grhdata(PrimerIndice + FramesTotales).Frames(ii) = PrimerIndice + ii - 1
Next ii
Dim tS As String
tS = Round(FramesTotales / 2)
Grhdata(PrimerIndice + FramesTotales).Speed = CSng(tS & tS & tS & "." & tS & tS & tS)

If Check1.Value = 1 Then
    Dim nINdex As Integer
    nINdex = UBound(FxData) + 1
    Call AgregaFx(nINdex)
    FxData(nINdex).FX.GrhIndex = PrimerIndice + FramesTotales
    FxData(nINdex).FX.Speed = CSng(tS & tS & tS & "." & tS & tS & tS)
    EstadoNoGuardado(e_EstadoIndexador.FX) = True
End If

Call frmMain.CambiarEstado(e_EstadoIndexador.Grh)
Call frmMain.BuscarNuevoF(PrimerIndice)

Exit Sub
errh:
MsgBox "error: " & Err.Description & "  " & Err.Number

End Sub

Private Sub Command10_Click()
TextDatos3(8).Text = Val(TextDatos3(8).Text) + 32
End Sub

Private Sub Command11_Click()
If Val(TextDatos3(7).Text) >= 32 Then TextDatos3(7).Text = Val(TextDatos3(7).Text) - 32
End Sub

Private Sub Command12_Click()
If Val(TextDatos3(8).Text) >= 32 Then TextDatos3(8).Text = Val(TextDatos3(8).Text) - 32
End Sub



Private Sub Command2_Click()
' creacion de un grafico normal:

Dim FramesTotales As Integer
Dim FramesAncho As Integer, FramesAlto As Integer
Dim NumeroBMP As Integer, PrimerIndice As Long
Dim existenciaBMP As Integer
Dim Alto As Long, Ancho As Long
Dim AltoBMP As Long, AnchoBMP As Long
Dim BitCount As Integer
Dim ii As Integer
Dim FramesX As Long, FramesY As Long
Dim actualFrame As Integer
Dim curx As Long, cury As Long
Dim respuesta As Byte

On Error GoTo errh

'Comprobamos si hay datos invalidos:
For ii = 1 To 6
    If Val(FormAuto.TextDatos2(ii).Text) <= 0 Then
        FormAuto.TextDatos2(ii).Text = 0
    End If
Next ii

'Recogemos los valores necesarios
FramesTotales = Val(FormAuto.TextDatos2(6).Text)
PrimerIndice = Val(FormAuto.TextDatos2(5).Text)
NumeroBMP = Val(FormAuto.TextDatos2(4).Text)
Alto = Val(FormAuto.TextDatos2(2).Text)
Ancho = Val(FormAuto.TextDatos2(3).Text)
FramesAncho = Val(FormAuto.TextDatos2(0).Text)
FramesAlto = Val(FormAuto.TextDatos2(1).Text)

'comprobamos que hay hueco
If (Not hayGrHlibres(PrimerIndice, FramesTotales + 4)) Or PrimerIndice <= 0 Or PrimerIndice > MAXGrH Then
    MsgBox "No hay sitio para la animacion" & vbCrLf & "Sobreescribir x implementar"
    
    Exit Sub 'Realmente el implementar el sobreescribir seria solo quitar esta linea. Pero , sin la opcion de deshacer ahora mismo, no es muy recomendable
End If

If CheckAuto.Value = vbChecked Then
    ii = Val(Text1.Text)
    If errorEnIndice() Then
        respuesta = MsgBox("El grafico Indice de autoindexacion indicado ya existe, ¿estas seguro de sobreescribirlo?", 4, "Aviso")
    Else
        respuesta = vbYes
    End If
    If respuesta <> vbYes Then Exit Sub
End If
actualFrame = 0
curx = 0
cury = 0
If ComboTipoAnim.listIndex = 0 Then
    If optionDimension(0).Value Then
        ' Frames en el mismo BMP
        For FramesY = 1 To FramesTotales
                Grhdata(PrimerIndice + FramesY - 1).FileNum = NumeroBMP
                Grhdata(PrimerIndice + FramesY - 1).Frames(1) = FramesY
                Grhdata(PrimerIndice + FramesY - 1).NumFrames = 1
                Grhdata(PrimerIndice + FramesY - 1).pixelHeight = Alto
                Grhdata(PrimerIndice + FramesY - 1).pixelWidth = Ancho
                Grhdata(PrimerIndice + FramesY - 1).sX = PosicionNormales(FramesY).x
                Grhdata(PrimerIndice + FramesY - 1).sY = PosicionNormales(FramesY).y
                Grhdata(PrimerIndice + FramesY - 1).TileHeight = Grhdata(PrimerIndice + FramesY - 1).pixelHeight / TilePixelHeight
                Grhdata(PrimerIndice + FramesY - 1).TileWidth = Grhdata(PrimerIndice + FramesY - 1).pixelWidth / TilePixelWidth
        Next FramesY
        
    ElseIf optionDimension(1).Value Then
        For FramesY = 1 To 4
            For FramesX = 1 To FramesAncho
                Grhdata(PrimerIndice + actualFrame).FileNum = NumeroBMP
                Grhdata(PrimerIndice + actualFrame).Frames(1) = PrimerIndice + actualFrame
                Grhdata(PrimerIndice + actualFrame).NumFrames = 1
                Grhdata(PrimerIndice + actualFrame).pixelHeight = Alto
                Grhdata(PrimerIndice + actualFrame).pixelWidth = Ancho
                Grhdata(PrimerIndice + actualFrame).sX = Ancho * curx
                Grhdata(PrimerIndice + actualFrame).sY = Alto * cury
                Grhdata(PrimerIndice + actualFrame).TileHeight = Grhdata(PrimerIndice + actualFrame).pixelHeight / TilePixelHeight
                Grhdata(PrimerIndice + actualFrame).TileWidth = Grhdata(PrimerIndice + actualFrame).pixelWidth / TilePixelWidth
                curx = curx + 1
                actualFrame = actualFrame + 1
                If actualFrame >= FramesTotales Then GoTo TerminarAnim
            Next FramesX
            curx = 0
            cury = cury + 1
        Next FramesY
    ElseIf optionDimension(3).Value Then
        FramesAncho = 4
        For FramesY = 1 To 4
            For FramesX = 1 To 4
                
                Grhdata(PrimerIndice + actualFrame).FileNum = NumeroBMP
                Grhdata(PrimerIndice + actualFrame).Frames(1) = PrimerIndice + actualFrame
                Grhdata(PrimerIndice + actualFrame).NumFrames = 1
                
                If FramesY = 3 Or FramesY = 4 Then
                    Grhdata(PrimerIndice + actualFrame).pixelHeight = 94
                    Grhdata(PrimerIndice + actualFrame).pixelWidth = 66
                    Grhdata(PrimerIndice + actualFrame).sX = curx * 66
                    Grhdata(PrimerIndice + actualFrame).sY = 133 + ((FramesY - 3) * 94)
                Else
                    Grhdata(PrimerIndice + actualFrame).pixelHeight = 68
                    Grhdata(PrimerIndice + actualFrame).pixelWidth = 96
                    Grhdata(PrimerIndice + actualFrame).sX = 96 * curx
                    Grhdata(PrimerIndice + actualFrame).sY = 68 * cury
                End If
                Grhdata(PrimerIndice + actualFrame).TileHeight = Grhdata(PrimerIndice + actualFrame).pixelHeight / TilePixelHeight
                Grhdata(PrimerIndice + actualFrame).TileWidth = Grhdata(PrimerIndice + actualFrame).pixelWidth / TilePixelWidth
                curx = curx + 1
                actualFrame = actualFrame + 1
                If actualFrame >= FramesTotales Then GoTo TerminarAnim
            Next FramesX
            curx = 0
            cury = cury + 1
        Next FramesY
    ElseIf optionDimension(4).Value Then
        FramesAncho = 4
        For FramesY = 1 To 4
            For FramesX = 1 To 4
                
                Grhdata(PrimerIndice + actualFrame).FileNum = NumeroBMP
                Grhdata(PrimerIndice + actualFrame).Frames(1) = PrimerIndice + actualFrame
                Grhdata(PrimerIndice + actualFrame).NumFrames = 1
                
                If FramesY = 3 Or FramesY = 4 Then
                    Grhdata(PrimerIndice + actualFrame).pixelHeight = 96
                    Grhdata(PrimerIndice + actualFrame).pixelWidth = 66
                    Grhdata(PrimerIndice + actualFrame).sX = curx * 66
                    Grhdata(PrimerIndice + actualFrame).sY = 159 + ((FramesY - 3) * 96)
                Else
                    Grhdata(PrimerIndice + actualFrame).pixelHeight = 80
                    Grhdata(PrimerIndice + actualFrame).pixelWidth = 96
                    Grhdata(PrimerIndice + actualFrame).sX = 96 * curx
                    Grhdata(PrimerIndice + actualFrame).sY = 80 * cury
                End If
                Grhdata(PrimerIndice + actualFrame).TileHeight = Grhdata(PrimerIndice + actualFrame).pixelHeight / TilePixelHeight
                Grhdata(PrimerIndice + actualFrame).TileWidth = Grhdata(PrimerIndice + actualFrame).pixelWidth / TilePixelWidth
                curx = curx + 1
                actualFrame = actualFrame + 1
                If actualFrame >= FramesTotales Then GoTo TerminarAnim
            Next FramesX
            curx = 0
            cury = cury + 1
        Next FramesY
    Else
        For FramesY = 1 To FramesTotales
                Grhdata(PrimerIndice + FramesY - 1).FileNum = NumeroBMP
                Grhdata(PrimerIndice + FramesY - 1).Frames(1) = FramesY
                Grhdata(PrimerIndice + FramesY - 1).NumFrames = 1
                Grhdata(PrimerIndice + FramesY - 1).pixelHeight = Alto
                Grhdata(PrimerIndice + FramesY - 1).pixelWidth = Ancho
                Grhdata(PrimerIndice + FramesY - 1).sX = PosicionNormales2(FramesY).x
                Grhdata(PrimerIndice + FramesY - 1).sY = PosicionNormales2(FramesY).y
                Grhdata(PrimerIndice + FramesY - 1).TileHeight = Grhdata(PrimerIndice + FramesY - 1).pixelHeight / TilePixelHeight
                Grhdata(PrimerIndice + FramesY - 1).TileWidth = Grhdata(PrimerIndice + FramesY - 1).pixelWidth / TilePixelWidth
        Next FramesY
    End If
ElseIf ComboTipoAnim.listIndex = 1 Then 'cabeza o casco
    For FramesY = 1 To FramesAlto
        For FramesX = 1 To FramesAncho
            Grhdata(PrimerIndice + actualFrame).FileNum = NumeroBMP
            Grhdata(PrimerIndice + actualFrame).Frames(1) = PrimerIndice + actualFrame
            Grhdata(PrimerIndice + actualFrame).NumFrames = 1
            Grhdata(PrimerIndice + actualFrame).pixelHeight = Alto
            Grhdata(PrimerIndice + actualFrame).pixelWidth = Ancho
            Grhdata(PrimerIndice + actualFrame).sX = Ancho * curx + Val(TextDatos2(7).Text)
            Grhdata(PrimerIndice + actualFrame).sY = Alto * cury + Val(TextDatos2(8).Text)
            Grhdata(PrimerIndice + actualFrame).TileHeight = Grhdata(PrimerIndice + actualFrame).pixelHeight / TilePixelHeight
            Grhdata(PrimerIndice + actualFrame).TileWidth = Grhdata(PrimerIndice + actualFrame).pixelWidth / TilePixelWidth
            curx = curx + 1
            actualFrame = actualFrame + 1
            If actualFrame >= FramesTotales Then GoTo TerminarAnim
        Next FramesX
        curx = 0
        cury = cury + 1
    Next FramesY
End If

Dim tS As String

' No me gustan los Goto pero... es lo q hay xD
TerminarAnim:
 EstadoNoGuardado(e_EstadoIndexador.Grh) = True
If ComboTipoAnim.listIndex = 0 Then
    If optionDimension(0).Value Then 'indexacion clasica de bodys
        Grhdata(PrimerIndice + FramesTotales).FileNum = NumeroBMP
        Grhdata(PrimerIndice + FramesTotales).NumFrames = 6
        Grhdata(PrimerIndice + FramesTotales).Speed = 333.333
        Grhdata(PrimerIndice + FramesTotales).TileHeight = Alto / TilePixelHeight
        Grhdata(PrimerIndice + FramesTotales).TileWidth = Ancho / TilePixelWidth
        Grhdata(PrimerIndice + FramesTotales).pixelHeight = Alto
        Grhdata(PrimerIndice + FramesTotales).pixelWidth = Ancho
        For ii = 1 To 6
            Grhdata(PrimerIndice + FramesTotales).Frames(ii) = PrimerIndice + ii - 1
        Next ii
        Grhdata(PrimerIndice + FramesTotales + 1).FileNum = NumeroBMP
        Grhdata(PrimerIndice + FramesTotales + 1).NumFrames = 6
        Grhdata(PrimerIndice + FramesTotales + 1).Speed = 333.333
        Grhdata(PrimerIndice + FramesTotales + 1).TileHeight = Alto / TilePixelHeight
        Grhdata(PrimerIndice + FramesTotales + 1).TileWidth = Ancho / TilePixelWidth
        Grhdata(PrimerIndice + FramesTotales + 1).pixelHeight = Alto
        Grhdata(PrimerIndice + FramesTotales + 1).pixelWidth = Ancho
        For ii = 7 To 12
            Grhdata(PrimerIndice + FramesTotales + 1).Frames(ii - 6) = PrimerIndice + ii - 1
        Next ii
        Grhdata(PrimerIndice + FramesTotales + 2).FileNum = NumeroBMP
        Grhdata(PrimerIndice + FramesTotales + 2).NumFrames = 5
        Grhdata(PrimerIndice + FramesTotales + 2).Speed = 252.525
        Grhdata(PrimerIndice + FramesTotales + 2).TileHeight = Alto / TilePixelHeight
        Grhdata(PrimerIndice + FramesTotales + 2).TileWidth = Ancho / TilePixelWidth
        Grhdata(PrimerIndice + FramesTotales + 2).pixelHeight = Alto
        Grhdata(PrimerIndice + FramesTotales + 2).pixelWidth = Ancho
        For ii = 13 To 17
            Grhdata(PrimerIndice + FramesTotales + 2).Frames(ii - 12) = PrimerIndice + ii - 1
        Next ii
        Grhdata(PrimerIndice + FramesTotales + 3).FileNum = NumeroBMP
        Grhdata(PrimerIndice + FramesTotales + 3).NumFrames = 5
        Grhdata(PrimerIndice + FramesTotales + 3).Speed = 252.525
        Grhdata(PrimerIndice + FramesTotales + 3).TileWidth = Ancho / TilePixelWidth
                
        Grhdata(PrimerIndice + FramesTotales + 3).pixelHeight = Alto
        Grhdata(PrimerIndice + FramesTotales + 3).pixelWidth = Ancho
        For ii = 18 To 22
            Grhdata(PrimerIndice + FramesTotales + 3).Frames(ii - 17) = PrimerIndice + ii - 1
        Next ii
    ElseIf optionDimension(1).Value Then 'indexacion npc standar
        Grhdata(PrimerIndice + FramesTotales).FileNum = NumeroBMP
        Grhdata(PrimerIndice + FramesTotales).NumFrames = FramesAncho
        tS = Round(FramesAncho / 2)
        Grhdata(PrimerIndice + FramesTotales).Speed = CSng(tS & tS & tS & "." & tS & tS & tS)
        Grhdata(PrimerIndice + FramesTotales).TileHeight = Alto / TilePixelHeight
        Grhdata(PrimerIndice + FramesTotales).TileWidth = Ancho / TilePixelWidth
        Grhdata(PrimerIndice + FramesTotales).pixelHeight = Alto
        Grhdata(PrimerIndice + FramesTotales).pixelWidth = Ancho
        For ii = 1 To FramesAncho
            Grhdata(PrimerIndice + FramesTotales).Frames(ii) = PrimerIndice + ii - 1
        Next ii
        Grhdata(PrimerIndice + FramesTotales + 1).FileNum = NumeroBMP
        Grhdata(PrimerIndice + FramesTotales + 1).NumFrames = FramesAncho
        tS = Round(FramesAncho / 2)
        Grhdata(PrimerIndice + FramesTotales + 1).Speed = CSng(tS & tS & tS & "." & tS & tS & tS)
        Grhdata(PrimerIndice + FramesTotales + 1).TileHeight = Alto / TilePixelHeight
        Grhdata(PrimerIndice + FramesTotales + 1).TileWidth = Ancho / TilePixelWidth
        Grhdata(PrimerIndice + FramesTotales + 1).pixelHeight = Alto
        Grhdata(PrimerIndice + FramesTotales + 1).pixelWidth = Ancho
        For ii = FramesAncho + 1 To FramesAncho * 2
            Grhdata(PrimerIndice + FramesTotales + 1).Frames(ii - FramesAncho) = PrimerIndice + ii - 1
        Next ii
        Grhdata(PrimerIndice + FramesTotales + 2).FileNum = NumeroBMP
        Grhdata(PrimerIndice + FramesTotales + 2).NumFrames = FramesAncho
        tS = Round(FramesAncho / 2)
        Grhdata(PrimerIndice + FramesTotales + 2).Speed = CSng(tS & tS & tS & "." & tS & tS & tS)
        Grhdata(PrimerIndice + FramesTotales + 2).TileHeight = Alto / TilePixelHeight
        Grhdata(PrimerIndice + FramesTotales + 2).TileWidth = Ancho / TilePixelWidth
        Grhdata(PrimerIndice + FramesTotales + 2).pixelHeight = Alto
        Grhdata(PrimerIndice + FramesTotales + 2).pixelWidth = Ancho
        For ii = (FramesAncho * 2) + 1 To FramesAncho * 3
            Grhdata(PrimerIndice + FramesTotales + 2).Frames(ii - (FramesAncho * 2)) = PrimerIndice + ii - 1
        Next ii
        Grhdata(PrimerIndice + FramesTotales + 3).FileNum = NumeroBMP
        Grhdata(PrimerIndice + FramesTotales + 3).NumFrames = FramesAncho
        tS = Round(FramesAncho / 2)
        Grhdata(PrimerIndice + FramesTotales + 3).Speed = CSng(tS & tS & tS & "." & tS & tS & tS)
        Grhdata(PrimerIndice + FramesTotales + 3).TileWidth = Ancho / TilePixelWidth
                
        Grhdata(PrimerIndice + FramesTotales + 3).pixelHeight = Alto
        Grhdata(PrimerIndice + FramesTotales + 3).pixelWidth = Ancho
        For ii = (FramesAncho * 3) + 1 To FramesAncho * 4
            Grhdata(PrimerIndice + FramesTotales + 3).Frames(ii - (FramesAncho * 3)) = PrimerIndice + ii - 1
        Next ii
        
    ElseIf optionDimension(3).Value Then 'indexacion BARCA 1
        Grhdata(PrimerIndice + FramesTotales).FileNum = NumeroBMP
        Grhdata(PrimerIndice + FramesTotales).NumFrames = 4
        Grhdata(PrimerIndice + FramesTotales).Speed = 222.222
        Grhdata(PrimerIndice + FramesTotales).TileHeight = 68 / TilePixelHeight
        Grhdata(PrimerIndice + FramesTotales).TileWidth = 96 / TilePixelWidth
        Grhdata(PrimerIndice + FramesTotales).pixelHeight = 68
        Grhdata(PrimerIndice + FramesTotales).pixelWidth = 96
        For ii = 1 To FramesAncho
            Grhdata(PrimerIndice + FramesTotales).Frames(ii) = PrimerIndice + ii - 1
        Next ii
        Grhdata(PrimerIndice + FramesTotales + 1).FileNum = NumeroBMP
        Grhdata(PrimerIndice + FramesTotales + 1).NumFrames = 4
        Grhdata(PrimerIndice + FramesTotales + 1).Speed = 222.222
        Grhdata(PrimerIndice + FramesTotales + 1).TileHeight = 68 / TilePixelHeight
        Grhdata(PrimerIndice + FramesTotales + 1).TileWidth = 96 / TilePixelWidth
        Grhdata(PrimerIndice + FramesTotales + 1).pixelHeight = 68
        Grhdata(PrimerIndice + FramesTotales + 1).pixelWidth = 96
        For ii = FramesAncho + 1 To FramesAncho * 2
            Grhdata(PrimerIndice + FramesTotales + 1).Frames(ii - FramesAncho) = PrimerIndice + ii - 1
        Next ii
        Grhdata(PrimerIndice + FramesTotales + 2).FileNum = NumeroBMP
        Grhdata(PrimerIndice + FramesTotales + 2).NumFrames = 4
        Grhdata(PrimerIndice + FramesTotales + 2).Speed = 222.222
        Grhdata(PrimerIndice + FramesTotales + 2).TileHeight = 94 / TilePixelHeight
        Grhdata(PrimerIndice + FramesTotales + 2).TileWidth = 66 / TilePixelWidth
        Grhdata(PrimerIndice + FramesTotales + 2).pixelHeight = 94
        Grhdata(PrimerIndice + FramesTotales + 2).pixelWidth = 66
        For ii = (FramesAncho * 2) + 1 To FramesAncho * 3
            Grhdata(PrimerIndice + FramesTotales + 2).Frames(ii - (FramesAncho * 2)) = PrimerIndice + ii - 1
        Next ii
        Grhdata(PrimerIndice + FramesTotales + 3).FileNum = NumeroBMP
        Grhdata(PrimerIndice + FramesTotales + 3).NumFrames = 4
        Grhdata(PrimerIndice + FramesTotales + 3).Speed = 222.222
        Grhdata(PrimerIndice + FramesTotales + 3).TileWidth = 66 / TilePixelWidth
        Grhdata(PrimerIndice + FramesTotales + 3).TileHeight = 94 / TilePixelHeight
        Grhdata(PrimerIndice + FramesTotales + 3).pixelHeight = 94
        Grhdata(PrimerIndice + FramesTotales + 3).pixelWidth = 66
        For ii = (FramesAncho * 3) + 1 To FramesAncho * 4
            Grhdata(PrimerIndice + FramesTotales + 3).Frames(ii - (FramesAncho * 3)) = PrimerIndice + ii - 1
        Next ii
    ElseIf optionDimension(4).Value Then 'indexacion BARCA 2
        Grhdata(PrimerIndice + FramesTotales).FileNum = NumeroBMP
        Grhdata(PrimerIndice + FramesTotales).NumFrames = 4
        Grhdata(PrimerIndice + FramesTotales).Speed = 222.222
        Grhdata(PrimerIndice + FramesTotales).TileHeight = 80 / TilePixelHeight
        Grhdata(PrimerIndice + FramesTotales).TileWidth = 96 / TilePixelWidth
        Grhdata(PrimerIndice + FramesTotales).pixelHeight = 80
        Grhdata(PrimerIndice + FramesTotales).pixelWidth = 96
        For ii = 1 To FramesAncho
            Grhdata(PrimerIndice + FramesTotales).Frames(ii) = PrimerIndice + ii - 1
        Next ii
        Grhdata(PrimerIndice + FramesTotales + 1).FileNum = NumeroBMP
        Grhdata(PrimerIndice + FramesTotales + 1).NumFrames = 4
        Grhdata(PrimerIndice + FramesTotales + 1).Speed = 222.222
        Grhdata(PrimerIndice + FramesTotales + 1).TileHeight = 80 / TilePixelHeight
        Grhdata(PrimerIndice + FramesTotales + 1).TileWidth = 96 / TilePixelWidth
        Grhdata(PrimerIndice + FramesTotales + 1).pixelHeight = 80
        Grhdata(PrimerIndice + FramesTotales + 1).pixelWidth = 96
        For ii = FramesAncho + 1 To FramesAncho * 2
            Grhdata(PrimerIndice + FramesTotales + 1).Frames(ii - FramesAncho) = PrimerIndice + ii - 1
        Next ii
        Grhdata(PrimerIndice + FramesTotales + 2).FileNum = NumeroBMP
        Grhdata(PrimerIndice + FramesTotales + 2).NumFrames = 4
        Grhdata(PrimerIndice + FramesTotales + 2).Speed = 222.222
        Grhdata(PrimerIndice + FramesTotales + 2).TileHeight = 96 / TilePixelHeight
        Grhdata(PrimerIndice + FramesTotales + 2).TileWidth = 66 / TilePixelWidth
        Grhdata(PrimerIndice + FramesTotales + 2).pixelHeight = 96
        Grhdata(PrimerIndice + FramesTotales + 2).pixelWidth = 66
        For ii = (FramesAncho * 2) + 1 To FramesAncho * 3
            Grhdata(PrimerIndice + FramesTotales + 2).Frames(ii - (FramesAncho * 2)) = PrimerIndice + ii - 1
        Next ii
        Grhdata(PrimerIndice + FramesTotales + 3).FileNum = NumeroBMP
        Grhdata(PrimerIndice + FramesTotales + 3).NumFrames = 4
        Grhdata(PrimerIndice + FramesTotales + 3).Speed = 222.222
        Grhdata(PrimerIndice + FramesTotales + 3).TileWidth = 66 / TilePixelWidth
        Grhdata(PrimerIndice + FramesTotales + 3).TileHeight = 96 / TilePixelHeight
        Grhdata(PrimerIndice + FramesTotales + 3).pixelHeight = 96
        Grhdata(PrimerIndice + FramesTotales + 3).pixelWidth = 66
        For ii = (FramesAncho * 3) + 1 To FramesAncho * 4
            Grhdata(PrimerIndice + FramesTotales + 3).Frames(ii - (FramesAncho * 3)) = PrimerIndice + ii - 1
        Next ii
    Else 'indexacion ultimos bodys alkon
        Grhdata(PrimerIndice + FramesTotales).FileNum = NumeroBMP
        Grhdata(PrimerIndice + FramesTotales).NumFrames = 6
        Grhdata(PrimerIndice + FramesTotales).Speed = 333.333
        Grhdata(PrimerIndice + FramesTotales).TileHeight = Alto / TilePixelHeight
        Grhdata(PrimerIndice + FramesTotales).TileWidth = Ancho / TilePixelWidth
        Grhdata(PrimerIndice + FramesTotales).pixelHeight = Alto
        Grhdata(PrimerIndice + FramesTotales).pixelWidth = Ancho
       ' Grhdata(PrimerIndice + ActualFrame).pixelHeight / TilePixelHeight
        For ii = 1 To 6
            Grhdata(PrimerIndice + FramesTotales).Frames(ii) = PrimerIndice + ii - 1
        Next ii
        Grhdata(PrimerIndice + FramesTotales + 1).FileNum = NumeroBMP
        Grhdata(PrimerIndice + FramesTotales + 1).NumFrames = 6
        Grhdata(PrimerIndice + FramesTotales + 1).Speed = 333.333
        Grhdata(PrimerIndice + FramesTotales + 1).TileHeight = Alto / TilePixelHeight
        Grhdata(PrimerIndice + FramesTotales + 1).TileWidth = Ancho / TilePixelWidth
        Grhdata(PrimerIndice + FramesTotales + 1).pixelHeight = Alto
        Grhdata(PrimerIndice + FramesTotales + 1).pixelWidth = Ancho
        For ii = 7 To 12
        
            Grhdata(PrimerIndice + FramesTotales + 1).Frames(ii - 6) = PrimerIndice + ii - 1
        Next ii
        Grhdata(PrimerIndice + FramesTotales + 2).FileNum = NumeroBMP
        Grhdata(PrimerIndice + FramesTotales + 2).NumFrames = 5
        Grhdata(PrimerIndice + FramesTotales + 2).Speed = 252.525
        Grhdata(PrimerIndice + FramesTotales + 2).TileHeight = Alto / TilePixelHeight
        Grhdata(PrimerIndice + FramesTotales + 2).TileWidth = Ancho / TilePixelWidth
        Grhdata(PrimerIndice + FramesTotales + 2).pixelHeight = Alto
        Grhdata(PrimerIndice + FramesTotales + 2).pixelWidth = Ancho
        
        For ii = 13 To 17
            Grhdata(PrimerIndice + FramesTotales + 2).Frames(ii - 12) = PrimerIndice + ii - 1
        Next ii
        Grhdata(PrimerIndice + FramesTotales + 3).FileNum = NumeroBMP
        Grhdata(PrimerIndice + FramesTotales + 3).NumFrames = 5
        Grhdata(PrimerIndice + FramesTotales + 3).Speed = 252.525
        Grhdata(PrimerIndice + FramesTotales + 3).TileHeight = Alto / TilePixelHeight
        Grhdata(PrimerIndice + FramesTotales + 3).TileWidth = Ancho / TilePixelWidth
                
        Grhdata(PrimerIndice + FramesTotales + 3).pixelHeight = Alto
        Grhdata(PrimerIndice + FramesTotales + 3).pixelWidth = Ancho
        For ii = 18 To 22
            Grhdata(PrimerIndice + FramesTotales + 3).Frames(ii - 17) = PrimerIndice + ii - 1
        Next ii
    End If
ElseIf ComboTipoAnim.listIndex = 1 Then ' animacion con offset
    'Grhdata(PrimerIndice + FramesTotales).FileNum = NumeroBMP
    'Grhdata(PrimerIndice + FramesTotales).NumFrames = FramesTotales
    'Grhdata(PrimerIndice + FramesTotales).pixelHeight = Grhdata(PrimerIndice).pixelHeight
    'Grhdata(PrimerIndice + FramesTotales).pixelWidth = Grhdata(PrimerIndice).pixelWidth
    'Grhdata(PrimerIndice + FramesTotales).sX = Grhdata(PrimerIndice).sX
    'Grhdata(PrimerIndice + FramesTotales).sY = Grhdata(PrimerIndice).sY
    'Grhdata(PrimerIndice + FramesTotales).TileHeight = Grhdata(PrimerIndice).TileHeight
    'Grhdata(PrimerIndice + FramesTotales).TileWidth = Grhdata(PrimerIndice).TileWidth
End If

If ComboTipoAnim.listIndex = 0 Then 'ARMADURAS ESCUDOS Y ARMAS
    If CheckAuto.Value = vbChecked Then ' selecionado autoindexar como... n
    
        Select Case Combo2.listIndex
            Case 0 'body
                ii = UBound(BodyData) + 1
                Call AgregaBody(ii, False)
                BodyData(ii).HeadOffset.y = Val(ReadField(1, Text2.Text, Asc("º")))
                BodyData(ii).HeadOffset.x = Val(ReadField(2, Text2.Text, Asc("º")))
                BodyData(ii).Walk(1).GrhIndex = PrimerIndice + FramesTotales + 1
                BodyData(ii).Walk(2).GrhIndex = PrimerIndice + FramesTotales + 3
                BodyData(ii).Walk(3).GrhIndex = PrimerIndice + FramesTotales
                BodyData(ii).Walk(4).GrhIndex = PrimerIndice + FramesTotales + 2
                If optionDimension(3).Value Then
                    BodyData(ii).Walk(1).GrhIndex = PrimerIndice + FramesTotales + 3 ' + 1
                    BodyData(ii).Walk(2).GrhIndex = PrimerIndice + FramesTotales + 1 '+ 3
                    BodyData(ii).Walk(3).GrhIndex = PrimerIndice + FramesTotales + 2
                    BodyData(ii).Walk(4).GrhIndex = PrimerIndice + FramesTotales '+ 2
                ElseIf optionDimension(4).Value Then
                    BodyData(ii).Walk(1).GrhIndex = PrimerIndice + FramesTotales + 3 ' + 1
                    BodyData(ii).Walk(2).GrhIndex = PrimerIndice + FramesTotales + 1 '+ 3
                    BodyData(ii).Walk(3).GrhIndex = PrimerIndice + FramesTotales + 2
                    BodyData(ii).Walk(4).GrhIndex = PrimerIndice + FramesTotales '+ 2
                End If
                 EstadoNoGuardado(e_EstadoIndexador.Body) = True
                 Call frmMain.CambiarEstado(e_EstadoIndexador.Body)
                    Call frmMain.BuscarNuevoF(ii)
            Case 1 'arma
                ii = UBound(WeaponAnimData) + 1
                Call AgregaArma(ii, False)
                WeaponAnimData(ii).WeaponWalk(1).GrhIndex = PrimerIndice + FramesTotales + 1
                WeaponAnimData(ii).WeaponWalk(2).GrhIndex = PrimerIndice + FramesTotales + 3
                WeaponAnimData(ii).WeaponWalk(3).GrhIndex = PrimerIndice + FramesTotales
                WeaponAnimData(ii).WeaponWalk(4).GrhIndex = PrimerIndice + FramesTotales + 2
                 EstadoNoGuardado(e_EstadoIndexador.Armas) = True
                 Call frmMain.CambiarEstado(e_EstadoIndexador.Armas)
                    Call frmMain.BuscarNuevoF(ii)
            Case 2 'escudo
                ii = UBound(ShieldAnimData) + 1
                Call AgregaEscudo(ii, False)
                ShieldAnimData(ii).ShieldWalk(1).GrhIndex = PrimerIndice + FramesTotales + 1
                ShieldAnimData(ii).ShieldWalk(2).GrhIndex = PrimerIndice + FramesTotales + 3
                ShieldAnimData(ii).ShieldWalk(3).GrhIndex = PrimerIndice + FramesTotales
                ShieldAnimData(ii).ShieldWalk(4).GrhIndex = PrimerIndice + FramesTotales + 2
                 EstadoNoGuardado(e_EstadoIndexador.Escudos) = True
                 Call frmMain.CambiarEstado(e_EstadoIndexador.Escudos)
                Call frmMain.BuscarNuevoF(ii)
        End Select
    End If
End If
If ComboTipoAnim.listIndex = 1 Then 'head o casco
    'If CheckAuto.value = vbChecked Then
        If Combo1.listIndex = 0 Then  'Cabeza
            ii = UBound(HeadData) + 1
            Call AgregaCabeza(ii)
            HeadData(ii).Head(1).GrhIndex = PrimerIndice + 3
            HeadData(ii).Head(2).GrhIndex = PrimerIndice + 1
            HeadData(ii).Head(3).GrhIndex = PrimerIndice
            HeadData(ii).Head(4).GrhIndex = PrimerIndice + 2
             EstadoNoGuardado(e_EstadoIndexador.Cabezas) = True
             Call frmMain.CambiarEstado(e_EstadoIndexador.Cabezas)
            Call frmMain.BuscarNuevoF(ii)
        Else 'CASCOS
            If CascoAnimData(UBound(CascoAnimData)).Head(1).GrhIndex = 0 Then
                ii = UBound(CascoAnimData)
            Else
                ii = UBound(CascoAnimData) + 1
                Call AgregaCasco(ii)
            End If
            CascoAnimData(ii).Head(1).GrhIndex = PrimerIndice + 3
            CascoAnimData(ii).Head(2).GrhIndex = PrimerIndice + 1
            CascoAnimData(ii).Head(3).GrhIndex = PrimerIndice
            CascoAnimData(ii).Head(4).GrhIndex = PrimerIndice + 2
             EstadoNoGuardado(e_EstadoIndexador.Cascos) = True
             
             Call frmMain.CambiarEstado(e_EstadoIndexador.Cascos)
                Call frmMain.BuscarNuevoF(ii)
        End If
    'End If
End If
If CheckAuto.Value = vbUnchecked And ComboTipoAnim.listIndex = 0 Then
    Call frmMain.CambiarEstado(e_EstadoIndexador.Grh)
    Call frmMain.BuscarNuevoF(PrimerIndice)
End If


Exit Sub
errh:
MsgBox "error: " & Err.Description & "  " & Err.Number

End Sub

Private Sub Command3_Click()
'FormAuto.FrameAnim(0).Visible = True
'FormAuto.FrameAnim(1).Visible = False
'FormAuto.FrameAnim(2).Visible = False
frmSelectBMP.Show , frmMain
Unload Me
End Sub

Private Sub Command4_Click()
'Calcula automaticamente el ancho de los frames a partir de su numero
On Error GoTo errh
Dim FramesTotales As Integer
Dim FramesAncho As Integer, FramesAlto As Integer
Dim NumeroBMP As Integer, PrimerIndice As Long
Dim existenciaBMP As Integer
Dim Alto As Long, Ancho As Long
Dim AltoBMP As Long, AnchoBMP As Long
Dim BitCount As Integer
Dim ii As Integer

FramesTotales = Val(FormAuto.TextDatos3(6).Text)
PrimerIndice = Val(FormAuto.TextDatos2(5).Text)
NumeroBMP = Val(FormAuto.TextDatos3(4).Text)
Alto = Val(FormAuto.TextDatos3(2).Text)
Ancho = Val(FormAuto.TextDatos3(3).Text)
FramesAncho = Val(FormAuto.TextDatos3(0).Text)
FramesAlto = Val(FormAuto.TextDatos3(1).Text)

If FramesAncho < 1 Then Exit Sub
If FramesAlto < 1 Then Exit Sub

existenciaBMP = ExisteBMP(NumeroBMP)
If existenciaBMP > 0 Then
    Call GetTamañoBMP(NumeroBMP, AltoBMP, AnchoBMP, BitCount)
    FormAuto.TextDatos3(2).Text = CInt(AltoBMP / FramesAlto)
    FormAuto.TextDatos3(3).Text = CInt(AnchoBMP / FramesAncho)
End If

Exit Sub
errh:
MsgBox "error: " & Err.Description & "  " & Err.Number
End Sub


Private Sub Command7_Click()
MsgBox "Si es una textura, piso o superficies como costas, techos, pisos, paredes, etc. La indexacion es dividida en cuadrados de 32px x 32px para el correcto funcionamiento del worldeditor. cuando la imagen es para ser usada una unica vez dentro del mapa y no es una superficie continua(Entradas de dungeons, castillos, puertas, etc. El grafico es indexado en su tamaño original."
End Sub

Private Sub Command8_Click()
MsgBox "Cuando el grafico a ser indexado ocupa toda la imagen, entonces la indexacion se hace con el tamaño del BMP original. Cuando el grafico a ser indexado es solo una parte de la imagen, entonces las posiciones x e y dentro de la imagen tanto como el ancho y alto del grafico deben ser especificados."
End Sub

Private Sub Command9_Click()
TextDatos3(7).Text = Val(TextDatos3(7).Text) + 32
End Sub

Private Sub ISimple_Click()
FormAuto.FrameAnim(0).Visible = False
FormAuto.FrameAnim(1).Visible = False
FormAuto.FrameAnim(2).Visible = True
End Sub

Private Sub Command5_Click()


FormAuto.TextDatos3(5).Text = BuscarGrHlibres(IIf(Option2(0).Value = True, tFramesTotales, Val(TextDatos3(6).Text)))

End Sub


Private Sub crearsuperficie()


End Sub
Private Sub Command6_Click()
Dim FramesTotales As Integer
Dim FramesAncho As Integer, FramesAlto As Integer
Dim NumeroBMP As Integer, PrimerIndice As Long
Dim existenciaBMP As Integer
Dim Alto As Long, Ancho As Long
Dim AltoBMP As Long, AnchoBMP As Long
Dim BitCount As Integer
Dim ii As Integer
Dim FramesX As Long, FramesY As Long
Dim actualFrame As Long
Dim curx As Long, cury As Long
Dim offsetx As Integer, offsety As Integer
Dim primerasuperficie As Integer
'On Error GoTo errh

For ii = 0 To 6
    If ii <> 4 Then
        If Val(FormAuto.TextDatos3(ii).Text) <= 0 Then  'Val(FormAuto.TextDatos3(ii).Text) > 32000 Or
            FormAuto.TextDatos3(ii).Text = 0
        End If
    End If
Next ii


FramesTotales = Val(FormAuto.TextDatos3(6).Text)
PrimerIndice = Val(FormAuto.TextDatos3(5).Text)
NumeroBMP = Val(FormAuto.TextDatos3(4).Text)

offsetx = Val(FormAuto.TextDatos3(7).Text)
offsety = Val(FormAuto.TextDatos3(8).Text)


If Option2(0).Value = True Then 'Son texturas y pisos
    FramesAlto = tFramesAlto
    FramesAncho = tFramesAncho
    Ancho = tAncho
    Alto = tAlto
    FramesTotales = tFramesTotales
    If Check2.Value = vbChecked Then
        NumSuperficies = NumSuperficies + 1
        primerasuperficie = NumSuperficies
        ReDim Preserve SupData(0 To NumSuperficies)
    End If
Else 'Son decoraciones, entradas o arboles
    Alto = Val(FormAuto.TextDatos3(2).Text)
    Ancho = Val(FormAuto.TextDatos3(3).Text)
    FramesAncho = Val(FormAuto.TextDatos3(0).Text)
    FramesAlto = Val(FormAuto.TextDatos3(1).Text)
    If Check2.Value = vbChecked Then
        NumSuperficies = NumSuperficies + FramesTotales
        ReDim Preserve SupData(0 To NumSuperficies)
    End If
End If

If (Not hayGrHlibres(PrimerIndice, FramesTotales)) Or PrimerIndice <= 0 Or PrimerIndice > MAXGrH Then
    MsgBox "No hay sitio para la animacion" & vbCrLf & "Sobreescribir x implementar"
Exit Sub
End If

actualFrame = 0
curx = 0
cury = 0
'If Option1(0).value Then
    ' Frames en el mismo BMP

    For FramesY = 1 To FramesAlto
        For FramesX = 1 To FramesAncho
            Grhdata(PrimerIndice + actualFrame).FileNum = NumeroBMP
            Grhdata(PrimerIndice + actualFrame).Frames(1) = PrimerIndice + actualFrame
            Grhdata(PrimerIndice + actualFrame).NumFrames = 1
            Grhdata(PrimerIndice + actualFrame).pixelHeight = Alto
            Grhdata(PrimerIndice + actualFrame).pixelWidth = Ancho
            Grhdata(PrimerIndice + actualFrame).sX = (Ancho * curx) + offsetx
            Grhdata(PrimerIndice + actualFrame).sY = (Alto * cury) + offsety
            Grhdata(PrimerIndice + actualFrame).TileHeight = Grhdata(PrimerIndice + actualFrame).pixelHeight / TilePixelHeight
            Grhdata(PrimerIndice + actualFrame).TileWidth = Grhdata(PrimerIndice + actualFrame).pixelWidth / TilePixelWidth
            
            curx = curx + 1
            actualFrame = actualFrame + 1
            If Check2.Value = 1 Then
                If Option2(0).Value = False Then 'Son decoraciones o arboles
                    With SupData(NumSuperficies - FramesTotales + actualFrame)
                        .Nombre = txtNombre.Text & IIf(actualFrame = 0, "", " - " & actualFrame)
                        .Alto = 1
                        .Ancho = 1
                        .GrhIndex = PrimerIndice + actualFrame - 1
                    End With
                    Call guardarSuperficie(NumSuperficies - FramesTotales + actualFrame)
                End If
            End If
            
            If actualFrame >= FramesTotales Then
                If Check2.Value = 1 Then
                    If Option2(0).Value = True Then
                        With SupData(NumSuperficies)
                            .Nombre = txtNombre.Text
                            .Alto = FramesAlto
                            .Ancho = FramesAncho
                            .GrhIndex = PrimerIndice
                            .Capa = CByte(cmbCapa.List(cmbCapa.listIndex))
                            Call guardarSuperficie(NumSuperficies)
                        End With
                    End If
                End If
                GoTo TerminarAnim
            End If
        Next FramesX
        curx = 0
        cury = cury + 1
    Next FramesY
    

'Else
'    For FramesY = 1 To FramesTotales
'            Grhdata(PrimerIndice + ActualFrame).FileNum = NumeroBMP + ActualFrame
'            Grhdata(PrimerIndice + ActualFrame).Frames(1) = PrimerIndice + ActualFrame
'            Grhdata(PrimerIndice + ActualFrame).NumFrames = 1
'            Grhdata(PrimerIndice + ActualFrame).pixelHeight = Alto
'            Grhdata(PrimerIndice + ActualFrame).pixelWidth = Ancho
'            Grhdata(PrimerIndice + ActualFrame).sX = 0
'            Grhdata(PrimerIndice + ActualFrame).sY = 0
'            Grhdata(PrimerIndice + ActualFrame).TileHeight = Grhdata(PrimerIndice + ActualFrame).pixelHeight / TilePixelHeight
'            Grhdata(PrimerIndice + ActualFrame).TileWidth = Grhdata(PrimerIndice + ActualFrame).pixelWidth / TilePixelWidth
'            ActualFrame = ActualFrame + 1
'            If ActualFrame >= FramesTotales Then GoTo TerminarAnim
'    Next FramesY
'End If


TerminarAnim:

EstadoNoGuardado(e_EstadoIndexador.Grh) = True

If Check2.Value = 1 Then
    EstadoNoGuardado(e_EstadoIndexador.Superficies) = True
    MsgBox "Superficie indexada correctamente"
    
  '  Call frmMain.CambiarEstado(e_EstadoIndexador.Superficies)
 '''''  Call frmMain.BuscarNuevoF(primerasuperficie)
    
Else
  '  Call frmMain.CambiarEstado(e_EstadoIndexador.Grh)
  '  Call frmMain.BuscarNuevoF(PrimerIndice)
End If
Exit Sub
errh:
MsgBox "error: " & Err.Description & "  " & Err.Number
End Sub

Private Sub CommandBuscar2_Click()
    FormAuto.TextDatos2(5).Text = BuscarGrHlibres(Val(FormAuto.TextDatos2(6).Text) + 4)
End Sub

Private Sub CommandBuscar_Click()
    FormAuto.TextDatos(5).Text = BuscarGrHlibres(Val(FormAuto.TextDatos(6).Text) + 1)
End Sub

Private Sub CommandCalu_Click()
On Error GoTo errh
Dim FramesTotales As Integer
Dim FramesAncho As Integer, FramesAlto As Integer
Dim NumeroBMP As Integer, PrimerIndice As Long
Dim existenciaBMP As Integer
Dim Alto As Long, Ancho As Long
Dim AltoBMP As Long, AnchoBMP As Long
Dim BitCount As Integer
Dim ii As Integer

For ii = 1 To 6
    If Val(FormAuto.TextDatos(ii).Text) <= 0 Then
        FormAuto.TextDatos(ii).Text = 0
    End If
Next ii
If Val(FormAuto.TextDatos(6).Text) > 25 Then
    FormAuto.TextDatos(6).Text = 25
ElseIf Val(FormAuto.TextDatos(6).Text) <= 0 Then
    FormAuto.TextDatos(6).Text = 0
End If

FramesTotales = Val(FormAuto.TextDatos(6).Text)
PrimerIndice = Val(FormAuto.TextDatos(5).Text)
NumeroBMP = Val(FormAuto.TextDatos(4).Text)
Alto = Val(FormAuto.TextDatos(2).Text)
Ancho = Val(FormAuto.TextDatos(3).Text)
FramesAncho = Val(FormAuto.TextDatos(0).Text)
FramesAlto = Val(FormAuto.TextDatos(1).Text)

If FramesAncho < 1 Then Exit Sub
If FramesAlto < 1 Then Exit Sub

existenciaBMP = ExisteBMP(NumeroBMP)
If existenciaBMP > 0 Then
    Call GetTamañoBMP(NumeroBMP, AltoBMP, AnchoBMP, BitCount)
    FormAuto.TextDatos(2).Text = CInt(AltoBMP / FramesAlto)
    FormAuto.TextDatos(3).Text = CInt(AnchoBMP / FramesAncho)
End If
Exit Sub
errh:
MsgBox "error: " & Err.Description & "  " & Err.Number

End Sub

Private Sub CommandCalu2_Click()
'Calcula automaticamente el ancho de los frames a partir de su numero
On Error GoTo errh
Dim FramesTotales As Integer
Dim FramesAncho As Integer, FramesAlto As Integer
Dim NumeroBMP As Integer, PrimerIndice As Long
Dim existenciaBMP As Integer
Dim Alto As Long, Ancho As Long
Dim AltoBMP As Long, AnchoBMP As Long
Dim BitCount As Integer
Dim ii As Integer

For ii = 1 To 6
    If Val(FormAuto.TextDatos2(ii).Text) <= 0 Then
        FormAuto.TextDatos2(ii).Text = 0
    End If
Next ii
If Val(FormAuto.TextDatos2(6).Text) > 25 And ComboTipoAnim.listIndex > 1 Then
    FormAuto.TextDatos2(6).Text = 25
ElseIf Val(FormAuto.TextDatos2(6).Text) <= 0 Then
    FormAuto.TextDatos2(6).Text = 0
End If

FramesTotales = Val(FormAuto.TextDatos2(6).Text)
PrimerIndice = Val(FormAuto.TextDatos2(5).Text)
NumeroBMP = Val(FormAuto.TextDatos2(4).Text)
Alto = Val(FormAuto.TextDatos2(2).Text)
Ancho = Val(FormAuto.TextDatos2(3).Text)
FramesAncho = Val(FormAuto.TextDatos2(0).Text)
FramesAlto = Val(FormAuto.TextDatos2(1).Text)

If FramesAncho < 1 Then Exit Sub
If FramesAlto < 1 Then Exit Sub

existenciaBMP = ExisteBMP(NumeroBMP)
If existenciaBMP > 0 Then
    Call GetTamañoBMP(NumeroBMP, AltoBMP, AnchoBMP, BitCount)
    FormAuto.TextDatos2(2).Text = CInt(AltoBMP / FramesAlto)
    FormAuto.TextDatos2(3).Text = CInt(AnchoBMP / FramesAncho)
End If

Exit Sub
errh:
MsgBox "error: " & Err.Description & "  " & Err.Number
End Sub

Private Sub Form_Load()
Dim ii As Long
FormAuto.FrameAnim(0).Visible = True
FormAuto.FrameAnim(1).Visible = False
FormAuto.ComboTipoAnim.listIndex = 0
DibujarIndexaciones.activo = True
If EstadoIndexador <> e_EstadoIndexador.Resource Then Call frmMain.CambiarEstado(e_EstadoIndexador.Resource)

FormAuto.cmbCapa.AddItem "1"
FormAuto.cmbCapa.AddItem "2"
FormAuto.cmbCapa.AddItem "3"
FormAuto.cmbCapa.AddItem "4"
FormAuto.cmbCapa.listIndex = 0
For ii = 1 To 6
    PosicionNormales(ii).y = 0
Next ii

For ii = 7 To 12
    PosicionNormales(ii).y = 45
Next ii

For ii = 13 To 17
    PosicionNormales(ii).y = 90
Next ii

For ii = 18 To 22
    PosicionNormales(ii).y = 135
Next ii

PosicionNormales(1).x = 0
PosicionNormales(2).x = 25
PosicionNormales(3).x = 49
PosicionNormales(4).x = 73
PosicionNormales(5).x = 98
PosicionNormales(6).x = 123
PosicionNormales(7).x = 0
PosicionNormales(8).x = 25
PosicionNormales(9).x = 49
PosicionNormales(10).x = 73
PosicionNormales(11).x = 98
PosicionNormales(12).x = 123
PosicionNormales(13).x = 0
PosicionNormales(14).x = 25
PosicionNormales(15).x = 49
PosicionNormales(16).x = 73
PosicionNormales(17).x = 98
PosicionNormales(18).x = 0
PosicionNormales(19).x = 25
PosicionNormales(20).x = 49
PosicionNormales(21).x = 73
PosicionNormales(22).x = 98


For ii = 1 To 6
    PosicionNormales2(ii).y = 0
Next ii

For ii = 7 To 12
    PosicionNormales2(ii).y = 45
Next ii

For ii = 13 To 17
    PosicionNormales2(ii).y = 90
Next ii

For ii = 18 To 22
    PosicionNormales2(ii).y = 135
Next ii

For ii = 1 To 6
    PosicionNormales2(ii).x = (ii - 1) * 31
Next ii
For ii = 7 To 12
    PosicionNormales2(ii).x = (ii - 7) * 31
Next ii
For ii = 13 To 17
    PosicionNormales2(ii).x = (ii - 13) * 31
Next ii
For ii = 18 To 22
    PosicionNormales2(ii).x = (ii - 18) * 31
Next ii
'PosicionNormales2(1).X = 0
'PosicionNormales2(2).X = 31
'PosicionNormales2(3).X = 62
'PosicionNormales2(4).X = 93
'PosicionNormales2(5).X = 124
'PosicionNormales2(6).X = 155

'PosicionNormales2(7).X = 0
''PosicionNormales2(8).X = 30
'PosicionNormales2(9).X = 60
'PosicionNormales2(10).X = 90
'PosicionNormales2(11).X = 120
'PosicionNormales2(12).X = 150

'PosicionNormales2(13).X = 0
'PosicionNormales2(14).X = 25
'PosicionNormales2(15).X = 50
'PosicionNormales2(16).X = 75
'PosicionNormales2(17).X = 100

'PosicionNormales2(18).X = 0
'PosicionNormales2(19).X = 25
'PosicionNormales2(20).X = 50
'PosicionNormales2(21).X = 75
'PosicionNormales2(22).X = 100


TextDatos(6).Text = 1
TextDatos(2).Text = 16
TextDatos(3).Text = 16
TextDatos(4).Text = 1

'Text3.Text = "Nueva superficie"
End Sub



Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
DibujarIndexaciones.activo = False
End Sub


Private Sub Option1_Click(Index As Integer)
    If Index = 0 Then
       TextDatos(0).Enabled = True
       TextDatos(1).Enabled = True
    Else
        TextDatos(0).Enabled = False
        TextDatos(1).Enabled = False
        TextDatos(0).Text = 1
        TextDatos(1).Text = 1
    End If
End Sub

Private Sub Option2_Click(Index As Integer)
    If Option2(0).Value = True Then
        TextDatos3(0).Enabled = False
        TextDatos3(1).Enabled = False
        TextDatos3(6).Enabled = False
        TextDatos3(0).Text = "1"
        TextDatos3(1).Text = "1"
        TextDatos3(6).Text = "1"
        Call TextDatos3_Change(0)
        TextDatos(3).Text = "128"
        TextDatos(3).Text = "128"
    Else
        TextDatos3(0).Enabled = True
        TextDatos3(1).Enabled = True
        TextDatos3(6).Enabled = True
        Call TextDatos3_Change(0)
    End If
End Sub

Private Sub Optiondimension_Click(Index As Integer)
    
    Select Case Index
    Case 0
        CommandCalu2.Enabled = False
        FormAuto.TextDatos2(2).Enabled = False
        FormAuto.TextDatos2(3).Enabled = False
        TextDatos2(0).Enabled = False
        Combo2.Enabled = True
        Text2.Text = "-38º0"
        FormAuto.TextDatos2(0).Text = 6
        FormAuto.TextDatos2(1).Text = 4
        FormAuto.TextDatos2(6).Text = 22
        FormAuto.TextDatos2(2).Text = 46
        FormAuto.TextDatos2(3).Text = 26
    Case 1
        TextDatos2(0).Enabled = True
        FormAuto.TextDatos2(2).Enabled = True
        FormAuto.TextDatos2(3).Enabled = True
        Combo2.Enabled = False
        Combo2.listIndex = 0
        FormAuto.TextDatos2(1).Text = 0
        FormAuto.TextDatos2(1).Text = 4
        Text2.Text = "0º0"
    Case 2
        CommandCalu2.Enabled = False
        FormAuto.TextDatos2(2).Enabled = False
        FormAuto.TextDatos2(3).Enabled = False
        TextDatos2(0).Enabled = False
        Combo2.Enabled = True
        Text2.Text = "-38º0"
        FormAuto.TextDatos2(0).Text = 6
        FormAuto.TextDatos2(1).Text = 4
        FormAuto.TextDatos2(6).Text = 22
        FormAuto.TextDatos2(2).Text = 45
        FormAuto.TextDatos2(3).Text = 31
    End Select
    TextDatos2(0).Text = TextDatos2(0).Text
End Sub

Private Sub TabStrip1_Click()

End Sub

Private Sub Command13_Click(Index As Integer)
Dim ii As Long
    For ii = 0 To 2
        If ii = Index Then
            FrameAnim(ii).Visible = True
            Command13(ii).BackColor = &H8000000A
        Else
            FrameAnim(ii).Visible = False
            Command13(ii).BackColor = &H8000000F
        End If
    Next ii
End Sub

Public Sub loadTabStrip()
    Dim i As Long
    For i = 0 To 2
        If FrameAnim(i).Visible = True Then
            Command13(i).BackColor = &H8000000A
        Else
            Command13(i).BackColor = &H8000000F
        End If
    Next i

End Sub


Private Sub Text1_Change()
    On Error GoTo errh
    If Val(Text1.Text < 1) Then Text1.Text = 1
    If Val(Text1.Text > 32000) Then Text1.Text = 32000
    
    If errorEnIndice() Then
         Text1.BackColor = vbRed
    Else
        Text1.BackColor = vbWhite
    End If
    
    Exit Sub
errh:
MsgBox "error: " & Err.Description & "  " & Err.Number
End Sub

Private Function errorEnIndice() As Boolean
Select Case Combo2.listIndex
        Case 0 'body
            If Val(Text1.Text) <= UBound(BodyData) Then
                If BodyData(Val(Text1.Text)).Walk(1).GrhIndex > 0 Then
                    errorEnIndice = True
                End If
            End If
        Case 1 'arma
            If Val(Text1.Text) <= UBound(WeaponAnimData) Then
                If WeaponAnimData(Val(Text1.Text)).WeaponWalk(1).GrhIndex > 0 Then
                     errorEnIndice = True
                End If
            End If
        Case 2 'escudo
            If Val(Text1.Text) <= UBound(ShieldAnimData) Then
                If ShieldAnimData(Val(Text1.Text)).ShieldWalk(1).GrhIndex > 0 Then
                     errorEnIndice = True
                End If
            End If
        Case 3 'botas

    End Select
End Function
Private Sub Text2_Change()
Dim tempdouble1 As Double, tempdobule2 As Double

tempdouble1 = Val(ReadField(1, Text2.Text, Asc("º")))
tempdobule2 = Val(ReadField(2, Text2.Text, Asc("º")))

If tempdouble1 < -32000 Or tempdouble1 > 32000 Then
    Text2.Text = "0º" & tempdobule2
    tempdouble1 = 0
End If

If tempdobule2 < -32000 Or tempdobule2 > 32000 Then
    Text2.Text = tempdouble1 & "º0"
End If
        
End Sub

Private Sub TextDatos_Change(Index As Integer)
Dim FramesTotales As Integer
Dim FramesAncho As Integer, FramesAlto As Integer
Dim NumeroBMP As Integer, PrimerIndice As Long
Dim existenciaBMP As Integer
Dim Alto As Long, Ancho As Long
Dim AltoBMP As Long, AnchoBMP As Long
Dim BitCount As Integer
Dim ii As Integer
On Error GoTo errh
For ii = 1 To 6
    If Val(FormAuto.TextDatos(ii).Text) <= 0 Then
        FormAuto.TextDatos(ii).Text = 0
    End If
Next ii

If Index = 0 Or Index = 1 Then
    If Not (TextDatos(0) = 0 Or TextDatos(1) = 0) Then _
        TextDatos(6) = TextDatos(0) * TextDatos(1)
End If

If Val(FormAuto.TextDatos(6).Text) > 25 Then
    FormAuto.TextDatos(6).Text = 25
ElseIf Val(FormAuto.TextDatos(6).Text) <= 0 Then
    FormAuto.TextDatos(6).Text = 0
End If

FramesTotales = Val(FormAuto.TextDatos(6).Text)
PrimerIndice = Val(FormAuto.TextDatos(5).Text)
NumeroBMP = Val(FormAuto.TextDatos(4).Text)
Alto = Val(FormAuto.TextDatos(2).Text)
Ancho = Val(FormAuto.TextDatos(3).Text)
FramesAncho = Val(FormAuto.TextDatos(0).Text)
FramesAlto = Val(FormAuto.TextDatos(1).Text)

If FramesTotales > 0 Then
    FormAuto.CommandBuscar.Enabled = True
Else
    FormAuto.CommandBuscar.Enabled = False
End If

If Not hayGrHlibres(PrimerIndice, FramesTotales + 1) Then
    FormAuto.TextDatos(5).BackColor = vbRed
Else
    FormAuto.TextDatos(5).BackColor = vbWhite
End If

existenciaBMP = ExisteBMP(NumeroBMP)
If existenciaBMP > 0 Then
    FormAuto.TextDatos(4).BackColor = vbWhite
    CommandCalu.Enabled = True
    If EstadoIndexador <> e_EstadoIndexador.Resource Then Call frmMain.CambiarEstado(e_EstadoIndexador.Resource)
    Call CrearIndexacion(0)
    Call frmMain.BuscarNuevoF(NumeroBMP)
Else
    CommandCalu.Enabled = False
    FormAuto.TextDatos(4).BackColor = vbRed
    Exit Sub
End If

Call GetTamañoBMP(NumeroBMP, AltoBMP, AnchoBMP, BitCount)

If FramesAlto * Alto > AltoBMP Then
     FormAuto.TextDatos(2).BackColor = vbYellow
Else
     FormAuto.TextDatos(2).BackColor = vbWhite
End If

If FramesAncho * Ancho > AnchoBMP Then
     FormAuto.TextDatos(3).BackColor = vbYellow
Else
     FormAuto.TextDatos(3).BackColor = vbWhite
End If

Exit Sub
errh:
MsgBox "error: " & Err.Description & "  " & Err.Number
End Sub

Private Sub TextDatos2_Change(Index As Integer)
Dim FramesTotales As Integer
Dim FramesAncho As Integer, FramesAlto As Integer
Dim NumeroBMP As Integer, PrimerIndice As Long
Dim existenciaBMP As Integer
Dim Alto As Long, Ancho As Long
Dim AltoBMP As Long, AnchoBMP As Long
Dim BitCount As Integer
Dim ii As Integer
On Error GoTo errh
For ii = 1 To 6
    If Val(FormAuto.TextDatos2(ii).Text) <= 0 Then
        FormAuto.TextDatos2(ii).Text = 0
    End If
Next ii

If Val(FormAuto.TextDatos2(6).Text) > 25 And ComboTipoAnim.listIndex > 1 Then
    FormAuto.TextDatos2(6).Text = 25
ElseIf Val(FormAuto.TextDatos2(6).Text) <= 0 Then
    FormAuto.TextDatos2(6).Text = 0
End If


FramesTotales = Val(FormAuto.TextDatos2(6).Text)
PrimerIndice = Val(FormAuto.TextDatos2(5).Text)
NumeroBMP = Val(FormAuto.TextDatos2(4).Text)
Alto = Val(FormAuto.TextDatos2(2).Text)
Ancho = Val(FormAuto.TextDatos2(3).Text)
FramesAncho = Val(FormAuto.TextDatos2(0).Text)
FramesAlto = Val(FormAuto.TextDatos2(1).Text)

If ComboTipoAnim.listIndex = 0 And optionDimension(1).Value Then
    FormAuto.TextDatos2(6).Text = FramesAncho * 4
End If

If FramesTotales > 0 Then
    FormAuto.CommandBuscar2.Enabled = True
Else
    FormAuto.CommandBuscar2.Enabled = False
End If

If Not hayGrHlibres(PrimerIndice, FramesTotales + 4) Then
    FormAuto.TextDatos2(5).BackColor = vbRed
Else
    FormAuto.TextDatos2(5).BackColor = vbWhite
End If

existenciaBMP = ExisteBMP(NumeroBMP)
If existenciaBMP > 0 Then
    FormAuto.TextDatos2(4).BackColor = vbWhite
    CommandCalu2.Enabled = (optionDimension(1).Value) Or ComboTipoAnim.listIndex = 1
    If EstadoIndexador <> e_EstadoIndexador.Resource Then Call frmMain.CambiarEstado(e_EstadoIndexador.Resource)
    Call CrearIndexacion(1)
    Call frmMain.BuscarNuevoF(NumeroBMP)
Else
    CommandCalu2.Enabled = False
    FormAuto.TextDatos2(4).BackColor = vbRed
    Exit Sub
End If

Call GetTamañoBMP(NumeroBMP, AltoBMP, AnchoBMP, BitCount)

If FramesAlto * Alto > AltoBMP Then
     FormAuto.TextDatos2(2).BackColor = vbYellow
Else
     FormAuto.TextDatos2(2).BackColor = vbWhite
End If

If FramesAncho * Ancho > AnchoBMP Then
     FormAuto.TextDatos2(3).BackColor = vbYellow
Else
     FormAuto.TextDatos2(3).BackColor = vbWhite
End If
Exit Sub
errh:
MsgBox "error: " & Err.Description & "  " & Err.Number
End Sub
Public Sub CrearIndexacion(ByVal Index As Integer)
On Error Resume Next
Dim FramesY As Integer
Dim FramesX As Integer
Dim FramesTotales As Integer
Dim FramesAncho As Integer, FramesAlto As Integer
Dim NumeroBMP As Integer, Alto As Long, Ancho As Long
Dim actualFrame As Integer
Dim curx As Integer, cury As Integer
Dim BitCount As Integer





If Index = 0 Then
    FramesTotales = Val(FormAuto.TextDatos(6).Text)
    NumeroBMP = Val(FormAuto.TextDatos(4).Text)
    Alto = Val(FormAuto.TextDatos(2).Text)
    Ancho = Val(FormAuto.TextDatos(3).Text)
    FramesAncho = Val(FormAuto.TextDatos(0).Text)
    FramesAlto = Val(FormAuto.TextDatos(1).Text)
    DibujarIndexaciones.Ancho = Ancho
    DibujarIndexaciones.Alto = Alto
    DibujarIndexaciones.activo = True
    If Option1(0).Value Then
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
    Else
        DibujarIndexaciones.Inicios(1).x = 0
        DibujarIndexaciones.Inicios(1).y = 0
        DibujarIndexaciones.Total = 1
    End If
ElseIf Index = 1 Then
    FramesTotales = Val(FormAuto.TextDatos2(6).Text)
    NumeroBMP = Val(FormAuto.TextDatos2(4).Text)
    Alto = Val(FormAuto.TextDatos2(2).Text)
    Ancho = Val(FormAuto.TextDatos2(3).Text)
    FramesAncho = Val(FormAuto.TextDatos2(0).Text)
    FramesAlto = Val(FormAuto.TextDatos2(1).Text)
    DibujarIndexaciones.Ancho = Ancho
    DibujarIndexaciones.Alto = Alto
    
    If ComboTipoAnim.listIndex = 0 Then
        If optionDimension(0).Value Then
            ' Frames en el mismo BMP
            DibujarIndexaciones.activo = True
            For FramesY = 1 To FramesTotales
                DibujarIndexaciones.Inicios(FramesY).x = PosicionNormales(FramesY).x
                DibujarIndexaciones.Inicios(FramesY).y = PosicionNormales(FramesY).y
            Next FramesY
            DibujarIndexaciones.Total = FramesTotales
        ElseIf optionDimension(1).Value Then
            DibujarIndexaciones.activo = True
            For FramesY = 1 To 4
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
        Else
            DibujarIndexaciones.activo = True
            For FramesY = 1 To FramesTotales
                DibujarIndexaciones.Inicios(FramesY).x = PosicionNormales2(FramesY).x
                DibujarIndexaciones.Inicios(FramesY).y = PosicionNormales2(FramesY).y
                
            Next FramesY
            DibujarIndexaciones.Total = FramesTotales
        End If
    ElseIf ComboTipoAnim.listIndex = 1 Then
        DibujarIndexaciones.activo = True
        For FramesY = 1 To FramesAlto
            For FramesX = 1 To FramesAncho
                DibujarIndexaciones.Inicios(actualFrame + 1).x = Ancho * curx '+ Val(TextDatos2(7).Text)
                DibujarIndexaciones.Inicios(actualFrame + 1).y = Alto * cury '+ Val(TextDatos2(8).Text)
                curx = curx + 1
                actualFrame = actualFrame + 1
            Next FramesX
            curx = 0
            cury = cury + 1
        Next FramesY
        DibujarIndexaciones.Total = FramesTotales
    End If
ElseIf Index = 3 Then
    Dim offsetx As Integer, offsety As Integer
    FramesTotales = Val(FormAuto.TextDatos3(6).Text)
    NumeroBMP = Val(FormAuto.TextDatos3(4).Text)
    Alto = Val(FormAuto.TextDatos3(2).Text)
    Ancho = Val(FormAuto.TextDatos3(3).Text)
    FramesAncho = Val(FormAuto.TextDatos3(0).Text)
    FramesAlto = Val(FormAuto.TextDatos3(1).Text)
    offsetx = Val(FormAuto.TextDatos3(7).Text)
    offsety = Val(FormAuto.TextDatos3(8).Text)
    DibujarIndexaciones.Ancho = Ancho
    DibujarIndexaciones.Alto = Alto
    DibujarIndexaciones.activo = True
    'If Option1(0).value Then
        For FramesY = 1 To FramesAlto
            For FramesX = 1 To FramesAncho
                DibujarIndexaciones.Inicios(actualFrame + 1).x = (Ancho * curx) + offsetx
                DibujarIndexaciones.Inicios(actualFrame + 1).y = (Alto * cury) + offsety
                curx = curx + 1
                actualFrame = actualFrame + 1
            Next FramesX
            curx = 0
            cury = cury + 1
        Next FramesY
        DibujarIndexaciones.Total = FramesTotales
    'Else
   '     DibujarIndexaciones.Inicios(1).X = 0
    '    DibujarIndexaciones.Inicios(1).Y = 0
   '     DibujarIndexaciones.Total = 1
   ' End If
End If

End Sub
Private Sub TextDatos3_Change(Index As Integer)
Dim FramesTotales As Integer
Dim FramesAncho As Integer, FramesAlto As Integer
Dim NumeroBMP As Integer, PrimerIndice As Long
Dim existenciaBMP As Integer
Dim Alto As Long, Ancho As Long
Dim AltoBMP As Long, AnchoBMP As Long
Dim BitCount As Integer
Dim ii As Integer
'On Error GoTo errh


For ii = 1 To 6
    If Val(FormAuto.TextDatos3(ii).Text) <= 0 Then 'Val(FormAuto.TextDatos3(ii).Text) > 32000 Or
        FormAuto.TextDatos3(ii).Text = 0
    End If
Next ii

If Index = 0 Or Index = 1 Then
    If Not (TextDatos3(0) = 0 Or TextDatos3(1) = 0) Then _
        TextDatos3(6) = TextDatos3(0) * TextDatos3(1)
End If

If Val(FormAuto.TextDatos3(6).Text) > 60 Then
    FormAuto.TextDatos3(6).Text = 60
ElseIf Val(FormAuto.TextDatos3(6).Text) <= 0 Then
    FormAuto.TextDatos3(6).Text = 0
End If


PrimerIndice = Val(FormAuto.TextDatos3(5).Text)
NumeroBMP = Val(FormAuto.TextDatos3(4).Text) '
Alto = Val(FormAuto.TextDatos3(2).Text) '
Ancho = Val(FormAuto.TextDatos3(3).Text) '

If Option2(0).Value = False Then ' son graficos unicos
    FramesAncho = Val(FormAuto.TextDatos3(0).Text)
    FramesAlto = Val(FormAuto.TextDatos3(1).Text)
    FramesTotales = Val(FormAuto.TextDatos3(6).Text) '
Else 'Son pisos paredes o costas
    Select Case Alto
        Case 128
            FramesAlto = 128 / 32
            Alto = 32
            
        Case 256
            FramesAlto = 256 / 32
            Alto = 32
            
        Case 512
            FramesAlto = Alto / 32
            Alto = 32
        
        Case 384
            FramesAlto = Alto / 32
            Alto = 32
            
        Case Else
            FramesAlto = Alto / 32
            Alto = 32
    End Select
    
    Select Case Ancho
        Case 128
            FramesAncho = 128 / 32
            Ancho = 32
            
        Case 256
            FramesAncho = 256 / 32
            Ancho = 32
            
        Case 512
            FramesAncho = Ancho / 32
            Ancho = 32
            
        Case 384
            FramesAncho = Ancho / 32
            Ancho = 32
            
        Case Else
            FramesAncho = Ancho / 32
            Ancho = 32
    End Select
    FramesTotales = FramesAncho * FramesAlto
    'TextDatos(0).Text = FramesAncho
    'TextDatos(1).Text = FramesAlto
    'TextDatos(6).Text = FramesTotales
    tFramesAncho = FramesAncho
    tFramesAlto = FramesAlto
    tAncho = Alto
    tAlto = Ancho
    
End If
tFramesTotales = FramesTotales
If FramesTotales > 0 Then
    FormAuto.CommandBuscar.Enabled = True
Else
    FormAuto.CommandBuscar.Enabled = False
End If

If Not hayGrHlibres(PrimerIndice, FramesTotales) Then
    FormAuto.TextDatos3(5).BackColor = vbRed
Else
    FormAuto.TextDatos3(5).BackColor = vbWhite
End If

existenciaBMP = ExisteBMP(NumeroBMP)
If existenciaBMP > 0 Then
    FormAuto.TextDatos3(4).BackColor = vbWhite
    CommandCalu.Enabled = True
    If EstadoIndexador <> e_EstadoIndexador.Resource Then Call frmMain.CambiarEstado(e_EstadoIndexador.Resource)
    If Option2(0).Value = False Then
        Call CrearIndexacion(3)
    Else
            Dim offsetx As Integer, offsety As Integer, curx As Integer, cury As Integer, actualFrame As Integer, FramesY As Long, FramesX As Long
            NumeroBMP = Val(FormAuto.TextDatos3(4).Text)
            offsetx = Val(FormAuto.TextDatos3(7).Text)
            offsety = Val(FormAuto.TextDatos3(8).Text)
            DibujarIndexaciones.Ancho = Ancho
            DibujarIndexaciones.Alto = Alto
            DibujarIndexaciones.activo = True
            'If Option1(0).value Then
                For FramesY = 1 To FramesAlto
                    For FramesX = 1 To FramesAncho
                        DibujarIndexaciones.Inicios(actualFrame + 1).x = (Ancho * curx) + offsetx
                        DibujarIndexaciones.Inicios(actualFrame + 1).y = (Alto * cury) + offsety
                        curx = curx + 1
                        actualFrame = actualFrame + 1
                    Next FramesX
                    curx = 0
                    cury = cury + 1
                Next FramesY
                DibujarIndexaciones.Total = FramesTotales
    End If
    Call frmMain.BuscarNuevoF(NumeroBMP)
Else
    CommandCalu.Enabled = False
    FormAuto.TextDatos3(4).BackColor = vbRed
    Exit Sub
End If

Call GetTamañoBMP(NumeroBMP, AltoBMP, AnchoBMP, BitCount)

If Option2(0).Value = False Then
    If FramesAlto * Alto > AltoBMP Then
         FormAuto.TextDatos3(2).BackColor = vbYellow
    Else
         FormAuto.TextDatos3(2).BackColor = vbWhite
    End If
    
    If FramesAncho * Ancho > AnchoBMP Then
         FormAuto.TextDatos3(3).BackColor = vbYellow
    Else
         FormAuto.TextDatos3(3).BackColor = vbWhite
    End If
End If


Exit Sub
'errh:
'MsgBox "error: " & Err.Description & "  " & Err.Number
End Sub
Private Sub TextDatos33_Change(Index As Integer)
Dim NumeroBMP As Long
Dim PrimerIndice As Long
Dim ii As Long

For ii = 0 To 8
    If Val(FormAuto.TextDatos3(ii).Text) <= 0 Then
        FormAuto.TextDatos3(ii).Text = 0
    End If
Next ii

PrimerIndice = Val(FormAuto.TextDatos3(5).Text)
NumeroBMP = Val(FormAuto.TextDatos3(4).Text)

If Not hayGrHlibres(PrimerIndice, 1) Then
    FormAuto.TextDatos3(5).BackColor = vbRed
Else
    FormAuto.TextDatos3(5).BackColor = vbWhite
End If

If Index = 4 Then
    If ExisteBMP(NumeroBMP) > 0 Then
        If EstadoIndexador <> e_EstadoIndexador.Resource Then Call frmMain.CambiarEstado(e_EstadoIndexador.Resource)
        Call CrearIndexacion(3)
        Call frmMain.BuscarNuevoF(NumeroBMP)
        FormAuto.TextDatos3(4).BackColor = vbWhite
    Else
        FormAuto.TextDatos3(4).BackColor = vbRed
    End If
End If
    
End Sub
