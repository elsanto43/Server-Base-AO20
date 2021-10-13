VERSION 5.00
Begin VB.Form frmDats 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Dats"
   ClientHeight    =   11190
   ClientLeft      =   150
   ClientTop       =   495
   ClientWidth     =   15510
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   11190
   ScaleWidth      =   15510
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdLanzaSpells 
      Caption         =   "..."
      BeginProperty Font 
         Name            =   "Poor Richard"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   9000
      TabIndex        =   220
      Top             =   7800
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.ComboBox cmbTarget 
      Height          =   315
      Left            =   5040
      TabIndex        =   219
      Text            =   "Combo2"
      Top             =   5400
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.PictureBox picInv 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   510
      Left            =   3960
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   218
      Top             =   360
      Width           =   510
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   14160
      Top             =   960
   End
   Begin VB.CommandButton cmdRazasProhibidas 
      Caption         =   "Ver razas"
      BeginProperty Font 
         Name            =   "@Microsoft JhengHei"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   5280
      TabIndex        =   216
      Top             =   6240
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.CommandButton cmdClasesProhibidas 
      Caption         =   "Ver clases"
      BeginProperty Font 
         Name            =   "@Microsoft JhengHei"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   5280
      TabIndex        =   215
      Top             =   6600
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.ComboBox cmbSubtipo 
      Height          =   315
      Left            =   4080
      TabIndex        =   214
      Text            =   "Combo2"
      Top             =   4320
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.PictureBox picBuscar 
      Height          =   1815
      Left            =   120
      ScaleHeight     =   1755
      ScaleWidth      =   3315
      TabIndex        =   208
      Top             =   1200
      Visible         =   0   'False
      Width           =   3375
      Begin VB.CommandButton Command9 
         Caption         =   "Cancelar"
         Height          =   495
         Left            =   1680
         TabIndex        =   213
         Top             =   1080
         Width           =   1575
      End
      Begin VB.CommandButton Command8 
         Caption         =   "Buscar"
         Height          =   495
         Left            =   120
         TabIndex        =   212
         Top             =   1080
         Width           =   1575
      End
      Begin VB.TextBox txtBuscar 
         Height          =   285
         Left            =   120
         TabIndex        =   211
         Top             =   720
         Width           =   3135
      End
      Begin VB.OptionButton optObjtype 
         Caption         =   "Por tipo de objeto/npc/hechizo"
         Height          =   255
         Left            =   120
         TabIndex        =   210
         Top             =   360
         Width           =   2895
      End
      Begin VB.OptionButton optNombre 
         Caption         =   "Por nombre"
         Height          =   255
         Left            =   120
         TabIndex        =   209
         Top             =   120
         Value           =   -1  'True
         Width           =   1335
      End
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Buscar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   120
      TabIndex        =   207
      Top             =   600
      Width           =   3375
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Cancelar cambios"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   16800
      TabIndex        =   205
      Top             =   10080
      Visible         =   0   'False
      Width           =   3615
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Guardar OBJ.DAT"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   120
      TabIndex        =   204
      Top             =   10320
      Width           =   3375
   End
   Begin VB.CommandButton Command3 
      Caption         =   "..."
      BeginProperty Font 
         Name            =   "Poor Richard"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   9360
      TabIndex        =   203
      Top             =   8760
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.PictureBox picGrh 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   6255
      Left            =   7080
      ScaleHeight     =   415
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   543
      TabIndex        =   200
      Top             =   720
      Width           =   8175
   End
   Begin VB.ComboBox cmbObjType 
      Height          =   315
      Left            =   9240
      TabIndex        =   198
      Top             =   360
      Width           =   4095
   End
   Begin VB.PictureBox ProgressBar1 
      Height          =   615
      Left            =   9960
      ScaleHeight     =   555
      ScaleWidth      =   3675
      TabIndex        =   197
      Top             =   12240
      Width           =   3735
   End
   Begin VB.TextBox txtDatos 
      Height          =   285
      Index           =   95
      Left            =   7920
      TabIndex        =   172
      Text            =   "Text2"
      Top             =   19800
      Width           =   1815
   End
   Begin VB.TextBox txtDatos 
      Height          =   285
      Index           =   94
      Left            =   7920
      TabIndex        =   171
      Text            =   "Text2"
      Top             =   19440
      Width           =   1815
   End
   Begin VB.TextBox txtDatos 
      Height          =   285
      Index           =   93
      Left            =   7920
      TabIndex        =   170
      Text            =   "Text2"
      Top             =   19080
      Width           =   1815
   End
   Begin VB.TextBox txtDatos 
      Height          =   285
      Index           =   92
      Left            =   7920
      TabIndex        =   169
      Text            =   "Text2"
      Top             =   18720
      Width           =   1815
   End
   Begin VB.TextBox txtDatos 
      Height          =   285
      Index           =   91
      Left            =   7920
      TabIndex        =   168
      Text            =   "Text2"
      Top             =   18360
      Width           =   1815
   End
   Begin VB.TextBox txtDatos 
      Height          =   285
      Index           =   90
      Left            =   4440
      TabIndex        =   167
      Text            =   "Text2"
      Top             =   20520
      Width           =   1815
   End
   Begin VB.TextBox txtDatos 
      Height          =   285
      Index           =   89
      Left            =   4440
      TabIndex        =   166
      Text            =   "Text2"
      Top             =   20160
      Width           =   1815
   End
   Begin VB.TextBox txtDatos 
      Height          =   285
      Index           =   88
      Left            =   4440
      TabIndex        =   165
      Text            =   "Text2"
      Top             =   19800
      Width           =   1815
   End
   Begin VB.TextBox txtDatos 
      Height          =   285
      Index           =   87
      Left            =   4440
      TabIndex        =   164
      Text            =   "Text2"
      Top             =   19440
      Width           =   1815
   End
   Begin VB.TextBox txtDatos 
      Height          =   285
      Index           =   86
      Left            =   4440
      TabIndex        =   163
      Text            =   "Text2"
      Top             =   19080
      Width           =   1815
   End
   Begin VB.TextBox txtDatos 
      Height          =   285
      Index           =   85
      Left            =   4440
      TabIndex        =   162
      Text            =   "Text2"
      Top             =   18720
      Width           =   1815
   End
   Begin VB.TextBox txtDatos 
      Height          =   285
      Index           =   84
      Left            =   4440
      TabIndex        =   161
      Text            =   "Text2"
      Top             =   18360
      Width           =   1815
   End
   Begin VB.TextBox txtDatos 
      Height          =   285
      Index           =   83
      Left            =   960
      TabIndex        =   160
      Text            =   "Text2"
      Top             =   20160
      Width           =   1815
   End
   Begin VB.TextBox txtDatos 
      Height          =   285
      Index           =   82
      Left            =   960
      TabIndex        =   159
      Text            =   "Text2"
      Top             =   19800
      Width           =   1815
   End
   Begin VB.TextBox txtDatos 
      Height          =   285
      Index           =   81
      Left            =   960
      TabIndex        =   158
      Text            =   "Text2"
      Top             =   19440
      Width           =   1815
   End
   Begin VB.TextBox txtDatos 
      Height          =   285
      Index           =   80
      Left            =   960
      TabIndex        =   157
      Text            =   "Text2"
      Top             =   19080
      Width           =   1815
   End
   Begin VB.TextBox txtDatos 
      Height          =   285
      Index           =   79
      Left            =   960
      TabIndex        =   156
      Text            =   "Text2"
      Top             =   18720
      Width           =   1815
   End
   Begin VB.TextBox txtDatos 
      Height          =   285
      Index           =   78
      Left            =   960
      TabIndex        =   155
      Text            =   "Text2"
      Top             =   18360
      Width           =   1815
   End
   Begin VB.TextBox txtDatos 
      Height          =   285
      Index           =   77
      Left            =   4440
      TabIndex        =   154
      Text            =   "Text2"
      Top             =   20880
      Width           =   1815
   End
   Begin VB.TextBox txtDatos 
      Height          =   285
      Index           =   76
      Left            =   960
      TabIndex        =   153
      Text            =   "Text2"
      Top             =   20880
      Width           =   1815
   End
   Begin VB.TextBox txtDatos 
      Height          =   285
      Index           =   75
      Left            =   960
      TabIndex        =   152
      Text            =   "Text2"
      Top             =   20520
      Width           =   1815
   End
   Begin VB.TextBox txtDatos 
      Height          =   285
      Index           =   74
      Left            =   7920
      TabIndex        =   151
      Text            =   "Text2"
      Top             =   20880
      Width           =   1815
   End
   Begin VB.TextBox txtDatos 
      Height          =   285
      Index           =   73
      Left            =   7920
      TabIndex        =   150
      Text            =   "Text2"
      Top             =   20520
      Width           =   1815
   End
   Begin VB.TextBox txtDatos 
      Height          =   285
      Index           =   72
      Left            =   7920
      TabIndex        =   149
      Text            =   "Text2"
      Top             =   20160
      Width           =   1815
   End
   Begin VB.TextBox txtDatos 
      Height          =   285
      Index           =   71
      Left            =   960
      TabIndex        =   130
      Text            =   "Text2"
      Top             =   16200
      Width           =   1815
   End
   Begin VB.TextBox txtDatos 
      Height          =   285
      Index           =   70
      Left            =   960
      TabIndex        =   129
      Text            =   "Text2"
      Top             =   16560
      Width           =   1815
   End
   Begin VB.TextBox txtDatos 
      Height          =   285
      Index           =   69
      Left            =   960
      TabIndex        =   128
      Text            =   "Text2"
      Top             =   16920
      Width           =   1815
   End
   Begin VB.TextBox txtDatos 
      Height          =   285
      Index           =   68
      Left            =   960
      TabIndex        =   127
      Text            =   "Text2"
      Top             =   17280
      Width           =   1815
   End
   Begin VB.TextBox txtDatos 
      Height          =   285
      Index           =   67
      Left            =   960
      TabIndex        =   126
      Text            =   "Text2"
      Top             =   17640
      Width           =   1815
   End
   Begin VB.TextBox txtDatos 
      Height          =   285
      Index           =   66
      Left            =   960
      TabIndex        =   125
      Text            =   "Text2"
      Top             =   18000
      Width           =   1815
   End
   Begin VB.TextBox txtDatos 
      Height          =   285
      Index           =   65
      Left            =   7920
      TabIndex        =   124
      Text            =   "Text2"
      Top             =   16200
      Width           =   1815
   End
   Begin VB.TextBox txtDatos 
      Height          =   285
      Index           =   64
      Left            =   7920
      TabIndex        =   123
      Text            =   "Text2"
      Top             =   16560
      Width           =   1815
   End
   Begin VB.TextBox txtDatos 
      Height          =   285
      Index           =   63
      Left            =   7920
      TabIndex        =   122
      Text            =   "Text2"
      Top             =   16920
      Width           =   1815
   End
   Begin VB.TextBox txtDatos 
      Height          =   285
      Index           =   62
      Left            =   7920
      TabIndex        =   121
      Text            =   "Text2"
      Top             =   17280
      Width           =   1815
   End
   Begin VB.TextBox txtDatos 
      Height          =   285
      Index           =   61
      Left            =   4440
      TabIndex        =   120
      Text            =   "Text2"
      Top             =   16200
      Width           =   1815
   End
   Begin VB.TextBox txtDatos 
      Height          =   285
      Index           =   60
      Left            =   4440
      TabIndex        =   119
      Text            =   "Text2"
      Top             =   16560
      Width           =   1815
   End
   Begin VB.TextBox txtDatos 
      Height          =   285
      Index           =   59
      Left            =   4440
      TabIndex        =   118
      Text            =   "Text2"
      Top             =   16920
      Width           =   1815
   End
   Begin VB.TextBox txtDatos 
      Height          =   285
      Index           =   58
      Left            =   4440
      TabIndex        =   117
      Text            =   "Text2"
      Top             =   17280
      Width           =   1815
   End
   Begin VB.TextBox txtDatos 
      Height          =   285
      Index           =   57
      Left            =   4440
      TabIndex        =   116
      Text            =   "Text2"
      Top             =   17640
      Width           =   1815
   End
   Begin VB.TextBox txtDatos 
      Height          =   285
      Index           =   56
      Left            =   4440
      TabIndex        =   115
      Text            =   "Text2"
      Top             =   18000
      Width           =   1815
   End
   Begin VB.TextBox txtDatos 
      Height          =   285
      Index           =   55
      Left            =   7920
      TabIndex        =   114
      Text            =   "Text2"
      Top             =   17640
      Width           =   1815
   End
   Begin VB.TextBox txtDatos 
      Height          =   285
      Index           =   54
      Left            =   7920
      TabIndex        =   113
      Text            =   "Text2"
      Top             =   18000
      Width           =   1815
   End
   Begin VB.TextBox txtDatos 
      Height          =   285
      Index           =   53
      Left            =   9720
      TabIndex        =   94
      Text            =   "Text2"
      Top             =   13320
      Width           =   1815
   End
   Begin VB.TextBox txtDatos 
      Height          =   285
      Index           =   52
      Left            =   10200
      TabIndex        =   93
      Text            =   "Text2"
      Top             =   13200
      Width           =   1815
   End
   Begin VB.TextBox txtDatos 
      Height          =   285
      Index           =   51
      Left            =   9840
      TabIndex        =   92
      Text            =   "Text2"
      Top             =   12480
      Width           =   1815
   End
   Begin VB.TextBox txtDatos 
      Height          =   285
      Index           =   50
      Left            =   9360
      TabIndex        =   91
      Text            =   "Text2"
      Top             =   12480
      Width           =   1815
   End
   Begin VB.TextBox txtDatos 
      Height          =   285
      Index           =   49
      Left            =   9000
      TabIndex        =   90
      Text            =   "Text2"
      Top             =   12000
      Width           =   1815
   End
   Begin VB.TextBox txtDatos 
      Height          =   285
      Index           =   48
      Left            =   8040
      TabIndex        =   89
      Text            =   "Text2"
      Top             =   12240
      Width           =   1815
   End
   Begin VB.TextBox txtDatos 
      Height          =   285
      Index           =   47
      Left            =   7800
      TabIndex        =   88
      Text            =   "Text2"
      Top             =   13320
      Width           =   1815
   End
   Begin VB.TextBox txtDatos 
      Height          =   285
      Index           =   46
      Left            =   7800
      TabIndex        =   87
      Text            =   "Text2"
      Top             =   13680
      Width           =   1815
   End
   Begin VB.TextBox txtDatos 
      Height          =   285
      Index           =   45
      Left            =   7800
      TabIndex        =   86
      Text            =   "Text2"
      Top             =   14040
      Width           =   1815
   End
   Begin VB.TextBox txtDatos 
      Height          =   285
      Index           =   44
      Left            =   7800
      TabIndex        =   85
      Text            =   "Text2"
      Top             =   14400
      Width           =   1815
   End
   Begin VB.TextBox txtDatos 
      Height          =   285
      Index           =   43
      Left            =   7920
      TabIndex        =   84
      Text            =   "Text2"
      Top             =   13320
      Width           =   1815
   End
   Begin VB.TextBox txtDatos 
      Height          =   285
      Index           =   42
      Left            =   7920
      TabIndex        =   83
      Text            =   "Text2"
      Top             =   13680
      Width           =   1815
   End
   Begin VB.TextBox txtDatos 
      Height          =   285
      Index           =   41
      Left            =   7920
      TabIndex        =   82
      Text            =   "Text2"
      Top             =   14040
      Width           =   1815
   End
   Begin VB.TextBox txtDatos 
      Height          =   285
      Index           =   40
      Left            =   7920
      TabIndex        =   81
      Text            =   "Text2"
      Top             =   14400
      Width           =   1815
   End
   Begin VB.TextBox txtDatos 
      Height          =   285
      Index           =   39
      Left            =   7920
      TabIndex        =   80
      Text            =   "Text2"
      Top             =   14760
      Width           =   1815
   End
   Begin VB.TextBox txtDatos 
      Height          =   285
      Index           =   38
      Left            =   7920
      TabIndex        =   79
      Text            =   "Text2"
      Top             =   15120
      Width           =   1815
   End
   Begin VB.TextBox txtDatos 
      Height          =   285
      Index           =   37
      Left            =   7920
      TabIndex        =   78
      Text            =   "Text2"
      Top             =   15480
      Width           =   1815
   End
   Begin VB.TextBox txtDatos 
      Height          =   285
      Index           =   36
      Left            =   7920
      TabIndex        =   77
      Text            =   "Text2"
      Top             =   15840
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Guardar cambios"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   13200
      TabIndex        =   76
      Top             =   10080
      Visible         =   0   'False
      Width           =   3615
   End
   Begin VB.TextBox txtDatos 
      Height          =   285
      Index           =   35
      Left            =   4440
      TabIndex        =   75
      Text            =   "Text2"
      Top             =   15840
      Width           =   1815
   End
   Begin VB.TextBox txtDatos 
      Height          =   285
      Index           =   34
      Left            =   4440
      TabIndex        =   74
      Text            =   "Text2"
      Top             =   15480
      Width           =   1815
   End
   Begin VB.TextBox txtDatos 
      Height          =   285
      Index           =   33
      Left            =   4440
      TabIndex        =   73
      Text            =   "Text2"
      Top             =   15120
      Width           =   1815
   End
   Begin VB.TextBox txtDatos 
      Height          =   285
      Index           =   32
      Left            =   4440
      TabIndex        =   69
      Text            =   "Text2"
      Top             =   14760
      Width           =   1815
   End
   Begin VB.TextBox txtDatos 
      Height          =   285
      Index           =   31
      Left            =   4440
      TabIndex        =   68
      Text            =   "Text2"
      Top             =   14400
      Width           =   1815
   End
   Begin VB.TextBox txtDatos 
      Height          =   285
      Index           =   30
      Left            =   4440
      TabIndex        =   67
      Text            =   "Text2"
      Top             =   14040
      Width           =   1815
   End
   Begin VB.TextBox txtDatos 
      Height          =   285
      Index           =   29
      Left            =   4440
      TabIndex        =   63
      Text            =   "Text2"
      Top             =   13680
      Width           =   1815
   End
   Begin VB.TextBox txtDatos 
      Height          =   285
      Index           =   28
      Left            =   4560
      TabIndex        =   62
      Text            =   "Text2"
      Top             =   12600
      Width           =   1815
   End
   Begin VB.TextBox txtDatos 
      Height          =   285
      Index           =   27
      Left            =   4320
      TabIndex        =   61
      Text            =   "Text2"
      Top             =   14400
      Width           =   1815
   End
   Begin VB.TextBox txtDatos 
      Height          =   285
      Index           =   26
      Left            =   4320
      TabIndex        =   57
      Text            =   "Text2"
      Top             =   14040
      Width           =   1815
   End
   Begin VB.TextBox txtDatos 
      Height          =   285
      Index           =   25
      Left            =   4320
      TabIndex        =   56
      Text            =   "Text2"
      Top             =   13680
      Width           =   1815
   End
   Begin VB.TextBox txtDatos 
      Height          =   285
      Index           =   24
      Left            =   4320
      TabIndex        =   55
      Text            =   "Text2"
      Top             =   12720
      Width           =   1815
   End
   Begin VB.TextBox txtDatos 
      Height          =   285
      Index           =   23
      Left            =   3840
      TabIndex        =   51
      Text            =   "Text2"
      Top             =   12120
      Width           =   1815
   End
   Begin VB.TextBox txtDatos 
      Height          =   285
      Index           =   22
      Left            =   5400
      TabIndex        =   50
      Text            =   "Text2"
      Top             =   11280
      Width           =   1815
   End
   Begin VB.TextBox txtDatos 
      Height          =   285
      Index           =   21
      Left            =   11280
      MultiLine       =   -1  'True
      TabIndex        =   49
      Text            =   "frmDats.frx":0000
      Top             =   9000
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.TextBox txtDatos 
      Height          =   285
      Index           =   20
      Left            =   4200
      TabIndex        =   45
      Text            =   "Text2"
      Top             =   12360
      Width           =   1815
   End
   Begin VB.TextBox txtDatos 
      Height          =   285
      Index           =   19
      Left            =   4440
      TabIndex        =   44
      Text            =   "Text2"
      Top             =   12240
      Width           =   1815
   End
   Begin VB.TextBox txtDatos 
      Height          =   285
      Index           =   18
      Left            =   4200
      TabIndex        =   43
      Text            =   "Text2"
      Top             =   12120
      Width           =   1815
   End
   Begin VB.TextBox txtDatos 
      Height          =   285
      Index           =   17
      Left            =   960
      TabIndex        =   39
      Text            =   "Text2"
      Top             =   15840
      Width           =   1815
   End
   Begin VB.TextBox txtDatos 
      Height          =   285
      Index           =   16
      Left            =   960
      TabIndex        =   38
      Text            =   "Text2"
      Top             =   15480
      Width           =   1815
   End
   Begin VB.TextBox txtDatos 
      Height          =   285
      Index           =   15
      Left            =   1080
      TabIndex        =   37
      Text            =   "Text2"
      Top             =   14280
      Width           =   1815
   End
   Begin VB.TextBox txtDatos 
      Height          =   285
      Index           =   14
      Left            =   960
      TabIndex        =   33
      Text            =   "Text2"
      Top             =   14760
      Width           =   1815
   End
   Begin VB.TextBox txtDatos 
      Height          =   285
      Index           =   13
      Left            =   960
      TabIndex        =   32
      Text            =   "Text2"
      Top             =   14400
      Width           =   1815
   End
   Begin VB.TextBox txtDatos 
      Height          =   285
      Index           =   12
      Left            =   960
      TabIndex        =   31
      Text            =   "Text2"
      Top             =   14040
      Width           =   1815
   End
   Begin VB.TextBox txtDatos 
      Height          =   285
      Index           =   11
      Left            =   960
      TabIndex        =   27
      Text            =   "Text2"
      Top             =   13680
      Width           =   1815
   End
   Begin VB.TextBox txtDatos 
      Height          =   285
      Index           =   10
      Left            =   960
      TabIndex        =   26
      Text            =   "Text2"
      Top             =   13320
      Width           =   1815
   End
   Begin VB.TextBox txtDatos 
      Height          =   285
      Index           =   9
      Left            =   840
      TabIndex        =   25
      Text            =   "Text2"
      Top             =   14400
      Width           =   1815
   End
   Begin VB.TextBox txtDatos 
      Height          =   285
      Index           =   8
      Left            =   840
      TabIndex        =   21
      Text            =   "Text2"
      Top             =   14040
      Width           =   1815
   End
   Begin VB.TextBox txtDatos 
      Height          =   285
      Index           =   7
      Left            =   840
      TabIndex        =   20
      Text            =   "Text2"
      Top             =   13680
      Width           =   1815
   End
   Begin VB.TextBox txtDatos 
      Height          =   285
      Index           =   6
      Left            =   840
      TabIndex        =   19
      Text            =   "Text2"
      Top             =   13320
      Width           =   1815
   End
   Begin VB.TextBox txtDatos 
      Height          =   285
      Index           =   5
      Left            =   840
      TabIndex        =   15
      Text            =   "Text2"
      Top             =   12960
      Width           =   1815
   End
   Begin VB.TextBox txtDatos 
      Height          =   285
      Index           =   4
      Left            =   840
      MultiLine       =   -1  'True
      TabIndex        =   14
      Text            =   "frmDats.frx":0006
      Top             =   12600
      Width           =   1815
   End
   Begin VB.TextBox txtDatos 
      Height          =   285
      Index           =   3
      Left            =   960
      MultiLine       =   -1  'True
      TabIndex        =   13
      Text            =   "frmDats.frx":000C
      Top             =   11760
      Width           =   1815
   End
   Begin VB.TextBox txtDatos 
      Height          =   285
      Index           =   2
      Left            =   7200
      MultiLine       =   -1  'True
      TabIndex        =   9
      Text            =   "frmDats.frx":0012
      Top             =   11760
      Width           =   1815
   End
   Begin VB.TextBox txtDatos 
      Height          =   285
      Index           =   1
      Left            =   720
      MultiLine       =   -1  'True
      TabIndex        =   8
      Text            =   "frmDats.frx":0018
      Top             =   12480
      Width           =   1815
   End
   Begin VB.TextBox txtDatos 
      Height          =   285
      Index           =   0
      Left            =   1440
      MultiLine       =   -1  'True
      TabIndex        =   7
      Text            =   "frmDats.frx":001E
      Top             =   11520
      Width           =   1815
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   4920
      TabIndex        =   6
      Top             =   360
      Width           =   4215
   End
   Begin VB.ListBox List1 
      Height          =   8250
      Left            =   120
      TabIndex        =   1
      Top             =   1320
      Width           =   3375
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "frmDats.frx":0024
      Left            =   120
      List            =   "frmDats.frx":0026
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   240
      Width           =   3375
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Borrar objeto"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   11400
      TabIndex        =   206
      Top             =   7680
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Nuevo objeto"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   120
      TabIndex        =   202
      Top             =   9720
      Width           =   3375
   End
   Begin VB.Label lblEstado 
      Caption         =   "Estado: Guardado"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   495
      Left            =   3720
      TabIndex        =   217
      Top             =   10680
      Width           =   6495
   End
   Begin VB.Label Label3 
      Caption         =   "Clickea en una propiedad para saber mas sobre ella"
      Height          =   615
      Left            =   840
      TabIndex        =   201
      Top             =   12480
      Width           =   3375
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Tipo de objeto"
      Height          =   255
      Left            =   9360
      TabIndex        =   199
      Top             =   120
      Width           =   3855
   End
   Begin VB.Label Label1 
      Caption         =   "Label2"
      Height          =   255
      Index           =   96
      Left            =   6480
      TabIndex        =   196
      Top             =   19800
      Width           =   2055
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   255
      Index           =   95
      Left            =   6480
      TabIndex        =   195
      Top             =   19440
      Width           =   1935
   End
   Begin VB.Label Label1 
      Caption         =   "Label2"
      Height          =   255
      Index           =   94
      Left            =   6480
      TabIndex        =   194
      Top             =   19080
      Width           =   1815
   End
   Begin VB.Label Label1 
      Caption         =   "Label2"
      Height          =   255
      Index           =   93
      Left            =   6480
      TabIndex        =   193
      Top             =   18720
      Width           =   1935
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   255
      Index           =   92
      Left            =   6480
      TabIndex        =   192
      Top             =   18360
      Width           =   1935
   End
   Begin VB.Label Label1 
      Caption         =   "Label2"
      Height          =   255
      Index           =   91
      Left            =   3000
      TabIndex        =   191
      Top             =   20520
      Width           =   2055
   End
   Begin VB.Label Label1 
      Caption         =   "Label2"
      Height          =   255
      Index           =   90
      Left            =   3000
      TabIndex        =   190
      Top             =   20160
      Width           =   2055
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   255
      Index           =   89
      Left            =   3000
      TabIndex        =   189
      Top             =   19800
      Width           =   1935
   End
   Begin VB.Label Label1 
      Caption         =   "Label2"
      Height          =   255
      Index           =   88
      Left            =   3000
      TabIndex        =   188
      Top             =   19440
      Width           =   1815
   End
   Begin VB.Label Label1 
      Caption         =   "Label2"
      Height          =   255
      Index           =   87
      Left            =   3000
      TabIndex        =   187
      Top             =   19080
      Width           =   1935
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   255
      Index           =   86
      Left            =   3000
      TabIndex        =   186
      Top             =   18720
      Width           =   1935
   End
   Begin VB.Label Label1 
      Caption         =   "Label2"
      Height          =   255
      Index           =   85
      Left            =   3000
      TabIndex        =   185
      Top             =   18360
      Width           =   2055
   End
   Begin VB.Label Label1 
      Caption         =   "Label2"
      Height          =   255
      Index           =   84
      Left            =   -480
      TabIndex        =   184
      Top             =   20160
      Width           =   2055
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   255
      Index           =   83
      Left            =   -480
      TabIndex        =   183
      Top             =   19800
      Width           =   1935
   End
   Begin VB.Label Label1 
      Caption         =   "Label2"
      Height          =   255
      Index           =   82
      Left            =   -480
      TabIndex        =   182
      Top             =   19440
      Width           =   1815
   End
   Begin VB.Label Label1 
      Caption         =   "Label2"
      Height          =   255
      Index           =   81
      Left            =   -480
      TabIndex        =   181
      Top             =   19080
      Width           =   1935
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   255
      Index           =   80
      Left            =   -480
      TabIndex        =   180
      Top             =   18720
      Width           =   1935
   End
   Begin VB.Label Label1 
      Caption         =   "Label2"
      Height          =   255
      Index           =   79
      Left            =   -480
      TabIndex        =   179
      Top             =   18360
      Width           =   2055
   End
   Begin VB.Label Label1 
      Caption         =   "Label2"
      Height          =   255
      Index           =   78
      Left            =   3000
      TabIndex        =   178
      Top             =   20880
      Width           =   2055
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   255
      Index           =   77
      Left            =   -480
      TabIndex        =   177
      Top             =   20880
      Width           =   1935
   End
   Begin VB.Label Label1 
      Caption         =   "Label2"
      Height          =   255
      Index           =   76
      Left            =   -480
      TabIndex        =   176
      Top             =   20520
      Width           =   1815
   End
   Begin VB.Label Label1 
      Caption         =   "Label2"
      Height          =   255
      Index           =   75
      Left            =   6480
      TabIndex        =   175
      Top             =   20880
      Width           =   1935
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   255
      Index           =   74
      Left            =   6480
      TabIndex        =   174
      Top             =   20520
      Width           =   1935
   End
   Begin VB.Label Label1 
      Caption         =   "Label2"
      Height          =   255
      Index           =   73
      Left            =   6480
      TabIndex        =   173
      Top             =   20160
      Width           =   2055
   End
   Begin VB.Label Label1 
      Caption         =   "Label2"
      Height          =   255
      Index           =   72
      Left            =   -480
      TabIndex        =   148
      Top             =   16200
      Width           =   2055
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   255
      Index           =   71
      Left            =   -480
      TabIndex        =   147
      Top             =   16560
      Width           =   1935
   End
   Begin VB.Label Label1 
      Caption         =   "Label2"
      Height          =   255
      Index           =   70
      Left            =   -480
      TabIndex        =   146
      Top             =   16920
      Width           =   1935
   End
   Begin VB.Label Label1 
      Caption         =   "Label2"
      Height          =   255
      Index           =   69
      Left            =   -480
      TabIndex        =   145
      Top             =   17280
      Width           =   1815
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   255
      Index           =   68
      Left            =   -480
      TabIndex        =   144
      Top             =   17640
      Width           =   1935
   End
   Begin VB.Label Label1 
      Caption         =   "Label2"
      Height          =   255
      Index           =   67
      Left            =   -480
      TabIndex        =   143
      Top             =   18000
      Width           =   2055
   End
   Begin VB.Label Label1 
      Caption         =   "Label2"
      Height          =   255
      Index           =   66
      Left            =   6480
      TabIndex        =   142
      Top             =   16200
      Width           =   2055
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   255
      Index           =   65
      Left            =   6480
      TabIndex        =   141
      Top             =   16560
      Width           =   1935
   End
   Begin VB.Label Label1 
      Caption         =   "Label2"
      Height          =   255
      Index           =   64
      Left            =   6480
      TabIndex        =   140
      Top             =   16920
      Width           =   1935
   End
   Begin VB.Label Label1 
      Caption         =   "Label2"
      Height          =   255
      Index           =   63
      Left            =   6480
      TabIndex        =   139
      Top             =   17280
      Width           =   1815
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   255
      Index           =   62
      Left            =   3000
      TabIndex        =   138
      Top             =   16200
      Width           =   1935
   End
   Begin VB.Label Label1 
      Caption         =   "Label2"
      Height          =   255
      Index           =   61
      Left            =   3000
      TabIndex        =   137
      Top             =   16560
      Width           =   2055
   End
   Begin VB.Label Label1 
      Caption         =   "Label2"
      Height          =   255
      Index           =   60
      Left            =   3000
      TabIndex        =   136
      Top             =   16920
      Width           =   2055
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   255
      Index           =   59
      Left            =   3000
      TabIndex        =   135
      Top             =   17280
      Width           =   1935
   End
   Begin VB.Label Label1 
      Caption         =   "Label2"
      Height          =   255
      Index           =   58
      Left            =   3000
      TabIndex        =   134
      Top             =   17640
      Width           =   1935
   End
   Begin VB.Label Label1 
      Caption         =   "Label2"
      Height          =   255
      Index           =   57
      Left            =   3000
      TabIndex        =   133
      Top             =   18000
      Width           =   1815
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   255
      Index           =   56
      Left            =   6480
      TabIndex        =   132
      Top             =   17640
      Width           =   1935
   End
   Begin VB.Label Label1 
      Caption         =   "Label2"
      Height          =   255
      Index           =   55
      Left            =   6480
      TabIndex        =   131
      Top             =   18000
      Width           =   2055
   End
   Begin VB.Label Label1 
      Caption         =   "Label2"
      Height          =   255
      Index           =   54
      Left            =   6240
      TabIndex        =   112
      Top             =   12120
      Width           =   2055
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   255
      Index           =   53
      Left            =   6240
      TabIndex        =   111
      Top             =   12480
      Width           =   1935
   End
   Begin VB.Label Label1 
      Caption         =   "Label2"
      Height          =   255
      Index           =   52
      Left            =   6240
      TabIndex        =   110
      Top             =   12480
      Width           =   1935
   End
   Begin VB.Label Label1 
      Caption         =   "Label2"
      Height          =   255
      Index           =   51
      Left            =   6360
      TabIndex        =   109
      Top             =   12240
      Width           =   1815
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   255
      Index           =   50
      Left            =   6360
      TabIndex        =   108
      Top             =   12600
      Width           =   1935
   End
   Begin VB.Label Label1 
      Caption         =   "Label2"
      Height          =   255
      Index           =   49
      Left            =   6360
      TabIndex        =   107
      Top             =   12960
      Width           =   2055
   End
   Begin VB.Label Label1 
      Caption         =   "Label2"
      Height          =   255
      Index           =   48
      Left            =   6360
      TabIndex        =   106
      Top             =   13320
      Width           =   2055
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   255
      Index           =   47
      Left            =   6360
      TabIndex        =   105
      Top             =   13680
      Width           =   1935
   End
   Begin VB.Label Label1 
      Caption         =   "Label2"
      Height          =   255
      Index           =   46
      Left            =   6360
      TabIndex        =   104
      Top             =   14040
      Width           =   1935
   End
   Begin VB.Label Label1 
      Caption         =   "Label2"
      Height          =   255
      Index           =   45
      Left            =   6360
      TabIndex        =   103
      Top             =   14400
      Width           =   1815
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   255
      Index           =   44
      Left            =   6480
      TabIndex        =   102
      Top             =   13320
      Width           =   1935
   End
   Begin VB.Label Label1 
      Caption         =   "Label2"
      Height          =   255
      Index           =   43
      Left            =   6480
      TabIndex        =   101
      Top             =   13680
      Width           =   2055
   End
   Begin VB.Label Label1 
      Caption         =   "Label2"
      Height          =   255
      Index           =   42
      Left            =   6480
      TabIndex        =   100
      Top             =   14040
      Width           =   2055
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   255
      Index           =   41
      Left            =   6480
      TabIndex        =   99
      Top             =   14400
      Width           =   1935
   End
   Begin VB.Label Label1 
      Caption         =   "Label2"
      Height          =   255
      Index           =   40
      Left            =   6480
      TabIndex        =   98
      Top             =   14760
      Width           =   1935
   End
   Begin VB.Label Label1 
      Caption         =   "Label2"
      Height          =   255
      Index           =   39
      Left            =   6480
      TabIndex        =   97
      Top             =   15120
      Width           =   1815
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   255
      Index           =   38
      Left            =   6480
      TabIndex        =   96
      Top             =   15480
      Width           =   1935
   End
   Begin VB.Label Label1 
      Caption         =   "Label2"
      Height          =   255
      Index           =   37
      Left            =   6480
      TabIndex        =   95
      Top             =   15840
      Width           =   2055
   End
   Begin VB.Label Label1 
      Caption         =   "Label2"
      Height          =   255
      Index           =   36
      Left            =   3000
      TabIndex        =   72
      Top             =   15840
      Width           =   2055
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   255
      Index           =   35
      Left            =   3000
      TabIndex        =   71
      Top             =   15480
      Width           =   1935
   End
   Begin VB.Label Label1 
      Caption         =   "Label2"
      Height          =   255
      Index           =   34
      Left            =   3000
      TabIndex        =   70
      Top             =   15120
      Width           =   1815
   End
   Begin VB.Label Label1 
      Caption         =   "Label2"
      Height          =   255
      Index           =   33
      Left            =   3000
      TabIndex        =   66
      Top             =   14760
      Width           =   1935
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   255
      Index           =   32
      Left            =   3000
      TabIndex        =   65
      Top             =   14400
      Width           =   1935
   End
   Begin VB.Label Label1 
      Caption         =   "Label2"
      Height          =   255
      Index           =   31
      Left            =   3000
      TabIndex        =   64
      Top             =   14040
      Width           =   2055
   End
   Begin VB.Label Label1 
      Caption         =   "Label2"
      Height          =   255
      Index           =   30
      Left            =   3000
      TabIndex        =   60
      Top             =   13680
      Width           =   2055
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   255
      Index           =   29
      Left            =   3000
      TabIndex        =   59
      Top             =   13320
      Width           =   1935
   End
   Begin VB.Label Label1 
      Caption         =   "Label2"
      Height          =   255
      Index           =   28
      Left            =   2880
      TabIndex        =   58
      Top             =   14400
      Width           =   1815
   End
   Begin VB.Label Label1 
      Caption         =   "Label2"
      Height          =   255
      Index           =   27
      Left            =   2880
      TabIndex        =   54
      Top             =   14040
      Width           =   1935
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   255
      Index           =   26
      Left            =   2880
      TabIndex        =   53
      Top             =   13680
      Width           =   1935
   End
   Begin VB.Label Label1 
      Caption         =   "Label2"
      Height          =   255
      Index           =   25
      Left            =   2880
      TabIndex        =   52
      Top             =   13320
      Width           =   2055
   End
   Begin VB.Label Label1 
      Caption         =   "Label2"
      Height          =   255
      Index           =   24
      Left            =   2880
      TabIndex        =   48
      Top             =   12960
      Width           =   2055
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   255
      Index           =   23
      Left            =   2880
      TabIndex        =   47
      Top             =   12600
      Width           =   1935
   End
   Begin VB.Label Label1 
      Caption         =   "Label2"
      Height          =   255
      Index           =   22
      Left            =   2880
      TabIndex        =   46
      Top             =   12240
      Width           =   1815
   End
   Begin VB.Label Label1 
      Caption         =   "Label2"
      Height          =   255
      Index           =   21
      Left            =   2760
      TabIndex        =   42
      Top             =   12480
      Width           =   1935
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   255
      Index           =   20
      Left            =   2760
      TabIndex        =   41
      Top             =   12480
      Width           =   1935
   End
   Begin VB.Label Label1 
      Caption         =   "Label2"
      Height          =   255
      Index           =   19
      Left            =   2760
      TabIndex        =   40
      Top             =   12120
      Width           =   2055
   End
   Begin VB.Label Label1 
      Caption         =   "Label2"
      Height          =   255
      Index           =   18
      Left            =   -480
      TabIndex        =   36
      Top             =   15840
      Width           =   2055
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   255
      Index           =   17
      Left            =   -480
      TabIndex        =   35
      Top             =   15480
      Width           =   1935
   End
   Begin VB.Label Label1 
      Caption         =   "Label2"
      Height          =   255
      Index           =   16
      Left            =   -480
      TabIndex        =   34
      Top             =   15120
      Width           =   1815
   End
   Begin VB.Label Label1 
      Caption         =   "Label2"
      Height          =   255
      Index           =   15
      Left            =   -480
      TabIndex        =   30
      Top             =   14760
      Width           =   1935
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   255
      Index           =   14
      Left            =   -480
      TabIndex        =   29
      Top             =   14400
      Width           =   1935
   End
   Begin VB.Label Label1 
      Caption         =   "Label2"
      Height          =   255
      Index           =   13
      Left            =   -480
      TabIndex        =   28
      Top             =   14040
      Width           =   2055
   End
   Begin VB.Label Label1 
      Caption         =   "Label2"
      Height          =   255
      Index           =   12
      Left            =   -480
      TabIndex        =   24
      Top             =   13680
      Width           =   2055
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   255
      Index           =   11
      Left            =   -480
      TabIndex        =   23
      Top             =   13320
      Width           =   1935
   End
   Begin VB.Label Label1 
      Caption         =   "Label2"
      Height          =   255
      Index           =   10
      Left            =   -600
      TabIndex        =   22
      Top             =   14400
      Width           =   1815
   End
   Begin VB.Label Label1 
      Caption         =   "Label2"
      Height          =   255
      Index           =   9
      Left            =   -600
      TabIndex        =   18
      Top             =   14040
      Width           =   1935
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   255
      Index           =   8
      Left            =   -600
      TabIndex        =   17
      Top             =   13680
      Width           =   1935
   End
   Begin VB.Label Label1 
      Caption         =   "Label2"
      Height          =   255
      Index           =   7
      Left            =   -600
      TabIndex        =   16
      Top             =   13320
      Width           =   2055
   End
   Begin VB.Label Label1 
      Caption         =   "Label2"
      Height          =   255
      Index           =   6
      Left            =   -600
      TabIndex        =   12
      Top             =   12960
      Width           =   2055
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   255
      Index           =   5
      Left            =   -600
      TabIndex        =   11
      Top             =   12600
      Width           =   1935
   End
   Begin VB.Label Label1 
      Caption         =   "Label2"
      Height          =   255
      Index           =   4
      Left            =   -600
      TabIndex        =   10
      Top             =   12240
      Width           =   1815
   End
   Begin VB.Label Label1 
      Caption         =   "Label2"
      Height          =   255
      Index           =   3
      Left            =   -720
      TabIndex        =   5
      Top             =   12480
      Width           =   1935
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   255
      Index           =   2
      Left            =   -720
      TabIndex        =   4
      Top             =   12120
      Width           =   1935
   End
   Begin VB.Label Label1 
      Caption         =   "Label2"
      Height          =   255
      Index           =   1
      Left            =   1800
      TabIndex        =   3
      Top             =   12120
      Width           =   2055
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Nombre"
      Height          =   255
      Index           =   0
      Left            =   5040
      TabIndex        =   2
      Top             =   120
      Width           =   3975
   End
   Begin VB.Menu mnuArchivo 
      Caption         =   "Archivo"
      Begin VB.Menu mnuGuardar 
         Caption         =   "Guardar (CTRL+G)"
         Shortcut        =   ^G
      End
      Begin VB.Menu mnuGuardarTodo 
         Caption         =   "Guardar todo (CTRL+T)"
         Shortcut        =   ^T
      End
   End
End
Attribute VB_Name = "frmDats"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit



Private Type tAnim
    grhAnim As Integer
    Direccion As Byte
End Type
Dim currAnim As tAnim
Dim cargado As Boolean

Dim objtypeS() As String
Dim objtypeN() As Integer
Public CurrentIndex As Integer
Dim msgAyudaObjs(97) As String
Dim msgAyudaNPCs(97) As String
Dim msgAyudaHechizos(97) As String


Private Enum eNPCType
    NPCTYPE_COMUN = 0
    NPCTYPE_REVIVIR = 1
    NPCTYPE_GUARDIAS = 2
    NPCTYPE_ENTRENADOR = 3
    NPCTYPE_BANQUERO = 4
    NPCTYPE_NOBLE = 5
    NPCTYPE_APOSTADOR = 7
    NPCTYPE_TIENDA = 8
    NPCTYPE_QUEST = 18
    NPCTYPE_VIEJO = 11
    NPCTYPE_ROLERO = 19
    NPCTYPE_ROLERO2 = 20
    NPCTYPE_ROLERO3 = 21
    NPCTYPE_ROLERO4 = 22
    NPCTYPE_ROLERO5 = 23
    NPCTYPE_ROLERO6 = 24
    NPCTYPE_ROLERO7 = 25
    NPCTYPE_ROLERO8 = 26
    NPCTYPE_ROLERO9 = 27
    NPCTYPE_ROLERO10 = 28
    NPCTYPE_ROLERO11 = 29
    NPCTYPE_ROLERO12 = 30
    NPCTYPE_ROLERO13 = 31
    NPCTYPE_ROLERO14 = 32
    NPCTYPE_ROLERO15 = 33
    NPCTYPE_ROLERO16 = 34
    NPCTYPE_ROLERO17 = 35
    NPCTYPE_ROLERO18 = 36
    NPCTYPE_ROLERO19 = 37
    NPCTYPE_ROLERO20 = 38
    NPCTYPE_ROLERO21 = 39
    NPCTYPE_ROLERO22 = 40
    NPCTYPE_GEMSACRO = 41
    NPCTYPE_ROLERO24 = 42
    NPCTYPE_REMORT = 45
End Enum


Private Function GetObjtypeIndex(ByVal objtype As Integer) As Integer
    Dim X As Long
    For X = 0 To cmbObjType.ListCount
        If Val(ReadField(1, cmbObjType.List(X), Asc(" "))) = objtype Then
            GetObjtypeIndex = X
            Exit Function
        End If
        
    Next X
End Function

Private Sub ConfigureTxtLbl(ByVal Index As Byte, ByVal num As Byte, Optional ByVal Texto As String = "")
    Dim X As Long, Y As Long

        Y = IIf(estadoDat = eModo.Objetos, 1000, 700) + (num * 360)
        X = 3700
    If Y > 10000 Then
        Debug.Print num
        X = 7300
        Y = 6800 + ((num - 25) * 360)
    End If
    
    
    If estadoDat = Npc Then
        If Index = eNPCStats.NroItems Then
            Command3.Top = Y
            Command3.Left = X + 1440 + 1450
            Command3.Visible = True
        End If
        If Index = eNPCStats.LanzaSpells Then
            cmdLanzaSpells.Top = Y
            cmdLanzaSpells.Left = X + 1440 + 1450
            cmdLanzaSpells.Visible = True
        End If
        If Index = eNPCStats.Desc Then
            Label1(Index).Caption = "Descripcion"
        End If
    End If

    Label1(Index).Top = Y
    txtDatos(Index - 1).Top = Y
    Label1(Index).Left = X
    txtDatos(Index - 1).Left = X + 1440
    Label1(Index).Visible = True
    txtDatos(Index - 1).Visible = True
    Debug.Print Label1(Index).Caption & " - " & num
    If estadoDat = Objetos Then
        If Index = eObjStats.SubTipo Then
            txtDatos(Index - 1).Visible = False
            cmbSubtipo.Left = txtDatos(Index - 1).Left
            cmbSubtipo.Top = txtDatos(Index - 1).Top
            cmbSubtipo.Width = txtDatos(Index - 1).Width
            cmbSubtipo.listIndex = Val(Texto)
            cmbSubtipo.Visible = True
        End If
        If Index = eObjStats.Clasesprohib Then
            txtDatos(Index - 1).Visible = False
            cmdClasesProhibidas.Left = txtDatos(Index - 1).Left
            cmdClasesProhibidas.Top = txtDatos(Index - 1).Top
            cmdClasesProhibidas.Width = txtDatos(Index - 1).Width
            cmdClasesProhibidas.Visible = True
            
        End If
        If Index = eObjStats.Razasprohib Then
            txtDatos(Index - 1).Visible = False
            cmdRazasProhibidas.Left = txtDatos(Index - 1).Left
            cmdRazasProhibidas.Top = txtDatos(Index - 1).Top
            cmdRazasProhibidas.Width = txtDatos(Index - 1).Width
            cmdRazasProhibidas.Visible = True
            
        End If
    End If

    'If LenB(Texto) <> 0 Then
    txtDatos(Index - 1).Text = Texto
End Sub

Private Sub SetInfoHechizo(ByVal nINdex As Integer)
    Dim i As Long, hostilE As Boolean
    For i = 1 To 95
        Label1(i).Visible = False
        txtDatos(i - 1).Visible = False
    Next i
    Command3.Visible = False
    cmdLanzaSpells.Visible = False
    cmdRazasProhibidas.Visible = False
    cmdClasesProhibidas.Visible = False
    cmbSubtipo.Visible = False
    'cmbObjType.Visible = False
    
    If nINdex <= 0 Or nINdex > NumeroHechizos Then Exit Sub
    
    With Hechizos(nINdex)
        Hechizos(nINdex).Modificando = True
        Text1.Text = .Nombre
        Select Case .Tipo
            Case 1 'HP, MANA,
                ConfigureTxtLbl eHecStats.FXgrh, 1, .FXgrh
                ConfigureTxtLbl eHecStats.Desc, 2, .Desc
                ConfigureTxtLbl eHecStats.PalabrasMagicas, 3, .PalabrasMagicas
                ConfigureTxtLbl eHecStats.HechizeroMsg, 4, .HechizeroMsg
                ConfigureTxtLbl eHecStats.TargetMsg, 5, .TargetMsg
                ConfigureTxtLbl eHecStats.PropioMsg, 6, .PropioMsg
                ConfigureTxtLbl eHecStats.Target, 7, .Target
                ConfigureTxtLbl eHecStats.WAV, 8, .WAV
                ConfigureTxtLbl eHecStats.MinSkill, 9, .MinSkill
                ConfigureTxtLbl eHecStats.loops, 10, .loops
                
                ConfigureTxtLbl eHecStats.SubeHP, 11, .SubeHP
                msgAyudaHechizos(eHecStats.SubeHP) = "1> Cura vida" & vbCrLf & "2> Saca vida"
                ConfigureTxtLbl eHecStats.MinHP, 12, .MinHP
                ConfigureTxtLbl eHecStats.MaxHP, 13, .MaxHP
                
                ConfigureTxtLbl eHecStats.SubeHam, 14, .SubeHam
                msgAyudaHechizos(eHecStats.SubeHam) = "2> Baja hambre del objetivo"
                ConfigureTxtLbl eHecStats.MinHam, 15, .MinHam
                ConfigureTxtLbl eHecStats.MaxHam, 16, .MaxHam
                
                ConfigureTxtLbl eHecStats.SubeAgilidad, 17, .SubeAgilidad
                msgAyudaHechizos(eHecStats.SubeAgilidad) = "1> Sube agilidad del objetivo" & vbCrLf & "2>Baja agilidad del objetivo"
                ConfigureTxtLbl eHecStats.MinAgilidad, 18, .MinAgilidad
                ConfigureTxtLbl eHecStats.MaxAgilidad, 19, .MaxAgilidad
                
                ConfigureTxtLbl eHecStats.SubeFuerza, 20, .SubeFuerza
                msgAyudaHechizos(eHecStats.SubeFuerza) = "1> Sube fuerza del objetivo" & vbCrLf & "2>Baja fuerza del objetivo"
                ConfigureTxtLbl eHecStats.MinFuerza, 21, .MinFuerza
                ConfigureTxtLbl eHecStats.MaxFuerza, 22, .MaxFuerza
                
                ConfigureTxtLbl eHecStats.ManaRequerido, 23, .ManaRequerido
                ConfigureTxtLbl eHecStats.StaRequerido, 24, .StaRequerido
                ConfigureTxtLbl eHecStats.Baculo, 25, .Baculo
                
                ConfigureTxtLbl eHecStats.Nivel, 26, .Nivel
                msgAyudaHechizos(eHecStats.Nivel) = "Nivel minimo para usar el hechizo"
                ConfigureTxtLbl eHecStats.Resis, 27, .Resis
                
            Case 2 'estados del usuario
                ConfigureTxtLbl eHecStats.FXgrh, 1, .FXgrh
                ConfigureTxtLbl eHecStats.Desc, 2, .Desc
                ConfigureTxtLbl eHecStats.PalabrasMagicas, 3, .PalabrasMagicas
                ConfigureTxtLbl eHecStats.HechizeroMsg, 4, .HechizeroMsg
                ConfigureTxtLbl eHecStats.TargetMsg, 5, .TargetMsg
                ConfigureTxtLbl eHecStats.PropioMsg, 6, .PropioMsg
                ConfigureTxtLbl eHecStats.Target, 7, .Target
                ConfigureTxtLbl eHecStats.WAV, 8, .WAV
                ConfigureTxtLbl eHecStats.MinSkill, 9, .MinSkill
                ConfigureTxtLbl eHecStats.loops, 10, .loops
                
                ConfigureTxtLbl eHecStats.CuraVeneno, 11, .CuraVeneno ' =1 cura veneno
                ConfigureTxtLbl eHecStats.Envenena, 12, .Envenena ' =2 envenena
                ConfigureTxtLbl eHecStats.RemoverParalisis, 13, .RemoverParalisis ' =1 remueva para
                ConfigureTxtLbl eHecStats.Revivir, 14, .Revivir ' =1 revive 2=Resucita
                ConfigureTxtLbl eHecStats.Invisibilidad, 15, .Invisibilidad ' =1 invi
                ConfigureTxtLbl eHecStats.Paraliza, 16, .Paraliza ' 2=inmo  1=paralizar
                ConfigureTxtLbl eHecStats.Ceguera, 17, .Ceguera ' =1 ceguera
                ConfigureTxtLbl eHecStats.Estupidez, 18, .Estupidez ' =1 estupidiza 2=Remmueve estupidez
                ConfigureTxtLbl eHecStats.NoAtacar, 19, .NoAtacar ' =1 no atacar
                ConfigureTxtLbl eHecStats.Flecha, 20, .Flecha ' =1 aumenta golpe con arco

                ConfigureTxtLbl eHecStats.ManaRequerido, 21, .ManaRequerido
                ConfigureTxtLbl eHecStats.StaRequerido, 22, .StaRequerido
                ConfigureTxtLbl eHecStats.Baculo, 23, .Baculo
                
                ConfigureTxtLbl eHecStats.Nivel, 24, .Nivel
                msgAyudaHechizos(eHecStats.Nivel) = "Nivel minimo para usar el hechizo"
                ConfigureTxtLbl eHecStats.Resis, 25, .Resis
                
            
            Case 4 'invocaciones
            
                ConfigureTxtLbl eHecStats.FXgrh, 1, .FXgrh
                ConfigureTxtLbl eHecStats.Desc, 2, .Desc
                ConfigureTxtLbl eHecStats.PalabrasMagicas, 3, .PalabrasMagicas
                ConfigureTxtLbl eHecStats.HechizeroMsg, 4, .HechizeroMsg
                ConfigureTxtLbl eHecStats.Target, 5, .Target
                ConfigureTxtLbl eHecStats.WAV, 6, .WAV
                ConfigureTxtLbl eHecStats.MinSkill, 7, .MinSkill
                ConfigureTxtLbl eHecStats.loops, 8, .loops
                
                ConfigureTxtLbl eHecStats.Invoca, 9, .Invoca
                ConfigureTxtLbl eHecStats.NumNPC, 10, .NumNPC
                ConfigureTxtLbl eHecStats.cant, 11, .cant
                
                ConfigureTxtLbl eHecStats.ManaRequerido, 12, .ManaRequerido
                ConfigureTxtLbl eHecStats.StaRequerido, 13, .StaRequerido
                ConfigureTxtLbl eHecStats.Baculo, 14, .Baculo
                
                ConfigureTxtLbl eHecStats.Nivel, 15, .Nivel
                msgAyudaHechizos(eHecStats.Nivel) = "Nivel minimo para usar el hechizo"
                ConfigureTxtLbl eHecStats.Resis, 16, .Resis
                
            
            Case 6 'Hechizos de area
                
                ConfigureTxtLbl eHecStats.FXgrh, 1, .FXgrh
                ConfigureTxtLbl eHecStats.Desc, 2, .Desc
                ConfigureTxtLbl eHecStats.PalabrasMagicas, 3, .PalabrasMagicas
                ConfigureTxtLbl eHecStats.HechizeroMsg, 4, .HechizeroMsg
                ConfigureTxtLbl eHecStats.TargetMsg, 6, .TargetMsg
                ConfigureTxtLbl eHecStats.Target, 7, .Target
                
                ConfigureTxtLbl eHecStats.WAV, 8, .WAV
                ConfigureTxtLbl eHecStats.MinSkill, 9, .MinSkill
                ConfigureTxtLbl eHecStats.loops, 10, .loops
                
                ConfigureTxtLbl eHecStats.Invisibilidad, 11, .Invisibilidad
                ConfigureTxtLbl eHecStats.CuraArea, 12, .CuraArea
                
                ConfigureTxtLbl eHecStats.ManaRequerido, 13, .ManaRequerido
                ConfigureTxtLbl eHecStats.StaRequerido, 14, .StaRequerido
                ConfigureTxtLbl eHecStats.Baculo, 15, .Baculo
                
                ConfigureTxtLbl eHecStats.Nivel, 16, .Nivel
                msgAyudaHechizos(eHecStats.Nivel) = "Nivel minimo para usar el hechizo"
                ConfigureTxtLbl eHecStats.Resis, 17, .Resis
        End Select
        Hechizos(nINdex).Modificando = False
    End With
End Sub

Private Sub SetInfoNPC(ByVal npcIndex As Integer)
    Npclist(npcIndex).Modificando = True
    Dim i As Long, hostilE As Boolean
    For i = 1 To 95
        Label1(i).Visible = False
        txtDatos(i - 1).Visible = False
        txtDatos(i - 1).Text = ""
    Next i
    Command3.Visible = False
    cmdLanzaSpells.Visible = False

    Text1.Text = Npclist(npcIndex).name
    hostilE = Npclist(npcIndex).hostilE
    With Npclist(npcIndex)
        If hostilE = True Then
            ConfigureTxtLbl eNPCStats.AguaValida, 1, .flags.AguaValida
            ConfigureTxtLbl eNPCStats.TierraInvalida, 2, .flags.TierraInvalida
            ConfigureTxtLbl eNPCStats.Alineacion, 3, .Stats.Alineacion
            ConfigureTxtLbl eNPCStats.Atacable, 4, .Attackable
            ConfigureTxtLbl eNPCStats.Body, 5, .Char.Body
            ConfigureTxtLbl eNPCStats.Def, 6, .Stats.Def
            ConfigureTxtLbl eNPCStats.Domable, 7, .flags.Domable
            ConfigureTxtLbl eNPCStats.GiveEXP, 8, .GiveEXP
            ConfigureTxtLbl eNPCStats.GiveGLD, 9, .GiveGLD
            ConfigureTxtLbl eNPCStats.Head, 10, .Char.Head
            ConfigureTxtLbl eNPCStats.Heading, 11, .Char.Heading
            ConfigureTxtLbl eNPCStats.hostil, 12, .hostilE
            ConfigureTxtLbl eNPCStats.LanzaSpells, 13, .flags.LanzaSpells
            ConfigureTxtLbl eNPCStats.MinHP, 14, .Stats.MinHP
            ConfigureTxtLbl eNPCStats.MaxHP, 15, .Stats.MaxHP
            ConfigureTxtLbl eNPCStats.MinHit, 16, .Stats.MinHit
            ConfigureTxtLbl eNPCStats.MaxHit, 17, .Stats.MaxHit
            ConfigureTxtLbl eNPCStats.Movement, 18, .Movement
            ConfigureTxtLbl eNPCStats.NPCtype, 19, .NPCtype
            ConfigureTxtLbl eNPCStats.NroItems, 20, .Invent.NroItems
            ConfigureTxtLbl eNPCStats.PocaParalisis, 21, .flags.PocaParalisis
            ConfigureTxtLbl eNPCStats.PoderAtaque, 22, .PoderAtaque
            ConfigureTxtLbl eNPCStats.PoderEvasion, 23, .PoderEvasion
            ConfigureTxtLbl eNPCStats.Respawn, 24, .flags.Respawn
            ConfigureTxtLbl eNPCStats.RespawnOrigPos, 25, .flags.RespawnOrigPos
            ConfigureTxtLbl eNPCStats.Probabilidad, 26, .Probabilidad
    
            ConfigureTxtLbl eNPCStats.Snd1, 27, .flags.Snd1
            ConfigureTxtLbl eNPCStats.Snd2, 28, .flags.Snd2
            ConfigureTxtLbl eNPCStats.Snd3, 29, .flags.Snd3
            ConfigureTxtLbl eNPCStats.Snd4, 30, .flags.Snd4
            ConfigureTxtLbl eNPCStats.AutoCurar, 31, .AutoCurar
            ConfigureTxtLbl eNPCStats.VeInvis, 32, .VeInvis
            ConfigureTxtLbl eNPCStats.Veneno, 33, .Veneno
            ConfigureTxtLbl eNPCStats.GolpeExacto, 34, .flags.GolpeExacto
        Else
            ConfigureTxtLbl eNPCStats.Body, 1, .Char.Body
            ConfigureTxtLbl eNPCStats.Head, 2, .Char.Head
            ConfigureTxtLbl eNPCStats.Heading, 3, .Char.Heading
            ConfigureTxtLbl eNPCStats.Desc, 4, .Desc
            ConfigureTxtLbl eNPCStats.Comercia, 5, .Comercia
            ConfigureTxtLbl eNPCStats.AguaValida, 6, .flags.AguaValida
            ConfigureTxtLbl eNPCStats.Alineacion, 7, .Stats.Alineacion
            ConfigureTxtLbl eNPCStats.Atacable, 8, .Attackable
            ConfigureTxtLbl eNPCStats.Def, 9, .Stats.Def
            ConfigureTxtLbl eNPCStats.Domable, 10, .flags.Domable
            ConfigureTxtLbl eNPCStats.Faccion, 11, .flags.Faccion
            ConfigureTxtLbl eNPCStats.hostil, 12, .hostilE
            ConfigureTxtLbl eNPCStats.Inflacion, 13, .Inflacion
            ConfigureTxtLbl eNPCStats.InvReSpawn, 14, .InvReSpawn
            ConfigureTxtLbl eNPCStats.LanzaSpells, 15, .flags.LanzaSpells
            ConfigureTxtLbl eNPCStats.MaxHP, 16, .Stats.MaxHP
            ConfigureTxtLbl eNPCStats.MinHP, 17, .Stats.MinHP
            ConfigureTxtLbl eNPCStats.MaxHit, 18, .Stats.MaxHit
            ConfigureTxtLbl eNPCStats.MinHit, 19, .Stats.MinHit
            ConfigureTxtLbl eNPCStats.Movement, 20, .Movement
            ConfigureTxtLbl eNPCStats.NPCtype, 21, .NPCtype
            ConfigureTxtLbl eNPCStats.NroItems, 22, .Invent.NroItems
            ConfigureTxtLbl eNPCStats.PoderAtaque, 23, .PoderAtaque
            ConfigureTxtLbl eNPCStats.PoderEvasion, 24, .PoderEvasion
            ConfigureTxtLbl eNPCStats.Respawn, 25, .flags.Respawn
            ConfigureTxtLbl eNPCStats.RespawnOrigPos, 26, .flags.RespawnOrigPos
            ConfigureTxtLbl eNPCStats.Snd1, 27, .flags.Snd1
            ConfigureTxtLbl eNPCStats.Snd2, 28, .flags.Snd2
            ConfigureTxtLbl eNPCStats.Snd3, 29, .flags.Snd3
            ConfigureTxtLbl eNPCStats.Snd4, 30, .flags.Snd4
            ConfigureTxtLbl eNPCStats.Sound, 31, .flags.Sound
            ConfigureTxtLbl eNPCStats.TipoItems, 32, .TipoItems
           ' Label1(eNPCStats.TipoItems).Caption = "Tipo items"
        End If
    End With
    Npclist(npcIndex).Modificando = False
        'For i = 47 To 95
         '   Label1(i).Caption = ""
        'Next i
    'Else
    
    
    
    'End If
End Sub
Private Function getRazasString(ByVal OBJIndex As Integer) As String
    Dim i As Long, tmpStr As String
    For i = 1 To NUMRAZAS
        If ObjData(OBJIndex).RazaProhibida(i) <> 0 Then
            tmpStr = tmpStr & ObjData(OBJIndex).RazaProhibida(i) & "-"
        End If
    Next i
    If LenB(tmpStr) > 0 Then
        tmpStr = Left(tmpStr, LenB(tmpStr) - 1)
        getRazasString = tmpStr
    End If
End Function
Private Function getClasesString(ByVal OBJIndex As Integer) As String
    Dim i As Long, tmpStr As String
    For i = 1 To NUMCLASES
        If ObjData(OBJIndex).ClaseProhibida(i) <> 0 Then
            tmpStr = tmpStr & ObjData(OBJIndex).ClaseProhibida(i) & "-"
        End If
    Next i
    If LenB(tmpStr) > 0 Then
        tmpStr = Left(tmpStr, LenB(tmpStr) - 1)
        getClasesString = tmpStr
    End If
End Function



Private Sub SetInfoObj(ByVal OBJIndex As Integer)
    
    Dim Y As Long, X As Long
    For Y = 1 To 95
        Label1(Y).Visible = False
        txtDatos(Y - 1).Visible = False
    Next Y
    cmdRazasProhibidas.Visible = False
    cmdClasesProhibidas.Visible = False
        cmbSubtipo.Visible = False
    With ObjData(OBJIndex)
        Select Case .objtype
            Case eObjtype.Arboles
                ConfigureTxtLbl eObjStats.grhIndex, 0, .grhIndex
                ConfigureTxtLbl eObjStats.Agarrable, 1, .Agarrable
                ConfigureTxtLbl eObjStats.Arbol_Elfico, 2, .ArbolElfico
                ConfigureTxtLbl eObjStats.Resistencia, 3, .Resistencia
                ConfigureTxtLbl eObjStats.Info, 4, .Info
                
            Case eObjtype.Arma
                ConfigureTxtLbl eObjStats.grhIndex, 0, .grhIndex
                ConfigureTxtLbl eObjStats.Armaanim, 1, .WeaponAnim
                ConfigureTxtLbl eObjStats.Apuñala, 2, .Apuñala
                ConfigureTxtLbl eObjStats.MinHit, 3, .MinHit
                ConfigureTxtLbl eObjStats.MaxHit, 4, .MaxHit
                ConfigureTxtLbl eObjStats.Valor, 5, .Valor
                ConfigureTxtLbl eObjStats.SkillHerreria, 6, .SkHerreria
                ConfigureTxtLbl eObjStats.LingotesHierro, 7, .LingH
                ConfigureTxtLbl eObjStats.LingotesPlata, 8, .LingP
                ConfigureTxtLbl eObjStats.LingotesOro, 9, .LingO
                ConfigureTxtLbl eObjStats.Clasesprohib, 10, getClasesString(OBJIndex) 'PENDIENTEEEEEEEEEEEEEEEEEE
                ConfigureTxtLbl eObjStats.Dosmanos, 11, .Dosmanos
                ConfigureTxtLbl eObjStats.Baculo, 11, .Baculo
                ConfigureTxtLbl eObjStats.NoSeCae, 12, IIf(.NoSeCae = True, 1, 0)
                ConfigureTxtLbl eObjStats.Crucial, 13, .Crucial
                ConfigureTxtLbl eObjStats.proyectil, 14, .proyectil
                ConfigureTxtLbl eObjStats.Municion, 15, .Municion
                ConfigureTxtLbl eObjStats.Real, 16, .Real
                ConfigureTxtLbl eObjStats.Caos, 17, .Caos
                ConfigureTxtLbl eObjStats.Info, 18, .Info
                ConfigureTxtLbl eObjStats.Def, 19, .Def
                ConfigureTxtLbl eObjStats.plusMagia, 20, .plusMagia
                
            Case eObjtype.Armadura
                ConfigureTxtLbl eObjStats.grhIndex, 0, .grhIndex
                ConfigureTxtLbl eObjStats.SubTipo, 1, .SubTipo
                
                Select Case .SubTipo
                    Case 0
                        ConfigureTxtLbl eObjStats.Ropaje, 2, .Ropaje
                    
                    Case 1
                        
                        ConfigureTxtLbl eObjStats.CascoAnim, 2, .CascoAnim
                    Case 2
                        ConfigureTxtLbl eObjStats.Escudoanim, 2, .ShieldAnim
                End Select
                
                ConfigureTxtLbl eObjStats.MinDef, 3, .MinDef
                ConfigureTxtLbl eObjStats.MaxDef, 4, .MaxDef
                ConfigureTxtLbl eObjStats.Valor, 5, .Valor
                ConfigureTxtLbl eObjStats.SkillHerreria, 6, .SkHerreria
                ConfigureTxtLbl eObjStats.LingotesHierro, 7, .LingH
                ConfigureTxtLbl eObjStats.LingotesPlata, 8, .LingP
                ConfigureTxtLbl eObjStats.LingotesOro, 9, .LingO
                ConfigureTxtLbl eObjStats.SkillTacticas, 10, .SkillTacticas
                ConfigureTxtLbl eObjStats.Crucial, 11, .Crucial
                ConfigureTxtLbl eObjStats.HOMBRE, 12, .HOMBRE
                ConfigureTxtLbl eObjStats.MUJER, 13, .MUJER
                ConfigureTxtLbl eObjStats.Clasesprohib, 14, getClasesString(OBJIndex) 'PENDIENTEEEEEEEEEEEEEEEE
                ConfigureTxtLbl eObjStats.Razasprohib, 15, getRazasString(OBJIndex)
                
                ConfigureTxtLbl eObjStats.defensa, 16, .SkDefensa
                ConfigureTxtLbl eObjStats.NoSeCae, 17, IIf(.NoSeCae = True, 1, 0)
                
                ConfigureTxtLbl eObjStats.Real, 18, .Real
                ConfigureTxtLbl eObjStats.Caos, 19, .Caos
                ConfigureTxtLbl eObjStats.Jerarquia, 20, .Jerarquia
                ConfigureTxtLbl eObjStats.NoComerciable, 21, .NoComerciable
                ConfigureTxtLbl eObjStats.Gorro, 22, .Gorro
                ConfigureTxtLbl eObjStats.Def, 23, .Def
                ConfigureTxtLbl eObjStats.plusMagia, 24, .plusMagia
                ConfigureTxtLbl eObjStats.PielLobo, 25, .PielLobo
                ConfigureTxtLbl eObjStats.PielOso, 26, .PielOsoPardo
                ConfigureTxtLbl eObjStats.PielOsoPolar, 27, .PielOsoPolar

            Case eObjtype.Barcos
                ConfigureTxtLbl eObjStats.grhIndex, 0, .grhIndex
                ConfigureTxtLbl eObjStats.SkillCarpinteria, 1, .SkCarpinteria
                ConfigureTxtLbl eObjStats.Madera, 2, .Madera
                ConfigureTxtLbl eObjStats.Valor, 3, .Valor
                ConfigureTxtLbl eObjStats.Ropaje, 4, .Ropaje
                ConfigureTxtLbl eObjStats.MinSkill, 5, .MinSkill
                ConfigureTxtLbl eObjStats.MinHit, 6, .MinHit
                ConfigureTxtLbl eObjStats.MaxHit, 7, .MaxHit
                ConfigureTxtLbl eObjStats.MinDef, 8, .MinDef
                ConfigureTxtLbl eObjStats.MaxDef, 9, .MaxDef
                ConfigureTxtLbl eObjStats.NoComerciable, 10, .NoComerciable
                ConfigureTxtLbl eObjStats.Info, 11, .Info
                
            Case eObjtype.Bebidas
                ConfigureTxtLbl eObjStats.grhIndex, 0, .grhIndex
                ConfigureTxtLbl eObjStats.MinSed, 1, .MinSed
                ConfigureTxtLbl eObjStats.Valor, 2, .Valor
                ConfigureTxtLbl eObjStats.Crucial, 3, .Crucial
                ConfigureTxtLbl eObjStats.Info, 4, .Info
                
            Case eObjtype.Botellallena '34
                ConfigureTxtLbl eObjStats.grhIndex, 0, .grhIndex
                ConfigureTxtLbl eObjStats.MinSed, 1, .MinSed
                ConfigureTxtLbl eObjStats.Valor, 2, .Valor
                ConfigureTxtLbl eObjStats.MinSta, 3, .MinSta
                ConfigureTxtLbl eObjStats.IndexAbierta, 4, .IndexAbierta
                ConfigureTxtLbl eObjStats.IndexCerrada, 5, .IndexCerrada
                ConfigureTxtLbl eObjStats.Info, 6, .Info
                
            Case eObjtype.Botellavacia '33
                ConfigureTxtLbl eObjStats.grhIndex, 0, .grhIndex
                ConfigureTxtLbl eObjStats.Crucial, 1, .Crucial
                ConfigureTxtLbl eObjStats.Valor, 2, .Valor
                ConfigureTxtLbl eObjStats.IndexAbierta, 3, .IndexAbierta
                ConfigureTxtLbl eObjStats.IndexCerrada, 4, .IndexCerrada
                
            Case eObjtype.Carteles
                ConfigureTxtLbl eObjStats.grhIndex, 0, .grhIndex
                ConfigureTxtLbl eObjStats.GrhSecundario, 1, .GrhSecundario
                ConfigureTxtLbl eObjStats.Texto, 2, .Texto
                ConfigureTxtLbl eObjStats.Agarrable, 3, .Agarrable
                ConfigureTxtLbl eObjStats.Info, 4, .Info
                
            Case eObjtype.Contenedores
                ConfigureTxtLbl eObjStats.grhIndex, 0, .grhIndex
                ConfigureTxtLbl eObjStats.Llave, 1, .Llave
                ConfigureTxtLbl eObjStats.Resistencia, 2, .Resistencia
                ConfigureTxtLbl eObjStats.CantItems, 3, .MaxItems
                ConfigureTxtLbl eObjStats.IndexAbierta, 4, .IndexAbierta
                ConfigureTxtLbl eObjStats.IndexCerrada, 5, .IndexCerrada
                
            Case eObjtype.Dinero
                ConfigureTxtLbl eObjStats.grhIndex, 0, .grhIndex
                
            Case eObjtype.Flechas
                ConfigureTxtLbl eObjStats.grhIndex, 0, .grhIndex
                ConfigureTxtLbl eObjStats.MinHit, 1, .MinHit
                ConfigureTxtLbl eObjStats.MaxHit, 2, .MaxHit
                ConfigureTxtLbl eObjStats.Valor, 3, .Valor
                ConfigureTxtLbl eObjStats.SkillCarpinteria, 4, .SkCarpinteria
                ConfigureTxtLbl eObjStats.Madera, 5, .Madera
                ConfigureTxtLbl eObjStats.Municion, 6, .Municion
                
            Case eObjtype.Fogata
                ConfigureTxtLbl eObjStats.grhIndex, 0, .grhIndex
                ConfigureTxtLbl eObjStats.Agarrable, 1, .Agarrable
    
            Case eObjtype.Foros
                ConfigureTxtLbl eObjStats.grhIndex, 0, .grhIndex
                ConfigureTxtLbl eObjStats.ForoID, 1, .ForoID
                ConfigureTxtLbl eObjStats.Agarrable, 2, .Agarrable
                ConfigureTxtLbl eObjStats.GrhSecundario, 3, .GrhSecundario
                ConfigureTxtLbl eObjStats.Info, 4, .Info
                
            Case eObjtype.Libros
                 ConfigureTxtLbl eObjStats.grhIndex, 0, .grhIndex
                ConfigureTxtLbl eObjStats.GrhSecundario, 1, .GrhSecundario
                ConfigureTxtLbl eObjStats.Texto, 2, .Texto
                ConfigureTxtLbl eObjStats.Info, 3, .Info
                
            Case eObjtype.Fragua
                ConfigureTxtLbl eObjStats.grhIndex, 0, .grhIndex
                ConfigureTxtLbl eObjStats.Info, 1, .Info
                ConfigureTxtLbl eObjStats.Agarrable, 2, .Agarrable
                
            Case eObjtype.Herramientas
                ConfigureTxtLbl eObjStats.grhIndex, 0, .grhIndex
                ConfigureTxtLbl eObjStats.Valor, 1, .Valor
                ConfigureTxtLbl eObjStats.Armaanim, 2, .WeaponAnim
                ConfigureTxtLbl eObjStats.Crucial, 3, .Crucial
                ConfigureTxtLbl eObjStats.SkillCarpinteria, 4, .SkCarpinteria
                ConfigureTxtLbl eObjStats.Madera, 5, .Madera
                ConfigureTxtLbl eObjStats.SkillHerreria, 6, .SkHerreria
                ConfigureTxtLbl eObjStats.LingotesHierro, 7, .LingH
                
            Case eObjtype.Instrumentos
                ConfigureTxtLbl eObjStats.grhIndex, 0, .grhIndex
                ConfigureTxtLbl eObjStats.Valor, 1, .Valor
                ConfigureTxtLbl eObjStats.Snd1, 2, .Snd1
                ConfigureTxtLbl eObjStats.Snd2, 3, .Snd2
                ConfigureTxtLbl eObjStats.Snd3, 4, .Snd3
                ConfigureTxtLbl eObjStats.SkillCarpinteria, 5, .SkCarpinteria
                ConfigureTxtLbl eObjStats.Madera, 6, .Madera
                ConfigureTxtLbl eObjStats.MinInt, 7, .MinInt
            
            Case eObjtype.Leña
                ConfigureTxtLbl eObjStats.grhIndex, 0, .grhIndex
                ConfigureTxtLbl eObjStats.Valor, 1, .Valor
                ConfigureTxtLbl eObjStats.Crucial, 2, .Crucial
                
            Case eObjtype.Llaves
                ConfigureTxtLbl eObjStats.grhIndex, 0, .grhIndex
                ConfigureTxtLbl eObjStats.Clave, 1, .Clave
                ConfigureTxtLbl eObjStats.Valor, 2, .Valor
                ConfigureTxtLbl eObjStats.Info, 3, .Info
                ConfigureTxtLbl eObjStats.NoSeCae, 4, IIf(.NoSeCae = True, 1, 0)
                
            Case eObjtype.Manchas
                ConfigureTxtLbl eObjStats.grhIndex, 0, .grhIndex
                ConfigureTxtLbl eObjStats.Agarrable, 1, .Agarrable
                
            Case eObjtype.Minerales
                ConfigureTxtLbl eObjStats.grhIndex, 0, .grhIndex
                ConfigureTxtLbl eObjStats.Info, 1, .Info
                ConfigureTxtLbl eObjStats.MinSkill, 2, .MinSkill
                ConfigureTxtLbl eObjStats.LingoteIndex, 3, .LingoteIndex
                ConfigureTxtLbl eObjStats.Valor, 4, .Valor
                
            Case eObjtype.Pergaminos
                ConfigureTxtLbl eObjStats.grhIndex, 0, .grhIndex
                ConfigureTxtLbl eObjStats.Valor, 1, .Valor
                ConfigureTxtLbl eObjStats.Crucial, 2, .Crucial
                ConfigureTxtLbl eObjStats.HechizoIndex, 3, .HechizoIndex
            
            Case eObjtype.Piel
                ConfigureTxtLbl eObjStats.grhIndex, 0, .grhIndex
                ConfigureTxtLbl eObjStats.Valor, 1, .Valor
                ConfigureTxtLbl eObjStats.Crucial, 2, .Crucial
        
            Case eObjtype.pociones
                ConfigureTxtLbl eObjStats.grhIndex, 0, .grhIndex
                ConfigureTxtLbl eObjStats.Valor, 1, .Valor
                ConfigureTxtLbl eObjStats.TipoPocion, 2, .TipoPocion
                ConfigureTxtLbl eObjStats.MinModificador, 3, .MinModificador
                ConfigureTxtLbl eObjStats.MaxModificador, 4, .MaxModificador
                ConfigureTxtLbl eObjStats.DuracionEfecto, 5, .DuracionEfecto
                ConfigureTxtLbl eObjStats.Newbie, 6, .Newbie
                
            Case eObjtype.Puertas
                ConfigureTxtLbl eObjStats.grhIndex, 0, .grhIndex
                ConfigureTxtLbl eObjStats.Cerrada, 1, .Cerrada
                ConfigureTxtLbl eObjStats.Llave, 2, .Llave
                ConfigureTxtLbl eObjStats.Clave, 3, .Clave
                ConfigureTxtLbl eObjStats.Resistencia, 4, .Resistencia
                ConfigureTxtLbl eObjStats.Agarrable, 5, .Agarrable
                ConfigureTxtLbl eObjStats.IndexAbierta, 6, .IndexAbierta
                ConfigureTxtLbl eObjStats.IndexCerrada, 7, .IndexCerrada
                ConfigureTxtLbl eObjStats.IndexCerradaLlave, 8, .IndexCerradaLlave
            
            Case eObjtype.Teleports
                ConfigureTxtLbl eObjStats.grhIndex, 0, .grhIndex
                ConfigureTxtLbl eObjStats.Agarrable, 1, .Agarrable
    
            Case eObjtype.UseOnce
                ConfigureTxtLbl eObjStats.grhIndex, 0, .grhIndex
                ConfigureTxtLbl eObjStats.Minhambre, 1, .MinHam
                ConfigureTxtLbl eObjStats.Newbie, 2, .Newbie
                ConfigureTxtLbl eObjStats.Valor, 3, .Valor
            Case eObjtype.Warp
                ConfigureTxtLbl eObjStats.grhIndex, 0, .grhIndex
                ConfigureTxtLbl eObjStats.Map_Pasajes, 1, .WMapa
                ConfigureTxtLbl eObjStats.MapX, 2, .WX
                ConfigureTxtLbl eObjStats.MapY, 3, .WY
                ConfigureTxtLbl eObjStats.WarpI, 4, .WI
                ConfigureTxtLbl eObjStats.Valor, 5, .Valor
                
            Case eObjtype.Yacimiento
                ConfigureTxtLbl eObjStats.grhIndex, 0, .grhIndex
                ConfigureTxtLbl eObjStats.Info, 1, .Info
                ConfigureTxtLbl eObjStats.Agarrable, 2, .Agarrable
                ConfigureTxtLbl eObjStats.MineralIndex, 3, .MineralIndex
    
            Case eObjtype.Yunque
                ConfigureTxtLbl eObjStats.grhIndex, 0, .grhIndex
                ConfigureTxtLbl eObjStats.Agarrable, 1, .Agarrable
                ConfigureTxtLbl eObjStats.Info, 2, .Info
            
            Case eObjtype.Amuleto
                ConfigureTxtLbl eObjStats.grhIndex, 0, .grhIndex
                ConfigureTxtLbl eObjStats.Def, 1, .Def
                ConfigureTxtLbl eObjStats.plusMagia, 2, .plusMagia
                ConfigureTxtLbl eObjStats.MinDef, 3, .MinDef
                ConfigureTxtLbl eObjStats.MaxDef, 4, .MaxDef
                ConfigureTxtLbl eObjStats.NoSeCae, 5, IIf(.NoSeCae = True, 1, 0)
                ConfigureTxtLbl eObjStats.NoComerciable, 6, .NoComerciable
                ConfigureTxtLbl eObjStats.Aura, 7, .Aura
                ConfigureTxtLbl eObjStats.Remort, 8, .Remort
                
            Case eObjtype.Montura
                ConfigureTxtLbl eObjStats.grhIndex, 0, .grhIndex
                ConfigureTxtLbl eObjStats.MinDef, 1, .MinDef
                ConfigureTxtLbl eObjStats.MaxDef, 2, .MaxDef
                ConfigureTxtLbl eObjStats.Caballo, 3, .Caballo
                ConfigureTxtLbl eObjStats.Valor, 4, .Valor
                ConfigureTxtLbl eObjStats.Crucial, 5, .Crucial
                ConfigureTxtLbl eObjStats.NoSeCae, 6, IIf(.NoSeCae = True, 1, 0)
                ConfigureTxtLbl eObjStats.Def, 7, .Def
                ConfigureTxtLbl eObjStats.plusMagia, 8, .plusMagia
        End Select
        Text1.Text = .name
        If .objtype = 0 Then
            cmbObjType.listIndex = 0
            Exit Sub
        End If
        Dim nObjn As Integer
        nObjn = GetObjtypeIndex(.objtype)
        'If nObjn = 0 Then Exit Sub
        cmbObjType.listIndex = nObjn
        
    End With

    

    
End Sub

Private Sub cmbObjType_Change()
'Call SetInfoObj(Val(ReadField(1, cmbObjType.List(cmbObjType.listIndex), Asc(" "))))
End Sub

Private Sub cmbObjType_Click()

    'If estadoDat <> Objetos Then Exit Sub
    If estadoDat = Objetos Then
        If CurrentIndex > 0 And CurrentIndex <= numObjs Then
            With ObjData(CurrentIndex)
                If .Modificando = False Then
                    .objtype = Val(ReadField(1, cmbObjType.List(cmbObjType.listIndex), Asc(" ")))
                    Call SetInfoObj(CurrentIndex)
                    .FueModificado = True
                    DatNoGuardado(eModo.Objetos) = True
                End If
            End With
        End If
    ElseIf estadoDat = Hechizo Then
        If CurrentIndex > 0 And CurrentIndex <= NumeroHechizos Then
            With Hechizos(CurrentIndex)
                If .Modificando = False Then
                    .Tipo = Val(ReadField(1, cmbObjType.List(cmbObjType.listIndex), Asc(" ")))
                    Call SetInfoHechizo(CurrentIndex)
                    .FueModificado = True
                    DatNoGuardado(eModo.Hechizo) = True
                End If
            End With
        End If
    
    End If
     'ObjData(nIndex).objtype = Val(ReadField(1, cmbObjType.List(cmbObjType.listIndex), Asc(" ")))
     
 'Call SetObjData(nIndex)
End Sub


Public Sub CambiarEstadoDat(ByVal Index As eModo)
Dim i As Long
    Select Case Index
        Case eModo.Objetos
            Command2.Caption = "Nuevo objeto"
            Command4.Caption = "Guardar modificaciones en Obj.dat"
            cmbObjType.Visible = True
            Label2.Visible = True
            cmbObjType.Clear
            Label1(eObjStats.NoComerciable).Caption = "No comerciable"
            sObjStat(eObjStats.NoComerciable) = "NoComerciable"
    
            Label1(eObjStats.NoSeCae).Caption = "No se cae"
            sObjStat(eObjStats.NoSeCae) = "Nosecae"
            
            Label1(eObjStats.Agarrable).Caption = "Agarrable"
            sObjStat(eObjStats.Agarrable) = "Agarrable"
            
            Label1(eObjStats.objtype).Caption = "Objtype"
            sObjStat(eObjStats.objtype) = "Objtype"
            
            Label1(eObjStats.SubTipo).Caption = "Subtipo"
            sObjStat(eObjStats.SubTipo) = "Subtipo"
            msgAyudaObjs(eObjStats.SubTipo) = "(Armadura 0, Casco 1, Escudo 2)"
            
            Label1(eObjStats.Dosmanos).Caption = "Dosmanos"
            sObjStat(eObjStats.Dosmanos) = "Dosmanos"
            
            Label1(eObjStats.grhIndex).Caption = "Grhindex"
            sObjStat(eObjStats.grhIndex) = "GrhIndex"
            
            Label1(eObjStats.GrhSecundario).Caption = "Grhsecundario"
            sObjStat(eObjStats.GrhSecundario) = "VGrande"
            
            Label1(eObjStats.Jerarquia).Caption = "Jerarquia"
            sObjStat(eObjStats.Jerarquia) = "Jerarquia"
            
            Label1(eObjStats.Respawn).Caption = "Respawn"
            sObjStat(eObjStats.Respawn) = "Respawn"
            
            Label1(eObjStats.Def).Caption = "Def"
            sObjStat(eObjStats.Def) = "Def"
            
            Label1(eObjStats.MaxItems).Caption = "MaxItems"
            sObjStat(eObjStats.MaxItems) = "NroItems"
            
            Label1(eObjStats.Apuñala).Caption = "Apuñala"
            sObjStat(eObjStats.Apuñala) = "Apuñala"
            
            Label1(eObjStats.HechizoIndex).Caption = "Hechizo index"
            sObjStat(eObjStats.HechizoIndex) = "HechizoIndex"
            
            Label1(eObjStats.ForoID).Caption = "Foro ID"
            sObjStat(eObjStats.ForoID) = "ID"
            
            Label1(eObjStats.MinHP).Caption = "MinHP"
            sObjStat(eObjStats.MinHP) = "MinHP"
            Label1(eObjStats.MaxHP).Caption = "MaxHP"
            sObjStat(eObjStats.MaxHP) = "MaxHP"
            
            Label1(eObjStats.Map_Pasajes).Caption = "Mapa destino(Pasajes)"
            sObjStat(eObjStats.Map_Pasajes) = "WMapa"
            Label1(eObjStats.MapX).Caption = "MapX"
            sObjStat(eObjStats.MapX) = "WX"
            Label1(eObjStats.MapY).Caption = "MapY"
            sObjStat(eObjStats.MapY) = "WY"
            Label1(eObjStats.WarpI).Caption = "WarpI"
            sObjStat(eObjStats.WarpI) = "WI"
            
            Label1(eObjStats.Baculo).Caption = "Baculo"
            sObjStat(eObjStats.Baculo) = "Baculo"
            
            Label1(eObjStats.MineralIndex).Caption = "Mineral index"
            sObjStat(eObjStats.MineralIndex) = "MineralIndex"
            
            Label1(eObjStats.Texto).Caption = "Texto"
            sObjStat(eObjStats.Texto) = "Texto"
            
            Label1(eObjStats.proyectil).Caption = "Proyectil"
            sObjStat(eObjStats.proyectil) = "Proyectil"
            
            Label1(eObjStats.Municion).Caption = "Municion"
            sObjStat(eObjStats.Municion) = "Municiones"
            
            Label1(eObjStats.Crucial).Caption = "Crucial"
            sObjStat(eObjStats.Crucial) = "Crucial"
            
            Label1(eObjStats.Newbie).Caption = "Newbie"
            sObjStat(eObjStats.Newbie) = "Newbie"
            
            Label1(eObjStats.MinSta).Caption = "MinSTA"
            sObjStat(eObjStats.MinSta) = "MinST"
            
            Label1(eObjStats.TipoPocion).Caption = "Tipo pocion"
            sObjStat(eObjStats.TipoPocion) = "TipoPocion"
            
            Label1(eObjStats.MaxModificador).Caption = "Max modificador"
            sObjStat(eObjStats.MaxModificador) = "MaxModificador"
            
            Label1(eObjStats.MinModificador).Caption = "Min modificador"
            sObjStat(eObjStats.MinModificador) = "MinModificador"
            
            Label1(eObjStats.DuracionEfecto).Caption = "Duracion efecto"
            sObjStat(eObjStats.DuracionEfecto) = "DuracionEfecto"
            
            Label1(eObjStats.MinSkill).Caption = "MinSkill"
            sObjStat(eObjStats.MinSkill) = "MinSkill"
            
            Label1(eObjStats.LingoteIndex).Caption = "Lingote index"
            sObjStat(eObjStats.LingoteIndex) = "LingoteIndex"
            
            Label1(eObjStats.MinHit).Caption = "Min hit"
            sObjStat(eObjStats.MinHit) = "MinHit"
            
            Label1(eObjStats.MaxHit).Caption = "Max hit"
            sObjStat(eObjStats.MaxHit) = "MaxHit"
            
            Label1(eObjStats.Minhambre).Caption = "Min hambre"
            sObjStat(eObjStats.Minhambre) = "MinHam"
            
            Label1(eObjStats.MinSed).Caption = "Min sed"
            sObjStat(eObjStats.MinSed) = "MinAgu"
            
            Label1(eObjStats.defensa).Caption = "Defensa"
            sObjStat(eObjStats.defensa) = "skEscudos"
            
            Label1(eObjStats.MinDef).Caption = "Min def"
            sObjStat(eObjStats.MinDef) = "MinDef"
            
            Label1(eObjStats.MaxDef).Caption = "Max def"
            sObjStat(eObjStats.MaxDef) = "MaxDef"
            
            Label1(eObjStats.Ropaje).Caption = "Ropaje"
            sObjStat(eObjStats.Ropaje) = "NumRopaje"
            
            Label1(eObjStats.Armaanim).Caption = "Arma anim"
            sObjStat(eObjStats.Armaanim) = "Anim"
            
            Label1(eObjStats.Escudoanim).Caption = "Escudo anim"
            sObjStat(eObjStats.Escudoanim) = "Anim"
            
            Label1(eObjStats.CascoAnim).Caption = "Casco anim"
            sObjStat(eObjStats.CascoAnim) = "Anim"
            
            Label1(eObjStats.Gorro).Caption = "Gorro"
            sObjStat(eObjStats.Gorro) = "Gorro"
            
            Label1(eObjStats.Valor).Caption = "Valor"
            sObjStat(eObjStats.Valor) = "Valor"
            
            Label1(eObjStats.Cerrada).Caption = "Cerrada"
            sObjStat(eObjStats.Cerrada) = "Abierta"
            
            Label1(eObjStats.Llave).Caption = "Llave"
            sObjStat(eObjStats.Llave) = "Llave"
            
            Label1(eObjStats.Clave).Caption = "Clave"
            sObjStat(eObjStats.Clave) = "Clave"
            
            Label1(eObjStats.IndexAbierta).Caption = "Index abierta"
            sObjStat(eObjStats.IndexAbierta) = "IndexAbierta"
            Label1(eObjStats.IndexCerrada).Caption = "Index cerrada"
            sObjStat(eObjStats.IndexCerrada) = "IndexCerrada"
            Label1(eObjStats.IndexCerradaLlave).Caption = "Index cerrada llave"
            sObjStat(eObjStats.IndexCerradaLlave) = "IndexCerradaLlave"
            
            Label1(eObjStats.RazaEnana).Caption = "Raza enana"
            sObjStat(eObjStats.RazaEnana) = "RazaEnana"
            
            Label1(eObjStats.MUJER).Caption = "Mujer"
            sObjStat(eObjStats.MUJER) = "Vendible"
            Label1(eObjStats.HOMBRE).Caption = "Hombre"
            sObjStat(eObjStats.HOMBRE) = "Vendible"
            
            Label1(eObjStats.Envenena).Caption = "Envenena"
            sObjStat(eObjStats.Envenena) = "Envenena"
            
            Label1(eObjStats.SkillCombate).Caption = "Skill combate"
            sObjStat(eObjStats.SkillCombate) = "skCombate"
            Label1(eObjStats.SkillTacticas).Caption = "Skill tacticas"
            sObjStat(eObjStats.SkillTacticas) = "skTacticas"
            Label1(eObjStats.SkillProyectiles).Caption = "Skill proyectiles"
            sObjStat(eObjStats.SkillProyectiles) = "skProyectiles"
            Label1(eObjStats.SkillApuñalar).Caption = "Skill apuñalar"
            sObjStat(eObjStats.SkillApuñalar) = "skApuñalar"
            
            Label1(eObjStats.Resistencia).Caption = "Resistencia"
            sObjStat(eObjStats.Resistencia) = "Resistencia"
            
            Label1(eObjStats.Agarrable).Caption = "Agarrable"
            sObjStat(eObjStats.Agarrable) = "Agarrable"
            
            Label1(eObjStats.Arbol_Elfico).Caption = "Arbol elfico"
            sObjStat(eObjStats.Arbol_Elfico) = "ArbolElfico"
            
            Label1(eObjStats.LingotesOro).Caption = "Lingotes Oro"
            sObjStat(eObjStats.LingotesOro) = "LingO"
            Label1(eObjStats.LingotesPlata).Caption = "Lingotes Plata"
            sObjStat(eObjStats.LingotesPlata) = "LingP"
            Label1(eObjStats.LingotesHierro).Caption = "Lingotes Hierro"
            sObjStat(eObjStats.LingotesHierro) = "LingH"
            
            Label1(eObjStats.Madera).Caption = "Madera"
            sObjStat(eObjStats.Madera) = "Madera"
            Label1(eObjStats.Madera_Elfica).Caption = "Madera elfica"
            sObjStat(eObjStats.Madera_Elfica) = "MaderaElfica"
            
            Label1(eObjStats.Raices).Caption = "Raices"
            sObjStat(eObjStats.Raices) = "Raices"
            
            Label1(eObjStats.PielLobo).Caption = "Piel lobo"
            sObjStat(eObjStats.PielLobo) = "PielLobo"
            Label1(eObjStats.PielOso).Caption = "Piel oso pardo"
            sObjStat(eObjStats.PielOso) = "PielOsoPardo"
            Label1(eObjStats.PielOsoPolar).Caption = "Piel oso polar"
            sObjStat(eObjStats.PielOsoPolar) = "PielOsoPolar"
            
            Label1(eObjStats.SkillHerreria).Caption = "Skill Herreria"
            sObjStat(eObjStats.SkillHerreria) = "skHerreria"
            Label1(eObjStats.SkillCarpinteria).Caption = "Skill Carpinteria"
            sObjStat(eObjStats.SkillCarpinteria) = "skCarpinteria"
            Label1(eObjStats.SkillResistencia).Caption = "Skill resistencia"
            sObjStat(eObjStats.SkillResistencia) = "skResistencias"
            'Label1(eObjStats.Skilldefensa).Caption = "Skill defensa"
            'sObjStat(eObjStats.Skilldefensa) = "Vendible"
            Label1(eObjStats.Skillpociones).Caption = "Skill pociones"
            sObjStat(eObjStats.Skillpociones) = "skPociones"
            Label1(eObjStats.Skillsastreria).Caption = "Skill sastreria"
            sObjStat(eObjStats.Skillsastreria) = "skSastreria"
            
            Label1(eObjStats.Clasesprohib).Caption = "Clases prohibidas"
            msgAyudaObjs(eObjStats.Clasesprohib) = "Se ponen las clases prohibidas en el siguiente formato: '35-1-2-4-52'"
            sObjStat(eObjStats.Clasesprohib) = "NumClases"
            
            Label1(eObjStats.Razasprohib).Caption = "Razas prohibidas"
            msgAyudaObjs(eObjStats.Razasprohib) = "Se ponen las razas prohibidas en el siguiente formato: '3-1-2'"
            sObjStat(eObjStats.Razasprohib) = "NumRazas"
            
            Label1(eObjStats.Snd1).Caption = "SND1"
            sObjStat(eObjStats.Snd1) = "SND1"
            Label1(eObjStats.Snd2).Caption = "SND2"
            sObjStat(eObjStats.Snd2) = "SND2"
            Label1(eObjStats.Snd3).Caption = "SND3"
            sObjStat(eObjStats.Snd3) = "SND3"
            
            Label1(eObjStats.MinInt).Caption = "MinInt"
            sObjStat(eObjStats.MinInt) = "MinINT"
            
            Label1(eObjStats.Real).Caption = "Real"
            sObjStat(eObjStats.Real) = "Real"
            Label1(eObjStats.Caos).Caption = "Caos"
            sObjStat(eObjStats.Caos) = "Caos"
            Label1(eObjStats.CantItems).Caption = "CantItems"
            sObjStat(eObjStats.CantItems) = "CantItems"
            Label1(eObjStats.Info).Caption = "Info"
            sObjStat(eObjStats.Info) = "Info"
            
            Label1(eObjStats.Caballo).Caption = "Caballo Body"
            sObjStat(eObjStats.Caballo) = "Caballo"
            Label1(eObjStats.plusMagia).Caption = "Plus Magia"
            sObjStat(eObjStats.plusMagia) = "PlusMagia"
            Label1(eObjStats.Aura).Caption = "Aura"
            sObjStat(eObjStats.Aura) = "Aura"
            'Call CargarListaObj
            Label2.Caption = "Tipo de objeto"
            cmbObjType.AddItem "0 - <<<Seleccionar objtype>>>"
            cmbObjType.AddItem "1 - USEONCE "
            cmbObjType.AddItem "2 - ARMAS"
            cmbObjType.AddItem "3 - ROPAS/CASCOS/ESCUDOS/ARMADURAS"
            cmbObjType.AddItem "4 - ARBOLES"
            cmbObjType.AddItem "5 - DINERO"
            cmbObjType.AddItem "6 - PUERTAS"
            cmbObjType.AddItem "7 - CONTENEDORES"
            cmbObjType.AddItem "8 - CARTELES  "
            cmbObjType.AddItem "9 - LLAVES  "
            cmbObjType.AddItem "10 - FOROS  "
            cmbObjType.AddItem "11 - POCIONES  "
            cmbObjType.AddItem "13 - BEBIDA  "
            cmbObjType.AddItem "14 - LEÑA  "
            cmbObjType.AddItem "15 - FOGATA  "
            cmbObjType.AddItem "18 - HERRAMIENTAS  "
            cmbObjType.AddItem "22 - YACIMIENTO  "
            cmbObjType.AddItem "24 - PERGAMINOS  "
            cmbObjType.AddItem "19 - TELEPORT  "
            cmbObjType.AddItem "27 - YUNQUE  "
            cmbObjType.AddItem "28 - FRAGUA  "
            cmbObjType.AddItem "23 - MINERALES  "
            cmbObjType.AddItem "26 - INSTRUMENTOS  "
            cmbObjType.AddItem "31 - BARCOS  "
            cmbObjType.AddItem "32 - FLECHAS  "
            cmbObjType.AddItem "33 - BOTELLAVACIA  "
            cmbObjType.AddItem "34 - BOTELLALLENA  "
            cmbObjType.AddItem "35 - MANCHAS  "
            cmbObjType.AddItem "29 - GEMAZUL  "
            cmbObjType.AddItem "38 - GEMNARANJA  "
            cmbObjType.AddItem "39 - GEMCELESTE  "
            cmbObjType.AddItem "40 - GEMLILA  "
            cmbObjType.AddItem "41 - GEMROJO  "
            cmbObjType.AddItem "42 - GEMVERDE  "
            cmbObjType.AddItem "43 - GEMVIOLETA  "
            cmbObjType.AddItem "44 - AMULETO  "
            cmbObjType.AddItem "36 - RAIZ  "
            cmbObjType.AddItem "30 - PIEL  "
            cmbObjType.AddItem "45 - MONTURA  "
            cmbObjType.AddItem "37 - WARP  "
            cmbObjType.AddItem "1000 - CUALQUIERA  "
            For i = 1 To numObjs
                If LenB(ObjData(i).name) <> 0 Then
                    List1.AddItem i & " - " & ObjData(i).name
                Else
                    List1.AddItem i & " <<-SLOT VACIO->> "
                End If
            Next i
            Command2.Caption = "Crear nuevo objeto"
            Command4.Caption = "Guardar modificaciones en OBJ.dat"
            'Call LoadOBJData
        Case eModo.Npc
            cmbObjType.Visible = False
            Label2.Visible = False
            Command2.Caption = "Crear nuevo NPC"
            Command4.Caption = "Guardar modificaciones en NPCs.dat y NPCs-hostiles.dat"

            
            For i = 1 To 90
                frmDats.Label1(i).Caption = sNpcStat(i)
                
            Next i
            For i = 1 To MaxNPCnohostiles
                If Npclist(i).Char.Body <> 0 Or LenB(Npclist(i).name) > 0 Then
                    List1.AddItem i & " - " & Npclist(i).name
                Else
                    List1.AddItem i & " - <<-SLOT VACIO->>"
                End If
            Next i
            
            For i = 500 To MaxNPC + 500
                If Npclist(i).Char.Body <> 0 Or LenB(Npclist(i).name) > 0 Then
                    List1.AddItem i & " - " & Npclist(i).name & "(HOSTIL)"
                Else
                    List1.AddItem i & " - <<-SLOT VACIO->>"
                End If
            Next i
            
        Case eModo.Hechizo
            Dim Hechizo As Integer
            cmbObjType.Visible = True
            Label2.Visible = True
            
            cmbObjType.Clear
            cmbObjType.AddItem "0 - <<Selecccionar tipo de hechizo>>"
            cmbObjType.AddItem "1 - Sobre HP, MANA, STA, HAM Y SED"
            cmbObjType.AddItem "2 - Sobre estados de usuarios(Paralizar, remover, etc)"
            cmbObjType.AddItem "3 - Materializa"
            cmbObjType.AddItem "4 - Invocaciones"
            cmbObjType.AddItem "5 - Metamorfosis"
            cmbObjType.AddItem "6 - Hechizos de area"
            Label2.Caption = "Tipo de hechizo"
            
        
            For i = 1 To 90
                frmDats.Label1(i).Caption = sHecStat(i)
                
            Next i
            For Hechizo = 1 To NumeroHechizos
                If LenB(Hechizos(Hechizo).Nombre) <> 0 Then
                    List1.AddItem Hechizo & " - " & Hechizos(Hechizo).Nombre
                Else
                    List1.AddItem Hechizo & " <<-SLOT VACIO->> "
                End If

            Next Hechizo
            
            Command2.Caption = "Crear nuevo hechizo"
            Command4.Caption = "Guardar modificaciones en Hechizos.dat"
            cmbTarget.Clear
            cmbTarget.AddItem "0 - Selecciona target"
            cmbTarget.AddItem "1 - Usuarios"
            cmbTarget.AddItem "2 - NPCs"
            cmbTarget.AddItem "3 - Usuarios y npcs"
            cmbTarget.AddItem "4 - Terreno"
    End Select
    If Index = 0 Then Exit Sub
    estadoDat = Index
    Combo1.listIndex = Index
    List1.listIndex = 0
    List1_Click
End Sub

Public Sub SetNuevoDatInfo()

End Sub

Private Function fClaseProhibida(ByRef ClasesProhibidas() As Integer, ByVal Clase As Byte) As Boolean
    Dim i As Long
    For i = 1 To NUMCLASES
        If ClasesProhibidas(i) = Clase Then
            fClaseProhibida = True
            Exit Function
        End If
    Next i
    fClaseProhibida = False
End Function

Private Function fRazaProhibida(ByRef RazasProhibidas() As Integer, ByVal Raza As Byte) As Boolean
    Dim i As Long
    For i = 1 To NUMRAZAS
        If RazasProhibidas(i) = Raza Then
            fRazaProhibida = True
            Exit Function
        End If
    Next i
    fRazaProhibida = False
End Function

Private Sub cmbSubtipo_Change()
    If estadoDat <> Objetos Then Exit Sub
    If CurrentIndex > 0 And CurrentIndex <= numObjs Then
        With ObjData(CurrentIndex)
            If .Modificando = False Then
                .SubTipo = Val(ReadField(1, frmDats.cmbSubtipo.List(frmDats.cmbSubtipo.listIndex), Asc(" ")))
                .FueModificado = True
                DatNoGuardado(eModo.Objetos) = True
            End If
            
            Select Case .SubTipo
                    Case 0
                        ConfigureTxtLbl eObjStats.Ropaje, 2, .Ropaje
                    
                    Case 1
                        ConfigureTxtLbl eObjStats.CascoAnim, 2, .CascoAnim
                        
                    Case 2
                        ConfigureTxtLbl eObjStats.Escudoanim, 2, .ShieldAnim
                End Select
        End With
    End If
End Sub

Private Sub cmbSubtipo_Click()
    If estadoDat <> Objetos Then Exit Sub
    If CurrentIndex > 0 And CurrentIndex <= numObjs Then
        With ObjData(CurrentIndex)
            If .Modificando = False Then
                .SubTipo = Val(ReadField(1, frmDats.cmbSubtipo.List(frmDats.cmbSubtipo.listIndex), Asc(" ")))
                .FueModificado = True
                DatNoGuardado(eModo.Objetos) = True
            End If
            
            Select Case .SubTipo
                    Case 0
                        ConfigureTxtLbl eObjStats.Ropaje, 2, .Ropaje
                    
                    Case 1
                        ConfigureTxtLbl eObjStats.CascoAnim, 2, .CascoAnim
                        
                    Case 2
                        ConfigureTxtLbl eObjStats.Escudoanim, 2, .ShieldAnim
                End Select
        End With
    End If
End Sub

Private Sub cmdClasesProhibidas_Click()
    If estadoDat <> Objetos Then Exit Sub
    If CurrentIndex > numObjs Then Exit Sub
    If CurrentIndex <= 0 Then Exit Sub
    frmProhibidas.Show , Me
    frmProhibidas.ModoProhibidos = clases
    frmProhibidas.List1.Clear
    frmProhibidas.List2.Clear
    frmProhibidas.Caption = "Clases prohibidas"
    Dim i As Long
    For i = 1 To NUMCLASES 'Aqui generamos la lista de clases prohibidas
        With ObjData(CurrentIndex)
            If .ClaseProhibida(i) > 0 And .ClaseProhibida(i) <= NUMCLASES Then
                frmProhibidas.List1.AddItem .ClaseProhibida(i) & " - " & ListaClases(.ClaseProhibida(i))
            End If
        End With
    Next i
    
    For i = 1 To NUMCLASES 'Aqui la lista de clases permitidas
        If ListaClases(i) <> "" Then 'Si esta registrada esta clase
            If fClaseProhibida(ObjData(CurrentIndex).ClaseProhibida, i) = False Then
                frmProhibidas.List2.AddItem i & " - " & ListaClases(i)
            End If
        End If
    Next i
End Sub

Private Sub cmdLanzaSpells_Click()
    frmLanzaSpells.LoadLanzaSpells
    frmLanzaSpells.Show , Me
End Sub

Private Sub cmdRazasProhibidas_Click()
    If estadoDat <> Objetos Then Exit Sub
    If CurrentIndex > numObjs Then Exit Sub
    If CurrentIndex <= 0 Then Exit Sub
    frmProhibidas.Show , Me
    frmProhibidas.ModoProhibidos = razas
    frmProhibidas.List1.Clear
    frmProhibidas.List2.Clear
    frmProhibidas.Caption = "Razas prohibidas"
    Dim i As Long
    For i = 1 To NUMRAZAS
        With ObjData(CurrentIndex)
            If .RazaProhibida(i) > 0 And .RazaProhibida(i) <= NUMRAZAS Then
                frmProhibidas.List1.AddItem .RazaProhibida(i) & " " & ListaRazas(.RazaProhibida(i))
            End If
        End With
    Next i
    
    For i = 1 To NUMRAZAS
        If ListaRazas(i) <> "" Then 'Si esta registrada esta clase
            If fRazaProhibida(ObjData(CurrentIndex).RazaProhibida, i) = False Then
                frmProhibidas.List2.AddItem i & " - " & ListaRazas(i)
            End If
        End If
    Next i
    
End Sub

Private Sub Combo1_Click()
    'If nModo = Combo1.listIndex Then GoTo sig
    If estadoDat = (Combo1.listIndex) Then Exit Sub
    If configDats.modo = 1 Then Exit Sub
    List1.Clear
    Call CambiarEstadoDat(Combo1.listIndex)
    
sig:
    
End Sub


Private Sub Command1_Click()
    Dim ii As Long, nPath As String
    Dim num As Integer, nArray() As String
    Dim i As Long
    Dim nObj As Integer
    Select Case estadoDat
        Case eModo.Objetos
            
            'Call WriteVar(NPATH, "OBJ" & nObj, "Name", Text1.Text)
            'Call WriteVar(NPATH, "OBJ" & nObj, "Objtype", ReadField(1, cmbObjType.List(cmbObjType.listIndex), Asc(" ")))
            
            
            'For ii = 1 To 94
            '    If txtDatos(ii).Visible = True Then
            '        If ii <> 80 And ii <> 81 Then
            '            Call WriteVar(NPATH, "OBJ" & nObj, sObjStat(ii), txtDatos(ii - 1).Text)
            '        End If
            '        If ii = 80 Then  'Clases prohib. '59-35-32-30'
                    
            '            nArray = Split(txtDatos(ii - 1).Text, "-")
            '            num = UBound(nArray) + 1
            '
            '            For i = 1 To num
            '                Call WriteVar(NPATH, "OBJ" & nObj, "CP" & i, nArray(i - 1)) 'ObjData(Object).ClaseProhibida(i) = INIDarClaveInt(A, S, "CP" & i)
            '            Next
            '        End If
            '        If ii = 81 Then 'Raza prohibida. '2-3-4'
            '            nArray = Split(txtDatos(ii - 1).Text, "-")
            '            num = UBound(nArray) + 1
            '
            '            For i = 1 To num
            '                Call WriteVar(NPATH, "OBJ" & nObj, "RP" & i, nArray(i - 1)) 'ObjData(Object).ClaseProhibida(i) = INIDarClaveInt(A, S, "CP" & i)
            '            Next
            '        End If
            '    End If
            ' Next ii
            
    End Select
End Sub

Private Sub Command2_Click()
    Dim i As Long, nINdex As Integer
    Dim buscarslot As Byte
    Dim hostilE As Byte
    Select Case estadoDat
        Case eModo.Hechizo
            
            buscarslot = MsgBox("¿Deseas buscar un slot de HECHIZOS.dat que este vacío?", vbYesNo, "Nuevo hechizo")
            
            If buscarslot = vbYes Then
                
                For i = 1 To NumeroHechizos
                    If Hechizos(i).Nombre = vbNullString And Hechizos(i).Tipo = 0 Then 'And ObjData(i).grhIndex = 0 Then 'Esta vacio
                        nINdex = i
                        Exit For
                    End If
                Next i
                
                If nINdex > 0 Then
                    List1.listIndex = nINdex - 1
                Else
                    MsgBox "No se ha encontrado ningun slot vacio. Se creará uno nuevo. "
                    List1.AddItem NumeroHechizos + 1 & " - Nuevo hechizo"
                    NumeroHechizos = NumeroHechizos + 1
                    ReDim Preserve Hechizos(1 To NumeroHechizos) As tHechizo
                    'ReDim ObjData(0 To numObjs)
                    List1.listIndex = List1.ListCount - 1
                End If
            ElseIf buscarslot = vbNo Then
                List1.AddItem NumeroHechizos + 1 & " - Nuevo hechizo"
                NumeroHechizos = NumeroHechizos + 1
                ReDim Preserve Hechizos(1 To NumeroHechizos) As tHechizo
                'ReDim ObjData(0 To numObjs)
                List1.listIndex = List1.ListCount - 1
            End If
        Case eModo.Objetos
            
            buscarslot = MsgBox("¿Deseas buscar un slot de OBJ.dat que este vacío?", vbYesNo, "Nuevo objeto")
            
            If buscarslot = vbYes Then
                
                For i = 1 To numObjs
                    If ObjData(i).name = "" And ObjData(i).objtype = 0 And ObjData(i).grhIndex = 0 Then 'Esta vacio
                        nINdex = i
                        Exit For
                    End If
                Next i
                
                If nINdex > 0 Then
                    List1.listIndex = nINdex - 1
                Else
                    MsgBox "No se ha encontrado ningun slot vacio."
                End If
            ElseIf buscarslot = vbNo Then
                List1.AddItem numObjs + 1 & " - Nuevo objeto"
                numObjs = numObjs + 1
                ReDim Preserve ObjData(0 To numObjs) As ObjData
                'ReDim ObjData(0 To numObjs)
                List1.listIndex = List1.ListCount - 1
            End If
            
        Case eModo.Npc
            Dim xx As Long

            Dim nSlot As Integer
            Dim tINT As Integer, tI As Integer, lSlot As Integer
            hostilE = MsgBox("¿Este NPC será hostil? Si es hostil se guardara en el indice 'NPCs-HOSTILES.dat'; si NO es hostil en 'NPCs.dat'", vbYesNo, "Nuevo NPC")
            buscarslot = MsgBox("¿Deseas buscar un slot de NPC que este vacío?", vbYesNo, "Nuevo NPC")
            'buscarslot = vbNo
            
            If hostilE = vbYes Then
                hostilE = 1
            ElseIf hostilE = vbNo Then
                hostilE = 0
                
            End If
            
            If buscarslot = vbYes Then
                If hostilE = 1 Then
                    For i = 500 To 500 + MaxNPC
                        If Npclist(i).name = "" And Npclist(i).Char.Body = 0 Then 'Esta vacio
                            nINdex = i
                            Exit For
                        End If
                    Next i
                Else 'No es hostil
                    For i = 1 To MaxNPCnohostiles
                        If Npclist(i).name = "" And Npclist(i).Char.Body = 0 Then 'Esta vacio
                            nINdex = i
                            Exit For
                        End If
                    Next i
                End If
                If nINdex > 0 Then
                    
                    
                    For tINT = 0 To frmDats.List1.ListCount - 1
                        If Val(ReadField(1, frmDats.List1.List(tINT), Asc(" "))) = nINdex Then
                            lSlot = tINT
                            Exit For
                        End If
                    Next tINT
                   nSlot = nINdex
                   
                   frmDats.List1.listIndex = lSlot
                    
                    'nINdex
                    'nSlot = nINdex
                    
                Else
                    If hostilE = 1 Then
                        MsgBox "No se ha encontrado ningun slot vacio. Se creará uno nuevo"
                        frmDats.List1.AddItem 500 + MaxNPC + 1 & " - Nuevo NPC"
                        MaxNPC = MaxNPC + 1
                        nSlot = MaxNPC
                        ReDim Preserve Npclist(1 To 500 + MaxNPC) As Npc
                        frmDats.List1.listIndex = frmDats.List1.ListCount - 1
                    Else 'No es hostiles
                        
                        With frmDats.List1
                        'tint ti lslot
                            For tINT = 0 To .ListCount - 1
                                If Val(ReadField(1, .List(tINT), Asc(" "))) = 500 Then
                                    lSlot = tINT
                                    Exit For
                                End If
                            Next tINT
                            
                            .AddItem .List(.ListCount - 1)
                            For tI = .ListCount - 2 To (lSlot + 1) Step -1
                                'Desde el primer npc hostil hasta el final de la lista
                                .List(tI) = .List(tI - 1)
                            Next tI
                            nSlot = lSlot
                            frmDats.List1.listIndex = lSlot
                        End With
                    End If
                End If
            ElseIf buscarslot = vbNo Then
                
               If hostilE = 1 Then
                    frmDats.List1.AddItem 500 + MaxNPC + 1 & " - Nuevo NPC(Hostil)"
                    MaxNPC = MaxNPC + 1
                    nSlot = MaxNPC + 500
                    ReDim Preserve Npclist(1 To 500 + MaxNPC) As Npc
                    frmDats.List1.listIndex = frmDats.List1.ListCount - 1
                Else 'No es hostiles
                        
                    With frmDats.List1
                    'tint ti lslot
                        For tINT = 0 To .ListCount - 1
                            If Val(ReadField(1, .List(tINT), Asc(" "))) = 500 Then
                                lSlot = tINT
                                Exit For
                            End If
                        Next tINT
                        
                        'MsgBox .ListCount
                        .AddItem (.List(.ListCount - 1))
                        For tI = .ListCount - 2 To (lSlot + 1) Step -1
                            'Desde el primer npc hostil hasta el final de la lista
                            .List(tI) = .List(tI - 1)
                        Next tI
                        MaxNPCnohostiles = MaxNPCnohostiles + 1
                        .List(lSlot) = MaxNPCnohostiles & " - Nuevo NPC"
                        nSlot = lSlot
                        frmDats.List1.listIndex = lSlot
                    End With
                End If
            End If
           
            
            DatNoGuardado(eModo.Npc) = True
            
            Npclist(nSlot).name = "Nuevo NPC"
            Npclist(nSlot).Char.Body = DataIndexActual
            Npclist(nSlot).FueModificado = True
             DatNoGuardado(eModo.Npc) = True
    End Select
End Sub

Private Sub Command3_Click()
    Call DibujarInventarioNPC(Val(ReadField(1, List1.List(List1.listIndex), Asc(" "))))
End Sub

Public Sub GuardarEstadoActual()
    Dim OBJIndex As Long, nProh As Byte
    Dim npcIndex As Long
    Dim nPath As String
    Dim i As Long
    Dim AvisoStr As String
    Dim numChanges As Integer
    Select Case estadoDat
        Case eModo.Objetos
            If DatNoGuardado(eModo.Objetos) = False Then MsgBox "No se ha registrado ningun cambio hasta ahora en los objetos": Exit Sub
            If MsgBox("¿Deseas sobreescribir los cambios en '" & ConfigDir.Dats & "\OBJ.dat'? (Si presionas NO, los cambios se guardaran en un archivo adicional de nombre OBJ1.dat)", vbYesNo, "Sobreescribir archivo") = vbYes Then
                nPath = ConfigDir.Dats & "/OBJ.dat"
            Else
                nPath = ConfigDir.Dats & "/OBJ1.dat"
            End If
            AvisoStr = "Se han guardado los cambios de los siguientes objetos: "
            Call WriteVar(nPath, "INIT", "NumOBJs", numObjs)
            
            For OBJIndex = 1 To numObjs
                With ObjData(OBJIndex)
                    If .FueModificado = True Then 'Solo guardamos los que fueron modificados
                        AvisoStr = AvisoStr & OBJIndex & "(" & .name & ")" & ", "
                        numChanges = numChanges + 1
                        Call WriteVar(nPath, "OBJ" & OBJIndex, "Name", .name)
                        Call WriteVar(nPath, "OBJ" & OBJIndex, "ObjType", .objtype)
                        Select Case .objtype
                            Case eObjtype.Arboles
                                Call WriteVar(nPath, "OBJ" & OBJIndex, sObjStat(eObjStats.grhIndex), .grhIndex)
                                Call WriteVar(nPath, "OBJ" & OBJIndex, sObjStat(eObjStats.Agarrable), .Agarrable)
                                Call WriteVar(nPath, "OBJ" & OBJIndex, sObjStat(eObjStats.Arbol_Elfico), .ArbolElfico)
                                Call WriteVar(nPath, "OBJ" & OBJIndex, sObjStat(eObjStats.Resistencia), .Resistencia)
                                Call WriteVar(nPath, "OBJ" & OBJIndex, sObjStat(eObjStats.Info), .Info)
                                
                            Case eObjtype.Arma
                                Call WriteVar(nPath, "OBJ" & OBJIndex, sObjStat(eObjStats.grhIndex), .grhIndex)
                                Call WriteVar(nPath, "OBJ" & OBJIndex, sObjStat(eObjStats.Armaanim), .WeaponAnim)
                                Call WriteVar(nPath, "OBJ" & OBJIndex, sObjStat(eObjStats.Apuñala), .Apuñala)
                                Call WriteVar(nPath, "OBJ" & OBJIndex, sObjStat(eObjStats.MinHit), .MinHit)
                                Call WriteVar(nPath, "OBJ" & OBJIndex, sObjStat(eObjStats.MaxHit), .MaxHit)
                                Call WriteVar(nPath, "OBJ" & OBJIndex, sObjStat(eObjStats.Valor), .Valor)
                                Call WriteVar(nPath, "OBJ" & OBJIndex, sObjStat(eObjStats.SkillHerreria), .SkHerreria)
                                Call WriteVar(nPath, "OBJ" & OBJIndex, sObjStat(eObjStats.LingotesHierro), .LingH)
                                Call WriteVar(nPath, "OBJ" & OBJIndex, sObjStat(eObjStats.LingotesPlata), .LingP)
                                Call WriteVar(nPath, "OBJ" & OBJIndex, sObjStat(eObjStats.LingotesOro), .LingO)
                                nProh = 0
                                For i = 1 To NUMCLASES
                                    If .ClaseProhibida(i) <> 0 Then
                                        nProh = nProh + 1
                                        Call WriteVar(nPath, "OBJ" & OBJIndex, "CP" & nProh, .ClaseProhibida(i))
                                    End If
                                Next i
                                'call writevar(nPath, "OBJ" & OBJINDEX, sObjStat(eObjStats.Clasesprohib), getClasesString(Objindex) 'PENDIENTEEEEEEEEEEEEEEEEEE
                                Call WriteVar(nPath, "OBJ" & OBJIndex, sObjStat(eObjStats.Dosmanos), .Dosmanos)
                                Call WriteVar(nPath, "OBJ" & OBJIndex, sObjStat(eObjStats.Baculo), .Baculo)
                                Call WriteVar(nPath, "OBJ" & OBJIndex, sObjStat(eObjStats.NoSeCae), .NoSeCae)
                                Call WriteVar(nPath, "OBJ" & OBJIndex, sObjStat(eObjStats.Crucial), .Crucial)
                                Call WriteVar(nPath, "OBJ" & OBJIndex, sObjStat(eObjStats.proyectil), .proyectil)
                                Call WriteVar(nPath, "OBJ" & OBJIndex, sObjStat(eObjStats.Municion), .Municion)
                                Call WriteVar(nPath, "OBJ" & OBJIndex, sObjStat(eObjStats.Real), .Real)
                                Call WriteVar(nPath, "OBJ" & OBJIndex, sObjStat(eObjStats.Caos), .Caos)
                                Call WriteVar(nPath, "OBJ" & OBJIndex, sObjStat(eObjStats.Info), .Info)
                                Call WriteVar(nPath, "OBJ" & OBJIndex, sObjStat(eObjStats.Def), .Def)
                                Call WriteVar(nPath, "OBJ" & OBJIndex, sObjStat(eObjStats.plusMagia), .plusMagia)
                                
                            Case eObjtype.Armadura
                                Call WriteVar(nPath, "OBJ" & OBJIndex, sObjStat(eObjStats.grhIndex), .grhIndex)
                                Call WriteVar(nPath, "OBJ" & OBJIndex, sObjStat(eObjStats.SubTipo), .SubTipo)
                                Select Case .SubTipo
                                    Case 0 'armadura o ropaje
                                        Call WriteVar(nPath, "OBJ" & OBJIndex, sObjStat(eObjStats.Ropaje), .Ropaje)
                                    
                                    Case 1 'cascos
                                        Call WriteVar(nPath, "OBJ" & OBJIndex, sObjStat(eObjStats.CascoAnim), .CascoAnim)
                                
                                    Case 2 'escudos
                                        Call WriteVar(nPath, "OBJ" & OBJIndex, sObjStat(eObjStats.Escudoanim), .ShieldAnim)
                                End Select
                                Call WriteVar(nPath, "OBJ" & OBJIndex, sObjStat(eObjStats.MinDef), .MinDef)
                                Call WriteVar(nPath, "OBJ" & OBJIndex, sObjStat(eObjStats.MaxDef), .MaxDef)
                                Call WriteVar(nPath, "OBJ" & OBJIndex, sObjStat(eObjStats.Valor), .Valor)
                                Call WriteVar(nPath, "OBJ" & OBJIndex, sObjStat(eObjStats.SkillHerreria), .SkHerreria)
                                Call WriteVar(nPath, "OBJ" & OBJIndex, sObjStat(eObjStats.LingotesHierro), .LingH)
                                Call WriteVar(nPath, "OBJ" & OBJIndex, sObjStat(eObjStats.LingotesPlata), .LingP)
                                Call WriteVar(nPath, "OBJ" & OBJIndex, sObjStat(eObjStats.LingotesOro), .LingO)
                                Call WriteVar(nPath, "OBJ" & OBJIndex, sObjStat(eObjStats.SkillTacticas), .SkillTacticas)
                                Call WriteVar(nPath, "OBJ" & OBJIndex, sObjStat(eObjStats.Crucial), .Crucial)
                                Call WriteVar(nPath, "OBJ" & OBJIndex, sObjStat(eObjStats.HOMBRE), .HOMBRE)
                                Call WriteVar(nPath, "OBJ" & OBJIndex, sObjStat(eObjStats.MUJER), .MUJER)
                                nProh = 0
                                For i = 1 To NUMCLASES
                                    If .ClaseProhibida(i) <> 0 Then
                                        nProh = nProh + 1
                                        Call WriteVar(nPath, "OBJ" & OBJIndex, "CP" & nProh, .ClaseProhibida(i))
                                    End If
                                Next i
                                nProh = 0
                                For i = 1 To NUMRAZAS
                                    If .RazaProhibida(i) <> 0 Then
                                        nProh = nProh + 1
                                        Call WriteVar(nPath, "OBJ" & OBJIndex, "RP" & nProh, .RazaProhibida(i))
                                    End If
                                Next i
    
                                Call WriteVar(nPath, "OBJ" & OBJIndex, sObjStat(eObjStats.defensa), .SkDefensa)
                                Call WriteVar(nPath, "OBJ" & OBJIndex, sObjStat(eObjStats.NoSeCae), .NoSeCae)
                                
                                Call WriteVar(nPath, "OBJ" & OBJIndex, sObjStat(eObjStats.Real), .Real)
                                Call WriteVar(nPath, "OBJ" & OBJIndex, sObjStat(eObjStats.Caos), .Caos)
                                Call WriteVar(nPath, "OBJ" & OBJIndex, sObjStat(eObjStats.Jerarquia), .Jerarquia)
                                Call WriteVar(nPath, "OBJ" & OBJIndex, sObjStat(eObjStats.NoComerciable), .NoComerciable)
                                Call WriteVar(nPath, "OBJ" & OBJIndex, sObjStat(eObjStats.Gorro), .Gorro)
                                Call WriteVar(nPath, "OBJ" & OBJIndex, sObjStat(eObjStats.Info), .Info)
                                Call WriteVar(nPath, "OBJ" & OBJIndex, sObjStat(eObjStats.Def), .Def)
                                Call WriteVar(nPath, "OBJ" & OBJIndex, sObjStat(eObjStats.plusMagia), .plusMagia)
                                
                            Case eObjtype.Barcos
                                Call WriteVar(nPath, "OBJ" & OBJIndex, sObjStat(eObjStats.grhIndex), .grhIndex)
                                Call WriteVar(nPath, "OBJ" & OBJIndex, sObjStat(eObjStats.SkillCarpinteria), .SkCarpinteria)
                                Call WriteVar(nPath, "OBJ" & OBJIndex, sObjStat(eObjStats.Madera), .Madera)
                                Call WriteVar(nPath, "OBJ" & OBJIndex, sObjStat(eObjStats.Valor), .Valor)
                                Call WriteVar(nPath, "OBJ" & OBJIndex, sObjStat(eObjStats.Ropaje), .Ropaje)
                                Call WriteVar(nPath, "OBJ" & OBJIndex, sObjStat(eObjStats.MinSkill), .MinSkill)
                                Call WriteVar(nPath, "OBJ" & OBJIndex, sObjStat(eObjStats.MinHit), .MinHit)
                                Call WriteVar(nPath, "OBJ" & OBJIndex, sObjStat(eObjStats.MaxHit), .MaxHit)
                                Call WriteVar(nPath, "OBJ" & OBJIndex, sObjStat(eObjStats.MinDef), .MinDef)
                                Call WriteVar(nPath, "OBJ" & OBJIndex, sObjStat(eObjStats.MaxDef), .MaxDef)
                                Call WriteVar(nPath, "OBJ" & OBJIndex, sObjStat(eObjStats.NoComerciable), .NoComerciable)
                                Call WriteVar(nPath, "OBJ" & OBJIndex, sObjStat(eObjStats.Info), .Info)
                                
                            Case eObjtype.Bebidas
                                Call WriteVar(nPath, "OBJ" & OBJIndex, sObjStat(eObjStats.grhIndex), .grhIndex)
                                Call WriteVar(nPath, "OBJ" & OBJIndex, sObjStat(eObjStats.MinSed), .MinSed)
                                Call WriteVar(nPath, "OBJ" & OBJIndex, sObjStat(eObjStats.Valor), .Valor)
                                Call WriteVar(nPath, "OBJ" & OBJIndex, sObjStat(eObjStats.Crucial), .Crucial)
                                Call WriteVar(nPath, "OBJ" & OBJIndex, sObjStat(eObjStats.Info), .Info)
                                
                            Case eObjtype.Botellallena '34
                                Call WriteVar(nPath, "OBJ" & OBJIndex, sObjStat(eObjStats.grhIndex), .grhIndex)
                                Call WriteVar(nPath, "OBJ" & OBJIndex, sObjStat(eObjStats.MinSed), .MinSed)
                                Call WriteVar(nPath, "OBJ" & OBJIndex, sObjStat(eObjStats.Valor), .Valor)
                                Call WriteVar(nPath, "OBJ" & OBJIndex, sObjStat(eObjStats.MinSta), .MinSta)
                                Call WriteVar(nPath, "OBJ" & OBJIndex, sObjStat(eObjStats.IndexAbierta), .IndexAbierta)
                                Call WriteVar(nPath, "OBJ" & OBJIndex, sObjStat(eObjStats.IndexCerrada), .IndexCerrada)
                                Call WriteVar(nPath, "OBJ" & OBJIndex, sObjStat(eObjStats.Info), .Info)
                                
                            Case eObjtype.Botellavacia '33
                                Call WriteVar(nPath, "OBJ" & OBJIndex, sObjStat(eObjStats.grhIndex), .grhIndex)
                                Call WriteVar(nPath, "OBJ" & OBJIndex, sObjStat(eObjStats.Crucial), .Crucial)
                                Call WriteVar(nPath, "OBJ" & OBJIndex, sObjStat(eObjStats.Valor), .Valor)
                                Call WriteVar(nPath, "OBJ" & OBJIndex, sObjStat(eObjStats.IndexAbierta), .IndexAbierta)
                                Call WriteVar(nPath, "OBJ" & OBJIndex, sObjStat(eObjStats.IndexCerrada), .IndexCerrada)
                                
                            Case eObjtype.Carteles
                                Call WriteVar(nPath, "OBJ" & OBJIndex, sObjStat(eObjStats.grhIndex), .grhIndex)
                                Call WriteVar(nPath, "OBJ" & OBJIndex, sObjStat(eObjStats.GrhSecundario), .GrhSecundario)
                                Call WriteVar(nPath, "OBJ" & OBJIndex, sObjStat(eObjStats.Texto), .Texto)
                                Call WriteVar(nPath, "OBJ" & OBJIndex, sObjStat(eObjStats.Agarrable), .Agarrable)
                                Call WriteVar(nPath, "OBJ" & OBJIndex, sObjStat(eObjStats.Info), .Info)
                                
                            Case eObjtype.Contenedores
                                Call WriteVar(nPath, "OBJ" & OBJIndex, sObjStat(eObjStats.grhIndex), .grhIndex)
                                Call WriteVar(nPath, "OBJ" & OBJIndex, sObjStat(eObjStats.Llave), .Llave)
                                Call WriteVar(nPath, "OBJ" & OBJIndex, sObjStat(eObjStats.Resistencia), .Resistencia)
                                Call WriteVar(nPath, "OBJ" & OBJIndex, sObjStat(eObjStats.CantItems), .MaxItems)
                                Call WriteVar(nPath, "OBJ" & OBJIndex, sObjStat(eObjStats.IndexAbierta), .IndexAbierta)
                                Call WriteVar(nPath, "OBJ" & OBJIndex, sObjStat(eObjStats.IndexCerrada), .IndexCerrada)
                                
                            Case eObjtype.Dinero
                                Call WriteVar(nPath, "OBJ" & OBJIndex, sObjStat(eObjStats.grhIndex), .grhIndex)
                                
                            Case eObjtype.Flechas
                                Call WriteVar(nPath, "OBJ" & OBJIndex, sObjStat(eObjStats.grhIndex), .grhIndex)
                                Call WriteVar(nPath, "OBJ" & OBJIndex, sObjStat(eObjStats.MinHit), .MinHit)
                                Call WriteVar(nPath, "OBJ" & OBJIndex, sObjStat(eObjStats.MaxHit), .MaxHit)
                                Call WriteVar(nPath, "OBJ" & OBJIndex, sObjStat(eObjStats.Valor), .Valor)
                                Call WriteVar(nPath, "OBJ" & OBJIndex, sObjStat(eObjStats.SkillCarpinteria), .SkCarpinteria)
                                Call WriteVar(nPath, "OBJ" & OBJIndex, sObjStat(eObjStats.Madera), .Madera)
                                Call WriteVar(nPath, "OBJ" & OBJIndex, sObjStat(eObjStats.Municion), .Municion)
                                
                            Case eObjtype.Fogata
                                Call WriteVar(nPath, "OBJ" & OBJIndex, sObjStat(eObjStats.grhIndex), .grhIndex)
                                Call WriteVar(nPath, "OBJ" & OBJIndex, sObjStat(eObjStats.Agarrable), .Agarrable)
                    
                            Case eObjtype.Foros
                                Call WriteVar(nPath, "OBJ" & OBJIndex, sObjStat(eObjStats.grhIndex), .grhIndex)
                                Call WriteVar(nPath, "OBJ" & OBJIndex, sObjStat(eObjStats.ForoID), .ForoID)
                                Call WriteVar(nPath, "OBJ" & OBJIndex, sObjStat(eObjStats.Agarrable), .Agarrable)
                                Call WriteVar(nPath, "OBJ" & OBJIndex, sObjStat(eObjStats.GrhSecundario), .GrhSecundario)
                                Call WriteVar(nPath, "OBJ" & OBJIndex, sObjStat(eObjStats.Info), .Info)
                                
                            Case eObjtype.Libros
                                 Call WriteVar(nPath, "OBJ" & OBJIndex, sObjStat(eObjStats.grhIndex), .grhIndex)
                                Call WriteVar(nPath, "OBJ" & OBJIndex, sObjStat(eObjStats.GrhSecundario), .GrhSecundario)
                                Call WriteVar(nPath, "OBJ" & OBJIndex, sObjStat(eObjStats.Texto), .Texto)
                                Call WriteVar(nPath, "OBJ" & OBJIndex, sObjStat(eObjStats.Info), .Info)
                                
                            Case eObjtype.Fragua
                                Call WriteVar(nPath, "OBJ" & OBJIndex, sObjStat(eObjStats.grhIndex), .grhIndex)
                                Call WriteVar(nPath, "OBJ" & OBJIndex, sObjStat(eObjStats.Info), .Info)
                                Call WriteVar(nPath, "OBJ" & OBJIndex, sObjStat(eObjStats.Agarrable), .Agarrable)
                                
                            Case eObjtype.Herramientas
                                Call WriteVar(nPath, "OBJ" & OBJIndex, sObjStat(eObjStats.grhIndex), .grhIndex)
                                Call WriteVar(nPath, "OBJ" & OBJIndex, sObjStat(eObjStats.Valor), .Valor)
                                Call WriteVar(nPath, "OBJ" & OBJIndex, sObjStat(eObjStats.Armaanim), .WeaponAnim)
                                Call WriteVar(nPath, "OBJ" & OBJIndex, sObjStat(eObjStats.Crucial), .Crucial)
                                Call WriteVar(nPath, "OBJ" & OBJIndex, sObjStat(eObjStats.SkillCarpinteria), .SkCarpinteria)
                                Call WriteVar(nPath, "OBJ" & OBJIndex, sObjStat(eObjStats.Madera), .Madera)
                                Call WriteVar(nPath, "OBJ" & OBJIndex, sObjStat(eObjStats.SkillHerreria), .SkHerreria)
                                Call WriteVar(nPath, "OBJ" & OBJIndex, sObjStat(eObjStats.LingotesHierro), .LingH)
                                
                            Case eObjtype.Instrumentos
                                Call WriteVar(nPath, "OBJ" & OBJIndex, sObjStat(eObjStats.grhIndex), .grhIndex)
                                Call WriteVar(nPath, "OBJ" & OBJIndex, sObjStat(eObjStats.Valor), .Valor)
                                Call WriteVar(nPath, "OBJ" & OBJIndex, sObjStat(eObjStats.Snd1), .Snd1)
                                Call WriteVar(nPath, "OBJ" & OBJIndex, sObjStat(eObjStats.Snd2), .Snd2)
                                Call WriteVar(nPath, "OBJ" & OBJIndex, sObjStat(eObjStats.Snd3), .Snd3)
                                Call WriteVar(nPath, "OBJ" & OBJIndex, sObjStat(eObjStats.SkillCarpinteria), .SkCarpinteria)
                                Call WriteVar(nPath, "OBJ" & OBJIndex, sObjStat(eObjStats.Madera), .Madera)
                                Call WriteVar(nPath, "OBJ" & OBJIndex, sObjStat(eObjStats.MinInt), .MinInt)
                            
                            Case eObjtype.Leña
                                Call WriteVar(nPath, "OBJ" & OBJIndex, sObjStat(eObjStats.grhIndex), .grhIndex)
                                Call WriteVar(nPath, "OBJ" & OBJIndex, sObjStat(eObjStats.Valor), .Valor)
                                Call WriteVar(nPath, "OBJ" & OBJIndex, sObjStat(eObjStats.Crucial), .Crucial)
                                
                            Case eObjtype.Llaves
                                Call WriteVar(nPath, "OBJ" & OBJIndex, sObjStat(eObjStats.grhIndex), .grhIndex)
                                Call WriteVar(nPath, "OBJ" & OBJIndex, sObjStat(eObjStats.Clave), .Clave)
                                Call WriteVar(nPath, "OBJ" & OBJIndex, sObjStat(eObjStats.Valor), .Valor)
                                Call WriteVar(nPath, "OBJ" & OBJIndex, sObjStat(eObjStats.Info), .Info)
                                Call WriteVar(nPath, "OBJ" & OBJIndex, sObjStat(eObjStats.NoSeCae), .NoSeCae)
                                
                            Case eObjtype.Manchas
                                Call WriteVar(nPath, "OBJ" & OBJIndex, sObjStat(eObjStats.grhIndex), .grhIndex)
                                Call WriteVar(nPath, "OBJ" & OBJIndex, sObjStat(eObjStats.Agarrable), .Agarrable)
                                
                            Case eObjtype.Minerales
                                Call WriteVar(nPath, "OBJ" & OBJIndex, sObjStat(eObjStats.grhIndex), .grhIndex)
                                Call WriteVar(nPath, "OBJ" & OBJIndex, sObjStat(eObjStats.Info), .Info)
                                Call WriteVar(nPath, "OBJ" & OBJIndex, sObjStat(eObjStats.MinSkill), .MinSkill)
                                Call WriteVar(nPath, "OBJ" & OBJIndex, sObjStat(eObjStats.LingoteIndex), .LingoteIndex)
                                Call WriteVar(nPath, "OBJ" & OBJIndex, sObjStat(eObjStats.Valor), .Valor)
                                
                            Case eObjtype.Pergaminos
                                Call WriteVar(nPath, "OBJ" & OBJIndex, sObjStat(eObjStats.grhIndex), .grhIndex)
                                Call WriteVar(nPath, "OBJ" & OBJIndex, sObjStat(eObjStats.Valor), .Valor)
                                Call WriteVar(nPath, "OBJ" & OBJIndex, sObjStat(eObjStats.Crucial), .Crucial)
                                Call WriteVar(nPath, "OBJ" & OBJIndex, sObjStat(eObjStats.HechizoIndex), .HechizoIndex)
                            
                            Case eObjtype.Piel
                                Call WriteVar(nPath, "OBJ" & OBJIndex, sObjStat(eObjStats.grhIndex), .grhIndex)
                                Call WriteVar(nPath, "OBJ" & OBJIndex, sObjStat(eObjStats.Valor), .Valor)
                                Call WriteVar(nPath, "OBJ" & OBJIndex, sObjStat(eObjStats.Crucial), .Crucial)
                        
                            Case eObjtype.pociones
                                Call WriteVar(nPath, "OBJ" & OBJIndex, sObjStat(eObjStats.grhIndex), .grhIndex)
                                Call WriteVar(nPath, "OBJ" & OBJIndex, sObjStat(eObjStats.Valor), .Valor)
                                Call WriteVar(nPath, "OBJ" & OBJIndex, sObjStat(eObjStats.TipoPocion), .TipoPocion)
                                Call WriteVar(nPath, "OBJ" & OBJIndex, sObjStat(eObjStats.MinModificador), .MinModificador)
                                Call WriteVar(nPath, "OBJ" & OBJIndex, sObjStat(eObjStats.MaxModificador), .MaxModificador)
                                Call WriteVar(nPath, "OBJ" & OBJIndex, sObjStat(eObjStats.DuracionEfecto), .DuracionEfecto)
                                Call WriteVar(nPath, "OBJ" & OBJIndex, sObjStat(eObjStats.Newbie), .Newbie)
                                
                            Case eObjtype.Puertas
                                Call WriteVar(nPath, "OBJ" & OBJIndex, sObjStat(eObjStats.grhIndex), .grhIndex)
                                Call WriteVar(nPath, "OBJ" & OBJIndex, sObjStat(eObjStats.Cerrada), .Cerrada)
                                Call WriteVar(nPath, "OBJ" & OBJIndex, sObjStat(eObjStats.Llave), .Llave)
                                Call WriteVar(nPath, "OBJ" & OBJIndex, sObjStat(eObjStats.Clave), .Clave)
                                Call WriteVar(nPath, "OBJ" & OBJIndex, sObjStat(eObjStats.Resistencia), .Resistencia)
                                Call WriteVar(nPath, "OBJ" & OBJIndex, sObjStat(eObjStats.Agarrable), .Agarrable)
                                Call WriteVar(nPath, "OBJ" & OBJIndex, sObjStat(eObjStats.IndexAbierta), .IndexAbierta)
                                Call WriteVar(nPath, "OBJ" & OBJIndex, sObjStat(eObjStats.IndexCerrada), .IndexCerrada)
                                Call WriteVar(nPath, "OBJ" & OBJIndex, sObjStat(eObjStats.IndexCerradaLlave), .IndexCerradaLlave)
                            
                            Case eObjtype.Teleports
                                Call WriteVar(nPath, "OBJ" & OBJIndex, sObjStat(eObjStats.grhIndex), .grhIndex)
                                Call WriteVar(nPath, "OBJ" & OBJIndex, sObjStat(eObjStats.Agarrable), .Agarrable)
                    
                            Case eObjtype.UseOnce
                                Call WriteVar(nPath, "OBJ" & OBJIndex, sObjStat(eObjStats.grhIndex), .grhIndex)
                                Call WriteVar(nPath, "OBJ" & OBJIndex, sObjStat(eObjStats.Minhambre), .MinHam)
                                Call WriteVar(nPath, "OBJ" & OBJIndex, sObjStat(eObjStats.Newbie), .Newbie)
                                Call WriteVar(nPath, "OBJ" & OBJIndex, sObjStat(eObjStats.Valor), .Valor)
                            Case eObjtype.Warp
                                Call WriteVar(nPath, "OBJ" & OBJIndex, sObjStat(eObjStats.grhIndex), .grhIndex)
                                Call WriteVar(nPath, "OBJ" & OBJIndex, sObjStat(eObjStats.Map_Pasajes), .WMapa)
                                Call WriteVar(nPath, "OBJ" & OBJIndex, sObjStat(eObjStats.MapX), .WX)
                                Call WriteVar(nPath, "OBJ" & OBJIndex, sObjStat(eObjStats.MapY), .WY)
                                Call WriteVar(nPath, "OBJ" & OBJIndex, sObjStat(eObjStats.WarpI), .WI)
                                Call WriteVar(nPath, "OBJ" & OBJIndex, sObjStat(eObjStats.Valor), .Valor)
                                
                            Case eObjtype.Yacimiento
                                Call WriteVar(nPath, "OBJ" & OBJIndex, sObjStat(eObjStats.grhIndex), .grhIndex)
                                Call WriteVar(nPath, "OBJ" & OBJIndex, sObjStat(eObjStats.Info), .Info)
                                Call WriteVar(nPath, "OBJ" & OBJIndex, sObjStat(eObjStats.Agarrable), .Agarrable)
                                Call WriteVar(nPath, "OBJ" & OBJIndex, sObjStat(eObjStats.MineralIndex), .MineralIndex)
                    
                            Case eObjtype.Yunque
                                Call WriteVar(nPath, "OBJ" & OBJIndex, sObjStat(eObjStats.grhIndex), .grhIndex)
                                Call WriteVar(nPath, "OBJ" & OBJIndex, sObjStat(eObjStats.Agarrable), .Agarrable)
                                Call WriteVar(nPath, "OBJ" & OBJIndex, sObjStat(eObjStats.Info), .Info)
                            
                            Case eObjtype.Amuleto
                                Call WriteVar(nPath, "OBJ" & OBJIndex, sObjStat(eObjStats.grhIndex), .grhIndex)
                                Call WriteVar(nPath, "OBJ" & OBJIndex, sObjStat(eObjStats.Def), .Def)
                                Call WriteVar(nPath, "OBJ" & OBJIndex, sObjStat(eObjStats.plusMagia), .plusMagia)
                                Call WriteVar(nPath, "OBJ" & OBJIndex, sObjStat(eObjStats.MinDef), .MinDef)
                                Call WriteVar(nPath, "OBJ" & OBJIndex, sObjStat(eObjStats.MaxDef), .MaxDef)
                                Call WriteVar(nPath, "OBJ" & OBJIndex, sObjStat(eObjStats.NoSeCae), .NoSeCae)
                                Call WriteVar(nPath, "OBJ" & OBJIndex, sObjStat(eObjStats.NoComerciable), .NoComerciable)
                                Call WriteVar(nPath, "OBJ" & OBJIndex, sObjStat(eObjStats.Aura), .Aura)
                                Call WriteVar(nPath, "OBJ" & OBJIndex, sObjStat(eObjStats.Remort), .Remort)
                                
                            Case eObjtype.Montura
                                Call WriteVar(nPath, "OBJ" & OBJIndex, sObjStat(eObjStats.grhIndex), .grhIndex)
                                Call WriteVar(nPath, "OBJ" & OBJIndex, sObjStat(eObjStats.MinDef), .MinDef)
                                Call WriteVar(nPath, "OBJ" & OBJIndex, sObjStat(eObjStats.MaxDef), .MaxDef)
                                Call WriteVar(nPath, "OBJ" & OBJIndex, sObjStat(eObjStats.Caballo), .Caballo)
                                Call WriteVar(nPath, "OBJ" & OBJIndex, sObjStat(eObjStats.Valor), .Valor)
                                Call WriteVar(nPath, "OBJ" & OBJIndex, sObjStat(eObjStats.Crucial), .Crucial)
                                Call WriteVar(nPath, "OBJ" & OBJIndex, sObjStat(eObjStats.NoSeCae), .NoSeCae)
                                Call WriteVar(nPath, "OBJ" & OBJIndex, sObjStat(eObjStats.Def), .Def)
                                Call WriteVar(nPath, "OBJ" & OBJIndex, sObjStat(eObjStats.plusMagia), .plusMagia)
                        End Select
                        .FueModificado = False
                    End If 'Solo si fue modificado
                End With
            Next OBJIndex
            
            AvisoStr = Left$(AvisoStr, LenB(AvisoStr) - 2)
            MsgBox AvisoStr
            If numChanges = 0 Then MsgBox "No se ha registrado ningun cambio hasta ahora en los objetos"
            DatNoGuardado(eModo.Objetos) = False
            
        Case eModo.Npc
            Dim noSobre As Boolean
            If DatNoGuardado(eModo.Npc) = False Then MsgBox "No se ha registrado ningun cambio hasta ahora en ningun NPC": Exit Sub
            If MsgBox("¿Deseas sobreescribir los cambios en '" & ConfigDir.Dats & "\NPCs.dat' y '" & ConfigDir.Dats & "\NPCs-Hostiles.dat'? (Si presionas NO, los cambios se guardaran en archivos adicionales con el nombre de 'NPCs1.dat' y 'NPCs-Hostiles1.dat')", vbYesNo, "Sobreescribir archivos") = vbYes Then
                noSobre = False
            Else
                noSobre = True
            End If
            AvisoStr = "Se han guardado los cambios de los siguientes NPCs: "
            
            Call WriteVar(ConfigDir.Dats & "/NPCs-HOSTILES.dat", "INIT", "NumNPCs", MaxNPC)
            
            For npcIndex = 1 To MaxNPC + 500
                If npcIndex < 500 And npcIndex > MaxNPCnohostiles Then GoTo sig ' NPCIndex
                
                With Npclist(npcIndex)
                    If noSobre Then
                        If npcIndex > 499 Then
                            nPath = ConfigDir.Dats & "/NPCs-HOSTILES1.dat"
                        Else
                            nPath = ConfigDir.Dats & "/NPCs1.dat"
                        End If
                    Else
                        If npcIndex > 499 Then
                            nPath = ConfigDir.Dats & "/NPCs-HOSTILES.dat"
                        Else
                            nPath = ConfigDir.Dats & "/NPCs.dat"
                        End If
                    End If
                    If .FueModificado = True Then
                        AvisoStr = AvisoStr & npcIndex & "(" & .name & ")" & ", "
                        numChanges = numChanges + 1
                        
                        WriteVar nPath, "NPC" & npcIndex, "Name", .name
                        If .hostilE = True Then
                        
                            WriteVar nPath, "NPC" & npcIndex, sNpcStat(eNPCStats.AguaValida), .flags.AguaValida
                            WriteVar nPath, "NPC" & npcIndex, sNpcStat(eNPCStats.TierraInvalida), .flags.TierraInvalida
                            WriteVar nPath, "NPC" & npcIndex, sNpcStat(eNPCStats.Alineacion), .Stats.Alineacion
                            WriteVar nPath, "NPC" & npcIndex, sNpcStat(eNPCStats.Atacable), .Attackable
                            WriteVar nPath, "NPC" & npcIndex, sNpcStat(eNPCStats.Body), .Char.Body
                            WriteVar nPath, "NPC" & npcIndex, sNpcStat(eNPCStats.Def), .Stats.Def
                            WriteVar nPath, "NPC" & npcIndex, sNpcStat(eNPCStats.Domable), .flags.Domable
                            WriteVar nPath, "NPC" & npcIndex, sNpcStat(eNPCStats.GiveEXP), .GiveEXP
                            WriteVar nPath, "NPC" & npcIndex, sNpcStat(eNPCStats.GiveGLD), .GiveGLD
                            WriteVar nPath, "NPC" & npcIndex, sNpcStat(eNPCStats.Head), .Char.Head
                            WriteVar nPath, "NPC" & npcIndex, sNpcStat(eNPCStats.Heading), .Char.Heading
                            WriteVar nPath, "NPC" & npcIndex, sNpcStat(eNPCStats.hostil), .hostilE
                            If .flags.LanzaSpells > 0 Then
                                WriteVar nPath, "NPC" & npcIndex, sNpcStat(eNPCStats.LanzaSpells), .flags.LanzaSpells 'asdasd
                                
                                For i = 1 To Npclist(npcIndex).flags.LanzaSpells
                                    Call WriteVar(nPath, "NPC" & npcIndex, "Sp" & i, Npclist(npcIndex).Spells(i))
                                Next
                            End If
                            WriteVar nPath, "NPC" & npcIndex, sNpcStat(eNPCStats.MinHP), .Stats.MinHP
                            WriteVar nPath, "NPC" & npcIndex, sNpcStat(eNPCStats.MaxHP), .Stats.MaxHP
                            WriteVar nPath, "NPC" & npcIndex, sNpcStat(eNPCStats.MinHit), .Stats.MinHit
                            WriteVar nPath, "NPC" & npcIndex, sNpcStat(eNPCStats.MaxHit), .Stats.MaxHit
                            WriteVar nPath, "NPC" & npcIndex, sNpcStat(eNPCStats.Movement), .Movement
                            WriteVar nPath, "NPC" & npcIndex, sNpcStat(eNPCStats.NPCtype), .NPCtype
                            
                           ' Obj1 = 1 - 100
                         '   Obj2 = 27 - 100
                           ' Obj3 = 158 - 100
                            WriteVar nPath, "NPC" & npcIndex, sNpcStat(eNPCStats.NroItems), .Invent.NroItems  'asdasd
                            If .Invent.NroItems > 0 Then
                                For i = 1 To .Invent.NroItems
                                    WriteVar nPath, "NPC" & npcIndex, "Obj" & i, .Invent.Object(i).OBJIndex & "-" & .Invent.Object(i).Amount
                                Next i
                            End If
                            WriteVar nPath, "NPC" & npcIndex, sNpcStat(eNPCStats.PocaParalisis), .flags.PocaParalisis
                            WriteVar nPath, "NPC" & npcIndex, sNpcStat(eNPCStats.PoderAtaque), .PoderAtaque
                            WriteVar nPath, "NPC" & npcIndex, sNpcStat(eNPCStats.PoderEvasion), .PoderEvasion
                            WriteVar nPath, "NPC" & npcIndex, sNpcStat(eNPCStats.Respawn), .flags.Respawn
                            WriteVar nPath, "NPC" & npcIndex, sNpcStat(eNPCStats.RespawnOrigPos), .flags.RespawnOrigPos
                            WriteVar nPath, "NPC" & npcIndex, sNpcStat(eNPCStats.Probabilidad), .Probabilidad
                            
                            WriteVar nPath, "NPC" & npcIndex, sNpcStat(eNPCStats.Snd1), .flags.Snd1
                            WriteVar nPath, "NPC" & npcIndex, sNpcStat(eNPCStats.Snd2), .flags.Snd2
                            WriteVar nPath, "NPC" & npcIndex, sNpcStat(eNPCStats.Snd3), .flags.Snd3
                            WriteVar nPath, "NPC" & npcIndex, sNpcStat(eNPCStats.Snd4), .flags.Snd4
                            WriteVar nPath, "NPC" & npcIndex, sNpcStat(eNPCStats.AutoCurar), .AutoCurar
                            WriteVar nPath, "NPC" & npcIndex, sNpcStat(eNPCStats.VeInvis), .VeInvis
                            WriteVar nPath, "NPC" & npcIndex, sNpcStat(eNPCStats.Veneno), .Veneno
                            WriteVar nPath, "NPC" & npcIndex, sNpcStat(eNPCStats.GolpeExacto), .flags.GolpeExacto
                        Else
                            WriteVar nPath, "NPC" & npcIndex, sNpcStat(eNPCStats.Body), .Char.Body
                            WriteVar nPath, "NPC" & npcIndex, sNpcStat(eNPCStats.Head), .Char.Head
                            WriteVar nPath, "NPC" & npcIndex, sNpcStat(eNPCStats.Heading), .Char.Heading
                            WriteVar nPath, "NPC" & npcIndex, sNpcStat(eNPCStats.Desc), .Desc
                            WriteVar nPath, "NPC" & npcIndex, sNpcStat(eNPCStats.Comercia), .Comercia
                            WriteVar nPath, "NPC" & npcIndex, sNpcStat(eNPCStats.AguaValida), .flags.AguaValida
                            WriteVar nPath, "NPC" & npcIndex, sNpcStat(eNPCStats.Alineacion), .Stats.Alineacion
                            WriteVar nPath, "NPC" & npcIndex, sNpcStat(eNPCStats.Atacable), .Attackable
                            WriteVar nPath, "NPC" & npcIndex, sNpcStat(eNPCStats.Def), .Stats.Def
                            WriteVar nPath, "NPC" & npcIndex, sNpcStat(eNPCStats.Domable), .flags.Domable
                            WriteVar nPath, "NPC" & npcIndex, sNpcStat(eNPCStats.Faccion), .flags.Faccion
                            WriteVar nPath, "NPC" & npcIndex, sNpcStat(eNPCStats.hostil), .hostilE
                            WriteVar nPath, "NPC" & npcIndex, sNpcStat(eNPCStats.Inflacion), .Inflacion
                            WriteVar nPath, "NPC" & npcIndex, sNpcStat(eNPCStats.InvReSpawn), .InvReSpawn
                            If .flags.LanzaSpells > 0 Then
                                WriteVar nPath, "NPC" & npcIndex, sNpcStat(eNPCStats.LanzaSpells), .flags.LanzaSpells 'asdasd
                                
                                For i = 1 To Npclist(npcIndex).flags.LanzaSpells
                                    Call WriteVar(nPath, "NPC" & npcIndex, "Sp" & i, Npclist(npcIndex).Spells(i))
                                Next
                            End If
                            WriteVar nPath, "NPC" & npcIndex, sNpcStat(eNPCStats.MaxHP), .Stats.MaxHP
                            WriteVar nPath, "NPC" & npcIndex, sNpcStat(eNPCStats.MinHP), .Stats.MinHP
                            WriteVar nPath, "NPC" & npcIndex, sNpcStat(eNPCStats.MaxHit), .Stats.MaxHit
                            WriteVar nPath, "NPC" & npcIndex, sNpcStat(eNPCStats.MinHit), .Stats.MinHit
                            WriteVar nPath, "NPC" & npcIndex, sNpcStat(eNPCStats.Movement), .Movement
                            WriteVar nPath, "NPC" & npcIndex, sNpcStat(eNPCStats.NPCtype), .NPCtype
                            WriteVar nPath, "NPC" & npcIndex, sNpcStat(eNPCStats.NroItems), .Invent.NroItems  'asdasd
                            If .Invent.NroItems > 0 Then
                                For i = 1 To .Invent.NroItems
                                    WriteVar nPath, "NPC" & npcIndex, "Obj" & i, .Invent.Object(i).OBJIndex & "-" & .Invent.Object(i).Amount
                                Next i
                            End If
                            WriteVar nPath, "NPC" & npcIndex, sNpcStat(eNPCStats.PoderAtaque), .PoderAtaque
                            WriteVar nPath, "NPC" & npcIndex, sNpcStat(eNPCStats.PoderEvasion), .PoderEvasion
                            WriteVar nPath, "NPC" & npcIndex, sNpcStat(eNPCStats.Respawn), .flags.Respawn
                            WriteVar nPath, "NPC" & npcIndex, sNpcStat(eNPCStats.RespawnOrigPos), .flags.RespawnOrigPos
                            WriteVar nPath, "NPC" & npcIndex, sNpcStat(eNPCStats.Snd1), .flags.Snd1
                            WriteVar nPath, "NPC" & npcIndex, sNpcStat(eNPCStats.Snd2), .flags.Snd2
                            WriteVar nPath, "NPC" & npcIndex, sNpcStat(eNPCStats.Snd3), .flags.Snd3
                            WriteVar nPath, "NPC" & npcIndex, sNpcStat(eNPCStats.Snd4), .flags.Snd4
                            WriteVar nPath, "NPC" & npcIndex, sNpcStat(eNPCStats.Sound), .flags.Sound
                            WriteVar nPath, "NPC" & npcIndex, sNpcStat(eNPCStats.TipoItems), .TipoItems
                        End If
                        .FueModificado = False
                    End If
                End With
sig:
            Next npcIndex
            AvisoStr = Left$(AvisoStr, LenB(AvisoStr) - 2)
            MsgBox AvisoStr
            If numChanges = 0 Then MsgBox "No se ha registrado ningun cambio hasta ahora en los NPCs"
            DatNoGuardado(eModo.Npc) = False
            
            
            
        Case eModo.Hechizo
            Dim Hechizo As Integer
            If DatNoGuardado(eModo.Hechizo) = False Then MsgBox "No se ha registrado ningun cambio hasta ahora en los objetos": Exit Sub
            If MsgBox("¿Deseas sobreescribir los cambios en '" & ConfigDir.Dats & "\OBJ.dat'? (Si presionas NO, los cambios se guardaran en un archivo adicional de nombre OBJ1.dat)", vbYesNo, "Sobreescribir archivo") = vbYes Then
                nPath = ConfigDir.Dats & "/Hechizos.dat"
            Else
                nPath = ConfigDir.Dats & "/Hechizos1.dat"
            End If
            AvisoStr = "Se han guardado los cambios de los siguientes hechizos: "
            Call WriteVar(nPath, "INIT", "NumeroHechizos", NumeroHechizos)
            
            For Hechizo = 1 To NumeroHechizos
                With Hechizos(Hechizo) '(OBJIndex)
                    If .FueModificado = True Then 'Solo guardamos los que fueron modificados
                        AvisoStr = AvisoStr & Hechizo & "(" & .Nombre & ")" & ", "
                        numChanges = numChanges + 1
                        Call WriteVar(nPath, "Hechizo" & Hechizo, "Nombre", .Nombre)
                        Call WriteVar(nPath, "Hechizo" & Hechizo, "Tipo", .Tipo)
                        Select Case .Tipo
                            Case 1 'HP, MANA,
                                WriteVar nPath, "Hechizo" & Hechizo, sHecStat(eHecStats.FXgrh), .FXgrh

                                WriteVar nPath, "Hechizo" & Hechizo, sHecStat(eHecStats.Desc), .Desc
                                WriteVar nPath, "Hechizo" & Hechizo, sHecStat(eHecStats.PalabrasMagicas), .PalabrasMagicas
                                WriteVar nPath, "Hechizo" & Hechizo, sHecStat(eHecStats.HechizeroMsg), .HechizeroMsg
                                WriteVar nPath, "Hechizo" & Hechizo, sHecStat(eHecStats.TargetMsg), .TargetMsg
                                WriteVar nPath, "Hechizo" & Hechizo, sHecStat(eHecStats.PropioMsg), .PropioMsg
                                WriteVar nPath, "Hechizo" & Hechizo, sHecStat(eHecStats.Target), .Target
                                WriteVar nPath, "Hechizo" & Hechizo, sHecStat(eHecStats.WAV), .WAV
                                WriteVar nPath, "Hechizo" & Hechizo, sHecStat(eHecStats.MinSkill), .MinSkill
                                WriteVar nPath, "Hechizo" & Hechizo, sHecStat(eHecStats.loops), .loops
                                
                                WriteVar nPath, "Hechizo" & Hechizo, sHecStat(eHecStats.SubeHP), .SubeHP
                                msgAyudaHechizos(eHecStats.SubeHP) = "1> Cura vida" & vbCrLf & "2> Saca vida"
                                WriteVar nPath, "Hechizo" & Hechizo, sHecStat(eHecStats.MinHP), .MinHP
                                WriteVar nPath, "Hechizo" & Hechizo, sHecStat(eHecStats.MaxHP), .MaxHP
                                
                                WriteVar nPath, "Hechizo" & Hechizo, sHecStat(eHecStats.SubeHam), .SubeHam
                                msgAyudaHechizos(eHecStats.SubeHam) = "2> Baja hambre del objetivo"
                                WriteVar nPath, "Hechizo" & Hechizo, sHecStat(eHecStats.MinHam), .MinHam
                                WriteVar nPath, "Hechizo" & Hechizo, sHecStat(eHecStats.MaxHam), .MaxHam
                                
                                WriteVar nPath, "Hechizo" & Hechizo, sHecStat(eHecStats.SubeAgilidad), .SubeAgilidad
                                msgAyudaHechizos(eHecStats.SubeAgilidad) = "1> Sube agilidad del objetivo" & vbCrLf & "2>Baja agilidad del objetivo"
                                WriteVar nPath, "Hechizo" & Hechizo, sHecStat(eHecStats.MinAgilidad), .MinAgilidad
                                WriteVar nPath, "Hechizo" & Hechizo, sHecStat(eHecStats.MaxAgilidad), .MaxAgilidad
                                
                                WriteVar nPath, "Hechizo" & Hechizo, sHecStat(eHecStats.SubeFuerza), .SubeFuerza
                                msgAyudaHechizos(eHecStats.SubeFuerza) = "1> Sube fuerza del objetivo" & vbCrLf & "2>Baja fuerza del objetivo"
                                WriteVar nPath, "Hechizo" & Hechizo, sHecStat(eHecStats.MinFuerza), .MinFuerza
                                WriteVar nPath, "Hechizo" & Hechizo, sHecStat(eHecStats.MaxFuerza), .MaxFuerza
                                
                                WriteVar nPath, "Hechizo" & Hechizo, sHecStat(eHecStats.ManaRequerido), .ManaRequerido
                                WriteVar nPath, "Hechizo" & Hechizo, sHecStat(eHecStats.StaRequerido), .StaRequerido
                                WriteVar nPath, "Hechizo" & Hechizo, sHecStat(eHecStats.Baculo), .Baculo
                                
                                WriteVar nPath, "Hechizo" & Hechizo, sHecStat(eHecStats.Nivel), .Nivel
                                msgAyudaHechizos(eHecStats.Nivel) = "Nivel minimo para usar el hechizo"
                                WriteVar nPath, "Hechizo" & Hechizo, sHecStat(eHecStats.Resis), .Resis
                                
                            Case 2 'estados del usuario
                                WriteVar nPath, "Hechizo" & Hechizo, sHecStat(eHecStats.FXgrh), .FXgrh
                                WriteVar nPath, "Hechizo" & Hechizo, sHecStat(eHecStats.Desc), .Desc
                                WriteVar nPath, "Hechizo" & Hechizo, sHecStat(eHecStats.PalabrasMagicas), .PalabrasMagicas
                                WriteVar nPath, "Hechizo" & Hechizo, sHecStat(eHecStats.HechizeroMsg), .HechizeroMsg
                                WriteVar nPath, "Hechizo" & Hechizo, sHecStat(eHecStats.TargetMsg), .TargetMsg
                                WriteVar nPath, "Hechizo" & Hechizo, sHecStat(eHecStats.PropioMsg), .PropioMsg
                                WriteVar nPath, "Hechizo" & Hechizo, sHecStat(eHecStats.Target), .Target
                                WriteVar nPath, "Hechizo" & Hechizo, sHecStat(eHecStats.WAV), .WAV
                                WriteVar nPath, "Hechizo" & Hechizo, sHecStat(eHecStats.MinSkill), .MinSkill
                                WriteVar nPath, "Hechizo" & Hechizo, sHecStat(eHecStats.loops), .loops
                                
                                WriteVar nPath, "Hechizo" & Hechizo, sHecStat(eHecStats.CuraVeneno), .CuraVeneno ' =1 cura veneno
                                WriteVar nPath, "Hechizo" & Hechizo, sHecStat(eHecStats.Envenena), .Envenena ' =2 envenena
                                WriteVar nPath, "Hechizo" & Hechizo, sHecStat(eHecStats.RemoverParalisis), .RemoverParalisis ' =1 remueva para
                                WriteVar nPath, "Hechizo" & Hechizo, sHecStat(eHecStats.Revivir), .Revivir ' =1 revive 2=Resucita
                                WriteVar nPath, "Hechizo" & Hechizo, sHecStat(eHecStats.Invisibilidad), .Invisibilidad ' =1 invi
                                WriteVar nPath, "Hechizo" & Hechizo, sHecStat(eHecStats.Paraliza), .Paraliza ' 2=inmo  1=paralizar
                                WriteVar nPath, "Hechizo" & Hechizo, sHecStat(eHecStats.Ceguera), .Ceguera ' =1 ceguera
                                WriteVar nPath, "Hechizo" & Hechizo, sHecStat(eHecStats.Estupidez), .Estupidez ' =1 estupidiza 2=Remmueve estupidez
                                WriteVar nPath, "Hechizo" & Hechizo, sHecStat(eHecStats.NoAtacar), .NoAtacar ' =1 no atacar
                                WriteVar nPath, "Hechizo" & Hechizo, sHecStat(eHecStats.Flecha), .Flecha ' =1 aumenta golpe con arco
                
                                WriteVar nPath, "Hechizo" & Hechizo, sHecStat(eHecStats.ManaRequerido), .ManaRequerido
                                WriteVar nPath, "Hechizo" & Hechizo, sHecStat(eHecStats.StaRequerido), .StaRequerido
                                WriteVar nPath, "Hechizo" & Hechizo, sHecStat(eHecStats.Baculo), .Baculo
                                
                                WriteVar nPath, "Hechizo" & Hechizo, sHecStat(eHecStats.Nivel), .Nivel
                                msgAyudaHechizos(eHecStats.Nivel) = "Nivel minimo para usar el hechizo"
                                WriteVar nPath, "Hechizo" & Hechizo, sHecStat(eHecStats.Resis), .Resis
                                
                            
                            Case 4 'invocaciones
                            
                                WriteVar nPath, "Hechizo" & Hechizo, sHecStat(eHecStats.FXgrh), .FXgrh
                                WriteVar nPath, "Hechizo" & Hechizo, sHecStat(eHecStats.Desc), .Desc
                                WriteVar nPath, "Hechizo" & Hechizo, sHecStat(eHecStats.PalabrasMagicas), .PalabrasMagicas
                                WriteVar nPath, "Hechizo" & Hechizo, sHecStat(eHecStats.HechizeroMsg), .HechizeroMsg
                                WriteVar nPath, "Hechizo" & Hechizo, sHecStat(eHecStats.Target), .Target
                                WriteVar nPath, "Hechizo" & Hechizo, sHecStat(eHecStats.WAV), .WAV
                                WriteVar nPath, "Hechizo" & Hechizo, sHecStat(eHecStats.MinSkill), .MinSkill
                                WriteVar nPath, "Hechizo" & Hechizo, sHecStat(eHecStats.loops), .loops
                                
                                WriteVar nPath, "Hechizo" & Hechizo, sHecStat(eHecStats.Invoca), .Invoca
                                WriteVar nPath, "Hechizo" & Hechizo, sHecStat(eHecStats.NumNPC), .NumNPC
                                WriteVar nPath, "Hechizo" & Hechizo, sHecStat(eHecStats.cant), .cant
                                
                                WriteVar nPath, "Hechizo" & Hechizo, sHecStat(eHecStats.ManaRequerido), .ManaRequerido
                                WriteVar nPath, "Hechizo" & Hechizo, sHecStat(eHecStats.StaRequerido), .StaRequerido
                                WriteVar nPath, "Hechizo" & Hechizo, sHecStat(eHecStats.Baculo), .Baculo
                                
                                WriteVar nPath, "Hechizo" & Hechizo, sHecStat(eHecStats.Nivel), .Nivel
                                msgAyudaHechizos(eHecStats.Nivel) = "Nivel minimo para usar el hechizo"
                                WriteVar nPath, "Hechizo" & Hechizo, sHecStat(eHecStats.Resis), .Resis
                                
                            
                            Case 6 'Hechizos de area
                                
                                WriteVar nPath, "Hechizo" & Hechizo, sHecStat(eHecStats.FXgrh), .FXgrh
                                WriteVar nPath, "Hechizo" & Hechizo, sHecStat(eHecStats.Desc), .Desc
                                WriteVar nPath, "Hechizo" & Hechizo, sHecStat(eHecStats.PalabrasMagicas), .PalabrasMagicas
                                WriteVar nPath, "Hechizo" & Hechizo, sHecStat(eHecStats.HechizeroMsg), .HechizeroMsg
                                WriteVar nPath, "Hechizo" & Hechizo, sHecStat(eHecStats.TargetMsg), .TargetMsg
                                WriteVar nPath, "Hechizo" & Hechizo, sHecStat(eHecStats.Target), .Target
                                
                                WriteVar nPath, "Hechizo" & Hechizo, sHecStat(eHecStats.WAV), .WAV
                                WriteVar nPath, "Hechizo" & Hechizo, sHecStat(eHecStats.MinSkill), .MinSkill
                                WriteVar nPath, "Hechizo" & Hechizo, sHecStat(eHecStats.loops), .loops
                                
                                WriteVar nPath, "Hechizo" & Hechizo, sHecStat(eHecStats.Invisibilidad), .Invisibilidad
                                WriteVar nPath, "Hechizo" & Hechizo, sHecStat(eHecStats.CuraArea), .CuraArea
                                
                                WriteVar nPath, "Hechizo" & Hechizo, sHecStat(eHecStats.ManaRequerido), .ManaRequerido
                                WriteVar nPath, "Hechizo" & Hechizo, sHecStat(eHecStats.StaRequerido), .StaRequerido
                                WriteVar nPath, "Hechizo" & Hechizo, sHecStat(eHecStats.Baculo), .Baculo
                                
                                WriteVar nPath, "Hechizo" & Hechizo, sHecStat(eHecStats.Nivel), .Nivel
                                msgAyudaHechizos(eHecStats.Nivel) = "Nivel minimo para usar el hechizo"
                                WriteVar nPath, "Hechizo" & Hechizo, sHecStat(eHecStats.Resis), .Resis
                        End Select
                        .FueModificado = False
                    End If 'Solo si fue modificado
                End With
            Next
            
            AvisoStr = Left$(AvisoStr, LenB(AvisoStr) - 2)
            MsgBox AvisoStr
            If numChanges = 0 Then MsgBox "No se ha registrado ningun cambio hasta ahora en los objetos"
            DatNoGuardado(eModo.Hechizo) = False
    End Select
End Sub

Private Sub Command4_Click()
GuardarEstadoActual
End Sub

Private Sub Command7_Click()
    If estadoDat = 0 Then MsgBox "Elige el modo de dat primero": Exit Sub
    
    picBuscar.Visible = Not picBuscar.Visible
    If picBuscar.Visible = False Then
        List1.Top = 1320
        List1.Height = 8250
    Else
        List1.Top = 1320 + picBuscar.Height
        List1.Height = 8250 - picBuscar.Height + 100
    End If
End Sub

Private Sub Command8_Click()
    Dim i As Long
    If optNombre.Value = True Then 'Buscar por nombre
        Select Case estadoDat
            Case eModo.Objetos
                List1.Clear
                For i = 1 To numObjs
                    If InStr(1, UCase$(ObjData(i).name), UCase$(txtBuscar.Text)) Then
                        List1.AddItem i & " - " & ObjData(i).name
                    End If
                Next i
                
            Case eModo.Npc
                List1.Clear
                For i = 1 To MaxNPC + 500
                    If InStr(1, UCase$(Npclist(i).name), UCase$(txtBuscar.Text)) Then
                        List1.AddItem i & " - " & Npclist(i).name
                    End If
                Next i
                
            Case eModo.Hechizo
                List1.Clear
                For i = 1 To NumeroHechizos
                    If InStr(1, UCase$(Hechizos(i).Nombre), UCase$(txtBuscar.Text)) Then
                        List1.AddItem i & " - " & Hechizos(i).Nombre
                    End If
                Next i
        End Select
    Else 'Buscar por objtype/npctype
        Select Case estadoDat
            Case eModo.Objetos
                List1.Clear
                For i = 1 To numObjs
                    If Val(txtBuscar.Text) = ObjData(i).objtype Then
                        List1.AddItem i & " - " & ObjData(i).name
                    End If
                Next i
                
            Case eModo.Npc
                List1.Clear
                For i = 1 To MaxNPC + 500
                    If Val(txtBuscar.Text) = Npclist(i).NPCtype Then
                        List1.AddItem i & " - " & Npclist(i).name
                    End If
                Next i
                                
            Case eModo.Hechizo
                List1.Clear
                For i = 1 To NumeroHechizos
                    If Val(txtBuscar.Text) = Hechizos(i).Tipo Then
                        List1.AddItem i & " - " & Hechizos(i).Nombre
                    End If
                Next i
        End Select
    End If
End Sub

Private Sub Command9_Click()
    Call CambiarEstadoDat(estadoDat)
   ' Select Case estadoDat
           ' Case eModo.Objetos
             '   List1.Clear
             '   Dim I As Long
             '   For I = 1 To numObjs
              '      If LenB(ObjData(I).name) <> 0 Then _
              '          List1.AddItem I & " - " & ObjData(I).name
              '  Next I
       ' End Select
        
        List1.Top = 1320
        List1.Height = 8250
        picBuscar.Visible = False
End Sub

Private Sub Form_Load()
    Dim X As Long
    For X = 0 To 95
        txtDatos(X).Text = ""
    Next X
    Combo1.Clear
    Combo1.AddItem "<<Elegir modo>>"
    Combo1.AddItem "Objetos"
    Combo1.AddItem "Npc's"
    Combo1.AddItem "Hechizos"

    Combo1.listIndex = 0
    
    cmbSubtipo.Clear
    cmbSubtipo.AddItem "0 - Armaduras"
    cmbSubtipo.AddItem "1 - Cascos"
    cmbSubtipo.AddItem "2 - Escudos"
End Sub


Private Sub Form_Terminate()
    cargado = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
    estadoDat = 0
End Sub


Private Sub Label1_Click(Index As Integer)

    Select Case estadoDat
        Case eModo.Objetos
            If LenB(msgAyudaObjs(Index)) > 0 Then _
                MsgBox msgAyudaObjs(Index)
                        
        Case eModo.Npc
            If LenB(msgAyudaNPCs(Index)) > 0 Then _
                MsgBox msgAyudaNPCs(Index)
                
        Case eModo.Hechizo
            If LenB(msgAyudaHechizos(Index)) > 0 Then _
                MsgBox msgAyudaHechizos(Index)
            
    End Select
End Sub

Private Sub List1_Click()
    On Local Error Resume Next
    Dim nINdex As Integer, tmpGrh As Integer
    nINdex = Val(ReadField(1, List1.List(List1.listIndex), Asc(" ")))
    CurrentIndex = nINdex
    Select Case estadoDat
        Case eModo.Objetos
            
            frmDats.picGrh.Cls
            If nINdex <= 0 Then Exit Sub
            ObjData(nINdex).Modificando = True
            Call SetInfoObj(nINdex)
            DibujarObjeto nINdex

            ObjData(nINdex).Modificando = False
            
        Case eModo.Npc
            frmDats.picInv.Visible = False
            frmDats.picGrh.Cls
            If nINdex = 0 Then Exit Sub

            Npclist(nINdex).Modificando = True
            SetInfoNPC nINdex
            

            Npclist(nINdex).Modificando = False
            'End If
            
        Case eModo.Hechizo
            frmDats.picInv.Visible = False
            frmDats.picGrh.Cls
            If nINdex = 0 Then Exit Sub

            Hechizos(nINdex).Modificando = True
            
            SetInfoHechizo nINdex
            

            Hechizos(nINdex).Modificando = False
            
    End Select
End Sub
Public Sub DibujarObjeto(ByVal OBJIndex As Integer)
    
    Dim i As Long, nGrhI As Integer, ln As String
    Dim auxr As RECT, auxr2 As RECT
    Dim X As Integer, Y As Integer
    
    If OBJIndex > numObjs Then Exit Sub
    With ObjData(OBJIndex)
        
        frmDats.picInv.Cls
        frmDats.picGrh.Cls
        BackBufferSurface.BltColorFill auxr, 0
        
        If .grhIndex > 0 Then
            If Grhdata(.grhIndex).pixelHeight = 32 And Grhdata(.grhIndex).pixelWidth = 32 Then
                frmDats.picInv.Visible = True
                Call dibujagrh(BackBufferSurface, Grhdata(.grhIndex), 0, 0)
                auxr.Left = 0
                auxr.Top = 0
                auxr.Bottom = 32
                auxr.Right = 32
                auxr2 = auxr
                BackBufferSurface.BltToDC frmDats.picInv.hdc, auxr2, auxr
                
                BackBufferSurface.BltColorFill auxr, 0
                frmDats.picInv.Refresh

                
            Else
                frmDats.picInv.Visible = False
                Call dibujarGrh4(Grhdata(.grhIndex), 0, 0)
                auxr.Left = 0
                auxr.Top = 0
                auxr.Bottom = Grhdata(.grhIndex).pixelHeight
                auxr.Right = Grhdata(.grhIndex).pixelWidth
                auxr2 = auxr
                BackBufferSurface.BltToDC frmDats.picGrh.hdc, auxr2, auxr
            End If
        End If
        
        'Surface.BltFast offsetX, offsetY, SurfaceDB.Surface(Grh.FileNum), r2, DDBLTFAST_SRCCOLORKEY Or DDBLTFAST_WAIT

        'BackBufferSurface.SetForeColor IIf(frmStatOptions.selectedslot > .Invent.NroItems, vbRed, vbGreen)
        'BackBufferSurface.setDrawStyle DrawStyleConstants.vbDashDot
        'BackBufferSurface.DrawBox x, y, x + 32, y + 32
        

    End With
    frmStatOptions.Picture1.Refresh
End Sub

Public Sub DibujarInventarioNPC(ByVal npcIndex As Integer)
    
    If frmStatOptions.Visible = False Then frmStatOptions.Show , Me
    Dim i As Long, nGrhI As Integer, ln As String
    Dim auxr As RECT, auxr2 As RECT
    Dim X As Integer, Y As Integer
    If npcIndex > (500 + MaxNPC) Then Exit Sub
    With Npclist(npcIndex)
        .Modificando = True
        If .Invent.NroItems <> 0 Then
            BackBufferSurface.BltColorFill auxr, 0
            For i = 1 To Npclist(npcIndex).Invent.NroItems
                nGrhI = 0

                If .Invent.Object(i).OBJIndex > 0 And .Invent.Object(i).OBJIndex <= numObjs Then _
                    nGrhI = ObjData(.Invent.Object(i).OBJIndex).grhIndex 'Val(GetVar(ConfigDir.Dats & "/obj.dat", "OBJ" & .Invent.Object(i).ObjIndex, "GrhIndex"))
                If nGrhI = 0 Then nGrhI = 1
                
                Call dibujaSlot(BackBufferSurface, Grhdata(nGrhI), CByte(i), .Invent.Object(i).Amount)
                
                If frmStatOptions.selectedslot = i Then
                    frmStatOptions.Text1.Text = .Invent.Object(i).OBJIndex
                    frmStatOptions.Text2.Text = .Invent.Object(i).Amount
                    frmStatOptions.Label2.Caption = ObjData(.Invent.Object(i).OBJIndex).name
                End If
            Next i
            
            GetSlotOffset frmStatOptions.selectedslot, X, Y
            BackBufferSurface.SetForeColor IIf(frmStatOptions.selectedslot > .Invent.NroItems, vbRed, vbGreen)
            BackBufferSurface.setDrawStyle DrawStyleConstants.vbDashDot
            BackBufferSurface.DrawBox X, Y, X + 32, Y + 32
            
            frmStatOptions.Label4.Caption = "Numitems: " & .Invent.NroItems
            auxr.Left = 0
            auxr.Top = 0
            auxr.Bottom = 32 * 4
            auxr.Right = 32 * 5
            auxr2 = auxr
            BackBufferSurface.BltToDC frmStatOptions.Picture1.hdc, auxr2, auxr
        End If
        .Modificando = False
    End With
    frmStatOptions.Picture1.Refresh
End Sub







Private Sub mnuGuardarTodo_Click()
    Dim antiguoEstado As Byte
    antiguoEstado = estadoDat
    
    estadoDat = eModo.Objetos
    GuardarEstadoActual
    estadoDat = eModo.Npc
    GuardarEstadoActual
    'estadoDat = eModo.Hechizo
    'GuardarEstadoActual
    
    estadoDat = antiguoEstado
End Sub

Private Sub Text1_Change()
    'nombre
    Select Case estadoDat
        Case eModo.Objetos
            If CurrentIndex > 0 And CurrentIndex <= numObjs Then
                With ObjData(CurrentIndex)
                    If .Modificando = False Then
                        .name = Text1.Text
                        List1.List(List1.listIndex) = CurrentIndex & " - " & Text1.Text
                        .FueModificado = True
                        DatNoGuardado(eModo.Objetos) = True
                    End If
                End With
            End If
        
        Case eModo.Npc
            If CurrentIndex > 0 And CurrentIndex <= (MaxNPC + 500) Then
                With Npclist(CurrentIndex)
                    If .Modificando = False Then
                        .name = Text1.Text
                        List1.List(List1.listIndex) = CurrentIndex & " - " & Text1.Text
                        .FueModificado = True
                        DatNoGuardado(eModo.Npc) = True
                    End If
                End With
            End If
            
        Case eModo.Hechizo
            If CurrentIndex > 0 And CurrentIndex <= NumeroHechizos Then
                With Hechizos(CurrentIndex)
                    If .Modificando = False Then
                        .Nombre = Text1.Text
                        List1.List(List1.listIndex) = CurrentIndex & " - " & Text1.Text
                        .FueModificado = True
                        DatNoGuardado(eModo.Hechizo) = True
                    End If
                End With
            End If
    End Select
End Sub

Private Sub Timer1_Timer()
    Dim i As Long, nGrhI As Integer, grhD As Integer
    Dim auxr As RECT, auxr2 As RECT
    Dim Y As Integer
    Dim tStr As String
    Static offsetx As Integer
    Static frame(3) As Byte
        If estadoDat > 0 And estadoDat < 4 Then
        If DatNoGuardado(estadoDat) = True Then
            lblEstado.Caption = "Estado: No guardado"
            lblEstado.ForeColor = vbRed
        Else
            lblEstado.Caption = "Estado: Guardado"
            lblEstado.ForeColor = vbGreen
        End If
    End If
    If Me.WindowState = vbMinimized Then Exit Sub
    
    If CurrentIndex <> 0 Then
        If estadoDat = eModo.Objetos Then
            With ObjData(CurrentIndex)
            
                If .objtype = 2 Or .objtype = 3 Then
                    picGrh.Cls
    
                    BackBufferSurface.BltColorFill auxr, 0
                    For i = 1 To 4
                        If .objtype = 2 Then 'armas
                            If .WeaponAnim > 0 Then nGrhI = (WeaponAnimData(.WeaponAnim).WeaponWalk(i).grhIndex) '.Frames(1)
                        ElseIf .objtype = 3 Then
                            If .SubTipo = 0 And .Ropaje > 0 And .Ropaje <= UBound(BodyData) Then nGrhI = (BodyData(.Ropaje).Walk(i).grhIndex) '.Frames(1)
                            If .SubTipo = 1 And .CascoAnim > 0 And .CascoAnim <= UBound(CascoAnimData) Then grhD = CascoAnimData(.CascoAnim).Head(i).grhIndex
                            If .SubTipo = 2 And .ShieldAnim > 0 And .ShieldAnim <= UBound(ShieldAnimData) Then nGrhI = (ShieldAnimData(.ShieldAnim).ShieldWalk(i).grhIndex) '.Frames(1)
                        End If
                        
                        If Not (.SubTipo = 1 And .objtype = 3) Then 'Si no es un casco
                            If nGrhI > 0 Then
                                If Grhdata(nGrhI).NumFrames > 0 Then
                                    frame(i - 1) = frame(i - 1) + 1
                                    If frame(i - 1) > Grhdata(nGrhI).NumFrames Then frame(i - 1) = 1
                                    
                                    
                                End If
                                grhD = Grhdata(nGrhI).Frames(frame(i - 1))
                            End If
                        End If
                        If i = 1 Then
                            tStr = "NORTE"
                        ElseIf i = 2 Then
                            tStr = "ESTE"
                        ElseIf i = 3 Then
                            tStr = "SUR"
                        ElseIf i = 4 Then
                            tStr = "OESTE"
                        End If
                        
                        If grhD <= 0 Then Exit Sub
                        
                        If (i > 2) And (Grhdata(grhD).pixelWidth * 2 + 20 > picGrh.ScaleWidth) Then
                            Call dibujagrh(BackBufferSurface, Grhdata(grhD), (i - 1) * (Grhdata(grhD).pixelWidth + 10), 10)
                            Call BackBufferSurface.DrawText(((i - 1) * (Grhdata(grhD).pixelWidth + 10)), 0, tStr, False)
                        Else
                        
                            Call dibujagrh(BackBufferSurface, Grhdata(grhD), (i - 1) * (Grhdata(grhD).pixelWidth + 10), 10)
                            Call BackBufferSurface.DrawText(((i - 1) * (Grhdata(grhD).pixelWidth + 10)), 0, tStr, False)
                        End If
                    Next i
                    
                    auxr.Left = 0
                    auxr.Top = 0
                    auxr.Bottom = Grhdata(grhD).pixelHeight + 10
                    auxr.Right = (4) * (Grhdata(grhD).pixelWidth + 10) ' Grhdata(nGrhI).pixelWidth
                    auxr2 = auxr
                    BackBufferSurface.BltToDC frmDats.picGrh.hdc, auxr2, auxr
                    picGrh.Refresh
                End If
            End With
        ElseIf estadoDat = eModo.Npc Then
            With Npclist(CurrentIndex)
                If .Char.Body > 0 Then
                    picGrh.Cls
    
                    BackBufferSurface.BltColorFill auxr, 0
                    For i = 1 To 4
                        If .Char.Body > UBound(BodyData) Then Exit Sub
                        nGrhI = (BodyData(.Char.Body).Walk(i).grhIndex) '.Frames(1)

                        If Grhdata(nGrhI).NumFrames > 0 Then
                            frame(i - 1) = frame(i - 1) + 1
                            If frame(i - 1) > Grhdata(nGrhI).NumFrames Then frame(i - 1) = 1
                            
                            
                        End If
                        grhD = Grhdata(nGrhI).Frames(frame(i - 1))
                        
                        
                        
                        If i = 1 Then
                            tStr = "NORTE"
                        ElseIf i = 2 Then
                            tStr = "ESTE"
                        ElseIf i = 3 Then
                            tStr = "SUR"
                        ElseIf i = 4 Then
                            tStr = "OESTE"
                        End If
                        
                        If grhD <= 0 Then Exit Sub
                        
                        If .Char.Head > 0 Then
                            Y = 10
                            If HeadData(.Char.Head).Head(i).grhIndex > 0 Then
                                Call dibujagrh(BackBufferSurface, Grhdata(HeadData(.Char.Head).Head(i).grhIndex), (i - 1) * (Grhdata(grhD).pixelWidth + 10) + 7, 10)
                            End If
                        End If
                        
                        If (i > 2) And (Grhdata(grhD).pixelWidth * 4 + 20 > picGrh.ScaleWidth) Then
                            Y = 10 + Grhdata(grhD).pixelHeight
                            Call dibujagrh(BackBufferSurface, Grhdata(grhD), (i - 3) * (Grhdata(grhD).pixelWidth + 10), Y + 10)
                            Call BackBufferSurface.DrawText(((i - 3) * (Grhdata(grhD).pixelWidth + 10)), Y, tStr, False)
                            'auxr.Left = 0
                            'auxr.Top = 0
                            'auxr.Bottom = Grhdata(grhD).pixelHeight + 25
                            'auxr.Right = (2) * (Grhdata(grhD).pixelWidth + 10) ' Grhdata(nGrhI).pixelWidth
                        Else
                            
                            Call dibujagrh(BackBufferSurface, Grhdata(grhD), offsetx, Y + 10) '(i - 1) * (Grhdata(grhD).pixelWidth + 10), y + 10)
                            Call BackBufferSurface.DrawText(offsetx, 0, tStr, False)
                            offsetx = offsetx + Grhdata(grhD).pixelWidth + 10
                           ' auxr.Left = 0
                           ' auxr.Top = 0
                           ' auxr.Bottom = Grhdata(grhD).pixelHeight + 25
                           ' auxr.Right = (4) * (Grhdata(grhD).pixelWidth + 10) ' Grhdata(nGrhI).pixelWidth
                        End If
                    
                    Next i
                    offsetx = 0
                    auxr.Left = 0
                    auxr.Top = 0
                    auxr.Bottom = picGrh.ScaleHeight 'Grhdata(grhD).pixelHeight + 25
                    auxr.Right = picGrh.ScaleWidth '(4) * (Grhdata(grhD).pixelWidth + 10) ' Grhdata(nGrhI).pixelWidth
                    
                    auxr2 = auxr
                    BackBufferSurface.BltToDC frmDats.picGrh.hdc, auxr2, auxr
                    picGrh.Refresh
                End If
            End With
        ElseIf estadoDat = eModo.Hechizo Then
            
                If CurrentIndex <= NumeroHechizos Then
                    With Hechizos(CurrentIndex)
                
                        picGrh.Cls
        
                        BackBufferSurface.BltColorFill auxr, 0
                            
                            If .FXgrh <= 0 Then Exit Sub
                            
                            nGrhI = FxData(.FXgrh).FX.grhIndex
                            
                            'If Not (.SubTipo = 1 And .objtype = 3) Then
                            If nGrhI > 0 Then
                                If Grhdata(nGrhI).NumFrames > 0 Then
                                    frame(1) = frame(1) + 1
                                    If frame(1) > Grhdata(nGrhI).NumFrames Then frame(1) = 1
                                    
                                    
                                End If
                                grhD = Grhdata(nGrhI).Frames(frame(1))
                            End If
                           ' End If
    
                            If grhD <= 0 Then Exit Sub
    
                            Call dibujagrh(BackBufferSurface, Grhdata(grhD), 0, 0)
                           ' Call BackBufferSurface.DrawText(((i - 1) * (Grhdata(grhD).pixelWidth + 10)), 0, tStr, False)
    
                        
                        auxr.Left = 0
                        auxr.Top = 0
                        auxr.Bottom = Grhdata(grhD).pixelHeight + 10
                        auxr.Right = (Grhdata(grhD).pixelWidth + 10)  ' Grhdata(nGrhI).pixelWidth
                        auxr2 = auxr
                        BackBufferSurface.BltToDC frmDats.picGrh.hdc, auxr2, auxr
                        picGrh.Refresh
                
                    End With
            End If
        End If
    End If
                

End Sub

Private Sub txtDatos_Change(Index As Integer)
    Dim nINdex As Integer
    Dim tmpGrh As Integer
    If List1.listIndex < 0 Then Exit Sub
    nINdex = Val(ReadField(1, List1.List(List1.listIndex), Asc(" ")))
    
    If nINdex <= 0 Then Exit Sub
    Select Case estadoDat
        Case eModo.Objetos
        
            If nINdex > numObjs Then Exit Sub
            
            If ObjData(nINdex).Modificando = False Then 'Esta siendo modificado por el usuario
                ObjData(nINdex).FueModificado = True
                GrabarDatos nINdex, Index
                DatNoGuardado(eModo.Objetos) = True
                
                If (Index + 1) = eObjStats.grhIndex Then
                    DibujarObjeto nINdex
                End If
            End If
            
            
        Case eModo.Npc
        
            'If nINdex > (500 + MaxNPC) Then Exit Sub
            
            If Npclist(nINdex).Modificando = False Then 'Esta siendo modificado por el usuario
                Npclist(nINdex).FueModificado = True
                GrabarDatos nINdex, Index
                DatNoGuardado(eModo.Npc) = True
            End If

        Case eModo.Hechizo
        
            If nINdex > NumeroHechizos Then Exit Sub
            
            If Hechizos(nINdex).Modificando = False Then 'Esta siendo modificado por el usuario
                Hechizos(nINdex).FueModificado = True
                GrabarDatos nINdex, Index
                DatNoGuardado(eModo.Hechizo) = True
                
                'If (Index + 1) = eHecStats.FXgrh Then
               '     If ObjData(nINdex).grhIndex > 0 Then
               '         frmDats.picGrh.Cls
               '         Call dibujarGrh4(Grhdata(ObjData(nINdex).grhIndex))
               '         frmDats.picGrh.Refresh
               '     End If
               ' End If
            End If
    End Select
End Sub
































Private Sub txtDatos_GotFocus(Index As Integer)
    txtDatos(Index).BackColor = RGB(200, 255, 200)
    If estadoDat = eModo.Npc And (Index + 1) = eNPCStats.Desc Then
        'txtDatos(Index).MultiLine = True
        txtDatos(Index).Height = 1200 'txtDatos(Index).Height * 2
        txtDatos(Index).Width = 6000 ' txtDatos(Index).Width * 2
        txtDatos(Index).ZOrder 0
        'txtDatos(Index).ScrollBars = 2
    End If
    
    If estadoDat = eModo.Objetos And (Index) = 21 Then 'eNPCStats.Desc Then
        'txtDatos(Index).MultiLine = True
        txtDatos(Index).Height = 1200 'txtDatos(Index).Height * 2
        txtDatos(Index).Width = 6000 ' txtDatos(Index).Width * 2
        txtDatos(Index).ZOrder 0
        'txtDatos(Index).ScrollBars = 2
    End If
    
    If estadoDat = eModo.Hechizo Then
        If (Index + 1) = eHecStats.Desc Or (Index + 1) = eHecStats.PalabrasMagicas Or _
                (Index + 1) = eHecStats.HechizeroMsg Or (Index + 1) = eHecStats.PropioMsg Or _
                (Index + 1) = eHecStats.TargetMsg Then
                
                
                txtDatos(Index).Height = 1200 'txtDatos(Index).Height * 2
                txtDatos(Index).Width = 6000 ' txtDatos(Index).Width * 2
                txtDatos(Index).ZOrder 0
        
        End If
    
    End If
End Sub

Private Sub txtDatos_LostFocus(Index As Integer)
    txtDatos(Index).BackColor = vbWhite
    If estadoDat = eModo.Npc And (Index + 1) = eNPCStats.Desc Then
        'txtDatos(Index).MultiLine = False
        txtDatos(Index).Height = 285 ' txtDatos(Index).Height / 2
        txtDatos(Index).Width = 1815 'txtDatos(Index).Width / 2
        'txtDatos(Index).ScrollBars = 0
    End If
    If estadoDat = eModo.Objetos And (Index) = 21 Then 'eNPCStats.Desc Then
        'txtDatos(Index).MultiLine = True
        txtDatos(Index).Height = 285 'txtDatos(Index).Height * 2
        txtDatos(Index).Width = 1815 ' txtDatos(Index).Width * 2

        'txtDatos(Index).ScrollBars = 2
    End If
    
        If estadoDat = eModo.Hechizo Then
        If (Index + 1) = eHecStats.Desc Or (Index + 1) = eHecStats.PalabrasMagicas Or _
                (Index + 1) = eHecStats.HechizeroMsg Or (Index + 1) = eHecStats.PropioMsg Or _
                (Index + 1) = eHecStats.TargetMsg Then
                
            txtDatos(Index).Height = 285 'txtDatos(Index).Height * 2
            txtDatos(Index).Width = 1815 ' txtDatos(Index).Width * 2
        
        End If
    
    End If
End Sub


Public Sub DatearClickDerecho(ByVal Tipo As eModo, ByVal eIndex As e_EstadoIndexador)
    
    'configDats.modo = eIndex
    'configDats.index = DataIndexActual
    estadoDat = Tipo
    configDats.modo = 1
    frmDats.Show , frmMain
    DoEvents
    Dim xx As Long, buscarslot As Byte
    Dim nINdex As Integer
    Dim nSlot As Integer
    Dim tINT As Integer, tI As Integer, lSlot As Integer
    Call frmDats.CambiarEstadoDat(estadoDat)
    DoEvents
    Select Case Tipo
        Case eModo.Objetos
            
            'ReDim Preserve ObjData(0 To numObjs) As ObjData
            
            'frmDats.List1.AddItem numObjs & " - Nuevo objeto"
            buscarslot = MsgBox("¿Deseas buscar un slot de OBJ.dat que este vacío?", vbYesNo, "Nuevo objeto")
            
            If buscarslot = vbYes Then
                
                For xx = 1 To numObjs
                    If ObjData(xx).name = "" And ObjData(xx).objtype = 0 And ObjData(xx).grhIndex = 0 Then 'Esta vacio
                        nINdex = xx
                        Exit For
                    End If
                Next xx
                
                
                If nINdex > 0 Then
                    frmDats.List1.listIndex = nINdex - 1
                    nSlot = nINdex
                Else
                    MsgBox "No se ha encontrado ningun slot vacio."
                    frmDats.List1.AddItem numObjs + 1 & " - Nuevo objeto"
                    numObjs = numObjs + 1
                    ReDim Preserve ObjData(0 To numObjs) As ObjData
                    'ReDim ObjData(0 To numObjs)
                    frmDats.List1.listIndex = frmDats.List1.ListCount - 1
                    nSlot = numObjs
                End If
            ElseIf buscarslot = vbNo Then
                frmDats.List1.AddItem numObjs + 1 & " - Nuevo objeto"
                numObjs = numObjs + 1
                ReDim Preserve ObjData(0 To numObjs) As ObjData
                nSlot = numObjs
                'ReDim ObjData(0 To numObjs)
                frmDats.List1.listIndex = frmDats.List1.ListCount - 1
            End If
            
            
            Select Case eIndex
            
                Case e_EstadoIndexador.Grh
                    ObjData(nSlot).grhIndex = GRHActual
                    ObjData(nSlot).name = "Nuevo objeto"
                    
                Case e_EstadoIndexador.Armas
                    ObjData(nSlot).objtype = 2
                    ObjData(nSlot).WeaponAnim = DataIndexActual
                    ObjData(nSlot).name = "Nueva arma"
                        
                Case e_EstadoIndexador.Body
                    ObjData(nSlot).objtype = 3
                    ObjData(nSlot).SubTipo = 0
                    ObjData(nSlot).Ropaje = DataIndexActual
                    ObjData(nSlot).name = "Nuevo ropaje/armadura"
                    
                Case e_EstadoIndexador.Cascos
                    ObjData(nSlot).objtype = 3
                    ObjData(nSlot).SubTipo = 1
                    ObjData(nSlot).CascoAnim = DataIndexActual
                    ObjData(nSlot).name = "Nuevo casco"
                    
                Case e_EstadoIndexador.Escudos
                    ObjData(nSlot).objtype = 3
                    ObjData(nSlot).SubTipo = 2
                    ObjData(nSlot).ShieldAnim = DataIndexActual
                    ObjData(nSlot).name = "Nuevo escudo"
                    
            End Select
            
            
            ObjData(nSlot).FueModificado = True
            DatNoGuardado(Tipo) = True
            
            frmDats.List1.listIndex = frmDats.List1.listIndex - 1
             frmDats.List1.listIndex = frmDats.List1.listIndex + 1
            frmDats.Text1.Text = ObjData(nSlot).name
        Case eModo.Npc
            Dim hostilE As Byte
            Dim i As Long
            hostilE = MsgBox("¿Este NPC será hostil? Si es hostil se guardara en el indice 'NPCs-HOSTILES.dat'; si NO es hostil en 'NPCs.dat'", vbYesNo, "Nuevo NPC")
            buscarslot = MsgBox("¿Deseas buscar un slot de NPC que este vacío?", vbYesNo, "Nuevo NPC")
            'buscarslot = vbNo
            
            If hostilE = vbYes Then
                hostilE = 1
            ElseIf hostilE = vbNo Then
                hostilE = 0
                
            End If
            
            If buscarslot = vbYes Then
                If hostilE = 1 Then
                    For i = 500 To 500 + MaxNPC
                        If Npclist(i).name = "" And Npclist(i).Char.Body = 0 Then 'Esta vacio
                            nINdex = i
                            Exit For
                        End If
                    Next i
                Else 'No es hostil
                    For i = 1 To MaxNPCnohostiles
                        If Npclist(i).name = "" And Npclist(i).Char.Body = 0 Then 'Esta vacio
                            nINdex = i
                            Exit For
                        End If
                    Next i
                End If
                If nINdex > 0 Then
                    
                    
                    For tINT = 0 To frmDats.List1.ListCount - 1
                        If Val(ReadField(1, frmDats.List1.List(tINT), Asc(" "))) = nINdex Then
                            lSlot = tINT
                            Exit For
                        End If
                    Next tINT
                    
                   nSlot = nINdex
                   Npclist(nSlot).Modificando = True
                   frmDats.List1.listIndex = lSlot
                    
                    'nINdex
                    'nSlot = nINdex
                    
                Else
                    If hostilE = 1 Then
                        MsgBox "No se ha encontrado ningun slot vacio. Se creará uno nuevo"
                        MaxNPC = MaxNPC + 1
                        nSlot = MaxNPC + 500
                        ReDim Preserve Npclist(1 To nSlot) As Npc
                        Npclist(nSlot).Modificando = True
                        
                        frmDats.List1.AddItem nSlot & " - Nuevo NPC"
                        frmDats.List1.listIndex = frmDats.List1.ListCount - 1
                        
                        
                        Npclist(nSlot).Char.Body = DataIndexActual
                        DatNoGuardado(eModo.Npc) = True
                        Npclist(nSlot).name = "Nuevo NPC"
                        SetInfoNPC nSlot
                        Npclist(nSlot).Modificando = False
                        Exit Sub
                    Else 'No es hostiles
                        
                        With frmDats.List1
                        'tint ti lslot
                            For tINT = 0 To .ListCount - 1
                                If Val(ReadField(1, .List(tINT), Asc(" "))) = 500 Then
                                    lSlot = tINT
                                    Exit For
                                End If
                            Next tINT
                            MaxNPCnohostiles = MaxNPCnohostiles + 1
                            MsgBox "No se ha encontrado ningun slot vacio. Se creará uno nuevo"
                            .AddItem .List(.ListCount - 1)
                            For tI = .ListCount - 2 To (lSlot + 1) Step -1
                                'Desde el primer npc hostil hasta el final de la lista
                                .List(tI) = .List(tI - 1)
                            Next tI
                            nSlot = MaxNPCnohostiles
                            Npclist(nSlot).Modificando = True
                            frmDats.List1.listIndex = lSlot

                        End With
                    End If
                End If
            ElseIf buscarslot = vbNo Then
                
               If hostilE = 1 Then
                    frmDats.List1.AddItem 500 + MaxNPC + 1 & " - Nuevo NPC(Hostil)"
                    MaxNPC = MaxNPC + 1
                    nSlot = MaxNPC + 500
                    ReDim Preserve Npclist(1 To 500 + MaxNPC) As Npc
                    Npclist(nSlot).Modificando = True
                    frmDats.List1.listIndex = frmDats.List1.ListCount - 1
                Else 'No es hostiles
                        
                    With frmDats.List1
                    'tint ti lslot
                        For tINT = 0 To .ListCount - 1
                            If Val(ReadField(1, .List(tINT), Asc(" "))) = 500 Then
                                lSlot = tINT
                                Exit For
                            End If
                        Next tINT
                        'MsgBox .ListCount
                        .AddItem (.List(.ListCount - 1))
                        For tI = .ListCount - 2 To (lSlot + 1) Step -1
                            'Desde el primer npc hostil hasta el final de la lista
                            .List(tI) = .List(tI - 1)
                        Next tI
                        MaxNPCnohostiles = MaxNPCnohostiles + 1
                        .List(lSlot) = MaxNPCnohostiles & " - Nuevo NPC"
                        nSlot = MaxNPCnohostiles
                        Npclist(nSlot).Modificando = True
                        frmDats.List1.listIndex = lSlot
                        
                    End With
                End If
            End If
            
            Npclist(nSlot).name = "Nuevo NPC"
            Npclist(nSlot).Char.Body = DataIndexActual
            Npclist(nSlot).FueModificado = True
            
            Npclist(nSlot).Modificando = True
            SetInfoNPC nSlot
            Npclist(nSlot).Modificando = False
            
            DatNoGuardado(Tipo) = True
            
            
        Case eModo.Hechizo
            buscarslot = MsgBox("¿Deseas buscar un slot de HECHIZOS.dat que este vacío?", vbYesNo, "Nuevo hechizo")
            
            If buscarslot = vbYes Then
                
                For xx = 1 To NumeroHechizos
                    If Hechizos(xx).Nombre = "" And Hechizos(xx).Tipo = 0 And Hechizos(xx).FXgrh = 0 Then 'Esta vacio
                        nINdex = xx
                        Exit For
                    End If
                Next xx
                
                
                If nINdex > 0 Then
                    frmDats.List1.listIndex = nINdex - 1
                    nSlot = nINdex
                Else
                    MsgBox "No se ha encontrado ningun slot vacio."
                    frmDats.List1.AddItem NumeroHechizos + 1 & " - Nuevo objeto"
                    NumeroHechizos = NumeroHechizos + 1
                    ReDim Preserve Hechizos(1 To NumeroHechizos) As tHechizo
                    'ReDim ObjData(0 To numObjs)
                    frmDats.List1.listIndex = frmDats.List1.ListCount - 1
                    nSlot = NumeroHechizos
                End If
            ElseIf buscarslot = vbNo Then
                frmDats.List1.AddItem NumeroHechizos + 1 & " - Nuevo hechizo"
                NumeroHechizos = NumeroHechizos + 1
                ReDim Preserve Hechizos(1 To NumeroHechizos) As tHechizo
                nSlot = NumeroHechizos
                'ReDim ObjData(0 To numObjs)
                frmDats.List1.listIndex = frmDats.List1.ListCount - 1
            End If
            Hechizos(nSlot).Nombre = "Nuevo hechizo"
            Hechizos(nSlot).FXgrh = DataIndexActual
            Hechizos(nSlot).FueModificado = True
            DatNoGuardado(Tipo) = True
            
            frmDats.cmbObjType.listIndex = 0
            frmDats.List1.listIndex = frmDats.List1.listIndex - 1
             frmDats.List1.listIndex = frmDats.List1.listIndex + 1
            
    End Select
    configDats.modo = 0
                'cmbObjType.AddItem "2 - ARMAS"
            'cmbObjType.AddItem "3 - ROPAS/CASCOS/ESCUDOS/ARMADURAS"
            
            'cmbSubtipo.AddItem "0 - Armaduras"
   ' cmbSubtipo.AddItem "1 - Cascos"
   ' cmbSubtipo.AddItem "2 - Escudos"
End Sub


