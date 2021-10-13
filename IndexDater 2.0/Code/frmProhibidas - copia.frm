VERSION 5.00
Begin VB.Form frmProhibidas 
   BackColor       =   &H00000000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Clases prohibidas"
   ClientHeight    =   3720
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   5235
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3720
   ScaleWidth      =   5235
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command4 
      Caption         =   "Cancelar"
      Height          =   495
      Left            =   2640
      TabIndex        =   9
      Top             =   2880
      Width           =   2415
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Guardar"
      Height          =   495
      Left            =   120
      TabIndex        =   8
      Top             =   2880
      Width           =   2415
   End
   Begin VB.CommandButton Command2 
      Caption         =   "<-"
      Height          =   255
      Left            =   2400
      TabIndex        =   7
      Top             =   1800
      Width           =   375
   End
   Begin VB.CommandButton Command1 
      Caption         =   "->"
      Height          =   255
      Left            =   2400
      TabIndex        =   6
      Top             =   840
      Width           =   375
   End
   Begin VB.ListBox List2 
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   2400
      Left            =   2760
      TabIndex        =   3
      Top             =   360
      Width           =   2295
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   120
      TabIndex        =   2
      Top             =   3840
      Width           =   2295
   End
   Begin VB.ListBox List1 
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   2400
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   2295
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Permitidas"
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
      Left            =   2760
      TabIndex        =   5
      Top             =   120
      Width           =   2295
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Prohibidas"
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
      TabIndex        =   4
      Top             =   120
      Width           =   2295
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      Caption         =   "Num clases prohibidas: 25"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   2640
      TabIndex        =   1
      Top             =   3840
      Width           =   2295
   End
End
Attribute VB_Name = "frmProhibidas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public ModoProhibidos As eProModo
Private Sub Command1_Click()
    If List1.listIndex < 0 Then Exit Sub
    List2.AddItem List1.List(List1.listIndex)
    List2.listIndex = List2.ListCount - 1
    List1.RemoveItem List1.listIndex
End Sub

Private Sub Command2_Click()
    If List2.listIndex < 0 Then Exit Sub
    List1.AddItem List2.List(List2.listIndex)
    List1.listIndex = List1.ListCount - 1
    List2.RemoveItem List2.listIndex
End Sub

Private Sub Command3_Click()
    Dim i As Long, num As Byte
    Select Case ModoProhibidos
        Case eProModo.clases
            If List1.ListCount > 0 Then
                For i = 1 To (List1.ListCount)
                    ObjData(frmDats.CurrentIndex).ClaseProhibida(i) = Val(ReadField(1, List1.List(i - 1), Asc(" ")))
                    num = num + 1
                Next i
            End If
            For i = (List1.ListCount + 1) To NUMCLASES
                ObjData(frmDats.CurrentIndex).ClaseProhibida(i) = 0
            Next i
            
            ObjData(frmDats.CurrentIndex).FueModificado = True
            DatNoGuardado(eModo.Objetos) = True
        Case eProModo.razas
            For i = 1 To (List1.ListCount)
                ObjData(frmDats.CurrentIndex).RazaProhibida(i) = Val(ReadField(1, List1.List(i - 1), Asc(" ")))
            Next i
            For i = (List1.ListCount + 1) To NUMRAZAS
                ObjData(frmDats.CurrentIndex).RazaProhibida(i) = 0
            Next i
            
            ObjData(frmDats.CurrentIndex).FueModificado = True
            DatNoGuardado(eModo.Objetos) = True
    End Select
    Unload Me
End Sub

Private Sub Command4_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Dim i As Long
    Select Case ModoProhibidos
        Case eProModo.clases

            For i = 1 To NUMCLASES
                If ListaClases(i) <> "" Then
                    Combo1.AddItem i & " - " & ListaClases(i)
                End If
            Next i
                
        Case eProModo.razas

            
            For i = 1 To NUMRAZAS
                If ListaRazas(i) <> "" Then
                    Combo1.AddItem i & " - " & ListaRazas(i)
                End If
            Next i
    End Select
End Sub

Private Sub frmProhibidas_Click()

End Sub

