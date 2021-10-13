VERSION 5.00
Begin VB.Form frmLanzaSpells 
   BackColor       =   &H00000000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Clases prohibidas"
   ClientHeight    =   9765
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   5175
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9765
   ScaleWidth      =   5175
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command4 
      Caption         =   "Cancelar"
      Height          =   495
      Left            =   2640
      TabIndex        =   7
      Top             =   9120
      Width           =   2415
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Guardar"
      Height          =   495
      Left            =   120
      TabIndex        =   6
      Top             =   9120
      Width           =   2415
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Quitar hechizo"
      Height          =   495
      Left            =   120
      TabIndex        =   5
      Top             =   8400
      Width           =   4935
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Agregar hechizo"
      Height          =   495
      Left            =   120
      TabIndex        =   4
      Top             =   4800
      Width           =   4935
   End
   Begin VB.ListBox List2 
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   2400
      Left            =   120
      TabIndex        =   1
      Top             =   5880
      Width           =   4935
   End
   Begin VB.ListBox List1 
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   4350
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   4935
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Hechizos que lanza el NPC seleccionado"
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
      TabIndex        =   3
      Top             =   5640
      Width           =   3615
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Lista de hechizos"
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
Attribute VB_Name = "frmLanzaSpells"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


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
    With Npclist(frmDats.CurrentIndex)
        
        .flags.LanzaSpells = List2.ListCount
        If .flags.LanzaSpells > 0 Then
            ReDim .Spells(1 To .flags.LanzaSpells)
            For i = 1 To (List2.ListCount)
                Npclist(frmDats.CurrentIndex).Spells(i) = Val(ReadField(1, List2.List(i - 1), Asc(" ")))
            Next i
        
        End If
        
        frmDats.txtDatos(eNPCStats.LanzaSpells - 1).Text = .flags.LanzaSpells
        .FueModificado = True
        DatNoGuardado(eModo.Npc) = True
    End With
    Unload Me
End Sub

Private Sub Command4_Click()
    Unload Me
End Sub

Public Sub LoadLanzaSpells()
    Dim loopC As Integer
    Dim i As Long
    Dim nSpell As Integer
    
    List2.Clear
    List1.Clear
    
    For i = 1 To NumeroHechizos
        If Not IsLanzaSpell(i) Then
            If Hechizos(i).Nombre <> "" Then
            
                List1.AddItem i & " - " & Hechizos(i).Nombre
            End If
        End If
    Next i
    
    Dim npcIndex As Integer
    npcIndex = frmDats.CurrentIndex
    If Npclist(npcIndex).flags.LanzaSpells > 0 Then
        For loopC = 1 To Npclist(npcIndex).flags.LanzaSpells
            nSpell = Npclist(npcIndex).Spells(loopC)
            If nSpell > 0 And nSpell <= NumeroHechizos Then
                List2.AddItem nSpell & " - " & Hechizos(nSpell).Nombre
            End If
        Next loopC
    End If
End Sub

Private Function IsLanzaSpell(ByVal hIndex As Integer) As Boolean
Dim npcIndex As Integer
    npcIndex = frmDats.CurrentIndex
    If Npclist(npcIndex).flags.LanzaSpells <= 0 Then Exit Function
    Dim loopC As Integer
    For loopC = 1 To Npclist(npcIndex).flags.LanzaSpells
        If Npclist(npcIndex).Spells(loopC) = hIndex Then IsLanzaSpell = True: Exit Function
    Next loopC
    
    IsLanzaSpell = False
End Function

