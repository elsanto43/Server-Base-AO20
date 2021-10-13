VERSION 5.00
Begin VB.Form frmStatOptions 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Inventario del npc"
   ClientHeight    =   7320
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   2700
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   488
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   180
   StartUpPosition =   1  'CenterOwner
   WhatsThisHelp   =   -1  'True
   Begin VB.CommandButton Command3 
      Caption         =   "Modificar objeto"
      Height          =   315
      Left            =   600
      TabIndex        =   9
      Top             =   4200
      Width           =   1455
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Borrar item"
      Height          =   375
      Left            =   120
      TabIndex        =   7
      Top             =   5520
      Width           =   2415
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Agregar item"
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   5040
      Width           =   2415
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   1320
      TabIndex        =   4
      Top             =   3840
      Width           =   855
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   1320
      TabIndex        =   2
      Top             =   3480
      Width           =   855
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   2895
      Left            =   120
      ScaleHeight     =   191
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   162
      TabIndex        =   0
      Top             =   120
      Width           =   2460
   End
   Begin VB.Label Label5 
      BackColor       =   &H00000000&
      Caption         =   $"frmStatOptions.frx":0000
      ForeColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   120
      TabIndex        =   10
      Top             =   6000
      Width           =   2415
   End
   Begin VB.Label Label4 
      Caption         =   "Numitems: "
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   4680
      Width           =   2415
   End
   Begin VB.Label Label2 
      BackColor       =   &H00000000&
      Caption         =   "Label2"
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   3120
      Width           =   2415
   End
   Begin VB.Label Label3 
      BackColor       =   &H00000000&
      Caption         =   "Cantidad:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   480
      TabIndex        =   3
      Top             =   3840
      Width           =   975
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      Caption         =   "ObjIndex: "
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   480
      TabIndex        =   1
      Top             =   3480
      Width           =   975
   End
End
Attribute VB_Name = "frmStatOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
    Dim Over As Boolean
Public selectedslot As Byte

Private Sub Command1_Click()
Dim nINdex As Integer
nINdex = frmDats.CurrentIndex
If nINdex > (500 + MaxNPC) Or nINdex <= 0 Then Exit Sub
If selectedslot = 0 Or selectedslot > 30 Then Exit Sub
With Npclist(nINdex)
    If .Invent.NroItems > 29 Then MsgBox "El npc no puede tener mas de 30 objetos": Exit Sub
    .Invent.NroItems = .Invent.NroItems + 1
    .Invent.Object(.Invent.NroItems).Amount = 0
    .Invent.Object(.Invent.NroItems).OBJIndex = 0
    selectedslot = .Invent.NroItems
    Npclist(nINdex).FueModificado = True
    DatNoGuardado(eModo.Npc) = True
    Call frmDats.DibujarInventarioNPC(nINdex)
    frmDats.txtDatos(eNPCStats.NroItems - 1).Text = .Invent.NroItems
End With

End Sub

Private Sub Command2_Click()
    Dim nINdex As Integer, loopC As Long
    nINdex = Val(ReadField(1, frmDats.List1.List(frmDats.List1.listIndex), Asc(" ")))
    If nINdex > (500 + MaxNPC) Or nINdex <= 0 Then Exit Sub

    With Npclist(nINdex)
        If selectedslot = 0 Or selectedslot > .Invent.NroItems Then Exit Sub
        If selectedslot <> .Invent.NroItems Then  ' no es el ultimo objeto, tenemos q reposicionar todos
            For loopC = selectedslot To .Invent.NroItems
                .Invent.Object(selectedslot) = .Invent.Object(selectedslot + 1)
            Next loopC
            '.Invent.Object(.Invent.NroItems).OBJIndex = 0
            '.Invent.Object(.Invent.NroItems).Amount = 0
        Else ' es el ultimo slot
            '.Invent.Object(selectedslot).OBJIndex = 0
            '.Invent.Object(selectedslot).Amount = 0
        End If
        .Invent.NroItems = .Invent.NroItems - 1

        .FueModificado = True
        DatNoGuardado(eModo.Npc) = True
        'selectedslot = .Invent.NroItems
        Call frmDats.DibujarInventarioNPC(nINdex)
        frmDats.txtDatos(eNPCStats.NroItems - 1).Text = .Invent.NroItems
    End With
End Sub

Private Sub Command3_Click()
    Dim nINdex As Integer
    nINdex = Val(ReadField(1, frmDats.List1.List(frmDats.List1.listIndex), Asc(" ")))
    If nINdex > (500 + MaxNPC) Or nINdex <= 0 Then Exit Sub
    If selectedslot = 0 Or selectedslot > 30 Then Exit Sub
    With Npclist(nINdex)
        .Invent.Object(selectedslot).OBJIndex = Val(Text1.Text)
        .Invent.Object(selectedslot).Amount = Val(Text2.Text)
        .FueModificado = True
        DatNoGuardado(eModo.Npc) = True
        Call frmDats.DibujarInventarioNPC(nINdex)
    End With
End Sub

Private Sub Form_Load()

 selectedslot = 1
End Sub


Private Sub Picture1_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim nINdex As Integer, npci As Integer
    npci = frmDats.CurrentIndex
    nINdex = ClickItem(x, y)
    If Not (nINdex > 30 Or nINdex = 0) Then
        'nINdex = Npclist(npci).Invent.NroItems
        frmStatOptions.selectedslot = nINdex
    End If
    Call frmDats.DibujarInventarioNPC(npci)
End Sub


Private Function ClickItem(ByVal x As Single, ByVal y As Single) As Integer
    Dim filaY As Byte
    Dim filaX As Byte
    If (y > 0 And y <= 32) Then
        filaY = 1
    ElseIf (y > 32 And y <= 64) Then
        filaY = 2
    ElseIf (y > 64 And y <= 96) Then
        filaY = 3
    ElseIf (y > 96 And y <= 128) Then
        filaY = 4
    End If
    
    If (x > 0 And x <= 32) Then
        filaX = 1
    ElseIf (x > 32 And x <= 64) Then
        filaX = 2
    ElseIf (x > 64 And x <= 96) Then
        filaX = 3
    ElseIf (x > 96 And x <= 128) Then
        filaX = 4
    ElseIf (x > 128) Then
        filaX = 5
    End If
    
    If filaY = 0 Then Exit Function
    If filaX = 0 Then Exit Function
    ClickItem = ((filaY - 1) * 5) + filaX
End Function

