VERSION 5.00
Begin VB.Form frmIniciarSubasta 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Iniciar subasta"
   ClientHeight    =   4605
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   4110
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4605
   ScaleWidth      =   4110
   StartUpPosition =   3  'Windows Default
   Begin AOLibre.uAOButton uAOButton1 
      Height          =   495
      Left            =   120
      TabIndex        =   5
      Top             =   3960
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   873
      TX              =   "Iniciar"
      ENAB            =   -1  'True
      FCOL            =   7314354
      OCOL            =   16777215
      PICE            =   "frmIniciarSubasta.frx":0000
      PICF            =   "frmIniciarSubasta.frx":001C
      PICH            =   "frmIniciarSubasta.frx":0038
      PICV            =   "frmIniciarSubasta.frx":0054
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   2040
      TabIndex        =   4
      Text            =   "10000"
      Top             =   3600
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   2040
      TabIndex        =   2
      Text            =   "1"
      Top             =   3240
      Width           =   615
   End
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   2955
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3855
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Oferta inicial"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   960
      TabIndex        =   3
      Top             =   3630
      Width           =   1095
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Cantidad"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1200
      TabIndex        =   1
      Top             =   3270
      Width           =   1095
   End
End
Attribute VB_Name = "frmIniciarSubasta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    Dim i As Long
    For i = 1 To Inventario.MaxObjs
        If Inventario.Amount(i) > 0 Then
            List1.AddItem Inventario.Amount(i) & " - " & Inventario.ItemName(i)
        Else
            List1.AddItem "<<<Vacio>>>"
        End If
    Next i
End Sub

Private Sub Text1_Change()
    If Val(text1.Text) > 10000 Then text1.Text = "10000"
    If Val(text1.Text) <= 1 Then text1.Text = "1"
End Sub

Private Sub Text2_Change()
    Text2.Text = Val(Text2.Text)
End Sub

Private Sub uAOButton1_Click()
    If List1.ListIndex >= 0 Then
        If MsgBox("Seguro que deseas subastar " & Val(text1.Text) & " - " & Inventario.ItemName(List1.ListIndex + 1) & " con una oferta inicial de: " & Val(Text2.Text), vbYesNo) = vbYes Then
            Call WriteIniciarSubasta(List1.ListIndex + 1, CInt(text1.Text), CLng(Text2.Text))
            Unload Me
        End If
    End If
End Sub
