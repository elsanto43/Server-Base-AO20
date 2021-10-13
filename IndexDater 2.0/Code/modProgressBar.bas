Attribute VB_Name = "modProgressBar"
Option Explicit

Private Type tBar
    Max As Long
    Value As Long
    Width As Long
End Type

Public ProgressBar(1) As tBar

Public Sub Init(ByVal Max As Long, Optional ByVal Width As Long = 465, Optional ByVal Index As Byte = 0)
    ProgressBar(Index).Max = Max
    ProgressBar(Index).Value = 0
    ProgressBar(Index).Width = Width
End Sub

Public Sub Restart(ByVal Max As Long, Optional ByVal Index As Byte = 0)
    ProgressBar(Index).Max = Max
    ProgressBar(Index).Value = 0
End Sub

Public Sub Update(ByVal Value As Long, Optional ByVal Index As Byte = 0)
    If Value > ProgressBar(Index).Max Then Value = ProgressBar(Index).Max
    ProgressBar(Index).Value = Value
    Select Case Index
        Case 0
            If ProgressBar(Index).Max > 0 Then _
            frmCargando.shpProgress.Width = Value / ProgressBar(Index).Max * ProgressBar(Index).Width
            DoEvents
    End Select
End Sub

