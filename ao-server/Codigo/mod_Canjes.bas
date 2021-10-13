Attribute VB_Name = "mod_Canjes"
Option Explicit

Private Type tCanje
    pideTrofeosOro As Byte
    pideTrofeosPlata As Byte
    pideTrofeosBronce As Byte
    
    recompensaItem As obj
End Type

Public Canjes() As tCanje
Public numCanjes As Integer

Public Const Trofeo_Oro As Integer = 222
Public Const Trofeo_Plata As Integer = 223
Public Const Trofeo_Bronce As Integer = 224

Public Sub CanjearItem(ByVal iCanje As Integer, ByVal ui As Integer)
    If iCanje < 1 Or iCanje > numCanjes Then Exit Sub
    With Canjes(iCanje)
        If .pideTrofeosOro > 0 Then
            If TieneObjetos(Trofeo_Oro, .pideTrofeosOro, ui) = False Then
                Call WriteConsoleMsg(ui, "No tienes suficientes trofeos de oro para canjear este objeto", FontTypeNames.FONTTYPE_GUILD)
                Exit Sub
            End If
        End If
        If .pideTrofeosPlata > 0 Then
            If TieneObjetos(Trofeo_Plata, .pideTrofeosPlata, ui) = False Then
                Call WriteConsoleMsg(ui, "No tienes suficientes trofeos de plata para canjear este objeto", FontTypeNames.FONTTYPE_GUILD)
                Exit Sub
            End If
        End If
        If .pideTrofeosBronce > 0 Then
            If TieneObjetos(Trofeo_Bronce, .pideTrofeosBronce, ui) = False Then
                Call WriteConsoleMsg(ui, "No tienes suficientes trofeos de bronce para canjear este objeto", FontTypeNames.FONTTYPE_GUILD)
                Exit Sub
            End If
        End If
        
        If Not MeterItemEnInventario(ui, .recompensaItem) Then
            Call WriteConsoleMsg(ui, "Has espacio en tu inventario para recibir el objeto primero", FontTypeNames.FONTTYPE_GUILD)
            Exit Sub
            'evitamos tirarle el objeto al piso
        End If
        
        If .pideTrofeosOro > 0 Then Call QuitarObjetos(Trofeo_Oro, .pideTrofeosOro, ui)
        If .pideTrofeosPlata > 0 Then Call QuitarObjetos(Trofeo_Plata, .pideTrofeosPlata, ui)
        If .pideTrofeosBronce > 0 Then Call QuitarObjetos(Trofeo_Bronce, .pideTrofeosBronce, ui)
        
        
    End With
End Sub

Public Sub CargarCanjes()
    Dim ler As clsIniManager
    Dim i As Long, buff As String
    
    Set ler = New clsIniManager
    Call ler.Initialize(DatPath & "Canjes.dat")
    
    numCanjes = val(ler.GetValue("INIT", "NumCanjes"))
    ReDim Canjes(1 To numCanjes) As tCanje
    For i = 1 To numCanjes
        With Canjes(i)
            .pideTrofeosOro = ler.GetValue("CANJE" & i, "TrofeosOro")
            .pideTrofeosPlata = ler.GetValue("CANJE" & i, "TrofeosPlata")
            .pideTrofeosBronce = ler.GetValue("CANJE" & i, "TrofeosBronce")
            
            buff = ler.GetValue("CANJE" & i, "Obj")
            
            .recompensaItem.ObjIndex = val(ReadField(1, buff, 45))
            .recompensaItem.Amount = val(ReadField(2, buff, 45))
        End With
    Next i
End Sub
