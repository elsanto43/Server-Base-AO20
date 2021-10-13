Attribute VB_Name = "GeneralES"
Option Explicit
'********************************************************************************
'********************************************************************************
'********************************************************************************
'*********************** Funciones de Carga *************************************
'********************************************************************************
'********************************************************************************
Public reConfigurarPath As Boolean
Public Numgrhs As Long
Public Sub LoadGrhData(Optional ByVal FileNamePath As String = vbNullString)
'*****************************************************************
'Loads Grh.dat
'*****************************************************************

On Error GoTo ErrorHandler

Dim Grh As Long
Dim frame As Integer
Dim TempInt As Integer
Dim ArchivoAbrir As String
Dim handle As Integer
Dim fileversion As Long
'modProgressBar.Restart  32000
'Resize arrays

'Open files

If FileNamePath = vbNullString Then
    If SavePath = 0 Then
        ArchivoAbrir = ConfigDir.Inits & "/Graficos.ind"
    Else
        ArchivoAbrir = ConfigDir.Inits & "/Graficos" & SavePath & ".ind"
    End If
Else
    ArchivoAbrir = FileNamePath
End If

If Not FileExist(ArchivoAbrir, vbNormal) Then
    MsgBox "Error al cargar: " & vbCrLf & ArchivoAbrir & vbCrLf & "El archivo no existe"
    reConfigurarPath = True
    Exit Sub
End If

handle = FreeFile

    Open ArchivoAbrir For Binary Access Read As handle
    Seek handle, 1

    Get handle, , fileversion

    Get handle, , Grh
    
    Numgrhs = Grh
    modProgressBar.Restart Grh
    
    ReDim Grhdata(0 To Grh) As Grhdata
    Dim iGrh As Long
    
While Not EOF(handle)
        Get handle, , iGrh
        If iGrh <> 0 Then
            With Grhdata(iGrh)
                'Get number of frames
                Get handle, , .NumFrames
                If .NumFrames <= 0 Then GoTo ErrorHandler
                
                ' = True
                
                
               ' ReDim .Frames(1 To .NumFrames)
                If .NumFrames > 1 Then
                
                    'Read a animation GRH set
                    For frame = 1 To .NumFrames
                        Get handle, , .Frames(frame)
                        If .Frames(frame) <= 0 Or .Frames(frame) > Grh Then
                            GoTo ErrorHandler
                        End If
                    Next frame
                    
                    Get handle, , .Speed
                    
                    If .Speed <= 0 Then GoTo ErrorHandler
                    
                    'Compute width and height
                    .pixelHeight = Grhdata(.Frames(1)).pixelHeight
                    If .pixelHeight <= 0 Then GoTo ErrorHandler
                    
                    .pixelWidth = Grhdata(.Frames(1)).pixelWidth
                    If .pixelWidth <= 0 Then GoTo ErrorHandler
                    
                    .TileWidth = Grhdata(.Frames(1)).TileWidth
                    If .TileWidth <= 0 Then GoTo ErrorHandler
                    
                    .TileHeight = Grhdata(.Frames(1)).TileHeight
                    If .TileHeight <= 0 Then GoTo ErrorHandler
                Else
                    'Read in normal GRH data
                    Get handle, , .FileNum
                    If .FileNum <= 0 Then GoTo ErrorHandler
                    
                    Get handle, , .sX
                    If .sX < 0 Then GoTo ErrorHandler
                    
                    Get handle, , .sY
                    If .sY < 0 Then GoTo ErrorHandler
                    
                    Get handle, , .pixelWidth
                    If .pixelWidth <= 0 Then GoTo ErrorHandler
                    
                    Get handle, , .pixelHeight
                    If .pixelHeight <= 0 Then GoTo ErrorHandler
                    
                    'Compute width and height
                    .TileWidth = .pixelWidth / TilePixelHeight
                    .TileHeight = .pixelHeight / TilePixelWidth
                    
                    .Frames(1) = iGrh
                End If
            End With
        End If
        
        modProgressBar.Update ProgressBar(0).Value + 1
    Wend
    
    

'************************************************

Close handle

Exit Sub

ErrorHandler:
Close handle
MsgBox "Error while loading the Grh.dat! Stopped at GRH number: " & Grh

End Sub


Public Sub CargarCuerposDat(Optional ByVal FileNamePath As String = vbNullString)
On Error Resume Next
Dim N As Integer, i As Integer
Dim NumCuerpos As Integer

Dim MisCuerpos() As tIndiceCuerpoLong
Dim ArchivoAbrir As String
Dim loopC As Integer


If FileNamePath = vbNullString Then
    If SavePath = 0 Then
        ArchivoAbrir = App.Path & "\encode\Body.dat"
    Else
        ArchivoAbrir = App.Path & "\encode\Body" & SavePath & ".dat"
    End If
Else
    ArchivoAbrir = FileNamePath
End If

If Not FileExist(ArchivoAbrir, vbNormal) Then
    MsgBox "Error al cargar: " & vbCrLf & ArchivoAbrir & vbCrLf & "El archivo no existe"
    reConfigurarPath = True
    Exit Sub
End If
Dim Leer As New clsIniReader

Call Leer.Initialize(ArchivoAbrir)

NumCuerpos = Val(Leer.GetValue("INIT", "NumBodies"))
ReDim MisCuerpos(0 To NumCuerpos + 1) As tIndiceCuerpoLong

ReDim BodyData(0 To NumCuerpos) As BodyData

For loopC = 1 To NumCuerpos
    InitGrh BodyData(loopC).Walk(1), Val(Leer.GetValue("Body" & loopC, "WALK1")), 0
    InitGrh BodyData(loopC).Walk(2), Val(Leer.GetValue("Body" & loopC, "WALK2")), 0
    InitGrh BodyData(loopC).Walk(3), Val(Leer.GetValue("Body" & loopC, "WALK3")), 0
    InitGrh BodyData(loopC).Walk(4), Val(Leer.GetValue("body" & loopC, "WALK4")), 0
    BodyData(loopC).HeadOffset.x = Val(Leer.GetValue("body" & loopC, "HeadOffsetX"))
    BodyData(loopC).HeadOffset.y = Val(Leer.GetValue("body" & loopC, "HeadOffsety"))
Next loopC
Set Leer = Nothing

End Sub


Public Sub CargarCabezasdat(Optional ByVal FileNamePath As String = vbNullString)
On Error Resume Next
Dim N As Integer, i As Integer, Index As Integer
Dim ArchivoAbrir As String
Dim loopC As Long

Dim Miscabezas() As tIndiceCabezaLong


If FileNamePath = vbNullString Then
    If SavePath = 0 Then
        ArchivoAbrir = App.Path & "\encode\Head.dat"
    Else
        ArchivoAbrir = App.Path & "\encode\Head" & SavePath & ".dat"
    End If
Else
    ArchivoAbrir = FileNamePath
End If

If Not FileExist(ArchivoAbrir, vbNormal) Then
    MsgBox "Error al cargar: " & vbCrLf & ArchivoAbrir & vbCrLf & "El archivo no existe"
    reConfigurarPath = True
    Exit Sub
End If

Dim Leer As New clsIniReader

Call Leer.Initialize(ArchivoAbrir)


Numheads = Val(Leer.GetValue("INIT", "NumHeads"))

ReDim HeadData(0 To Numheads) As HeadData

For i = 1 To Numheads
    InitGrh HeadData(i).Head(1), Val(Leer.GetValue("Head" & i, "Head1")), 0
    InitGrh HeadData(i).Head(2), Val(Leer.GetValue("Head" & i, "Head2")), 0
    InitGrh HeadData(i).Head(3), Val(Leer.GetValue("Head" & i, "Head3")), 0
    InitGrh HeadData(i).Head(4), Val(Leer.GetValue("Head" & i, "Head4")), 0
    DoEvents
    frmMain.LUlitError.Caption = "cabeza: " & i
Next i

Set Leer = Nothing
End Sub

Public Sub CargarCascosDat(Optional ByVal FileNamePath As String = vbNullString)
On Error Resume Next
Dim ArchivoAbrir As String
Dim N As Integer, i As Integer, NumCascos As Integer, Index As Integer

Dim Miscabezas() As tIndiceCabezaLong


If FileNamePath = vbNullString Then
    If SavePath = 0 Then
        ArchivoAbrir = App.Path & "\encode\Cascos.dat"
    Else
        ArchivoAbrir = App.Path & "\encode\Cascos" & SavePath & ".dat"
    End If
Else
    ArchivoAbrir = FileNamePath
End If

If Not FileExist(ArchivoAbrir, vbNormal) Then
    reConfigurarPath = True
    MsgBox "Error al cargar: " & vbCrLf & ArchivoAbrir & vbCrLf & "El archivo no existe"
    Exit Sub
End If


NumCascos = Val(GetVar(ArchivoAbrir, "INIT", "NumCascos"))
'Resize array
ReDim CascoAnimData(0 To NumCascos) As HeadData

For i = 1 To NumCascos
    InitGrh CascoAnimData(i).Head(1), Val(GetVar(ArchivoAbrir, "Casco" & i, "Head1")), 0
    InitGrh CascoAnimData(i).Head(2), Val(GetVar(ArchivoAbrir, "Casco" & i, "Head2")), 0
    InitGrh CascoAnimData(i).Head(3), Val(GetVar(ArchivoAbrir, "Casco" & i, "Head3")), 0
    InitGrh CascoAnimData(i).Head(4), Val(GetVar(ArchivoAbrir, "Casco" & i, "Head4")), 0
Next i


End Sub


Public Sub CargarCabezas(Optional ByVal FileNamePath As String = vbNullString)
On Error Resume Next
Dim N As Integer, i As Integer, Index As Integer
Dim ArchivoAbrir As String

Dim Miscabezas() As tIndiceCabeza
Dim MiscabezasLong() As tIndiceCabezaLong

If FileNamePath = vbNullString Then
    If SavePath = 0 Then
        ArchivoAbrir = ConfigDir.Inits & "\Cabezas.ind"
    Else
        ArchivoAbrir = ConfigDir.Inits & "\Cabezas" & SavePath & ".ind"
    End If
Else
    ArchivoAbrir = FileNamePath
End If

If Not FileExist(ArchivoAbrir, vbNormal) Then
    reConfigurarPath = True '
    MsgBox "Error al cargar: " & vbCrLf & ArchivoAbrir & vbCrLf & "El archivo no existe"
'    If UBound(HeadData()) = 0 Then
  '      ReDim HeadData(1) As HeadData
  '  End If
    
    Exit Sub
End If

N = FreeFile
Open ArchivoAbrir For Binary Access Read As #N

'cabecera
Get #N, , MiCabecera

'num de cabezas
Get #N, , Numheads

modProgressBar.Restart Numheads

'Resize array
ReDim HeadData(0 To Numheads) As HeadData
ReDim Miscabezas(0 To Numheads + 1) As tIndiceCabeza
ReDim MiscabezasLong(0 To Numheads + 1) As tIndiceCabezaLong

'If UsarGrhLong Then
    For i = 1 To Numheads
        Get #N, , MiscabezasLong(i)
        InitGrh HeadData(i).Head(1), MiscabezasLong(i).Head(1), 0
        InitGrh HeadData(i).Head(2), MiscabezasLong(i).Head(2), 0
        InitGrh HeadData(i).Head(3), MiscabezasLong(i).Head(3), 0
        InitGrh HeadData(i).Head(4), MiscabezasLong(i).Head(4), 0
        modProgressBar.Update ProgressBar(0).Value + 1
    Next i
'Else
'    For i = 1 To Numheads
  '      Get #N, , Miscabezas(i)
 '       InitGrh HeadData(i).Head(1), Miscabezas(i).Head(1), 0
 '       InitGrh HeadData(i).Head(2), Miscabezas(i).Head(2), 0
  '      InitGrh HeadData(i).Head(3), Miscabezas(i).Head(3), 0
 '       InitGrh HeadData(i).Head(4), Miscabezas(i).Head(4), 0
    '    modProgressBar.Update ProgressBar(0).Value + 1
 '   Next i

'End If

Close #N

End Sub

Public Sub CargarCascos(Optional ByVal FileNamePath As String = vbNullString)
On Error Resume Next
Dim ArchivoAbrir As String
Dim N As Integer, i As Integer, NumCascos As Integer, Index As Integer

Dim Miscabezas() As tIndiceCabeza
Dim MiscabezasLong() As tIndiceCabezaLong

If FileNamePath = vbNullString Then
    If SavePath = 0 Then
        ArchivoAbrir = ConfigDir.Inits & "\Cascos.ind"
    Else
        ArchivoAbrir = ConfigDir.Inits & "\Cascos" & SavePath & ".ind"
    End If
Else
    ArchivoAbrir = FileNamePath
End If

If Not FileExist(ArchivoAbrir, vbNormal) Then
    reConfigurarPath = True '
    MsgBox "Error al cargar: " & vbCrLf & ArchivoAbrir & vbCrLf & "El archivo no existe"
'    If UBound(CascoAnimData()) = 0 Then
  '      ReDim CascoAnimData(1) As HeadData
 '   End If
    Exit Sub
End If
N = FreeFile
Open ArchivoAbrir For Binary Access Read As #N

'cabecera
Get #N, , MiCabecera

'num de cabezas
Get #N, , NumCascos

'Resize arra
ReDim CascoAnimData(0 To NumCascos + 1) As HeadData
ReDim Miscabezas(0 To NumCascos + 1) As tIndiceCabeza
ReDim MiscabezasLong(0 To Numheads + 1) As tIndiceCabezaLong

'If UsarGrhLong Then
    For i = 1 To NumCascos
        Get #N, , MiscabezasLong(i)
        InitGrh CascoAnimData(i).Head(1), MiscabezasLong(i).Head(1), 0
        InitGrh CascoAnimData(i).Head(2), MiscabezasLong(i).Head(2), 0
        InitGrh CascoAnimData(i).Head(3), MiscabezasLong(i).Head(3), 0
        InitGrh CascoAnimData(i).Head(4), MiscabezasLong(i).Head(4), 0
    Next i
'Else
 '   For i = 1 To NumCascos
 '       Get #N, , Miscabezas(i)
 '       InitGrh CascoAnimData(i).Head(1), Miscabezas(i).Head(1), 0
  '      InitGrh CascoAnimData(i).Head(2), Miscabezas(i).Head(2), 0
 '       InitGrh CascoAnimData(i).Head(3), Miscabezas(i).Head(3), 0
 '       InitGrh CascoAnimData(i).Head(4), Miscabezas(i).Head(4), 0
 '   Next i
'End If

Close #N

End Sub
Public Sub CargarSuperficies()
    Dim nPath As String
    Dim i As Long
    nPath = ConfigDir.InitWE & "\indices.ini"
    
    If Not FileExist(nPath, vbNormal) Then
        reConfigurarPath = True '
        MsgBox "Error al cargar: " & vbCrLf & nPath & vbCrLf & "El archivo no existe"
    End If

    
    NumSuperficies = Val(GetVar(nPath, "INIT", "Referencias"))
    modProgressBar.Restart NumSuperficies + 2


    ReDim SupData(0 To NumSuperficies)
    For i = 0 To NumSuperficies
        With SupData(i)
            .Nombre = GetVar(nPath, "REFERENCIA" & i, "Nombre")
            .GrhIndex = Val(GetVar(nPath, "REFERENCIA" & i, "GrhIndice"))
            .Ancho = Val(GetVar(nPath, "REFERENCIA" & i, "Ancho"))
            .Alto = Val(GetVar(nPath, "REFERENCIA" & i, "Alto"))
            .Capa = Val(GetVar(nPath, "REFERENCIA" & i, "Capa"))
        End With
        
        modProgressBar.Update ProgressBar(0).Value + 1
    Next i
End Sub

Public Sub guardarSuperficie(ByVal Index As Integer)
    Call WriteVar(ConfigDir.InitWE & "\indices.ini", "INIT", "Referencias", NumSuperficies)
    Call WriteVar(ConfigDir.InitWE & "\indices.ini", "REFERENCIA" & Index, "Nombre", SupData(Index).Nombre)
    Call WriteVar(ConfigDir.InitWE & "\indices.ini", "REFERENCIA" & Index, "GrhIndice", SupData(Index).GrhIndex)
    Call WriteVar(ConfigDir.InitWE & "\indices.ini", "REFERENCIA" & Index, "Ancho", SupData(Index).Ancho)
    Call WriteVar(ConfigDir.InitWE & "\indices.ini", "REFERENCIA" & Index, "Alto", SupData(Index).Alto)
    Call WriteVar(ConfigDir.InitWE & "\indices.ini", "REFERENCIA" & Index, "Capa", SupData(Index).Capa)
End Sub


Sub CargarCuerpos(Optional ByVal FileNamePath As String = vbNullString)
On Error Resume Next
Dim N As Integer, i As Integer
Dim NumCuerpos As Integer
Dim MisCuerpos() As tIndiceCuerpoLong
Dim ArchivoAbrir As String

N = FreeFile



If FileNamePath = vbNullString Then
    If SavePath = 0 Then
        ArchivoAbrir = ConfigDir.Inits & "\Personajes.ind"
    Else
        ArchivoAbrir = ConfigDir.Inits & "\Personajes" & SavePath & ".ind"
    End If
Else
    ArchivoAbrir = FileNamePath
End If

If Not FileExist(ArchivoAbrir, vbNormal) Then
   reConfigurarPath = True '
   MsgBox "Error al cargar: " & vbCrLf & ArchivoAbrir & vbCrLf & "El archivo no existe"
    'If UBound(BodyData()) = 0 Then
   '     ReDim BodyData(1) As BodyData
    'End If
    Exit Sub
End If
Open ArchivoAbrir For Binary Access Read As #N

'cabecera
Get #N, , MiCabecera

'num de cabezas
Get #N, , NumCuerpos

modProgressBar.Restart NumCuerpos

'Resize array
ReDim BodyData(0 To NumCuerpos) As BodyData
ReDim MisCuerpos(0 To NumCuerpos + 1) As tIndiceCuerpoLong

    For i = 1 To NumCuerpos
        Get #N, , MisCuerpos(i)
        InitGrh BodyData(i).Walk(1), MisCuerpos(i).Body(1), 0
        InitGrh BodyData(i).Walk(2), MisCuerpos(i).Body(2), 0
        InitGrh BodyData(i).Walk(3), MisCuerpos(i).Body(3), 0
        InitGrh BodyData(i).Walk(4), MisCuerpos(i).Body(4), 0
        BodyData(i).HeadOffset.x = MisCuerpos(i).HeadOffsetX
        BodyData(i).HeadOffset.y = MisCuerpos(i).HeadOffsetY
        modProgressBar.Update ProgressBar(0).Value + 1

    Next i


Close #N

End Sub



Public Sub CargarFxs(Optional ByVal FileNamePath As String = vbNullString)
    On Error Resume Next
    
    Dim N As Integer, i As Integer
    Dim ArchivoAbrir As String
    N = FreeFile

    Dim FileManager As clsIniReader
    
    If FileNamePath = vbNullString Then
        If SavePath = 0 Then
            ArchivoAbrir = ConfigDir.Inits & "\Fxs.ini"
        Else
            ArchivoAbrir = ConfigDir.Inits & "\Fxs" & SavePath & ".ind"
        End If
    Else
        ArchivoAbrir = FileNamePath
    End If
    
    If Not FileExist(ArchivoAbrir, vbNormal) Then
        reConfigurarPath = True '
        MsgBox "Error al cargar: " & vbCrLf & ArchivoAbrir & vbCrLf & "El archivo no existe"
        Exit Sub
    End If


    Set FileManager = New clsIniReader
    Call FileManager.Initialize(ArchivoAbrir)
    
    'Resize array
    ReDim FxData(0 To FileManager.GetValue("INIT", "NumFxs")) As FxData
    
    For i = 1 To numfxs
        With FxData(i)
            Call InitGrh(.FX, Val(FileManager.GetValue("FX" & CStr(i), "Animacion")))
            .offsetx = Val(FileManager.GetValue("FX" & CStr(i), "OffsetX"))
            .offsety = Val(FileManager.GetValue("FX" & CStr(i), "OffsetY"))
        End With
        'FxData(i).OffsetX = MisFxs(i).OffsetX
       ' FxData(i).offsety = MisFxs(i).offsety
        modProgressBar.Update ProgressBar(0).Value + 1


    Next i


End Sub

'********************************************************************************
'********************************************************************************
'********************************************************************************
'********************** Funciones de guardado ***********************************
'********************************************************************************
'********************************************************************************

Public Sub GuardarCabezas(Optional ByVal FileNamePath As String = vbNullString)
On Error GoTo errhandler

Dim N As Integer, i As Integer
N = FreeFile
Dim ArchivoAbrir As String
If FileNamePath = vbNullString Then
    If SavePath = 0 Then
        ArchivoAbrir = ConfigDir.Inits & "\Cabezas.ind"
    Else
        ArchivoAbrir = ConfigDir.Inits & "\Cabezas" & SavePath & ".ind"
    End If
Else
    ArchivoAbrir = FileNamePath
End If

If Not ComprobarSobreescribir(ArchivoAbrir) Then Exit Sub


Open ArchivoAbrir For Binary As #N
'Escribimos la cabecera
Put #N, , MiCabecera
'Guardamos las cabezas

Put #N, , CInt(UBound(HeadData)) 'numheads
Dim Miscabezas() As tIndiceCabeza
ReDim Miscabezas(0 To UBound(HeadData) + 1) As tIndiceCabeza
ReDim MiscabezasLong(0 To UBound(HeadData) + 1) As tIndiceCabezaLong

If UsarGrhLong Then
    For i = 1 To UBound(HeadData)
        MiscabezasLong(i).Head(1) = HeadData(i).Head(1).GrhIndex
        MiscabezasLong(i).Head(2) = HeadData(i).Head(2).GrhIndex
        MiscabezasLong(i).Head(3) = HeadData(i).Head(3).GrhIndex
        MiscabezasLong(i).Head(4) = HeadData(i).Head(4).GrhIndex
        Put #N, , MiscabezasLong(i)
    Next i
Else
    For i = 1 To UBound(HeadData)
        Miscabezas(i).Head(1) = HeadData(i).Head(1).GrhIndex
        Miscabezas(i).Head(2) = HeadData(i).Head(2).GrhIndex
        Miscabezas(i).Head(3) = HeadData(i).Head(3).GrhIndex
        Miscabezas(i).Head(4) = HeadData(i).Head(4).GrhIndex
        Put #N, , Miscabezas(i)
    Next i
End If
Close #N

frmMain.LUlitError.Caption = "Guardado: " & ArchivoAbrir

EstadoNoGuardado(e_EstadoIndexador.Cabezas) = False

Exit Sub
errhandler:
Call MsgBox("Error en cabeza" & i)

End Sub
Public Sub GuardarCabezasDat(Optional ByVal FileNamePath As String = vbNullString)
On Error GoTo errhandler

Dim N As Integer, i As Integer
N = FreeFile
Dim ArchivoAbrir As String
If FileNamePath = vbNullString Then
    If SavePath = 0 Then
        ArchivoAbrir = App.Path & "\encode\Head.dat"
    Else
        ArchivoAbrir = App.Path & "\encode\Head" & SavePath & ".dat"
    End If
Else
    ArchivoAbrir = FileNamePath
End If

If Not ComprobarSobreescribir(ArchivoAbrir) Then Exit Sub


Call WriteVar(ArchivoAbrir, "INIT", "NumHeads", CInt(UBound(HeadData)))

For i = 1 To UBound(HeadData)
    If HeadData(i).Head(1).GrhIndex > 0 Then
        Call WriteVar(ArchivoAbrir, "Head" & i, "Head1", HeadData(i).Head(1).GrhIndex)
        Call WriteVar(ArchivoAbrir, "Head" & i, "Head2", HeadData(i).Head(2).GrhIndex)
        Call WriteVar(ArchivoAbrir, "Head" & i, "Head3", HeadData(i).Head(3).GrhIndex)
        Call WriteVar(ArchivoAbrir, "Head" & i, "Head4", HeadData(i).Head(4).GrhIndex)
        DoEvents
        frmMain.LUlitError.Caption = "cabeza: " & i
    End If
Next i

frmMain.LUlitError.Caption = "Guardado: " & ArchivoAbrir

EstadoNoGuardado(e_EstadoIndexador.Cabezas) = False

Exit Sub

errhandler:
Call MsgBox("Error en cabeza" & i)

End Sub

Public Sub GuardarSuperficies()
    Dim i As Long
    Dim nPath As String
    
    nPath = ConfigDir.InitWE & "\indices.ini"
    
    Call WriteVar(nPath, "INIT", "Referencias", NumSuperficies)
    
    For i = 0 To NumSuperficies
        With SupData(i)
            Call WriteVar(nPath, "REFERENCIA" & i, "GrhIndice", .GrhIndex)
            If .GrhIndex > 0 And .GrhIndex < UBound(Grhdata) Then
                Call WriteVar(nPath, "REFERENCIA" & i, "Nombre", .Nombre)
                Call WriteVar(nPath, "REFERENCIA" & i, "Ancho", .Ancho)
                Call WriteVar(nPath, "REFERENCIA" & i, "Alto", .Alto)
            End If
        End With
    Next i
    
End Sub
Public Sub GuardarFxs(Optional ByVal FileNamePath As String = vbNullString)
On Error GoTo errhandler

numfxs = UBound(FxData)
ReDim FxDataI(0 To numfxs + 1) As tIndiceFx


Dim N As Integer, i As Integer
N = FreeFile
Dim ArchivoAbrir As String

If FileNamePath = vbNullString Then
    If SavePath = 0 Then
        ArchivoAbrir = ConfigDir.Inits & "\Fxs.ind"
    Else
        ArchivoAbrir = ConfigDir.Inits & "\Fxs" & SavePath & ".ind"
    End If
Else
    ArchivoAbrir = FileNamePath
End If

If Not ComprobarSobreescribir(ArchivoAbrir) Then Exit Sub

Open ArchivoAbrir For Binary As #N

'Escribimos la cabecera
Put #N, , MiCabecera
'Guardamos las cabezas
Put #N, , numfxs

If UsarGrhLong Then
    For i = 1 To numfxs
        MisFxslong(i).Animacion = FxData(i).FX.GrhIndex
        MisFxslong(i).offsetx = FxData(i).offsetx
        MisFxslong(i).offsety = FxData(i).offsety
        Put #N, , MisFxslong(i)
    Next i
Else
    For i = 1 To numfxs
        FxDataI(i).Animacion = FxData(i).FX.GrhIndex
        FxDataI(i).offsetx = FxData(i).offsetx
        FxDataI(i).offsety = FxData(i).offsety
        Put #N, , FxDataI(i)
    Next i
End If
Close #N

frmMain.LUlitError.Caption = "Guardado: " & ArchivoAbrir

EstadoNoGuardado(e_EstadoIndexador.FX) = False

Exit Sub
errhandler:
Call MsgBox("Error en FX " & i)

End Sub

Public Sub GuardarFxsDat(Optional ByVal FileNamePath As String = vbNullString)
On Error GoTo errhandler


numfxs = UBound(FxData)


Dim N As Integer, i As Integer
N = FreeFile
Dim ArchivoAbrir As String


If FileNamePath = vbNullString Then
    If SavePath = 0 Then
        ArchivoAbrir = App.Path & "\encode\fx.dat"
    Else
        ArchivoAbrir = App.Path & "\encode\fx" & SavePath & ".dat"
    End If
Else
    ArchivoAbrir = FileNamePath
End If

If Not ComprobarSobreescribir(ArchivoAbrir) Then Exit Sub

Call WriteVar(ArchivoAbrir, "INIT", "NumFxs", numfxs)

For i = 1 To numfxs
    If FxData(i).FX.GrhIndex > 0 Then
        Call WriteVar(ArchivoAbrir, "Fx" & i, "Animacion", FxData(i).FX.GrhIndex)
        Call WriteVar(ArchivoAbrir, "Fx" & i, "OffsetX", FxData(i).offsetx)
        Call WriteVar(ArchivoAbrir, "Fx" & i, "OffsetY", FxData(i).offsety)
        frmMain.LUlitError.Caption = "Fx : " & i
        DoEvents
    End If
Next i


frmMain.LUlitError.Caption = "Guardado: " & ArchivoAbrir

EstadoNoGuardado(e_EstadoIndexador.FX) = False

Exit Sub
errhandler:
Call MsgBox("Error en FX " & i)

End Sub

Public Sub GuardarBodys(Optional ByVal FileNamePath As String = vbNullString)
On Error GoTo errhandler

Dim N As Integer, i As Integer
N = FreeFile
Dim ArchivoAbrir As String


If FileNamePath = vbNullString Then
    If SavePath = 0 Then
        ArchivoAbrir = ConfigDir.Inits & "\Personajes.ind"
    Else
        ArchivoAbrir = ConfigDir.Inits & "\Personajes" & SavePath & ".ind"
    End If
Else
    ArchivoAbrir = FileNamePath
End If

If Not ComprobarSobreescribir(ArchivoAbrir) Then Exit Sub

Open ArchivoAbrir For Binary As #N
'Escribimos la cabecera
Put #N, , MiCabecera
'Guardamos las cabezas


Dim MisCuerpos() As tIndiceCuerpo
Dim MisCuerposLong() As tIndiceCuerpoLong
ReDim MisCuerpos(0 To UBound(BodyData) + 1) As tIndiceCuerpo

ReDim MisCuerposLong(0 To UBound(BodyData) + 1) As tIndiceCuerpoLong




Put #N, , CInt(UBound(BodyData)) 'numheads


    For i = 1 To UBound(BodyData)
        MisCuerposLong(i).Body(1) = BodyData(i).Walk(1).GrhIndex
        MisCuerposLong(i).Body(2) = BodyData(i).Walk(2).GrhIndex
        MisCuerposLong(i).Body(3) = BodyData(i).Walk(3).GrhIndex
        MisCuerposLong(i).Body(4) = BodyData(i).Walk(4).GrhIndex
        MisCuerposLong(i).HeadOffsetX = BodyData(i).HeadOffset.x
        MisCuerposLong(i).HeadOffsetY = BodyData(i).HeadOffset.y
        Put #N, , MisCuerposLong(i)
    Next i

Close #N

frmMain.LUlitError.Caption = "Guardado: " & ArchivoAbrir

EstadoNoGuardado(e_EstadoIndexador.Body) = False

Exit Sub
errhandler:
Call MsgBox("Error en cuerpo " & i & " . " & Err.Description)

End Sub
Public Sub GuardarBodysDat(Optional ByVal FileNamePath As String = vbNullString)
On Error GoTo errhandler

Dim N As Integer, i As Integer
N = FreeFile
Dim ArchivoAbrir As String


If FileNamePath = vbNullString Then
    If SavePath = 0 Then
        ArchivoAbrir = App.Path & "\encode\Body.dat"
    Else
        ArchivoAbrir = App.Path & "\encode\Body" & SavePath & ".dat"
    End If
Else
    ArchivoAbrir = FileNamePath
End If

If Not ComprobarSobreescribir(ArchivoAbrir) Then Exit Sub


Call WriteVar(ArchivoAbrir, "INIT", "NumBodies", CInt(UBound(BodyData))) 'numheads

For i = 1 To UBound(BodyData)
    If BodyData(i).Walk(1).GrhIndex > 0 Then
        Call WriteVar(ArchivoAbrir, "Body" & i, "WALK1", BodyData(i).Walk(1).GrhIndex)
        Call WriteVar(ArchivoAbrir, "Body" & i, "WALK2", BodyData(i).Walk(2).GrhIndex)
        Call WriteVar(ArchivoAbrir, "Body" & i, "WALK3", BodyData(i).Walk(3).GrhIndex)
        Call WriteVar(ArchivoAbrir, "Body" & i, "WALK4", BodyData(i).Walk(4).GrhIndex)
        Call WriteVar(ArchivoAbrir, "Body" & i, "HeadOffsetX", BodyData(i).HeadOffset.x)
        Call WriteVar(ArchivoAbrir, "Body" & i, "HeadOffsety", BodyData(i).HeadOffset.y)
        frmMain.LUlitError.Caption = "body : " & i
        DoEvents
    End If
Next i


frmMain.LUlitError.Caption = "Guardado: " & ArchivoAbrir
Exit Sub

EstadoNoGuardado(e_EstadoIndexador.Body) = False

errhandler:
Call MsgBox("Error en cuerpo " & i & " . " & Err.Description)

End Sub

Public Sub GuardarCascos(Optional ByVal FileNamePath As String = vbNullString)
On Error GoTo errhandler

Dim N As Integer, i As Integer
N = FreeFile
Dim ArchivoAbrir As String
If FileNamePath = vbNullString Then
    If SavePath = 0 Then
        ArchivoAbrir = ConfigDir.Inits & "\Cascos.ind"
    Else
        ArchivoAbrir = ConfigDir.Inits & "\Cascos" & SavePath & ".ind"
    End If
Else
    ArchivoAbrir = FileNamePath
End If

If Not ComprobarSobreescribir(ArchivoAbrir) Then Exit Sub

Open ArchivoAbrir For Binary As #N
'Escribimos la cabecera
Put #N, , MiCabecera
'Guardamos las cabezas

Dim Miscabezas() As tIndiceCabeza
ReDim Miscabezas(0 To UBound(CascoAnimData) + 1) As tIndiceCabeza
ReDim MiscabezasLong(0 To UBound(CascoAnimData) + 1) As tIndiceCabezaLong

Put #N, , CInt(UBound(CascoAnimData)) 'numheads

If UsarGrhLong Then
    For i = 1 To UBound(CascoAnimData)
        MiscabezasLong(i).Head(1) = CascoAnimData(i).Head(1).GrhIndex
        MiscabezasLong(i).Head(2) = CascoAnimData(i).Head(2).GrhIndex
        MiscabezasLong(i).Head(3) = CascoAnimData(i).Head(3).GrhIndex
        MiscabezasLong(i).Head(4) = CascoAnimData(i).Head(4).GrhIndex
        Put #N, , MiscabezasLong(i)
    Next i
Else
    
    For i = 1 To UBound(CascoAnimData)
        Miscabezas(i).Head(1) = CascoAnimData(i).Head(1).GrhIndex
        Miscabezas(i).Head(2) = CascoAnimData(i).Head(2).GrhIndex
        Miscabezas(i).Head(3) = CascoAnimData(i).Head(3).GrhIndex
        Miscabezas(i).Head(4) = CascoAnimData(i).Head(4).GrhIndex
        Put #N, , Miscabezas(i)
    Next i
End If
Close #N

frmMain.LUlitError.Caption = "Guardado: " & ArchivoAbrir

EstadoNoGuardado(e_EstadoIndexador.Cascos) = False

Exit Sub
errhandler:
Call MsgBox("Error en casco " & i)

End Sub

Public Sub GuardarCascosDat(Optional ByVal FileNamePath As String = vbNullString)
On Error GoTo errhandler

Dim N As Integer, i As Integer
N = FreeFile
Dim ArchivoAbrir As String
If FileNamePath = vbNullString Then
    If SavePath = 0 Then
        ArchivoAbrir = App.Path & "\encode\Cascos.dat"
    Else
        ArchivoAbrir = App.Path & "\encode\Cascos" & SavePath & ".dat"
    End If
Else
    ArchivoAbrir = FileNamePath
End If

If Not ComprobarSobreescribir(ArchivoAbrir) Then Exit Sub


Call WriteVar(ArchivoAbrir, "INIT", "NumCascos", CInt(UBound(CascoAnimData)))

For i = 1 To UBound(CascoAnimData)
    If CascoAnimData(i).Head(1).GrhIndex > 0 Then
        Call WriteVar(ArchivoAbrir, "Casco" & i, "Head1", CascoAnimData(i).Head(1).GrhIndex)
        Call WriteVar(ArchivoAbrir, "Casco" & i, "Head2", CascoAnimData(i).Head(2).GrhIndex)
        Call WriteVar(ArchivoAbrir, "Casco" & i, "Head3", CascoAnimData(i).Head(3).GrhIndex)
        Call WriteVar(ArchivoAbrir, "Casco" & i, "Head4", CascoAnimData(i).Head(4).GrhIndex)
    End If
Next i


frmMain.LUlitError.Caption = "Guardado: " & ArchivoAbrir

EstadoNoGuardado(e_EstadoIndexador.Cascos) = False

Exit Sub
errhandler:
Call MsgBox("Error en casco " & i)

End Sub
Public Sub GuardarArmas(Optional ByVal FileNamePath As String = vbNullString)
On Error GoTo errhandler
Dim Narchivo As String
Dim N As Integer, i As Integer
Dim ArchivoAbrir As String

If FileNamePath = vbNullString Then
    If SavePath = 0 Then
        ArchivoAbrir = ConfigDir.Inits & "\armas.dat"
    Else
        ArchivoAbrir = ConfigDir.Inits & "\armas" & SavePath & ".dat"
    End If
Else
    ArchivoAbrir = FileNamePath
End If

If Not ComprobarSobreescribir(ArchivoAbrir) Then Exit Sub

Narchivo = ArchivoAbrir
Call WriteVar(Narchivo, "INIT", "NumArmas", UBound(WeaponAnimData))
For i = 1 To UBound(WeaponAnimData)
    If WeaponAnimData(i).WeaponWalk(1).GrhIndex > 0 Then
        Call WriteVar(Narchivo, "ARMA" & i, "Dir1", WeaponAnimData(i).WeaponWalk(1).GrhIndex)
        Call WriteVar(Narchivo, "ARMA" & i, "Dir2", WeaponAnimData(i).WeaponWalk(2).GrhIndex)
        Call WriteVar(Narchivo, "ARMA" & i, "Dir3", WeaponAnimData(i).WeaponWalk(3).GrhIndex)
        Call WriteVar(Narchivo, "ARMA" & i, "Dir4", WeaponAnimData(i).WeaponWalk(4).GrhIndex)
    End If
Next i
frmMain.LUlitError.Caption = "Guardado: " & ArchivoAbrir

EstadoNoGuardado(e_EstadoIndexador.Armas) = False

Exit Sub
errhandler:
Call MsgBox("Error en arma " & i)
End Sub
Public Sub GuardarEscudos(Optional ByVal FileNamePath As String = vbNullString)
On Error GoTo errhandler
Dim Narchivo As String
Dim N As Integer, i As Integer
Dim ArchivoAbrir As String

If FileNamePath = vbNullString Then
    If SavePath = 0 Then
        ArchivoAbrir = ConfigDir.Inits & "\Escudos.dat"
    Else
        ArchivoAbrir = ConfigDir.Inits & "\Escudos" & SavePath & ".dat"
    End If
Else
    ArchivoAbrir = FileNamePath
End If

If Not ComprobarSobreescribir(ArchivoAbrir) Then Exit Sub

Narchivo = ArchivoAbrir
Call WriteVar(Narchivo, "INIT", "NumEscudos", UBound(ShieldAnimData))
For i = 1 To UBound(ShieldAnimData)
    If ShieldAnimData(i).ShieldWalk(1).GrhIndex > 0 Then
        Call WriteVar(Narchivo, "ESC" & i, "Dir1", ShieldAnimData(i).ShieldWalk(1).GrhIndex)
        Call WriteVar(Narchivo, "ESC" & i, "Dir2", ShieldAnimData(i).ShieldWalk(2).GrhIndex)
        Call WriteVar(Narchivo, "ESC" & i, "Dir3", ShieldAnimData(i).ShieldWalk(3).GrhIndex)
        Call WriteVar(Narchivo, "ESC" & i, "Dir4", ShieldAnimData(i).ShieldWalk(4).GrhIndex)
    End If
Next i
frmMain.LUlitError.Caption = "Guardado: " & ArchivoAbrir

EstadoNoGuardado(e_EstadoIndexador.Escudos) = False

Exit Sub
errhandler:
Call MsgBox("Error en escudo " & i)
End Sub


'********************************************************************************
'********************************************************************************
'********************************************************************************
'******************************** Botones ***************************************
'********************************************************************************
'********************************************************************************
'********************************************************************************

Public Sub BotonGuardado(Optional ByVal FileNamePath As String = vbNullString)

Select Case EstadoIndexador
    Case e_EstadoIndexador.Grh
        Call SaveGrhData(FileNamePath)
    Case e_EstadoIndexador.Body
        Call GuardarBodys(FileNamePath)
    Case e_EstadoIndexador.Cabezas
        Call GuardarCabezas(FileNamePath)
    Case e_EstadoIndexador.Cascos
        Call GuardarCascos(FileNamePath)
    Case e_EstadoIndexador.Escudos
        Call GuardarEscudos(FileNamePath)
    Case e_EstadoIndexador.Armas
        Call GuardarArmas(FileNamePath)
    Case e_EstadoIndexador.FX
        Call GuardarFxs(FileNamePath)
    Case e_EstadoIndexador.Superficies
        Call GuardarSuperficies
        
End Select
End Sub
Public Sub BotonGuardadoDat(Optional ByVal FileNamePath As String = vbNullString)

Select Case EstadoIndexador
    Case e_EstadoIndexador.Grh
        Call SaveGrhDataDat(FileNamePath)
    Case e_EstadoIndexador.Body
        Call GuardarBodysDat(FileNamePath)
    Case e_EstadoIndexador.Cabezas
        Call GuardarCabezasDat(FileNamePath)
    Case e_EstadoIndexador.Cascos
        Call GuardarCascosDat(FileNamePath)
    Case e_EstadoIndexador.Escudos
        Call GuardarEscudos(FileNamePath)
    Case e_EstadoIndexador.Armas
        Call GuardarArmas(FileNamePath)
    Case e_EstadoIndexador.FX
        Call GuardarFxsDat(FileNamePath)
End Select
End Sub
Public Sub BotonCargado(Optional ByVal FileNamePath As String = vbNullString)
Dim respuesta As Byte
Dim TempLong As Long

respuesta = MsgBox("ATENCION Si contunias perderas los cambios no guardados", 4, "¡¡ADVERTENCIA!!")
If respuesta <> vbYes Then
    Exit Sub
End If
        
frmMain.Visor.Cls
Select Case EstadoIndexador
    Case e_EstadoIndexador.Grh
        Call LoadGrhData(FileNamePath)
        Call RenuevaListaGrH
    Case e_EstadoIndexador.Body
        Call CargarCuerpos(FileNamePath)
        Call RenuevaListaBodys
    Case e_EstadoIndexador.Cabezas
        Call CargarCabezas(FileNamePath)
        Call RenuevaListaCabezas
    Case e_EstadoIndexador.Cascos
        Call CargarCascos(FileNamePath)
        Call RenuevaListaCascos
    Case e_EstadoIndexador.Escudos
        Call CargarAnimEscudos(FileNamePath)
        Call RenuevaListaEscudos
    Case e_EstadoIndexador.Armas
        Call CargarAnimArmas(FileNamePath)
        Call RenuevaListaArmas

    Case e_EstadoIndexador.FX
        Call CargarFxs(FileNamePath)
        Call RenuevaListaFX
End Select
If EstadoIndexador = e_EstadoIndexador.Grh Then
    TempLong = ListaindexGrH(GRHActual)
    If TempLong >= frmMain.Lista.ListCount Then TempLong = 0
    frmMain.Lista.listIndex = TempLong
Else
    TempLong = ListaindexGrH(DataIndexActual)
    If TempLong >= frmMain.Lista.ListCount Then TempLong = 0
    frmMain.Lista.listIndex = TempLong
End If
End Sub

Public Sub BotonCargadoDat(Optional ByVal FileNamePath As String = vbNullString)
Dim respuesta As Byte
Dim TempLong As Long

respuesta = MsgBox("ATENCION Si contunias perderas los cambios no guardados", 4, "¡¡ADVERTENCIA!!")
If respuesta <> vbYes Then
    Exit Sub
End If
        
frmMain.Visor.Cls
Select Case EstadoIndexador
    Case e_EstadoIndexador.Grh
        Call LoadGrhDataDat(FileNamePath)
        Call RenuevaListaGrH
    Case e_EstadoIndexador.Body
        Call CargarCuerposDat(FileNamePath)
        Call RenuevaListaBodys
    Case e_EstadoIndexador.Cabezas
        Call CargarCabezasdat(FileNamePath)
        Call RenuevaListaCabezas
    Case e_EstadoIndexador.Cascos
        Call CargarCascosDat(FileNamePath)
        Call RenuevaListaCascos
    Case e_EstadoIndexador.Escudos
        Call CargarAnimEscudos(FileNamePath)
        Call RenuevaListaEscudos
    Case e_EstadoIndexador.Armas
        Call CargarAnimArmas(FileNamePath)
        Call RenuevaListaArmas

    Case e_EstadoIndexador.FX
        Call CargarFxsDat(FileNamePath)
        Call RenuevaListaFX
End Select
If EstadoIndexador = e_EstadoIndexador.Grh Then
    TempLong = ListaindexGrH(GRHActual)
    frmMain.Lista.listIndex = TempLong
Else
    TempLong = ListaindexGrH(DataIndexActual)
    frmMain.Lista.listIndex = TempLong
End If
End Sub
Sub CargarConfig()
        ConfigDir.Inits = GetVar(App.Path & "\Conf.ini", "Config", "Inits")
    ConfigDir.Graficos = GetVar(App.Path & "\Conf.ini", "Config", "Graficos")
    ConfigDir.Dats = GetVar(App.Path & "\Conf.ini", "Config", "Dats")
    ConfigDir.InitWE = GetVar(App.Path & "\Conf.ini", "Config", "InitWE")
End Sub
Sub guardarConfig()
    Call WriteVar(App.Path & "\Conf.ini", "Config", "Inits", ConfigDir.Inits)
    Call WriteVar(App.Path & "\Conf.ini", "Config", "Graficos", ConfigDir.Graficos)
    Call WriteVar(App.Path & "\Conf.ini", "Config", "Dats", ConfigDir.Dats)
    Call WriteVar(App.Path & "\Conf.ini", "Config", "InitWE", ConfigDir.InitWE)
End Sub
