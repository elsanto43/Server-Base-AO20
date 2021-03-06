Attribute VB_Name = "ES"
'Argentum Online 0.12.2
'Copyright (C) 2002 Marquez Pablo Ignacio
'
'This program is free software; you can redistribute it and/or modify
'it under the terms of the Affero General Public License;
'either version 1 of the License, or any later version.
'
'This program is distributed in the hope that it will be useful,
'but WITHOUT ANY WARRANTY; without even the implied warranty of
'MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'Affero General Public License for more details.
'
'You should have received a copy of the Affero General Public License
'along with this program; if not, you can find it at http://www.affero.org/oagpl.html
'
'Argentum Online is based on Baronsoft's VB6 Online RPG
'You can contact the original creator of ORE at aaron@baronsoft.com
'for more information about ORE please visit http://www.baronsoft.com/
'
'
'You can contact me at:
'morgolock@speedy.com.ar
'www.geocities.com/gmorgolock
'Calle 3 numero 983 piso 7 dto A
'La Plata - Pcia, Buenos Aires - Republica Argentina
'Codigo Postal 1900
'Pablo Ignacio Marquez

Option Explicit

#If False Then

    Dim x, y, n, map, Mapa, Email, max, Value As Variant

#End If

Public Sub CargarSpawnList()
    '****************************************************************************************
    'Author: Unknown
    'Last Modification: 27/03/2020
    'Cargo la lista de NPC's hostiles desde el NPC's.dat
    ' - Omitimos los NPC's pretorianos ya que deben invocarse mediante su respectivo comando.
    '****************************************************************************************
    If frmMain.Visible Then frmMain.txtStatus.Text = "Cargando Invokar.dat"
    
    ReDim SpawnList(1 To val(LeerNPCs.GetValue("INIT", "NumNPCs"))) As tCriaturasEntrenador
    
    Dim i As Integer: i = 0
    
    Dim LoopC As Long
    For LoopC = 1 To UBound(SpawnList)
        
        If val(LeerNPCs.GetValue("NPC" & LoopC, "Hostile")) = 1 And _
           val(LeerNPCs.GetValue("NPC" & LoopC, "NpcType")) <> 10 Then
            
            i = i + 1
            
            SpawnList(i).NpcIndex = LoopC
            SpawnList(i).NpcName = LeerNPCs.GetValue("NPC" & LoopC, "Name")
            
        End If
        
    Next
    
    ' Hacemos el trim a la lista.
    ReDim Preserve SpawnList(1 To i) As tCriaturasEntrenador
    
    If frmMain.Visible Then frmMain.txtStatus.Text = "Lista de NPC's hostiles se cargo correctamente"
    
End Sub

Function EsAdmin(ByRef Name As String) As Boolean
    '***************************************************
    'Author: Unknown
    'Last Modification: 27/03/2011
    '27/03/2011 - ZaMa: Utilizo la clase para saber los datos.
    '***************************************************
    EsAdmin = (val(Administradores.GetValue("Admin", Name)) = 1)

End Function

Function EsDios(ByRef Name As String) As Boolean
    '***************************************************
    'Author: Unknown
    'Last Modification: 27/03/2011
    '27/03/2011 - ZaMa: Utilizo la clase para saber los datos.
    '***************************************************
    EsDios = (val(Administradores.GetValue("Dios", Name)) = 1)

End Function

Function EsSemiDios(ByRef Name As String) As Boolean
    '***************************************************
    'Author: Unknown
    'Last Modification: 27/03/2011
    '27/03/2011 - ZaMa: Utilizo la clase para saber los datos.
    '***************************************************
    EsSemiDios = (val(Administradores.GetValue("SemiDios", Name)) = 1)

End Function

Function EsGmEspecial(ByRef Name As String) As Boolean
    '***************************************************
    'Author: ZaMa
    'Last Modification: 27/03/2011
    '27/03/2011 - ZaMa: Utilizo la clase para saber los datos.
    '***************************************************
    EsGmEspecial = (val(Administradores.GetValue("Especial", Name)) = 1)

End Function

Function EsConsejero(ByRef Name As String) As Boolean
    '***************************************************
    'Author: Unknown
    'Last Modification: 27/03/2011
    '27/03/2011 - ZaMa: Utilizo la clase para saber los datos.
    '***************************************************
    EsConsejero = (val(Administradores.GetValue("Consejero", Name)) = 1)

End Function

Function EsRolesMaster(ByRef Name As String) As Boolean
    '***************************************************
    'Author: Unknown
    'Last Modification: 27/03/2011
    '27/03/2011 - ZaMa: Utilizo la clase para saber los datos.
    '***************************************************
    EsRolesMaster = (val(Administradores.GetValue("RM", Name)) = 1)

End Function

Public Function EsGmChar(ByRef Name As String) As Boolean
    '***************************************************
    'Author: ZaMa
    'Last Modification: 27/03/2011
    'Returns true if char is administrative user.
    '***************************************************
    
    Dim EsGm As Boolean
    
    ' Admin?
    EsGm = EsAdmin(Name)

    ' Dios?
    If Not EsGm Then EsGm = EsDios(Name)

    ' Semidios?
    If Not EsGm Then EsGm = EsSemiDios(Name)

    ' Consejero?
    If Not EsGm Then EsGm = EsConsejero(Name)

    EsGmChar = EsGm

End Function

Public Sub loadAdministrativeUsers()
    'Admines     => Admin
    'Dioses      => Dios
    'SemiDioses  => SemiDios
    'Especiales  => Especial
    'Consejeros  => Consejero
    'RoleMasters => RM
    If frmMain.Visible Then frmMain.txtStatus.Text = "Cargando Administradores/Dioses/Gms."

    'Si esta mierda tuviese array asociativos el codigo seria tan lindo.
    Dim buf  As Integer

    Dim i    As Long

    Dim Name As String
       
    ' Public container
    Set Administradores = New clsIniManager
    
    ' Server ini info file
    Dim ServerIni As clsIniManager

    Set ServerIni = New clsIniManager
    
    Call ServerIni.Initialize(IniPath & "Server.ini")
       
    ' Admines
    buf = val(ServerIni.GetValue("INIT", "Admines"))
    
    For i = 1 To buf
        Name = UCase$(ServerIni.GetValue("Admines", "Admin" & i))
        
        If Left$(Name, 1) = "*" Or Left$(Name, 1) = "+" Then Name = Right$(Name, Len(Name) - 1)
        
        ' Add key
        Call Administradores.ChangeValue("Admin", Name, "1")

    Next i
    
    ' Dioses
    buf = val(ServerIni.GetValue("INIT", "Dioses"))
    
    For i = 1 To buf
        Name = UCase$(ServerIni.GetValue("Dioses", "Dios" & i))
        
        If Left$(Name, 1) = "*" Or Left$(Name, 1) = "+" Then Name = Right$(Name, Len(Name) - 1)
        
        ' Add key
        Call Administradores.ChangeValue("Dios", Name, "1")
        
    Next i
    
    ' Especiales
    buf = val(ServerIni.GetValue("INIT", "Especiales"))
    
    For i = 1 To buf
        Name = UCase$(ServerIni.GetValue("Especiales", "Especial" & i))
        
        If Left$(Name, 1) = "*" Or Left$(Name, 1) = "+" Then Name = Right$(Name, Len(Name) - 1)
        
        ' Add key
        Call Administradores.ChangeValue("Especial", Name, "1")
        
    Next i
    
    ' SemiDioses
    buf = val(ServerIni.GetValue("INIT", "SemiDioses"))
    
    For i = 1 To buf
        Name = UCase$(ServerIni.GetValue("SemiDioses", "SemiDios" & i))
        
        If Left$(Name, 1) = "*" Or Left$(Name, 1) = "+" Then Name = Right$(Name, Len(Name) - 1)
        
        ' Add key
        Call Administradores.ChangeValue("SemiDios", Name, "1")
        
    Next i
    
    ' Consejeros
    buf = val(ServerIni.GetValue("INIT", "Consejeros"))
        
    For i = 1 To buf
        Name = UCase$(ServerIni.GetValue("Consejeros", "Consejero" & i))
        
        If Left$(Name, 1) = "*" Or Left$(Name, 1) = "+" Then Name = Right$(Name, Len(Name) - 1)
        
        ' Add key
        Call Administradores.ChangeValue("Consejero", Name, "1")
        
    Next i
    
    ' RolesMasters
    buf = val(ServerIni.GetValue("INIT", "RolesMasters"))
        
    For i = 1 To buf
        Name = UCase$(ServerIni.GetValue("RolesMasters", "RM" & i))
        
        If Left$(Name, 1) = "*" Or Left$(Name, 1) = "+" Then Name = Right$(Name, Len(Name) - 1)
        
        ' Add key
        Call Administradores.ChangeValue("RM", Name, "1")
    Next i
    
    Set ServerIni = Nothing

    If frmMain.Visible Then frmMain.txtStatus.Text = Date & " " & time & " - Los Administradores/Dioses/Gms se han cargado correctamente."

End Sub

Public Function GetCharPrivs(ByRef UserName As String) As PlayerType
    '****************************************************
    'Author: ZaMa
    'Last Modification: 18/11/2010
    'Reads the user's charfile and retrieves its privs.
    '***************************************************

    Dim Privs As PlayerType

    If EsAdmin(UserName) Then
        Privs = PlayerType.Admin
        
    ElseIf EsDios(UserName) Then
        Privs = PlayerType.Dios

    ElseIf EsSemiDios(UserName) Then
        Privs = PlayerType.SemiDios
        
    ElseIf EsConsejero(UserName) Then
        Privs = PlayerType.Consejero
    
    Else
        Privs = PlayerType.User

    End If

    GetCharPrivs = Privs

End Function

Public Function TxtDimension(ByVal Name As String) As Long
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    Dim n As Integer, cad As String, Tam As Long

    n = FreeFile(1)
    Open Name For Input As #n
    Tam = 0

    Do While Not EOF(n)
        Tam = Tam + 1
        Line Input #n, cad
    Loop
    Close n
    TxtDimension = Tam

End Function

Public Sub CargarForbidenWords()
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************
    If frmMain.Visible Then frmMain.txtStatus.Text = "Cargando Nombres prohibidos (NombresInvalidos.txt)."

    ReDim ForbidenNames(1 To TxtDimension(DatPath & "NombresInvalidos.txt"))

    Dim n As Integer, i As Integer

    n = FreeFile(1)
    Open DatPath & "NombresInvalidos.txt" For Input As #n
    
    For i = 1 To UBound(ForbidenNames)
        Line Input #n, ForbidenNames(i)
    Next i
    
    Close n

    If frmMain.Visible Then frmMain.txtStatus.Text = Date & " " & time & " - NombresInvalidos.txt han cargado con exito."

End Sub

Public Sub CargarHechizos()
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    '###################################################
    '#               ATENCION PELIGRO                  #
    '###################################################
    '
    '   NO USAR GetVar PARA LEER Hechizos.dat !!!!
    '
    'El que ose desafiar esta LEY, se las tendra que ver
    'con migo. Para leer Hechizos.dat se debera usar
    'la nueva clase clsLeerInis.
    '
    'Alejo
    '
    '###################################################

    On Error GoTo errHandler

    If frmMain.Visible Then frmMain.txtStatus.Text = "Cargando Hechizos."
    
    Dim Hechizo As Integer

    Dim leer    As clsIniManager

    Set leer = New clsIniManager
    
    Call leer.Initialize(DatPath & "Hechizos.dat")
    
    'obtiene el numero de hechizos
    NumeroHechizos = val(leer.GetValue("INIT", "NumeroHechizos"))
    
    ReDim Hechizos(1 To NumeroHechizos) As tHechizo
    
    frmCargando.cargar.min = 0
    frmCargando.cargar.max = NumeroHechizos
    frmCargando.cargar.Value = 0
    
    'Llena la lista
    For Hechizo = 1 To NumeroHechizos

        With Hechizos(Hechizo)
            '.Nombre = Leer.GetValue("Hechizo" & Hechizo, "Nombre")
            '.desc = Leer.GetValue("Hechizo" & Hechizo, "Desc")
            '.PalabrasMagicas = Leer.GetValue("Hechizo" & Hechizo, "PalabrasMagicas")
            
            '.HechizeroMsg = Leer.GetValue("Hechizo" & Hechizo, "HechizeroMsg")
            '.TargetMsg = Leer.GetValue("Hechizo" & Hechizo, "TargetMsg")
            '.PropioMsg = Leer.GetValue("Hechizo" & Hechizo, "PropioMsg")
            
            .tiempoCasteo = val(leer.GetValue("Hechizo" & Hechizo, "TiempoCasteo"))
            .Tipo = val(leer.GetValue("Hechizo" & Hechizo, "Tipo"))
            .WAV = val(leer.GetValue("Hechizo" & Hechizo, "WAV"))
            .FXgrh = val(leer.GetValue("Hechizo" & Hechizo, "Fxgrh"))
            
            .loops = val(leer.GetValue("Hechizo" & Hechizo, "Loops"))
            
            '    .Resis = val(Leer.GetValue("Hechizo" & Hechizo, "Resis"))
            
            .SubeHP = val(leer.GetValue("Hechizo" & Hechizo, "SubeHP"))
            .MinHp = val(leer.GetValue("Hechizo" & Hechizo, "MinHP"))
            .MaxHp = val(leer.GetValue("Hechizo" & Hechizo, "MaxHP"))
            
            .SubeMana = val(leer.GetValue("Hechizo" & Hechizo, "SubeMana"))
            .MiMana = val(leer.GetValue("Hechizo" & Hechizo, "MinMana"))
            .MaMana = val(leer.GetValue("Hechizo" & Hechizo, "MaxMana"))
            
            .SubeSta = val(leer.GetValue("Hechizo" & Hechizo, "SubeSta"))
            .MinSta = val(leer.GetValue("Hechizo" & Hechizo, "MinSta"))
            .MaxSta = val(leer.GetValue("Hechizo" & Hechizo, "MaxSta"))
            
            .SubeHam = val(leer.GetValue("Hechizo" & Hechizo, "SubeHam"))
            .MinHam = val(leer.GetValue("Hechizo" & Hechizo, "MinHam"))
            .MaxHam = val(leer.GetValue("Hechizo" & Hechizo, "MaxHam"))
            
            .SubeSed = val(leer.GetValue("Hechizo" & Hechizo, "SubeSed"))
            .MinSed = val(leer.GetValue("Hechizo" & Hechizo, "MinSed"))
            .MaxSed = val(leer.GetValue("Hechizo" & Hechizo, "MaxSed"))
            
            .SubeAgilidad = val(leer.GetValue("Hechizo" & Hechizo, "SubeAG"))
            .MinAgilidad = val(leer.GetValue("Hechizo" & Hechizo, "MinAG"))
            .MaxAgilidad = val(leer.GetValue("Hechizo" & Hechizo, "MaxAG"))
            
            .SubeFuerza = val(leer.GetValue("Hechizo" & Hechizo, "SubeFU"))
            .MinFuerza = val(leer.GetValue("Hechizo" & Hechizo, "MinFU"))
            .MaxFuerza = val(leer.GetValue("Hechizo" & Hechizo, "MaxFU"))
            
            .SubeCarisma = val(leer.GetValue("Hechizo" & Hechizo, "SubeCA"))
            .MinCarisma = val(leer.GetValue("Hechizo" & Hechizo, "MinCA"))
            .MaxCarisma = val(leer.GetValue("Hechizo" & Hechizo, "MaxCA"))
            
            .Invisibilidad = val(leer.GetValue("Hechizo" & Hechizo, "Invisibilidad"))
            .Paraliza = val(leer.GetValue("Hechizo" & Hechizo, "Paraliza"))
            .Inmoviliza = val(leer.GetValue("Hechizo" & Hechizo, "Inmoviliza"))
            .RemoverParalisis = val(leer.GetValue("Hechizo" & Hechizo, "RemoverParalisis"))
            .RemoverEstupidez = val(leer.GetValue("Hechizo" & Hechizo, "RemoverEstupidez"))
            .RemueveInvisibilidadParcial = val(leer.GetValue("Hechizo" & Hechizo, "RemueveInvisibilidadParcial"))
            
            .CuraVeneno = val(leer.GetValue("Hechizo" & Hechizo, "CuraVeneno"))
            .Envenena = val(leer.GetValue("Hechizo" & Hechizo, "Envenena"))
            .Maldicion = val(leer.GetValue("Hechizo" & Hechizo, "Maldicion"))
            .RemoverMaldicion = val(leer.GetValue("Hechizo" & Hechizo, "RemoverMaldicion"))
            .Bendicion = val(leer.GetValue("Hechizo" & Hechizo, "Bendicion"))
            .Revivir = val(leer.GetValue("Hechizo" & Hechizo, "Revivir"))
            
            .Ceguera = val(leer.GetValue("Hechizo" & Hechizo, "Ceguera"))
            .Estupidez = val(leer.GetValue("Hechizo" & Hechizo, "Estupidez"))
            
            .Warp = val(leer.GetValue("Hechizo" & Hechizo, "Warp"))
            
            .Invoca = val(leer.GetValue("Hechizo" & Hechizo, "Invoca"))
            .NumNpc = val(leer.GetValue("Hechizo" & Hechizo, "NumNpc"))
            .cant = val(leer.GetValue("Hechizo" & Hechizo, "Cant"))
            .Mimetiza = val(leer.GetValue("hechizo" & Hechizo, "Mimetiza"))
            
            '    .Materializa = val(Leer.GetValue("Hechizo" & Hechizo, "Materializa"))
            '    .ItemIndex = val(Leer.GetValue("Hechizo" & Hechizo, "ItemIndex"))
            
            .MinSkill = val(leer.GetValue("Hechizo" & Hechizo, "MinSkill"))
            .ManaRequerido = val(leer.GetValue("Hechizo" & Hechizo, "ManaRequerido"))
            
            'Barrin 30/9/03
            .StaRequerido = val(leer.GetValue("Hechizo" & Hechizo, "StaRequerido"))
            
            .Target = val(leer.GetValue("Hechizo" & Hechizo, "Target"))
            frmCargando.cargar.Value = frmCargando.cargar.Value + 1
            
            .NeedStaff = val(leer.GetValue("Hechizo" & Hechizo, "NeedStaff"))
            .StaffAffected = CBool(val(leer.GetValue("Hechizo" & Hechizo, "StaffAffected")))

        End With

    Next Hechizo
    
    Set leer = Nothing

    If frmMain.Visible Then frmMain.txtStatus.Text = Date & " " & time & " - Los hechizos se han cargado con exito."
    
    Exit Sub

errHandler:
    MsgBox "Error cargando hechizos.dat " & Err.Number & ": " & Err.description
 
End Sub

Sub LoadMotd()
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************
    If frmMain.Visible Then frmMain.txtStatus.Text = "Cargando archivo MOTD.INI."

    Dim i As Integer
    
    MaxLines = val(GetVar(App.path & "\Dat\Motd.ini", "INIT", "NumLines"))
    
    ReDim MOTD(1 To MaxLines)

    For i = 1 To MaxLines
        MOTD(i).texto = GetVar(App.path & "\Dat\Motd.ini", "Motd", "Line" & i)
        MOTD(i).Formato = vbNullString
    Next i

    If frmMain.Visible Then frmMain.txtStatus.Text = Date & " " & time & " - El archivo MOTD.INI fue cargado con exito"

End Sub

Public Sub DoBackUp()
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************
    If frmMain.Visible Then frmMain.txtStatus.Text = Date & " " & time & " - Los hechizos se han cargado con exito."

    haciendoBK = True
    
    ' Lo saco porque elimina elementales y mascotas - Maraxus
    ''''''''''''''lo pongo aca x sugernecia del yind
    'For i = 1 To LastNPC
    '    If Npclist(i).flags.NPCActive Then
    '        If Npclist(i).Contadores.TiempoExistencia > 0 Then
    '            Call MuereNpc(i, 0)
    '        End If
    '    End If
    'Next i
    '''''''''''/'lo pongo aca x sugernecia del yind
    
    Call SendData(SendTarget.ToAll, 0, PrepareMessagePauseToggle())

    Call WorldSave
    Call modGuilds.v_RutinaElecciones
    
    Call SendData(SendTarget.ToAll, 0, PrepareMessagePauseToggle())
    
    'Aqui solo vamos a hacer un request a los endpoints de la aplicacion en Node.js
    'el repositorio para hacer funcionar esto, es este: https://github.com/ao-libre/ao-api-server
    'Si no tienen interes en usarlo pueden desactivarlo en el Server.ini
    If ConexionAPI Then
        Call ApiEndpointBackupCharfiles
        Call ApiEndpointBackupCuentas
        Call ApiEndpointBackupLogs
        Call ApiEndpointSendWorldSaveMessageDiscord
    End If

    haciendoBK = False
    
    'Log
    On Error Resume Next

    Dim nfile As Integer

    If frmMain.Visible Then frmMain.txtStatus.Text = Date & " " & time & " - El WorldSave (backup) se hizo correctamente."

    nfile = FreeFile ' obtenemos un canal
    Open App.path & "\logs\BackUps.log" For Append Shared As #nfile
    Print #nfile, Date & " " & time
    Close #nfile

End Sub

Public Sub GrabarMapa(ByVal map As Long, ByRef MAPFILE As String)
    '***************************************************
    'Author: Unknown
    'Last Modification: 12/01/2011
    '10/08/2010 - Pato: Implemento el clsByteBuffer para el grabado de mapas
    '28/10/2010:ZaMa - Ahora no se hace backup de los pretorianos.
    '12/01/2011 - Amraphen: Ahora no se hace backup de NPCs prohibidos (Pretorianos, Mascotas, Invocados )
    '***************************************************

    On Error Resume Next

    Dim FreeFileMap As Long
    Dim FreeFileInf As Long

    Dim y           As Long
    Dim x           As Long

    Dim ByFlags     As Byte

    Dim LoopC       As Long

    Dim MapWriter   As clsByteBuffer
    Dim InfWriter   As clsByteBuffer
    Dim IniManager  As clsIniManager

    Dim NpcInvalido As Boolean
    
    Set MapWriter = New clsByteBuffer
    Set InfWriter = New clsByteBuffer
    Set IniManager = New clsIniManager
    
    If FileExist(MAPFILE & ".map", vbNormal) Then
        Call Kill(MAPFILE & ".map")
    End If
    
    If FileExist(MAPFILE & ".inf", vbNormal) Then
        Call Kill(MAPFILE & ".inf")
    End If
    
    'Open .map file
    FreeFileMap = FreeFile
    
    Open MAPFILE & ".Map" For Binary As FreeFileMap
    
    Call MapWriter.initializeWriter(FreeFileMap)
    
    'Open .inf file
    FreeFileInf = FreeFile
    Open MAPFILE & ".Inf" For Binary As FreeFileInf
    
    Call InfWriter.initializeWriter(FreeFileInf)
    
    'map Header
    Call MapWriter.putInteger(MapInfo(map).MapVersion)
        
    Call MapWriter.putString(MiCabecera.Desc, False)
    Call MapWriter.putLong(MiCabecera.crc)
    Call MapWriter.putLong(MiCabecera.MagicWord)
    
    Call MapWriter.putDouble(0)
    
    'inf Header
    Call InfWriter.putDouble(0)
    Call InfWriter.putInteger(0)
    
    'Write .map file
    For y = YMinMapSize To YMaxMapSize
        For x = XMinMapSize To XMaxMapSize

            With MapData(map, x, y)
                ByFlags = 0
                
                If .Blocked Then ByFlags = ByFlags Or 1
                If .Graphic(2) Then ByFlags = ByFlags Or 2
                If .Graphic(3) Then ByFlags = ByFlags Or 4
                If .Graphic(4) Then ByFlags = ByFlags Or 8
                If .trigger Then ByFlags = ByFlags Or 16
                
                Call MapWriter.putByte(ByFlags)
                
                Call MapWriter.putLong(.Graphic(1))
                
                For LoopC = 2 To 4
                    If .Graphic(LoopC) Then Call MapWriter.putLong(.Graphic(LoopC))
                Next LoopC
                
                If .trigger Then Call MapWriter.putInteger(CInt(.trigger))
                
                '.inf file
                ByFlags = 0
                
                If .ObjInfo.ObjIndex > 0 Then
                    
                    If ObjData(.ObjInfo.ObjIndex).OBJType = eOBJType.otFogata Then
                        .ObjInfo.ObjIndex = 0
                        .ObjInfo.amount = 0
                    End If

                End If
    
                If .TileExit.map Then ByFlags = ByFlags Or 1
                
                ' No hacer backup de los NPCs invalidos (Pretorianos, Mascotas, Invocados )
                If .NpcIndex Then
                    
                    NpcInvalido = (Npclist(.NpcIndex).NPCtype = eNPCType.Pretoriano) Or _
                                  (Npclist(.NpcIndex).MaestroUser > 0)
                    
                    If Not NpcInvalido Then ByFlags = ByFlags Or 2

                End If
                
                If .ObjInfo.ObjIndex Then ByFlags = ByFlags Or 4
                
                Call InfWriter.putByte(ByFlags)
                
                If .TileExit.map Then
                    Call InfWriter.putInteger(.TileExit.map)
                    Call InfWriter.putInteger(.TileExit.x)
                    Call InfWriter.putInteger(.TileExit.y)
                End If
                
                If .NpcIndex And Not NpcInvalido Then Call InfWriter.putInteger(Npclist(.NpcIndex).Numero)
                
                If .ObjInfo.ObjIndex Then
                    Call InfWriter.putInteger(.ObjInfo.ObjIndex)
                    Call InfWriter.putInteger(.ObjInfo.amount)
                End If
                
                NpcInvalido = False

            End With

        Next x
    Next y
    
    Call MapWriter.saveBuffer
    Call InfWriter.saveBuffer
    
    'Close .map file
    Close FreeFileMap

    'Close .inf file
    Close FreeFileInf
    
    Set MapWriter = Nothing
    Set InfWriter = Nothing

    With MapInfo(map)
        'write .dat file
        Call IniManager.ChangeValue("Mapa" & map, "Name", .Name)
        Call IniManager.ChangeValue("Mapa" & map, "MusicNum", .Music)
        Call IniManager.ChangeValue("Mapa" & map, "MagiaSinefecto", .MagiaSinEfecto)
        Call IniManager.ChangeValue("Mapa" & map, "InviSinEfecto", .InviSinEfecto)
        Call IniManager.ChangeValue("Mapa" & map, "ResuSinEfecto", .ResuSinEfecto)
        Call IniManager.ChangeValue("Mapa" & map, "StartPos", .StartPos.map & "-" & .StartPos.x & "-" & .StartPos.y)
        Call IniManager.ChangeValue("Mapa" & map, "OnDeathGoTo", .OnDeathGoTo.map & "-" & .OnDeathGoTo.x & "-" & .OnDeathGoTo.y)
    
        Call IniManager.ChangeValue("Mapa" & map, "Terreno", TerrainByteToString(.Terreno))
        Call IniManager.ChangeValue("Mapa" & map, "Zona", .Zona)
        Call IniManager.ChangeValue("Mapa" & map, "Restringir", RestrictByteToString(.Restringir))
        Call IniManager.ChangeValue("Mapa" & map, "BackUp", Str(.BackUp))
    
        If .Pk Then
            Call IniManager.ChangeValue("Mapa" & map, "Pk", "0")
        Else
            Call IniManager.ChangeValue("Mapa" & map, "Pk", "1")

        End If
        
        Call IniManager.ChangeValue("Mapa" & map, "OcultarSinEfecto", .OcultarSinEfecto)
        Call IniManager.ChangeValue("Mapa" & map, "InvocarSinEfecto", .InvocarSinEfecto)
        Call IniManager.ChangeValue("Mapa" & map, "NoEncriptarMP", .NoEncriptarMP)
        Call IniManager.ChangeValue("Mapa" & map, "RoboNpcsPermitido", .RoboNpcsPermitido)
    
        Call IniManager.DumpFile(MAPFILE & ".dat")

    End With
    
    Set IniManager = Nothing

End Sub

Sub LoadArmasHerreria()
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************
    
    If frmMain.Visible Then frmMain.txtStatus.Text = "Cargando armas crafteables por Herreria."
    
    Dim n As Integer, lc As Integer
    
    n = val(GetVar(DatPath & "ArmasHerrero.dat", "INIT", "NumArmas"))
    
    ReDim Preserve ArmasHerrero(1 To n) As Integer
    
    For lc = 1 To n
        ArmasHerrero(lc) = val(GetVar(DatPath & "ArmasHerrero.dat", "Arma" & lc, "Index"))
    Next lc
    
    If frmMain.Visible Then frmMain.txtStatus.Text = Date & " " & time & " - Se cargo las armas crafteables por Herreria. Operacion Realizada con exito."
    
End Sub

Sub LoadArmadurasHerreria()
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************
        
    If frmMain.Visible Then frmMain.txtStatus.Text = "Cargando armaduras crafteables por Herreria."

    Dim n As Integer, lc As Integer
    
    n = val(GetVar(DatPath & "ArmadurasHerrero.dat", "INIT", "NumArmaduras"))
    
    ReDim Preserve ArmadurasHerrero(1 To n) As Integer
    
    For lc = 1 To n
        ArmadurasHerrero(lc) = val(GetVar(DatPath & "ArmadurasHerrero.dat", "Armadura" & lc, "Index"))
    Next lc
    
    If frmMain.Visible Then frmMain.txtStatus.Text = Date & " " & time & " - Se cargo las armaduras crafteables por Herreria. Operacion Realizada con exito."
    
End Sub

Sub LoadBalance()
    '***************************************************
    'Author: Unknown
    'Last Modification: 15/04/2010
    '15/04/2010: ZaMa - Agrego recompensas faccionarias.
    '***************************************************

    Dim i As Long

    If frmMain.Visible Then frmMain.txtStatus.Text = "Cargando el archivo Balance.dat"
    
    'Modificadores de Clase
    For i = 1 To NUMCLASES

        With ModClase(i)
            .Evasion = val(GetVar(DatPath & "Balance.dat", "MODEVASION", ListaClases(i)))
            .AtaqueArmas = val(GetVar(DatPath & "Balance.dat", "MODATAQUEARMAS", ListaClases(i)))
            .AtaqueProyectiles = val(GetVar(DatPath & "Balance.dat", "MODATAQUEPROYECTILES", ListaClases(i)))
            .AtaqueWrestling = val(GetVar(DatPath & "Balance.dat", "MODATAQUEWRESTLING", ListaClases(i)))
            .DanoArmas = val(GetVar(DatPath & "Balance.dat", "MODDANOARMAS", ListaClases(i)))
            .DanoProyectiles = val(GetVar(DatPath & "Balance.dat", "MODDANOPROYECTILES", ListaClases(i)))
            .DanoWrestling = val(GetVar(DatPath & "Balance.dat", "MODDANOWRESTLING", ListaClases(i)))
            .Escudo = val(GetVar(DatPath & "Balance.dat", "MODESCUDO", ListaClases(i)))

        End With

    Next i
    
    'Modificadores de Raza
    For i = 1 To NUMRAZAS

        With ModRaza(i)
            .Fuerza = val(GetVar(DatPath & "Balance.dat", "MODRAZA", ListaRazas(i) + "Fuerza"))
            .Agilidad = val(GetVar(DatPath & "Balance.dat", "MODRAZA", ListaRazas(i) + "Agilidad"))
            .Inteligencia = val(GetVar(DatPath & "Balance.dat", "MODRAZA", ListaRazas(i) + "Inteligencia"))
            .Carisma = val(GetVar(DatPath & "Balance.dat", "MODRAZA", ListaRazas(i) + "Carisma"))
            .Constitucion = val(GetVar(DatPath & "Balance.dat", "MODRAZA", ListaRazas(i) + "Constitucion"))

        End With

    Next i
    
    'Modificadores de Vida
    For i = 1 To NUMCLASES
        ModVida(i) = val(GetVar(DatPath & "Balance.dat", "MODVIDA", ListaClases(i)))
    Next i
    
    'Distribucion de Vida
    For i = 1 To 5
        DistribucionEnteraVida(i) = val(GetVar(DatPath & "Balance.dat", "DISTRIBUCION", "E" + CStr(i)))
    Next i

    For i = 1 To 4
        DistribucionSemienteraVida(i) = val(GetVar(DatPath & "Balance.dat", "DISTRIBUCION", "S" + CStr(i)))
    Next i
    
    'Extra
    PorcentajeRecuperoMana = val(GetVar(DatPath & "Balance.dat", "EXTRA", "PorcentajeRecuperoMana"))

    'Party
    ExponenteNivelParty = val(GetVar(DatPath & "Balance.dat", "PARTY", "ExponenteNivelParty"))
    
    ' Recompensas faccionarias
    For i = 1 To NUM_RANGOS_FACCION
        RecompensaFacciones(i - 1) = val(GetVar(DatPath & "Balance.dat", "RECOMPENSAFACCION", "Rango" & i))
    Next i
    
    If frmMain.Visible Then frmMain.txtStatus.Text = Date & " " & time & " - Se cargo con exito el archivo Balance.dat"

End Sub

Sub LoadObjCarpintero()
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************
    
    If frmMain.Visible Then frmMain.txtStatus.Text = "Cargando los objetos crafteables via Carpinteria"
    
    Dim n As Integer, lc As Integer
    
    n = val(GetVar(DatPath & "ObjCarpintero.dat", "INIT", "NumObjs"))
    
    ReDim Preserve ObjCarpintero(1 To n) As Integer
    
    For lc = 1 To n
        ObjCarpintero(lc) = val(GetVar(DatPath & "ObjCarpintero.dat", "Obj" & lc, "Index"))
    Next lc
    
    If frmMain.Visible Then frmMain.txtStatus.Text = Date & " " & time & " - Se cargo con exito los objetos crafteables via Carpinteria."

End Sub

Sub LoadObjArtesano()
    '***************************************************
    'Author: WyroX
    'Last Modification: 27/01/2020
    '***************************************************
    
    If frmMain.Visible Then frmMain.txtStatus.Text = "Cargando los objetos crafteables del Artesano"
    
    Dim n As Integer, lc As Integer
    
    n = val(GetVar(DatPath & "ObjArtesano.dat", "INIT", "NumObjs"))
    
    ReDim Preserve ObjArtesano(1 To n) As Integer
    
    For lc = 1 To n
        ObjArtesano(lc) = val(GetVar(DatPath & "ObjArtesano.dat", "Obj" & lc, "Index"))
    Next lc
    
    If frmMain.Visible Then frmMain.txtStatus.Text = Date & " " & time & " - Se cargo con exito los objetos crafteables del Artesano."

End Sub

Sub LoadOBJData()
    '*****************************************************************************************
    'Author: Unknown
    'Last Modification: 06/02/2020
    '03/02/2020: WyroX - Agrego nivel y skill minimo a ciertos objetos. Nuevas habilidades para anillos
    '06/02/2020: WyroX - MinSkill queda solo para barcos y lingotes (porque tienen una comprobacion especial).
    '                             - Skill requerido modificable para items equipables
    '*****************************************************************************************

    '###################################################
    '#               ATENCION PELIGRO                  #
    '###################################################
    '
    ' NO USAR GetVar PARA LEER DESDE EL OBJ.DAT !!!!
    '
    'El que ose desafiar esta LEY, se las tendra que ver
    'con migo. Para leer desde el OBJ.DAT se debera usar
    'la nueva clase clsLeerInis.
    '
    'Alejo
    '
    '###################################################

    'Call LogTarea("Sub LoadOBJData")

    On Error GoTo errHandler

    If frmMain.Visible Then frmMain.txtStatus.Text = "Cargando base de datos de los objetos."
    
    '*****************************************************************
    'Carga la lista de objetos
    '*****************************************************************
    Dim Object As Integer

    Dim leer   As clsIniManager

    Set leer = New clsIniManager
    
    Call leer.Initialize(DatPath & "Obj.dat")
    
    'obtiene el numero de obj
    NumObjDatas = val(leer.GetValue("INIT", "NumObjs"))
    
    frmCargando.cargar.min = 0
    frmCargando.cargar.max = NumObjDatas
    frmCargando.cargar.Value = 0
    
    ReDim Preserve ObjData(1 To NumObjDatas) As ObjData
    
    'Llena la lista
    For Object = 1 To NumObjDatas

        With ObjData(Object)
            .Name = leer.GetValue("OBJ" & Object, "Name")
            
            'Pablo (ToxicWaste) Log de Objetos.
            .Log = val(leer.GetValue("OBJ" & Object, "Log"))
            .NoLog = val(leer.GetValue("OBJ" & Object, "NoLog"))
            '07/09/07
            
            .GrhIndex = val(leer.GetValue("OBJ" & Object, "GrhIndex"))

            If .GrhIndex = 0 Then
                .GrhIndex = .GrhIndex

            End If
            
            .OBJType = val(leer.GetValue("OBJ" & Object, "ObjType"))
            .SubTipo = val(leer.GetValue("OBJ" & Object, "SubTipo"))
            
            .Newbie = val(leer.GetValue("OBJ" & Object, "Newbie"))
            
            Select Case .OBJType

                Case eOBJType.otArmadura
                    .Real = val(leer.GetValue("OBJ" & Object, "Real"))
                    .Caos = val(leer.GetValue("OBJ" & Object, "Caos"))
                    .LingH = val(leer.GetValue("OBJ" & Object, "LingH"))
                    .LingP = val(leer.GetValue("OBJ" & Object, "LingP"))
                    .LingO = val(leer.GetValue("OBJ" & Object, "LingO"))
                    .SkHerreria = val(leer.GetValue("OBJ" & Object, "SkHerreria"))
                    .MinLevel = val(leer.GetValue("OBJ" & Object, "MinLevel"))
                
                Case eOBJType.otEscudo
                    .ShieldAnim = val(leer.GetValue("OBJ" & Object, "Anim"))
                    .LingH = val(leer.GetValue("OBJ" & Object, "LingH"))
                    .LingP = val(leer.GetValue("OBJ" & Object, "LingP"))
                    .LingO = val(leer.GetValue("OBJ" & Object, "LingO"))
                    .SkHerreria = val(leer.GetValue("OBJ" & Object, "SkHerreria"))
                    .Real = val(leer.GetValue("OBJ" & Object, "Real"))
                    .Caos = val(leer.GetValue("OBJ" & Object, "Caos"))
                    .MinLevel = val(leer.GetValue("OBJ" & Object, "MinLevel"))
                
                Case eOBJType.otCasco
                    .CascoAnim = val(leer.GetValue("OBJ" & Object, "Anim"))
                    .LingH = val(leer.GetValue("OBJ" & Object, "LingH"))
                    .LingP = val(leer.GetValue("OBJ" & Object, "LingP"))
                    .LingO = val(leer.GetValue("OBJ" & Object, "LingO"))
                    .SkHerreria = val(leer.GetValue("OBJ" & Object, "SkHerreria"))
                    .Real = val(leer.GetValue("OBJ" & Object, "Real"))
                    .Caos = val(leer.GetValue("OBJ" & Object, "Caos"))
                    .MinLevel = val(leer.GetValue("OBJ" & Object, "MinLevel"))
                
                Case eOBJType.otWeapon
                    .WeaponAnim = val(leer.GetValue("OBJ" & Object, "Anim"))
                    .Apunala = val(leer.GetValue("OBJ" & Object, "Apunala"))
                    .Envenena = val(leer.GetValue("OBJ" & Object, "Envenena"))
                    .MaxHIT = val(leer.GetValue("OBJ" & Object, "MaxHIT"))
                    .MinHIT = val(leer.GetValue("OBJ" & Object, "MinHIT"))
                    .proyectil = val(leer.GetValue("OBJ" & Object, "Proyectil"))
                    .Municion = val(leer.GetValue("OBJ" & Object, "Municiones"))
                    .StaffPower = val(leer.GetValue("OBJ" & Object, "StaffPower"))
                    .StaffDamageBonus = val(leer.GetValue("OBJ" & Object, "StaffDamageBonus"))
                    .Refuerzo = val(leer.GetValue("OBJ" & Object, "Refuerzo"))
                    
                    .LingH = val(leer.GetValue("OBJ" & Object, "LingH"))
                    .LingP = val(leer.GetValue("OBJ" & Object, "LingP"))
                    .LingO = val(leer.GetValue("OBJ" & Object, "LingO"))
                    .SkHerreria = val(leer.GetValue("OBJ" & Object, "SkHerreria"))
                    .Real = val(leer.GetValue("OBJ" & Object, "Real"))
                    .Caos = val(leer.GetValue("OBJ" & Object, "Caos"))
                    
                    .WeaponRazaEnanaAnim = val(leer.GetValue("OBJ" & Object, "RazaEnanaAnim"))
                    .MinLevel = val(leer.GetValue("OBJ" & Object, "MinLevel"))
                
                Case eOBJType.otInstrumentos
                    .Snd1 = val(leer.GetValue("OBJ" & Object, "SND1"))
                    .Snd2 = val(leer.GetValue("OBJ" & Object, "SND2"))
                    .Snd3 = val(leer.GetValue("OBJ" & Object, "SND3"))
                    'Pablo (ToxicWaste)
                    .Real = val(leer.GetValue("OBJ" & Object, "Real"))
                    .Caos = val(leer.GetValue("OBJ" & Object, "Caos"))
                
                Case eOBJType.otPuertas, eOBJType.otBotellaVacia, eOBJType.otBotellaLlena
                    .IndexAbierta = val(leer.GetValue("OBJ" & Object, "IndexAbierta"))
                    .IndexCerrada = val(leer.GetValue("OBJ" & Object, "IndexCerrada"))
                    .IndexCerradaLlave = val(leer.GetValue("OBJ" & Object, "IndexCerradaLlave"))
                
                Case otPociones
                    .MaxModificador = val(leer.GetValue("OBJ" & Object, "MaxModificador"))
                    .MinModificador = val(leer.GetValue("OBJ" & Object, "MinModificador"))
                    .DuracionEfecto = val(leer.GetValue("OBJ" & Object, "DuracionEfecto"))
                
                Case eOBJType.otBarcos
                    .MinSkill = val(leer.GetValue("OBJ" & Object, "MinSkill"))
                    .MinLevel = val(leer.GetValue("OBJ" & Object, "MinLevel"))
                    .MaxHIT = val(leer.GetValue("OBJ" & Object, "MaxHIT"))
                    .MinHIT = val(leer.GetValue("OBJ" & Object, "MinHIT"))
                    .Real = val(leer.GetValue("OBJ" & Object, "Real"))
                    .Caos = val(leer.GetValue("OBJ" & Object, "Caos"))
                
                Case eOBJType.otFlechas
                    .MaxHIT = val(leer.GetValue("OBJ" & Object, "MaxHIT"))
                    .MinHIT = val(leer.GetValue("OBJ" & Object, "MinHIT"))
                    .Envenena = val(leer.GetValue("OBJ" & Object, "Envenena"))
                    .Paraliza = val(leer.GetValue("OBJ" & Object, "Paraliza"))
                    .MinLevel = val(leer.GetValue("OBJ" & Object, "MinLevel"))

                Case eOBJType.otMonturas
                    .MinSkill = val(leer.GetValue("OBJ" & Object, "MinSkill"))
                    .MinLevel = val(leer.GetValue("OBJ" & Object, "MinLevel"))
                    .MaxHIT = val(leer.GetValue("OBJ" & Object, "MaxHIT"))
                    .MinHIT = val(leer.GetValue("OBJ" & Object, "MinHIT"))

                Case eOBJType.otMinerales
                    .MinSkill = val(leer.GetValue("OBJ" & Object, "MinSkill"))

                Case eOBJType.otAnillo 'Pablo (ToxicWaste)
                    .LingH = val(leer.GetValue("OBJ" & Object, "LingH"))
                    .LingP = val(leer.GetValue("OBJ" & Object, "LingP"))
                    .LingO = val(leer.GetValue("OBJ" & Object, "LingO"))
                    .SkHerreria = val(leer.GetValue("OBJ" & Object, "SkHerreria"))
                    .MaxHIT = val(leer.GetValue("OBJ" & Object, "MaxHIT"))
                    .MinHIT = val(leer.GetValue("OBJ" & Object, "MinHIT"))
                    .MinLevel = val(leer.GetValue("OBJ" & Object, "MinLevel"))
                    '(WyroX)
                    .ImpideParalizar = val(leer.GetValue("OBJ" & Object, "ImpideParalizar")) <> 0
                    .ImpideAturdir = val(leer.GetValue("OBJ" & Object, "ImpideAturdir")) <> 0
                    .ImpideCegar = val(leer.GetValue("OBJ" & Object, "ImpideCegar")) <> 0
                    '(/WyroX)
                    
                Case eOBJType.otTeleport
                    .Radio = val(leer.GetValue("OBJ" & Object, "Radio"))
                    
                Case eOBJType.otForos
                    Call AddForum(leer.GetValue("OBJ" & Object, "ID"))
                    
                Case eOBJType.otPergaminos
                    .MinLevel = val(leer.GetValue("OBJ" & Object, "MinLevel"))

            End Select
            
            .Ropaje = val(leer.GetValue("OBJ" & Object, "NumRopaje"))
            .HechizoIndex = val(leer.GetValue("OBJ" & Object, "HechizoIndex"))
            
            .LingoteIndex = val(leer.GetValue("OBJ" & Object, "LingoteIndex"))
            
            .MineralIndex = val(leer.GetValue("OBJ" & Object, "MineralIndex"))
            
            .MaxHp = val(leer.GetValue("OBJ" & Object, "MaxHP"))
            .MinHp = val(leer.GetValue("OBJ" & Object, "MinHP"))
            
            .Mujer = val(leer.GetValue("OBJ" & Object, "Mujer"))
            .Hombre = val(leer.GetValue("OBJ" & Object, "Hombre"))
            
            .MinHam = val(leer.GetValue("OBJ" & Object, "MinHam"))
            .MinSed = val(leer.GetValue("OBJ" & Object, "MinAgu"))
            
            .MinDef = val(leer.GetValue("OBJ" & Object, "MINDEF"))
            .MaxDef = val(leer.GetValue("OBJ" & Object, "MAXDEF"))
            .def = (.MinDef + .MaxDef) / 2
            
            .RazaEnana = val(leer.GetValue("OBJ" & Object, "RazaEnana"))
            .RazaDrow = val(leer.GetValue("OBJ" & Object, "RazaDrow"))
            .RazaElfa = val(leer.GetValue("OBJ" & Object, "RazaElfa"))
            .RazaGnoma = val(leer.GetValue("OBJ" & Object, "RazaGnoma"))
            .RazaHumana = val(leer.GetValue("OBJ" & Object, "RazaHumana"))
            
            .valor = val(leer.GetValue("OBJ" & Object, "Valor"))
            
            .Crucial = val(leer.GetValue("OBJ" & Object, "Crucial"))
            
            .Cerrada = val(leer.GetValue("OBJ" & Object, "abierta"))

            If .Cerrada = 1 Then
                .Llave = val(leer.GetValue("OBJ" & Object, "Llave"))
                .Clave = val(leer.GetValue("OBJ" & Object, "Clave"))

            End If
            
            'Puertas y llaves
            .Clave = val(leer.GetValue("OBJ" & Object, "Clave"))
            
            .texto = leer.GetValue("OBJ" & Object, "Texto")
            .GrhSecundario = val(leer.GetValue("OBJ" & Object, "VGrande"))
            
            .Agarrable = val(leer.GetValue("OBJ" & Object, "Agarrable"))
            .ForoID = leer.GetValue("OBJ" & Object, "ID")
            
            .Acuchilla = val(leer.GetValue("OBJ" & Object, "Acuchilla"))
            
            .Guante = val(leer.GetValue("OBJ" & Object, "Guante"))
            
            'CHECK: !!! Esto es provisorio hasta que los de Dateo cambien los valores de string a numerico
            Dim i As Integer

            Dim n As Integer

            Dim S As String

            For i = 1 To NUMCLASES
                S = UCase$(leer.GetValue("OBJ" & Object, "CP" & i))
                n = 1

                Do While LenB(S) > 0 And UCase$(ListaClases(n)) <> S
                    n = n + 1
                Loop
                .ClaseProhibida(i) = IIf(LenB(S) > 0, n, 0)
            Next i
            
            .DefensaMagicaMax = val(leer.GetValue("OBJ" & Object, "DefensaMagicaMax"))
            .DefensaMagicaMin = val(leer.GetValue("OBJ" & Object, "DefensaMagicaMin"))
            
            .SkCarpinteria = val(leer.GetValue("OBJ" & Object, "SkCarpinteria"))
            
            If .SkCarpinteria > 0 Then .Madera = val(leer.GetValue("OBJ" & Object, "Madera"))
            .MaderaElfica = val(leer.GetValue("OBJ" & Object, "MaderaElfica"))
            
            ReDim .ItemCrafteo(1 To MAX_ITEMS_CRAFTEO) As CraftingItem
            
            For i = 1 To MAX_ITEMS_CRAFTEO
                S = leer.GetValue("OBJ" & Object, "ItemCrafteo" & i)
                If Len(S) <= 0 Then Exit For

                .ItemCrafteo(i).ObjIndex = val(ReadField(1, S, Asc("-")))
                .ItemCrafteo(i).amount = val(ReadField(2, S, Asc("-")))
            Next i
            
            If i > 1 Then
                ReDim Preserve .ItemCrafteo(1 To i - 1) As CraftingItem
            Else
                Erase .ItemCrafteo
            End If

           ' Skill minimo
            S = leer.GetValue("OBJ" & Object, "SkillRequerido")
            If Len(S) > 0 Then
                .SkillCantidad = val(ReadField(2, S, Asc("-")))
    
                S = Replace(UCase$(ReadField(1, S, Asc("-"))), "+", " ")
                For i = 1 To NUMSKILLS
                    If S = UCase$(SkillsNames(i)) Then
                        .SkillRequerido = i
                        Exit For
                    End If
                Next i
            End If

            'Bebidas
            .MinSta = val(leer.GetValue("OBJ" & Object, "MinST"))
            
            .NoSeCae = val(leer.GetValue("OBJ" & Object, "NoSeCae"))
            
            .Upgrade = val(leer.GetValue("OBJ" & Object, "Upgrade"))
            
            frmCargando.cargar.Value = frmCargando.cargar.Value + 1

        End With

    Next Object
    
    Set leer = Nothing
    
    ' Inicializo los foros faccionarios
    Call AddForum(FORO_CAOS_ID)
    Call AddForum(FORO_REAL_ID)

    If frmMain.Visible Then frmMain.txtStatus.Text = Date & " " & time & " - Se cargo base de datos de los objetos. Operacion Realizada con exito."
    
    Exit Sub
errHandler:
    MsgBox "error cargando objetos " & Err.Number & ": " & Err.description

End Sub

Sub LoadUserStats(ByVal Userindex As Integer, ByRef UserFile As clsIniManager)

    '*************************************************
    'Author: Unknown
    'Last modified: 11/19/2009
    '11/19/2009: Pato - Load the EluSkills and ExpSkills
    '*************************************************
    Dim LoopC As Long

    With UserList(Userindex)
        With .Stats

            For LoopC = 1 To NUMATRIBUTOS
                .UserAtributos(LoopC) = CByte(UserFile.GetValue("ATRIBUTOS", "AT" & LoopC))
                .UserAtributosBackUP(LoopC) = CByte(.UserAtributos(LoopC))
            Next LoopC
        
            For LoopC = 1 To NUMSKILLS
                .UserSkills(LoopC) = CByte(UserFile.GetValue("SKILLS", "SK" & LoopC))
                .EluSkills(LoopC) = CLng(UserFile.GetValue("SKILLS", "ELUSK" & LoopC))
                .ExpSkills(LoopC) = CLng(UserFile.GetValue("SKILLS", "EXPSK" & LoopC))
            Next LoopC
        
            For LoopC = 1 To MAXUSERHECHIZOS
                .UserHechizos(LoopC) = CInt(UserFile.GetValue("Hechizos", "H" & LoopC))
            Next LoopC
        
            .Gld = CLng(UserFile.GetValue("STATS", "GLD"))
            .Banco = CLng(UserFile.GetValue("STATS", "BANCO"))
        
            .MaxHp = CInt(UserFile.GetValue("STATS", "MaxHP"))
            .MinHp = CInt(UserFile.GetValue("STATS", "MinHP"))
        
            .MinSta = CInt(UserFile.GetValue("STATS", "MinSTA"))
            .MaxSta = CInt(UserFile.GetValue("STATS", "MaxSTA"))
        
            .MaxMAN = CInt(UserFile.GetValue("STATS", "MaxMAN"))
            .MinMAN = CInt(UserFile.GetValue("STATS", "MinMAN"))
        
            .MaxHIT = CInt(UserFile.GetValue("STATS", "MaxHIT"))
            .MinHIT = CInt(UserFile.GetValue("STATS", "MinHIT"))
        
            .MaxAGU = CByte(UserFile.GetValue("STATS", "MaxAGU"))
            .MinAGU = CByte(UserFile.GetValue("STATS", "MinAGU"))
        
            .MaxHam = CByte(UserFile.GetValue("STATS", "MaxHAM"))
            .MinHam = CByte(UserFile.GetValue("STATS", "MinHAM"))
        
            .SkillPts = CInt(UserFile.GetValue("STATS", "SkillPtsLibres"))
        
            .Exp = CDbl(UserFile.GetValue("STATS", "EXP"))
            .ELU = CLng(UserFile.GetValue("STATS", "ELU"))
            .ELV = CByte(UserFile.GetValue("STATS", "ELV"))
            
            .InventLevel = CByte(UserFile.GetValue("STATS", "InventLevel"))
        
            .UsuariosMatados = CLng(UserFile.GetValue("MUERTES", "UserMuertes"))
            .NPCsMuertos = CInt(UserFile.GetValue("MUERTES", "NpcsMuertes"))

        End With
    
        With .flags

            If CByte(UserFile.GetValue("CONSEJO", "PERTENECE")) Then .Privilegios = .Privilegios Or PlayerType.RoyalCouncil
        
            If CByte(UserFile.GetValue("CONSEJO", "PERTENECECAOS")) Then .Privilegios = .Privilegios Or PlayerType.ChaosCouncil

        End With

    End With

End Sub


Sub LoadUserInit(ByVal Userindex As Integer, ByRef UserFile As clsIniManager)

    '*************************************************
    'Author: Unknown
    'Last modified: 19/11/2019
    'Loads the Users RECORDs
    '23/01/2007 Pablo (ToxicWaste) - Agrego NivelIngreso, FechaIngreso, MatadosIngreso y NextRecompensa.
    '23/01/2007 Pablo (ToxicWaste) - Quito CriminalesMatados de Stats porque era redundante.
    '19/11/2019 Recox - Casteo todas las propiedades a su tipo de dato en Declares para evitar errores
    '*************************************************
    Dim LoopC As Long

    Dim ln    As String
    
    With UserList(Userindex)
        With .Faccion

            .CiudadanosMatados = CLng(UserFile.GetValue("FACCIONES", "CiudMatados"))
            .CriminalesMatados = CLng(UserFile.GetValue("FACCIONES", "CrimMatados"))
            .NeutralesMatados = CLng(UserFile.GetValue("FACCIONES", "NeutMatados"))
            .Bando = CByte(UserFile.GetValue("FACCIONES", "Bando"))
            .Jerarquia = CByte(UserFile.GetValue("FACCIONES", "Jerarquia"))
        End With
        
        With .flags
            .Muerto = CByte(UserFile.GetValue("FLAGS", "Muerto"))
            .Escondido = CByte(UserFile.GetValue("FLAGS", "Escondido"))
            
            .Hambre = CByte(UserFile.GetValue("FLAGS", "Hambre"))
            .Sed = CByte(UserFile.GetValue("FLAGS", "Sed"))
            .Desnudo = CByte(UserFile.GetValue("FLAGS", "Desnudo"))
            .Navegando = CByte(UserFile.GetValue("FLAGS", "Navegando"))
            .Envenenado = CByte(UserFile.GetValue("FLAGS", "Envenenado"))
            .Paralizado = CByte(UserFile.GetValue("FLAGS", "Paralizado"))
            
            'Matrix
            .lastMap = val(UserFile.GetValue("FLAGS", "LastMap"))

        End With

        .Counters.Pena = CLng(UserFile.GetValue("COUNTERS", "Pena"))
        .Counters.AsignedSkills = CByte(val(UserFile.GetValue("COUNTERS", "SkillsAsignados")))
        
        .Email = UserFile.GetValue("CONTACTO", "Email")
        
        'Cargando Amigos
        If UserFile.KeyExists("AMIGOS") Then

            For LoopC = 1 To MAXAMIGOS
                .Amigos(LoopC).Nombre = UserFile.GetValue("AMIGOS", "NOMBRE" & LoopC)
                .Amigos(LoopC).Ignorado = CByte(UserFile.GetValue("AMIGOS", "IGNORADO" & LoopC))
            Next LoopC

        Else ' Si no existe AMIGOS entonces se crean:

            Dim i As Long
            For i = 1 To MAXAMIGOS
                
                .Amigos(i).Nombre = vbNullString
                .Amigos(i).Ignorado = 0
                .Amigos(i).index = 0

            Next i

        End If

        .AccountHash = CStr(UserFile.GetValue("INIT", "AccountHash"))
        .Genero = CByte(UserFile.GetValue("INIT", "Genero"))
        .Clase = CByte(UserFile.GetValue("INIT", "Clase"))
        .raza = CByte(UserFile.GetValue("INIT", "Raza"))
        .Hogar = CByte(UserFile.GetValue("INIT", "Hogar"))
        .Char.heading = CInt(UserFile.GetValue("INIT", "Heading"))
        
        With .OrigChar
            .Head = CInt(UserFile.GetValue("INIT", "Head"))
            .body = CInt(UserFile.GetValue("INIT", "Body"))
            .WeaponAnim = CInt(UserFile.GetValue("INIT", "Arma"))
            .ShieldAnim = CInt(UserFile.GetValue("INIT", "Escudo"))
            .CascoAnim = CInt(UserFile.GetValue("INIT", "Casco"))
            .heading = eHeading.SOUTH

        End With
        
        #If ConUpTime Then
            .UpTime = CLng(UserFile.GetValue("INIT", "UpTime"))
        #End If

        .Desc = UserFile.GetValue("INIT", "Desc")
        
        .Pos.map = CInt(ReadField(1, UserFile.GetValue("INIT", "Position"), 45))
        .Pos.x = CInt(ReadField(2, UserFile.GetValue("INIT", "Position"), 45))
        .Pos.y = CInt(ReadField(3, UserFile.GetValue("INIT", "Position"), 45))
        
        .Invent.NroItems = CInt(UserFile.GetValue("Inventory", "CantidadItems"))
        
        .BancoInvent.NroItems = CInt(UserFile.GetValue("BancoInventory", "CantidadItems"))

        'Lista de objetos del banco
        For LoopC = 1 To MAX_BANCOINVENTORY_SLOTS
            ln = UserFile.GetValue("BancoInventory", "Obj" & LoopC)
            If (val(ReadField(1, ln, 45))) > NumObjDatas Then
                .BancoInvent.Object(LoopC).ObjIndex = 0
                .BancoInvent.Object(LoopC).amount = 0
            Else
                .BancoInvent.Object(LoopC).ObjIndex = CInt(ReadField(1, ln, 45))
                .BancoInvent.Object(LoopC).amount = CInt(ReadField(2, ln, 45))
            End If
        Next LoopC
        
        'Lista de objetos
        For LoopC = 1 To MAX_INVENTORY_SLOTS
            ln = UserFile.GetValue("Inventory", "Obj" & LoopC)
            If (val(ReadField(1, ln, 45))) > NumObjDatas Then
                .Invent.Object(LoopC).ObjIndex = 0
                .Invent.Object(LoopC).amount = 0
                .Invent.Object(LoopC).Equipped = 0
            Else
                .Invent.Object(LoopC).ObjIndex = val(ReadField(1, ln, 45))
                .Invent.Object(LoopC).amount = val(ReadField(2, ln, 45))
                .Invent.Object(LoopC).Equipped = val(ReadField(3, ln, 45))
            End If

        Next LoopC
        
        .Invent.WeaponEqpSlot = CByte(UserFile.GetValue("Inventory", "WeaponEqpSlot"))
        .Invent.ArmourEqpSlot = CByte(UserFile.GetValue("Inventory", "ArmourEqpSlot"))
        .Invent.EscudoEqpSlot = CByte(UserFile.GetValue("Inventory", "EscudoEqpSlot"))
        .Invent.CascoEqpSlot = CByte(UserFile.GetValue("Inventory", "CascoEqpSlot"))
        .Invent.BarcoSlot = CByte(UserFile.GetValue("Inventory", "BarcoSlot"))
        
        'Si no existe MonturaEqpSlot, se agrega al charfile.
        If Not UserFile.KeyExists("MonturaEqpSlot") Then
            .Invent.MonturaEqpSlot = 0
        Else
            .Invent.MonturaEqpSlot = CByte(UserFile.GetValue("Inventory", "MonturaEqpSlot"))
        End If
        
        .Invent.MunicionEqpSlot = CByte(UserFile.GetValue("Inventory", "MunicionSlot"))
        .Invent.AnilloEqpSlot = CByte(UserFile.GetValue("Inventory", "AnilloSlot"))
        
        .NroMascotas = CInt(UserFile.GetValue("MASCOTAS", "NroMascotas"))

        For LoopC = 1 To MAXMASCOTAS
            .MascotasType(LoopC) = val(UserFile.GetValue("MASCOTAS", "MAS" & LoopC))
        Next LoopC
        
        ln = UserFile.GetValue("Guild", "GUILDINDEX")

        If IsNumeric(ln) Then
            .GuildIndex = CInt(ln)
        Else
            .GuildIndex = 0

        End If

    End With

End Sub

Function GetVar(ByVal File As String, _
                ByVal Main As String, _
                ByVal Var As String, _
                Optional EmptySpaces As Long = 1024) As String
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    Dim sSpaces  As String ' This will hold the input that the program will retrieve

    Dim szReturn As String ' This will be the defaul value if the string is not found
      
    szReturn = vbNullString
      
    sSpaces = Space$(EmptySpaces) ' This tells the computer how long the longest string can be
      
    GetPrivateProfileString Main, Var, szReturn, sSpaces, EmptySpaces, File
      
    GetVar = RTrim$(sSpaces)
    GetVar = Left$(GetVar, Len(GetVar) - 1)
  
End Function

Sub CargarBackUp()
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    If frmMain.Visible Then frmMain.txtStatus.Text = "Cargando backup."
    
    Dim map       As Integer

    Dim tFileName As String
    
    On Error GoTo man
        
    NumMaps = val(GetVar(DatPath & "Map.dat", "INIT", "NumMaps"))
        
    frmCargando.cargar.min = 0
    frmCargando.cargar.max = NumMaps
    frmCargando.cargar.Value = 0
        
    MapPath = GetVar(DatPath & "Map.dat", "INIT", "MapPath")
        
    ReDim MapData(1 To NumMaps, XMinMapSize To XMaxMapSize, YMinMapSize To YMaxMapSize) As MapBlock
    ReDim MapInfo(1 To NumMaps) As MapInfo
        
    For map = 1 To NumMaps

        If val(GetVar(App.path & MapPath & "Mapa" & map & ".Dat", "Mapa" & map, "BackUp")) <> 0 Then
            tFileName = App.path & "\WorldBackUp\Mapa" & map
                
            If Not FileExist(tFileName & ".*") Then 'Miramos que exista al menos uno de los 3 archivos, sino lo cargamos de la carpeta de los mapas
                tFileName = App.path & MapPath & "Mapa" & map

            End If

        Else
            tFileName = App.path & MapPath & "Mapa" & map

        End If
            
        Call CargarMapa(map, tFileName)
            
        frmCargando.cargar.Value = frmCargando.cargar.Value + 1
        DoEvents
    Next map
    
    Exit Sub

    If frmMain.Visible Then frmMain.txtStatus.Text = Date & " " & time & " - Se termino de cargar el backup."

man:
    MsgBox ("Error durante la carga de mapas, el mapa " & map & " contiene errores")
    Call LogError(Date & " " & Err.description & " " & Err.HelpContext & " " & Err.HelpFile & " " & Err.source)
 
End Sub

Sub LoadMapData()
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    If frmMain.Visible Then frmMain.txtStatus.Text = "Cargando mapas..."
    
    Dim map       As Integer

    Dim tFileName As String
    
    On Error GoTo man
        
    NumMaps = val(GetVar(DatPath & "Map.dat", "INIT", "NumMaps"))
        
    frmCargando.cargar.min = 0
    frmCargando.cargar.max = NumMaps
    frmCargando.cargar.Value = 0
        
    MapPath = GetVar(DatPath & "Map.dat", "INIT", "MapPath")
        
    ReDim MapData(1 To NumMaps, XMinMapSize To XMaxMapSize, YMinMapSize To YMaxMapSize) As MapBlock
    ReDim MapInfo(1 To NumMaps) As MapInfo
          
    For map = 1 To NumMaps
            
        tFileName = App.path & MapPath & "Mapa" & map
        Call CargarMapa(map, tFileName)
            
        frmCargando.cargar.Value = frmCargando.cargar.Value + 1
        DoEvents
    Next map
    
    Exit Sub

    If frmMain.Visible Then frmMain.txtStatus.Text = Date & " " & time & " - Se cargaron todos los mapas. Operacion Realizada con exito."

man:
    MsgBox ("Error durante la carga de mapas, el mapa " & map & " contiene errores")
    Call LogError(Date & " " & Err.description & " " & Err.HelpContext & " " & Err.HelpFile & " " & Err.source)

End Sub

Public Sub CargarMapa(ByVal map As Long, ByRef MAPFl As String)
    '***************************************************
    'Author: Unknown
    'Last Modification: 10/08/2010
    '10/08/2010 - Pato: Implemento el clsByteBuffer y el clsIniManager para la carga de mapa
    '***************************************************

    On Error GoTo errh

    Dim hFile     As Integer
    Dim x         As Long
    Dim y         As Long
    Dim ByFlags   As Byte
    Dim npcfile   As String
    
    Dim leer      As clsIniManager
    Dim MapReader As clsByteBuffer
    Dim InfReader As clsByteBuffer

    Dim buff()    As Byte
    
    Set MapReader = New clsByteBuffer
    Set InfReader = New clsByteBuffer
    Set leer = New clsIniManager
    
    npcfile = DatPath & "NPCs.dat"
    
    hFile = FreeFile
    
    'Leemos el archivo ".MAP"
    Open MAPFl & ".map" For Binary As #hFile
        Seek hFile, 1
        ReDim buff(LOF(hFile) - 1) As Byte
        Get #hFile, , buff
    Close hFile
    
    Call MapReader.initializeReader(buff)

    'Leemos el archivo ".INF"
    Open MAPFl & ".inf" For Binary As #hFile
        Seek hFile, 1
        ReDim buff(LOF(hFile) - 1) As Byte
        Get #hFile, , buff
    Close hFile
    
    Call InfReader.initializeReader(buff)
    
    'map Header
    MapInfo(map).MapVersion = MapReader.getInteger
    
    With MiCabecera
        .Desc = MapReader.getString(Len(MiCabecera.Desc))
        .crc = MapReader.getLong
        .MagicWord = MapReader.getLong
    End With
    
    Call MapReader.getDouble

    'inf Header
    Call InfReader.getDouble
    Call InfReader.getInteger

    For y = YMinMapSize To YMaxMapSize
        For x = XMinMapSize To XMaxMapSize

            With MapData(map, x, y)
                '.map file
                ByFlags = MapReader.getByte

                If ByFlags And 1 Then .Blocked = 1
                
                'Layer 1
                .Graphic(1) = MapReader.getLong

                'Layer 2 used?
                If ByFlags And 2 Then .Graphic(2) = MapReader.getLong

                'Layer 3 used?
                If ByFlags And 4 Then .Graphic(3) = MapReader.getLong

                'Layer 4 used?
                If ByFlags And 8 Then .Graphic(4) = MapReader.getLong

                'Trigger used?
                If ByFlags And 16 Then .trigger = MapReader.getInteger

                '.inf file
                ByFlags = InfReader.getByte

                If ByFlags And 1 Then
                    .TileExit.map = InfReader.getInteger
                    .TileExit.x = InfReader.getInteger
                    .TileExit.y = InfReader.getInteger
                End If

                If ByFlags And 2 Then
                    'Get and make NPC
                    .NpcIndex = InfReader.getInteger

                    If .NpcIndex > 0 Then

                        'Si el npc debe hacer respawn en la pos
                        'original la guardamos
                        If val(GetVar(npcfile, "NPC" & .NpcIndex, "PosOrig")) = 1 Then
                            .NpcIndex = OpenNPC(.NpcIndex)
                            Npclist(.NpcIndex).Orig.map = map
                            Npclist(.NpcIndex).Orig.x = x
                            Npclist(.NpcIndex).Orig.y = y
                        Else
                            .NpcIndex = OpenNPC(.NpcIndex)

                        End If

                        Npclist(.NpcIndex).Pos.map = map
                        Npclist(.NpcIndex).Pos.x = x
                        Npclist(.NpcIndex).Pos.y = y

                        Call MakeNPCChar(True, 0, .NpcIndex, map, x, y)

                    End If

                End If

                If ByFlags And 4 Then
                    'Get and make Object
                    .ObjInfo.ObjIndex = InfReader.getInteger
                    .ObjInfo.amount = InfReader.getInteger

                End If

            End With

        Next x
    Next y
    
    Call leer.Initialize(MAPFl & ".dat")
    
    With MapInfo(map)
        .Name = leer.GetValue("Mapa" & map, "Name")
        .Music = leer.GetValue("Mapa" & map, "MusicNum")
        .MusicMp3 = leer.GetValue("Mapa" & map, "MusicNumMp3")
        
        .StartPos.map = val(ReadField(1, leer.GetValue("Mapa" & map, "StartPos"), Asc("-")))
        .StartPos.x = val(ReadField(2, leer.GetValue("Mapa" & map, "StartPos"), Asc("-")))
        .StartPos.y = val(ReadField(3, leer.GetValue("Mapa" & map, "StartPos"), Asc("-")))
        
        .OnDeathGoTo.map = val(ReadField(1, leer.GetValue("Mapa" & map, "OnDeathGoTo"), Asc("-")))
        .OnDeathGoTo.x = val(ReadField(2, leer.GetValue("Mapa" & map, "OnDeathGoTo"), Asc("-")))
        .OnDeathGoTo.y = val(ReadField(3, leer.GetValue("Mapa" & map, "OnDeathGoTo"), Asc("-")))
        
        .MagiaSinEfecto = val(leer.GetValue("Mapa" & map, "MagiaSinEfecto"))
        .InviSinEfecto = val(leer.GetValue("Mapa" & map, "InviSinEfecto"))
        .ResuSinEfecto = val(leer.GetValue("Mapa" & map, "ResuSinEfecto"))
        .OcultarSinEfecto = val(leer.GetValue("Mapa" & map, "OcultarSinEfecto"))
        .InvocarSinEfecto = val(leer.GetValue("Mapa" & map, "InvocarSinEfecto"))
        
        .NoEncriptarMP = val(leer.GetValue("Mapa" & map, "NoEncriptarMP"))

        .RoboNpcsPermitido = val(leer.GetValue("Mapa" & map, "RoboNpcsPermitido"))
        
        If val(leer.GetValue("Mapa" & map, "Pk")) = 0 Then
            .Pk = True
        Else
            .Pk = False

        End If
        
        .Terreno = TerrainStringToByte(leer.GetValue("Mapa" & map, "Terreno"))
        .Zona = leer.GetValue("Mapa" & map, "Zona")
        .Restringir = RestrictStringToByte(leer.GetValue("Mapa" & map, "Restringir"))
        .BackUp = val(leer.GetValue("Mapa" & map, "BACKUP"))

    End With
    
    Set MapReader = Nothing
    Set InfReader = Nothing
    Set leer = Nothing
    
    Erase buff
    Exit Sub

errh:
    Call LogError("Error cargando mapa: " & map & " - Pos: " & x & "," & y & "." & Err.description)

    Set MapReader = Nothing
    Set InfReader = Nothing
    Set leer = Nothing

End Sub

Sub LoadSini()
'***************************************************
'Author: Unknown
'Last Modification: 13/11/2019 (Recox)
'CHOTS: Database params
'Cucsifae: Agregados multiplicadores exp y oro
'CHOTS: Agregado multiplicador oficio
'CHOTS: Agregado min y max Dados
'Jopi: Uso de clsIniManager para cargar los valores.
'Recox: Cargamos si el centinela esta activo o no.
'***************************************************

    Dim Temporal As Long
    
    Dim Lector As clsIniManager
    Set Lector = New clsIniManager
    
    If frmMain.Visible Then
        frmMain.txtStatus.Text = "Cargando info de inicio del server."
    End If
    
    Call Lector.Initialize(IniPath & "Server.ini")
    
    BootDelBackUp = CBool(val(Lector.GetValue("INIT", "IniciarDesdeBackUp")))
    
    'Misc
    Puerto = val(Lector.GetValue("INIT", "StartPort"))
    LastSockListen = val(Lector.GetValue("INIT", "LastSockListen"))
    HideMe = CBool(Lector.GetValue("INIT", "Hide"))
    AllowMultiLogins = CBool(val(Lector.GetValue("INIT", "AllowMultiLogins")))
    IdleLimit = val(Lector.GetValue("INIT", "IdleLimit"))
    LimiteConexionesPorIp = val(Lector.GetValue("INIT", "LimiteConexionesPorIp"))
    
    'Lee la version correcta del cliente
    ULTIMAVERSION = Lector.GetValue("INIT", "VersionBuildCliente")
    
    STAT_MAXELV = val(Lector.GetValue("INIT", "NivelMaximo"))
    
    ExpMultiplier = val(Lector.GetValue("INIT", "ExpMulti"))
    OroMultiplier = val(Lector.GetValue("INIT", "OroMulti"))
    OficioMultiplier = val(Lector.GetValue("INIT", "OficioMulti"))
    DiceMinimum = val(Lector.GetValue("INIT", "MinDados"))
    DiceMaximum = val(Lector.GetValue("INIT", "MaxDados"))
    
    DropItemsAlMorir = CBool(Lector.GetValue("INIT", "DropItemsAlMorir"))
    
    ArtesaniaCosto = val(Lector.GetValue("INIT", "ArtesaniaCosto"))

    'Esto es para ver si el centinela esta activo o no.
    isCentinelaActivated = CBool(val(Lector.GetValue("INIT", "CentinelaAuditoriaTrabajoActivo")))

    PuedeCrearPersonajes = val(Lector.GetValue("INIT", "PuedeCrearPersonajes"))
    ServerSoloGMs = val(Lector.GetValue("INIT", "ServerSoloGMs"))
    
    MAPA_PRETORIANO = val(Lector.GetValue("CLAN-PRETORIANO", "Mapa"))
    PRETORIANO_X = val(Lector.GetValue("CLAN-PRETORIANO", "X"))
    PRETORIANO_Y = val(Lector.GetValue("CLAN-PRETORIANO", "Y"))
    
    EnTesting = CBool(Lector.GetValue("INIT", "Testing"))
    
    ContadorAntiPiquete = val(Lector.GetValue("INIT", "ContadorAntiPiquete"))
    MinutosCarcelPiquete = val(Lector.GetValue("INIT", "MinutosCarcelPiquete"))

    'Usar Mundo personalizado / Use custom world
    UsarMundoPropio = CBool(Lector.GetValue("MUNDO", "UsarMundoPropio"))
    
    OroDirectoABille = val(Lector.GetValue("INIT", "OroDirectoABille"))
    
    'Inventario Inicial
    InventarioUsarConfiguracionPersonalizada = CBool(val(Lector.GetValue("INVENTARIO", "InventarioUsarConfiguracionPersonalizada")))

    'Atributos Iniciales
    EstadisticasInicialesUsarConfiguracionPersonalizada = CBool(val(Lector.GetValue("ESTADISTICASINICIALESPJ", "Activado")))

    'Intervalos
    SanaIntervaloSinDescansar = val(Lector.GetValue("INTERVALOS", "SanaIntervaloSinDescansar"))
    StaminaIntervaloSinDescansar = val(Lector.GetValue("INTERVALOS", "StaminaIntervaloSinDescansar"))
    SanaIntervaloDescansar = val(Lector.GetValue("INTERVALOS", "SanaIntervaloDescansar"))
    StaminaIntervaloDescansar = val(Lector.GetValue("INTERVALOS", "StaminaIntervaloDescansar"))
    IntervaloSed = val(Lector.GetValue("INTERVALOS", "IntervaloSed"))
    IntervaloHambre = val(Lector.GetValue("INTERVALOS", "IntervaloHambre"))
    IntervaloVeneno = val(Lector.GetValue("INTERVALOS", "IntervaloVeneno"))
    IntervaloParalizado = val(Lector.GetValue("INTERVALOS", "IntervaloParalizado"))
    IntervaloInvisible = val(Lector.GetValue("INTERVALOS", "IntervaloInvisible"))
    IntervaloFrio = val(Lector.GetValue("INTERVALOS", "IntervaloFrio"))
    IntervaloWavFx = val(Lector.GetValue("INTERVALOS", "IntervaloWAVFX"))
    IntervaloInvocacion = val(Lector.GetValue("INTERVALOS", "IntervaloInvocacion"))
    IntervaloParaConexion = val(Lector.GetValue("INTERVALOS", "IntervaloParaConexion"))
    IntervaloUserPuedeCastear = val(Lector.GetValue("INTERVALOS", "IntervaloLanzaHechizo"))
    IntervaloUserPuedeTrabajar = val(Lector.GetValue("INTERVALOS", "IntervaloTrabajo"))
    IntervaloUserPuedeAtacar = val(Lector.GetValue("INTERVALOS", "IntervaloUserPuedeAtacar"))
    
    'TODO : Agregar estos intervalos al form!!!
    IntervaloMagiaGolpe = val(Lector.GetValue("INTERVALOS", "IntervaloMagiaGolpe"))
    IntervaloGolpeMagia = val(Lector.GetValue("INTERVALOS", "IntervaloGolpeMagia"))
    IntervaloGolpeUsar = val(Lector.GetValue("INTERVALOS", "IntervaloGolpeUsar"))
    
    '&&&&&&&&&&&&&&&&&&&&& TIMERS &&&&&&&&&&&&&&&&&&&&&&&
    IntervaloPuedeSerAtacado = val(Lector.GetValue("TIMERS", "IntervaloPuedeSerAtacado"))
    IntervaloAtacable = val(Lector.GetValue("TIMERS", "IntervaloAtacable"))
    IntervaloOwnedNpc = val(Lector.GetValue("TIMERS", "IntervaloOwnedNpc"))

    MinutosWs = val(Lector.GetValue("INTERVALOS", "IntervaloWS"))

    If MinutosWs < 60 Then MinutosWs = 180
    
    MinutosGuardarUsuarios = val(Lector.GetValue("INTERVALOS", "IntervaloGuardarUsuarios"))
    IntervaloCerrarConexion = val(Lector.GetValue("INTERVALOS", "IntervaloCerrarConexion"))
    IntervaloUserPuedeUsar = val(Lector.GetValue("INTERVALOS", "IntervaloUserPuedeUsar"))
    IntervaloFlechasCazadores = val(Lector.GetValue("INTERVALOS", "IntervaloFlechasCazadores"))
    
    IntervaloOculto = val(Lector.GetValue("INTERVALOS", "IntervaloOculto"))
    
    '&&&&&&&&&&&&&&&&&&&&& SUERTE &&&&&&&&&&&&&&&&&&&&&&&
    DificultadPescar = val(Lector.GetValue("DIFICULTAD", "DificultadPescar"))
    DificultadTalar = val(Lector.GetValue("DIFICULTAD", "DificultadTalar"))
    DificultadMinar = val(Lector.GetValue("DIFICULTAD", "DificultadMinar"))
    '&&&&&&&&&&&&&&&&&&&&& FIN TIMERS &&&&&&&&&&&&&&&&&&&&&&&
      
    '&&&&&&&&&&&&&&&&&&&&& Evento Pesca &&&&&&&&&&&&&&&&&&&&&&&
    PescaEvent.Activado = val(Lector.GetValue("EVENTOPESCA", "Activado"))
    PescaEvent.Tiempo = val(Lector.GetValue("EVENTOPESCA", "Tiempo"))
    PescaEvent.CantidadDeZonas = val(Lector.GetValue("EVENTOPESCA", "CantidadDeZonas"))
    Call LoadPeces
    '&&&&&&&&&&&&&&&&&&&&& Fin Evento Pesca &&&&&&&&&&&&&&&&&&&&&&&

    RecordUsuariosOnline = val(Lector.GetValue("INIT", "Record"))

    ' HappyHour
    Dim lDayNumberTemp As Long
    Dim sDayName As String
    
    iniHappyHourActivado = CBool(val(Lector.GetValue("HAPPYHOUR", "Activado")))
    For lDayNumberTemp = 1 To 7
        sDayName = Lector.GetValue("HAPPYHOUR", "Dia" & lDayNumberTemp)
        HappyHourDays(lDayNumberTemp).Hour = val(ReadField(1, sDayName, 45)) ' GSZAO
        HappyHourDays(lDayNumberTemp).Multi = val(ReadField(2, sDayName, 45)) ' 0.13.5
        
        If HappyHourDays(lDayNumberTemp).Hour < 0 Or HappyHourDays(lDayNumberTemp).Hour > 23 Then
            HappyHourDays(lDayNumberTemp).Hour = 20 ' Hora de 0 a 23.
        End If
        
        If HappyHourDays(lDayNumberTemp).Multi < 0 Then
            HappyHourDays(lDayNumberTemp).Multi = 0
        End If
    Next

    'Conexion con la API hecha en Node.js
    'Mas info aqui: https://github.com/ao-libre/ao-api-server/
    ConexionAPI = CBool(Lector.GetValue("CONEXIONAPI", "Activado"))
    ApiUrlServer = Lector.GetValue("CONEXIONAPI", "UrlServer")
    ApiPath = Lector.GetValue("CONEXIONAPI", "ApiPath")

    'CHOTS | Database
    Database_Enabled = CBool(val(Lector.GetValue("DATABASE", "Enabled")))
    Database_DataSource = Lector.GetValue("DATABASE", "DSN")
    Database_Host = Lector.GetValue("DATABASE", "Host")
    Database_Name = Lector.GetValue("DATABASE", "Name")
    Database_Username = Lector.GetValue("DATABASE", "Username")
    Database_Password = Lector.GetValue("DATABASE", "Password")
      
    'Max users
    Temporal = val(Lector.GetValue("INIT", "MaxUsers"))

    If MaxUsers = 0 Then
        MaxUsers = Temporal
        ReDim UserList(1 To MaxUsers) As User

    End If
    
    '&&&&&&&&&&&&&&&&&&&&& BALANCE &&&&&&&&&&&&&&&&&&&&&&&
    'Se agrego en LoadBalance y en el Balance.dat
    'PorcentajeRecuperoMana = val(Lector.GetValue("BALANCE", "PorcentajeRecuperoMana"))
    
    ''&&&&&&&&&&&&&&&&&&&&& FIN BALANCE &&&&&&&&&&&&&&&&&&&&&&&
    Call Statistics.Initialize

    'En caso que usemos mundo propio, cargamos el mapa y la coordeanas donde se hara el spawn inicial'
    If UsarMundoPropio Then
        CustomSpawnMap.map = Lector.GetValue("MUNDO", "Mapa")
        CustomSpawnMap.x = Lector.GetValue("MUNDO", "X")
        CustomSpawnMap.y = Lector.GetValue("MUNDO", "Y")
    End If
    
    Set Lector = Nothing
    
    Set ConsultaPopular = New ConsultasPopulares
    Call ConsultaPopular.LoadData
    
    ' Admins
    Call loadAdministrativeUsers

    If frmMain.Visible Then frmMain.txtStatus.Text = Date & " " & time & " - Se cargo la info de inicio del server (Sinfo.ini)"
    
End Sub

Sub CargarCiudades()
    
    '***************************************************
    'Author: Jopi
    'Last Modification: 15/05/2019 (Jopi)
    'Jopi: Uso de clsIniManager para cargar los valores.
    '***************************************************
    If frmMain.Visible Then frmMain.txtStatus.Text = "Cargando Ciudades.dat"
    
    Dim Lector As clsIniManager: Set Lector = New clsIniManager
    
    Call Lector.Initialize(DatPath & "Ciudades.dat")
        
        With Ullathorpe
            .map = Lector.GetValue("Ullathorpe", "Mapa")
            .x = Lector.GetValue("Ullathorpe", "X")
            .y = Lector.GetValue("Ullathorpe", "Y")
        End With
        
        With Nix
            .map = Lector.GetValue("Nix", "Mapa")
            .x = Lector.GetValue("Nix", "X")
            .y = Lector.GetValue("Nix", "Y")
        End With
        
        With Banderbill
            .map = Lector.GetValue("Banderbill", "Mapa")
            .x = Lector.GetValue("Banderbill", "X")
            .y = Lector.GetValue("Banderbill", "Y")
        End With
      
        With Lindos
            .map = Lector.GetValue("Lindos", "Mapa")
            .x = Lector.GetValue("Lindos", "X")
            .y = Lector.GetValue("Lindos", "Y")
        End With
        
        With Arghal
            .map = Lector.GetValue("Arghal", "Mapa")
            .x = Lector.GetValue("Arghal", "X")
            .y = Lector.GetValue("Arghal", "Y")
        End With
        
        With Arkhein
            .map = Lector.GetValue("Arkhein", "Mapa")
            .x = Lector.GetValue("Arkhein", "X")
            .y = Lector.GetValue("Arkhein", "Y")
        End With
        
        With Nemahuak
            .map = Lector.GetValue("Nemahuak", "Mapa")
            .x = Lector.GetValue("Nemahuak", "X")
            .y = Lector.GetValue("Nemahuak", "Y")
        End With
        
        With Prision
            .map = Lector.GetValue("Prision", "Mapa")
            .x = Lector.GetValue("Prision", "X")
            .y = Lector.GetValue("Prision", "Y")
        End With
        
        With Libertad
            .map = Lector.GetValue("Prision-Afuera", "Mapa")
            .x = Lector.GetValue("Prision-Afuera", "X")
            .y = Lector.GetValue("Prision-Afuera", "Y")
        End With

        With Gotland
            .map = Lector.GetValue("Gotland", "Mapa")
            .x = Lector.GetValue("Gotland", "X")
            .y = Lector.GetValue("Gotland", "Y")
        End With

        With Perdida
            .map = Lector.GetValue("Perdida", "Mapa")
            .x = Lector.GetValue("Perdida", "X")
            .y = Lector.GetValue("Perdida", "Y")
        End With

        With Totem
            .map = Lector.GetValue("Totem", "Mapa")
            .x = Lector.GetValue("Totem", "X")
            .y = Lector.GetValue("Totem", "Y")
        End With

    Set Lector = Nothing
    
    Ciudades(eCiudad.cUllathorpe) = Ullathorpe
    Ciudades(eCiudad.cNix) = Nix
    Ciudades(eCiudad.cBanderbill) = Banderbill
    Ciudades(eCiudad.cLindos) = Lindos
    Ciudades(eCiudad.cArghal) = Arghal
    Ciudades(eCiudad.cArkhein) = Arkhein
    Ciudades(eCiudad.cGotland) = Gotland
    Ciudades(eCiudad.cPerdida) = Perdida
    Ciudades(eCiudad.cTotem) = Totem

    If frmMain.Visible Then frmMain.txtStatus.Text = Date & " " & time & " - Se cargaron las ciudades.dat"

End Sub

Sub WriteVar(ByVal File As String, _
             ByVal Main As String, _
             ByVal Var As String, _
             ByVal Value As String)
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    'Escribe VAR en un archivo
    '***************************************************

    writeprivateprofilestring Main, Var, Value, File
    
End Sub

Sub SaveUserToCharfile(ByVal Userindex As Integer, Optional ByVal SaveTimeOnline As Boolean = True)
    '*************************************************
    'Author: Unknown
    'Last modified: 10/10/2010 (Pato)
    'Saves the Users RECORDs
    '23/01/2007 Pablo (ToxicWaste) - Agrego NivelIngreso, FechaIngreso, MatadosIngreso y NextRecompensa.
    '11/19/2009: Pato - Save the EluSkills and ExpSkills
    '12/01/2010: ZaMa - Los druidas pierden la inmunidad de ser atacados cuando pierden el efecto del mimetismo.
    '10/10/2010: Pato - Saco el WriteVar e implemento la clase clsIniManager
    '18/09/2018: CHOTS - Nuevo nombre de la funcion, solo realiza el grabado
    '19/11/2019: Recox - Cambie el casteo de muchas propiedades, para evitar y arreglar errores
    '*************************************************

    On Error GoTo ErrorHandler

    Dim Manager  As clsIniManager

    Dim Existe   As Boolean

    Dim UserFile As String

    With UserList(Userindex)

        UserFile = CharPath & UCase$(.Name) & ".chr"
    
        Set Manager = New clsIniManager
    
        If FileExist(UserFile) Then
            Call Manager.Initialize(UserFile)
        
            If FileExist(UserFile & ".bk") Then Call Kill(UserFile & ".bk")
            Name UserFile As UserFile & ".bk"
        
            Existe = True

        End If
    
        Dim LoopC As Long
    
        Call Manager.ChangeValue("FLAGS", "Muerto", CByte(.flags.Muerto))
        Call Manager.ChangeValue("FLAGS", "Escondido", CByte(.flags.Escondido))
        Call Manager.ChangeValue("FLAGS", "Hambre", CByte(.flags.Hambre))
        Call Manager.ChangeValue("FLAGS", "Sed", CByte(.flags.Sed))
        Call Manager.ChangeValue("FLAGS", "Desnudo", CByte(.flags.Desnudo))
        Call Manager.ChangeValue("FLAGS", "Ban", CByte(.flags.Ban))
        Call Manager.ChangeValue("FLAGS", "Navegando", CByte(.flags.Navegando))
        Call Manager.ChangeValue("FLAGS", "Envenenado", CByte(.flags.Envenenado))
        Call Manager.ChangeValue("FLAGS", "Paralizado", CByte(.flags.Paralizado))
        'Matrix
        Call Manager.ChangeValue("FLAGS", "LastMap", CInt(.flags.lastMap))
    
        Call Manager.ChangeValue("CONSEJO", "PERTENECE", IIf(.flags.Privilegios And PlayerType.RoyalCouncil, "1", "0"))
        Call Manager.ChangeValue("CONSEJO", "PERTENECECAOS", IIf(.flags.Privilegios And PlayerType.ChaosCouncil, "1", "0"))
    
        For LoopC = 1 To MAXAMIGOS
            Call Manager.ChangeValue("AMIGOS", "Nombre" & LoopC, .Amigos(LoopC).Nombre)
            Call Manager.ChangeValue("AMIGOS", "IGNORADO" & LoopC, CStr(.Amigos(LoopC).Ignorado))
        Next LoopC

        Call Manager.ChangeValue("COUNTERS", "Pena", CLng(.Counters.Pena))
        Call Manager.ChangeValue("COUNTERS", "SkillsAsignados", CByte(.Counters.AsignedSkills))
    
        Call Manager.ChangeValue("FACCIONES", "CiudMatados", CLng(.Faccion.CiudadanosMatados))
        Call Manager.ChangeValue("FACCIONES", "CrimMatados", CLng(.Faccion.CriminalesMatados))
        Call Manager.ChangeValue("FACCIONES", "NeutMatados", CLng(.Faccion.NeutralesMatados))
        Call Manager.ChangeValue("FACCIONES", "Bando", CByte(.Faccion.Bando))
        Call Manager.ChangeValue("FACCIONES", "Jerarquia", CLng(.Faccion.Jerarquia))
        
    
        'Fueron modificados los atributos del usuario?
        If Not .flags.TomoPocion Then

            For LoopC = 1 To UBound(.Stats.UserAtributos)
                Call Manager.ChangeValue("ATRIBUTOS", "AT" & LoopC, CStr(.Stats.UserAtributos(LoopC)))
            Next LoopC

        Else

            For LoopC = 1 To UBound(.Stats.UserAtributos)
                '.Stats.UserAtributos(LoopC) = .Stats.UserAtributosBackUP(LoopC)
                Call Manager.ChangeValue("ATRIBUTOS", "AT" & LoopC, CStr(.Stats.UserAtributosBackUP(LoopC)))
            Next LoopC

        End If
    
        For LoopC = 1 To UBound(.Stats.UserSkills)
            Call Manager.ChangeValue("SKILLS", "SK" & LoopC, CStr(.Stats.UserSkills(LoopC)))
            Call Manager.ChangeValue("SKILLS", "ELUSK" & LoopC, CStr(.Stats.EluSkills(LoopC)))
            Call Manager.ChangeValue("SKILLS", "EXPSK" & LoopC, CStr(.Stats.ExpSkills(LoopC)))
        Next LoopC
    
        Call Manager.ChangeValue("CONTACTO", "Email", CStr(.Email))
    
        Call Manager.ChangeValue("INIT", "AccountHash", CStr(.AccountHash))
        Call Manager.ChangeValue("INIT", "Genero", CByte(.Genero))
        Call Manager.ChangeValue("INIT", "Raza", CByte(.raza))
        Call Manager.ChangeValue("INIT", "Hogar", CByte(.Hogar))
        Call Manager.ChangeValue("INIT", "Clase", CByte(.Clase))
        Call Manager.ChangeValue("INIT", "Desc", CStr(.Desc))
    
        Call Manager.ChangeValue("INIT", "Heading", CByte(.Char.heading))
        Call Manager.ChangeValue("INIT", "Head", CInt(.OrigChar.Head))
    
        If .flags.Muerto = 0 Then
            If .Char.body <> 0 Then
                Call Manager.ChangeValue("INIT", "Body", CInt(.Char.body))

            End If

        End If
    
        Call Manager.ChangeValue("INIT", "Arma", CInt(.Char.WeaponAnim))
        Call Manager.ChangeValue("INIT", "Escudo", CInt(.Char.ShieldAnim))
        Call Manager.ChangeValue("INIT", "Casco", CInt(.Char.CascoAnim))
    
        #If ConUpTime Then
    
            If SaveTimeOnline Then

                Dim TempDate As Date

                TempDate = Now - .LogOnTime
                .LogOnTime = Now
                .UpTime = .UpTime + (Abs(Day(TempDate) - 30) * 24 * 3600) + Hour(TempDate) * 3600 + Minute(TempDate) * 60 + Second(TempDate)
                Call Manager.ChangeValue("INIT", "UpTime", CLng(.UpTime))

            End If

        #End If
    
        'First time around?
        If Manager.GetValue("INIT", "LastIP1") = vbNullString Then
            Call Manager.ChangeValue("INIT", "LastIP1", .ip & " - " & Date & ":" & time)
            'Is it a different ip from last time?
        ElseIf .ip <> Left$(Manager.GetValue("INIT", "LastIP1"), InStr(1, Manager.GetValue("INIT", "LastIP1"), " ") - 1) Then

            Dim i As Integer

            For i = 5 To 2 Step -1
                Call Manager.ChangeValue("INIT", "LastIP" & i, Manager.GetValue("INIT", "LastIP" & CStr(i - 1)))
            Next i

            Call Manager.ChangeValue("INIT", "LastIP1", .ip & " - " & Date & ":" & time)
            'Same ip, just update the date
        Else
            Call Manager.ChangeValue("INIT", "LastIP1", .ip & " - " & Date & ":" & time)

        End If
    
        Call Manager.ChangeValue("INIT", "Position", .Pos.map & "-" & .Pos.x & "-" & .Pos.y)
    
        Call Manager.ChangeValue("STATS", "GLD", CLng(.Stats.Gld))
        Call Manager.ChangeValue("STATS", "BANCO", CLng(.Stats.Banco))
    
        Call Manager.ChangeValue("STATS", "MaxHP", CInt(.Stats.MaxHp))
        Call Manager.ChangeValue("STATS", "MinHP", CInt(.Stats.MinHp))
    
        Call Manager.ChangeValue("STATS", "MaxSTA", CInt(.Stats.MaxSta))
        Call Manager.ChangeValue("STATS", "MinSTA", CInt(.Stats.MinSta))
    
        Call Manager.ChangeValue("STATS", "MaxMAN", CInt(.Stats.MaxMAN))
        Call Manager.ChangeValue("STATS", "MinMAN", CInt(.Stats.MinMAN))
    
        Call Manager.ChangeValue("STATS", "MaxHIT", CInt(.Stats.MaxHIT))
        Call Manager.ChangeValue("STATS", "MinHIT", CInt(.Stats.MinHIT))
    
        Call Manager.ChangeValue("STATS", "MaxAGU", CByte(.Stats.MaxAGU))
        Call Manager.ChangeValue("STATS", "MinAGU", CByte(.Stats.MinAGU))
    
        Call Manager.ChangeValue("STATS", "MaxHAM", CByte(.Stats.MaxHam))
        Call Manager.ChangeValue("STATS", "MinHAM", CByte(.Stats.MinHam))
    
        Call Manager.ChangeValue("STATS", "SkillPtsLibres", CInt(.Stats.SkillPts))
    
        Call Manager.ChangeValue("STATS", "EXP", CDbl(.Stats.Exp))
        Call Manager.ChangeValue("STATS", "ELV", CByte(.Stats.ELV))
      
        Call Manager.ChangeValue("STATS", "ELU", CLng(.Stats.ELU))
    
        Call Manager.ChangeValue("STATS", "InventLevel", CByte(.Stats.InventLevel))
    
        Call Manager.ChangeValue("MUERTES", "UserMuertes", CLng(.Stats.UsuariosMatados))
        Call Manager.ChangeValue("MUERTES", "NpcsMuertes", CInt(.Stats.NPCsMuertos))
      
        Call Manager.ChangeValue("BancoInventory", "CantidadItems", CInt(.BancoInvent.NroItems))

        For LoopC = 1 To MAX_BANCOINVENTORY_SLOTS
            Call Manager.ChangeValue("BancoInventory", "Obj" & LoopC, .BancoInvent.Object(LoopC).ObjIndex & "-" & .BancoInvent.Object(LoopC).amount)
        Next LoopC
      
        'Save Inv
        Call Manager.ChangeValue("Inventory", "CantidadItems", CInt(.Invent.NroItems))
    
        For LoopC = 1 To MAX_INVENTORY_SLOTS
            Call Manager.ChangeValue("Inventory", "Obj" & LoopC, .Invent.Object(LoopC).ObjIndex & "-" & .Invent.Object(LoopC).amount & "-" & .Invent.Object(LoopC).Equipped)
        Next LoopC
    
        Call Manager.ChangeValue("Inventory", "WeaponEqpSlot", CByte(.Invent.WeaponEqpSlot))
        Call Manager.ChangeValue("Inventory", "ArmourEqpSlot", CByte(.Invent.ArmourEqpSlot))
        Call Manager.ChangeValue("Inventory", "CascoEqpSlot", CByte(.Invent.CascoEqpSlot))
        Call Manager.ChangeValue("Inventory", "EscudoEqpSlot", CByte(.Invent.EscudoEqpSlot))
        Call Manager.ChangeValue("Inventory", "BarcoSlot", CByte(.Invent.BarcoSlot))
        Call Manager.ChangeValue("Inventory", "MonturaEqpSlot", CByte(.Invent.MonturaEqpSlot))
        Call Manager.ChangeValue("Inventory", "MunicionSlot", CByte(.Invent.MunicionEqpSlot))
        Call Manager.ChangeValue("Inventory", "AnilloSlot", CByte(.Invent.AnilloEqpSlot))
    
        Dim cad As String
    
        For LoopC = 1 To MAXUSERHECHIZOS
            cad = .Stats.UserHechizos(LoopC)
            Call Manager.ChangeValue("HECHIZOS", "H" & LoopC, cad)
        Next
    
        Dim NroMascotas As Long

        NroMascotas = .NroMascotas
    
        For LoopC = 1 To MAXMASCOTAS

            ' Mascota valida?
            If .MascotasIndex(LoopC) > 0 Then

                ' Nos aseguramos que la criatura no fue invocada
                If Npclist(.MascotasIndex(LoopC)).Contadores.TiempoExistencia = 0 Then
                    cad = .MascotasType(LoopC)
                Else 'Si fue invocada no la guardamos
                    cad = "0"
                    NroMascotas = NroMascotas - 1

                End If

                Call Manager.ChangeValue("MASCOTAS", "MAS" & LoopC, cad)
            Else
                cad = .MascotasType(LoopC)
                Call Manager.ChangeValue("MASCOTAS", "MAS" & LoopC, cad)

            End If
    
        Next
    
        Call Manager.ChangeValue("MASCOTAS", "NroMascotas", CInt(NroMascotas))
    
        'Devuelve el head de muerto
        If .flags.Muerto = 1 Then
            .Char.Head = iCabezaMuerto

        End If

    End With

    Call SaveQuestStats(Userindex, Manager)

    Call Manager.DumpFile(UserFile)

    Set Manager = Nothing

    If Existe Then Call Kill(UserFile & ".bk")

    Exit Sub

ErrorHandler:
    Call LogError("Error en SaveUserToCharfile: " & UserFile & " -- " & Err.Number & ": " & Err.description)

    Set Manager = Nothing

End Sub

Sub BackUPnPc(ByVal NpcIndex As Integer, ByVal hFile As Integer)
    '***************************************************
    'Author: Unknown
    'Last Modification: 10/09/2010
    '10/09/2010 - Pato: Optimice el BackUp de NPCs
    '***************************************************

    Dim LoopC As Integer
    
    Print #hFile, "[NPC" & Npclist(NpcIndex).Numero & "]"
    
    With Npclist(NpcIndex)
        'General
        Print #hFile, "Name=" & .Name
        Print #hFile, "Desc=" & .Desc
        Print #hFile, "Head=" & val(.Char.Head)
        Print #hFile, "Body=" & val(.Char.body)
        Print #hFile, "Heading=" & val(.Char.heading)
        Print #hFile, "Movement=" & val(.Movement)
        Print #hFile, "Attackable=" & val(.Attackable)
        Print #hFile, "Comercia=" & val(.Comercia)
        Print #hFile, "TipoItems=" & val(.TipoItems)
        Print #hFile, "Hostil=" & val(.Hostile)
        Print #hFile, "GiveEXP=" & val(.GiveEXP)
        Print #hFile, "GiveGLD=" & val(.GiveGLD)
        Print #hFile, "InvReSpawn=" & val(.InvReSpawn)
        Print #hFile, "NpcType=" & val(.NPCtype)
        
        'Stats
        Print #hFile, "Alineacion=" & val(.Stats.Alineacion)
        Print #hFile, "DEF=" & val(.Stats.def)
        Print #hFile, "MaxHit=" & val(.Stats.MaxHIT)
        Print #hFile, "MaxHp=" & val(.Stats.MaxHp)
        Print #hFile, "MinHit=" & val(.Stats.MinHIT)
        Print #hFile, "MinHp=" & val(.Stats.MinHp)
        
        'Flags
        Print #hFile, "ReSpawn=" & val(.flags.Respawn)
        Print #hFile, "BackUp=" & val(.flags.BackUp)
        Print #hFile, "Domable=" & val(.flags.Domable)
        
        'Inventario
        Print #hFile, "NroItems=" & val(.Invent.NroItems)

        If .Invent.NroItems > 0 Then

            For LoopC = 1 To .Invent.NroItems
                Print #hFile, "Obj" & LoopC & "=" & .Invent.Object(LoopC).ObjIndex & "-" & .Invent.Object(LoopC).amount
            Next LoopC

        End If
        
        Print #hFile, ""

    End With

End Sub

Sub CargarNpcBackUp(ByVal NpcIndex As Integer, ByVal NpcNumber As Integer)
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    'Status
    If frmMain.Visible Then frmMain.txtStatus.Text = "Cargando backup Npc"
    
    Dim npcfile As String
    
    'If NpcNumber > 499 Then
    '    npcfile = DatPath & "bkNPCs-HOSTILES.dat"
    'Else
    npcfile = DatPath & "bkNPCs.dat"
    'End If
    
    With Npclist(NpcIndex)
    
        .Numero = NpcNumber
        .Name = GetVar(npcfile, "NPC" & NpcNumber, "Name")
        .Desc = GetVar(npcfile, "NPC" & NpcNumber, "Desc")
        .Movement = val(GetVar(npcfile, "NPC" & NpcNumber, "Movement"))
        .NPCtype = val(GetVar(npcfile, "NPC" & NpcNumber, "NpcType"))
        
        .Char.body = val(GetVar(npcfile, "NPC" & NpcNumber, "Body"))
        .Char.Head = val(GetVar(npcfile, "NPC" & NpcNumber, "Head"))
        .Char.heading = val(GetVar(npcfile, "NPC" & NpcNumber, "Heading"))
        
        .Attackable = val(GetVar(npcfile, "NPC" & NpcNumber, "Attackable"))
        .Comercia = val(GetVar(npcfile, "NPC" & NpcNumber, "Comercia"))
        .Hostile = val(GetVar(npcfile, "NPC" & NpcNumber, "Hostile"))
        .GiveEXP = val(GetVar(npcfile, "NPC" & NpcNumber, "GiveEXP"))
        
        .GiveGLD = val(GetVar(npcfile, "NPC" & NpcNumber, "GiveGLD"))
        
        .InvReSpawn = val(GetVar(npcfile, "NPC" & NpcNumber, "InvReSpawn"))
        
        .Stats.MaxHp = val(GetVar(npcfile, "NPC" & NpcNumber, "MaxHP"))
        .Stats.MinHp = val(GetVar(npcfile, "NPC" & NpcNumber, "MinHP"))
        .Stats.MaxHIT = val(GetVar(npcfile, "NPC" & NpcNumber, "MaxHIT"))
        .Stats.MinHIT = val(GetVar(npcfile, "NPC" & NpcNumber, "MinHIT"))
        .Stats.def = val(GetVar(npcfile, "NPC" & NpcNumber, "DEF"))
        .Stats.Alineacion = val(GetVar(npcfile, "NPC" & NpcNumber, "Alineacion"))
        
        Dim LoopC As Integer

        Dim ln    As String

        .Invent.NroItems = val(GetVar(npcfile, "NPC" & NpcNumber, "NROITEMS"))

        If .Invent.NroItems > 0 Then

            For LoopC = 1 To MAX_INVENTORY_SLOTS
                ln = GetVar(npcfile, "NPC" & NpcNumber, "Obj" & LoopC)
                .Invent.Object(LoopC).ObjIndex = val(ReadField(1, ln, 45))
                .Invent.Object(LoopC).amount = val(ReadField(2, ln, 45))
               
            Next LoopC

        Else

            For LoopC = 1 To MAX_INVENTORY_SLOTS
                .Invent.Object(LoopC).ObjIndex = 0
                .Invent.Object(LoopC).amount = 0
            Next LoopC

        End If
        
        For LoopC = 1 To MAX_NPC_DROPS
            ln = GetVar(npcfile, "NPC" & NpcNumber, "Drop" & LoopC)
            .Drop(LoopC).ObjIndex = val(ReadField(1, ln, 45))
            .Drop(LoopC).amount = val(ReadField(2, ln, 45))
        Next LoopC
        
        .flags.NPCActive = True
        .flags.Respawn = val(GetVar(npcfile, "NPC" & NpcNumber, "ReSpawn"))
        .flags.BackUp = val(GetVar(npcfile, "NPC" & NpcNumber, "BackUp"))
        .flags.Domable = val(GetVar(npcfile, "NPC" & NpcNumber, "Domable"))
        .flags.RespawnOrigPos = val(GetVar(npcfile, "NPC" & NpcNumber, "OrigPos"))
        
        'Tipo de items con los que comercia
        .TipoItems = val(GetVar(npcfile, "NPC" & NpcNumber, "TipoItems"))

    End With

    If frmMain.Visible Then frmMain.txtStatus.Text = Date & " " & time & " - Se cargo el archivo bkNPCs.dat"

End Sub

Sub Ban(ByVal BannedName As String, ByVal Baneador As String, ByVal Motivo As String)
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    Call WriteVar(App.path & "\logs\" & "BanDetail.dat", BannedName, "BannedBy", Baneador)
    Call WriteVar(App.path & "\logs\" & "BanDetail.dat", BannedName, "Reason", Motivo)
    
    'Log interno del servidor, lo usa para hacer un UNBAN general de toda la gente banned
    Dim mifile As Integer

    mifile = FreeFile
    Open App.path & "\logs\GenteBanned.log" For Append Shared As #mifile
    Print #mifile, BannedName
    Close #mifile

End Sub

Public Sub CargaApuestas()
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************
    If frmMain.Visible Then frmMain.txtStatus.Text = "Cargando apuestas.dat"

    Apuestas.Ganancias = val(GetVar(DatPath & "apuestas.dat", "Main", "Ganancias"))
    Apuestas.Perdidas = val(GetVar(DatPath & "apuestas.dat", "Main", "Perdidas"))
    Apuestas.Jugadas = val(GetVar(DatPath & "apuestas.dat", "Main", "Jugadas"))

    If frmMain.Visible Then frmMain.txtStatus.Text = Date & " " & time & " - Se cargo el archivo apuestas.dat"

End Sub

Public Sub generateMatrix(ByVal Mapa As Integer)
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    Dim i As Integer

    Dim j As Integer
    
    ReDim distanceToCities(1 To NumMaps) As HomeDistance
    
    For j = 1 To NUMCIUDADES
        For i = 1 To NumMaps
            distanceToCities(i).distanceToCity(j) = -1
        Next i
    Next j
    
    For j = 1 To NUMCIUDADES
        For i = 1 To 4

            Select Case i

                Case eHeading.NORTH
                    Call setDistance(getLimit(Ciudades(j).map, eHeading.NORTH), j, i, 0, 1)

                Case eHeading.EAST
                    Call setDistance(getLimit(Ciudades(j).map, eHeading.EAST), j, i, 1, 0)

                Case eHeading.SOUTH
                    Call setDistance(getLimit(Ciudades(j).map, eHeading.SOUTH), j, i, 0, 1)

                Case eHeading.WEST
                    Call setDistance(getLimit(Ciudades(j).map, eHeading.WEST), j, i, -1, 0)

            End Select

        Next i
    Next j

End Sub

Public Sub setDistance(ByVal Mapa As Integer, _
                       ByVal city As Byte, _
                       ByVal side As Integer, _
                       Optional ByVal x As Integer = 0, _
                       Optional ByVal y As Integer = 0)
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    Dim i   As Integer

    Dim lim As Integer

    If Mapa <= 0 Or Mapa > NumMaps Then Exit Sub

    If distanceToCities(Mapa).distanceToCity(city) >= 0 Then Exit Sub

    If Mapa = Ciudades(city).map Then
        distanceToCities(Mapa).distanceToCity(city) = 0
    Else
        distanceToCities(Mapa).distanceToCity(city) = Abs(x) + Abs(y)

    End If

    For i = 1 To 4
        lim = getLimit(Mapa, i)

        If lim > 0 Then

            Select Case i

                Case eHeading.NORTH
                    Call setDistance(lim, city, i, x, y + 1)

                Case eHeading.EAST
                    Call setDistance(lim, city, i, x + 1, y)

                Case eHeading.SOUTH
                    Call setDistance(lim, city, i, x, y - 1)

                Case eHeading.WEST
                    Call setDistance(lim, city, i, x - 1, y)

            End Select

        End If

    Next i

End Sub

Public Function getLimit(ByVal Mapa As Integer, ByVal side As Byte) As Integer

    '***************************************************
    'Author: Budi
    'Last Modification: 31/01/2010
    'Retrieves the limit in the given side in the given map.
    'TODO: This should be set in the .inf map file.
    '***************************************************
    Dim x As Long

    Dim y As Long

    If Mapa <= 0 Then Exit Function

    For x = 15 To 87
        For y = 0 To 3

            Select Case side

                Case eHeading.NORTH
                    getLimit = MapData(Mapa, x, 7 + y).TileExit.map

                Case eHeading.EAST
                    getLimit = MapData(Mapa, 92 - y, x).TileExit.map

                Case eHeading.SOUTH
                    getLimit = MapData(Mapa, x, 94 - y).TileExit.map

                Case eHeading.WEST
                    getLimit = MapData(Mapa, 9 + y, x).TileExit.map

            End Select

            If getLimit > 0 Then Exit Function
        Next y
    Next x

End Function

Sub SendUserBovedaTxtFromCharfile(ByVal sendIndex As Integer, ByVal charName As String)
    '***************************************************
    'Author: Unknown
    'Last Modification: 19/09/2018
    'CHOTS: Lo movi a esta funcion porque tiene mas sentido
    '***************************************************

    On Error Resume Next

    Dim j        As Integer
    Dim CharFile As String, Tmp As String
    Dim ObjInd   As Long, ObjCant As Long

    CharFile = CharPath & charName & ".chr"

    If FileExist(CharFile, vbNormal) Then
        Call WriteConsoleMsg(sendIndex, charName, FontTypeNames.FONTTYPE_INFO)
        Call WriteConsoleMsg(sendIndex, "Tiene " & GetVar(CharFile, "BancoInventory", "CantidadItems") & " objetos.", FontTypeNames.FONTTYPE_INFO)

        For j = 1 To MAX_BANCOINVENTORY_SLOTS
            Tmp = GetVar(CharFile, "BancoInventory", "Obj" & j)
            ObjInd = ReadField(1, Tmp, Asc("-"))
            ObjCant = ReadField(2, Tmp, Asc("-"))

            If ObjInd > 0 Then
                Call WriteConsoleMsg(sendIndex, "Objeto " & j & " " & ObjData(ObjInd).Name & " Cantidad:" & ObjCant, FontTypeNames.FONTTYPE_INFO)

            End If

        Next
    Else
        Call WriteConsoleMsg(sendIndex, "Usuario inexistente: " & charName, FontTypeNames.FONTTYPE_INFO)

    End If

End Sub

Sub SendUserMiniStatsTxtFromCharfile(ByVal sendIndex As Integer, ByVal charName As String)

    '*************************************************
    'Author: Unknown
    'Last modified: 19/19/2018
    'Shows the users Stats when the user is offline.
    '23/01/2007 Pablo (ToxicWaste) - Agrego de funciones y mejora de distribucion de parametros.
    '19/09/2018 CHOTS - Movido a FileIO
    '*************************************************
    Dim CharFile      As String

    Dim Ban           As String

    Dim BanDetailPath As String
    
    BanDetailPath = App.path & "\logs\" & "BanDetail.dat"
    CharFile = CharPath & charName & ".chr"
    
    If FileExist(CharFile) Then
        Call WriteConsoleMsg(sendIndex, "Pj: " & charName, FontTypeNames.FONTTYPE_INFO)
        Call WriteConsoleMsg(sendIndex, "Ciudadanos matados: " & GetVar(CharFile, "FACCIONES", "CiudMatados") & " CriminalesMatados: " & GetVar(CharFile, "FACCIONES", "CrimMatados") & " usuarios matados: " & GetVar(CharFile, "MUERTES", "UserMuertes"), FontTypeNames.FONTTYPE_INFO)
        Call WriteConsoleMsg(sendIndex, "NPCs muertos: " & GetVar(CharFile, "MUERTES", "NpcsMuertes"), FontTypeNames.FONTTYPE_INFO)
        Call WriteConsoleMsg(sendIndex, "Clase: " & ListaClases(GetVar(CharFile, "INIT", "Clase")), FontTypeNames.FONTTYPE_INFO)
        Call WriteConsoleMsg(sendIndex, "Pena: " & GetVar(CharFile, "COUNTERS", "PENA"), FontTypeNames.FONTTYPE_INFO)
        
        If CByte(GetVar(CharFile, "FACCIONES", "EjercitoReal")) = 1 Then
            Call WriteConsoleMsg(sendIndex, "Ejercito real desde: " & GetVar(CharFile, "FACCIONES", "FechaIngreso"), FontTypeNames.FONTTYPE_INFO)
            Call WriteConsoleMsg(sendIndex, "Ingreso en nivel: " & CInt(GetVar(CharFile, "FACCIONES", "NivelIngreso")) & " con " & CInt(GetVar(CharFile, "FACCIONES", "MatadosIngreso")) & " ciudadanos matados.", FontTypeNames.FONTTYPE_INFO)
            Call WriteConsoleMsg(sendIndex, "Veces que ingreso: " & CByte(GetVar(CharFile, "FACCIONES", "Reenlistadas")), FontTypeNames.FONTTYPE_INFO)
        
        ElseIf CByte(GetVar(CharFile, "FACCIONES", "EjercitoCaos")) = 1 Then
            Call WriteConsoleMsg(sendIndex, "Legion oscura desde: " & GetVar(CharFile, "FACCIONES", "FechaIngreso"), FontTypeNames.FONTTYPE_INFO)
            Call WriteConsoleMsg(sendIndex, "Ingreso en nivel: " & CInt(GetVar(CharFile, "FACCIONES", "NivelIngreso")), FontTypeNames.FONTTYPE_INFO)
            Call WriteConsoleMsg(sendIndex, "Veces que ingreso: " & CByte(GetVar(CharFile, "FACCIONES", "Reenlistadas")), FontTypeNames.FONTTYPE_INFO)
        
        ElseIf CByte(GetVar(CharFile, "FACCIONES", "rExReal")) = 1 Then
            Call WriteConsoleMsg(sendIndex, "Fue ejercito real", FontTypeNames.FONTTYPE_INFO)
            Call WriteConsoleMsg(sendIndex, "Veces que ingreso: " & CByte(GetVar(CharFile, "FACCIONES", "Reenlistadas")), FontTypeNames.FONTTYPE_INFO)
        
        ElseIf CByte(GetVar(CharFile, "FACCIONES", "rExCaos")) = 1 Then
            Call WriteConsoleMsg(sendIndex, "Fue legion oscura", FontTypeNames.FONTTYPE_INFO)
            Call WriteConsoleMsg(sendIndex, "Veces que ingreso: " & CByte(GetVar(CharFile, "FACCIONES", "Reenlistadas")), FontTypeNames.FONTTYPE_INFO)

        End If
        
        Call WriteConsoleMsg(sendIndex, "Asesino: " & CLng(GetVar(CharFile, "REP", "Asesino")), FontTypeNames.FONTTYPE_INFO)
        Call WriteConsoleMsg(sendIndex, "Noble: " & CLng(GetVar(CharFile, "REP", "Nobles")), FontTypeNames.FONTTYPE_INFO)
        
        If IsNumeric(GetVar(CharFile, "Guild", "GUILDINDEX")) Then
            Call WriteConsoleMsg(sendIndex, "Clan: " & modGuilds.GuildName(CInt(GetVar(CharFile, "Guild", "GUILDINDEX"))), FontTypeNames.FONTTYPE_INFO)

        End If
        
        Ban = GetVar(CharFile, "FLAGS", "Ban")
        Call WriteConsoleMsg(sendIndex, "Ban: " & Ban, FontTypeNames.FONTTYPE_INFO)
        
        If Ban = "1" Then
            Call WriteConsoleMsg(sendIndex, "Ban por: " & GetVar(CharFile, charName, "BannedBy") & " Motivo: " & GetVar(BanDetailPath, charName, "Reason"), FontTypeNames.FONTTYPE_INFO)

        End If

    Else
        Call WriteConsoleMsg(sendIndex, "El pj no existe: " & charName, FontTypeNames.FONTTYPE_INFO)

    End If

End Sub

Sub SendUserInvTxtFromCharfile(ByVal sendIndex As Integer, ByVal charName As String)
    '***************************************************
    'Author: Unknown
    'Last Modification: 19/09/2018
    '19/09/2018 CHOTS - Movido a FileIO
    '***************************************************

    On Error Resume Next

    Dim j        As Long

    Dim CharFile As String, Tmp As String

    Dim ObjInd   As Long, ObjCant As Long
    
    CharFile = CharPath & charName & ".chr"
    
    If FileExist(CharFile, vbNormal) Then
        Call WriteConsoleMsg(sendIndex, charName, FontTypeNames.FONTTYPE_INFO)
        Call WriteConsoleMsg(sendIndex, "Tiene " & GetVar(CharFile, "Inventory", "CantidadItems") & " objetos.", FontTypeNames.FONTTYPE_INFO)
        
        For j = 1 To MAX_INVENTORY_SLOTS
            Tmp = GetVar(CharFile, "Inventory", "Obj" & j)
            ObjInd = ReadField(1, Tmp, Asc("-"))
            ObjCant = ReadField(2, Tmp, Asc("-"))

            If ObjInd > 0 Then
                Call WriteConsoleMsg(sendIndex, "Objeto " & j & " " & ObjData(ObjInd).Name & " Cantidad:" & ObjCant, FontTypeNames.FONTTYPE_INFO)

            End If

        Next j

    Else
        Call WriteConsoleMsg(sendIndex, "Usuario inexistente: " & charName, FontTypeNames.FONTTYPE_INFO)

    End If

End Sub

Sub SendUserOROTxtFromCharfile(ByVal sendIndex As Integer, ByVal charName As String)
    '***************************************************
    'Author: Unknown
    'Last Modification: 19/09/2018
    '19/09/2018 CHOTS - Movido a FileIO
    '***************************************************

    Dim CharFile As String
    
    On Error Resume Next

    CharFile = CharPath & charName & ".chr"
    
    If FileExist(CharFile, vbNormal) Then
        Call WriteConsoleMsg(sendIndex, charName, FontTypeNames.FONTTYPE_INFO)
        Call WriteConsoleMsg(sendIndex, "Tiene " & GetVar(CharFile, "STATS", "BANCO") & " en el banco.", FontTypeNames.FONTTYPE_INFO)
    Else
        Call WriteConsoleMsg(sendIndex, "Usuario inexistente: " & charName, FontTypeNames.FONTTYPE_INFO)

    End If

End Sub

Sub SendUserStatsTxtCharfile(ByVal sendIndex As Integer, ByVal Nombre As String)
    '***************************************************
    'Author: Unknown
    'Last Modification: 19/09/2018
    '19/09/2018 CHOTS - Movido a FileIO
    '***************************************************

    If PersonajeExiste(Nombre) Then
        Call WriteConsoleMsg(sendIndex, "Pj Inexistente", FontTypeNames.FONTTYPE_INFO)
    Else
        Call WriteConsoleMsg(sendIndex, "Estadisticas de: " & Nombre, FontTypeNames.FONTTYPE_INFO)
        Call WriteConsoleMsg(sendIndex, "Nivel: " & GetVar(CharPath & Nombre & ".chr", "stats", "elv") & "  EXP: " & GetVar(CharPath & Nombre & ".chr", "stats", "Exp") & "/" & GetVar(CharPath & Nombre & ".chr", "stats", "elu"), FontTypeNames.FONTTYPE_INFO)
        Call WriteConsoleMsg(sendIndex, "Energia: " & GetVar(CharPath & Nombre & ".chr", "stats", "minsta") & "/" & GetVar(CharPath & Nombre & ".chr", "stats", "maxSta"), FontTypeNames.FONTTYPE_INFO)
        Call WriteConsoleMsg(sendIndex, "Salud: " & GetVar(CharPath & Nombre & ".chr", "stats", "MinHP") & "/" & GetVar(CharPath & Nombre & ".chr", "Stats", "MaxHP") & "  Mana: " & GetVar(CharPath & Nombre & ".chr", "Stats", "MinMAN") & "/" & GetVar(CharPath & Nombre & ".chr", "Stats", "MaxMAN"), FontTypeNames.FONTTYPE_INFO)
        
        Call WriteConsoleMsg(sendIndex, "Menor Golpe/Mayor Golpe: " & GetVar(CharPath & Nombre & ".chr", "stats", "MaxHIT"), FontTypeNames.FONTTYPE_INFO)
        
        Call WriteConsoleMsg(sendIndex, "Oro: " & GetVar(CharPath & Nombre & ".chr", "stats", "GLD"), FontTypeNames.FONTTYPE_INFO)
        
        #If ConUpTime Then

            Dim TempSecs As Long

            Dim TempStr  As String

            TempSecs = GetVar(CharPath & Nombre & ".chr", "INIT", "UpTime")
            TempStr = (TempSecs \ 86400) & " Dias, " & ((TempSecs Mod 86400) \ 3600) & " Horas, " & ((TempSecs Mod 86400) Mod 3600) \ 60 & " Minutos, " & (((TempSecs Mod 86400) Mod 3600) Mod 60) & " Segundos."
            Call WriteConsoleMsg(sendIndex, "Tiempo Logeado: " & TempStr, FontTypeNames.FONTTYPE_INFO)
        #End If
    
        Call WriteConsoleMsg(sendIndex, "Dados: " & GetVar(CharPath & Nombre & ".chr", "ATRIBUTOS", "AT1") & ", " & GetVar(CharPath & Nombre & ".chr", "ATRIBUTOS", "AT2") & ", " & GetVar(CharPath & Nombre & ".chr", "ATRIBUTOS", "AT3") & ", " & GetVar(CharPath & Nombre & ".chr", "ATRIBUTOS", "AT4") & ", " & GetVar(CharPath & Nombre & ".chr", "ATRIBUTOS", "AT5"), FontTypeNames.FONTTYPE_INFO)

    End If

End Sub

Public Sub ReloadNPCByIndex(ByVal NpcIndex As Integer)

    On Error GoTo errHandler

    Dim NpcNumber As Integer
    Dim LoopC As Long
    Dim ln As String

    With Npclist(NpcIndex)
        NpcNumber = .Numero
        .Name = LeerNPCs.GetValue("NPC" & NpcNumber, "Name")
        .Desc = LeerNPCs.GetValue("NPC" & NpcNumber, "Desc")
        .level = LeerNPCs.GetValue("NPC" & NpcNumber, "Level")
        
        If .level = 0 Then .level = 30
        
        .Movement = val(LeerNPCs.GetValue("NPC" & NpcNumber, "Movement"))
        .flags.OldMovement = .Movement

        .flags.AguaValida = val(LeerNPCs.GetValue("NPC" & NpcNumber, "AguaValida"))
        .flags.TierraInvalida = val(LeerNPCs.GetValue("NPC" & NpcNumber, "TierraInValida"))
        .flags.Faccion = val(LeerNPCs.GetValue("NPC" & NpcNumber, "Faccion"))
        .flags.AtacaDoble = val(LeerNPCs.GetValue("NPC" & NpcNumber, "AtacaDoble"))

        .NPCtype = val(LeerNPCs.GetValue("NPC" & NpcNumber, "NpcType"))

        .Char.body = val(LeerNPCs.GetValue("NPC" & NpcNumber, "Body"))
        .Char.Head = val(LeerNPCs.GetValue("NPC" & NpcNumber, "Head"))

        .Attackable = val(LeerNPCs.GetValue("NPC" & NpcNumber, "Attackable"))
        .Comercia = val(LeerNPCs.GetValue("NPC" & NpcNumber, "Comercia"))
        .Hostile = val(LeerNPCs.GetValue("NPC" & NpcNumber, "Hostile"))

        .GiveEXP = val(LeerNPCs.GetValue("NPC" & NpcNumber, "GiveEXP")) * ExpMultiplier

        If HappyHourActivated And (HappyHour <> 0) Then
            .GiveEXP = .GiveEXP * HappyHour
        End If

        .flags.ExpCount = .GiveEXP

        .Veneno = val(LeerNPCs.GetValue("NPC" & NpcNumber, "Veneno"))

        .flags.Domable = val(LeerNPCs.GetValue("NPC" & NpcNumber, "Domable"))

        .GiveGLD = val(LeerNPCs.GetValue("NPC" & NpcNumber, "GiveGLD"))

        .QuestNumber = val(LeerNPCs.GetValue("NPC" & NpcNumber, "QuestNumber"))

        .PoderAtaque = val(LeerNPCs.GetValue("NPC" & NpcNumber, "PoderAtaque"))
        .PoderEvasion = val(LeerNPCs.GetValue("NPC" & NpcNumber, "PoderEvasion"))

        .InvReSpawn = val(LeerNPCs.GetValue("NPC" & NpcNumber, "InvReSpawn"))

        With .Stats
            .MaxHp = val(LeerNPCs.GetValue("NPC" & NpcNumber, "MaxHP"))
            '.MinHp = val(LeerNPCs.GetValue("NPC" & npcNumber, "MinHP"))
            .MaxHIT = val(LeerNPCs.GetValue("NPC" & NpcNumber, "MaxHIT"))
            .MinHIT = val(LeerNPCs.GetValue("NPC" & NpcNumber, "MinHIT"))
            .def = val(LeerNPCs.GetValue("NPC" & NpcNumber, "DEF"))
            .defM = val(LeerNPCs.GetValue("NPC" & NpcNumber, "DEFm"))
            .Alineacion = val(LeerNPCs.GetValue("NPC" & NpcNumber, "Alineacion"))

        End With

        .Invent.NroItems = val(LeerNPCs.GetValue("NPC" & NpcNumber, "NROITEMS"))

        For LoopC = 1 To .Invent.NroItems
            ln = LeerNPCs.GetValue("NPC" & NpcNumber, "Obj" & LoopC)
            .Invent.Object(LoopC).ObjIndex = val(ReadField(1, ln, 45))
            .Invent.Object(LoopC).amount = val(ReadField(2, ln, 45))
        Next LoopC

        For LoopC = 1 To MAX_NPC_DROPS
            ln = LeerNPCs.GetValue("NPC" & NpcNumber, "Drop" & LoopC)
            .Drop(LoopC).ObjIndex = val(ReadField(1, ln, 45))

            If .Drop(LoopC).ObjIndex = iORO Then
                .Drop(LoopC).amount = val(ReadField(2, ln, 45)) * OroMultiplier
            Else
                .Drop(LoopC).amount = val(ReadField(2, ln, 45))

            End If

        Next LoopC

        .flags.LanzaSpells = val(LeerNPCs.GetValue("NPC" & NpcNumber, "LanzaSpells"))

        If .flags.LanzaSpells > 0 Then ReDim .Spells(1 To .flags.LanzaSpells)

        For LoopC = 1 To .flags.LanzaSpells
            .Spells(LoopC) = val(LeerNPCs.GetValue("NPC" & NpcNumber, "Sp" & LoopC))
        Next LoopC

        If .NPCtype = eNPCType.Entrenador Then
            .NroCriaturas = val(LeerNPCs.GetValue("NPC" & NpcNumber, "NroCriaturas"))
            ReDim .Criaturas(1 To .NroCriaturas) As tCriaturasEntrenador

            For LoopC = 1 To .NroCriaturas
                .Criaturas(LoopC).NpcIndex = LeerNPCs.GetValue("NPC" & NpcNumber, "CI" & LoopC)
                .Criaturas(LoopC).NpcName = LeerNPCs.GetValue("NPC" & NpcNumber, "CN" & LoopC)
            Next LoopC

        End If

        With .flags

            'If Respawn Then
            '    .Respawn = val(LeerNPCs.GetValue("NPC" & npcNumber, "ReSpawn"))
            'Else
            '    .Respawn = 1
            'End If

            .BackUp = val(LeerNPCs.GetValue("NPC" & NpcNumber, "BackUp"))
            .RespawnOrigPos = val(LeerNPCs.GetValue("NPC" & NpcNumber, "OrigPos"))
            .AfectaParalisis = val(LeerNPCs.GetValue("NPC" & NpcNumber, "AfectaParalisis"))

            .Snd1 = val(LeerNPCs.GetValue("NPC" & NpcNumber, "Snd1"))
            .Snd2 = val(LeerNPCs.GetValue("NPC" & NpcNumber, "Snd2"))
            .Snd3 = val(LeerNPCs.GetValue("NPC" & NpcNumber, "Snd3"))

        End With

        '<<<<<<<<<<<<<< Expresiones >>>>>>>>>>>>>>>>
        .NroExpresiones = val(LeerNPCs.GetValue("NPC" & NpcNumber, "NROEXP"))

        If .NroExpresiones > 0 Then
            ReDim .Expresiones(1 To .NroExpresiones) As String
        End If

        For LoopC = 1 To .NroExpresiones
            .Expresiones(LoopC) = LeerNPCs.GetValue("NPC" & NpcNumber, "Exp" & LoopC)
        Next LoopC
        '<<<<<<<<<<<<<< Expresiones >>>>>>>>>>>>>>>>

        'Tipo de items con los que comercia
        .TipoItems = val(LeerNPCs.GetValue("NPC" & NpcNumber, "TipoItems"))

        .Ciudad = val(LeerNPCs.GetValue("NPC" & NpcNumber, "Ciudad"))

    End With

    Exit Sub

errHandler:
    Call LogError("Error en ReloadNPCIndexByFile - Err: " & Err.Number & " " & Err.description)

End Sub
