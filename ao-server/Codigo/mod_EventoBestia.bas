Attribute VB_Name = "mod_EventoBestia"
Option Explicit

Public npci_Bestia As Integer
Public Const NPCINDEX_Bestia As Integer = 409
Public Const Bestia_Map As Integer = 2
Private Const bestia_xentrance As Byte = 55
Private Const bestia_yentrance As Byte = 55
Public bestia_tp As WorldPos
Public bestia_contador As Byte
Public bestia_partys(1 To 2) As Integer


Public Function PartyIsDeath(ByVal Userindex As Integer) As Boolean

    '*************************************************
    'Author: Unknown
    'Last modified: 11/27/09 (Budi)
    'Adapte la funcion a los nuevos metodos de clsParty
    '*************************************************
    Dim i                                    As Integer

    Dim pi                                   As Integer

    Dim MembersOnline(1 To PARTY_MAXMEMBERS) As Integer
    
    Dim lError As String
    
    Dim tmpTeam As Byte

    pi = UserList(Userindex).PartyIndex
    
    tmpTeam = 1
    PartyIsDeath = True
    If pi > 0 Then
        Call Parties(pi).ObtenerMiembrosOnline(MembersOnline())
        
        For i = 1 To PARTY_MAXMEMBERS
            
            If MembersOnline(i) > 0 Then
            
                If UserList(MembersOnline(i)).flags.Muerto = 0 Then
                    PartyIsDeath = False
                    Exit Function
                End If
            End If
            
        Next i
    End If
    
End Function

Private Function team_isOnDungeon(ByVal pi As Integer) As Boolean
    Dim MembersOnline(1 To PARTY_MAXMEMBERS) As Integer
    Dim i As Long
    
    team_isOnDungeon = True
    
    For i = 1 To PARTY_MAXMEMBERS
        
        If MembersOnline(i) > 0 Then
            
            If UserList(MembersOnline(i)).Pos.map <> Bestia_Map Then
                team_isOnDungeon = False
            End If
            
        End If
        
    Next i
    
End Function

Public Sub Bestia_UserDie(ByVal ui As Integer)
    Dim MembersOnline(1 To PARTY_MAXMEMBERS) As Integer
    Dim i As Long
    If PartyIsDeath(ui) Then
        For i = 1 To PARTY_MAXMEMBERS
            
            If MembersOnline(i) > 0 Then
            
                Call WarpUserChar(MembersOnline(i), bestia_tp.map, bestia_tp.x, bestia_tp.y, True)
                
            End If
            
        Next i
    End If
End Sub

Public Sub Bestia_PasarMinuto()
    Dim i As Long, tempIndex As Integer
        If bestia_contador > 0 Then
        
            bestia_contador = bestia_contador - 1
            
            If bestia_contador = 30 Then
                MapData(bestia_tp.map, bestia_tp.x, bestia_tp.y).ObjInfo.ObjIndex = 0
                MapData(bestia_tp.map, bestia_tp.x, bestia_tp.y).ObjInfo.amount = 0
                MapData(bestia_tp.map, bestia_tp.x, bestia_tp.y).TileExit.map = 0
                MapData(bestia_tp.map, bestia_tp.x, bestia_tp.y).TileExit.x = 0
                MapData(bestia_tp.map, bestia_tp.x, bestia_tp.y).TileExit.y = 0
            End If
            
            If bestia_contador <= 5 Then
                Call SendData(SendTarget.toMap, Bestia_Map, PrepareMessageConsoleMsg("El dungeon cerrara en " & bestia_contador & " minutos.", FontTypeNames.FONTTYPE_GUILD))
            End If
            
            If bestia_contador = 0 Then
                Call SendData(SendTarget.toMap, Bestia_Map, PrepareMessageConsoleMsg("El portal ha sido cerrado.", FontTypeNames.FONTTYPE_GUILD))
                For i = 1 To ConnGroups(Bestia_Map).Count()
                    tempIndex = ConnGroups(Bestia_Map).Item(i)
                    
                    If UserList(tempIndex).ConnIDValida Then
                        Call WarpUserChar(tempIndex, bestia_tp.map, bestia_tp.x, bestia_tp.y, True)
                    End If
            
                Next i
            End If
            
        End If
End Sub
Public Sub SummonBestia()
    Dim map As Integer, miPos As WorldPos
    Dim x As Long, y As Long
    
    map = RandomNumber(1, NumMaps)
    
    Do While MapInfo(map).Pk = False
        map = RandomNumber(1, NumMaps)
        
    Loop
    
    Call MensajeGlobal("La bestia ha aparecido en el mapa " & map & " - '" & MapInfo(map).Name & "', al morir abrira un portal al dungeon secreto donde solo podran entrar 2 Partys de hasta 5 usuarios. Quien mate al equipo contrario primero dominara el dungeon por 30 minutos", FontTypeNames.FONTTYPE_CONSEJOCAOSVesA)
    
    
    npci_Bestia = CrearNPC(NPCINDEX_Bestia, map, miPos)
End Sub

Public Function teamIsValid(ByVal Userindex As Integer) As Boolean

    '*************************************************
    'Author: Unknown
    'Last modified: 11/27/09 (Budi)
    'Adapte la funcion a los nuevos metodos de clsParty
    '*************************************************
    Dim i                                    As Integer

    Dim pi                                   As Integer

    Dim MembersOnline(1 To PARTY_MAXMEMBERS) As Integer
    
    Dim lError As String
    
    Dim tmpTeam As Byte

    pi = UserList(Userindex).PartyIndex
    
    tmpTeam = 1
    If pi > 0 Then
        Call Parties(pi).ObtenerMiembrosOnline(MembersOnline())
        
        For i = 1 To PARTY_MAXMEMBERS
            
            If MembersOnline(i) > 0 Then
            
                tmpTeam = tmpTeam + 1
                'team(tmpTeam) = MembersOnline(i)
                If tmpTeam > 5 Then ' se lleno el team
                    teamIsValid = False
                    Exit Function
                End If
                    
            End If
            
        Next i
        If tmpTeam <= 5 And tmpTeam > 2 Then teamIsValid = True
    End If
    
End Function

Public Function bestia_PuedeEntrar(ByVal ui As Integer) As Boolean
    Dim pi As Integer
    pi = UserList(ui).PartyIndex
    If pi > 0 Then
        If pi = bestia_partys(1) Or pi = bestia_partys(2) Then
            bestia_PuedeEntrar = True: Exit Function
        End If
        
        If teamIsValid(ui) = True Then
            bestia_PuedeEntrar = True
            If bestia_partys(1) > 0 Then
                bestia_partys(2) = pi
                'es la 2da party en entrar, cerramos el teleport.
                bestia_contador = 34
                Call MensajeGlobal("Ya han entrado 2 integrantes de partys diferntes al dungeon, tienen 3 minutos para entrar sus compañeros.", FontTypeNames.FONTTYPE_CITIZEN)
                
            Else
                bestia_partys(1) = pi
            End If
        Else
            Call WriteConsoleMsg(ui, "Debes formar una party de minimo 3 y maximo 5 jugadores para entrar al dungeon.", FontTypeNames.FONTTYPE_WARNING)
        End If
    Else
        Call WriteConsoleMsg(ui, "Debes formar una party de minimo 3 y maximo 5 jugadores para entrar al dungeon.", FontTypeNames.FONTTYPE_WARNING)
    End If
End Function

Public Sub Bestia_crearTP(ByVal map As Integer, ByVal x As Byte, ByVal y As Byte)

    Dim ET As obj
    
    ET.amount = 1
    ET.ObjIndex = TELEP_OBJ_INDEX ' + Radio
    
    With MapData(map, x, y)
        .TileExit.map = Bestia_Map
        .TileExit.x = bestia_xentrance
        .TileExit.y = bestia_yentrance
    
    End With
    
    bestia_tp.map = map
    bestia_tp.x = x
    bestia_tp.y = y
    If MapData(map, x, y).ObjInfo.ObjIndex > 0 Then MapData(map, x, y).ObjInfo.ObjIndex = 0
    
    Call MakeObj(ET, map, x, y)
    
    Call MensajeGlobal("La bestia ha muerto! Se ha abierto un portal en el mapa " & map & " x: " & x & " y: " & y, FontTypeNames.FONTTYPE_GUILD)
    
End Sub














