Attribute VB_Name = "mod_Torneos"
Option Explicit

Private Type tTeam
    Jugador() As Integer
    GanoRonda As Boolean
    YaPeleo As Boolean
    adversarioTeam As Byte
    Ocupado As Boolean
    arenai As Byte
    valid As Boolean
    teamStr As String
End Type

'
Private Type tTorneo
    faseString As String
    TamañoEquipos As Byte 'EN un torneo 1vs1, el team tiene 1 usuario.
    Equipo() As tTeam
    Activo As Boolean 'El evento esta activo?
    Comenzo As Boolean ' Ha comenzado la pelea?
    fase As Byte 'Fase del torneo(Ej: un torneo de 16 teams tiene 4 fases
                          ' 8vs8 -- quedan 8. 4vs4 quedan 4 -- 2vs2 quedan 2 1vs1 queda el ganador.
                          'un torneo de 64 teams tiene 8 fases.
    TeamsRestantes As Byte 'Al comenzar nueva ronda,
                                             'ej_ Quedan 4 equipos. Se coloca como valor 4. Al ir moriendo los equipos se va restando
    CuposRestantes As Byte
    uniqueid As Long
    gFinal As Boolean
    autoCancelar As Byte
End Type

Private Type tMapEvent
    Map As Integer
    X As Byte
    Y As Byte
    X2 As Byte
    Y2 As Byte
    Ocupada As Boolean
End Type


Public Arena(1 To 16) As tMapEvent

Private Const max_arenas_torneo As Byte = 16
Private Const Mapa_Torneo As Byte = 222
Private Const Torneo_EsperaX As Byte = 33
Private Const Torneo_EsperaY As Byte = 33

Public Torneo As tTorneo

Public Sub CargarArenasTorneo()
    Dim ler As clsIniManager
    Dim i As Long
    
    ler = New clsIniManager
    
    Call ler.Initialize(DatPath & "ArenasTorneo.dat")
    
    For i = LBound(Arena) To UBound(Arena)
        Arena(i).X = ler.GetValue("ZONA" & CStr(i), "X")
        Arena(i).X2 = ler.GetValue("ZONA" & CStr(i), "X2")
        Arena(i).Y = ler.GetValue("ZONA" & CStr(i), "Y")
        Arena(i).Y2 = ler.GetValue("ZONA" & CStr(i), "Y2")
    Next
    
    Set ler = Nothing
End Sub

Public Sub cancelar_Torneo(Optional ByVal faltaParticipantes As Boolean = False)
    Dim i As Long
    Dim X As Long
    
    With Torneo
        If .Activo Then
            Call MensajeGlobal("Torneo> El evento ha sido cancelado" & IIf(faltaParticipantes = True, " por falta de participantes.", "."), FontTypeNames.FONTTYPE_GUILD)
            
            For i = 1 To .fase * 2
                With .Equipo(i)
                    For X = 1 To Torneo.TamañoEquipos
                        If .Jugador(X) > 0 Then
                            Call WarpUserChar(.Jugador(X), UserList(.Jugador(X)).userTorneo.lastPos.Map, UserList(.Jugador(X)).userTorneo.lastPos.X, UserList(.Jugador(X)).userTorneo.lastPos.Y, True)
                            With UserList(.Jugador(X))
                                .userTorneo.EnTorneo = False
                                .userTorneo.teamIndex = 0
                            End With
                        End If
                    Next X
                End With
            Next i

            .Activo = False
        End If
    End With
End Sub

Public Sub CheckTorneoData(ByVal ui As Integer, ByVal uniqueid As Long, ByVal torneo_data As String)
    
    Dim ti As Byte
    Dim fase As Byte
    Dim advt As Byte
    Dim gano As Byte
    Dim i As Long
    Dim users() As String
    Dim compaI As Integer
    
    If Torneo.Activo Then
        If Torneo.uniqueid = uniqueid Then
            If Torneo.TamañoEquipos = 1 Then
                ti = val(ReadField(1, torneo_data, "-"))
                fase = val(ReadField(2, torneo_data, "-"))
                gano = val(ReadField(3, torneo_data, "-"))
                
                If Torneo.fase = fase Then
                    'no paso de ronda todavia, lo metemos en el array y lo teletransportamos.
                    If ti Mod 2 = 0 Then advt = ti - 1 Else advt = ti + 1
                    If Torneo.Equipo(advt).GanoRonda = False And gano = 1 Then
                        Torneo.Equipo(ti).Jugador(1) = ui
                        Torneo.Equipo(ti).GanoRonda = True
                        Torneo.Equipo(ti).YaPeleo = True
                        
                        With UserList(ui)
                            .userTorneo.lastPos = .Pos
                            .userTorneo.EnTorneo = True
                            .userTorneo.teamIndex = ti
                        End With
                        
                        Call WarpUserChar(ui, Mapa_Torneo, Torneo_EsperaX + 3, Torneo_EsperaY, True)
                        
                    End If
                End If
            Else
            
                users = Split(torneo_data, "-")
                For i = 0 To UBound(users)
                    If UCase$(users(i)) <> UCase$(UserList(ui).Name) Then
                        compaI = NameIndex(users(i))
                        If compaI > 0 Then Exit For
                    End If
                Next i
                ti = UserList(compaI).userTorneo.teamIndex
                
                For i = 1 To Torneo.TamañoEquipos
                    If Torneo.Equipo(ti).Jugador(i) = 0 Then _
                        Torneo.Equipo(ti).Jugador(i) = ui: Exit For
                Next i
                With UserList(ui)
                    .userTorneo.lastPos = .Pos
                    .userTorneo.EnTorneo = True
                    .userTorneo.teamIndex = ti
                    Call WarpUserChar(ui, UserList(compaI).Pos.Map, UserList(compaI).Pos.X, UserList(compaI).Pos.Y, True)
                End With
            End If
        End If
    End If
End Sub


Public Function get_torneodata(ByVal ui As Integer) As String
    Dim i As Long, ti As Byte, tStr As String
    ti = UserList(ui).userTorneo.teamIndex
    If Torneo.TamañoEquipos > 1 Then
        
        get_torneodata = Torneo.Equipo(ti).teamStr
    Else
        tStr = ti & "-" & Torneo.fase & "-" & IIf(Torneo.Equipo(ti).GanoRonda = True, 1, 0)
    End If
    get_torneodata = tStr
End Function

Public Sub CrearTorneo(ByVal tamanoEquipo As Byte, cupos As Byte)
    With Torneo
        .Activo = True
        .fase = cupos / 2
        .TamañoEquipos = tamanoEquipo
        .TeamsRestantes = cupos
        .CuposRestantes = cupos
        .uniqueid = CLng(RandomNumber(0, 9) & RandomNumber(0, 9) & RandomNumber(0, 9) & RandomNumber(1, 9) & RandomNumber(1, 9) & RandomNumber(1, 9))
        .autoCancelar = 10
        'el torneo creado crea una clave de 6 digitos unica del evento, para el sistema de recuperacion de pj al torneo x desconexion.
    End With
    
    Call MensajeGlobal("Torneo> Ha iniciado un torneo con modalidad " & tamanoEquipo & "vs" & tamanoEquipo & _
                                ", con espacio para " & cupos & IIf(tamanoEquipo > 1, " equipos. Para ingresar primero debes crear una party con todos los integrantes de tu equipo dentro, luego escribe /TORNEO para inscribirte.", " jugadores. Para ingresar al torneo escribe /TORNEO"), FontTypeNames.FONTTYPE_INFO)

End Sub

Private Function dameTeamSlot() As Byte
    Dim i As Long
    For i = 1 To Torneo.fase * 2
        If Torneo.Equipo(i).Ocupado = False Then
            dameTeamSlot = i
            Exit Function
        End If
    Next i
End Function

Public Sub IngresaTorneo(ByVal ui As Integer)
    With UserList(ui)
        Dim Error As String, team() As Integer, i As Long, nSlot As Byte
        Dim teamStr As String
        Call PuedeTorneo(ui, Error)
        
        If LenB(Error) > 0 Then
            Call WriteConsoleMsg(ui, "Torneo> " & Error, FontTypeNames.FONTTYPE_CONSEJOCAOS)
            Exit Sub
        End If
        
        If Torneo.TamañoEquipos > 1 Then
            ReDim team(1 To Torneo.TamañoEquipos)
            
            If getTeam(ui, team) Then
                'el equipo esta en condiciones de entrar
                'los warpeamos y guardamos en la db del torneo
                nSlot = dameTeamSlot '(Torneo.Fase * 2) - Torneo.CuposRestantes
                
                Torneo.Equipo(nSlot).Ocupado = True
                For i = 1 To Torneo.TamañoEquipos
                    
                    Torneo.Equipo(nSlot).Jugador(i) = team(i)
                    
                    With UserList(team(i))
                        .userTorneo.EnTorneo = True
                        .userTorneo.lastPos = .Pos
                        .userTorneo.teamIndex = nSlot
                    End With
                    
                    teamStr = teamStr & UserList(team(i)).Name & "-"
                    
                    Call WarpUserChar(team(i), Mapa_Torneo, Torneo_EsperaX + RandomNumber(1, 5), Torneo_EsperaY + RandomNumber(1, 5), True)
                    
                Next i
                Torneo.Equipo(nSlot).teamStr = Left$(teamStr, LenB(teamStr) - 1)
                Torneo.CuposRestantes = Torneo.CuposRestantes - 1
                
                If Torneo.CuposRestantes = 0 Then
                    Call IniciarPelea
                End If
                
            End If
            
        Else ' es torneo 1vs1. mas facil
                'lo warpeamos y guardamos en la db del torneo
                nSlot = dameTeamSlot '(Torneo.Fase * 2) - Torneo.CuposRestantes
                
                Torneo.Equipo(nSlot).Ocupado = True

                Torneo.Equipo(nSlot).Jugador(1) = ui
                
                With UserList(ui)
                    .userTorneo.EnTorneo = True
                    .userTorneo.lastPos = .Pos
                    .userTorneo.teamIndex = nSlot
                End With
                
                Call WarpUserChar(ui, Mapa_Torneo, Torneo_EsperaX + RandomNumber(1, 5), Torneo_EsperaY + RandomNumber(1, 5), True)

                Torneo.CuposRestantes = Torneo.CuposRestantes - 1
                
                If Torneo.CuposRestantes = 0 Then
                    Call IniciarPelea
                End If
        End If
        
    End With
End Sub


Private Sub IniciarPelea()
    'Call MensajeGlobal("Torneo> Ha comenzado la pelea", FontTypeNames.FONTTYPE_GUILD)
    
    Dim i As Long, X As Long, currTeam As Byte
    Dim strTeam1 As String, strTeam2 As String
    Dim teamvalid(1) As Boolean
    Dim winner As Byte
    
    For i = 1 To IIf((Torneo.TeamsRestantes / 2) >= 16, 16, (Torneo.TeamsRestantes / 2))
        currTeam = i * 2
    
        strTeam1 = vbNullString
        strTeam2 = vbNullString
        
        Arena(i).Ocupada = True
        
        Torneo.Equipo(currTeam).adversarioTeam = currTeam - 1
        Torneo.Equipo(currTeam - 1).adversarioTeam = currTeam
        
        Torneo.Equipo(currTeam).arenai = i
        Torneo.Equipo(currTeam - 1).arenai = i
        Torneo.Equipo(currTeam).GanoRonda = False
        Torneo.Equipo(currTeam - 1).GanoRonda = False
        Torneo.Equipo(currTeam).YaPeleo = True
        Torneo.Equipo(currTeam - 1).YaPeleo = True
        
        Torneo.Equipo(currTeam).valid = False
        Torneo.Equipo(currTeam - 1).valid = False
        
        For X = 1 To Torneo.TamañoEquipos
            If Torneo.Equipo(currTeam).Jugador(X) > 0 Then
                Call WarpUserChar(Torneo.Equipo(currTeam).Jugador(X), Mapa_Torneo, Arena(i).X, Arena(i).Y, False)
                Call WritePauseToggle(Torneo.Equipo(currTeam).Jugador(X))
                
                If UserList(Torneo.Equipo(currTeam).Jugador(X)).flags.Muerto > 0 Then _
                    Call RevivirUsuario(Torneo.Equipo(currTeam).Jugador(X))
                
                Torneo.Equipo(currTeam).valid = True
                
                UserList(Torneo.Equipo(currTeam).Jugador(X)).userTorneo.cuentaRegresiva = 8
                
            End If
            
            If Torneo.Equipo(currTeam - 1).Jugador(X) > 0 Then
                Call WarpUserChar(Torneo.Equipo(currTeam - 1).Jugador(X), Mapa_Torneo, Arena(i).X2 - (X - 1), Arena(i).Y2, False)
                Call WritePauseToggle(Torneo.Equipo(currTeam - 1).Jugador(X)) '
                
                If UserList(Torneo.Equipo(currTeam - 1).Jugador(X)).flags.Muerto > 0 Then _
                    Call RevivirUsuario(Torneo.Equipo(currTeam - 1).Jugador(X))
                    
                Torneo.Equipo(currTeam - 1).valid = True
                
                UserList(Torneo.Equipo(currTeam - 1).Jugador(X)).userTorneo.cuentaRegresiva = 8

            End If
        Next X

       ' If teamvalid(0) = False Then
      '      Call finalizarPelea(currTeam - 1, currTeam, i)
       ' ElseIf teamvalid(1) = False Then
       '     Call finalizarPelea(currTeam, currTeam - 1, i)
        If teamvalid(1) = True And teamvalid(0) = True Then
            Call MensajeGlobal("Torneo> ARENA " & i & ": (" & Torneo.Equipo(currTeam).teamStr & ") vs (" & Torneo.Equipo(currTeam - 1).teamStr, FontTypeNames.FONTTYPE_GUILD)
        End If
    Next i
    
    For i = 1 To IIf(Torneo.TeamsRestantes >= 32, 32, Torneo.TeamsRestantes)
        'recorro los teams por si habia uno invalido, lo mando a la siguiente ronda.
        If Torneo.Equipo(i).valid = False Then
            'If Torneo.TeamsRestantes = 4 Then ' ES LA FINAL, ENTONCES GANA EL TORNEO X DESCONEXION.
            '    If i Mod 2 = 0 Then 'ES PAR
            '        Call GanaTorneo(i - 1, i)
            '        Call MensajeGlobal("Torneo> El equipo (" & Torneo.Equipo(i - 1).teamStr & ") ha ganado el torneo por desconexion de el/los oponente(s)", FontTypeNames.FONTTYPE_GUILD)
                    
            '    Else 'ES IMPAR
            '        Call GanaTorneo(i + 1, i)
            '        Call MensajeGlobal("Torneo> El equipo (" & Torneo.Equipo(i).teamStr & ") ha ganado el torneo por desconexion de el/los oponente(s)", FontTypeNames.FONTTYPE_GUILD)
                
            '    End If
           ' Else
                If i Mod 2 = 0 Then 'ES PAR
                    Call finalizarPelea(i - 1, i, Torneo.Equipo(i).arenai)
                    Call MensajeGlobal("Torneo> El equipo (" & Torneo.Equipo(i - 1).teamStr & ") ha pasado automaticamente a la siguiente ronda por desconexion de el/los oponente(s)", FontTypeNames.FONTTYPE_GUILD)
                    
                Else 'ES IMPAR
                    Call finalizarPelea(i + 1, i, Torneo.Equipo(i).arenai)
                    Call MensajeGlobal("Torneo> El equipo (" & Torneo.Equipo(i).teamStr & ") ha pasado automaticamente a la siguiente ronda por desconexion de el/los oponente(s)", FontTypeNames.FONTTYPE_GUILD)
                
                End If
            'End If
        End If
    Next i
End Sub

Public Sub userTorneo_PasaSegundo(ByVal ui As Integer)
    With UserList(ui)
        If .userTorneo.cuentaRegresiva > 0 Then
            .userTorneo.cuentaRegresiva = .userTorneo.cuentaRegresiva - 1
            If .userTorneo.cuentaRegresiva = 0 Then
                Call WriteConsoleMsg(ui, "Torneo> ¡¡¡YAAAA!!!", FontTypeNames.FONTTYPE_GUILD)
                Call WritePauseToggle(ui)
            Else
                Call WriteConsoleMsg(ui, "Torneo> ¡¡¡" & .userTorneo.cuentaRegresiva & "!!!", FontTypeNames.FONTTYPE_GUILD)
            End If
        End If
    End With
End Sub

Public Function torneo_PuedeAtacar(ByVal Atacante As Integer, ByVal victima As Integer) As Boolean
    If UserList(Atacante).userTorneo.teamIndex = UserList(victima).userTorneo.teamIndex Then
        torneo_PuedeAtacar = False
        Exit Function
    End If
    
    If UserList(Atacante).userTorneo.cuentaRegresiva > 0 Then
        torneo_PuedeAtacar = False
        Exit Function
    End If
    
    torneo_PuedeAtacar = True
End Function

Private Function teamIsDeath(ByVal team As Byte, Optional ByVal onlydesc As Boolean = False) As Boolean
        
        Dim i As Long
        
        teamIsDeath = True
        
        For i = 1 To Torneo.TamañoEquipos
            If Torneo.Equipo(team).Jugador(i) > 0 Then
                If onlydesc = False Then
                    If UserList(Torneo.Equipo(team).Jugador(i)).flags.Muerto = 0 Then
                        teamIsDeath = False: Exit Function
                    End If
                End If
            End If
        Next i
        
End Function

Private Sub GranFinal()
        'Reordenamos los teams:
        '#1 = FINALISTA1
        '#2 = FINALISTA2
        '#3 = 3ERPUESTO1
        '#4 = 3ERPUESTO2
        Dim tmpTeams(1 To 4) As tTeam, i As Long
        Dim Count As Byte, count1 As Byte
        With Torneo
            .TeamsRestantes = 4
            .gFinal = True
            Count = 1
            count1 = 3
            For i = 1 To 4
                If .Equipo(i).GanoRonda = True Then
                    tmpTeams(Count) = .Equipo(i)
                    tmpTeams(Count).GanoRonda = False
                    tmpTeams(Count).YaPeleo = False
                    Count = Count + 1
                Else
                    tmpTeams(count1) = .Equipo(i)
                    tmpTeams(count1).GanoRonda = False
                    tmpTeams(count1).YaPeleo = False
                    count1 = count1 + 1
                End If
            Next i
            .Equipo = tmpTeams
            Call IniciarPelea
        End With
End Sub

Public Sub PasarRonda()
'este sub actualiza las variables del torneo, y da inicio a las luchas correspondientes a la siguiente ronda (INICIARPELEA al final del sub)
    Dim i As Long, X As Long, tmpTeams() As tTeam
    Dim b As Byte
    With Torneo
        .fase = .fase / 2
        .TeamsRestantes = .fase * 2
        b = 1
        ReDim tmpTeams(1 To .TeamsRestantes)
        
        For i = 1 To .TeamsRestantes * 2 ' recorremos el array anterior de teams y guardamos los ganadores
            If .Equipo(i).GanoRonda = True Then tmpTeams(b) = .Equipo(i): b = b + 1
            
        Next i
        ReDim .Equipo(1 To .TeamsRestantes)
         .Equipo = tmpTeams
         
         
        Select Case .TeamsRestantes
            Case 16
                .faseString = "Octavos de final"
                
            Case 8
                .faseString = "Cuartos de Final"
                
            Case 4
                .faseString = "Semifinal"
                
            Case 2
                .faseString = "Gran Final"
            
            Case Else
                .faseString = "Fase eliminatoria(Quedan " & .TeamsRestantes & IIf(.TamañoEquipos = 1, " jugadores)", " equipos)")
                
        End Select
        Call IniciarPelea
    End With
End Sub

Private Function buscarArena() As Byte
    Dim i As Long
    For i = 1 To max_arenas_torneo
        If Arena(i).Ocupada = False Then buscarArena = i: Exit Function
    Next i
End Function

Private Sub buscarLuchasIniciar()
    Dim i As Long, X As Long
    Dim team1 As Byte, team2 As Byte
    Dim iArena As Byte
    For i = 1 To Torneo.fase * 2
        If Torneo.Equipo(i).YaPeleo = False And (i Mod 2 = 0) Then
                team1 = i - 1
                team2 = i

                iArena = buscarArena
                
                Arena(iArena).Ocupada = True
                
                Torneo.Equipo(team1).adversarioTeam = team2
                Torneo.Equipo(team2).adversarioTeam = team1
                
                Torneo.Equipo(team1).GanoRonda = False
                Torneo.Equipo(team2).GanoRonda = False
                
                Torneo.Equipo(team1).YaPeleo = True
                Torneo.Equipo(team2).YaPeleo = True
                
                For X = 1 To Torneo.TamañoEquipos
                    If Torneo.Equipo(team1).Jugador(X) > 0 Then
                    
                        Call WarpUserChar(Torneo.Equipo(team1).Jugador(X), Mapa_Torneo, Arena(iArena).X, Arena(iArena).Y, False)
                        Call WritePauseToggle(Torneo.Equipo(team1).Jugador(X))
                        
                        UserList(Torneo.Equipo(team1).Jugador(X)).userTorneo.cuentaRegresiva = 8
                        
                    End If
                    
                    If Torneo.Equipo(team2).Jugador(X) > 0 Then
                    
                        Call WarpUserChar(Torneo.Equipo(team2).Jugador(X), Mapa_Torneo, Arena(iArena).X2 - (X - 1), Arena(iArena).Y2, False)
                        Call WritePauseToggle(Torneo.Equipo(team2).Jugador(X)) '
                        
                        UserList(Torneo.Equipo(team2).Jugador(X)).userTorneo.cuentaRegresiva = 8
                        
                    End If
                    
                Next X
                
                Call MensajeGlobal("Torneo> ARENA N" & iArena & ": (" & Torneo.Equipo(team1).teamStr & ") vs (" & Torneo.Equipo(team2).teamStr, FontTypeNames.FONTTYPE_GUILD)

        End If
    Next i
End Sub

Private Sub finalizarPelea(ByVal teamw As Byte, ByVal teaml As Byte, ByVal arenai As Byte)
    Dim i As Long, miPos As WorldPos
    
    Torneo.TeamsRestantes = Torneo.TeamsRestantes - 1
    Torneo.Equipo(teaml).GanoRonda = False

    Torneo.Equipo(teamw).GanoRonda = True
    
    For i = 1 To Torneo.TamañoEquipos
        'warpeamos team perdedor a lastpos
        If Torneo.Equipo(teaml).Jugador(i) > 0 Then
            miPos = UserList(Torneo.Equipo(teaml).Jugador(i)).userTorneo.lastPos
            UserList(Torneo.Equipo(teaml).Jugador(i)).userTorneo.EnTorneo = False
            UserList(Torneo.Equipo(teaml).Jugador(i)).userTorneo.teamIndex = 0
            Call WarpUserChar(Torneo.Equipo(teaml).Jugador(i), miPos.Map, miPos.X, miPos.Y, True)
        End If
        
        'warpeamos team ganador a la sala de espera.
        If Torneo.Equipo(teamw).Jugador(i) > 0 Then
            If Torneo.gFinal = True Then
                miPos = UserList(Torneo.Equipo(teamw).Jugador(i)).userTorneo.lastPos
                UserList(Torneo.Equipo(teamw).Jugador(i)).userTorneo.EnTorneo = False
                UserList(Torneo.Equipo(teamw).Jugador(i)).userTorneo.teamIndex = 0
                Call WarpUserChar(Torneo.Equipo(teamw).Jugador(i), miPos.Map, miPos.X, miPos.Y, True)
            Else
                Call WarpUserChar(Torneo.Equipo(teamw).Jugador(i), Mapa_Torneo, Torneo_EsperaX + RandomNumber(1, 5), Torneo_EsperaY + RandomNumber(1, 5), True)
            End If
        End If
        
    Next i
    Arena(arenai).Ocupada = False
    
    buscarLuchasIniciar 'vemos si hay luchas pendientes, debido a la cantidad limitada de arenas para pelear.
End Sub

Public Sub torneo_Muere(ByVal ui As Integer, ByVal desconexion As Boolean)

    With UserList(ui)
        Dim i As Long, advTeam As Byte, teamIndex As Byte, miPos As WorldPos
        teamIndex = .userTorneo.teamIndex
        advTeam = Torneo.Equipo(.userTorneo.teamIndex).adversarioTeam ' recuperamos el teamindex adversario(que gano la ronda)
        
        If desconexion Then
            For i = 1 To Torneo.TamañoEquipos
                If Torneo.Equipo(UserList(ui).userTorneo.teamIndex).Jugador(i) Then
                    Torneo.Equipo(UserList(ui).userTorneo.teamIndex).Jugador(i) = 0
                End If
            Next i
            Call WarpUserChar(ui, UserList(ui).userTorneo.lastPos.Map, UserList(ui).userTorneo.lastPos.X, UserList(ui).userTorneo.lastPos.Y, True)
        End If
        
        If Torneo.Comenzo = True Then
            If teamIsDeath(.userTorneo.teamIndex) Then 'solo terminamos la pelea si todo el team esta muerto
            'da igual si el torneo es 1vs1, deberia ejecutarse correctamente aqui tambien.
                
                Call finalizarPelea(advTeam, teamIndex, Torneo.Equipo(teamIndex).arenai)
                
                'If Torneo.TeamsRestantes = 1 And Torneo.fase = 2 Then 'termina el torneo
                '    Call GanaTorneo(advTeam, .userTorneo.teamindex)
                
                ''***FASE =2 Semifinal
            
                If Torneo.TeamsRestantes = Torneo.fase And Torneo.TeamsRestantes > 2 Then 'termino esta fase
                    Call PasarRonda
                ElseIf Torneo.TeamsRestantes = Torneo.fase And Torneo.fase = 2 And Torneo.gFinal = False Then
                    Call GranFinal
                ElseIf Torneo.gFinal = True And Torneo.TeamsRestantes = 2 Then
                    'Ya finalizo el torneo, y mando a todos para casa.
                    Call FinTorneo
                End If
                
            End If
        Else
            If teamIsDeath(.userTorneo.teamIndex, True) Then
                'no se llenaron los cupos todavia, entonces liberamos el cupo xq todo el team se desconecto.
                Torneo.CuposRestantes = Torneo.CuposRestantes + 1
                
            End If
        End If
    End With
End Sub

Private Sub DarPremiosaTeam(ByVal teamIndex As Byte, ByVal trofeo As Byte)
    Dim i As Long, miObj As obj, ui As Integer
    Dim players() As String
    miObj.amount = 1
    
    Select Case trofeo
        Case 1 'oro
            Call MensajeGlobal("Torneo> 1er puesto para" & IIf(Torneo.TamañoEquipos = 1, ": ", " el equipo de: ") & Torneo.Equipo(teamIndex).teamStr, FontTypeNames.FONTTYPE_GUILD)
            miObj.ObjIndex = Trofeo_Oro
            
        Case 2 'plata
            Call MensajeGlobal("Torneo> 2do puesto para" & IIf(Torneo.TamañoEquipos = 1, ": ", " el equipo de: ") & Torneo.Equipo(teamIndex).teamStr, FontTypeNames.FONTTYPE_GUILD)
            miObj.ObjIndex = Trofeo_Plata
            
        Case 3 'bronce
            Call MensajeGlobal("Torneo> 3er puesto para" & IIf(Torneo.TamañoEquipos = 1, ": ", " el equipo de: ") & Torneo.Equipo(teamIndex).teamStr, FontTypeNames.FONTTYPE_GUILD)
            miObj.ObjIndex = Trofeo_Bronce
            
    End Select
    
    players = Split(Torneo.Equipo(teamIndex).teamStr, "-")
    
    For i = 1 To Torneo.TamañoEquipos
        ui = NameIndex(players(i))
        If ui > 0 Then
            If Not MeterItemEnInventario(Torneo.Equipo(teamIndex).Jugador(i), miObj) Then
                Call TirarItemAlPiso(UserList(Torneo.Equipo(teamIndex).Jugador(i)).userTorneo.lastPos, miObj)
                Call WriteConsoleMsg(Torneo.Equipo(teamIndex).Jugador(i), "Torneo> No tenias suficiente espacio en el inventario, tu Trofeo cayo en el piso.", FontTypeNames.FONTTYPE_INFO)
            End If
        Else
            Call darItemaUserOffline(players(i), miObj.ObjIndex, miObj.amount)
        End If
    Next i
    
End Sub

Public Sub FinTorneo()
    '   entrega premio notifica x global y manda  a todos a posicioon inicial
    Dim i As Long

    For i = 1 To 4
        If i < 3 Then
            'finalistas.
            If Torneo.Equipo(i).GanoRonda = True Then
                DarPremiosaTeam i, 1
            Else
                DarPremiosaTeam i, 2
            End If
        Else
            'x el 3er puesto
            If Torneo.Equipo(i).GanoRonda = True Then
                DarPremiosaTeam i, 3
            End If
        End If
    Next i
End Sub

Public Function getTeam(ByVal Userindex As Integer, ByRef team() As Integer) As Boolean

    '*************************************************
    'Author: Unknown
    'Last modified: 11/27/09 (Budi)
    'Adapte la funcion a los nuevos metodos de clsParty
    '*************************************************
    Dim i                                    As Integer

    Dim PI                                   As Integer

    Dim MembersOnline(1 To PARTY_MAXMEMBERS) As Integer
    
    Dim lError As String
    
    Dim tmpTeam As Byte

    PI = UserList(Userindex).PartyIndex
    
    tmpTeam = 1
    If PI > 0 Then
        Call Parties(PI).ObtenerMiembrosOnline(MembersOnline())
        team(tmpTeam) = Userindex
        For i = 1 To PARTY_MAXMEMBERS
            
            If MembersOnline(i) > 0 Then
               ' Text = Text & " - " & UserList(MembersOnline(i)).Name & " (" & Fix(Parties(PI).MiExperiencia(MembersOnline(i))) & ")"
                lError = ""
                Call PuedeTorneo(MembersOnline(i), lError)
                If LenB(lError) > 0 Then
                    Call WriteConsoleMsg(Userindex, "Torneo> Uno de los usuarios no cumple los requisitos para entrar", FontTypeNames.FONTTYPE_CONSEJOCAOS)
                    getTeam = False
                Else 'cumple los requisitos
                    tmpTeam = tmpTeam + 1
                    team(tmpTeam) = MembersOnline(i)
                    If tmpTeam = Torneo.TamañoEquipos Then ' se lleno el team
                        getTeam = True
                        Exit Function
                    End If
                End If
                
            End If
            
        Next i
        
    Else
        getTeam = "Debes crear una party con los integrantes de tu equipo antes de ingresar"
    End If
    
End Function

Private Sub PuedeTorneo(ByVal Userindex As Integer, ByRef lError As String)
    With UserList(Userindex)
        
        If Torneo.Activo = False Then
            lError = "Evento inactivo"
            Exit Sub
        End If
        
        If Torneo.Comenzo = True Then
            lError = "Las inscripciones ya fueron cerradas"
            Exit Sub
        End If
        
        If Torneo.CuposRestantes <= 0 Then
            lError = "Cupos completos"
            Exit Sub
        End If
        
        If (.flags.Muerto <> 0) Then
            lError = "Estás muerto"
            Exit Sub
        End If
        
        If (.Counters.Pena <> 0) Then
            lError = "Estás en la cárcel"
            Exit Sub
        End If
        
        If .Stats.ELV < 25 Then
            lError = "Necesitas ser nivel 25"
            Exit Sub
        End If
    
        If MapInfo(.Pos.Map).Pk = True Then
            lError = "Estás en una zona insegura"
            Exit Sub
        End If
        
        If .EnEvento = True Or .UserDeath.EnDeath = True Or .flags.SlotReto <> 0 Then
            lError = "Ya estás en un evento"
            Exit Sub
        End If
        
        If .Stats.Gld < 2000 Then
            lError = "No tenes suficiente oro"
            Exit Sub
        End If
        
    End With
End Sub




