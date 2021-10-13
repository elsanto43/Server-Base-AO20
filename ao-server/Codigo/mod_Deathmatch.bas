Attribute VB_Name = "mod_Deathmatch"
Option Explicit

Private Type tDeath
    UIndex() As Integer
    Activo As Boolean
    cuentaRegresiva As Integer
    CuposRestantes As Byte
    cupos As Byte
    CaenItems As Boolean
    UsersRestantes As Byte
    Comenzado As Boolean
End Type
'AVISO CUANDO COMIENZA
Public Death As tDeath
'Death: 295, 48,81 espera
'death: 295,49,50 pelea
Public Const MAPA_DEATH As Integer = 49
Private Const ESPERA_X As Byte = 29
Private Const ESPERA_Y As Byte = 30
Private Const PELEA_X As Byte = 66
Private Const PELEA_Y As Byte = 30
Public Sub IniciarDeath(ByVal cupos As Byte, ByVal CaenItems As Boolean)
        
    With Death
        If .Activo = True Then
            Call SendData(SendTarget.ToAdmins, 0, PrepareMessageConsoleMsg("Ya hay un DeathMatch en curso", FontTypeNames.FONTTYPE_INFOBOLD))
            Exit Sub
        End If
        .Activo = True
        .cuentaRegresiva = 10
        .cupos = cupos
        .CuposRestantes = cupos
        .CaenItems = CaenItems
        .UsersRestantes = cupos
        .Comenzado = False
        'Call LimpiarMapa(MAPA_DEATH)
        ReDim .UIndex(1 To cupos)
        Call MensajeGlobal("DeathMatch> El evento ha comenzado. Cupos disponibles: " & cupos & ". " & IIf(CaenItems = False, "Los items no se caen", "CAEN LOS ITEMS ") & ". Escribe /ENTRARDEATH para ingresar", FontTypeNames.FONTTYPE_GUILD)
    End With
End Sub

Public Sub Muere_Death(ByVal Userindex As Integer, Optional ByVal desconexion As Boolean = False)
    With UserList(Userindex)
    
        If .UserDeath.EnDeath = False Then Exit Sub
        .UserDeath.EnDeath = False
        .EnEvento = False
        
        If Death.CaenItems Then
            If Death.Comenzado = True Then
                Call TirarTodosLosItems(Userindex)
            End If
        End If
        
        Call WarpUserChar(Userindex, .UserDeath.lastPos.Map, .UserDeath.lastPos.x, .UserDeath.lastPos.Y, True)
        
        Dim LoopC As Long
        For LoopC = 1 To Death.cupos
            If Death.UIndex(LoopC) = Userindex Then
                Death.UIndex(LoopC) = 0
            End If
        Next LoopC
        
        If desconexion = True And Death.Comenzado = False Then
            Death.CuposRestantes = Death.CuposRestantes + 1
            Call MensajeGlobal("DeathMatch> Se ha liberado un cupo por la desconexión de " & .Name, FontTypeNames.FONTTYPE_GUILD)
        End If
        
        If Death.Comenzado = True Then
            Death.UsersRestantes = Death.UsersRestantes - 1
            
            If desconexion Then
                Call MensajeGlobal("DeathMatch> " & .Name & " ha sido descalificado por desconexion" & IIf(Death.UsersRestantes >= 2, ". Quedan " & Death.UsersRestantes & " usuarios vivos.", "."), FontTypeNames.FONTTYPE_GUILD)
            Else
                Call MensajeGlobal("DeathMatch> " & .Name & " ha muerto" & IIf(Death.UsersRestantes > 2, ". Quedan " & Death.UsersRestantes & " usuarios vivos.", "."), FontTypeNames.FONTTYPE_GUILD)
            End If
            
            If Death.UsersRestantes = 1 Then
                Call Death_Finish
            End If
        End If
    End With
End Sub

Private Sub Death_Finish()
    With Death
        '.Activo = False
       ' .Comenzado = False
        '.CuentaRegresiva = 10
        Dim LoopC As Long, winner As Integer
        For LoopC = 1 To .cupos
            If .UIndex(LoopC) > 0 Then
                winner = .UIndex(LoopC)
                Exit For
            End If
        Next LoopC

        Call MensajeGlobal("DeathMatch> Evento finalizado. Ganador: " & UserList(winner).Name & ". Premio: 50000 monedas de oro", FontTypeNames.FONTTYPE_GUILD)
        With UserList(winner)
            .Stats.Gld = .Stats.Gld + 50000
            Call WriteUpdateGold(winner)
            If Death.CaenItems = False Then
                WarpUserChar winner, .UserDeath.lastPos.Map, .UserDeath.lastPos.x, .UserDeath.lastPos.Y, True
                Death.Activo = False
                Death.Comenzado = False
                Death.cuentaRegresiva = 10
            Else
                WriteConsoleMsg winner, "DeathMatch> Tenés 1 minuto para agarrar items.", FontTypeNames.FONTTYPE_GUILD
                .UserDeath.SecondsBack = 60
            End If
        End With
        
    End With
End Sub

Public Sub EnterDeath(ByVal Userindex As Integer)
    With UserList(Userindex)
        Dim lError As String '<=esta es la variable
        Call PuedeDeath(Userindex, lError)
        If LenB(lError) <> 0 Then
            Call WriteConsoleMsg(Userindex, "DeathMatch> " & lError, FontTypeNames.FONTTYPE_INFO)
            Exit Sub 'Si tiene algun error, le decimos cual es y salimos.
        End If
        With .UserDeath
            .EnDeath = True
            .lastPos = UserList(Userindex).Pos
        End With
        
        .EnEvento = True
        With Death
            .CuposRestantes = .CuposRestantes - 1
            Dim LoopC As Long, Find As Byte
            For LoopC = 1 To .cupos
                If .UIndex(LoopC) <= 0 Then
                    Find = CByte(LoopC)
                    Exit For
                End If
            Next LoopC
            .UIndex(Find) = Userindex
            WarpUserChar Userindex, MAPA_DEATH, ESPERA_X, ESPERA_Y, True
            Call MensajeGlobal("DeathMatch> " & UserList(Userindex).Name & " ha ingresado al evento.", FontTypeNames.FONTTYPE_GUILD)
            If .CuposRestantes = 0 Then
                Death_Go
            End If
        End With
    End With
End Sub

Public Sub PassSecondDeath()
    With Death
        'Death_Finish
        If .Activo And .Comenzado = True And .cuentaRegresiva >= 0 Then
            Select Case .cuentaRegresiva
                Case 0
                    Call MensajeGlobal("DeathMatch> ¡Ya!", FontTypeNames.FONTTYPE_GUILD)
                    Call DEATH_GO1
                
                Case Else
                    Call MensajeGlobal("DeathMatch> ¡" & .cuentaRegresiva & "!", FontTypeNames.FONTTYPE_GUILD)
            
            End Select
            .cuentaRegresiva = .cuentaRegresiva - 1
        End If
    End With
End Sub
Sub CancelarDeath()
    With Death
        If .Activo = False Then Exit Sub
        Dim x As Long
        For x = 1 To .cupos
            If .UIndex(x) > 0 Then
                WarpUserChar .UIndex(x), UserList(.UIndex(x)).UserDeath.lastPos.Map, UserList(.UIndex(x)).UserDeath.lastPos.x, UserList(.UIndex(x)).UserDeath.lastPos.Y, True
                UserList(.UIndex(x)).UserDeath.EnDeath = False
                UserList(.UIndex(x)).EnEvento = False
            End If
        Next x
        .Activo = False
        .CaenItems = False
        .Comenzado = False
        .cuentaRegresiva = 10
        Call MensajeGlobal("DeathMatch> El evento ha sido cancelado", FontTypeNames.FONTTYPE_GUILD)
    End With
End Sub

Function death_PuedeAtacar(ByVal Userindex As Integer) As Boolean
    With Death
        If .Activo = True And .Comenzado = True And .cuentaRegresiva <= 0 Then
            death_PuedeAtacar = True
            Exit Function
        End If
        
        If .Activo = True And .cuentaRegresiva > 0 Then
            death_PuedeAtacar = False
            WriteConsoleMsg Userindex, "DeathMatch> Espera que termine la cuenta regresiva", FontTypeNames.FONTTYPE_GUILD
        End If
    End With
End Function

Private Sub DEATH_GO1()
    Dim x As Long
    For x = 1 To Death.cupos
        If Death.UIndex(x) > 0 Then
            WritePauseToggle Death.UIndex(x)
        End If
    Next x
End Sub

Private Sub Death_Go()
    With Death
        .Comenzado = True
        '.CuentaRegresiva = 10
        
        Dim x As Long
        For x = 1 To .cupos
            If .UIndex(x) > 0 Then
                WarpUserChar .UIndex(x), MAPA_DEATH, PELEA_X, PELEA_Y, True
                WritePauseToggle .UIndex(x)
            End If
        Next x
    End With
End Sub

Private Sub PuedeDeath(ByVal Userindex As Integer, ByRef lError As String)
    With UserList(Userindex)
        
        If Death.Activo = False Then
            lError = "Evento inactivo"
            Exit Sub
        End If
        
        If Death.CuposRestantes <= 0 Then
            lError = "Cupos completos"
            Exit Sub
        End If
        
        If .UserDeath.EnDeath = True Then
            lError = "Ya estás en el evento"
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
        
        If .EnEvento = True Then
            lError = "Ya estás en un evento"
            Exit Sub
        End If
        
        If .Stats.Gld < 2000 Then
            lError = "No tenes suficiente oro"
            Exit Sub
        End If
        
        
        
    End With
End Sub

Public Sub MensajeGlobal(ByVal Chat As String, ByVal FontIndex As FontTypeNames)
    Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg(Chat, FontIndex))
End Sub

