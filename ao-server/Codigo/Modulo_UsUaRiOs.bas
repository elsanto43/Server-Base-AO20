Attribute VB_Name = "UsUaRiOs"
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

'????????????????????????????
'                        Modulo Usuarios
'????????????????????????????
'Rutinas de los usuarios
'????????????????????????????


Public Sub RevivirUsuario(ByVal Userindex As Integer)
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    With UserList(Userindex)
        .flags.Muerto = 0
        .Stats.MinHp = .Stats.UserAtributos(eAtributos.Constitucion)
        
        If .Stats.MinHp > .Stats.MaxHp Then
            .Stats.MinHp = .Stats.MaxHp

        End If
        
        If .flags.Navegando = 1 Then
            Call ToggleBoatBody(Userindex)
        Else
            Call DarCuerpoDesnudo(Userindex)
            
            .Char.Head = .OrigChar.Head

        End If
        
        If .flags.Traveling Then
            .flags.Traveling = 0
            .Counters.goHome = 0
            Call WriteMultiMessage(Userindex, eMessages.CancelHome)

        End If
        
        Call ChangeUserChar(Userindex, .Char.body, .Char.Head, .Char.heading, .Char.WeaponAnim, .Char.ShieldAnim, .Char.CascoAnim)
        Call WriteUpdateUserStats(Userindex)

    End With

End Sub

Public Sub ToggleBoatBody(ByVal Userindex As Integer)
    '***************************************************
    'Author: ZaMa
    'Last Modification: 25/07/2010
    'Gives boat body depending on user alignment.
    '25/07/2010: ZaMa - Now makes difference depending on faccion and atacable status.
    '***************************************************

    Dim Ropaje        As Integer

    Dim EsFaccionario As Boolean

    Dim NewBody       As Integer
    
    With UserList(Userindex)
 
        .Char.Head = 0

        If .Invent.BarcoObjIndex = 0 Then Exit Sub
        
        Ropaje = ObjData(.Invent.BarcoObjIndex).Ropaje
        
        ' Criminales y caos
        If esLegion(Userindex) Then
            
            EsFaccionario = UserList(Userindex).Faccion.Jerarquia > 0
            
            Select Case Ropaje

                Case iBarca

                    If EsFaccionario Then
                        NewBody = iBarcaCaos
                    Else
                        NewBody = iBarcaPk

                    End If
                
                Case iGalera

                    If EsFaccionario Then
                        NewBody = iGaleraCaos
                    Else
                        NewBody = iGaleraPk

                    End If
                    
                Case iGaleon

                    If EsFaccionario Then
                        NewBody = iGaleonCaos
                    Else
                        NewBody = iGaleonPk

                    End If

                Case iFragataFantasmal
                    NewBody = iFragataFantasmal

            End Select
        
            ' Ciudas y Armadas
        ElseIf esArmada(Userindex) Then
            
            EsFaccionario = esArmada(Userindex)

                Select Case Ropaje

                    Case iBarca

                        If EsFaccionario Then
                            NewBody = iBarcaReal
                        Else
                            NewBody = iBarcaCiuda

                        End If
                    
                    Case iGalera

                        If EsFaccionario Then
                            NewBody = iGaleraReal
                        Else
                            NewBody = iGaleraCiuda

                        End If
                        
                    Case iGaleon

                        If EsFaccionario Then
                            NewBody = iGaleonReal
                        Else
                            NewBody = iGaleonCiuda

                        End If

                    Case iFragataFantasmal
                        NewBody = iFragataFantasmal

                End Select
            
                '
        Else 'es neutral
            
                Select Case Ropaje

                    Case iBarca

                      '  If EsFaccionario Then
                    '        NewBody = iBarcaReal
                      ' Else
                            NewBody = iBarcaCiudaAtacable

                     '
                    Case iGalera

                   '     If EsFaccionario Then
                     '       NewBody = iGaleraReal
                     '   Else
                            NewBody = iGaleraCiudaAtacable

                    '    End If
                        
                    Case iGaleon

                     '   If EsFaccionario Then
                     '       NewBody = iGaleonReal
                '        Else
                          NewBody = iGaleonCiudaAtacable

                    '    End If

                    Case iFragataFantasmal
                        NewBody = iFragataFantasmal

                End Select
            
        End If
        
        .Char.body = NewBody
        .Char.ShieldAnim = NingunEscudo
        .Char.WeaponAnim = NingunArma
        .Char.CascoAnim = NingunCasco

    End With

End Sub


Public Sub ToggleMonturaBody(ByVal Userindex As Integer)
    '***************************************************
    'Author: Recix
    'Last Modification: 12/01/2020
    'Gives montura body
    '***************************************************

    With UserList(Userindex)
        
        If .Invent.MonturaObjIndex = 0 Then Exit Sub
 
        .Char.body = ObjData(.Invent.MonturaObjIndex).Ropaje
        .Char.ShieldAnim = NingunEscudo
        .Char.WeaponAnim = NingunArma

    End With

End Sub

Public Sub ChangeUserChar(ByVal Userindex As Integer, _
                          ByVal body As Integer, _
                          ByVal Head As Integer, _
                          ByVal heading As Byte, _
                          ByVal Arma As Integer, _
                          ByVal Escudo As Integer, _
                          ByVal casco As Integer)

    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************
    With UserList(Userindex).Char
        .body = body
        .Head = Head
        .heading = heading
        .WeaponAnim = Arma
        .ShieldAnim = Escudo
        .CascoAnim = casco
        
        Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessageCharacterChange(body, Head, heading, .CharIndex, Arma, Escudo, .FX, .loops, casco))
        Call SendData(SendTarget.ToPCAreaButIndex, Userindex, PrepareMessageHeadingChange(heading, .CharIndex))

    End With

End Sub

Public Function GetWeaponAnim(ByVal Userindex As Integer, _
                              ByVal ObjIndex As Integer) As Integer

    '***************************************************
    'Author: Torres Patricio (Pato)
    'Last Modification: 03/29/10
    '
    '***************************************************
    Dim Tmp As Integer

    With UserList(Userindex)
        Tmp = ObjData(ObjIndex).WeaponRazaEnanaAnim
            
        If Tmp > 0 Then
            If .raza = eRaza.Enano Or .raza = eRaza.Gnomo Then
                GetWeaponAnim = Tmp
                Exit Function

            End If

        End If
        
        GetWeaponAnim = ObjData(ObjIndex).WeaponAnim

    End With

End Function


Public Sub EraseUserChar(ByVal Userindex As Integer, ByVal IsAdminInvisible As Boolean)
    '*************************************************
    'Author: Unknown
    'Last modified: 08/01/2009
    '08/01/2009: ZaMa - No se borra el char de un admin invisible en todos los clientes excepto en su mismo cliente.
    '*************************************************

    On Error GoTo ErrorHandler
    
    With UserList(Userindex)
        CharList(.Char.CharIndex) = 0
        
        If .Char.CharIndex = LastChar Then

            Do Until CharList(LastChar) > 0
                LastChar = LastChar - 1

                If LastChar <= 1 Then Exit Do
            Loop

        End If
        
        ' Si esta invisible, solo el sabe de su propia existencia, es innecesario borrarlo en los demas clientes
        If IsAdminInvisible Then
            Call UserList(Userindex).outgoingData.WriteASCIIStringFixed(PrepareMessageCharacterRemove(.Char.CharIndex))
        Else
            'Le mandamos el mensaje para que borre el personaje a los clientes que esten cerca
            Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessageCharacterRemove(.Char.CharIndex))

        End If
        
        Call QuitarUser(Userindex, .Pos.map)
        
        MapData(.Pos.map, .Pos.x, .Pos.y).Userindex = 0
        .Char.CharIndex = 0

    End With
    
    NumChars = NumChars - 1
    Exit Sub
    
ErrorHandler:
    
    Dim UserName  As String

    Dim CharIndex As Integer
    
    If Userindex > 0 Then
        UserName = UserList(Userindex).Name
        CharIndex = UserList(Userindex).Char.CharIndex

    End If

    Call LogError("Error en EraseUserchar " & Err.Number & ": " & Err.description & ". User: " & UserName & "(UI: " & Userindex & " - CI: " & CharIndex & ")")

End Sub

Public Sub RefreshCharStatus(ByVal Userindex As Integer)

    '*************************************************
    'Author: Tararira
    'Last modified: 04/07/2009
    'Refreshes the status and tag of UserIndex.
    '04/07/2009: ZaMa - Ahora mantenes la fragata fantasmal si estas muerto.
    '*************************************************
    Dim ClanTag   As String

    Dim NickColor As Byte

    Dim NuevaA    As Boolean

    Dim GI        As Integer

    Dim tStr      As String
    
    With UserList(Userindex)

        If .GuildIndex > 0 Then
            ClanTag = modGuilds.GuildName(.GuildIndex)
            ClanTag = " <" & ClanTag & ">"

        End If
        
        NickColor = GetNickColor(Userindex)
        
        If .showName Then
            Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessageUpdateTagAndStatus(Userindex, NickColor, .Name & ClanTag))
        Else
            Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessageUpdateTagAndStatus(Userindex, NickColor, vbNullString))
        End If
        
        'Si esta navengando, se cambia la barca.
        If .flags.Navegando Then
            If .flags.Muerto = 1 Then
                .Char.body = iFragataFantasmal
            Else
                Call ToggleBoatBody(Userindex)
            End If
            
            Call ChangeUserChar(Userindex, .Char.body, .Char.Head, .Char.heading, .Char.WeaponAnim, .Char.ShieldAnim, .Char.CascoAnim)

        End If
        
        'Cambio de alineacion
        GI = .GuildIndex
 
        If GI > 0 Then
            NuevaA = False
            'URGENCIA CLANES
            If NuevaA Then
               Call SendData(SendTarget.ToGuildMembers, GI, PrepareMessageConsoleMsg("El clan ha pasado a tener alineacion " & modGuilds.GuildAlignment(GI) & "!", FontTypeNames.FONTTYPE_GUILD))
                tStr = modGuilds.GuildName(GI)
                Call LogClanes("El clan " & tStr & " cambio de alineacion!")
            End If
            
        End If

    End With

End Sub

Public Function GetNickColor(ByVal Userindex As Integer) As Byte
    '*************************************************
    'Author: ZaMa
    'Last modified: 15/01/2010
    '
    '*************************************************
    
    With UserList(Userindex)
        
        Select Case .Faccion.Bando
            Case eFaccion.Armada
                If .Faccion.Jerarquia = 0 Then
                    GetNickColor = eNickColor.ieCiudadano
                Else
                    GetNickColor = eNickColor.ieArmada
                End If
            
            Case eFaccion.Legion
                If .Faccion.Jerarquia = 0 Then
                    GetNickColor = eNickColor.ieCriminal
                Else
                    GetNickColor = eNickColor.ieLegion
                End If
            
            Case eFaccion.Neutral
                GetNickColor = eNickColor.ieNeutral
                
        End Select

    End With
    
End Function

Public Sub MakeUserChar(ByVal toMap As Boolean, _
                        ByVal sndIndex As Integer, _
                        ByVal Userindex As Integer, _
                        ByVal map As Integer, _
                        ByVal x As Integer, _
                        ByVal y As Integer, _
                        Optional ButIndex As Boolean = False)
    '*************************************************
    'Author: Unknown
    'Last modified: 15/01/2010
    '23/07/2009: Budi - Ahora se envia el nick
    '15/01/2010: ZaMa - Ahora se envia el color del nick.
    '*************************************************

    On Error GoTo errHandler

    Dim CharIndex  As Integer

    Dim ClanTag    As String

    Dim NickColor  As Byte

    Dim UserName   As String

    Dim Privileges As Byte
    
    With UserList(Userindex)
    
        If InMapBounds(map, x, y) Then

            'If needed make a new character in list
            If .Char.CharIndex = 0 Then
                CharIndex = NextOpenCharIndex
                .Char.CharIndex = CharIndex
                CharList(CharIndex) = Userindex

            End If
            
            'Place character on map if needed
            If toMap Then MapData(map, x, y).Userindex = Userindex
            
            'Send make character command to clients
            If Not toMap Then
                If .GuildIndex > 0 Then
                    ClanTag = modGuilds.GuildName(.GuildIndex)
                End If
                
                NickColor = GetNickColor(Userindex)
                Privileges = .flags.Privilegios
                
                'Preparo el nick
                If .showName Then
                    UserName = .Name
                    
                    If .flags.EnConsulta Then
                        UserName = UserName & " " & TAG_CONSULT_MODE
                    Else

                        If UserList(sndIndex).flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.RoleMaster) Then
                            If LenB(ClanTag) <> 0 Then UserName = UserName & " <" & ClanTag & ">"
                        Else

                            If (.flags.invisible Or .flags.Oculto) And (Not .flags.AdminInvisible = 1) And .flags.Navegando = 0 Then
                                UserName = UserName & " " & TAG_USER_INVISIBLE
                            Else

                                If LenB(ClanTag) <> 0 Then UserName = UserName & " <" & ClanTag & ">"

                            End If

                        End If

                    End If

                End If
            
                Call WriteCharacterCreate(sndIndex, .Char.body, .Char.Head, .Char.heading, .Char.CharIndex, x, y, .Char.WeaponAnim, .Char.ShieldAnim, .Char.FX, 999, .Char.CascoAnim, UserName, NickColor, Privileges)
            Else
                'Hide the name and clan - set privs as normal user
                Call AgregarUser(Userindex, .Pos.map, ButIndex)

            End If

        End If

    End With

    Exit Sub

errHandler:
    LogError ("MakeUserChar: num: " & Err.Number & " desc: " & Err.description)
    'Resume Next
    Call CloseSocket(Userindex)

End Sub

''
' Checks if the user gets the next level.
'
' @param UserIndex Specifies reference to user

Public Sub CheckUserLevel(ByVal Userindex As Integer, Optional ByVal PrintInConsole As Boolean = True)

    '*************************************************
    'Author: Unknown
    'Last modified: 06/09/2019
    'Chequea que el usuario no halla alcanzado el siguiente nivel,
    'de lo contrario le da la vida, mana, etc, correspodiente.
    '07/08/2006 Integer - Modificacion de los valores
    '01/10/2007 Tavo - Corregido el BUG de STAT_MAXELV
    '24/01/2007 Pablo (ToxicWaste) - Agrego modificaciones en ELU al subir de nivel.
    '24/01/2007 Pablo (ToxicWaste) - Agrego modificaciones de la subida de mana de los magos por lvl.
    '13/03/2007 Pablo (ToxicWaste) - Agrego diferencias entre el 18 y el 19 en Constitucion.
    '09/01/2008 Pablo (ToxicWaste) - Ahora el incremento de vida por Consitucion se controla desde Balance.dat
    '12/09/2008 Marco Vanotti (Marco) - Ahora si se llega a nivel 25 y esta en un clan, se lo expulsa para no sumar antifaccion
    '02/03/2009 ZaMa - Arreglada la validacion de expulsion para miembros de clanes faccionarios que llegan a 25.
    '11/19/2009 Pato - Modifico la nueva formula de mana ganada para el bandido y se la limito a 499
    '02/04/2010: ZaMa - Modifico la ganancia de hit por nivel del ladron.
    '08/04/2011: Amraphen - Arreglada la distribucion de probabilidades para la vida en el caso de promedio entero.
    '06/09/2019: Jopi - Guardado de usuario al pasar de nivel.
    '*************************************************
    Dim Pts              As Integer
    Dim AumentoHIT       As Integer
    Dim AumentoMANA      As Integer
    Dim AumentoSTA       As Integer
    Dim AumentoHP        As Integer
    Dim WasNewbie        As Boolean
    Dim Promedio         As Double
    Dim aux              As Integer
    Dim DistVida(1 To 5) As Integer
    Dim GI               As Integer 'Guild Index
    
    On Error GoTo errHandler
    
    WasNewbie = EsNewbie(Userindex)
    
    With UserList(Userindex)

        Do While .Stats.Exp >= .Stats.ELU
            
            'Checkea si alcanzo el maximo nivel
            If .Stats.ELV >= STAT_MAXELV Then
                .Stats.Exp = 0
                .Stats.ELU = 0
                Exit Sub

            End If
            
            'Store it!
            Call Statistics.UserLevelUp(Userindex)
            
            If PrintInConsole Then
                Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessagePlayWave(SND_NIVEL, .Pos.x, .Pos.y))
                Call WriteConsoleMsg(Userindex, "Has subido de nivel!", FontTypeNames.FONTTYPE_INFO)
            End If
            
            If .Stats.ELV = 1 Then
                Pts = 10
            Else
                'For multiple levels being rised at once
                Pts = Pts + 5

            End If
            
            .Stats.ELV = .Stats.ELV + 1
            
            .Stats.Exp = .Stats.Exp - .Stats.ELU
            
            'Nueva subida de exp x lvl. Pablo (ToxicWaste)
            If .Stats.ELV < 15 Then
                .Stats.ELU = .Stats.ELU * 1.4
            ElseIf .Stats.ELV < 21 Then
                .Stats.ELU = .Stats.ELU * 1.35
            ElseIf .Stats.ELV < 26 Then
                .Stats.ELU = .Stats.ELU * 1.3
            ElseIf .Stats.ELV < 35 Then
                .Stats.ELU = .Stats.ELU * 1.2
            ElseIf .Stats.ELV < 40 Then
                .Stats.ELU = .Stats.ELU * 1.3
            Else
                .Stats.ELU = .Stats.ELU * 1.375

            End If
            
            'Calculo subida de vida
            Promedio = ModVida(.Clase) - (21 - .Stats.UserAtributos(eAtributos.Constitucion)) * 0.5
            aux = RandomNumber(0, 100)
            
            If Promedio - Int(Promedio) = 0.5 Then
                'Es promedio semientero
                DistVida(1) = DistribucionSemienteraVida(1)
                DistVida(2) = DistVida(1) + DistribucionSemienteraVida(2)
                DistVida(3) = DistVida(2) + DistribucionSemienteraVida(3)
                DistVida(4) = DistVida(3) + DistribucionSemienteraVida(4)
                
                If aux <= DistVida(1) Then
                    AumentoHP = Promedio + 1.5
                ElseIf aux <= DistVida(2) Then
                    AumentoHP = Promedio + 0.5
                ElseIf aux <= DistVida(3) Then
                    AumentoHP = Promedio - 0.5
                Else
                    AumentoHP = Promedio - 1.5

                End If

            Else
                'Es promedio entero
                DistVida(1) = DistribucionEnteraVida(1)
                DistVida(2) = DistVida(1) + DistribucionEnteraVida(2)
                DistVida(3) = DistVida(2) + DistribucionEnteraVida(3)
                DistVida(4) = DistVida(3) + DistribucionEnteraVida(4)
                DistVida(5) = DistVida(4) + DistribucionEnteraVida(5)
                
                If aux <= DistVida(1) Then
                    AumentoHP = Promedio + 2
                ElseIf aux <= DistVida(2) Then
                    AumentoHP = Promedio + 1
                ElseIf aux <= DistVida(3) Then
                    AumentoHP = Promedio
                ElseIf aux <= DistVida(4) Then
                    AumentoHP = Promedio - 1
                Else
                    AumentoHP = Promedio - 2

                End If
                
            End If
        
            Select Case .Clase

                Case eClass.Warrior
                    AumentoHIT = IIf(.Stats.ELV > 35, 2, 3)
                    AumentoSTA = AumentoSTDef
                
                Case eClass.Hunter
                    AumentoHIT = IIf(.Stats.ELV > 35, 2, 3)
                    AumentoSTA = AumentoSTDef
                
                Case eClass.Pirat
                    AumentoHIT = 3
                    AumentoSTA = AumentoSTDef
                
                Case eClass.Paladin
                    AumentoHIT = IIf(.Stats.ELV > 35, 1, 3)
                    AumentoMANA = .Stats.UserAtributos(eAtributos.Inteligencia)
                    AumentoSTA = AumentoSTDef
                
                Case eClass.Thief
                    AumentoHIT = 2
                    AumentoSTA = AumentoSTLadron
                
                Case eClass.Mage
                    AumentoHIT = 1
                    AumentoMANA = 2.8 * .Stats.UserAtributos(eAtributos.Inteligencia)
                    AumentoSTA = AumentoSTMago
                
                Case eClass.Worker
                    AumentoHIT = 2
                    AumentoSTA = AumentoSTTrabajador
                
                Case eClass.Cleric
                    AumentoHIT = 2
                    AumentoMANA = 2 * .Stats.UserAtributos(eAtributos.Inteligencia)
                    AumentoSTA = AumentoSTDef
                
                Case eClass.Druid
                    AumentoHIT = 2
                    AumentoMANA = 2 * .Stats.UserAtributos(eAtributos.Inteligencia)
                    AumentoSTA = AumentoSTDef
                
                Case eClass.Assasin
                    AumentoHIT = IIf(.Stats.ELV > 35, 1, 3)
                    AumentoMANA = .Stats.UserAtributos(eAtributos.Inteligencia)
                    AumentoSTA = AumentoSTDef
                
                Case eClass.Bard
                    AumentoHIT = 2
                    AumentoMANA = 2 * .Stats.UserAtributos(eAtributos.Inteligencia)
                    AumentoSTA = AumentoSTDef
                    
                Case eClass.Bandit
                    AumentoHIT = IIf(.Stats.ELV > 35, 1, 3)
                    AumentoMANA = .Stats.UserAtributos(eAtributos.Inteligencia) / 3 * 2
                    AumentoSTA = AumentoStBandido
                
                Case Else
                    AumentoHIT = 2
                    AumentoSTA = AumentoSTDef

            End Select
            
            'Actualizamos HitPoints
            .Stats.MaxHp = .Stats.MaxHp + AumentoHP

            If .Stats.MaxHp > STAT_MAXHP Then .Stats.MaxHp = STAT_MAXHP
            
            'Actualizamos Stamina
            .Stats.MaxSta = .Stats.MaxSta + AumentoSTA

            If .Stats.MaxSta > STAT_MAXSTA Then .Stats.MaxSta = STAT_MAXSTA
            
            'Actualizamos Mana
            .Stats.MaxMAN = .Stats.MaxMAN + AumentoMANA

            If .Stats.MaxMAN > STAT_MAXMAN Then .Stats.MaxMAN = STAT_MAXMAN
            
            'Actualizamos Golpe Maximo
            .Stats.MaxHIT = .Stats.MaxHIT + AumentoHIT

            If .Stats.ELV < 36 Then
                If .Stats.MaxHIT > STAT_MAXHIT_UNDER36 Then .Stats.MaxHIT = STAT_MAXHIT_UNDER36
            Else

                If .Stats.MaxHIT > STAT_MAXHIT_OVER36 Then .Stats.MaxHIT = STAT_MAXHIT_OVER36

            End If
            
            'Actualizamos Golpe Minimo
            .Stats.MinHIT = .Stats.MinHIT + AumentoHIT

            If .Stats.ELV < 36 Then
                If .Stats.MinHIT > STAT_MAXHIT_UNDER36 Then .Stats.MinHIT = STAT_MAXHIT_UNDER36
            Else

                If .Stats.MinHIT > STAT_MAXHIT_OVER36 Then .Stats.MinHIT = STAT_MAXHIT_OVER36

            End If
            
            
            'Notificamos al user
            If PrintInConsole Then
                If AumentoHP > 0 Then
                    Call WriteConsoleMsg(Userindex, "Has ganado " & AumentoHP & " puntos de vida.", FontTypeNames.FONTTYPE_INFO)

                End If

                If AumentoSTA > 0 Then
                    Call WriteConsoleMsg(Userindex, "Has ganado " & AumentoSTA & " puntos de energia.", FontTypeNames.FONTTYPE_INFO)

                End If

                If AumentoMANA > 0 Then
                    Call WriteConsoleMsg(Userindex, "Has ganado " & AumentoMANA & " puntos de mana.", FontTypeNames.FONTTYPE_INFO)

                End If

                If AumentoHIT > 0 Then
                    Call WriteConsoleMsg(Userindex, "Tu golpe maximo aumento en " & AumentoHIT & " puntos.", FontTypeNames.FONTTYPE_INFO)
                    Call WriteConsoleMsg(Userindex, "Tu golpe minimo aumento en " & AumentoHIT & " puntos.", FontTypeNames.FONTTYPE_INFO)

                End If
            End If
            
            Call LogDesarrollo(.Name & " paso a nivel " & .Stats.ELV & " gano HP: " & AumentoHP)
            
            .Stats.MinHp = .Stats.MaxHp

            'If user is in a party, we modify the variable p_sumaniveleselevados
            Call mdParty.ActualizarSumaNivelesElevados(Userindex)
            'If user reaches lvl 25 and he is in a guild, we check the guild's alignment and expulses the user if guild has factionary alignment
        
            If .Stats.ELV = 25 Then
                GI = .GuildIndex

                If GI > 0 Then
                    If modGuilds.GuildAlignment(GI) = "Del Mal" Or modGuilds.GuildAlignment(GI) = "Real" Then
                        'We get here, so guild has factionary alignment, we have to expulse the user
                        Call modGuilds.m_EcharMiembroDeClan(-1, .Name)
                        
                        If PrintInConsole Then
                            Call SendData(SendTarget.ToGuildMembers, GI, PrepareMessageConsoleMsg(.Name & " deja el clan.", FontTypeNames.FONTTYPE_GUILD))
                            Call WriteConsoleMsg(Userindex, "Ya tienes la madurez suficiente como para decidir bajo que estandarte pelearas! Por esta razon, hasta tanto no te enlistes en la faccion bajo la cual tu clan esta alineado, estaras excluido del mismo.", FontTypeNames.FONTTYPE_GUILD)
                        End If

                    End If

                End If

            End If

        Loop
        
        'If it ceased to be a newbie, remove newbie items and get char away from newbie dungeon
        If Not EsNewbie(Userindex) And WasNewbie Then
            Call QuitarNewbieObj(Userindex)

            If MapInfo(.Pos.map).Restringir = eRestrict.restrict_newbie Then
                Call WarpUserChar(Userindex, 1, 50, 50, True)

                If PrintInConsole Then
                    Call WriteConsoleMsg(Userindex, "Debes abandonar el Dungeon Newbie.", FontTypeNames.FONTTYPE_INFO)
                End If

            End If

        End If
        
        'Send all gained skill points at once (if any)
        If Pts > 0 Then
            Call WriteLevelUp(Userindex, Pts)
            
            .Stats.SkillPts = .Stats.SkillPts + Pts
            If PrintInConsole Then
                Call WriteConsoleMsg(Userindex, "Has ganado un total de " & Pts & " skillpoints.", FontTypeNames.FONTTYPE_INFO)
            End If

        End If
        
    End With
    
    'Guardamos los datos del usuario.
    Call SaveUser(Userindex, True)
    
    Call WriteUpdateUserStats(Userindex)
    
    Exit Sub

errHandler:
    Call LogError("Error en la subrutina CheckUserLevel - Error : " & Err.Number & " - Description : " & Err.description)

End Sub

Public Function PuedeAtravesarAgua(ByVal Userindex As Integer) As Boolean
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    PuedeAtravesarAgua = UserList(Userindex).flags.Navegando = 1 Or UserList(Userindex).flags.Vuela = 1

End Function

Sub MoveUserChar(ByVal Userindex As Integer, ByVal nHeading As eHeading)

    '*************************************************
    'Author: Unknown
    'Last Modification: 06/04/2020
    'Moves the char, sending the message to everyone in range.
    '30/03/2009: ZaMa - Now it's legal to move where a casper is, changing its pos to where the moving char was.
    '28/05/2009: ZaMa - When you are moved out of an Arena, the resurrection safe is activated.
    '13/07/2009: ZaMa - Now all the clients don't know when an invisible admin moves, they force the admin to move.
    '13/07/2009: ZaMa - Invisible admins aren't allowed to force dead characater to move
    '06/04/2020: FrankoH298 - Ahora no se puede entrar a las casas montado.
    '*************************************************
    Dim nPos          As WorldPos

    Dim sailing       As Boolean

    Dim CasperIndex   As Integer

    Dim CasperHeading As eHeading

    Dim isAdminInvi   As Boolean
    
    sailing = PuedeAtravesarAgua(Userindex)
    nPos = UserList(Userindex).Pos
    Call HeadtoPos(nHeading, nPos)
        
    isAdminInvi = (UserList(Userindex).flags.AdminInvisible = 1)
    
    If MoveToLegalPos(UserList(Userindex).Pos.map, nPos.x, nPos.y, sailing, Not sailing) Then

        'No se puede caminar con monturas en casas, bajo techo o dungeons
        If UserList(Userindex).flags.Equitando And _
           (MapData(UserList(Userindex).Pos.map, nPos.x, nPos.y).trigger = eTrigger.CASA Or _
           MapData(UserList(Userindex).Pos.map, nPos.x, nPos.y).trigger = eTrigger.BAJOTECHO Or _
           MapInfo(UserList(Userindex).Pos.map).Zona = Dungeon) Then _

            Exit Sub
        End If

        'si no estoy solo en el mapa...
        If MapInfo(UserList(Userindex).Pos.map).numUsers > 1 Then
               
            CasperIndex = MapData(UserList(Userindex).Pos.map, nPos.x, nPos.y).Userindex

            'Si hay un usuario, y paso la validacion, entonces es un casper
            If CasperIndex > 0 Then

                ' Los admins invisibles no pueden patear caspers
                If Not isAdminInvi Then
                    
                    If TriggerZonaPelea(Userindex, CasperIndex) = TRIGGER6_PROHIBE Then
                        If UserList(CasperIndex).flags.SeguroResu = False Then
                            UserList(CasperIndex).flags.SeguroResu = True
                            Call WriteMultiMessage(CasperIndex, eMessages.ResuscitationSafeOn)

                        End If

                    End If
    
                    With UserList(CasperIndex)
                        CasperHeading = InvertHeading(nHeading)
                        Call HeadtoPos(CasperHeading, .Pos)
                    
                        ' Si es un admin invisible, no se avisa a los demas clientes
                        If Not .flags.AdminInvisible = 1 Then Call SendData(SendTarget.ToPCAreaButIndex, CasperIndex, PrepareMessageCharacterMove(.Char.CharIndex, .Pos.x, .Pos.y))
                        
                        Call WriteForceCharMove(CasperIndex, CasperHeading)
                            
                        'Update map and char
                        .Char.heading = CasperHeading
                        MapData(.Pos.map, .Pos.x, .Pos.y).Userindex = CasperIndex

                    End With
                
                    'Actualizamos las areas de ser necesario
                    Call Areas.CheckUpdateNeededUser(CasperIndex, CasperHeading)

                End If

            End If
            
            ' Si es un admin invisible, no se avisa a los demas clientes
            If Not isAdminInvi Then Call SendData(SendTarget.ToPCAreaButIndex, Userindex, PrepareMessageCharacterMove(UserList(Userindex).Char.CharIndex, nPos.x, nPos.y))
            
        End If
        
        ' Los admins invisibles no pueden patear caspers
        If Not (isAdminInvi And (CasperIndex <> 0)) Then

            Dim oldUserIndex As Integer
            
            With UserList(Userindex)
                oldUserIndex = MapData(.Pos.map, .Pos.x, .Pos.y).Userindex
                
                ' Si no hay intercambio de pos con nadie
                If oldUserIndex = Userindex Then
                    MapData(.Pos.map, .Pos.x, .Pos.y).Userindex = 0

                End If
                
                .Pos = nPos
                .Char.heading = nHeading
                MapData(.Pos.map, .Pos.x, .Pos.y).Userindex = Userindex
                
                If HaySacerdote(Userindex) Then Call AccionParaSacerdote(Userindex)
                
                Call DoTileEvents(Userindex, .Pos.map, .Pos.x, .Pos.y)

            End With
            
            'Actualizamos las areas de ser necesario
            Call Areas.CheckUpdateNeededUser(Userindex, nHeading)
        Else
            Call WritePosUpdate(Userindex)

        End If

    Else
        Call WritePosUpdate(Userindex)

    End If
    
    If UserList(Userindex).Counters.Trabajando Then UserList(Userindex).Counters.Trabajando = UserList(Userindex).Counters.Trabajando - 1

    If UserList(Userindex).Counters.Ocultando Then UserList(Userindex).Counters.Ocultando = UserList(Userindex).Counters.Ocultando - 1

End Sub

Public Function InvertHeading(ByVal nHeading As eHeading) As eHeading

    '*************************************************
    'Author: ZaMa
    'Last modified: 30/03/2009
    'Returns the heading opposite to the one passed by val.
    '*************************************************
    Select Case nHeading

        Case eHeading.EAST
            InvertHeading = WEST

        Case eHeading.WEST
            InvertHeading = EAST

        Case eHeading.SOUTH
            InvertHeading = NORTH

        Case eHeading.NORTH
            InvertHeading = SOUTH

    End Select

End Function

Sub ChangeUserInv(ByVal Userindex As Integer, ByVal Slot As Byte, ByRef Object As UserObj)
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    UserList(Userindex).Invent.Object(Slot) = Object
    Call WriteChangeInventorySlot(Userindex, Slot)

End Sub

Function NextOpenCharIndex() As Integer
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    Dim LoopC As Long
    
    For LoopC = 1 To MAXCHARS

        If CharList(LoopC) = 0 Then
            NextOpenCharIndex = LoopC
            NumChars = NumChars + 1
            
            If LoopC > LastChar Then LastChar = LoopC
            
            Exit Function

        End If

    Next LoopC

End Function

Function NextOpenUser() As Integer
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    Dim LoopC As Long
    
    For LoopC = 1 To MaxUsers + 1

        If LoopC > MaxUsers Then Exit For
        If (UserList(LoopC).ConnID = -1 And UserList(LoopC).flags.UserLogged = False) Then Exit For
    Next LoopC
    
    NextOpenUser = LoopC

End Function

Public Sub LiberarSlot(ByVal Userindex As Integer)
    '***************************************************
    'Author: Torres Patricio (Pato)
    'Last Modification: 01/10/2012
    '
    '***************************************************

    With UserList(Userindex)
        .ConnID = -1
        .ConnIDValida = False

    End With

    If Userindex = LastUser Then

        Do While (LastUser > 0) And (UserList(LastUser).ConnID = -1)
            LastUser = LastUser - 1
            If LastUser = 0 Then Exit Do
        Loop

    End If

End Sub

Public Sub SendUserStatsTxt(ByVal sendIndex As Integer, ByVal Userindex As Integer)
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    Dim GuildI As Integer
    
    With UserList(Userindex)
        Call WriteConsoleMsg(sendIndex, "Estadisticas de: " & .Name, FontTypeNames.FONTTYPE_INFO)
        Call WriteConsoleMsg(sendIndex, "Nivel: " & .Stats.ELV & "  EXP: " & .Stats.Exp & "/" & .Stats.ELU, FontTypeNames.FONTTYPE_INFO)
        Call WriteConsoleMsg(sendIndex, "Salud: " & .Stats.MinHp & "/" & .Stats.MaxHp & "  Mana: " & .Stats.MinMAN & "/" & .Stats.MaxMAN & "  Energia: " & .Stats.MinSta & "/" & .Stats.MaxSta, FontTypeNames.FONTTYPE_INFO)
        
        If .Invent.WeaponEqpObjIndex > 0 Then
            Call WriteConsoleMsg(sendIndex, "Menor Golpe/Mayor Golpe: " & .Stats.MinHIT & "/" & .Stats.MaxHIT & " (" & ObjData(.Invent.WeaponEqpObjIndex).MinHIT & "/" & ObjData(.Invent.WeaponEqpObjIndex).MaxHIT & ")", FontTypeNames.FONTTYPE_INFO)
        Else
            Call WriteConsoleMsg(sendIndex, "Menor Golpe/Mayor Golpe: " & .Stats.MinHIT & "/" & .Stats.MaxHIT, FontTypeNames.FONTTYPE_INFO)

        End If
        
        If .Invent.ArmourEqpObjIndex > 0 Then
            If .Invent.EscudoEqpObjIndex > 0 Then
                Call WriteConsoleMsg(sendIndex, "(CUERPO) Min Def/Max Def: " & ObjData(.Invent.ArmourEqpObjIndex).MinDef + ObjData(.Invent.EscudoEqpObjIndex).MinDef & "/" & ObjData(.Invent.ArmourEqpObjIndex).MaxDef + ObjData(.Invent.EscudoEqpObjIndex).MaxDef, FontTypeNames.FONTTYPE_INFO)
            Else
                Call WriteConsoleMsg(sendIndex, "(CUERPO) Min Def/Max Def: " & ObjData(.Invent.ArmourEqpObjIndex).MinDef & "/" & ObjData(.Invent.ArmourEqpObjIndex).MaxDef, FontTypeNames.FONTTYPE_INFO)

            End If

        Else
            Call WriteConsoleMsg(sendIndex, "(CUERPO) Min Def/Max Def: 0", FontTypeNames.FONTTYPE_INFO)

        End If
        
        If .Invent.CascoEqpObjIndex > 0 Then
            Call WriteConsoleMsg(sendIndex, "(CABEZA) Min Def/Max Def: " & ObjData(.Invent.CascoEqpObjIndex).MinDef & "/" & ObjData(.Invent.CascoEqpObjIndex).MaxDef, FontTypeNames.FONTTYPE_INFO)
        Else
            Call WriteConsoleMsg(sendIndex, "(CABEZA) Min Def/Max Def: 0", FontTypeNames.FONTTYPE_INFO)

        End If
        
        GuildI = .GuildIndex

        If GuildI > 0 Then
            Call WriteConsoleMsg(sendIndex, "Clan: " & modGuilds.GuildName(GuildI), FontTypeNames.FONTTYPE_INFO)

            If UCase$(modGuilds.GuildLeader(GuildI)) = UCase$(.Name) Then
                Call WriteConsoleMsg(sendIndex, "Status: Lider", FontTypeNames.FONTTYPE_INFO)

            End If

            'guildpts no tienen objeto
        End If
        
        #If ConUpTime Then

            Dim TempDate As Date

            Dim TempSecs As Long

            Dim TempStr  As String

            TempDate = Now - .LogOnTime
            TempSecs = (.UpTime + (Abs(Day(TempDate) - 30) * 24 * 3600) + (Hour(TempDate) * 3600) + (Minute(TempDate) * 60) + Second(TempDate))
            TempStr = (TempSecs \ 86400) & " Dias, " & ((TempSecs Mod 86400) \ 3600) & " Horas, " & ((TempSecs Mod 86400) Mod 3600) \ 60 & " Minutos, " & (((TempSecs Mod 86400) Mod 3600) Mod 60) & " Segundos."
            Call WriteConsoleMsg(sendIndex, "Logeado hace: " & Hour(TempDate) & ":" & Minute(TempDate) & ":" & Second(TempDate), FontTypeNames.FONTTYPE_INFO)
            Call WriteConsoleMsg(sendIndex, "Total: " & TempStr, FontTypeNames.FONTTYPE_INFO)
        #End If

        If .flags.Traveling = 1 Then
            Call WriteConsoleMsg(sendIndex, "Tiempo restante para llegar a tu hogar: " & GetHomeArrivalTime(Userindex) & " segundos.", FontTypeNames.FONTTYPE_INFO)

        End If
        
        Call WriteConsoleMsg(sendIndex, "Oro: " & .Stats.Gld & "  Posicion: " & .Pos.x & "," & .Pos.y & " en mapa " & .Pos.map, FontTypeNames.FONTTYPE_INFO)
        Call WriteConsoleMsg(sendIndex, "Dados: " & .Stats.UserAtributos(eAtributos.Fuerza) & ", " & .Stats.UserAtributos(eAtributos.Agilidad) & ", " & .Stats.UserAtributos(eAtributos.Inteligencia) & ", " & .Stats.UserAtributos(eAtributos.Carisma) & ", " & .Stats.UserAtributos(eAtributos.Constitucion), FontTypeNames.FONTTYPE_INFO)

    End With

End Sub

Sub SendUserMiniStatsTxt(ByVal sendIndex As Integer, ByVal Userindex As Integer)

    '*************************************************
    'Author: Unknown
    'Last modified: 23/01/2007
    'Shows the users Stats when the user is online.
    '23/01/2007 Pablo (ToxicWaste) - Agrego de funciones y mejora de distribucion de parametros.
    '*************************************************
    With UserList(Userindex)
        Call WriteConsoleMsg(sendIndex, "Pj: " & .Name, FontTypeNames.FONTTYPE_INFO)
        Call WriteConsoleMsg(sendIndex, "Neutrales matados: " & .Faccion.NeutralesMatados & " Ciudadanos matados: " & .Faccion.CiudadanosMatados & " Criminales matados: " & .Faccion.CriminalesMatados & " usuarios matados: " & .Stats.UsuariosMatados, FontTypeNames.FONTTYPE_INFO)
        Call WriteConsoleMsg(sendIndex, "NPCs muertos: " & .Stats.NPCsMuertos, FontTypeNames.FONTTYPE_INFO)
        Call WriteConsoleMsg(sendIndex, "Clase: " & ListaClases(.Clase), FontTypeNames.FONTTYPE_INFO)
        Call WriteConsoleMsg(sendIndex, "Pena: " & (.Counters.Pena / 40), FontTypeNames.FONTTYPE_INFO)
        
        If .Faccion.Bando = eFaccion.Armada Then
            Call WriteConsoleMsg(sendIndex, "Faccion: Alianza", FontTypeNames.FONTTYPE_INFO)
            
        ElseIf .Faccion.Bando = eFaccion.Neutral Then
            Call WriteConsoleMsg(sendIndex, "Faccion: Neutral", FontTypeNames.FONTTYPE_INFO)
            
        
        ElseIf .Faccion.Bando = eFaccion.Legion Then
            Call WriteConsoleMsg(sendIndex, "Faccion: Legion", FontTypeNames.FONTTYPE_INFO)
            
        End If

        If .GuildIndex > 0 Then
            Call WriteConsoleMsg(sendIndex, "Clan: " & GuildName(.GuildIndex), FontTypeNames.FONTTYPE_INFO)

        End If

    End With

End Sub

Sub SendUserInvTxt(ByVal sendIndex As Integer, ByVal Userindex As Integer)
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    On Error Resume Next

    Dim j As Long
    
    With UserList(Userindex)
        Call WriteConsoleMsg(sendIndex, .Name, FontTypeNames.FONTTYPE_INFO)
        Call WriteConsoleMsg(sendIndex, "Tiene " & .Invent.NroItems & " objetos.", FontTypeNames.FONTTYPE_INFO)
        
        For j = 1 To .CurrentInventorySlots

            If .Invent.Object(j).ObjIndex > 0 Then
                Call WriteConsoleMsg(sendIndex, "Objeto " & j & " " & ObjData(.Invent.Object(j).ObjIndex).Name & " Cantidad:" & .Invent.Object(j).amount, FontTypeNames.FONTTYPE_INFO)

            End If

        Next j

    End With

End Sub

Sub SendUserSkillsTxt(ByVal sendIndex As Integer, ByVal Userindex As Integer)
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    On Error Resume Next

    Dim j As Integer
    
    Call WriteConsoleMsg(sendIndex, UserList(Userindex).Name, FontTypeNames.FONTTYPE_INFO)
    
    For j = 1 To NUMSKILLS
        Call WriteConsoleMsg(sendIndex, SkillsNames(j) & " = " & UserList(Userindex).Stats.UserSkills(j), FontTypeNames.FONTTYPE_INFO)
    Next j
    
    Call WriteConsoleMsg(sendIndex, "SkillLibres:" & UserList(Userindex).Stats.SkillPts, FontTypeNames.FONTTYPE_INFO)

End Sub



Sub NPCAtacado(ByVal NpcIndex As Integer, ByVal Userindex As Integer)

    '**********************************************
    'Author: Unknown
    'Last Modification: 02/04/2010
    '24/01/2007 -> Pablo (ToxicWaste): Agrego para que se actualize el tag si corresponde.
    '24/07/2007 -> Pablo (ToxicWaste): Guardar primero que ataca NPC y el que atacas ahora.
    '06/28/2008 -> NicoNZ: Los elementales al atacarlos por su amo no se paran mas al lado de el sin hacer nada.
    '02/04/2010: ZaMa: Un ciuda no se vuelve mas criminal al atacar un npc no hostil.
    '**********************************************
    Dim EraCriminal As Boolean
    
    'Guardamos el usuario que ataco el npc.
    Npclist(NpcIndex).flags.AttackedBy = UserList(Userindex).Name
    
    'Npc que estabas atacando.
    Dim LastNpcHit As Integer

    LastNpcHit = UserList(Userindex).flags.NPCAtacado
    'Guarda el NPC que estas atacando ahora.
    UserList(Userindex).flags.NPCAtacado = NpcIndex
    
    'Revisamos robo de npc.
    'Guarda el primer nick que lo ataca.
    If Npclist(NpcIndex).flags.AttackedFirstBy = vbNullString Then

        'El que le pegabas antes ya no es tuyo
        If LastNpcHit <> 0 Then
            If Npclist(LastNpcHit).flags.AttackedFirstBy = UserList(Userindex).Name Then
                Npclist(LastNpcHit).flags.AttackedFirstBy = vbNullString

            End If

        End If

        Npclist(NpcIndex).flags.AttackedFirstBy = UserList(Userindex).Name
    ElseIf Npclist(NpcIndex).flags.AttackedFirstBy <> UserList(Userindex).Name Then

        'Estas robando NPC
        'El que le pegabas antes ya no es tuyo
        If LastNpcHit <> 0 Then
            If Npclist(LastNpcHit).flags.AttackedFirstBy = UserList(Userindex).Name Then
                Npclist(LastNpcHit).flags.AttackedFirstBy = vbNullString

            End If

        End If

    End If
    
    If Npclist(NpcIndex).MaestroUser > 0 Then
        If Npclist(NpcIndex).MaestroUser <> Userindex Then
            Call AllMascotasAtacanUser(Userindex, Npclist(NpcIndex).MaestroUser)

        End If

    End If
    

        Npclist(NpcIndex).Movement = TipoAI.NPCDEFENSA
        Npclist(NpcIndex).Hostile = 1


   ' End If

End Sub

Public Function PuedeApunalar(ByVal Userindex As Integer) As Boolean
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************
    
    Dim WeaponIndex As Integer
     
    With UserList(Userindex)
        
        WeaponIndex = .Invent.WeaponEqpObjIndex
        
        If WeaponIndex > 0 Then
            If ObjData(WeaponIndex).Apunala = 1 Then
                PuedeApunalar = .Stats.UserSkills(eSkill.Apunalar) >= MIN_APUNALAR Or .Clase = eClass.Assasin

            End If

        End If
        
    End With
    
End Function

Public Function PuedeAcuchillar(ByVal Userindex As Integer) As Boolean
    '***************************************************
    'Author: ZaMa
    'Last Modification: 25/01/2010 (ZaMa)
    '
    '***************************************************
    
    Dim WeaponIndex As Integer
    
    With UserList(Userindex)

        If .Clase = eClass.Pirat Then
        
            WeaponIndex = .Invent.WeaponEqpObjIndex

            If WeaponIndex > 0 Then
                PuedeAcuchillar = (ObjData(WeaponIndex).Acuchilla = 1)

            End If

        End If

    End With
    
End Function

Sub SubirSkill(ByVal Userindex As Integer, _
               ByVal Skill As Integer, _
               ByVal Acerto As Boolean)

    '*************************************************
    'Author: Unknown
    'Last modified: 11/19/2009
    '11/19/2009 Pato - Implement the new system to train the skills.
    '*************************************************
    With UserList(Userindex)

        If .flags.Hambre = 0 And .flags.Sed = 0 Then
            If .Counters.AsignedSkills < 10 Then
                If Not .flags.UltimoMensaje = 7 Then
                    Call WriteConsoleMsg(Userindex, "Para poder entrenar un skill debes asignar los 10 skills iniciales.", FontTypeNames.FONTTYPE_INFO)
                    .flags.UltimoMensaje = 7

                End If
                
                Exit Sub

            End If
                
            With .Stats

                If .UserSkills(Skill) = MAXSKILLPOINTS Then Exit Sub
                
                Dim Lvl As Integer

                Lvl = .ELV
                
                If Lvl > UBound(LevelSkill) Then Lvl = UBound(LevelSkill)
                
                If .UserSkills(Skill) >= LevelSkill(Lvl).LevelValue Then Exit Sub
                
                If Acerto Then
                    .ExpSkills(Skill) = .ExpSkills(Skill) + EXP_ACIERTO_SKILL
                Else
                    .ExpSkills(Skill) = .ExpSkills(Skill) + EXP_FALLO_SKILL

                End If
                
                If .ExpSkills(Skill) >= .EluSkills(Skill) Then
                    .UserSkills(Skill) = .UserSkills(Skill) + 1
                    
                    Call WriteConsoleMsg(Userindex, "Has mejorado tu skill " & SkillsNames(Skill) & " en un punto! Ahora tienes " & .UserSkills(Skill) & " pts.", FontTypeNames.FONTTYPE_INFO)
                    
                    .Exp = .Exp + 50

                    If .Exp > MAXEXP Then .Exp = MAXEXP
                    
                    Call WriteConsoleMsg(Userindex, "Has ganado 50 puntos de experiencia!", FontTypeNames.FONTTYPE_FIGHT)
                    
                    Call WriteUpdateExp(Userindex)
                    Call CheckUserLevel(Userindex)
                    Call CheckEluSkill(Userindex, Skill, False)

                End If

            End With

        End If

    End With

End Sub

''
' Muere un usuario
'
' @param UserIndex  Indice del usuario que muere
'

Public Sub UserDie(ByVal Userindex As Integer, Optional ByVal AttackerIndex As Integer = 0)

    '************************************************
    'Author: Uknown
    'Last Modified: 12/01/2010 (ZaMa)
    '04/15/2008: NicoNZ - Ahora se resetea el counter del invi
    '13/02/2009: ZaMa - Ahora se borran las mascotas cuando moris en agua.
    '27/05/2009: ZaMa - El seguro de resu no se activa si estas en una arena.
    '21/07/2009: Marco - Al morir se desactiva el comercio seguro.
    '16/11/2009: ZaMa - Al morir perdes la criatura que te pertenecia.
    '27/11/2009: Budi - Al morir envia los atributos originales.
    '12/01/2010: ZaMa - Los druidas pierden la inmunidad de ser atacados cuando mueren.
    '************************************************
    On Error GoTo ErrorHandler

    Dim i           As Long

    Dim aN          As Integer
    
    Dim iSoundDeath As Integer
    
    With UserList(Userindex)

        'Sonido
        If .Genero = eGenero.Mujer Then
            If HayAgua(.Pos.map, .Pos.x, .Pos.y) Then
                iSoundDeath = e_SoundIndex.MUERTE_MUJER_AGUA
            Else
                iSoundDeath = e_SoundIndex.MUERTE_MUJER

            End If

        Else

            If HayAgua(.Pos.map, .Pos.x, .Pos.y) Then
                iSoundDeath = e_SoundIndex.MUERTE_HOMBRE_AGUA
            Else
                iSoundDeath = e_SoundIndex.MUERTE_HOMBRE

            End If

        End If
        
        Call ReproducirSonido(SendTarget.ToPCArea, Userindex, iSoundDeath)
        
        'Quitar el dialogo del user muerto
        Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessageRemoveCharDialog(.Char.CharIndex))
        
        .Stats.MinHp = 0
        .Stats.MinSta = 0
        .flags.AtacadoPorUser = 0
        .flags.Envenenado = 0
        .flags.Muerto = 1

        .Counters.Trabajando = 0
        
        ' No se activa en arenas
        If TriggerZonaPelea(Userindex, Userindex) <> TRIGGER6_PERMITE Then
            .flags.SeguroResu = True
            Call WriteMultiMessage(Userindex, eMessages.ResuscitationSafeOn) 'Call WriteResuscitationSafeOn(UserIndex)
        Else
            .flags.SeguroResu = False
            Call WriteMultiMessage(Userindex, eMessages.ResuscitationSafeOff) 'Call WriteResuscitationSafeOff(UserIndex)

        End If
        
        aN = .flags.AtacadoPorNpc

        If aN > 0 Then
            Npclist(aN).Movement = Npclist(aN).flags.OldMovement
            Npclist(aN).Hostile = Npclist(aN).flags.OldHostil
            Npclist(aN).flags.AttackedBy = vbNullString

        End If
        
        aN = .flags.NPCAtacado

        If aN > 0 Then
            If Npclist(aN).flags.AttackedFirstBy = .Name Then
                Npclist(aN).flags.AttackedFirstBy = vbNullString

            End If

        End If

        .flags.AtacadoPorNpc = 0
        .flags.NPCAtacado = 0
        
        Call PerdioNpc(Userindex, False)
        
        '<<<< Equitando >>>>
        If .flags.Equitando = 1 Then
            Call UnmountMontura(Userindex)
            Call WriteEquitandoToggle(Userindex)
            
        End If
        
        '<<<< Paralisis >>>>
        If .flags.Paralizado = 1 Then
            .flags.Paralizado = 0
            Call WriteParalizeOK(Userindex)

        End If
        
        '<<< Estupidez >>>
        If .flags.Estupidez = 1 Then
            .flags.Estupidez = 0
            Call WriteDumbNoMore(Userindex)

        End If
        
        '<<<< Descansando >>>>
        If .flags.Descansar Then
            .flags.Descansar = False
            Call WriteRestOK(Userindex)

        End If
        
        '<<<< Meditando >>>>
        If .flags.Meditando Then
            .flags.Meditando = False
            Call WriteMeditateToggle(Userindex)

        End If
        
        '<<<< Invisible >>>>
        If .flags.invisible = 1 Or .flags.Oculto = 1 Then
            .flags.Oculto = 0
            .flags.invisible = 0
            .Counters.TiempoOculto = 0
            .Counters.Invisibilidad = 0
            
            'Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageSetInvisible(.Char.CharIndex, False))
            Call SetInvisible(Userindex, UserList(Userindex).Char.CharIndex, False)

        End If
        
        If TriggerZonaPelea(Userindex, Userindex) <> eTrigger6.TRIGGER6_PERMITE Then

            If DropItemsAlMorir Then
                
                ' Si estas en zona segura no se caen los items.
                If MapInfo(.Pos.map).Pk Then
                
                    ' << Si es newbie no pierde el inventario >>
                    If Not EsNewbie(Userindex) Then
                        Call TirarTodo(Userindex)
                    Else
                        Call TirarTodosLosItemsNoNewbies(Userindex)
    
                    End If
                    
                End If

            End If

        End If
        
        ' DESEQUIPA TODOS LOS OBJETOS
        'desequipar armadura
        If .Invent.ArmourEqpObjIndex > 0 Then
            Call Desequipar(Userindex, .Invent.ArmourEqpSlot)

        End If
        
        'desequipar arma
        If .Invent.WeaponEqpObjIndex > 0 Then
            Call Desequipar(Userindex, .Invent.WeaponEqpSlot)

        End If
        
        'desequipar casco
        If .Invent.CascoEqpObjIndex > 0 Then
            Call Desequipar(Userindex, .Invent.CascoEqpSlot)

        End If
        
        'desequipar herramienta
        If .Invent.AnilloEqpSlot > 0 Then
            Call Desequipar(Userindex, .Invent.AnilloEqpSlot)

        End If
        
        'desequipar municiones
        If .Invent.MunicionEqpObjIndex > 0 Then
            Call Desequipar(Userindex, .Invent.MunicionEqpSlot)

        End If
        
        'desequipar escudo
        If .Invent.EscudoEqpObjIndex > 0 Then
            Call Desequipar(Userindex, .Invent.EscudoEqpSlot)

        End If
        
        ' << Reseteamos los posibles FX sobre el personaje >>
        If .Char.loops = INFINITE_LOOPS Then
            .Char.FX = 0
            .Char.loops = 0

        End If
        
        ' << Restauramos el mimetismo
        If .flags.Mimetizado = 1 Then
            .Char.body = .CharMimetizado.body
            .Char.Head = .CharMimetizado.Head
            .Char.CascoAnim = .CharMimetizado.CascoAnim
            .Char.ShieldAnim = .CharMimetizado.ShieldAnim
            .Char.WeaponAnim = .CharMimetizado.WeaponAnim
            .Counters.Mimetismo = 0
            .flags.Mimetizado = 0
            ' Puede ser atacado por npcs (cuando resucite)
            .flags.Ignorado = False

        End If
        
        ' << Restauramos los atributos >>
        If .flags.TomoPocion = True Then

            For i = 1 To 5
                .Stats.UserAtributos(i) = .Stats.UserAtributosBackUP(i)
            Next i

        End If
        
        '<< Cambiamos la apariencia del char >>
        If .flags.Navegando = 0 Then
            .Char.body = iCuerpoMuerto
            .Char.Head = iCabezaMuerto
            .Char.ShieldAnim = NingunEscudo
            .Char.WeaponAnim = NingunArma
            .Char.CascoAnim = NingunCasco
        Else
            .Char.body = iFragataFantasmal

        End If
        
        For i = 1 To MAXMASCOTAS

            If .MascotasIndex(i) > 0 Then
                Call MuereNpc(.MascotasIndex(i), 0)
                ' Si estan en agua o zona segura
            Else
                .MascotasType(i) = 0

            End If

        Next i
        
        .NroMascotas = 0
        
        '<< Actualizamos clientes >>
        Call ChangeUserChar(Userindex, .Char.body, .Char.Head, .Char.heading, NingunArma, NingunEscudo, NingunCasco)
        Call WriteUpdateUserStats(Userindex)
        Call WriteUpdateStrenghtAndDexterity(Userindex)

        '<<Castigos por party>>
        If .PartyIndex > 0 Then
            Call mdParty.ObtenerExito(Userindex, .Stats.ELV * -10 * mdParty.CantMiembros(Userindex), .Pos.map, .Pos.x, .Pos.y)

        End If
        
        '<<Cerramos comercio seguro>>
        Call LimpiarComercioSeguro(Userindex)
        
        'DEATHMATCH
        Call mod_Deathmatch.Muere_Death(Userindex)
        
        If .Pos.map = mod_EventoBestia.Bestia_Map Then
            Call Bestia_UserDie(Userindex)
        End If
        'torneos nVSn
        If .userTorneo.EnTorneo = True Then
            Call mod_Torneos.torneo_Muere(Userindex, False)
        End If
        
        ' Hay que teletransportar?
        Dim Mapa As Integer

        Mapa = .Pos.map

        Dim MapaTelep As Integer

        MapaTelep = MapInfo(Mapa).OnDeathGoTo.map
        
        If MapaTelep <> 0 Then
            Call WriteConsoleMsg(Userindex, "Tu estado no te permite permanecer en el mapa!!!", FontTypeNames.FONTTYPE_INFOBOLD)
            Call WarpUserChar(Userindex, MapaTelep, MapInfo(Mapa).OnDeathGoTo.x, MapInfo(Mapa).OnDeathGoTo.y, True, True)

        End If
        
        ' Retos nVSn. User muere
        If AttackerIndex <> 0 Then
            If .flags.SlotReto > 0 Then
                Call Retos.UserDieFight(Userindex, AttackerIndex, False)
            End If
        End If
    End With

    Exit Sub

ErrorHandler:
    Call LogError("Error en SUB USERDIE. Error: " & Err.Number & " Descripcion: " & Err.description)

End Sub

Public Sub ContarMuerte(ByVal Muerto As Integer, ByVal Atacante As Integer)
    '***************************************************
    'Author: Unknown
    'Last Modification: 13/07/2010
    '13/07/2010: ZaMa - Los matados en estado atacable ya no suman frag.
    '***************************************************

    If EsNewbie(Muerto) Then Exit Sub
        
    With UserList(Atacante)

        If TriggerZonaPelea(Muerto, Atacante) = TRIGGER6_PERMITE Then Exit Sub
        
        If esLegion(Muerto) Then
            If .flags.LastCrimMatado <> UserList(Muerto).Name Then
                .flags.LastCrimMatado = UserList(Muerto).Name

                If .Faccion.CriminalesMatados < MAXUSERMATADOS Then .Faccion.CriminalesMatados = .Faccion.CriminalesMatados + IIf(eventox2 = eEventoX2.kills, 2, 1)

            End If

        ElseIf esNeutral(Muerto) Then
            If .flags.LastCrimMatado <> UserList(Muerto).Name Then
                .flags.LastCrimMatado = UserList(Muerto).Name

                If .Faccion.NeutralesMatados < MAXUSERMATADOS Then .Faccion.NeutralesMatados = .Faccion.NeutralesMatados + IIf(eventox2 = eEventoX2.kills, 2, 1)

            End If
        ElseIf esArmada(Muerto) Then

            If .flags.LastCiudMatado <> UserList(Muerto).Name Then
                .flags.LastCiudMatado = UserList(Muerto).Name

                If .Faccion.CiudadanosMatados < MAXUSERMATADOS Then .Faccion.CiudadanosMatados = .Faccion.CiudadanosMatados + IIf(eventox2 = eEventoX2.kills, 2, 1)

            End If

        End If
        
        If .Stats.UsuariosMatados < MAXUSERMATADOS Then .Stats.UsuariosMatados = .Stats.UsuariosMatados + IIf(eventox2 = eEventoX2.kills, 2, 1)

    End With

End Sub

Sub Tilelibre(ByRef Pos As WorldPos, _
              ByRef nPos As WorldPos, _
              ByRef obj As obj, _
              ByRef PuedeAgua As Boolean, _
              ByRef PuedeTierra As Boolean)

    '**************************************************************
    'Author: Unknown
    'Last Modify Date: 18/09/2010
    '23/01/2007 -> Pablo (ToxicWaste): El agua es ahora un TileLibre agregando las condiciones necesarias.
    '18/09/2010: ZaMa - Aplico optimizacion de busqueda de tile libre en forma de rombo.
    '**************************************************************
    On Error GoTo errHandler

    Dim Found As Boolean

    Dim LoopC As Integer

    Dim tX    As Long

    Dim tY    As Long
    
    nPos = Pos
    tX = Pos.x
    tY = Pos.y
    
    LoopC = 1
    
    ' La primera posicion es valida?
    If LegalPos(Pos.map, nPos.x, nPos.y, PuedeAgua, PuedeTierra, True) Then
        
        If Not HayObjeto(Pos.map, nPos.x, nPos.y, obj.ObjIndex, obj.amount) Then
            Found = True

        End If
        
    End If
    
    ' Busca en las demas posiciones, en forma de "rombo"
    If Not Found Then

        While (Not Found) And LoopC <= 16

            If RhombLegalTilePos(Pos, tX, tY, LoopC, obj.ObjIndex, obj.amount, PuedeAgua, PuedeTierra) Then
                nPos.x = tX
                nPos.y = tY
                Found = True

            End If
        
            LoopC = LoopC + 1
        Wend
        
    End If
    
    If Not Found Then
        nPos.x = 0
        nPos.y = 0

    End If
    
    Exit Sub
    
errHandler:
    Call LogError("Error en Tilelibre. Error: " & Err.Number & " - " & Err.description)

End Sub

Sub WarpUserChar(ByVal Userindex As Integer, _
                 ByVal map As Integer, _
                 ByVal x As Integer, _
                 ByVal y As Integer, _
                 ByVal FX As Boolean, _
                 Optional ByVal Teletransported As Boolean)

    '**************************************************************
    'Author: Unknown
    'Last Modify Date: 11/23/2010
    '15/07/2009 - ZaMa: Automatic toogle navigate after warping to water.
    '13/11/2009 - ZaMa: Now it's activated the timer which determines if the npc can atacak the user.
    '16/09/2010 - ZaMa: No se pierde la visibilidad al cambiar de mapa al estar navegando invisible.
    '11/23/2010 - C4b3z0n: Ahora si no se permite Invi o Ocultar en el mapa al que cambias, te lo saca
    '**************************************************************
    Dim OldMap As Integer

    Dim OldX   As Integer

    Dim OldY   As Integer
    
    With UserList(Userindex)
        'Quitar el dialogo
        Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessageRemoveCharDialog(.Char.CharIndex))
        
        OldMap = .Pos.map
        OldX = .Pos.x
        OldY = .Pos.y

        Call EraseUserChar(Userindex, .flags.AdminInvisible = 1)
        
        If OldMap <> map Then
            Call WriteChangeMap(Userindex, map, MapInfo(.Pos.map).MapVersion)
            
            If .flags.Privilegios And PlayerType.User Then 'El chequeo de invi/ocultar solo afecta a Usuarios (C4b3z0n)

                Dim AhoraVisible As Boolean 'Para enviar el mensaje de invi y hacer visible (C4b3z0n)

                Dim WasInvi      As Boolean

                'Chequeo de flags de mapa por invisibilidad (C4b3z0n)
                If MapInfo(map).InviSinEfecto > 0 And .flags.invisible = 1 Then
                    .flags.invisible = 0
                    .Counters.Invisibilidad = 0
                    AhoraVisible = True
                    WasInvi = True 'si era invi, para el string

                End If

                'Chequeo de flags de mapa por ocultar (C4b3z0n)
                If MapInfo(map).OcultarSinEfecto > 0 And .flags.Oculto = 1 Then
                    AhoraVisible = True
                    .flags.Oculto = 0
                    .Counters.TiempoOculto = 0

                End If
                
                If AhoraVisible Then 'Si no era visible y ahora es, le avisa. (C4b3z0n)
                    Call SetInvisible(Userindex, .Char.CharIndex, False)

                    If WasInvi Then 'era invi
                        Call WriteConsoleMsg(Userindex, "Has vuelto a ser visible ya que no esta permitida la invisibilidad en este mapa.", FontTypeNames.FONTTYPE_INFO)
                    Else 'estaba oculto
                        Call WriteConsoleMsg(Userindex, "Has vuelto a ser visible ya que no esta permitido ocultarse en este mapa.", FontTypeNames.FONTTYPE_INFO)

                    End If

                End If

            End If

            'Si tiene MP3 el mapa mandamos que lo reproduzca, sino reproducimos el MIDI de toda la vida
            If MapInfo(map).MusicMp3 <> vbNullString Then
                Call WritePlayMp3(Userindex, MapInfo(map).MusicMp3)
            Else
                Call WritePlayMidi(Userindex, val(ReadField(1, MapInfo(map).Music, 45)))
            End If

            'Update new Map Users
            MapInfo(map).numUsers = MapInfo(map).numUsers + 1
            
            'Update old Map Users
            MapInfo(OldMap).numUsers = MapInfo(OldMap).numUsers - 1

            If MapInfo(OldMap).numUsers < 0 Then
                MapInfo(OldMap).numUsers = 0
            End If
        
            'Si el mapa al que entro NO ES superficial AND en el que estaba TAMPOCO ES superficial, ENTONCES
            Dim nextMap, previousMap As Boolean

            nextMap = IIf(distanceToCities(map).distanceToCity(.Hogar) >= 0, True, False)
            previousMap = IIf(distanceToCities(.Pos.map).distanceToCity(.Hogar) >= 0, True, False)

            If previousMap And nextMap Then '138 => 139 (Ambos superficiales, no tiene que pasar nada)
                'NO PASA NADA PORQUE NO ENTRO A UN DUNGEON.
            ElseIf previousMap And Not nextMap Then '139 => 140 (139 es superficial, 140 no. Por lo tanto 139 es el ultimo mapa superficial)
                .flags.lastMap = .Pos.map
            ElseIf Not previousMap And nextMap Then '140 => 139 (140 es no es superficial, 139 si. Por lo tanto, el ultimo mapa es 0 ya que no esta en un dungeon)
                .flags.lastMap = 0
            ElseIf Not previousMap And Not nextMap Then '140 => 141 (Ninguno es superficial, el ultimo mapa es el mismo de antes)
                .flags.lastMap = .flags.lastMap

            End If
            
            Call WriteRemoveAllDialogs(Userindex)

        End If
        
        .Pos.x = x
        .Pos.y = y
        .Pos.map = map
        
        Call MakeUserChar(True, map, Userindex, map, x, y)
        Call WriteUserCharIndexInServer(Userindex)
        
        Call DoTileEvents(Userindex, map, x, y)
        
        If Teletransported Then
            If .flags.Traveling = 1 Then
                .flags.Traveling = 0
                .Counters.goHome = 0
                Call WriteMultiMessage(Userindex, eMessages.CancelHome)

            End If

        End If
        
        If FX And .flags.AdminInvisible = 0 Then 'FX
            Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessagePlayWave(SND_WARP, x, y))
            Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessageCreateFX(.Char.CharIndex, FXIDs.FXWARP, 0))

        End If
        
        If .NroMascotas Then Call WarpMascotas(Userindex)
        
        ' No puede ser atacado cuando cambia de mapa, por cierto tiempo
        Call IntervaloPermiteSerAtacado(Userindex, True)
        
        ' Perdes el npc al cambiar de mapa
        Call PerdioNpc(Userindex, False)
        
        ' Automatic toogle navigate
        If (.flags.Privilegios And (PlayerType.User Or PlayerType.Consejero)) = 0 Then
            If HayAgua(.Pos.map, .Pos.x, .Pos.y) Then
                If .flags.Navegando = 0 Then
                    .flags.Navegando = 1
                        
                    'Tell the client that we are navigating.
                    Call WriteNavigateToggle(Userindex)

                End If

            Else

                If .flags.Navegando = 1 Then
                    .flags.Navegando = 0
                            
                    'Tell the client that we are navigating.
                    Call WriteNavigateToggle(Userindex)

                End If

            End If

        End If
      
    End With

End Sub

Private Sub WarpMascotas(ByVal Userindex As Integer)

    '************************************************
    'Author: Uknown
    'Last Modified: 26/10/2010
    '13/02/2009: ZaMa - Arreglado respawn de mascotas al cambiar de mapa.
    '13/02/2009: ZaMa - Las mascotas no regeneran su vida al cambiar de mapa (Solo entre mapas inseguros).
    '11/05/2009: ZaMa - Chequeo si la mascota pueden spwnear para asiganrle los stats.
    '26/10/2010: ZaMa - Ahora las mascotas rapswnean de forma aleatoria.
    '************************************************
    Dim i                As Integer

    Dim petType          As Integer

    Dim PetRespawn       As Boolean

    Dim PetTiempoDeVida  As Integer

    Dim NroPets          As Integer

    Dim InvocadosMatados As Integer

    Dim canWarp          As Boolean

    Dim index            As Integer

    Dim iMinHP           As Integer
    
    NroPets = UserList(Userindex).NroMascotas
    canWarp = (MapInfo(UserList(Userindex).Pos.map).Pk = True)
    
    For i = 1 To MAXMASCOTAS
        index = UserList(Userindex).MascotasIndex(i)
        
        If index > 0 Then

            ' si la mascota tiene tiempo de vida > 0 significa q fue invocada => we kill it
            If Npclist(index).Contadores.TiempoExistencia > 0 Then
                Call QuitarNPC(index)
                UserList(Userindex).MascotasIndex(i) = 0
                InvocadosMatados = InvocadosMatados + 1
                NroPets = NroPets - 1
                
                petType = 0
            Else
                'Store data and remove NPC to recreate it after warp
                'PetRespawn = Npclist(index).flags.Respawn = 0
                petType = UserList(Userindex).MascotasType(i)
                'PetTiempoDeVida = Npclist(index).Contadores.TiempoExistencia
                
                ' Guardamos el hp, para restaurarlo uando se cree el npc
                iMinHP = Npclist(index).Stats.MinHp
                
                Call QuitarNPC(index)
                
                ' Restauramos el valor de la variable
                UserList(Userindex).MascotasType(i) = petType

            End If

        ElseIf UserList(Userindex).MascotasType(i) > 0 Then
            'Store data and remove NPC to recreate it after warp
            PetRespawn = True
            petType = UserList(Userindex).MascotasType(i)
            PetTiempoDeVida = 0
        Else
            petType = 0

        End If
        
        If petType > 0 And canWarp Then
        
            Dim SpawnPos As WorldPos
        
            SpawnPos.map = UserList(Userindex).Pos.map
            SpawnPos.x = UserList(Userindex).Pos.x + RandomNumber(-3, 3)
            SpawnPos.y = UserList(Userindex).Pos.y + RandomNumber(-3, 3)
        
            index = SpawnNpc(petType, SpawnPos, False, PetRespawn)
            
            'Controlamos que se sumoneo OK - should never happen. Continue to allow removal of other pets if not alone
            ' Exception: Pets don't spawn in water if they can't swim
            If index = 0 Then
                Call WriteConsoleMsg(Userindex, "Tus mascotas no pueden transitar este mapa.", FontTypeNames.FONTTYPE_INFO)
            Else
                UserList(Userindex).MascotasIndex(i) = index

                ' Nos aseguramos de que conserve el hp, si estaba danado
                Npclist(index).Stats.MinHp = IIf(iMinHP = 0, Npclist(index).Stats.MinHp, iMinHP)
            
                Npclist(index).MaestroUser = Userindex
                Npclist(index).Contadores.TiempoExistencia = PetTiempoDeVida
                Call FollowAmo(index)

            End If

        End If

    Next i
    
    If InvocadosMatados > 0 Then
        Call WriteConsoleMsg(Userindex, "Pierdes el control de tus mascotas invocadas.", FontTypeNames.FONTTYPE_INFO)

    End If
    
    If Not canWarp Then
        Call WriteConsoleMsg(Userindex, "No se permiten mascotas en zona segura. estas te esperaran afuera.", FontTypeNames.FONTTYPE_INFO)

    End If
    
    UserList(Userindex).NroMascotas = NroPets

End Sub

Public Sub WarpMascota(ByVal Userindex As Integer, ByVal PetIndex As Integer)

    '************************************************
    'Author: ZaMa
    'Last Modified: 18/11/2009
    'Warps a pet without changing its stats
    '************************************************
    Dim petType   As Integer

    Dim NpcIndex  As Integer

    Dim iMinHP    As Integer

    Dim TargetPos As WorldPos
    
    With UserList(Userindex)
        
        TargetPos.map = .flags.TargetMap
        TargetPos.x = .flags.TargetX
        TargetPos.y = .flags.TargetY
        
        NpcIndex = .MascotasIndex(PetIndex)
            
        'Store data and remove NPC to recreate it after warp
        petType = .MascotasType(PetIndex)
        
        ' Guardamos el hp, para restaurarlo cuando se cree el npc
        iMinHP = Npclist(NpcIndex).Stats.MinHp
        
        Call QuitarNPC(NpcIndex)
        
        ' Restauramos el valor de la variable
        .MascotasType(PetIndex) = petType
        .NroMascotas = .NroMascotas + 1
        NpcIndex = SpawnNpc(petType, TargetPos, False, False)
        
        'Controlamos que se sumoneo OK - should never happen. Continue to allow removal of other pets if not alone
        ' Exception: Pets don't spawn in water if they can't swim
        If NpcIndex = 0 Then
            Call WriteConsoleMsg(Userindex, "Tu mascota no pueden transitar este sector del mapa, intenta invocarla en otra parte.", FontTypeNames.FONTTYPE_INFO)
        Else
            .MascotasIndex(PetIndex) = NpcIndex

            With Npclist(NpcIndex)
                ' Nos aseguramos de que conserve el hp, si estaba danado
                .Stats.MinHp = IIf(iMinHP = 0, .Stats.MinHp, iMinHP)
            
                .MaestroUser = Userindex
                .Movement = TipoAI.SigueAmo
                .Target = 0
                .TargetNPC = 0

            End With
            
            Call FollowAmo(NpcIndex)

        End If

    End With

End Sub

''
' Se inicia la salida de un usuario.
'
' @param    UserIndex   El index del usuario que va a salir

Sub Cerrar_Usuario(ByVal Userindex As Integer)

    '***************************************************
    'Author: Unknown
    'Last Modification: 16/09/2010
    '16/09/2010 - ZaMa: Cuando se va el invi estando navegando, no se saca el invi (ya esta visible).
    '***************************************************
    Dim isNotVisible As Boolean

    Dim HiddenPirat  As Boolean
    
    With UserList(Userindex)

        If .flags.UserLogged And Not .Counters.Saliendo Then
            .Counters.Saliendo = True
            .Counters.Salir = IIf((.flags.Privilegios And PlayerType.User) And MapInfo(.Pos.map).Pk, IntervaloCerrarConexion, 0)
            
            isNotVisible = (.flags.Oculto Or .flags.invisible)

            If isNotVisible Then
                .flags.invisible = 0
                
                If .flags.Oculto Then
                    If .flags.Navegando = 1 Then
                        If .Clase = eClass.Pirat Then
                            ' Pierde la apariencia de fragata fantasmal
                            Call ToggleBoatBody(Userindex)
                            Call WriteConsoleMsg(Userindex, "Has recuperado tu apariencia normal!", FontTypeNames.FONTTYPE_INFO)
                            Call ChangeUserChar(Userindex, .Char.body, .Char.Head, .Char.heading, NingunArma, NingunEscudo, NingunCasco)
                            HiddenPirat = True

                        End If

                    End If

                End If
                
                .flags.Oculto = 0
                
                ' Para no repetir mensajes
                If Not HiddenPirat Then Call WriteConsoleMsg(Userindex, "Has vuelto a ser visible.", FontTypeNames.FONTTYPE_INFO)
                
                ' Si esta navegando ya esta visible
                If .flags.Navegando = 0 Then
                    Call SetInvisible(Userindex, .Char.CharIndex, False)

                End If

            End If
            
            If .flags.Traveling = 1 Then
                Call WriteMultiMessage(Userindex, eMessages.CancelHome)
                .flags.Traveling = 0
                .Counters.goHome = 0

            End If
            
            Call WriteConsoleMsg(Userindex, "Cerrando...Se cerrara el juego en " & .Counters.Salir & " segundos...", FontTypeNames.FONTTYPE_INFO)

        End If

    End With

End Sub

''
' Cancels the exit of a user. If it's disconnected it's reset.
'
' @param    UserIndex   The index of the user whose exit is being reset.

Public Sub CancelExit(ByVal Userindex As Integer)

    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 04/02/08
    '
    '***************************************************
    If UserList(Userindex).Counters.Saliendo Then

        ' Is the user still connected?
        If UserList(Userindex).ConnIDValida Then
            UserList(Userindex).Counters.Saliendo = False
            UserList(Userindex).Counters.Salir = 0
            Call WriteConsoleMsg(Userindex, "/salir cancelado.", FontTypeNames.FONTTYPE_WARNING)
        Else
            'Simply reset
            UserList(Userindex).Counters.Salir = IIf((UserList(Userindex).flags.Privilegios And PlayerType.User) And MapInfo(UserList(Userindex).Pos.map).Pk, IntervaloCerrarConexion, 0)

        End If

    End If

End Sub


''
'Checks if a given body index is a boat or not.
'
'@param body    The body index to bechecked.
'@return    True if the body is a boat, false otherwise.

Public Function BodyIsBoat(ByVal body As Integer) As Boolean

    '**************************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modify Date: 10/07/2008
    'Checks if a given body index is a boat
    '**************************************************************
    'TODO : This should be checked somehow else. This is nasty....
    If body = iFragataReal Or body = iFragataCaos Or body = iBarcaPk Or body = iGaleraPk Or body = iGaleonPk Or body = iBarcaCiuda Or body = iGaleraCiuda Or body = iGaleonCiuda Or body = iFragataFantasmal Then
        BodyIsBoat = True

    End If

End Function

Public Sub SetInvisible(ByVal Userindex As Integer, _
                        ByVal userCharIndex As Integer, _
                        ByVal invisible As Boolean)
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    Dim sndNick As String

    With UserList(Userindex)
        Call SendData(SendTarget.ToUsersAndRmsAndCounselorsAreaButGMs, Userindex, PrepareMessageSetInvisible(userCharIndex, invisible))
    
        sndNick = .Name
    
        If invisible Then
            sndNick = sndNick & " " & TAG_USER_INVISIBLE
        Else

            If .GuildIndex > 0 Then
                sndNick = sndNick & " <" & modGuilds.GuildName(.GuildIndex) & ">"

            End If

        End If
    
        Call SendData(SendTarget.ToGMsAreaButRmsOrCounselors, Userindex, PrepareMessageCharacterChangeNick(userCharIndex, sndNick))

    End With

End Sub

Public Sub SetConsulatMode(ByVal Userindex As Integer)
    '***************************************************
    'Author: Torres Patricio (Pato)
    'Last Modification: 05/06/10
    '
    '***************************************************

    Dim sndNick As String

    With UserList(Userindex)
        sndNick = .Name
    
        If .flags.EnConsulta Then
            sndNick = sndNick & " " & TAG_CONSULT_MODE
        Else

            If .GuildIndex > 0 Then
                sndNick = sndNick & " <" & modGuilds.GuildName(.GuildIndex) & ">"

            End If

        End If
    
        Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessageCharacterChangeNick(.Char.CharIndex, sndNick))

    End With

End Sub

Public Function IsArena(ByVal Userindex As Integer) As Boolean
    '**************************************************************
    'Author: ZaMa
    'Last Modify Date: 10/11/2009
    'Returns true if the user is in an Arena
    '**************************************************************
    IsArena = (TriggerZonaPelea(Userindex, Userindex) = TRIGGER6_PERMITE)

End Function

Public Sub PerdioNpc(ByVal Userindex As Integer, _
                     Optional ByVal CheckPets As Boolean = True)
    '**************************************************************
    'Author: ZaMa
    'Last Modify Date: 11/07/2010 (ZaMa)
    'The user loses his owned npc
    '18/01/2010: ZaMa - Las mascotas dejan de atacar al npc que se perdio.
    '11/07/2010: ZaMa - Coloco el indice correcto de las mascotas y ahora siguen al amo si existen.
    '13/07/2010: ZaMa - Ahora solo dejan de atacar las mascotas si estan atacando al npc que pierde su amo.
    '**************************************************************

    Dim PetCounter As Long

    Dim PetIndex   As Integer

    Dim NpcIndex   As Integer
    
    With UserList(Userindex)
        
        NpcIndex = .flags.OwnedNpc

        If NpcIndex > 0 Then
            
            If CheckPets Then

                ' Dejan de atacar las mascotas
                If .NroMascotas > 0 Then

                    For PetCounter = 1 To MAXMASCOTAS
                    
                        PetIndex = .MascotasIndex(PetCounter)
                        
                        If PetIndex > 0 Then

                            ' Si esta atacando al npc deja de hacerlo
                            If Npclist(PetIndex).TargetNPC = NpcIndex Then
                                Call FollowAmo(PetIndex)

                            End If

                        End If
                        
                    Next PetCounter

                End If

            End If
            
            ' Reset flags
            Npclist(NpcIndex).Owner = 0
            .flags.OwnedNpc = 0

        End If

    End With

End Sub

Public Sub ApropioNpc(ByVal Userindex As Integer, ByVal NpcIndex As Integer)
    '**************************************************************
    'Author: ZaMa
    'Last Modify Date: 27/07/2010 (zaMa)
    'The user owns a new npc
    '18/01/2010: ZaMa - El sistema no aplica a zonas seguras.
    '19/04/2010: ZaMa - Ahora los admins no se pueden apropiar de npcs.
    '27/07/2010: ZaMa - El sistema no aplica a mapas seguros.
    '**************************************************************

    With UserList(Userindex)

        ' Los admins no se pueden apropiar de npcs
        If EsGm(Userindex) Then Exit Sub
        
        Dim Mapa As Integer

        Mapa = .Pos.map
        
        ' No aplica a triggers seguras
        If MapData(Mapa, .Pos.x, .Pos.y).trigger = eTrigger.ZONASEGURA Then Exit Sub
        
        ' No se aplica a mapas seguros
        If MapInfo(Mapa).Pk = False Then Exit Sub
        
        ' No aplica a algunos mapas que permiten el robo de npcs
        If MapInfo(Mapa).RoboNpcsPermitido = 1 Then Exit Sub
        
        ' Pierde el npc anterior
        If .flags.OwnedNpc > 0 Then Npclist(.flags.OwnedNpc).Owner = 0
        
        ' Si tenia otro dueno, lo perdio aca
        Npclist(NpcIndex).Owner = Userindex
        .flags.OwnedNpc = NpcIndex

    End With
    
    ' Inicializo o actualizo el timer de pertenencia
    Call IntervaloPerdioNpc(Userindex, True)

End Sub

Public Function GetDireccion(ByVal Userindex As Integer, _
                             ByVal OtherUserIndex As Integer) As String

    '**************************************************************
    'Author: ZaMa
    'Last Modify Date: 17/11/2009
    'Devuelve la direccion hacia donde esta el usuario
    '**************************************************************
    Dim x As Integer

    Dim y As Integer
    
    x = UserList(Userindex).Pos.x - UserList(OtherUserIndex).Pos.x
    y = UserList(Userindex).Pos.y - UserList(OtherUserIndex).Pos.y
    
    If x = 0 And y > 0 Then
        GetDireccion = "Sur"
    ElseIf x = 0 And y < 0 Then
        GetDireccion = "Norte"
    ElseIf x > 0 And y = 0 Then
        GetDireccion = "Este"
    ElseIf x < 0 And y = 0 Then
        GetDireccion = "Oeste"
    ElseIf x > 0 And y < 0 Then
        GetDireccion = "NorEste"
    ElseIf x < 0 And y < 0 Then
        GetDireccion = "NorOeste"
    ElseIf x > 0 And y > 0 Then
        GetDireccion = "SurEste"
    ElseIf x < 0 And y > 0 Then
        GetDireccion = "SurOeste"

    End If

End Function

Public Function SameFaccion(ByVal Userindex As Integer, _
                            ByVal OtherUserIndex As Integer) As Boolean
    '**************************************************************
    'Author: ZaMa
    'Last Modify Date: 17/11/2009
    'Devuelve True si son de la misma faccion
    '**************************************************************
    SameFaccion = (UserList(Userindex).Faccion.Bando = UserList(OtherUserIndex).Faccion.Bando) '(esCaos(Userindex) And esCaos(OtherUserIndex)) Or (esArmada(Userindex) And esArmada(OtherUserIndex))

End Function

Public Function FarthestPet(ByVal Userindex As Integer) As Integer

    '**************************************************************
    'Author: ZaMa
    'Last Modify Date: 18/11/2009
    'Devuelve el indice de la mascota mas lejana.
    '**************************************************************
    On Error GoTo errHandler
    
    Dim PetIndex      As Integer

    Dim Distancia     As Integer

    Dim OtraDistancia As Integer
    
    With UserList(Userindex)

        If .NroMascotas = 0 Then Exit Function
    
        For PetIndex = 1 To MAXMASCOTAS

            ' Solo pos invocar criaturas que exitan!
            If .MascotasIndex(PetIndex) > 0 Then

                ' Solo aplica a mascota, nada de elementales..
                If Npclist(.MascotasIndex(PetIndex)).Contadores.TiempoExistencia = 0 Then
                    If FarthestPet = 0 Then
                        ' Por si tiene 1 sola mascota
                        FarthestPet = PetIndex
                        Distancia = Abs(.Pos.x - Npclist(.MascotasIndex(PetIndex)).Pos.x) + Abs(.Pos.y - Npclist(.MascotasIndex(PetIndex)).Pos.y)
                    Else
                        ' La distancia de la proxima mascota
                        OtraDistancia = Abs(.Pos.x - Npclist(.MascotasIndex(PetIndex)).Pos.x) + Abs(.Pos.y - Npclist(.MascotasIndex(PetIndex)).Pos.y)

                        ' Esta mas lejos?
                        If OtraDistancia > Distancia Then
                            Distancia = OtraDistancia
                            FarthestPet = PetIndex

                        End If

                    End If

                End If

            End If

        Next PetIndex

    End With

    Exit Function
    
errHandler:
    Call LogError("Error en FarthestPet")

End Function

''
' Set the EluSkill value at the skill.
'
' @param UserIndex  Specifies reference to user
' @param Skill      Number of the skill to check
' @param Allocation True If the motive of the modification is the allocation, False if the skill increase by training

Public Sub CheckEluSkill(ByVal Userindex As Integer, _
                         ByVal Skill As Byte, _
                         ByVal Allocation As Boolean)
    '*************************************************
    'Author: Torres Patricio (Pato)
    'Last modified: 11/20/2009
    '
    '*************************************************

    With UserList(Userindex).Stats

        If .UserSkills(Skill) < MAXSKILLPOINTS Then
            If Allocation Then
                .ExpSkills(Skill) = 0
            Else
                .ExpSkills(Skill) = .ExpSkills(Skill) - .EluSkills(Skill)

            End If
        
            .EluSkills(Skill) = ELU_SKILL_INICIAL * 1.05 ^ .UserSkills(Skill)
        Else
            .ExpSkills(Skill) = 0
            .EluSkills(Skill) = 0

        End If

    End With

End Sub

Public Function HasEnoughItems(ByVal Userindex As Integer, _
                               ByVal ObjIndex As Integer, _
                               ByVal amount As Long) As Boolean
    '**************************************************************
    'Author: ZaMa
    'Last Modify Date: 25/11/2009
    'Cheks Wether the user has the required amount of items in the inventory or not
    '**************************************************************

    Dim Slot          As Long

    Dim ItemInvAmount As Long
    
    With UserList(Userindex)

        For Slot = 1 To .CurrentInventorySlots

            ' Si es el item que busco
            If .Invent.Object(Slot).ObjIndex = ObjIndex Then
                ' Lo sumo a la cantidad total
                ItemInvAmount = ItemInvAmount + .Invent.Object(Slot).amount

            End If

        Next Slot

    End With
    
    HasEnoughItems = amount <= ItemInvAmount

End Function

Public Function TotalOfferItems(ByVal ObjIndex As Integer, _
                                ByVal Userindex As Integer) As Long

    '**************************************************************
    'Author: ZaMa
    'Last Modify Date: 25/11/2009
    'Cheks the amount of items the user has in offerSlots.
    '**************************************************************
    Dim Slot As Byte
    
    For Slot = 1 To MAX_OFFER_SLOTS

        ' Si es el item que busco
        If UserList(Userindex).ComUsu.Objeto(Slot) = ObjIndex Then
            ' Lo sumo a la cantidad total
            TotalOfferItems = TotalOfferItems + UserList(Userindex).ComUsu.cant(Slot)

        End If

    Next Slot

End Function

Public Function getMaxInventorySlots(ByVal Userindex As Integer) As Byte
    '***************************************************
    'Author: Unknown
    'Last Modification: 30/09/2020
    '
    '***************************************************

    If UserList(Userindex).Stats.InventLevel > 0 Then
        getMaxInventorySlots = MAX_USERINVENTORY_SLOTS + UserList(Userindex).Stats.InventLevel * SLOTS_PER_ROW_INVENTORY
    Else
        getMaxInventorySlots = MAX_USERINVENTORY_SLOTS
    End If

End Function

Public Sub goHome(ByVal Userindex As Integer)
    '***************************************************
    'Author: Budi
    'Last Modification: 01/06/2010
    '01/06/2010: ZaMa - Ahora usa otro tipo de intervalo
    '***************************************************

    Dim Distance As Long

    Dim Tiempo   As Long
    
    With UserList(Userindex)

        If .flags.Muerto = 1 Then
            If .flags.lastMap = 0 Then
                Distance = distanceToCities(.Pos.map).distanceToCity(.Hogar)
            Else
                Distance = distanceToCities(.flags.lastMap).distanceToCity(.Hogar) + GOHOME_PENALTY

            End If
            
            Tiempo = (Distance + 1) * 20 'seg
            
            If Tiempo > 60 Then
                Tiempo = 60
            End If
            
            Call IntervaloGoHome(Userindex, Tiempo * 1000, True)
                
            Call WriteMultiMessage(Userindex, eMessages.Home, Distance, Tiempo, , MapInfo(Ciudades(.Hogar).map).Name)
        Else
            Call WriteConsoleMsg(Userindex, "Debes estar muerto para poder utilizar este comando.", FontTypeNames.FONTTYPE_FIGHT)

        End If
        
    End With
    
End Sub

Public Sub setHome(ByVal Userindex As Integer, _
                   ByVal newHome As eCiudad, _
                   ByVal NpcIndex As Integer)

    '***************************************************
    'Author: Budi
    'Last Modification: 01/06/2010
    '30/04/2010: ZaMa - Ahora el npc avisa que se cambio de hogar.
    '01/06/2010: ZaMa - Ahora te avisa si ya tenes ese hogar.
    '***************************************************
    If newHome < eCiudad.cUllathorpe Or newHome > eCiudad.cLastCity - 1 Then Exit Sub
    
    If UserList(Userindex).Hogar <> newHome Then
        UserList(Userindex).Hogar = newHome
    
        Call WriteChatOverHead(Userindex, "Bienvenido a nuestra humilde comunidad, este es ahora tu nuevo hogar!!!", Npclist(NpcIndex).Char.CharIndex, vbWhite)
    Else
        Call WriteChatOverHead(Userindex, "Ya eres miembro de nuestra humilde comunidad!!!", Npclist(NpcIndex).Char.CharIndex, vbWhite)

    End If

End Sub

Public Function GetHomeArrivalTime(ByVal Userindex As Integer) As Integer

    '**************************************************************
    'Author: ZaMa
    'Last Modify by: ZaMa
    'Last Modify Date: 01/06/2010
    'Calculates the time left to arrive home.
    '**************************************************************
    Dim TActual As Long
    
    TActual = GetTickCount() And &H7FFFFFFF
    
    With UserList(Userindex)
        GetHomeArrivalTime = (.Counters.goHome - TActual) * 0.001

    End With

End Function

Public Sub HomeArrival(ByVal Userindex As Integer)
    '**************************************************************
    'Author: ZaMa
    'Last Modify by: ZaMa
    'Last Modify Date: 01/06/2010
    'Teleports user to its home.
    '**************************************************************
    
    Dim tX   As Integer

    Dim tY   As Integer

    Dim tMap As Integer

    With UserList(Userindex)

        'Antes de que el pj llegue a la ciudad, lo hacemos dejar de navegar para que no se buguee.
        If .flags.Navegando = 1 Then
            .Char.body = iCuerpoMuerto
            .Char.Head = iCabezaMuerto
            .Char.ShieldAnim = NingunEscudo
            .Char.WeaponAnim = NingunArma
            .Char.CascoAnim = NingunCasco
            
            .flags.Navegando = 0
            
            Call WriteNavigateToggle(Userindex)

            'Le sacamos el navegando, pero no le mostramos a los demas porque va a ser sumoneado hasta ulla.
        End If
        
        tX = Ciudades(.Hogar).x
        tY = Ciudades(.Hogar).y
        tMap = Ciudades(.Hogar).map
        
        Call FindLegalPos(Userindex, tMap, tX, tY)
        Call WarpUserChar(Userindex, tMap, tX, tY, True)
        
        Call WriteMultiMessage(Userindex, eMessages.FinishHome)
        
        .flags.Traveling = 0
        .Counters.goHome = 0
        
    End With
    
End Sub
