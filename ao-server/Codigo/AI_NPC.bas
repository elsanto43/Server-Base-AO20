Attribute VB_Name = "AI"
'Argentum Online 0.12.2
'Copyright (C) 2002 M�rquez Pablo Ignacio
'
'This program is free software; you can redistribute it and/or modify
'it under the terms of the Affero General Public License;
'either version 1 of the License, or any later version.


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
'Calle 3 n�mero 983 piso 7 dto A
'La Plata - Pcia, Buenos Aires - Republica Argentina
'C�digo Postal 1900
'Pablo Ignacio M�rquez

Option Explicit

Public Enum TipoAI

    ESTATICO = 1
    MueveAlAzar = 2
    NpcMaloAtacaUsersBuenos = 3
    NPCDEFENSA = 4
    GuardiasAtacanCriminales = 5
    NpcObjeto = 6
    SigueAmo = 8
    NpcAtacaNpc = 9
    NpcPathfinding = 10
    
    'Pretorianos
    SacerdotePretorianoAi = 20
    GuerreroPretorianoAi = 21
    MagoPretorianoAi = 22
    CazadorPretorianoAi = 23
    ReyPretoriano = 24

End Enum

Public Const ELEMENTALFUEGO  As Integer = 93
Public Const ELEMENTALTIERRA As Integer = 94
Public Const ELEMENTALAGUA   As Integer = 92

' WyroX: Tiles extra que ve el NPC mayor que los usuarios.
' Para que no queden tontos en un borde de la pantalla y puedas atacarlos.
Private Const VISION_EXTRA         As Byte = 2
Public Const RANGO_VISION_NPC_X    As Byte = RANGO_VISION_X + VISION_EXTRA
Public Const RANGO_VISION_NPC_Y    As Byte = RANGO_VISION_Y + VISION_EXTRA

'?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�
'?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�
'?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�
'                        Modulo AI_NPC
'?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�
'?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�
'?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�
'AI de los NPC
'?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�
'?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�
'?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�

Private Sub GuardiasAI(ByVal NpcIndex As Integer, ByVal DelCaos As Boolean)

    '***************************************************
    'Autor: Unknown (orginal version)
    'Last Modification: 12/01/2010 (ZaMa)
    '14/09/2009: ZaMa - Now npcs don't atack protected users.
    '12/01/2010: ZaMa - Los npcs no atacan druidas mimetizados con npcs
    '***************************************************
    Dim nPos          As WorldPos

    Dim headingloop   As Byte

    Dim ui            As Integer

    Dim UserProtected As Boolean
    
    With Npclist(NpcIndex)

        For headingloop = eHeading.NORTH To eHeading.WEST
            nPos = .Pos

            If .flags.Inmovilizado = 0 Or headingloop = .Char.heading Then
                Call HeadtoPos(headingloop, nPos)

                If InMapBounds(nPos.map, nPos.x, nPos.y) Then
                    ui = MapData(nPos.map, nPos.x, nPos.y).Userindex

                    If ui > 0 Then
                        'UserProtected = Not IntervaloPermiteSerAtacado(ui) And UserList(ui).flags.NoPuedeSerAtacado
                        UserProtected = UserProtected Or UserList(ui).flags.Ignorado Or UserList(ui).flags.EnConsulta
                        
                        If UserList(ui).flags.Muerto = 0 And UserList(ui).flags.AdminPerseguible And Not UserProtected Then

                            'guardias reales
                            If Not DelCaos Then
                                If esLegion(ui) Then
                                    If NpcAtacaUser(NpcIndex, ui) Then
                                        Call ChangeNPCChar(NpcIndex, .Char.body, .Char.Head, headingloop)

                                    End If

                                    Exit Sub
                                End If

                            Else 'guardias del mal

                                If esArmada(ui) Then
                                    If NpcAtacaUser(NpcIndex, ui) Then
                                        Call ChangeNPCChar(NpcIndex, .Char.body, .Char.Head, headingloop)

                                    End If

                                    Exit Sub

                                End If

                            End If

                        End If

                    End If

                End If

            End If  'not inmovil

        Next headingloop

    End With
    
    Call RestoreOldMovement(NpcIndex)

End Sub

''
' Handles the evil npcs' artificial intelligency.
'
' @param NpcIndex Specifies reference to the npc
Private Sub HostilMalvadoAI(ByVal NpcIndex As Integer)

    '**************************************************************
    'Author: Unknown
    'Last Modify Date: 12/01/2010 (ZaMa)
    '28/04/2009: ZaMa - Now those NPCs who doble attack, have 50% of posibility of casting a spell on user.
    '14/09/200*: ZaMa - Now npcs don't atack protected users.
    '12/01/2010: ZaMa - Los npcs no atacan druidas mimetizados con npcs
    '**************************************************************
    Dim nPos          As WorldPos

    Dim headingloop   As Byte

    Dim ui            As Integer

    Dim NPCI          As Integer

    Dim atacoPJ       As Boolean

    Dim UserProtected As Boolean
    
    atacoPJ = False
    
    With Npclist(NpcIndex)

        For headingloop = eHeading.NORTH To eHeading.WEST
            nPos = .Pos

            If .flags.Inmovilizado = 0 Or .Char.heading = headingloop Then
                Call HeadtoPos(headingloop, nPos)

                If InMapBounds(nPos.map, nPos.x, nPos.y) Then
                    ui = MapData(nPos.map, nPos.x, nPos.y).Userindex
                    NPCI = MapData(nPos.map, nPos.x, nPos.y).NpcIndex

                    If ui > 0 And Not atacoPJ Then
                        UserProtected = Not IntervaloPermiteSerAtacado(ui) And UserList(ui).flags.NoPuedeSerAtacado
                        UserProtected = UserProtected Or UserList(ui).flags.Ignorado Or UserList(ui).flags.EnConsulta
                        
                        If UserList(ui).flags.Muerto = 0 And UserList(ui).flags.AdminPerseguible And (Not UserProtected) Then
                            
                            atacoPJ = True

                            If .Movement = NpcObjeto Then

                                ' Los npc objeto no atacan siempre al mismo usuario
                                If RandomNumber(1, 3) = 3 Then atacoPJ = False

                            End If
                            
                            If atacoPJ Then
                                If .flags.LanzaSpells Then
                                    If .flags.AtacaDoble Then
                                        If (RandomNumber(0, 1)) Then
                                            If NpcAtacaUser(NpcIndex, ui) Then
                                                Call ChangeNPCChar(NpcIndex, .Char.body, .Char.Head, headingloop)

                                            End If

                                            Exit Sub

                                        End If

                                    End If
                                    
                                    Call ChangeNPCChar(NpcIndex, .Char.body, .Char.Head, headingloop)
                                    Call NpcLanzaUnSpell(NpcIndex, ui)

                                End If

                            End If

                            If NpcAtacaUser(NpcIndex, ui) Then
                                Call ChangeNPCChar(NpcIndex, .Char.body, .Char.Head, headingloop)

                            End If

                            Exit Sub

                        End If

                    ElseIf NPCI > 0 Then

                        If Npclist(NPCI).MaestroUser > 0 And Npclist(NPCI).flags.Paralizado = 0 Then
                            Call ChangeNPCChar(NpcIndex, .Char.body, .Char.Head, headingloop)
                            Call SistemaCombate.NpcAtacaNpc(NpcIndex, NPCI, False)
                            Exit Sub

                        End If

                    End If

                End If

            End If  'inmo

        Next headingloop

    End With
    
    Call RestoreOldMovement(NpcIndex)

End Sub

Private Sub HostilBuenoAI(ByVal NpcIndex As Integer)

    '***************************************************
    'Autor: Unknown (orginal version)
    'Last Modification: 12/01/2010 (ZaMa)
    '14/09/2009: ZaMa - Now npcs don't atack protected users.
    '12/01/2010: ZaMa - Los npcs no atacan druidas mimetizados con npcs
    '***************************************************
    Dim nPos          As WorldPos

    Dim headingloop   As eHeading

    Dim ui            As Integer

    Dim UserProtected As Boolean
    
    With Npclist(NpcIndex)

        For headingloop = eHeading.NORTH To eHeading.WEST
            nPos = .Pos

            If .flags.Inmovilizado = 0 Or .Char.heading = headingloop Then
                Call HeadtoPos(headingloop, nPos)

                If InMapBounds(nPos.map, nPos.x, nPos.y) Then
                    ui = MapData(nPos.map, nPos.x, nPos.y).Userindex

                    If ui > 0 Then
                        If UserList(ui).Name = .flags.AttackedBy Then
                        
                            UserProtected = Not IntervaloPermiteSerAtacado(ui) And UserList(ui).flags.NoPuedeSerAtacado
                            UserProtected = UserProtected Or UserList(ui).flags.Ignorado Or UserList(ui).flags.EnConsulta
                            
                            If UserList(ui).flags.Muerto = 0 And UserList(ui).flags.AdminPerseguible And Not UserProtected Then
                                If .flags.LanzaSpells > 0 Then
                                    Call NpcLanzaUnSpell(NpcIndex, ui)

                                End If
                                
                                If NpcAtacaUser(NpcIndex, ui) Then
                                    Call ChangeNPCChar(NpcIndex, .Char.body, .Char.Head, headingloop)

                                End If

                                Exit Sub

                            End If

                        End If

                    End If

                End If

            End If

        Next headingloop

    End With
    
    Call RestoreOldMovement(NpcIndex)

End Sub

Private Sub IrUsuarioCercano(ByVal NpcIndex As Integer)

    '***************************************************
    'Autor: Unknown (orginal version)
    'Last Modification: 25/07/2010 (ZaMa)
    '14/09/2009: ZaMa - Now npcs don't follow protected users.
    '12/01/2010: ZaMa - Los npcs no atacan druidas mimetizados con npcs
    '25/07/2010: ZaMa - Agrego una validacion temporal para evitar que los npcs ataquen a usuarios de mapas difernetes.
    '***************************************************
    Dim tHeading      As Byte

    Dim Userindex     As Integer

    Dim SignoNS       As Integer

    Dim SignoEO       As Integer

    Dim i             As Long

    Dim UserProtected As Boolean
    
    With Npclist(NpcIndex)

        If .flags.Inmovilizado = 1 Then

            Select Case .Char.heading

                Case eHeading.NORTH
                    SignoNS = -1
                    SignoEO = 0
                
                Case eHeading.EAST
                    SignoNS = 0
                    SignoEO = 1
                
                Case eHeading.SOUTH
                    SignoNS = 1
                    SignoEO = 0
                
                Case eHeading.WEST
                    SignoEO = -1
                    SignoNS = 0

            End Select
            
            For i = 1 To Areas.ConnGroups(.Pos.map).Count()
                Userindex = Areas.ConnGroups(.Pos.map).Item(i)
                
                'Is it in it's range of vision??
                If Abs(UserList(Userindex).Pos.x - .Pos.x) <= RANGO_VISION_NPC_X And Sgn(UserList(Userindex).Pos.x - .Pos.x) = SignoEO Then
                    If Abs(UserList(Userindex).Pos.y - .Pos.y) <= RANGO_VISION_NPC_Y And Sgn(UserList(Userindex).Pos.y - .Pos.y) = SignoNS Then
                        
                        UserProtected = Not IntervaloPermiteSerAtacado(Userindex) And UserList(Userindex).flags.NoPuedeSerAtacado
                        UserProtected = UserProtected Or UserList(Userindex).flags.Ignorado Or UserList(Userindex).flags.EnConsulta
                        
                        
                        If UserList(Userindex).flags.Muerto = 0 Then
                            If .level >= UserList(Userindex).Stats.ELV Then
                                If Not UserProtected Then
                                    If .flags.LanzaSpells <> 0 Then Call NpcLanzaUnSpell(NpcIndex, Userindex)
                                    Exit Sub
    
                                End If
                            End If
                        End If
                        
                    End If

                End If

            Next i
            
            ' No esta inmobilizado
        Else
            
            ' Tiene prioridad de seguir al usuario al que le pertenece si esta en el rango de vision
            Dim OwnerIndex As Integer
            
            OwnerIndex = .Owner

            If OwnerIndex > 0 Then
                
                ' TODO: Es temporal hatsa reparar un bug que hace que ataquen a usuarios de otros mapas
                If UserList(OwnerIndex).Pos.map = .Pos.map Then
                    
                    'Is it in it's range of vision??
                    If Abs(UserList(OwnerIndex).Pos.x - .Pos.x) <= RANGO_VISION_NPC_X Then
                        If Abs(UserList(OwnerIndex).Pos.y - .Pos.y) <= RANGO_VISION_NPC_Y Then
                            
                            ' va hacia el si no esta invi ni oculto
                            If UserList(OwnerIndex).flags.invisible = 0 And UserList(OwnerIndex).flags.Oculto = 0 And Not UserList(OwnerIndex).flags.EnConsulta And Not UserList(OwnerIndex).flags.Ignorado Then
                                If .flags.LanzaSpells <> 0 Then Call NpcLanzaUnSpell(NpcIndex, OwnerIndex)
                                    
                                tHeading = FindDirection(.Pos, UserList(OwnerIndex).Pos)
                                Call MoveNPCChar(NpcIndex, tHeading)
                                Exit Sub

                            End If

                        End If

                    End If
                
                    ' Esto significa que esta bugueado.. Lo logueo, y "reparo" el error a mano (Todo temporal)
                Else
                    Call LogError("El npc: " & .Name & "(" & NpcIndex & "), intenta atacar a " & UserList(OwnerIndex).Name & "(Index: " & OwnerIndex & ", Mapa: " & UserList(OwnerIndex).Pos.map & ") desde el mapa " & .Pos.map)
                    .Owner = 0

                End If
                
            End If
            
            ' No le pertenece a nadie o el dueno no esta en el rango de vision, sigue a cualquiera
            For i = 1 To Areas.ConnGroups(.Pos.map).Count()
                Userindex = Areas.ConnGroups(.Pos.map).Item(i)
                
                'Is it in it's range of vision??
                If Abs(UserList(Userindex).Pos.x - .Pos.x) <= RANGO_VISION_NPC_X Then
                    If Abs(UserList(Userindex).Pos.y - .Pos.y) <= RANGO_VISION_NPC_Y Then
                        
                        With UserList(Userindex)
                            
                            UserProtected = Not IntervaloPermiteSerAtacado(Userindex) And .flags.NoPuedeSerAtacado
                            UserProtected = UserProtected Or .flags.Ignorado Or .flags.EnConsulta
                            
                            If .flags.Muerto = 0 And .flags.invisible = 0 And .flags.Oculto = 0 And .flags.AdminPerseguible And Not UserProtected Then
                                If Npclist(NpcIndex).level >= UserList(Userindex).Stats.ELV Then
                                    If Npclist(NpcIndex).flags.LanzaSpells <> 0 Then Call NpcLanzaUnSpell(NpcIndex, Userindex)
                                    
                                    tHeading = FindDirection(Npclist(NpcIndex).Pos, .Pos)
                                    Call MoveNPCChar(NpcIndex, tHeading)
                                    Exit Sub
                                End If
                            End If
                            
                        End With
                        
                    End If

                End If

            Next i
            
            'Si llega aca es que no hab�a ning�n usuario cercano vivo.
            'A bailar. Pablo (ToxicWaste)
            If RandomNumber(0, 10) = 0 Then
                Call MoveNPCChar(NpcIndex, CByte(RandomNumber(eHeading.NORTH, eHeading.WEST)))

            End If
            
        End If

    End With
    
    Call RestoreOldMovement(NpcIndex)

End Sub

''
' Makes a Pet / Summoned Npc to Follow an enemy
'
' @param NpcIndex Specifies reference to the npc
Private Sub SeguirAgresor(ByVal NpcIndex As Integer)

    '**************************************************************
    'Author: Unknown
    'Last Modify by: Marco Vanotti (MarKoxX)
    'Last Modify Date: 08/16/2008
    '08/16/2008: MarKoxX - Now pets that do mel� attacks have to be near the enemy to attack.
    '**************************************************************
    Dim tHeading As Byte

    Dim ui       As Integer
    
    Dim i        As Long
    
    Dim SignoNS  As Integer

    Dim SignoEO  As Integer

    With Npclist(NpcIndex)

        If .flags.Paralizado = 1 Or .flags.Inmovilizado = 1 Then

            Select Case .Char.heading

                Case eHeading.NORTH
                    SignoNS = -1
                    SignoEO = 0
                
                Case eHeading.EAST
                    SignoNS = 0
                    SignoEO = 1
                
                Case eHeading.SOUTH
                    SignoNS = 1
                    SignoEO = 0
                
                Case eHeading.WEST
                    SignoEO = -1
                    SignoNS = 0

            End Select

            For i = 1 To Areas.ConnGroups(.Pos.map).Count()
                ui = Areas.ConnGroups(.Pos.map).Item(i)

                'Is it in it's range of vision??
                If Abs(UserList(ui).Pos.x - .Pos.x) <= RANGO_VISION_NPC_X And Sgn(UserList(ui).Pos.x - .Pos.x) = SignoEO Then
                    If Abs(UserList(ui).Pos.y - .Pos.y) <= RANGO_VISION_NPC_Y And Sgn(UserList(ui).Pos.y - .Pos.y) = SignoNS Then

                        If UserList(ui).Name = .flags.AttackedBy Then

                            If UserList(ui).flags.Muerto = 0 And UserList(ui).flags.invisible = 0 And UserList(ui).flags.Oculto = 0 Then
                                If .flags.LanzaSpells > 0 Then
                                    Call NpcLanzaUnSpell(NpcIndex, ui)
                                Else

                                    If Distancia(UserList(ui).Pos, Npclist(NpcIndex).Pos) <= 1 Then

                                        ' TODO : Set this a separate AI for Elementals and Druid's pets
                                        If Npclist(NpcIndex).Numero <> 92 Then
                                            Call NpcAtacaUser(NpcIndex, ui)

                                        End If

                                    End If

                                End If

                                Exit Sub

                            End If

                        End If
                        
                    End If

                End If
                
            Next i

        Else

            For i = 1 To Areas.ConnGroups(.Pos.map).Count()
                ui = Areas.ConnGroups(.Pos.map).Item(i)
                
                'Is it in it's range of vision??
                If Abs(UserList(ui).Pos.x - .Pos.x) <= RANGO_VISION_NPC_X Then
                    If Abs(UserList(ui).Pos.y - .Pos.y) <= RANGO_VISION_NPC_Y Then
                        
                        If UserList(ui).Name = .flags.AttackedBy Then
                            
                            If (UserList(ui).flags.Muerto = 0 And UserList(ui).flags.invisible = 0 And UserList(ui).flags.Oculto = 0) Or (.flags.SiguiendoGm = True) Then
                                If .flags.LanzaSpells > 0 Then
                                    Call NpcLanzaUnSpell(NpcIndex, ui)
                                Else

                                    If Distancia(UserList(ui).Pos, Npclist(NpcIndex).Pos) <= 1 Then

                                        ' TODO : Set this a separate AI for Elementals and Druid's pets
                                        If Npclist(NpcIndex).Numero <> 92 Then
                                            Call NpcAtacaUser(NpcIndex, ui)

                                        End If

                                    End If

                                End If
                                 
                                tHeading = FindDirection(.Pos, UserList(ui).Pos)
                                Call MoveNPCChar(NpcIndex, tHeading)
                                 
                                Exit Sub

                            End If

                        End If
                        
                    End If

                End If
                
            Next i

        End If

    End With
    
    Call RestoreOldMovement(NpcIndex)

End Sub

Private Sub RestoreOldMovement(ByVal NpcIndex As Integer)

    With Npclist(NpcIndex)

        If .MaestroUser = 0 Then
            .Movement = .flags.OldMovement
            .Hostile = .flags.OldHostil
            .flags.AttackedBy = vbNullString
            .flags.SiguiendoGm = False

        End If

    End With

End Sub

Private Sub PersigueCiudadano(ByVal NpcIndex As Integer)

    '***************************************************
    'Autor: Unknown (orginal version)
    'Last Modification: 12/01/2010 (ZaMa)
    '14/09/2009: ZaMa - Now npcs don't follow protected users.
    '12/01/2010: ZaMa - Los npcs no atacan druidas mimetizados con npcs.
    '***************************************************
    Dim Userindex     As Integer

    Dim tHeading      As Byte

    Dim i             As Long

    Dim UserProtected As Boolean
    
    With Npclist(NpcIndex)

        For i = 1 To Areas.ConnGroups(.Pos.map).Count()
            Userindex = Areas.ConnGroups(.Pos.map).Item(i)
                
            'Is it in it's range of vision??
            If Abs(UserList(Userindex).Pos.x - .Pos.x) <= RANGO_VISION_NPC_X Then
                If Abs(UserList(Userindex).Pos.y - .Pos.y) <= RANGO_VISION_NPC_Y Then
                    
                    If esArmada(Userindex) Then
                    
                        UserProtected = Not IntervaloPermiteSerAtacado(Userindex) And UserList(Userindex).flags.NoPuedeSerAtacado
                        UserProtected = UserProtected Or UserList(Userindex).flags.Ignorado Or UserList(Userindex).flags.EnConsulta
                        
                        If UserList(Userindex).flags.Muerto = 0 And UserList(Userindex).flags.invisible = 0 And UserList(Userindex).flags.Oculto = 0 And UserList(Userindex).flags.AdminPerseguible And Not UserProtected Then
                            
                            If .flags.LanzaSpells > 0 Then
                                Call NpcLanzaUnSpell(NpcIndex, Userindex)

                            End If

                            tHeading = FindDirection(.Pos, UserList(Userindex).Pos)
                            Call MoveNPCChar(NpcIndex, tHeading)
                            Exit Sub

                        End If

                    End If
                    
                End If

            End If
            
        Next i

    End With
    
    Call RestoreOldMovement(NpcIndex)

End Sub

Private Sub PersigueCriminal(ByVal NpcIndex As Integer)

    '***************************************************
    'Autor: Unknown (orginal version)
    'Last Modification: 12/01/2010 (ZaMa)
    '14/09/2009: ZaMa - Now npcs don't follow protected users.
    '12/01/2010: ZaMa - Los npcs no atacan druidas mimetizados con npcs.
    '***************************************************
    Dim Userindex     As Integer

    Dim tHeading      As Byte

    Dim i             As Long

    Dim SignoNS       As Integer

    Dim SignoEO       As Integer

    Dim UserProtected As Boolean
    
    With Npclist(NpcIndex)

        If .flags.Inmovilizado = 1 Then

            Select Case .Char.heading

                Case eHeading.NORTH
                    SignoNS = -1
                    SignoEO = 0
                
                Case eHeading.EAST
                    SignoNS = 0
                    SignoEO = 1
                
                Case eHeading.SOUTH
                    SignoNS = 1
                    SignoEO = 0
                
                Case eHeading.WEST
                    SignoEO = -1
                    SignoNS = 0

            End Select
            
            For i = 1 To Areas.ConnGroups(.Pos.map).Count()
                Userindex = Areas.ConnGroups(.Pos.map).Item(i)
                
                'Is it in it's range of vision??
                If Abs(UserList(Userindex).Pos.x - .Pos.x) <= RANGO_VISION_NPC_X And Sgn(UserList(Userindex).Pos.x - .Pos.x) = SignoEO Then
                    If Abs(UserList(Userindex).Pos.y - .Pos.y) <= RANGO_VISION_NPC_Y And Sgn(UserList(Userindex).Pos.y - .Pos.y) = SignoNS Then
                        
                        If esLegion(Userindex) Then

                            With UserList(Userindex)
                                 
                                UserProtected = Not IntervaloPermiteSerAtacado(Userindex) And .flags.NoPuedeSerAtacado
                                UserProtected = UserProtected Or UserList(Userindex).flags.Ignorado Or UserList(Userindex).flags.EnConsulta
                                 
                                If .flags.Muerto = 0 And .flags.invisible = 0 And .flags.Oculto = 0 And .flags.AdminPerseguible And Not UserProtected Then
                                     
                                    If Npclist(NpcIndex).flags.LanzaSpells > 0 Then
                                        Call NpcLanzaUnSpell(NpcIndex, Userindex)

                                    End If

                                    Exit Sub

                                End If

                            End With

                        End If
                        
                    End If

                End If

            Next i

        Else

            For i = 1 To Areas.ConnGroups(.Pos.map).Count()
                Userindex = Areas.ConnGroups(.Pos.map).Item(i)
                
                'Is it in it's range of vision??
                If Abs(UserList(Userindex).Pos.x - .Pos.x) <= RANGO_VISION_NPC_X Then
                    If Abs(UserList(Userindex).Pos.y - .Pos.y) <= RANGO_VISION_NPC_Y Then
                        
                        If esLegion(Userindex) Then
                            
                            UserProtected = Not IntervaloPermiteSerAtacado(Userindex) And UserList(Userindex).flags.NoPuedeSerAtacado
                            UserProtected = UserProtected Or UserList(Userindex).flags.Ignorado
                            
                            If UserList(Userindex).flags.Muerto = 0 And UserList(Userindex).flags.invisible = 0 And UserList(Userindex).flags.Oculto = 0 And UserList(Userindex).flags.AdminPerseguible And Not UserProtected Then

                                If .flags.LanzaSpells > 0 Then
                                    Call NpcLanzaUnSpell(NpcIndex, Userindex)

                                End If

                                If .flags.Inmovilizado = 1 Then Exit Sub
                                tHeading = FindDirection(.Pos, UserList(Userindex).Pos)
                                Call MoveNPCChar(NpcIndex, tHeading)
                                Exit Sub

                            End If

                        End If
                        
                    End If

                End If
                
            Next i

        End If

    End With
    
    Call RestoreOldMovement(NpcIndex)

End Sub

Private Sub SeguirAmo(ByVal NpcIndex As Integer)
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    Dim tHeading As Byte

    Dim ui       As Integer
    
    With Npclist(NpcIndex)

        If .Target = 0 And .TargetNPC = 0 Then
            ui = .MaestroUser
            
            If ui > 0 Then

                'Is it in it's range of vision??
                If Abs(UserList(ui).Pos.x - .Pos.x) <= RANGO_VISION_NPC_X Then
                    If Abs(UserList(ui).Pos.y - .Pos.y) <= RANGO_VISION_NPC_Y Then
                        If UserList(ui).flags.Muerto = 0 And UserList(ui).flags.invisible = 0 And UserList(ui).flags.Oculto = 0 And Distancia(.Pos, UserList(ui).Pos) > 3 Then
                            tHeading = FindDirection(.Pos, UserList(ui).Pos)
                            Call MoveNPCChar(NpcIndex, tHeading)
                            Exit Sub

                        End If

                    End If

                End If

            End If

        End If

    End With
    
    Call RestoreOldMovement(NpcIndex)

End Sub

Private Sub AiNpcAtacaNpc(ByVal NpcIndex As Integer)
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    Dim tHeading As Byte

    Dim x        As Long

    Dim y        As Long

    Dim NI       As Integer

    Dim bNoEsta  As Boolean
    
    Dim SignoNS  As Integer

    Dim SignoEO  As Integer
    
    With Npclist(NpcIndex)

        If .flags.Inmovilizado = 1 Then

            Select Case .Char.heading

                Case eHeading.NORTH
                    SignoNS = -1
                    SignoEO = 0
                
                Case eHeading.EAST
                    SignoNS = 0
                    SignoEO = 1
                
                Case eHeading.SOUTH
                    SignoNS = 1
                    SignoEO = 0
                
                Case eHeading.WEST
                    SignoEO = -1
                    SignoNS = 0

            End Select
            
            For y = .Pos.y To .Pos.y + SignoNS * RANGO_VISION_NPC_Y Step IIf(SignoNS = 0, 1, SignoNS)
                For x = .Pos.x To .Pos.x + SignoEO * RANGO_VISION_NPC_X Step IIf(SignoEO = 0, 1, SignoEO)

                    If x >= MinXBorder And x <= MaxXBorder And y >= MinYBorder And y <= MaxYBorder Then
                        NI = MapData(.Pos.map, x, y).NpcIndex

                        If NI > 0 Then
                            If .TargetNPC = NI Then
                                bNoEsta = True

                                If .Numero = ELEMENTALFUEGO Then
                                    Call NpcLanzaUnSpellSobreNpc(NpcIndex, NI)

                                    If Npclist(NI).NPCtype = DRAGON Then
                                        Call NpcLanzaUnSpellSobreNpc(NI, NpcIndex)

                                    End If

                                Else

                                    'aca verificamosss la distancia de ataque
                                    If Distancia(.Pos, Npclist(NI).Pos) <= 1 Then
                                        Call SistemaCombate.NpcAtacaNpc(NpcIndex, NI)

                                    End If

                                End If

                                Exit Sub

                            End If

                        End If

                    End If

                Next x
            Next y

        Else

            For y = .Pos.y - RANGO_VISION_NPC_Y To .Pos.y + RANGO_VISION_NPC_Y
                For x = .Pos.x - RANGO_VISION_NPC_Y To .Pos.x + RANGO_VISION_NPC_Y

                    If x >= MinXBorder And x <= MaxXBorder And y >= MinYBorder And y <= MaxYBorder Then
                        NI = MapData(.Pos.map, x, y).NpcIndex

                        If NI > 0 Then
                            If .TargetNPC = NI Then
                                bNoEsta = True

                                If .Numero = ELEMENTALFUEGO Then
                                    Call NpcLanzaUnSpellSobreNpc(NpcIndex, NI)

                                    If Npclist(NI).NPCtype = DRAGON Then
                                        Call NpcLanzaUnSpellSobreNpc(NI, NpcIndex)

                                    End If

                                Else

                                    'aca verificamosss la distancia de ataque
                                    If Distancia(.Pos, Npclist(NI).Pos) <= 1 Then
                                        Call SistemaCombate.NpcAtacaNpc(NpcIndex, NI)

                                    End If

                                End If

                                If .flags.Inmovilizado = 1 Then Exit Sub
                                If .TargetNPC = 0 Then Exit Sub
                                tHeading = FindDirection(.Pos, Npclist(MapData(.Pos.map, x, y).NpcIndex).Pos)
                                Call MoveNPCChar(NpcIndex, tHeading)
                                Exit Sub

                            End If

                        End If

                    End If

                Next x
            Next y

        End If
        
        If Not bNoEsta Then
            If .MaestroUser > 0 Then
                Call FollowAmo(NpcIndex)
            Else
                .Movement = .flags.OldMovement
                .Hostile = .flags.OldHostil

            End If

        End If

    End With

End Sub

Public Sub AiNpcObjeto(ByVal NpcIndex As Integer)

    '***************************************************
    'Autor: ZaMa
    'Last Modification: 14/09/2009 (ZaMa)
    '14/09/2009: ZaMa - Now npcs don't follow protected users.
    '***************************************************
    Dim Userindex     As Integer

    Dim i             As Long

    Dim UserProtected As Boolean
    
    With Npclist(NpcIndex)

        For i = 1 To Areas.ConnGroups(.Pos.map).Count()
            Userindex = Areas.ConnGroups(.Pos.map).Item(i)
            
            'Is it in it's range of vision??
            If Abs(UserList(Userindex).Pos.x - .Pos.x) <= RANGO_VISION_NPC_X Then
                If Abs(UserList(Userindex).Pos.y - .Pos.y) <= RANGO_VISION_NPC_Y Then
                    
                    With UserList(Userindex)
                        UserProtected = Not IntervaloPermiteSerAtacado(Userindex) And .flags.NoPuedeSerAtacado
                        
                        If .flags.Muerto = 0 And .flags.invisible = 0 And .flags.Oculto = 0 And .flags.AdminPerseguible And Not UserProtected Then
                            
                            ' No quiero que ataque siempre al primero
                            If RandomNumber(1, 3) < 3 Then
                                If Npclist(NpcIndex).flags.LanzaSpells > 0 Then
                                    Call NpcLanzaUnSpell(NpcIndex, Userindex)

                                End If
                            
                                Exit Sub

                            End If

                        End If

                    End With

                End If

            End If
            
        Next i

    End With

End Sub

Sub NPCAI(ByVal NpcIndex As Integer)

    '**************************************************************
    'Author: Unknown
    'Last Modify by: ZaMa
    'Last Modify Date: 15/11/2009
    '08/16/2008: MarKoxX - Now pets that do mel� attacks have to be near the enemy to attack.
    '15/11/2009: ZaMa - Implementacion de npc objetos ai.
    '**************************************************************
    On Error GoTo ErrorHandler

    With Npclist(NpcIndex)

        '<<<<<<<<<<< Ataques >>>>>>>>>>>>>>>>
        If .MaestroUser = 0 Then

            'Busca a alguien para atacar
            '�Es un guardia?
            If .NPCtype = eNPCType.GuardiaReal Then
                Call GuardiasAI(NpcIndex, False)
            ElseIf .NPCtype = eNPCType.Guardiascaos Then
                Call GuardiasAI(NpcIndex, True)
            ElseIf .Hostile And .Stats.Alineacion <> 0 Then
                Call HostilMalvadoAI(NpcIndex)
            ElseIf .Hostile And .Stats.Alineacion = 0 Then
                Call HostilBuenoAI(NpcIndex)

            End If

        Else

            'Evitamos que ataque a su amo, a menos
            'que el amo lo ataque.
            'Call HostilBuenoAI(NpcIndex)
        End If
        
        '<<<<<<<<<<<Movimiento>>>>>>>>>>>>>>>>
        Select Case .Movement

            Case TipoAI.MueveAlAzar

                If .flags.Inmovilizado = 1 Then Exit Sub
                If .NPCtype = eNPCType.GuardiaReal Then
                    If RandomNumber(1, 12) = 3 Then
                        Call MoveNPCChar(NpcIndex, CByte(RandomNumber(eHeading.NORTH, eHeading.WEST)))

                    End If
                    
                    Call PersigueCriminal(NpcIndex)
                    
                ElseIf .NPCtype = eNPCType.Guardiascaos Then

                    If RandomNumber(1, 12) = 3 Then
                        Call MoveNPCChar(NpcIndex, CByte(RandomNumber(eHeading.NORTH, eHeading.WEST)))

                    End If
                    
                    Call PersigueCiudadano(NpcIndex)
                    
                Else

                    If RandomNumber(1, 12) = 3 Then
                        Call MoveNPCChar(NpcIndex, CByte(RandomNumber(eHeading.NORTH, eHeading.WEST)))

                    End If

                End If
            
                'Va hacia el usuario cercano
            Case TipoAI.NpcMaloAtacaUsersBuenos
                Call IrUsuarioCercano(NpcIndex)
            
                'Va hacia el usuario que lo ataco(FOLLOW)
            Case TipoAI.NPCDEFENSA
                Call SeguirAgresor(NpcIndex)
            
                'Persigue criminales
            Case TipoAI.GuardiasAtacanCriminales
                Call PersigueCriminal(NpcIndex)
            
            Case TipoAI.SigueAmo

                If .flags.Inmovilizado = 1 Then Exit Sub
                Call SeguirAmo(NpcIndex)

                If RandomNumber(1, 12) = 3 Then
                    Call MoveNPCChar(NpcIndex, CByte(RandomNumber(eHeading.NORTH, eHeading.WEST)))

                End If
            
            Case TipoAI.NpcAtacaNpc
                Call AiNpcAtacaNpc(NpcIndex)
                
            Case TipoAI.NpcObjeto
                Call AiNpcObjeto(NpcIndex)
                
            Case TipoAI.NpcPathfinding

                If .flags.Inmovilizado = 1 Then Exit Sub
                If ReCalculatePath(NpcIndex) Then
                    Call PathFindingAI(NpcIndex)

                    'Existe el camino?
                    If .PFINFO.NoPath Then 'Si no existe nos movemos al azar
                        'Move randomly
                        Call MoveNPCChar(NpcIndex, RandomNumber(eHeading.NORTH, eHeading.WEST))

                    End If

                Else

                    If Not PathEnd(NpcIndex) Then
                        Call FollowPath(NpcIndex)
                    Else
                        .PFINFO.PathLenght = 0

                    End If

                End If

        End Select

    End With

    Exit Sub

ErrorHandler:

    With Npclist(NpcIndex)
        Call LogError("Error en NPCAI. Error: " & Err.Number & " - " & Err.description & ". " & "Npc: " & .Name & ", Index: " & NpcIndex & ", MaestroUser: " & .MaestroUser & ", MaestroNpc: " & .MaestroNpc & ", Mapa: " & .Pos.map & " x:" & .Pos.x & " y:" & .Pos.y & " Mov:" & .Movement & " TargU:" & .Target & " TargN:" & .TargetNPC)

    End With
    
    Dim MiNPC As npc

    MiNPC = Npclist(NpcIndex)
    Call QuitarNPC(NpcIndex)
    Call ReSpawnNpc(MiNPC)

End Sub

Function UserNear(ByVal NpcIndex As Integer) As Boolean
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    'Returns True if there is an user adjacent to the npc position.
    '***************************************************

    With Npclist(NpcIndex)
        UserNear = Not Int(Distance(.Pos.x, .Pos.y, UserList(.PFINFO.TargetUser).Pos.x, UserList(.PFINFO.TargetUser).Pos.y)) > 1

    End With

End Function

Function ReCalculatePath(ByVal NpcIndex As Integer) As Boolean
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    'Returns true if we have to seek a new path
    '***************************************************

    If Npclist(NpcIndex).PFINFO.PathLenght = 0 Then
        ReCalculatePath = True
    ElseIf Not UserNear(NpcIndex) And Npclist(NpcIndex).PFINFO.PathLenght = Npclist(NpcIndex).PFINFO.CurPos - 1 Then
        ReCalculatePath = True

    End If

End Function

Function PathEnd(ByVal NpcIndex As Integer) As Boolean
    '***************************************************
    'Author: Gulfas Morgolock
    'Last Modification: -
    'Returns if the npc has arrived to the end of its path
    '***************************************************
    PathEnd = Npclist(NpcIndex).PFINFO.CurPos = Npclist(NpcIndex).PFINFO.PathLenght

End Function

Function FollowPath(ByVal NpcIndex As Integer) As Boolean

    '***************************************************
    'Author: Gulfas Morgolock
    'Last Modification: -
    'Moves the npc.
    '***************************************************
    Dim tmpPos   As WorldPos

    Dim tHeading As Byte
    
    With Npclist(NpcIndex)
        tmpPos.map = .Pos.map
        tmpPos.x = .PFINFO.path(.PFINFO.CurPos).y ' invert� las coordenadas
        tmpPos.y = .PFINFO.path(.PFINFO.CurPos).x
        
        'Debug.Print "(" & tmpPos.X & "," & tmpPos.Y & ")"
        
        tHeading = FindDirection(.Pos, tmpPos)
        
        MoveNPCChar NpcIndex, tHeading
        
        .PFINFO.CurPos = .PFINFO.CurPos + 1

    End With

End Function

Function PathFindingAI(ByVal NpcIndex As Integer) As Boolean

    '***************************************************
    'Author: Gulfas Morgolock
    'Last Modification: -
    'This function seeks the shortest path from the Npc
    'to the user's location.
    '***************************************************
    Dim y As Long

    Dim x As Long
    
    With Npclist(NpcIndex)

        For y = .Pos.y - 10 To .Pos.y + 10    'Makes a loop that looks at
            For x = .Pos.x - 10 To .Pos.x + 10   '5 tiles in every direction
                
                'Make sure tile is legal
                If x > MinXBorder And x < MaxXBorder And y > MinYBorder And y < MaxYBorder Then
                    
                    'look for a user
                    If MapData(.Pos.map, x, y).Userindex > 0 Then

                        'Move towards user
                        Dim tmpUserIndex As Integer

                        tmpUserIndex = MapData(.Pos.map, x, y).Userindex

                        With UserList(tmpUserIndex)

                            If .flags.Muerto = 0 And .flags.invisible = 0 And .flags.Oculto = 0 And .flags.AdminPerseguible Then
                                'We have to invert the coordinates, this is because
                                'ORE refers to maps in converse way of my pathfinding
                                'routines.
                                Npclist(NpcIndex).PFINFO.Target.x = .Pos.y
                                Npclist(NpcIndex).PFINFO.Target.y = .Pos.x 'ops!
                                Npclist(NpcIndex).PFINFO.TargetUser = tmpUserIndex
                                Call SeekPath(NpcIndex)
                                Exit Function

                            End If

                        End With

                    End If

                End If

            Next x
        Next y

    End With

End Function

Sub NpcLanzaUnSpell(ByVal NpcIndex As Integer, ByVal Userindex As Integer)

    '**************************************************************
    'Author: Unknown
    'Last Modify by: -
    'Last Modify Date: -
    '**************************************************************
    With UserList(Userindex)

        If .flags.invisible = 1 Or .flags.Oculto = 1 Then Exit Sub

    End With
    
    Dim K As Integer

    K = RandomNumber(1, Npclist(NpcIndex).flags.LanzaSpells)
    Call NpcLanzaSpellSobreUser(NpcIndex, Userindex, Npclist(NpcIndex).Spells(K))

End Sub

Sub NpcLanzaUnSpellSobreNpc(ByVal NpcIndex As Integer, ByVal TargetNPC As Integer)
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    Dim K As Integer

    K = RandomNumber(1, Npclist(NpcIndex).flags.LanzaSpells)
    Call NpcLanzaSpellSobreNpc(NpcIndex, TargetNPC, Npclist(NpcIndex).Spells(K))

End Sub

Public Sub SacerdoteHealUser(ByVal Userindex As Integer)

    With UserList(Userindex)

        'Enviamos sonido de curar (Recox)
        Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessagePlayWave(SND_CURAR_SACERDOTE, .Pos.x, .Pos.y))

        .Stats.MinHp = .Stats.MaxHp

        Call WriteUpdateHP(Userindex)

        Call WriteConsoleMsg(Userindex, "El sacerdote te ha curado!!", FontTypeNames.FONTTYPE_INFO)

        'Si es newbie le sacamos todo, sino solo lo curamos. (Recox)
        If EsNewbie(Userindex) Then
            Call SacerdoteHealEffectsAndRestoreMana(Userindex)
        End If

        Call WriteUpdateUserStats(Userindex)
    End With

End Sub

Public Sub SacerdoteResucitateUser(ByVal Userindex As Integer)
    With UserList(Userindex)

        'Enviamos sonido de resucitacion (Recox)
        Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessagePlayWave(SND_RESUCITAR_SACERDOTE, .Pos.x, .Pos.y))
        
        Call RevivirUsuario(Userindex)
        Call WriteConsoleMsg(Userindex, "Has sido resucitado!!", FontTypeNames.FONTTYPE_INFO)

        'Si es newbie le sacamos todo, sino solo lo revivimos. (Recox)
        If EsNewbie(Userindex) Then
            Call SacerdoteHealEffectsAndRestoreMana(Userindex)
        End If

    End With
End Sub


Private Sub SacerdoteHealEffectsAndRestoreMana(ByVal Userindex As Integer)
    Dim MensajeAyuda As String
    MensajeAyuda = "Cuando dejes de ser newbie no lo hara mas el sacerdote y deberas comprar pociones o curarte con hechizos"

    With UserList(Userindex)
        ' Sacamos la maldicion.
        If .flags.Maldicion = 1 Then
            .flags.Maldicion = 0
            Call WriteConsoleMsg(Userindex, "El sacerdote te ha curado de la maldicion.", FontTypeNames.FONTTYPE_INFO)
            Call WriteConsoleMsg(Userindex, MensajeAyuda, FontTypeNames.FONTTYPE_INFO)
        End If
 
        ' Sacamos la ceguera.
        If .flags.Ceguera = 1 Then
            .flags.Ceguera = 0
            Call WriteConsoleMsg(Userindex, "El sacerdote te ha curado de la ceguera.", FontTypeNames.FONTTYPE_INFO)
            Call WriteConsoleMsg(Userindex, MensajeAyuda, FontTypeNames.FONTTYPE_INFO)
        End If

        ' Curamos su envenenamiento.
        If .flags.Envenenado = 1 Then
            .flags.Envenenado = 0
            Call WriteConsoleMsg(Userindex, "El sacerdote te ha curado del envenenamiento.", FontTypeNames.FONTTYPE_INFO)
            Call WriteConsoleMsg(Userindex, MensajeAyuda, FontTypeNames.FONTTYPE_INFO)
        End If

        ' Restauramos su mana.
        .Stats.MinMAN = .Stats.MaxMAN
        Call WriteUpdateMana(Userindex)
        Call WriteConsoleMsg(Userindex, "El sacerdote te ha restaurado el mana completamente.", FontTypeNames.FONTTYPE_INFO)
        Call WriteConsoleMsg(Userindex, MensajeAyuda, FontTypeNames.FONTTYPE_INFO)
    End With
End Sub
