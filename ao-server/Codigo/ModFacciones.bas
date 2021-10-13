Attribute VB_Name = "Mod_FaccionesOld"
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

Public Const NUM_RANGOS_FACCION       As Integer = 15

' Matriz que contiene las armaduras faccionarias segun raza, clase, faccion y defensa de armadura
'Public ArmadurasFaccion(1 To NUMCLASES, 1 To NUMRAZAS) As tFaccionArmaduras

' Contiene la cantidad de exp otorgada cada vez que aumenta el rango
Public RecompensaFacciones(NUM_RANGOS_FACCION)         As Long

Private Sub GiveFactionArmours(ByVal Userindex As Integer, ByVal IsCaos As Boolean)
    '***************************************************
    'Autor: ZaMa
    'Last Modification: 15/04/2010
    'Gives faction armours to user
    '***************************************************
    
    Dim miObj As obj

    Dim Rango     As Integer
    
    Dim i As Long
    Dim miFacc As Byte
    With UserList(Userindex)
        
        Rango = .Faccion.Jerarquia
        
        
        If IsCaos Then
            miFacc = eFaccion.Legion
        Else
            miFacc = eFaccion.Armada
        End If
        
        'primero le damos los premios de la jerarquia generales
    
        For i = 1 To Faccion(miFacc).Jerarquia(Rango).numPremios
            miObj.ObjIndex = Faccion(miFacc).Jerarquia(Rango).Premios(i).ObjIndex
            miObj.amount = Faccion(miFacc).Jerarquia(Rango).Premios(i).amount
            If Not MeterItemEnInventario(Userindex, miObj) Then
                Call TirarItemAlPiso(.Pos, miObj)
            End If
        Next i
        
        'ahora le damos los obj especifico de la clase y raza(altos o bajos)
        With Faccion(miFacc).Jerarquia(Rango).PremiosClase(.Clase)
            If UserList(Userindex).raza = eRaza.Enano Or UserList(Userindex).raza = eRaza.Gnomo Then
                For i = 0 To UBound(.ItemsBajos)
                    miObj.ObjIndex = .ItemsBajos(i).ObjIndex
                    miObj.ObjIndex = 1
                    If Not MeterItemEnInventario(Userindex, miObj) Then
                        Call TirarItemAlPiso(UserList(Userindex).Pos, miObj)
                    End If
                Next i
            Else
                For i = 0 To UBound(.ItemsAltos)
                    miObj.ObjIndex = .ItemsAltos(i).ObjIndex
                    miObj.ObjIndex = 1
                    If Not MeterItemEnInventario(Userindex, miObj) Then
                        Call TirarItemAlPiso(UserList(Userindex).Pos, miObj)
                    End If
                Next i
            End If
        End With
        
    End With

End Sub

Public Sub GiveExpReward(ByVal Userindex As Integer, ByVal Rango As Long)
    '***************************************************
    'Autor: ZaMa
    'Last Modification: 15/04/2010
    'Gives reward exp to user
    '***************************************************
    
    Dim GivenExp As Long
    
    With UserList(Userindex)
        
        GivenExp = RecompensaFacciones(Rango)
        
        .Stats.Exp = .Stats.Exp + GivenExp
        
        If .Stats.Exp > MAXEXP Then .Stats.Exp = MAXEXP
        
        Call WriteConsoleMsg(Userindex, "Has sido recompensado con " & GivenExp & " puntos de experiencia.", FontTypeNames.FONTTYPE_FIGHT)

        Call CheckUserLevel(Userindex)
        
    End With
    
End Sub

Public Sub EnlistarArmadaReal(ByVal Userindex As Integer)
    '***************************************************
    'Autor: Pablo (ToxicWaste) & Unknown (orginal version)
    'Last Modification: 15/04/2010
    'Handles the entrance of users to the "Armada Real"
    '15/03/2009: ZaMa - No se puede enlistar el fundador de un clan con alineacion neutral.
    '27/11/2009: ZaMa - Ahora no se puede enlistar un miembro de un clan neutro, por ende saque la antifaccion.
    '15/04/2010: ZaMa - Cambio en recompensas iniciales.
    '***************************************************

    With UserList(Userindex)

        If .Faccion.Bando = eFaccion.Armada Then
            Call WriteChatOverHead(Userindex, "Ya perteneces a las tropas reales!!! Ve a combatir criminales.", Str(Npclist(.flags.TargetNPC).Char.CharIndex), vbWhite)
            Exit Sub

        End If
    
        If .Faccion.Bando = eFaccion.Legion Then
            Call WriteChatOverHead(Userindex, "Maldito insolente!!! Vete de aqui seguidor de las sombras.", Str(Npclist(.flags.TargetNPC).Char.CharIndex), vbWhite)
            Exit Sub

        End If
    
        If .Faccion.Bando = eFaccion.Armada Then
            Call WriteChatOverHead(Userindex, "Debes convertirte en un ciudadano de la armada primero.", Str(Npclist(.flags.TargetNPC).Char.CharIndex), vbWhite)
            Exit Sub

        End If
    
        If .Faccion.CriminalesMatados < 30 Then
            Call WriteChatOverHead(Userindex, "Para unirte a nuestras fuerzas debes matar al menos 30 criminales, solo has matado " & .Faccion.CriminalesMatados & ".", Str(Npclist(.flags.TargetNPC).Char.CharIndex), vbWhite)
            Exit Sub

        End If
    
        If .Stats.ELV < 25 Then
            Call WriteChatOverHead(Userindex, "Para unirte a nuestras fuerzas debes ser al menos de nivel 25!!!", Str(Npclist(.flags.TargetNPC).Char.CharIndex), vbWhite)
            Exit Sub

        End If
     
        If .Faccion.CiudadanosMatados > 0 Then
            Call WriteChatOverHead(Userindex, "Has asesinado gente inocente, no aceptamos asesinos en las tropas reales!", Str(Npclist(.flags.TargetNPC).Char.CharIndex), vbWhite)
            Exit Sub

        End If
    
        If .GuildIndex > 0 Then
            If modGuilds.GuildAlignment(.GuildIndex) = "Neutral" Then
                Call WriteChatOverHead(Userindex, "Perteneces a un clan neutro, sal de el si quieres unirte a nuestras fuerzas!!!", Str(Npclist(.flags.TargetNPC).Char.CharIndex), vbWhite)
                Exit Sub

            End If

        End If
    
        .Faccion.Bando = eFaccion.Armada
        .Faccion.Jerarquia = 1

        Call WriteChatOverHead(Userindex, "Bienvenido al ejercito real!!! Aqui tienes tus vestimentas. Cumple bien tu labor exterminando criminales y me encargare de recompensarte.", Str(Npclist(.flags.TargetNPC).Char.CharIndex), vbWhite)
    
        ' TODO: Dejo esta variable por ahora, pero con chequear las reenlistadas deberia ser suficiente :S
      '  If .Faccion.RecibioArmaduraReal = 0 Then
        
            Call GiveFactionArmours(Userindex, False)
            Call GiveExpReward(Userindex, 0)
        
       ' End If
    
        If .flags.Navegando Then Call RefreshCharStatus(Userindex) 'Actualizamos la barca si esta navegando (NicoNZ)
    
        Call LogEjercitoReal(.Name & " ingreso el " & Date & " cuando era nivel " & .Stats.ELV)

    End With

End Sub

Public Sub RecompensaArmadaReal(ByVal Userindex As Integer)

    Dim Crimis    As Long

    Dim Lvl       As Byte

    Dim NextRecom As Long

    Dim Nobleza   As Long

    With UserList(Userindex)
        Lvl = .Stats.ELV
        Crimis = .Faccion.CriminalesMatados
        
        If Faccion(eFaccion.Armada).NumJerarquias = .Faccion.Jerarquia Then
            Call WriteChatOverHead(Userindex, "Eres uno de mis mejores soldados. Mataste " & Crimis & " criminales, sigue asi. Ya no tengo mas recompensa para darte que mi agradecimiento. Felicidades!", Str(Npclist(.flags.TargetNPC).Char.CharIndex), vbWhite)
                Exit Sub
        End If
        
        NextRecom = Faccion(eFaccion.Armada).Jerarquia(.Faccion.Jerarquia + 1).LegionesMatados ' .Faccion.NextRecompensa

        If Crimis < NextRecom Then
            Call WriteChatOverHead(Userindex, "Mata " & NextRecom - Crimis & " legiones más para pasar a la proxima jerarquía.", Str(Npclist(.flags.TargetNPC).Char.CharIndex), vbWhite)
            Exit Sub
            
        End If

        .Faccion.Jerarquia = .Faccion.Jerarquia + 1
        Call WriteChatOverHead(Userindex, "Aqui tienes tu recompensa " & TituloReal(Userindex) & "!!!", Str(Npclist(.flags.TargetNPC).Char.CharIndex), vbWhite)

        ' Recompensas de armaduras y exp
        Call GiveFactionArmours(Userindex, False)
        Call GiveExpReward(Userindex, .Faccion.Jerarquia)

    End With

End Sub

Public Function TituloReal(ByVal Userindex As Integer) As String
    TituloReal = Faccion(eFaccion.Armada).Jerarquia(UserList(Userindex).Faccion.Jerarquia).Titulo
End Function

Public Sub ExpulsarFaccionReal(ByVal Userindex As Integer, _
                               Optional Expulsado As Boolean = True)
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    ' 09/28/2010 C4b3z0n - Arreglado RT6 Overflow, el Desequipar() del escudo, ponia de parametro el ObjIndex del escudo en vez del EqpSlot.
    '***************************************************

    With UserList(Userindex)
        .Faccion.Bando = 0

        'Call PerderItemsFaccionarios(UserIndex)
        If Expulsado Then
            Call WriteConsoleMsg(Userindex, "Has sido expulsado del ejercito real!!!", FontTypeNames.FONTTYPE_FIGHT)
        Else
            Call WriteConsoleMsg(Userindex, "Te has retirado del ejercito real!!!", FontTypeNames.FONTTYPE_FIGHT)

        End If
    
        If .Invent.ArmourEqpObjIndex <> 0 Then

            'Desequipamos la armadura real si esta equipada
            If ObjData(.Invent.ArmourEqpObjIndex).Real = 1 Then Call Desequipar(Userindex, .Invent.ArmourEqpSlot)

        End If
    
        If .Invent.EscudoEqpObjIndex <> 0 Then

            'Desequipamos el escudo de caos si esta equipado
            If ObjData(.Invent.EscudoEqpObjIndex).Real = 1 Then Call Desequipar(Userindex, .Invent.EscudoEqpSlot)

        End If
    
        If .flags.Navegando Then Call RefreshCharStatus(Userindex) 'Actualizamos la barca si esta navegando (NicoNZ)

    End With

End Sub

Public Sub ExpulsarFaccionCaos(ByVal Userindex As Integer, _
                               Optional Expulsado As Boolean = True)
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    ' 09/28/2010 C4b3z0n - Arreglado RT6 Overflow, el Desequipar() del escudo, ponia de parametro el ObjIndex del escudo en vez del EqpSlot.
    '***************************************************

    With UserList(Userindex)
        .Faccion.Bando = 0

        'Call PerderItemsFaccionarios(UserIndex)
        If Expulsado Then
            Call WriteConsoleMsg(Userindex, "Has sido expulsado de la Legion Oscura!!!", FontTypeNames.FONTTYPE_FIGHT)
        Else
            Call WriteConsoleMsg(Userindex, "Te has retirado de la Legion Oscura!!!", FontTypeNames.FONTTYPE_FIGHT)

        End If
    
        If .Invent.ArmourEqpObjIndex <> 0 Then

            'Desequipamos la armadura de caos si esta equipada
            If ObjData(.Invent.ArmourEqpObjIndex).Caos = 1 Then Call Desequipar(Userindex, .Invent.ArmourEqpSlot)

        End If
    
        If .Invent.EscudoEqpObjIndex <> 0 Then

            'Desequipamos el escudo de caos si esta equipado
            If ObjData(.Invent.EscudoEqpObjIndex).Caos = 1 Then Call Desequipar(Userindex, .Invent.EscudoEqpSlot)

        End If
    
        If .flags.Navegando Then Call RefreshCharStatus(Userindex) 'Actualizamos la barca si esta navegando (NicoNZ)

    End With

End Sub

Public Sub EnlistarCaos(ByVal Userindex As Integer)
    With UserList(Userindex)

        If esNeutral(Userindex) Then
            Call WriteChatOverHead(Userindex, "Largate de aqui, bufon!!!", Str(Npclist(.flags.TargetNPC).Char.CharIndex), vbWhite)
            Exit Sub

        End If
    
        If .Faccion.Bando = eFaccion.Legion And .Faccion.Jerarquia > 0 Then
            Call WriteChatOverHead(Userindex, "Ya perteneces al ejercito de la legion oscura!!!", Str(Npclist(.flags.TargetNPC).Char.CharIndex), vbWhite)
            Exit Sub

        End If
    
        If .Faccion.Bando = eFaccion.Armada Then
            Call WriteChatOverHead(Userindex, "Las sombras reinaran en Argentum. Fuera de aqui insecto real!!!", Str(Npclist(.flags.TargetNPC).Char.CharIndex), vbWhite)
            Exit Sub

        End If
     
        If .Faccion.CiudadanosMatados < Faccion(eFaccion.Legion).Jerarquia(1).ArmadasMatados Then
            Call WriteChatOverHead(Userindex, "Para unirte a nuestras fuerzas debes matar al menos " & Faccion(eFaccion.Legion).Jerarquia(1).ArmadasMatados & " ciudadanos, solo has matado " & .Faccion.CiudadanosMatados & ".", Str(Npclist(.flags.TargetNPC).Char.CharIndex), vbWhite)
            Exit Sub

        End If
    
        If .Stats.ELV < 25 Then
            Call WriteChatOverHead(Userindex, "Para unirte a nuestras fuerzas debes ser al menos nivel 25!!!", Str(Npclist(.flags.TargetNPC).Char.CharIndex), vbWhite)
            Exit Sub

        End If
    
        .Faccion.Bando = eFaccion.Legion
        .Faccion.Jerarquia = 1
    
        Call WriteChatOverHead(Userindex, "Bienvenido a nuestro ejercito!!! Aqui tienes tus armaduras. Derrama sangre ciudadana y real, y seras recompensado, lo prometo.", Str(Npclist(.flags.TargetNPC).Char.CharIndex), vbWhite)
    
            Call GiveFactionArmours(Userindex, True)
            Call GiveExpReward(Userindex, 0)
    
        If .flags.Navegando Then Call RefreshCharStatus(Userindex) 'Actualizamos la barca si esta navegando (NicoNZ)

        Call LogEjercitoCaos(.Name & " ingreso el " & Date & " cuando era nivel " & .Stats.ELV)

    End With

End Sub

Public Sub RecompensaCaos(ByVal Userindex As Integer)
    Dim Ciudas    As Long

    Dim Lvl       As Byte

    Dim NextRecom As Long

    With UserList(Userindex)
        Lvl = .Stats.ELV
        Ciudas = .Faccion.CiudadanosMatados
        
        If Faccion(eFaccion.Legion).NumJerarquias = .Faccion.Jerarquia Then
            Call WriteChatOverHead(Userindex, "Eres uno de mis mejores soldados. Mataste " & Ciudas & " ciudadanos, sigue asi. Ya no tengo mas recompensa para darte que mi agradecimiento. Felicidades!", Str(Npclist(.flags.TargetNPC).Char.CharIndex), vbWhite)
            Exit Sub
        End If
        
        NextRecom = Faccion(eFaccion.Armada).Jerarquia(.Faccion.Jerarquia + 1).LegionesMatados ' .Faccion.NextRecompensa

        If Ciudas < NextRecom Then
            Call WriteChatOverHead(Userindex, "Mata " & NextRecom - Ciudas & " cuidadanos mas para recibir la proxima recompensa.", Str(Npclist(.flags.TargetNPC).Char.CharIndex), vbWhite)
            Exit Sub

        End If
        
        .Faccion.Jerarquia = .Faccion.Jerarquia + 1
    
        Call WriteChatOverHead(Userindex, "Bien hecho " & TituloCaos(Userindex) & ", aqui tienes tu recompensa!!!", Str(Npclist(.flags.TargetNPC).Char.CharIndex), vbWhite)
    
        ' Recompensas de armaduras y exp
        Call GiveFactionArmours(Userindex, True)
        Call GiveExpReward(Userindex, .Faccion.Jerarquia)
    
    End With

End Sub

Public Function TituloCaos(ByVal Userindex As Integer) As String
     TituloCaos = Faccion(eFaccion.Legion).Jerarquia(UserList(Userindex).Faccion.Jerarquia).Titulo
End Function

