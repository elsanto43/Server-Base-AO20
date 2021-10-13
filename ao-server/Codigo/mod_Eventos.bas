Attribute VB_Name = "mod_Eventos"
Option Explicit

Private Enum eMapasExpX2
    DungeonMaravel
    DungeonDragon
    DungeonKaka
    DungeonVoo
    DungeonTripto
End Enum 'Esto es solo
Private Type tMapEvent
    num As Integer
    Name As String
End Type
Public Const num_mapasexpX2 As Byte = 5

Public x2CurrMap As Byte

Public mapasx2(num_mapasexpX2 - 1) As tMapEvent

Private Sub GenerarEvento()
    If RandomNumber(1, 4) = 1 Then

            eventox2 = RandomNumber(1, 3)
            
            Select Case eventox2
                Case eEventoX2.Exp
                    x2CurrMap = RandomNumber(1, 5) - 1
                     Call MensajeGlobal("Experiencia x2> Todos los NPC's dan el doble de experiencia en: " & mapasx2(x2CurrMap).Name & "(Mapa " & mapasx2(x2CurrMap).num & "). Duración: 30 minutos", FontTypeNames.FONTTYPE_INFOBOLD)
                    
                Case eEventoX2.kills
                     Call MensajeGlobal("Kills x2> Ha comenzado el evento. Duración: 30 minutos. Los usuarios matados durante el evento contaran por 2.", FontTypeNames.FONTTYPE_INFOBOLD)
                     
                Case eEventoX2.oro
                     Call MensajeGlobal("Oro x2> Ha comenzado el evento. Duración: 30 minutos. Todos los NPC's dan el doble de oro.", FontTypeNames.FONTTYPE_INFOBOLD)
            
            End Select
        End If
End Sub

Public Sub PasaMinuto()
    Static minutos As Byte
    Static eventoT As Byte
    minutos = minutos + 1
    If minutos = 240 Then 'cada 4 horas se realiza un evento.
        Call GenerarEvento
        eventoT = 0
        minutos = 1
    End If
    
    
    If eventox2 <> 0 Then 'hay un evento en curso
            If eventoT < 30 Then
                    If (eventoT Mod 5 = 0) Then 'cada 5 minutos se avisa
                        Select Case eventox2
                            Case eEventoX2.Exp
                                 Call MensajeGlobal("Experiencia x2> Todos los NPC's dan el doble de experiencia en: " & mapasx2(x2CurrMap).Name & "(Mapa " & mapasx2(x2CurrMap).num & "). Quedan " & CStr(30 - eventoT) & " minutos.", FontTypeNames.FONTTYPE_INFOBOLD)
                                
                            Case eEventoX2.kills
                                 Call MensajeGlobal("Kills x2> Quedan " & CStr(30 - eventoT) & " minutos.", FontTypeNames.FONTTYPE_INFOBOLD)
                                 
                            Case eEventoX2.oro
                                 Call MensajeGlobal("Oro x2> Quedan " & CStr(30 - eventoT) & " minutos.", FontTypeNames.FONTTYPE_INFOBOLD)
                        
                        End Select
                    End If
                    If eventoT > 25 Then ' los ultimos 5 minutos avisa cada 1 minuto
                            Select Case eventox2
                                Case eEventoX2.Exp
                                     Call MensajeGlobal("Experiencia x2> Todos los NPC's dan el doble de experiencia en: " & mapasx2(x2CurrMap).Name & "(Mapa " & mapasx2(x2CurrMap).num & "). Quedan " & CStr(30 - eventoT) & " minutos.", FontTypeNames.FONTTYPE_INFOBOLD)
                                
                                Case eEventoX2.kills
                                     Call MensajeGlobal("Kills x2> Quedan " & CStr(30 - eventoT) & " minutos.", FontTypeNames.FONTTYPE_INFOBOLD)
                                     
                                Case eEventoX2.oro
                                     Call MensajeGlobal("Oro x2> Quedan " & CStr(30 - eventoT) & " minutos.", FontTypeNames.FONTTYPE_INFOBOLD)
                            
                            End Select
                    End If
            Else 'termina el evento
            
                    Select Case eventox2
                        Case eEventoX2.Exp
                             Call MensajeGlobal("Experiencia x2> El evento ha finalizado.", FontTypeNames.FONTTYPE_INFOBOLD)
                            
                        Case eEventoX2.kills
                             Call MensajeGlobal("Kills x2> El evento ha finalizado.", FontTypeNames.FONTTYPE_INFOBOLD)
                             
                        Case eEventoX2.oro
                             Call MensajeGlobal("Oro x2> El evento ha finalizado.", FontTypeNames.FONTTYPE_INFOBOLD)
                    
                    End Select
                    eventox2 = 0
                
            End If
        
    End If
    

        
End Sub

Public Sub CargarMapasExp()
    'aqui se configuran los mapas en los que se hara el evento especial de exp x2 cada 30 minutos en dungeons.
    mapasx2(eMapasExpX2.DungeonDragon).num = 251
    mapasx2(eMapasExpX2.DungeonMaravel).num = 55
    mapasx2(eMapasExpX2.DungeonKaka).num = 12
    mapasx2(eMapasExpX2.DungeonVoo).num = 77
    mapasx2(eMapasExpX2.DungeonTripto).num = 65
    
    mapasx2(eMapasExpX2.DungeonDragon).Name = "Dungeon Dragon"
    mapasx2(eMapasExpX2.DungeonMaravel).num = "Dungeon x"
    mapasx2(eMapasExpX2.DungeonKaka).num = "Dungeon x"
    mapasx2(eMapasExpX2.DungeonVoo).num = "Dungeon x"
    mapasx2(eMapasExpX2.DungeonTripto).num = "Dungeon x"
End Sub











