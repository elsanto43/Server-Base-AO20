Attribute VB_Name = "mod_Dats"
Option Explicit
Public MaxNPC As Integer
Public MaxNPCnohostiles As Integer
Private Type tConfigDats
    modo As Byte
    Index As Integer
End Type
Public configDats As tConfigDats


Public Sub InitializeDats()
        On Local Error Resume Next
        frmCargando.lblEstado.Caption = "Cargando OBJ.dat..."
        
            
           ' Call LoadOBJData
                    
        frmCargando.lblEstado.Caption = "Cargando NPCs-Hostiles.dat y NPCs.dat..."
        
           ' Call LoadAllNPC
            
        frmCargando.lblEstado.Caption = "Cargando Hechizos.dat..."
        
           ' Call CargarHechizos
            
        frmCargando.lblEstado.Caption = "Inicializando razas y clases..."
        
            ListaClases(1) = "Ciudadano"
            ListaClases(2) = "Trabajador"
            ListaClases(3) = "Experto en minerales"
            ListaClases(4) = "Minero"
            ListaClases(8) = "Herrero"
            ListaClases(13) = "Experto en uso de madera"
            ListaClases(14) = "Leñador"
            ListaClases(18) = "Carpintero"
            ListaClases(23) = "Pescador"
            ListaClases(27) = "Sastre"
            ListaClases(31) = "Alquimista"
            ListaClases(35) = "Luchador"
            ListaClases(36) = "Con uso de mana"
            ListaClases(37) = "Hechicero"
            ListaClases(38) = "Mago"
            ListaClases(39) = "Nigromante"
            ListaClases(40) = "Orden sagrada"
            ListaClases(41) = "Paladin"
            ListaClases(42) = "Clerigo"
            ListaClases(43) = "Naturalista"
            ListaClases(44) = "Bardo"
            ListaClases(45) = "Druida"
            ListaClases(46) = "Sigiloso"
            ListaClases(47) = "Asesino"
            ListaClases(48) = "Cazador"
            ListaClases(49) = "Sin uso de mana"
            ListaClases(50) = "Arquero"
            ListaClases(51) = "Guerrero"
            ListaClases(52) = "Caballero"
            ListaClases(53) = "Bandido"
            ListaClases(55) = "Pirata"
            ListaClases(56) = "Ladron"
            
            
                        
            ListaRazas(1) = "Humano"
            ListaRazas(2) = "Enano"
            ListaRazas(3) = "Elfo"
            ListaRazas(4) = "Elfo oscuro"
            ListaRazas(5) = "Gnomo"
            Unload frmCargando
End Sub

Sub LoadOBJData()

    

On Error GoTo errhandler
On Error GoTo 0



Dim Object As Integer

Dim nPath As String
Dim newdir As String

Dim Leer As clsIniReader

nPath = ConfigDir.Dats & "/Obj.dat"


If Not FileExist(nPath, vbNormal) Then MsgBox "No se ha encontrado obj.dat en el directorio configurado.", , "Error": reConfigurarPath = True: Exit Sub


Set Leer = New clsIniReader
Call Leer.Initialize(ConfigDir.Dats & "/Obj.dat")


numObjs = Leer.GetValue("INIT", "NumOBJs")

 'modProgressBar.Restart  numObjs
 modProgressBar.Restart numObjs
ReDim ObjData(0 To numObjs) As ObjData
      
  
For Object = 1 To numObjs
    ObjData(Object).name = Leer.GetValue("OBJ" & Object, "Name")

    'frmDats.List1.AddItem Object & " - " & ObjData(Object).name
    ObjData(Object).NoComerciable = Val(Leer.GetValue("OBJ" & Object, "NoComerciable"))
    
    ObjData(Object).GrhIndex = Val(Leer.GetValue("OBJ" & Object, "GrhIndex"))
    
    ObjData(Object).plusMagia = Val(Leer.GetValue("OBJ" & Object, "PlusMagia"))
    ObjData(Object).Def = Val(Leer.GetValue("OBJ" & Object, "Def"))
    
    ObjData(Object).NoSeCae = Val(Leer.GetValue("OBJ" & Object, "NoSeCae")) = 1
    ObjData(Object).objtype = Val(Leer.GetValue("OBJ" & Object, "ObjType"))
    ObjData(Object).ArbolElfico = Val(Leer.GetValue("OBJ" & Object, "ArbolElfico"))
    ObjData(Object).SubTipo = Val(Leer.GetValue("OBJ" & Object, "Subtipo"))
    ObjData(Object).Dosmanos = Val(Leer.GetValue("OBJ" & Object, "Dosmanos"))
    ObjData(Object).Newbie = Val(Leer.GetValue("OBJ" & Object, "Newbie"))
    ObjData(Object).itPart = Val(Leer.GetValue("OBJ" & Object, "itPart"))
    ObjData(Object).Aura = Val(Leer.GetValue("OBJ" & Object, "Aura"))
    ObjData(Object).SkPociones = Val(Leer.GetValue("OBJ" & Object, "SkPociones"))
    ObjData(Object).SkSastreria = Val(Leer.GetValue("OBJ" & Object, "SkSastreria"))
    ObjData(Object).Raices = Val(Leer.GetValue("OBJ" & Object, "Raices"))
    ObjData(Object).PielLobo = Val(Leer.GetValue("OBJ" & Object, "PielLobo"))
    ObjData(Object).PielOsoPardo = Val(Leer.GetValue("OBJ" & Object, "PielOsoPardo"))
    ObjData(Object).PielOsoPolar = Val(Leer.GetValue("OBJ" & Object, "PielOsoPolar"))
    ObjData(Object).Info = Leer.GetValue("OBJ" & Object, "Info")
        
    If ObjData(Object).SubTipo = OBJTYPE_ESCUDO Then
        ObjData(Object).ShieldAnim = Val(Leer.GetValue("OBJ" & Object, "Anim"))
        ObjData(Object).LingH = Val(Leer.GetValue("OBJ" & Object, "LingH"))
        ObjData(Object).LingP = Val(Leer.GetValue("OBJ" & Object, "LingP"))
        ObjData(Object).LingO = Val(Leer.GetValue("OBJ" & Object, "LingO"))

        ObjData(Object).SkHerreria = Val(Leer.GetValue("OBJ" & Object, "SkHerreria"))
    End If
    
    If ObjData(Object).SubTipo = OBJTYPE_CASCO Then
        ObjData(Object).CascoAnim = Val(Leer.GetValue("OBJ" & Object, "Anim"))
        ObjData(Object).LingH = Val(Leer.GetValue("OBJ" & Object, "LingH"))
        ObjData(Object).Gorro = Val(Leer.GetValue("OBJ" & Object, "Gorro"))
        ObjData(Object).LingP = Val(Leer.GetValue("OBJ" & Object, "LingP"))
        ObjData(Object).LingO = Val(Leer.GetValue("OBJ" & Object, "LingO"))
        ObjData(Object).SkHerreria = Val(Leer.GetValue("OBJ" & Object, "SkHerreria"))
    
    End If
    If ObjData(Object).objtype = OBJTYPE_MONTURA Then
        ObjData(Object).Caballo = Val(Leer.GetValue("OBJ" & Object, "Caballo"))
    End If
    ObjData(Object).Ropaje = Val(Leer.GetValue("OBJ" & Object, "NumRopaje"))
    ObjData(Object).HechizoIndex = Val(Leer.GetValue("OBJ" & Object, "HechizoIndex"))
    
    
    
        

    If ObjData(Object).objtype = OBJTYPE_WEAPON Then
        ObjData(Object).Baculo = Val(Leer.GetValue("OBJ" & Object, "Baculo"))
        ObjData(Object).WeaponAnim = Val(Leer.GetValue("OBJ" & Object, "Anim"))
        ObjData(Object).Apuñala = Val(Leer.GetValue("OBJ" & Object, "Apuñala"))
        ObjData(Object).Envenena = Val(Leer.GetValue("OBJ" & Object, "Envenena"))
        ObjData(Object).MaxHit = Val(Leer.GetValue("OBJ" & Object, "MaxHIT"))
        ObjData(Object).MinHit = Val(Leer.GetValue("OBJ" & Object, "MinHIT"))
        ObjData(Object).LingH = Val(Leer.GetValue("OBJ" & Object, "LingH"))
        ObjData(Object).LingP = Val(Leer.GetValue("OBJ" & Object, "LingP"))
        ObjData(Object).LingO = Val(Leer.GetValue("OBJ" & Object, "LingO"))
        ObjData(Object).SkHerreria = Val(Leer.GetValue("OBJ" & Object, "SkHerreria"))
        ObjData(Object).Real = Val(Leer.GetValue("OBJ" & Object, "Real"))
        ObjData(Object).Caos = Val(Leer.GetValue("OBJ" & Object, "Caos"))
        ObjData(Object).proyectil = Val(Leer.GetValue("OBJ" & Object, "Proyectil"))
        ObjData(Object).Municion = Val(Leer.GetValue("OBJ" & Object, "Municiones"))
        
    End If

    If ObjData(Object).objtype = OBJTYPE_ARMOUR Then
            ObjData(Object).LingH = Val(Leer.GetValue("OBJ" & Object, "LingH"))
            ObjData(Object).LingP = Val(Leer.GetValue("OBJ" & Object, "LingP"))
            ObjData(Object).LingO = Val(Leer.GetValue("OBJ" & Object, "LingO"))
            ObjData(Object).SkHerreria = Val(Leer.GetValue("OBJ" & Object, "SkHerreria"))
            ObjData(Object).Real = Val(Leer.GetValue("OBJ" & Object, "Real"))
            ObjData(Object).Caos = Val(Leer.GetValue("OBJ" & Object, "Caos"))
            ObjData(Object).Jerarquia = Val(Leer.GetValue("OBJ" & Object, "Jerarquia"))
    
    End If
    
    If ObjData(Object).objtype = OBJTYPE_HERRAMIENTAS Then
            ObjData(Object).LingH = Val(Leer.GetValue("OBJ" & Object, "LingH"))
            ObjData(Object).LingP = Val(Leer.GetValue("OBJ" & Object, "LingP"))
            ObjData(Object).LingO = Val(Leer.GetValue("OBJ" & Object, "LingO"))
            ObjData(Object).SkHerreria = Val(Leer.GetValue("OBJ" & Object, "SkHerreria"))
    End If
    
    If ObjData(Object).objtype = OBJTYPE_INSTRUMENTOS Then
        ObjData(Object).Snd1 = Val(Leer.GetValue("OBJ" & Object, "SND1"))
        ObjData(Object).Snd2 = Val(Leer.GetValue("OBJ" & Object, "SND2"))
        ObjData(Object).Snd3 = Val(Leer.GetValue("OBJ" & Object, "SND3"))
        ObjData(Object).MinInt = Val(Leer.GetValue("OBJ" & Object, "MinInt"))
    End If
    
    ObjData(Object).LingoteIndex = Val(Leer.GetValue("OBJ" & Object, "LingoteIndex"))
    
    If ObjData(Object).objtype = 31 Or ObjData(Object).objtype = 23 Then
        ObjData(Object).MinSkill = Val(Leer.GetValue("OBJ" & Object, "MinSkill"))
    End If
        
    ObjData(Object).MineralIndex = Val(Leer.GetValue("OBJ" & Object, "MineralIndex"))
    
    ObjData(Object).MaxHP = Val(Leer.GetValue("OBJ" & Object, "MaxHP"))
    ObjData(Object).MinHP = Val(Leer.GetValue("OBJ" & Object, "MinHP"))
    
    ObjData(Object).MUJER = Val(Leer.GetValue("OBJ" & Object, "Mujer"))
    ObjData(Object).HOMBRE = Val(Leer.GetValue("OBJ" & Object, "Hombre"))
    
    ObjData(Object).SkillCombate = Val(Leer.GetValue("OBJ" & Object, "SkCombate"))
    ObjData(Object).SkillTacticas = Val(Leer.GetValue("OBJ" & Object, "SkTacticas"))
    ObjData(Object).SkillProyectiles = Val(Leer.GetValue("OBJ" & Object, "SkProyectiles"))
    ObjData(Object).SkillApuñalar = Val(Leer.GetValue("OBJ" & Object, "SkApuñalar"))
    ObjData(Object).SkResistencia = Val(Leer.GetValue("OBJ" & Object, "SkResistencia"))
    ObjData(Object).SkDefensa = Val(Leer.GetValue("OBJ" & Object, "SkEscudos"))
    
    ObjData(Object).MinHam = Val(Leer.GetValue("OBJ" & Object, "MinHam"))
    ObjData(Object).MinSed = Val(Leer.GetValue("OBJ" & Object, "MinAgu"))
        
    ObjData(Object).MinDef = Val(Leer.GetValue("OBJ" & Object, "MINDEF"))
    ObjData(Object).MaxDef = Val(Leer.GetValue("OBJ" & Object, "MAXDEF"))

    ObjData(Object).Respawn = Val(Leer.GetValue("OBJ" & Object, "ReSpawn"))
    
    ObjData(Object).RazaEnana = Val(Leer.GetValue("OBJ" & Object, "RazaEnana"))
    
    ObjData(Object).Valor = Val(Leer.GetValue("OBJ" & Object, "Valor"))
    
    ObjData(Object).Crucial = Val(Leer.GetValue("OBJ" & Object, "Crucial"))
    
    ObjData(Object).Cerrada = Val(Leer.GetValue("OBJ" & Object, "abierta"))

    If ObjData(Object).Cerrada = 1 Then
            ObjData(Object).Llave = Val(Leer.GetValue("OBJ" & Object, "Llave"))
            ObjData(Object).Clave = Val(Leer.GetValue("OBJ" & Object, "Clave"))
    End If
    
    
    If ObjData(Object).objtype = OBJTYPE_PUERTAS Or ObjData(Object).objtype = OBJTYPE_BOTELLAVACIA Or ObjData(Object).objtype = OBJTYPE_BOTELLALLENA Then
        ObjData(Object).IndexAbierta = Val(Leer.GetValue("OBJ" & Object, "IndexAbierta"))
        ObjData(Object).IndexCerrada = Val(Leer.GetValue("OBJ" & Object, "IndexCerrada"))
        ObjData(Object).IndexCerradaLlave = Val(Leer.GetValue("OBJ" & Object, "IndexCerradaLlave"))
    End If
      If ObjData(Object).objtype = OBJTYPE_WARP Then
            ObjData(Object).WMapa = Val(Leer.GetValue("OBJ" & Object, "WMapa"))
            ObjData(Object).WX = Val(Leer.GetValue("OBJ" & Object, "WX"))
            ObjData(Object).WY = Val(Leer.GetValue("OBJ" & Object, "WY"))
            ObjData(Object).WI = Val(Leer.GetValue("OBJ" & Object, "WI"))
    End If
    
    ObjData(Object).Clave = Val(Leer.GetValue("OBJ" & Object, "Clave"))
    
    ObjData(Object).Texto = Leer.GetValue("OBJ" & Object, "Texto")
    ObjData(Object).GrhSecundario = Val(Leer.GetValue("OBJ" & Object, "VGrande"))
    
    ObjData(Object).Agarrable = Val(Leer.GetValue("OBJ" & Object, "Agarrable"))
    ObjData(Object).ForoID = Leer.GetValue("OBJ" & Object, "ID")
    
    
    Dim num As Integer
    
    num = Val(Leer.GetValue("OBJ" & Object, "NumClases"))
    
    Dim i As Integer
    For i = 1 To num
        ObjData(Object).ClaseProhibida(i) = Val(Leer.GetValue("OBJ" & Object, "CP" & i))
    Next
    
    num = Val(Leer.GetValue("OBJ" & Object, "NumRazas"))
     
    Dim d As Integer
    For d = 1 To num
        ObjData(Object).RazaProhibida(d) = Val(Leer.GetValue("OBJ" & Object, "RP" & d))
    Next
    
    ObjData(Object).Resistencia = Val(Leer.GetValue("OBJ" & Object, "Resistencia"))
    
    
    If ObjData(Object).objtype = 11 Then
        ObjData(Object).TipoPocion = Val(Leer.GetValue("OBJ" & Object, "TipoPocion"))
        ObjData(Object).MaxModificador = Val(Leer.GetValue("OBJ" & Object, "MaxModificador"))
        ObjData(Object).MinModificador = Val(Leer.GetValue("OBJ" & Object, "MinModificador"))
        ObjData(Object).DuracionEfecto = Val(Leer.GetValue("OBJ" & Object, "DuracionEfecto"))
    
    End If

    ObjData(Object).SkCarpinteria = Val(Leer.GetValue("OBJ" & Object, "SkCarpinteria"))
    
    If ObjData(Object).SkCarpinteria Then
        ObjData(Object).Madera = Val(Leer.GetValue("OBJ" & Object, "Madera"))
        ObjData(Object).MaderaElfica = Val(Leer.GetValue("OBJ" & Object, "MaderaElfica"))
    End If
    
    If ObjData(Object).objtype = OBJTYPE_BARCOS Then
            ObjData(Object).MaxHit = Val(Leer.GetValue("OBJ" & Object, "MaxHIT"))
            ObjData(Object).MinHit = Val(Leer.GetValue("OBJ" & Object, "MinHIT"))
    End If
    
    If ObjData(Object).objtype = OBJTYPE_FLECHAS Then
            ObjData(Object).MaxHit = Val(Leer.GetValue("OBJ" & Object, "MaxHIT"))
            ObjData(Object).MinHit = Val(Leer.GetValue("OBJ" & Object, "MinHIT"))
    End If
    
    ObjData(Object).MinSta = Val(Leer.GetValue("OBJ" & Object, "MinST"))

    modProgressBar.Update ProgressBar(0).Value + 1

Next Object

    Set Leer = Nothing
'Call INIDescarga(A)
'Call ExtraObjs

Exit Sub

errhandler:

'Call INIDescarga(A)

'Call LogErrorUrgente("Error cargando objetos: " & Err.Number & " : " & Err.Description)

End Sub





Sub LoadAllNPC()
    
    On Local Error Resume Next
Dim npcIndex As Long
Dim npcfile As String

Dim modoHostiles As Boolean
Dim Leer As clsIniReader

Dim nPath As String
nPath = ConfigDir.Dats & "/Npcs.dat"


If Not FileExist(nPath, vbNormal) Then MsgBox "No se ha encontrado NPCs.dat en el directorio configurado.", , "Error": reConfigurarPath = True: Exit Sub



Set Leer = New clsIniReader
Call Leer.Initialize(ConfigDir.Dats & "/NPCs.dat")

NumCuerpos = Val(Leer.GetValue("INIT", "NumBodies"))

MaxNPCnohostiles = Val(GetVar(ConfigDir.Dats & "/NPCs.dat", "INIT", "NumNPCs"))

MaxNPC = Val(GetVar(ConfigDir.Dats & "/NPCs-HOSTILES.dat", "INIT", "NumNPCs")) '+ 501


modProgressBar.Restart MaxNPC

ReDim Npclist(1 To MaxNPC + 501)
For npcIndex = 1 To MaxNPC + 501
    If npcIndex > 499 Then
        If modoHostiles = False Then
            modoHostiles = True
            Set Leer = Nothing
            DoEvents
            Set Leer = New clsIniReader
'            Leer.Initialize ConfigDir.Dats & "/NPCs-HOSTILES.dat"
        End If
    End If
    If npcIndex > MaxNPCnohostiles And npcIndex < MaxNPC Then
        npcIndex = 500
    End If
    'Debug.Print Val(GetVar(npcfile, "NPC" & NPCindex, "Body"))
    Npclist(npcIndex).Char.Body = Val(Leer.GetValue("NPC" & npcIndex, "Body"))
     Npclist(npcIndex).name = Leer.GetValue("NPC" & npcIndex, "Name")
    sNpcStat(eNPCStats.Body) = "Body"
    If Npclist(npcIndex).name <> "" And Npclist(npcIndex).Char.Body Then '   Val(Leer.GetValue( "NPC" & npcindex, "Body"))
       
        '
        Npclist(npcIndex).Desc = Leer.GetValue("NPC" & npcIndex, "Desc")
        sNpcStat(eNPCStats.Desc) = "Desc"
        
        Npclist(npcIndex).Movement = Val(Leer.GetValue("NPC" & npcIndex, "Movement"))
        sNpcStat(eNPCStats.Movement) = "Movement"

        Npclist(npcIndex).flags.AguaValida = Val(Leer.GetValue("NPC" & npcIndex, "AguaValida"))
        sNpcStat(eNPCStats.AguaValida) = "AguaValida"
        Npclist(npcIndex).flags.TierraInvalida = Val(Leer.GetValue("NPC" & npcIndex, "TierraInValida"))
        sNpcStat(eNPCStats.TierraInvalida) = "TierraInvalida"
        Npclist(npcIndex).flags.Faccion = Val(Leer.GetValue("NPC" & npcIndex, "Faccion"))
        sNpcStat(eNPCStats.Faccion) = "Faccion"
        
        Npclist(npcIndex).NPCtype = Val(Leer.GetValue("NPC" & npcIndex, "NpcType"))
        sNpcStat(eNPCStats.NPCtype) = "NpcType"
        
        
        Npclist(npcIndex).Char.Head = Val(Leer.GetValue("NPC" & npcIndex, "Head"))
        sNpcStat(eNPCStats.Head) = "Head"
        Npclist(npcIndex).Char.Heading = Val(Leer.GetValue("NPC" & npcIndex, "Heading"))
        sNpcStat(eNPCStats.Heading) = "Heading"
        Npclist(npcIndex).Attackable = Val(Leer.GetValue("NPC" & npcIndex, "Attackable"))
        sNpcStat(eNPCStats.Atacable) = "Attackable"
        Npclist(npcIndex).Comercia = Val(Leer.GetValue("NPC" & npcIndex, "Comercia"))
        sNpcStat(eNPCStats.Comercia) = "Comercia"
        Npclist(npcIndex).hostilE = Val(Leer.GetValue("NPC" & npcIndex, "Hostile"))
        sNpcStat(eNPCStats.hostil) = "Hostile"
        Npclist(npcIndex).InmuneParalisis = Val(Leer.GetValue("NPC" & npcIndex, "InmuneParalisis"))
        sNpcStat(eNPCStats.InmuneParalisis) = "InmuneParalisis"
      
        
        Npclist(npcIndex).MaxRecom = Val(Leer.GetValue("NPC" & npcIndex, "MaxRecom"))
        sNpcStat(eNPCStats.MaxRecompensa) = "MaxRecom"
        Npclist(npcIndex).MinRecom = Val(Leer.GetValue("NPC" & npcIndex, "MinRecom"))
        sNpcStat(eNPCStats.MinRecompensa) = "MinRecompensa"
        Npclist(npcIndex).Probabilidad = Val(Leer.GetValue("NPC" & npcIndex, "Probabilidad"))
        sNpcStat(eNPCStats.Probabilidad) = "Probabilidad"
        
        Npclist(npcIndex).GiveEXP = Val(Leer.GetValue("NPC" & npcIndex, "GiveEXP"))
        sNpcStat(eNPCStats.GiveEXP) = "GiveEXP"
        Npclist(npcIndex).Veneno = Val(Leer.GetValue("NPC" & npcIndex, "Veneno"))
        sNpcStat(eNPCStats.Veneno) = "Veneno"
        Npclist(npcIndex).flags.Domable = Val(Leer.GetValue("NPC" & npcIndex, "Domable"))
        sNpcStat(eNPCStats.Domable) = "Domable"
        
        Npclist(npcIndex).GiveGLD = Val(Leer.GetValue("NPC" & npcIndex, "GiveGLD"))
        sNpcStat(eNPCStats.GiveGLD) = "GiveGLD"
        Npclist(npcIndex).PoderAtaque = Val(Leer.GetValue("NPC" & npcIndex, "PoderAtaque"))
        sNpcStat(eNPCStats.PoderAtaque) = "PoderAtaque"
        Npclist(npcIndex).PoderEvasion = Val(Leer.GetValue("NPC" & npcIndex, "PoderEvasion"))
        sNpcStat(eNPCStats.PoderEvasion) = "PoderEvasion"
        
        Npclist(npcIndex).InvReSpawn = Val(Leer.GetValue("NPC" & npcIndex, "InvReSpawn"))
        sNpcStat(eNPCStats.InvReSpawn) = "InvReSpawn"
        Npclist(npcIndex).AutoCurar = Val(Leer.GetValue("NPC" & npcIndex, "autocurar"))
        sNpcStat(eNPCStats.AutoCurar) = "AutoCurar"
        
        Npclist(npcIndex).Stats.MaxHP = Val(Leer.GetValue("NPC" & npcIndex, "MaxHP"))
        sNpcStat(eNPCStats.MaxHP) = "MaxHP"
        Npclist(npcIndex).Stats.MinHP = Val(Leer.GetValue("NPC" & npcIndex, "MinHP"))
        sNpcStat(eNPCStats.MinHP) = "MinHP"
        Npclist(npcIndex).Stats.MaxHit = Val(Leer.GetValue("NPC" & npcIndex, "MaxHIT"))
        sNpcStat(eNPCStats.MaxHit) = "MaxHit"
        Npclist(npcIndex).Stats.MinHit = Val(Leer.GetValue("NPC" & npcIndex, "MinHIT"))
        sNpcStat(eNPCStats.MinHit) = "MinHit"
        Npclist(npcIndex).Stats.Def = Val(Leer.GetValue("NPC" & npcIndex, "DEF"))
        sNpcStat(eNPCStats.Def) = "DEF"
        Npclist(npcIndex).Stats.Alineacion = Val(Leer.GetValue("NPC" & npcIndex, "Alineacion"))
        sNpcStat(eNPCStats.Alineacion) = "Alineacion"
        Npclist(npcIndex).Stats.ImpactRate = Val(Leer.GetValue("NPC" & npcIndex, "ImpactRate"))
        sNpcStat(eNPCStats.ImpactRate) = "ImpactRate"
        
        Dim loopC As Integer
        Dim ln As String
        Npclist(npcIndex).Invent.NroItems = Val(Leer.GetValue("NPC" & npcIndex, "NROITEMS"))
        sNpcStat(eNPCStats.NroItems) = "NROITEMS"
        For loopC = 1 To Npclist(npcIndex).Invent.NroItems
            ln = Leer.GetValue("NPC" & npcIndex, "Obj" & loopC)
            Npclist(npcIndex).Invent.Object(loopC).OBJIndex = Val(ReadField(1, ln, 45))
            Npclist(npcIndex).Invent.Object(loopC).Amount = Val(ReadField(2, ln, 45))
        Next
        
        Npclist(npcIndex).flags.LanzaSpells = Val(Leer.GetValue("NPC" & npcIndex, "LanzaSpells"))
        sNpcStat(eNPCStats.LanzaSpells) = "LanzaSpells"
        If Npclist(npcIndex).flags.LanzaSpells Then ReDim Npclist(npcIndex).Spells(1 To Npclist(npcIndex).flags.LanzaSpells)
        For loopC = 1 To Npclist(npcIndex).flags.LanzaSpells
            Npclist(npcIndex).Spells(loopC) = Val(Leer.GetValue("NPC" & npcIndex, "Sp" & loopC))
        Next
        
        
       ' If Npclist(NpcIndex).NPCtype = NPCTYPE_ENTRENADOR Then
      '      Npclist(NpcIndex).NroCriaturas = Val(Leer.GetValue( "NPC" & NpcIndex, "NroCriaturas"))
     '       ReDim Npclist(NpcIndex).Criaturas(1 To Npclist(NpcIndex).NroCriaturas) As tCriaturasEntrenador
     '       For LoopC = 1 To Npclist(NpcIndex).NroCriaturas
      '          Npclist(NpcIndex).Criaturas(LoopC).NpcIndex =Leer.GetValue( "NPC" & NpcIndex, "CI" & LoopC)
     ''           Npclist(NpcIndex).Criaturas(LoopC).NpcName =Leer.GetValue( "NPC" & NpcIndex, "CN" & LoopC)
      '      Next
     '   End If
    
        
        Npclist(npcIndex).Inflacion = Val(Leer.GetValue("NPC" & npcIndex, "Inflacion"))
         sNpcStat(eNPCStats.Inflacion) = "Inflacion"
        
        'If Respawn Then
            Npclist(npcIndex).flags.Respawn = Val(Leer.GetValue("NPC" & npcIndex, "ReSpawn"))
             sNpcStat(eNPCStats.Respawn) = "ReSpawn"
        ''Else
        '   Npclist(NPCindex).flags.Respawn = 1
        'End If
        
        Npclist(npcIndex).flags.RespawnOrigPos = Val(Leer.GetValue("NPC" & npcIndex, "OrigPos"))
        sNpcStat(eNPCStats.RespawnOrigPos) = "OrigPos"
        Npclist(npcIndex).flags.AfectaParalisis = Val(Leer.GetValue("NPC" & npcIndex, "AfectaParalisis"))
        sNpcStat(eNPCStats.AfectaParalisis) = "AfectaParalisis"
        Npclist(npcIndex).flags.GolpeExacto = Val(Leer.GetValue("NPC" & npcIndex, "GolpeExacto"))
        sNpcStat(eNPCStats.GolpeExacto) = "GolpeExacto"
        Npclist(npcIndex).flags.PocaParalisis = Val(Leer.GetValue("NPC" & npcIndex, "PocaParalisis"))
        sNpcStat(eNPCStats.PocaParalisis) = "PocaParalisis"
        Npclist(npcIndex).VeInvis = Val(Leer.GetValue("NPC" & npcIndex, "veinvis"))
        sNpcStat(eNPCStats.VeInvis) = "VeInvis"
        
        
        Npclist(npcIndex).flags.Snd1 = Val(Leer.GetValue("NPC" & npcIndex, "Snd1"))
        sNpcStat(eNPCStats.Snd1) = "Snd1"
        Npclist(npcIndex).flags.Snd2 = Val(Leer.GetValue("NPC" & npcIndex, "Snd2"))
        sNpcStat(eNPCStats.Snd2) = "Snd2"
        Npclist(npcIndex).flags.Snd3 = Val(Leer.GetValue("NPC" & npcIndex, "Snd3"))
        sNpcStat(eNPCStats.Snd3) = "Snd3"
        Npclist(npcIndex).flags.Snd4 = Val(Leer.GetValue("NPC" & npcIndex, "Snd4"))
        sNpcStat(eNPCStats.Snd4) = "Snd4"
        Npclist(npcIndex).flags.Sound = Val(Leer.GetValue("NPC" & npcIndex, "sound"))
        sNpcStat(eNPCStats.Sound) = "sound"
        
        
        
        Npclist(npcIndex).TipoItems = Val(Leer.GetValue("NPC" & npcIndex, "TipoItems"))
        sNpcStat(eNPCStats.TipoItems) = "TipoItems"
        
        modProgressBar.Update ProgressBar(0).Value + 1

    
    End If 'body <> 0
Next npcIndex

'If NPCindex > LastNPC Then LastNPC = NPCindex
'NumNPCs = NumNPCs + 1

Set Leer = Nothing

'OpenNPC_Viejo = NPCindex
End Sub
    

Public Sub CargarHechizos()
On Error GoTo errhandler




If Not FileExist(ConfigDir.Dats & "/Hechizos.dat", vbNormal) Then MsgBox "No se ha encontrado Hechizos.dat en el directorio configurado.", , "Error": reConfigurarPath = True: Exit Sub


Dim Hechizo As Integer

NumeroHechizos = Val(GetVar(ConfigDir.Dats & "/Hechizos.dat", "INIT", "NumeroHechizos"))
ReDim Hechizos(1 To NumeroHechizos) As tHechizo
modProgressBar.Restart NumeroHechizos + 2
For Hechizo = 1 To NumeroHechizos
    modProgressBar.Update ProgressBar(0).Value + 1
    
    Hechizos(Hechizo).Nombre = GetVar(ConfigDir.Dats & "/Hechizos.dat", "Hechizo" & Hechizo, "Nombre")

    Hechizos(Hechizo).Desc = GetVar(ConfigDir.Dats & "/Hechizos.dat", "Hechizo" & Hechizo, "Desc")
    sHecStat(eHecStats.Desc) = "Desc"
    Hechizos(Hechizo).PalabrasMagicas = GetVar(ConfigDir.Dats & "/Hechizos.dat", "Hechizo" & Hechizo, "PalabrasMagicas")
    sHecStat(eHecStats.PalabrasMagicas) = "PalabrasMagicas"
    
    Hechizos(Hechizo).HechizeroMsg = GetVar(ConfigDir.Dats & "/Hechizos.dat", "Hechizo" & Hechizo, "HechizeroMsg")
    sHecStat(eHecStats.HechizeroMsg) = "HechizeroMsg"
    Hechizos(Hechizo).TargetMsg = GetVar(ConfigDir.Dats & "/Hechizos.dat", "Hechizo" & Hechizo, "TargetMsg")
    sHecStat(eHecStats.TargetMsg) = "TargetMsg"
    Hechizos(Hechizo).PropioMsg = GetVar(ConfigDir.Dats & "/Hechizos.dat", "Hechizo" & Hechizo, "PropioMsg")
    sHecStat(eHecStats.PropioMsg) = "PropioMsg"
    
    Hechizos(Hechizo).Tipo = Val(GetVar(ConfigDir.Dats & "/Hechizos.dat", "Hechizo" & Hechizo, "Tipo"))
    sHecStat(eHecStats.Tipo) = "Tipo"
    Hechizos(Hechizo).WAV = Val(GetVar(ConfigDir.Dats & "/Hechizos.dat", "Hechizo" & Hechizo, "WAV"))
    sHecStat(eHecStats.WAV) = "WAV"
    Hechizos(Hechizo).FXgrh = Val(GetVar(ConfigDir.Dats & "/Hechizos.dat", "Hechizo" & Hechizo, "Fxgrh"))
    sHecStat(eHecStats.FXgrh) = "FXgrh"
    
    Hechizos(Hechizo).Loops = Val(GetVar(ConfigDir.Dats & "/Hechizos.dat", "Hechizo" & Hechizo, "Loops"))
    sHecStat(eHecStats.Loops) = "Loops"
    
    Hechizos(Hechizo).Resis = Val(GetVar(ConfigDir.Dats & "/Hechizos.dat", "Hechizo" & Hechizo, "Resis"))
    sHecStat(eHecStats.Resis) = "Resis"
    Hechizos(Hechizo).Baculo = Val(GetVar(ConfigDir.Dats & "/Hechizos.dat", "Hechizo" & Hechizo, "Baculo"))
    sHecStat(eHecStats.Baculo) = "Baculo"
       
    Hechizos(Hechizo).SubeHP = Val(GetVar(ConfigDir.Dats & "/Hechizos.dat", "Hechizo" & Hechizo, "SubeHP"))
    sHecStat(eHecStats.SubeHP) = "SubeHP"
    Hechizos(Hechizo).MinHP = Val(GetVar(ConfigDir.Dats & "/Hechizos.dat", "Hechizo" & Hechizo, "MinHP"))
    sHecStat(eHecStats.MinHP) = "MinHP"
    Hechizos(Hechizo).MaxHP = Val(GetVar(ConfigDir.Dats & "/Hechizos.dat", "Hechizo" & Hechizo, "MaxHP"))
    sHecStat(eHecStats.MaxHP) = "MaxHP"
    
    Hechizos(Hechizo).SubeMana = Val(GetVar(ConfigDir.Dats & "/Hechizos.dat", "Hechizo" & Hechizo, "SubeMana"))
    sHecStat(eHecStats.SubeMana) = "SubeMana"
    Hechizos(Hechizo).MiMana = Val(GetVar(ConfigDir.Dats & "/Hechizos.dat", "Hechizo" & Hechizo, "MinMana"))
    sHecStat(eHecStats.MiMana) = "MinMana"
    Hechizos(Hechizo).MaMana = Val(GetVar(ConfigDir.Dats & "/Hechizos.dat", "Hechizo" & Hechizo, "MaxMana"))
    sHecStat(eHecStats.MaMana) = "MaxMana"
    
    Hechizos(Hechizo).SubeSta = Val(GetVar(ConfigDir.Dats & "/Hechizos.dat", "Hechizo" & Hechizo, "SubeSta"))
    sHecStat(eHecStats.SubeSta) = "SubeSta"
    Hechizos(Hechizo).MinSta = Val(GetVar(ConfigDir.Dats & "/Hechizos.dat", "Hechizo" & Hechizo, "MinSta"))
    sHecStat(eHecStats.MinSta) = "MinSta"
    Hechizos(Hechizo).MaxSta = Val(GetVar(ConfigDir.Dats & "/Hechizos.dat", "Hechizo" & Hechizo, "MaxSta"))
    sHecStat(eHecStats.MaxSta) = "MaxSta"
    
    Hechizos(Hechizo).SubeHam = Val(GetVar(ConfigDir.Dats & "/Hechizos.dat", "Hechizo" & Hechizo, "SubeHam"))
    sHecStat(eHecStats.SubeHam) = "SubeHam"
    Hechizos(Hechizo).MinHam = Val(GetVar(ConfigDir.Dats & "/Hechizos.dat", "Hechizo" & Hechizo, "MinHam"))
    sHecStat(eHecStats.MinHam) = "MinHam"
    Hechizos(Hechizo).MaxHam = Val(GetVar(ConfigDir.Dats & "/Hechizos.dat", "Hechizo" & Hechizo, "MaxHam"))
    sHecStat(eHecStats.MaxHam) = "MaxHam"
    
    Hechizos(Hechizo).SubeSed = Val(GetVar(ConfigDir.Dats & "/Hechizos.dat", "Hechizo" & Hechizo, "SubeSed"))
    sHecStat(eHecStats.SubeSed) = "SubeSed"
    Hechizos(Hechizo).MinSed = Val(GetVar(ConfigDir.Dats & "/Hechizos.dat", "Hechizo" & Hechizo, "MinSed"))
    sHecStat(eHecStats.MinSed) = "MinSed"
    Hechizos(Hechizo).MaxSed = Val(GetVar(ConfigDir.Dats & "/Hechizos.dat", "Hechizo" & Hechizo, "MaxSed"))
    sHecStat(eHecStats.MaxSed) = "MaxSed"
    
    Hechizos(Hechizo).SubeAgilidad = Val(GetVar(ConfigDir.Dats & "/Hechizos.dat", "Hechizo" & Hechizo, "SubeAG"))
    sHecStat(eHecStats.SubeAgilidad) = "SubeAG"
    Hechizos(Hechizo).MinAgilidad = Val(GetVar(ConfigDir.Dats & "/Hechizos.dat", "Hechizo" & Hechizo, "MinAG"))
    sHecStat(eHecStats.MinAgilidad) = "MinAG"
    Hechizos(Hechizo).MaxAgilidad = Val(GetVar(ConfigDir.Dats & "/Hechizos.dat", "Hechizo" & Hechizo, "MaxAG"))
    sHecStat(eHecStats.MaxAgilidad) = "MaxAG"
    
    Hechizos(Hechizo).SubeFuerza = Val(GetVar(ConfigDir.Dats & "/Hechizos.dat", "Hechizo" & Hechizo, "SubeFU"))
    sHecStat(eHecStats.SubeFuerza) = "SubeFU"
    Hechizos(Hechizo).MinFuerza = Val(GetVar(ConfigDir.Dats & "/Hechizos.dat", "Hechizo" & Hechizo, "MinFU"))
    sHecStat(eHecStats.MinFuerza) = "MinFU"
    Hechizos(Hechizo).MaxFuerza = Val(GetVar(ConfigDir.Dats & "/Hechizos.dat", "Hechizo" & Hechizo, "MaxFU"))
    sHecStat(eHecStats.MaxFuerza) = "MaxFU"
    
    Hechizos(Hechizo).SubeCarisma = Val(GetVar(ConfigDir.Dats & "/Hechizos.dat", "Hechizo" & Hechizo, "SubeCA"))
    sHecStat(eHecStats.SubeCarisma) = "SubeCA"
    Hechizos(Hechizo).MinCarisma = Val(GetVar(ConfigDir.Dats & "/Hechizos.dat", "Hechizo" & Hechizo, "MinCA"))
    sHecStat(eHecStats.MinCarisma) = "MinCA"
    Hechizos(Hechizo).MaxCarisma = Val(GetVar(ConfigDir.Dats & "/Hechizos.dat", "Hechizo" & Hechizo, "MaxCA"))
    sHecStat(eHecStats.MaxCarisma) = "MaxCA"
    
    Hechizos(Hechizo).Invisibilidad = Val(GetVar(ConfigDir.Dats & "/Hechizos.dat", "Hechizo" & Hechizo, "Invisibilidad"))
    sHecStat(eHecStats.Invisibilidad) = "Invisibilidad"
    Hechizos(Hechizo).Paraliza = Val(GetVar(ConfigDir.Dats & "/Hechizos.dat", "Hechizo" & Hechizo, "Paraliza"))
    sHecStat(eHecStats.Paraliza) = "Paraliza"
    
    Hechizos(Hechizo).Transforma = Val(GetVar(ConfigDir.Dats & "/Hechizos.dat", "Hechizo" & Hechizo, "Transforma"))
    sHecStat(eHecStats.Transforma) = "Transforma"
    Hechizos(Hechizo).Envenena = Val(GetVar(ConfigDir.Dats & "/Hechizos.dat", "Hechizo" & Hechizo, "Envenena"))
    sHecStat(eHecStats.Envenena) = "Envenena"
    Hechizos(Hechizo).Ceguera = Val(GetVar(ConfigDir.Dats & "/Hechizos.dat", "Hechizo" & Hechizo, "Ceguera"))
    sHecStat(eHecStats.Ceguera) = "Ceguera"
    Hechizos(Hechizo).Estupidez = Val(GetVar(ConfigDir.Dats & "/Hechizos.dat", "Hechizo" & Hechizo, "Estupidez"))
    sHecStat(eHecStats.Estupidez) = "Estupidez"

    Hechizos(Hechizo).Revivir = Val(GetVar(ConfigDir.Dats & "/Hechizos.dat", "Hechizo" & Hechizo, "Revivir"))
    sHecStat(eHecStats.Revivir) = "Revivir"
    Hechizos(Hechizo).Flecha = Val(GetVar(ConfigDir.Dats & "/Hechizos.dat", "Hechizo" & Hechizo, "Flecha"))
    sHecStat(eHecStats.Flecha) = "Flecha"
    
    Hechizos(Hechizo).Metamorfosis = Val(GetVar(ConfigDir.Dats & "/Hechizos.dat", "Hechizo" & Hechizo, "Metamorfosis"))
    sHecStat(eHecStats.Metamorfosis) = "Metamorfosis"
    Hechizos(Hechizo).Maldicion = Val(GetVar(ConfigDir.Dats & "/Hechizos.dat", "Hechizo" & Hechizo, "Maldicion"))
    sHecStat(eHecStats.Maldicion) = "Maldicion"
    Hechizos(Hechizo).Bendicion = Val(GetVar(ConfigDir.Dats & "/Hechizos.dat", "Hechizo" & Hechizo, "Bendicion"))
    sHecStat(eHecStats.Bendicion) = "Bendicion"
 
    Hechizos(Hechizo).RemoverParalisis = Val(GetVar(ConfigDir.Dats & "/Hechizos.dat", "Hechizo" & Hechizo, "RemoverParalisis"))
    sHecStat(eHecStats.RemoverParalisis) = "RemoverParalisis"
    Hechizos(Hechizo).CuraVeneno = Val(GetVar(ConfigDir.Dats & "/Hechizos.dat", "Hechizo" & Hechizo, "CuraVeneno"))
    sHecStat(eHecStats.CuraVeneno) = "CuraVeneno"
    Hechizos(Hechizo).RemoverMaldicion = Val(GetVar(ConfigDir.Dats & "/Hechizos.dat", "Hechizo" & Hechizo, "RemoverMaldicion"))
    sHecStat(eHecStats.RemoverMaldicion) = "RemoverMaldicion"
    
    Hechizos(Hechizo).Invoca = Val(GetVar(ConfigDir.Dats & "/Hechizos.dat", "Hechizo" & Hechizo, "Invoca"))
    sHecStat(eHecStats.Invoca) = "Invoca"
    Hechizos(Hechizo).NumNPC = Val(GetVar(ConfigDir.Dats & "/Hechizos.dat", "Hechizo" & Hechizo, "NumNpc"))
    sHecStat(eHecStats.NumNPC) = "NumNPC"
    Hechizos(Hechizo).cant = Val(GetVar(ConfigDir.Dats & "/Hechizos.dat", "Hechizo" & Hechizo, "Cant"))
    sHecStat(eHecStats.cant) = "Cant"
    
    Hechizos(Hechizo).Materializa = Val(GetVar(ConfigDir.Dats & "/Hechizos.dat", "Hechizo" & Hechizo, "Materializa"))
    sHecStat(eHecStats.Materializa) = "Materializa"
    Hechizos(Hechizo).CuraArea = Val(GetVar(ConfigDir.Dats & "/Hechizos.dat", "Hechizo" & Hechizo, "CuraArea"))
    sHecStat(eHecStats.CuraArea) = "CuraArea"
    Hechizos(Hechizo).ItemIndex = Val(GetVar(ConfigDir.Dats & "/Hechizos.dat", "Hechizo" & Hechizo, "ItemIndex"))
    sHecStat(eHecStats.ItemIndex) = "ItemIndex"
    
    Hechizos(Hechizo).Nivel = Val(GetVar(ConfigDir.Dats & "/Hechizos.dat", "Hechizo" & Hechizo, "Nivel"))
    sHecStat(eHecStats.Nivel) = "Nivel"
    Hechizos(Hechizo).MinSkill = Val(GetVar(ConfigDir.Dats & "/Hechizos.dat", "Hechizo" & Hechizo, "MinSkill"))
    sHecStat(eHecStats.MinSkill) = "MinSkill"
    Hechizos(Hechizo).ManaRequerido = Val(GetVar(ConfigDir.Dats & "/Hechizos.dat", "Hechizo" & Hechizo, "ManaRequerido"))
    sHecStat(eHecStats.ManaRequerido) = "ManaRequerido"
    Hechizos(Hechizo).StaRequerido = Val(GetVar(ConfigDir.Dats & "/Hechizos.dat", "Hechizo" & Hechizo, "StaRequerido"))
    sHecStat(eHecStats.StaRequerido) = "StaRequerido"
    
    Hechizos(Hechizo).Target = Val(GetVar(ConfigDir.Dats & "/Hechizos.dat", "Hechizo" & Hechizo, "Target"))
    sHecStat(eHecStats.Target) = "Target"

Next

Exit Sub

errhandler:
        
End Sub



Public Sub GrabarDatos(ByVal Index As Integer, ByVal TxtIndex As Byte)
'Esta funcion pasa el index del txtDatos para ser guardado en las respectivas variables correspondientes al index

    
    Dim nINdex As Integer
    Dim Value As String
    Value = frmDats.txtDatos(TxtIndex).Text
    nINdex = TxtIndex + 1
    
    
    Select Case estadoDat
        Case eModo.Objetos
            With ObjData(Index)
                Select Case nINdex
                    Case eObjStats.Agarrable
                        .Agarrable = Val(Value)
                    
                    Case eObjStats.Apuñala
                        .Apuñala = Val(Value)
                        
                    Case eObjStats.Arbol_Elfico
                        .ArbolElfico = Val(Value)
                                               
                    Case eObjStats.Armaanim
                        .WeaponAnim = Val(Value)
                                               
                    Case eObjStats.Aura
                        .Aura = Val(Value)
                                               
                    Case eObjStats.Baculo
                        .Baculo = Val(Value)
                                               
                    Case eObjStats.Caballo
                        .Caballo = Val(Value)
                                               
                    Case eObjStats.CantItems
                        .MaxItems = Val(Value)
                                               
                    Case eObjStats.Caos
                        .Caos = Val(Value)
                                               
                    Case eObjStats.CascoAnim
                        .CascoAnim = Val(Value)
                                               
                    Case eObjStats.Cerrada
                        .Cerrada = Val(Value)
                                               
                    Case eObjStats.Clasesprohib 'PENDIENTEEEEEEEEEEEEEEEEEEEEEEEEEEE
                        '.ArbolElfico = Val(value)
                                               
                    Case eObjStats.Clave
                        .Clave = Val(Value)
                                               
                    Case eObjStats.Crucial
                        .Crucial = Val(Value)
                                               
                    Case eObjStats.Def
                        .Def = Val(Value)
                                               
                    Case eObjStats.defensa
                        .SkDefensa = Val(Value)
                                               
                    Case eObjStats.Dosmanos
                        .Dosmanos = Val(Value)
                                               
                    Case eObjStats.DuracionEfecto
                        .DuracionEfecto = Val(Value)
                                               
                    Case eObjStats.Envenena
                        .Envenena = Val(Value)
                                               
                    Case eObjStats.Escudoanim
                        .ShieldAnim = Val(Value)
                                               
                    Case eObjStats.ForoID
                        .ForoID = Value
                                               
                    Case eObjStats.Gorro
                        .Gorro = Val(Value)
                                               
                    Case eObjStats.GrhIndex
                        .GrhIndex = Val(Value)
                                               
                    Case eObjStats.GrhSecundario
                        .GrhSecundario = Val(Value)
                       
                                            
                    Case eObjStats.HechizoIndex
                        .HechizoIndex = Val(Value)
                                               
                    Case eObjStats.HOMBRE
                        .HOMBRE = Val(Value)
                                               
                    Case eObjStats.IndexAbierta
                        .IndexAbierta = Val(Value)
                                               
                    Case eObjStats.IndexCerrada
                        .IndexCerrada = Val(Value)
                                               
                    Case eObjStats.IndexCerradaLlave
                        .IndexCerradaLlave = Val(Value)
                                               
                    Case eObjStats.Info
                        .Info = Value
                                              
                    Case eObjStats.Jerarquia
                        .Jerarquia = Val(Value)
                                               
                    Case eObjStats.LingoteIndex
                        .LingoteIndex = Val(Value)
                                               
                    Case eObjStats.LingotesHierro
                        .LingH = Val(Value)
                                               
                    Case eObjStats.LingotesOro
                        .LingO = Val(Value)
                                               
                    Case eObjStats.LingotesPlata
                        .LingP = Val(Value)
                                               
                    Case eObjStats.Llave
                        .Llave = Val(Value)
                                               
                    Case eObjStats.Madera
                        .Madera = Val(Value)
                                               
                    Case eObjStats.Madera_Elfica
                        .MaderaElfica = Val(Value)
                                               
                    Case eObjStats.Map_Pasajes
                        .WMapa = Val(Value)
                                               
                    Case eObjStats.MapX
                        .WX = Val(Value)
                                               
                    Case eObjStats.MapY
                        .WY = Val(Value)
                                               
                    Case eObjStats.MaxDef
                        .MaxDef = Val(Value)
                                               
                    Case eObjStats.MaxHit
                        .MaxHit = Val(Value)
                                               
                    Case eObjStats.MaxHP
                        .MaxHP = Val(Value)
                                               
                    Case eObjStats.MaxItems
                        .MaxItems = Val(Value)
                                               
                    Case eObjStats.MaxModificador
                        .MaxModificador = Val(Value)
                                               
                    Case eObjStats.MinDef
                        .MinDef = Val(Value)
                                               
                    Case eObjStats.MineralIndex
                        .MineralIndex = Val(Value)
                                               
                    Case eObjStats.Minhambre
                        .MinHam = Val(Value)
                                               
                    Case eObjStats.MinHit
                        .MinHit = Val(Value)
                                               
                    Case eObjStats.MinHP
                        .MinHP = Val(Value)
                                               
                    Case eObjStats.MinInt
                        .MinInt = Val(Value)
                            
                    Case eObjStats.MinModificador
                        .MinModificador = Val(Value)
                                               
                    Case eObjStats.MinSed
                        .MinSed = Val(Value)
                                               
                    Case eObjStats.MinSkill
                        .MinSkill = Val(Value)
                                               
                    Case eObjStats.MinSta
                        .MinSta = Val(Value)
                                               
                    Case eObjStats.MUJER
                        .MUJER = Val(Value)
                                               
                    Case eObjStats.Municion
                        .Municion = Val(Value)
                                               
                    Case eObjStats.Newbie
                        .Newbie = Val(Value)
                                               
                    Case eObjStats.NoComerciable
                        .NoComerciable = Val(Value)
                                               
                    Case eObjStats.NoSeCae
                        .NoSeCae = Val(Value)
                                               
                    Case eObjStats.PielLobo
                        .PielLobo = Val(Value)
                                               
                    Case eObjStats.PielOso
                        .PielOsoPardo = Val(Value)
                                               
                    Case eObjStats.PielOsoPolar
                        .PielOsoPardo = Val(Value)
                                               
                    Case eObjStats.plusMagia
                        .plusMagia = Val(Value)
                                               
                    Case eObjStats.proyectil
                        .proyectil = Val(Value)
                                               
                    Case eObjStats.Raices
                        .Raices = Val(Value)
                                               
                    Case eObjStats.RazaEnana
                        .RazaEnana = Val(Value)
                                               
                    Case eObjStats.Razasprohib 'AAAAAAAAAAAAAAAAAAAAAAAAAA
                        '.ArbolElfico = Val(value)
                                               
                    Case eObjStats.Real
                        .Real = Val(Value)
                                               
                    Case eObjStats.Remort
                        .Remort = Val(Value)
                                               
                    Case eObjStats.Resistencia
                        .Resistencia = Val(Value)
                                               
                    Case eObjStats.Respawn
                        .Respawn = Val(Value)
                                               
                    Case eObjStats.Ropaje
                        .Ropaje = Val(Value)
                                               
                    Case eObjStats.SkillApuñalar
                        .SkillApuñalar = Val(Value)
                                               
                    Case eObjStats.SkillCarpinteria
                        .SkCarpinteria = Val(Value)
                                               
                    Case eObjStats.SkillCombate
                        .SkillCombate = Val(Value)
                                               
                    Case eObjStats.Skilldefensa
                        .SkDefensa = Val(Value)
                                               
                    Case eObjStats.SkillHerreria
                        .SkHerreria = Val(Value)
                                               
                    Case eObjStats.Skillpociones
                        .SkPociones = Val(Value)
                                               
                    Case eObjStats.SkillProyectiles
                        .SkillProyectiles = Val(Value)
                                               
                    Case eObjStats.SkillResistencia
                        .SkResistencia = Val(Value)
                                               
                    Case eObjStats.Skillsastreria
                        .SkSastreria = Val(Value)
                                               
                    Case eObjStats.SkillTacticas
                        .SkillTacticas = Val(Value)
                                               
                    Case eObjStats.Snd1
                        .Snd1 = Val(Value)
                                               
                    Case eObjStats.Snd2
                        .Snd2 = Val(Value)
                                               
                    Case eObjStats.Snd3
                        .Snd3 = Val(Value)
                                               
                    Case eObjStats.SubTipo
                        .SubTipo = Val(ReadField(1, frmDats.cmbSubtipo.List(frmDats.cmbSubtipo.listIndex), Asc(" ")))
                        
                    Case eObjStats.objtype
                        .objtype = Val(ReadField(1, frmDats.cmbObjType.List(frmDats.cmbObjType.listIndex), Asc(" ")))
                                               
                    Case eObjStats.Texto
                        .Texto = (Value)
                                               
                    Case eObjStats.TipoPocion
                        .TipoPocion = Val(Value)
                                               
                    Case eObjStats.Valor
                        .Valor = Val(Value)
                                               
                    Case eObjStats.WarpI
                        .WI = Val(Value)
                       
                End Select
            End With
            
        Case eModo.Npc
            With Npclist(Index)
                Select Case nINdex
                    Case eNPCStats.AfectaParalisis
                        .flags.AfectaParalisis = Val(Value)
                    
                    Case eNPCStats.AguaValida
                        .flags.AguaValida = Val(Value)
                    
                    Case eNPCStats.Alineacion
                        .Stats.Alineacion = Val(Value)
                    
                    Case eNPCStats.Atacable
                        .Attackable = Val(Value)
                    
                    Case eNPCStats.AutoCurar
                        .AutoCurar = Val(Value)
                    
                    Case eNPCStats.Body
                        .Char.Body = Val(Value)
                    
                    Case eNPCStats.Comercia
                        .Comercia = Val(Value)
                    
                    Case eNPCStats.Def
                        .Stats.Def = Val(Value)
                    
                    Case eNPCStats.Desc
                        .Desc = (Value)
                    
                    Case eNPCStats.Domable
                        .flags.Domable = Val(Value)
                    
                    Case eNPCStats.Faccion
                        .flags.Faccion = Val(Value)
                    
                    Case eNPCStats.GiveEXP
                        .GiveEXP = Val(Value)
                    
                    Case eNPCStats.GiveGLD
                        .GiveGLD = Val(Value)
                    
                    Case eNPCStats.GolpeExacto
                        .flags.GolpeExacto = Val(Value)
                    
                    Case eNPCStats.Head
                        .Char.Head = Val(Value)
                    
                    Case eNPCStats.Heading
                        .Char.Heading = Val(Value)
                    
                    Case eNPCStats.hostil
                        .hostilE = Val(Value)
                    
                    Case eNPCStats.ImpactRate
                        .Stats.ImpactRate = Val(Value)
                    
                    Case eNPCStats.Inflacion
                        .Inflacion = Val(Value)
                    
                    Case eNPCStats.InmuneParalisis
                        .InmuneParalisis = Val(Value)
                    
                    Case eNPCStats.InvReSpawn
                        .InvReSpawn = Val(Value)
                    
                    Case eNPCStats.LanzaSpells
                        .flags.LanzaSpells = Val(Value)
                    
                    Case eNPCStats.MaxHit
                        .Stats.MaxHit = Val(Value)
                    
                    Case eNPCStats.MaxHP
                        .Stats.MaxHP = Val(Value)
                    
                    Case eNPCStats.MaxRecompensa
                        .MaxRecom = Val(Value)
                    
                    Case eNPCStats.MinHit
                        .Stats.MinHit = Val(Value)
                    
                    Case eNPCStats.MinHP
                        .Stats.MinHP = Val(Value)
                    
                    Case eNPCStats.MinRecompensa
                        .MinRecom = Val(Value)
                    
                    Case eNPCStats.Movement
                        .Movement = Val(Value)
                    
                    Case eNPCStats.NPCtype
                        .NPCtype = Val(Value)
                    
                    Case eNPCStats.NroCriaturas
                        .NroCriaturas = Val(Value)
                    
                    Case eNPCStats.NroItems
                        .Invent.NroItems = Val(Value)
                    
                    Case eNPCStats.PocaParalisis
                        .flags.PocaParalisis = Val(Value)
                    
                    Case eNPCStats.PoderAtaque
                        .PoderAtaque = Val(Value)
                    
                    Case eNPCStats.PoderEvasion
                        .PoderEvasion = Val(Value)
                    
                    Case eNPCStats.Probabilidad
                        .Probabilidad = Val(Value)
                    
                    Case eNPCStats.Respawn
                        .flags.Respawn = Val(Value)
                    
                    Case eNPCStats.RespawnOrigPos
                        .flags.RespawnOrigPos = Val(Value)
                    
                    Case eNPCStats.Snd1
                        .flags.Snd1 = Val(Value)
                    
                    Case eNPCStats.Snd2
                        .flags.Snd2 = Val(Value)
                    
                    Case eNPCStats.Snd3
                        .flags.Snd3 = Val(Value)
                    
                    Case eNPCStats.Snd4
                        .flags.Snd4 = Val(Value)
                    
                    Case eNPCStats.Sound
                        .flags.Sound = Val(Value)
                    
                    Case eNPCStats.TierraInvalida
                        .flags.TierraInvalida = Val(Value)
                    
                    Case eNPCStats.TipoItems
                        .TipoItems = Val(Value)
                    
                    Case eNPCStats.VeInvis
                        .VeInvis = Val(Value)
                    
                    Case eNPCStats.Veneno
                        .Veneno = Val(Value)
                    
                End Select
            End With
            
        Case eModo.Hechizo
            With Hechizos(Index)
                Select Case nINdex
                    Case eHecStats.Desc
                        Hechizos(Index).Desc = Value
                    
                    Case eHecStats.PalabrasMagicas
                        Hechizos(Index).PalabrasMagicas = Value
                    
                    Case eHecStats.HechizeroMsg
                        Hechizos(Index).HechizeroMsg = Value
                        
                    Case eHecStats.TargetMsg
                        Hechizos(Index).TargetMsg = Value
                        
                    Case eHecStats.PropioMsg
                        Hechizos(Index).PropioMsg = Value
                        
                    
                    Case eHecStats.Tipo
                        Hechizos(Index).Tipo = Val(Value)
                        
                    Case eHecStats.WAV
                        Hechizos(Index).WAV = Val(Value)
                        
                    Case eHecStats.FXgrh
                        Hechizos(Index).FXgrh = Val(Value)
                    
                    Case eHecStats.Loops
                        Hechizos(Index).Loops = Val(Value)
                        
                    
                    Case eHecStats.Resis
                        Hechizos(Index).Resis = Val(Value)
                        
                    Case eHecStats.Baculo
                        Hechizos(Index).Baculo = Val(Value)
                        
                    
                    Case eHecStats.SubeHP
                        Hechizos(Index).SubeHP = Val(Value)
                        
                    Case eHecStats.MinHP
                        Hechizos(Index).MinHP = Val(Value) '
                        
                    Case eHecStats.MaxHP
                        Hechizos(Index).MaxHP = Val(Value)
                        
                    
                    Case eHecStats.SubeMana
                        Hechizos(Index).SubeMana = Val(Value)
                        
                    Case eHecStats.MiMana
                        Hechizos(Index).MiMana = Val(Value)
                        
                    Case eHecStats.MaMana
                        Hechizos(Index).MaMana = Val(Value)
                        
                    Case eHecStats.SubeSta
                        Hechizos(Index).SubeSta = Val(Value)
                        
                    Case eHecStats.MinSta
                        Hechizos(Index).MinSta = Val(Value)
                        
                    Case eHecStats.MaxSta
                        Hechizos(Index).MaxSta = Val(Value)
                        
                    
                    Case eHecStats.SubeHam
                        Hechizos(Index).SubeHam = Val(Value)
                        
                    Case eHecStats.MinHam
                        Hechizos(Index).MinHam = Val(Value)
                        
                    Case eHecStats.MaxHam
                        Hechizos(Index).MaxHam = Val(Value)
                        
                    
                    Case eHecStats.SubeSed
                        Hechizos(Index).SubeSed = Val(Value)
                        
                    Case eHecStats.MinSed
                        Hechizos(Index).MinSed = Val(Value)
                        
                    Case eHecStats.MaxSed
                        Hechizos(Index).MaxSed = Val(Value)
                        
                    
                    Case eHecStats.SubeAgilidad
                        Hechizos(Index).SubeAgilidad = Val(Value)
                        
                        
                    Case eHecStats.MinAgilidad
                        Hechizos(Index).MinAgilidad = Val(Value) '
                        
                        
                    Case eHecStats.MaxAgilidad
                        Hechizos(Index).MaxAgilidad = Val(Value)
                        
                    
                    Case eHecStats.SubeFuerza
                        Hechizos(Index).SubeFuerza = Val(Value)
                        
                    Case eHecStats.MinFuerza
                        Hechizos(Index).MinFuerza = Val(Value)
                        
                    Case eHecStats.MaxFuerza
                        Hechizos(Index).MaxFuerza = Val(Value)
                        
                    
                    Case eHecStats.SubeCarisma
                        Hechizos(Index).SubeCarisma = Val(Value)
                        
                    Case eHecStats.MinCarisma
                        Hechizos(Index).MinCarisma = Val(Value)
                        
                    Case eHecStats.MaxCarisma
                        Hechizos(Index).MaxCarisma = Val(Value)
                        
                    
                    Case eHecStats.Invisibilidad
                        Hechizos(Index).Invisibilidad = Val(Value)
                        
                    Case eHecStats.Paraliza
                        Hechizos(Index).Paraliza = Val(Value)
                        
                    
                    Case eHecStats.Transforma
                        Hechizos(Index).Transforma = Val(Value)
                        
                    Case eHecStats.Envenena
                        Hechizos(Index).Envenena = Val(Value)
                        
                    Case eHecStats.Ceguera
                        Hechizos(Index).Ceguera = Val(Value)
                        
                    Case eHecStats.Estupidez
                        Hechizos(Index).Estupidez = Val(Value)
                        
                    
                    Case eHecStats.Revivir
                        Hechizos(Index).Revivir = Val(Value)
                        
                    Case eHecStats.Flecha
                        Hechizos(Index).Flecha = Val(Value)
                        
                    
                    Case eHecStats.Metamorfosis
                        Hechizos(Index).Metamorfosis = Val(Value)
                        
                    Case eHecStats.Maldicion
                        Hechizos(Index).Maldicion = Val(Value)
                        
                    Case eHecStats.Bendicion
                        Hechizos(Index).Bendicion = Val(Value)
                        
                    
                    Case eHecStats.RemoverParalisis
                        Hechizos(Index).RemoverParalisis = Val(Value)
                        
                    Case eHecStats.CuraVeneno
                        Hechizos(Index).CuraVeneno = Val(Value)
                        
                    Case eHecStats.RemoverMaldicion
                        Hechizos(Index).RemoverMaldicion = Val(Value)
                        
                    
                    Case eHecStats.Invoca
                        Hechizos(Index).Invoca = Val(Value)
                        
                    Case eHecStats.NumNPC
                        Hechizos(Index).NumNPC = Val(Value)
                        
                    Case eHecStats.cant
                        Hechizos(Index).cant = Val(Value)
                        
                    
                    Case eHecStats.Materializa
                        Hechizos(Index).Materializa = Val(Value)
                        
                    Case eHecStats.CuraArea
                        Hechizos(Index).CuraArea = Val(Value)
                        
                    Case eHecStats.ItemIndex
                        Hechizos(Index).ItemIndex = Val(Value)
                        
                    
                    Case eHecStats.Nivel
                        Hechizos(Index).Nivel = Val(Value)
                        
                    Case eHecStats.MinSkill
                        Hechizos(Index).MinSkill = Val(Value)
                        
                    Case eHecStats.ManaRequerido
                        Hechizos(Index).ManaRequerido = Val(Value)
                        
                    Case eHecStats.StaRequerido
                        Hechizos(Index).StaRequerido = Val(Value)
                        
                    
                    Case eHecStats.Target
                        Hechizos(Index).Target = Val(Value)
                        
                End Select
            End With
    End Select
End Sub








