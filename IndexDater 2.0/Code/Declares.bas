Attribute VB_Name = "Mod_Declaraciones"
'Argentum Online 0.9.0.9
'
'Copyright (C) 2002 Márquez Pablo Ignacio
'Copyright (C) 2002 Otto Perez
'Copyright (C) 2002 Aaron Perkins
'
'This program is free software; you can redistribute it and/or modify
'it under the terms of the GNU General Public License as published by
'the Free Software Foundation; either version 2 of the License, or
'any later version.
'
'This program is distributed in the hope that it will be useful,
'but WITHOUT ANY WARRANTY; without even the implied warranty of
'MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'GNU General Public License for more details.
'
'You should have received a copy of the GNU General Public License
'along with this program; if not, write to the Free Software
'Foundation, Inc., 59 Temple Place, Suite 330, Boston, MA  02111-1307  USA
'
'Argentum Online is based on Baronsoft's VB6 Online RPG
'You can contact the original creator of ORE at aaron@baronsoft.com
'for more information about ORE please visit http://www.baronsoft.com/
'
'
'You can contact me at:
'morgolock@speedy.com.ar
'www.geocities.com/gmorgolock
'Calle 3 número 983 piso 7 dto A
'La Plata - Pcia, Buenos Aires - Republica Argentina
'Código Postal 1900
'Pablo Ignacio Márquez


Option Explicit

'Objetos públicos
Public Enum eProModo
    clases = 1
    razas = 2
End Enum

Public SelectBMP() As Long
Public nSelectBMP As Byte

Public Enum eModo
    Objetos = 1
    Npc
    Hechizo
End Enum

Public estadoDat As eModo
Public DatNoGuardado(1 To 3) As Boolean

Public SurfaceDB As clsSurfaceManager   'No va new porque es unainterfaz, el new se pone al decidir que clase de objeto es
Public sNpcStat(97) As String
Public sObjStat(97) As String
Public sHecStat(97) As String

Public NumeroHechizos As Integer
'' The main timer of the game.

Public LOOPActual As Long
Public GRHActual As Long
Public DataIndexActual As Integer
Public Const VERSION_ACTUAL As String = "1.06"
'Sonidos


' Head index of the casper. Used to know if a char is killed
Public Const CASPER_HEAD As Integer = 500

Public LastFound As Long
Public BMPBuscado As Long

Public LoadingNew As Boolean

'Musica
Public Const MIdi_Inicio As Byte = 6

Public RawServersList As String
Public SavePath As Byte
Public DibujarFondo As Boolean
Public ColorFondo As Long

Public GrHCambiando As Boolean
Public TempGrh As Grhdata
Public tempDataIndex As BodyData

Public Type tColor
    r As Byte
    G As Byte
    b As Byte
End Type

Type IndexacionActual
    Total As Integer
    Inicios(1 To 10000) As Position
    activo As Boolean
    Ancho As Integer
    Alto As Integer
End Type

Type tDirectorios
    Inits As String
    Graficos As String
    Dats As String
    InitWE As String
End Type

Public ConfigDir As tDirectorios
Public ColoresPJ(0 To 50) As tColor
Public DibujarIndexaciones As IndexacionActual

Public Type tServerInfo
    Ip As String
    Puerto As Integer
    Desc As String
    PassRecPort As Integer
End Type

Public currentMidi As Long

Public ServersLst() As tServerInfo
Public ServersRecibidos As Boolean

Public CurServer As Integer

Public CreandoClan As Boolean
Public ClanName As String
Public Site As String

Public UserCiego As Boolean
Public UserEstupido As Boolean

Public NoRes As Boolean 'no cambiar la resolucion

Public RainBufferIndex As Long
Public FogataBufferIndex As Long

Public Const bCabeza = 1
Public Const bPiernaIzquierda = 2
Public Const bPiernaDerecha = 3
Public Const bBrazoDerecho = 4
Public Const bBrazoIzquierdo = 5
Public Const bTorso = 6

'Timers de GetTickCount
Public Const tAt = 2000
Public Const tUs = 600

Public Const PrimerBodyBarco = 84
Public Const UltimoBodyBarco = 87

Public NumEscudosAnims As Integer

Public ArmasHerrero(0 To 100) As Integer
Public ArmadurasHerrero(0 To 100) As Integer
Public ObjCarpintero(0 To 100) As Integer

Public Versiones(1 To 7) As Integer

Public UsaMacro As Boolean
Public CnTd As Byte

Public Trabajando As Boolean





Public Tips() As String * 255
Public Const LoopAdEternum As Integer = 999

'Direcciones
Public Enum E_Heading
    NORTH = 1
    EAST = 2
    SOUTH = 3
    WEST = 4
End Enum



Public Type BITMAPFILEHEADER
        bfType As Integer
        bfSize As Long
        bfReserved1 As Integer
        bfReserved2 As Integer
        bfOffBits As Long
End Type

'Info del encabezado del bmp
Public Type BITMAPINFOHEADER
        biSize As Long
        biWidth As Long
        biHeight As Long
        biPlanes As Integer
        biBitCount As Integer
        biCompression As Long
        biSizeImage As Long
        biXPelsPerMeter As Long
        biYPelsPerMeter As Long
        biClrUsed As Long
        biClrImportant As Long
End Type
Public Type RGBQUAD
    rgbBlue As Byte
    rgbGreen As Byte
    rgbRed As Byte
    rgbReserved As Byte
End Type
Public Type BITMAPINFO
    bmiHeader As BITMAPINFOHEADER
    bmiColors(255) As RGBQUAD
End Type

Public Type ArchivoBMP
    BMPFileHeader As BITMAPINFOHEADER
    bmpInfo As BITMAPINFO
    BMPData() As Byte
End Type

Public Type ResoGrap
    offset As Long
    Archivo As Byte
    tamaño As Long
End Type
Public Type RecursoGrafico
    Graficos() As ResoGrap
    UltimoGrafico As Long
End Type

Global Const DIB_RGB_COLORS = 0
Global Const MAXGrH = 50000

Public Nombres As Boolean

Public MixedKey As Long


Public DibujarWalk As Integer
'<-------------------------NUEVO-------------------------->
Public Comerciando As Boolean
'<-------------------------NUEVO-------------------------->

Public UserClase As String
Public UserSexo As String
Public UserRaza As String
Public UserEmail As String


Public Const NUMCLASES = 56
Public Const NUMRAZAS = 5


Public SkillPoints As Integer
Public Alocados As Integer
Public flags() As Integer
Public Oscuridad As Integer
Public logged As Boolean
Public ListaClases(1 To NUMCLASES) As String
Public ListaRazas(1 To NUMRAZAS) As String
Public UsingSkill As Integer
Public Enum e_EstadoIndexador
    Grh = 0
    Body = 1
    Cabezas = 2
    Cascos = 3
    Escudos = 4
    Armas = 5
    FX = 6
    Resource = 7
    Superficies = 8
End Enum
Public EstadoIndexador As e_EstadoIndexador

Public UltimoindexE(e_EstadoIndexador.Grh To e_EstadoIndexador.Superficies) As Long
Public EstadoNoGuardado(e_EstadoIndexador.Grh To e_EstadoIndexador.Superficies) As Boolean
Public MD5HushYo As String * 16




   
Public Enum FxMeditar
    CHICO = 4
    MEDIANO = 5
    GRANDE = 6
    XGRANDE = 16
End Enum
Public Enum EGuildPermiso
    VerBodeda = 1
    DepositarBodeda = 2
    RetirarBoveda = 3
    VerMiembro = 4
    AceptarMiembro = 5
    ExpulsarMiembro = 6
    CambiarGuildNews = 7
End Enum

Public cabezaActual As Integer
'Server stuff
Public RequestPosTimer As Integer 'Used in main loop
Public stxtbuffer As String 'Holds temp raw data from server
Public stxtbuffercmsg As String 'Holds temp raw data from server
Public SendNewChar As Boolean 'Used during login
Public Connected As Boolean 'True when connected to server
Public DownloadingMap As Boolean 'Currently downloading a map from server
Public UserMap As Integer

'String contants
Public Const ENDC As String * 1 = vbNullChar    'Endline character for talking with server
Public Const ENDL As String * 2 = vbCrLf        'Holds the Endline character for textboxes

'Control
Public prgRun As Boolean 'When true the program ends

Public IPdelServidor As String
Public PuertoDelServidor As String
Public ResourceF As RecursoGrafico
Public ResourceFile As Byte
Public UsarGrhLong As Boolean
Public IniciadoTodo As Boolean
'
'********** FUNCIONES API ***********
'

Public Declare Function GetTickCount Lib "kernel32" () As Long

'para escribir y leer variables
Public Declare Function writeprivateprofilestring Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationname As String, ByVal lpKeyname As Any, ByVal lpString As String, ByVal lpfilename As String) As Long
Public Declare Function getprivateprofilestring Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationname As String, ByVal lpKeyname As Any, ByVal lpdefault As String, ByVal lpreturnedstring As String, ByVal nsize As Long, ByVal lpfilename As String) As Long

'Teclado
Public Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer
Public Declare Function GetAsyncKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer
Public Declare Function StretchDIBits Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal DX As Long, ByVal dy As Long, ByVal srcX As Long, ByVal srcY As Long, ByVal wSrcWidth As Long, ByVal wSrcHeight As Long, lpBits As Any, lpBitsInfo As BITMAPINFO, ByVal wUsage As Long, ByVal dwRop As Long) As Long

Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)


Public Type tIndiceCuerpo
    Body(1 To 4) As Integer
    HeadOffsetX As Integer
    HeadOffsetY As Integer
End Type

Public Type tIndiceCuerpoLong
    Body(1 To 4) As Long
    HeadOffsetX As Integer
    HeadOffsetY As Integer
End Type

Public Type tIndiceFx
    Animacion As Long
    offsetx As Integer
    offsety As Integer
    FXTransparente As Boolean
End Type


Type InfoHerre
    Index As Integer
    Recompensa As Byte
End Type

Public Type ObjData
    Modificacion As Boolean
    Modificando As Boolean
    FueModificado As Boolean
    itPart As Integer

    name As String
    NoComerciable As Integer
    Aura As Integer
    Remort As Byte
    
    NoSeCae As Boolean
    objtype As Integer
    SubTipo As Integer
    Dosmanos As Integer
    GrhIndex As Integer
    GrhSecundario As Integer
    Jerarquia As Byte
    
    Respawn As Byte
    
    
    MaxItems As Integer
    'Conte As Inventario
    Apuñala As Byte
    
    HechizoIndex As Integer
    
    ForoID As String
    
    MinHP As Integer
    MaxHP As Integer
    
    WMapa As Integer
    WX As Integer
    WY As Integer
    WI As Integer
    
    Baculo As Byte
    
    MineralIndex As Integer

    
    proyectil As Integer
    Municion As Integer
    
    Crucial As Byte
    Newbie As Integer
    
    
    MinSta As Integer
    
    
    TipoPocion As Byte
    MaxModificador As Integer
    MinModificador As Integer
    DuracionEfecto As Long
    
    MinSkill As Integer
    LingoteIndex As Integer
    
    MinHit As Integer
    MaxHit As Integer
    
    MinHam As Integer
    MinSed As Integer
    
    Caballo As Integer
    
    Def As Integer
    
    MinDef As Integer
    MaxDef As Integer
    Ropaje As Integer
    
    plusMagia As Byte
    
    WeaponAnim As Integer
    ShieldAnim As Integer
    CascoAnim As Integer
    Gorro As Byte
    
    Valor As Long
    Info As String
    Cerrada As Integer
    Llave As Byte
    Clave As Long
    
    IndexAbierta As Integer
    IndexCerrada As Integer
    IndexCerradaLlave As Integer
    
    RazaEnana As Byte
    MUJER As Byte
    HOMBRE As Byte
    Envenena As Byte
    
    SkillCombate As Byte
    SkillTacticas As Byte
    SkillProyectiles As Byte
    SkillApuñalar As Byte
    
    Resistencia As Long
    
    Agarrable As Byte
    
    ArbolElfico As Byte
    LingH As Integer
    LingO As Integer
    LingP As Integer
    Madera As Integer
    MaderaElfica As Integer
    
     
    Raices As Integer
    PielLobo As Integer
    PielOsoPardo As Integer
    PielOsoPolar As Integer

    SkHerreria As Integer
    SkCarpinteria As Integer
    SkResistencia As Integer
    SkDefensa As Integer
        
    SkPociones As Integer
    SkSastreria As Integer
    
    Texto As String
    
    
    ClaseProhibida(1 To NUMCLASES) As Integer
    RazaProhibida(1 To NUMRAZAS) As Integer
    Snd1 As Integer
    Snd2 As Integer
    Snd3 As Integer
    MinInt As Integer
    
    Real As Integer
    Caos As Integer
    
End Type

Public Type Obj
    OBJIndex As Integer
    Amount As Integer
End Type

Type tHechizo
    Modificando As Boolean
    FueModificado As Boolean
    Modificacion As Boolean
    Nombre As String
    Desc As String
    PalabrasMagicas As String
    
    HechizeroMsg As String
    TargetMsg As String
    PropioMsg As String
    
    Resis As Byte
    
    Tipo As Byte
    WAV As Integer
    FXgrh As Integer
    Loops As Byte
    
    Baculo As Byte
    SubeHP As Byte
    MinHP As Integer
    MaxHP As Integer
    
    SubeMana As Byte
    MiMana As Integer
    MaMana As Integer
    
    SubeSta As Byte
    MinSta As Integer
    MaxSta As Integer
    
    SubeHam As Byte
    MinHam As Integer
    MaxHam As Integer
    
    SubeSed As Byte
    MinSed As Integer
    MaxSed As Integer
    
    SubeAgilidad As Byte
    MinAgilidad As Integer
    MaxAgilidad As Integer
    Metamorfosis As Byte
    Body As Integer
    SubeFuerza As Byte
    MinFuerza As Integer
    MaxFuerza As Integer
    
    SubeCarisma As Byte
    MinCarisma As Integer
    MaxCarisma As Integer

    RadioX As Byte
    RadioY As Byte
    
    Invisibilidad As Byte
    Transforma As Byte
    Paraliza As Byte
    Nivel As Byte
    RemoverParalisis As Byte
    CuraVeneno As Byte
    Envenena As Byte
    Maldicion As Byte
    RemoverMaldicion As Byte
    Bendicion As Byte
    Estupidez As Byte
    NoAtacar As Byte
    Ceguera As Byte
    Revivir As Byte
    Flecha As Byte
    Morph As Byte
    
    Invoca As Byte
    NumNPC As Integer
    cant As Integer
    
    Materializa As Byte
    CuraArea As Integer
    ItemIndex As Byte
    
    MinSkill As Integer
    StaRequerido As Integer
    ManaRequerido As Integer

    Target As Byte
End Type


Type NPCStats
    AutoCurar As Byte
    Alineacion As Byte
    MaxHP As Long
    MinHP As Long
    MaxHit As Integer
    MinHit As Integer
    Def As Integer
    ImpactRate As Integer
End Type

Type NpcCounters
    Paralisis As Long
    TiempoExistencia As Long
End Type

Type NPCFlags
    Apostador As Byte
    TiendaUser As Integer
    PocaParalisis As Byte
    AfectaParalisis As Byte
    GolpeExacto As Byte
    Domable As Integer
    Respawn As Byte
    NPCActive As Boolean
    Follow As Boolean
    Faccion As Byte
    LanzaSpells As Byte
    QuienParalizo As Integer
    OldMovement As Byte
    OldHostil As Byte
    
    NoMagia As Byte
    AguaValida As Byte
    TierraInvalida As Byte
    
    UseAINow As Boolean
    Sound As Integer
    AttackedBy As Integer
    RespawnOrigPos As Byte
    
    Envenenado As Byte
    Paralizado As Byte
    Invisible As Byte
    Maldicion As Byte
    Bendicion As Byte
    
    Snd1 As Integer
    Snd2 As Integer
    Snd3 As Integer
    Snd4 As Integer
    
End Type

Type tCriaturasEntrenador
    npcIndex As Integer
    NpcName As String
    tmpIndex As Integer
End Type

Type UserOBJ
    OBJIndex As Integer
    Amount As Integer
   ' Equipped As Byte

End Type

Type InventarioNPC
    Object(1 To 30) As UserOBJ
    NroItems As Integer
End Type
Type Char1
    itPart As Integer
    Aura As Integer
    CharIndex As Integer
    Head As Integer
    exBody As Integer
    Body As Integer
    
    WeaponAnim As Integer
    ShieldAnim As Integer
    CascoAnim As Integer
    
    FX As Integer
    Loops As Integer
    
    Heading As Byte
End Type

Type Npc
    Modificacion As Boolean
    FueModificado As Boolean
    Modificando As Boolean
    'quest As tQuestUser
    name As String
    Char As Char1
    Desc As String
    
    NPCtype As Integer
    Numero As Integer
    AutoCurar As Integer
    
    level As Integer
    
    InvReSpawn As Byte
    
    Comercia As Integer
    Target As Long
    TargetNpc As Long
    TipoItems As Integer
    
    Veneno As Byte
    
    Pos As WorldPos
    Orig As WorldPos
    SkillDomar As Integer
    
    Movement As Integer
    Attackable As Byte
    hostilE As Byte
    PoderAtaque As Long
    PoderEvasion As Long
     InmuneParalisis As Byte
    Inflacion As Long
    
    GiveEXP As Long
    GiveGLD As Long
    
    Stats As NPCStats
    flags As NPCFlags
    Contadores As NpcCounters
    
    Probabilidad As Integer
    MaxRecom As Integer
    MinRecom As Integer
    
    Invent As InventarioNPC
    CanAttack As Byte
    VeInvis As Byte
    
    NroExpresiones As Byte
    Expresiones() As String
    
    NroSpells As Byte
    Spells() As Integer
    
    NroCriaturas As Integer
    Criaturas() As tCriaturasEntrenador
   ' MaestroUser As Integer
    'MaestroNpc As Integer
  '  Mascotas As Integer

    'PFINFO As NpcPathFindingInfo
 
End Type


Public Npclist() As Npc
Public MapInfo() As MapInfo
Public Hechizos() As tHechizo
Public ObjData() As ObjData
Public SpawnList() As tCriaturasEntrenador

Public numObjs As Integer
'Public ArmasHerrero() As InfoHerre
'Public ArmadurasHerrero() As InfoHerre
Public CascosHerrero() As Integer
Public EscudosHerrero() As Integer
Public ObjDruida() As Integer
Public ObjSastre() As Integer
'Public ObjCarpintero() As InfoHerre

Public Const OBJTYPE_USEONCE = 1
Public Const OBJTYPE_WEAPON = 2
Public Const OBJTYPE_ARMOUR = 3
Public Const OBJTYPE_ARBOLES = 4
Public Const OBJTYPE_GUITA = 5
Public Const OBJTYPE_PUERTAS = 6
Public Const OBJTYPE_CONTENEDORES = 7
Public Const OBJTYPE_CARTELES = 8
Public Const OBJTYPE_LLAVES = 9
Public Const OBJTYPE_FOROS = 10
Public Const OBJTYPE_POCIONES = 11
Public Const OBJTYPE_BEBIDA = 13
Public Const OBJTYPE_LEÑA = 14
Public Const OBJTYPE_FOGATA = 15
Public Const OBJTYPE_HERRAMIENTAS = 18
Public Const OBJTYPE_YACIMIENTO = 22
Public Const OBJTYPE_PERGAMINOS = 24
Public Const OBJTYPE_TELEPORT = 19
Public Const OBJTYPE_YUNQUE = 27
Public Const OBJTYPE_FRAGUA = 28
Public Const OBJTYPE_MINERALES = 23
Public Const OBJTYPE_CUALQUIERA = 1000
Public Const OBJTYPE_INSTRUMENTOS = 26
Public Const OBJTYPE_BARCOS = 31
Public Const OBJTYPE_FLECHAS = 32
Public Const OBJTYPE_BOTELLAVACIA = 33
Public Const OBJTYPE_BOTELLALLENA = 34
Public Const OBJTYPE_MANCHAS = 35
Public Const OBJTYPE_GEMAZUL = 29
Public Const OBJTYPE_GEMNARANJA = 38
Public Const OBJTYPE_GEMCELESTE = 39
Public Const OBJTYPE_GEMLILA = 40
Public Const OBJTYPE_GEMROJO = 41
Public Const OBJTYPE_GEMVERDE = 42
Public Const OBJTYPE_GEMVIOLETA = 43
Public Const OBJTYPE_AMULETO = 44
Public Const OBJTYPE_RAIZ = 36
Public Const OBJTYPE_PIEL = 30
Public Const OBJTYPE_MONTURA = 45
Public Const OBJTYPE_WARP = 37


Public Const OBJTYPE_ARMADURA = 0
Public Const OBJTYPE_CASCO = 1
Public Const OBJTYPE_ESCUDO = 2
Public Const OBJTYPE_CAÑA = 138



Public Enum eObjStats
     NoComerciable = 1
     NoSeCae = 2
     objtype = 3
     SubTipo = 4
     Dosmanos = 5
     GrhIndex = 6
     GrhSecundario = 7
     Jerarquia = 8
     Respawn = 9
     MaxItems = 10
     Apuñala = 11
     HechizoIndex = 12
     ForoID = 13
     MinHP = 14
     MaxHP = 15
     Map_Pasajes = 16
     MapX = 17
     MapY = 18
     WarpI = 19
     Baculo = 20
     MineralIndex = 21
     Texto = 22
     proyectil = 23
     Municion = 24
     Crucial = 25
     Newbie = 26
     MinSta = 27
     TipoPocion = 28
     MaxModificador = 29
     MinModificador = 30
     DuracionEfecto = 31
     MinSkill = 32
     LingoteIndex = 33
     MinHit = 34
     MaxHit = 35
     Minhambre = 36
     MinSed = 37
     defensa = 38
     MinDef = 39
     MaxDef = 40
     Ropaje = 41
     Armaanim = 42
     Escudoanim = 43
     CascoAnim = 44
     Gorro = 45
     Valor = 46
     Cerrada = 47
     Llave = 48
     Clave = 49
     IndexAbierta = 50
     IndexCerrada = 51
     IndexCerradaLlave = 52
     RazaEnana = 53
     MUJER = 54
     HOMBRE = 55
     Envenena = 56
     SkillCombate = 57
     SkillTacticas = 58
     SkillProyectiles = 59
     SkillApuñalar = 60
     Resistencia = 61
     Agarrable = 62
     Arbol_Elfico = 63
     LingotesOro = 64
     LingotesPlata = 65
     LingotesHierro = 67
     Madera = 68
     Madera_Elfica = 69
     Raices = 70
     PielLobo = 71
     PielOso = 72
     PielOsoPolar = 73
     SkillHerreria = 74
     SkillCarpinteria = 75
     SkillResistencia = 76
     Skilldefensa = 77
     Skillpociones = 78
     Skillsastreria = 79
     Clasesprohib = 80 '59-35-32-30'
     Razasprohib = 81 '2-3-4' =
     Snd1 = 82
     Snd2 = 83
     Snd3 = 84
     MinInt = 85
     Real = 86
     Caos = 87
     CantItems = 88
     Info = 89
     plusMagia = 90
     Def = 91
     Aura = 92
     Remort = 93
     Caballo = 94
End Enum

Public Enum eNPCStats
    Desc = 1
    Movement
    AguaValida
    TierraInvalida
    Faccion
    NPCtype
    Body
    Head
    Heading
    Atacable
    Comercia
    hostil
    GiveEXP
    InmuneParalisis
    Veneno
    Domable
    MaxRecompensa
    MinRecompensa
    Probabilidad
    GiveGLD
    PoderAtaque
    PoderEvasion
    AutoCurar
    MaxHP
    MinHP
    MaxHit
    MinHit
    Def
    Alineacion
    ImpactRate
    InvReSpawn
    NroItems
    LanzaSpells
    NroCriaturas
    Inflacion
    Respawn
    AfectaParalisis
    RespawnOrigPos
    GolpeExacto
    PocaParalisis
    VeInvis
    Snd1
    Snd2
    Snd3
    Snd4
    Sound
    TipoItems
End Enum


Public Enum eHecStats
    'Nombre ' '(datpath & "Hechizos.dat", "Hechizo" & Hechizo, "Nombre")
    Desc = 1 ' '(datpath & "Hechizos.dat", "Hechizo" & Hechizo, "Desc")
    PalabrasMagicas ' '(datpath & "Hechizos.dat", "Hechizo" & Hechizo, "PalabrasMagicas")
    
    HechizeroMsg ' '(datpath & "Hechizos.dat", "Hechizo" & Hechizo, "HechizeroMsg")
    TargetMsg ' '(datpath & "Hechizos.dat", "Hechizo" & Hechizo, "TargetMsg")
    PropioMsg ' '(datpath & "Hechizos.dat", "Hechizo" & Hechizo, "PropioMsg")
    
    Tipo ' 'val('(datpath & "Hechizos.dat", "Hechizo" & Hechizo, "Tipo"))
    WAV ' 'val('(datpath & "Hechizos.dat", "Hechizo" & Hechizo, "WAV"))
    FXgrh ' 'val('(datpath & "Hechizos.dat", "Hechizo" & Hechizo, "Fxgrh"))
    
    Loops ' 'val('(datpath & "Hechizos.dat", "Hechizo" & Hechizo, "Loops"))
    
    Resis ' 'val('(datpath & "Hechizos.dat", "Hechizo" & Hechizo, "Resis"))
    Baculo ' 'val('(datpath & "Hechizos.dat", "Hechizo" & Hechizo, "Baculo"))
       
    SubeHP ' 'val('(datpath & "Hechizos.dat", "Hechizo" & Hechizo, "SubeHP"))
    MinHP ' 'val('(datpath & "Hechizos.dat", "Hechizo" & Hechizo, "MinHP"))
    MaxHP ' 'val('(datpath & "Hechizos.dat", "Hechizo" & Hechizo, "MaxHP"))
    
    SubeMana ' 'val('(datpath & "Hechizos.dat", "Hechizo" & Hechizo, "SubeMana"))
    MiMana ' 'val('(datpath & "Hechizos.dat", "Hechizo" & Hechizo, "MinMana"))
    MaMana ' 'val('(datpath & "Hechizos.dat", "Hechizo" & Hechizo, "MaxMana"))
    
    SubeSta ' 'val('(datpath & "Hechizos.dat", "Hechizo" & Hechizo, "SubeSta"))
    MinSta ' 'val('(datpath & "Hechizos.dat", "Hechizo" & Hechizo, "MinSta"))
    MaxSta ' 'val('(datpath & "Hechizos.dat", "Hechizo" & Hechizo, "MaxSta"))
    
    SubeHam ' 'val('(datpath & "Hechizos.dat", "Hechizo" & Hechizo, "SubeHam"))
    MinHam ' 'val('(datpath & "Hechizos.dat", "Hechizo" & Hechizo, "MinHam"))
    MaxHam ' 'val('(datpath & "Hechizos.dat", "Hechizo" & Hechizo, "MaxHam"))
    
    SubeSed ' 'val('(datpath & "Hechizos.dat", "Hechizo" & Hechizo, "SubeSed"))
    MinSed ' 'val('(datpath & "Hechizos.dat", "Hechizo" & Hechizo, "MinSed"))
    MaxSed ' 'val('(datpath & "Hechizos.dat", "Hechizo" & Hechizo, "MaxSed"))
    
    SubeAgilidad ' 'val('(datpath & "Hechizos.dat", "Hechizo" & Hechizo, "SubeAG"))
    MinAgilidad ' 'val('(datpath & "Hechizos.dat", "Hechizo" & Hechizo, "MinAG"))
    MaxAgilidad ' 'val('(datpath & "Hechizos.dat", "Hechizo" & Hechizo, "MaxAG"))
    
    SubeFuerza ' 'val('(datpath & "Hechizos.dat", "Hechizo" & Hechizo, "SubeFU"))
    MinFuerza ' 'val('(datpath & "Hechizos.dat", "Hechizo" & Hechizo, "MinFU"))
    MaxFuerza ' 'val('(datpath & "Hechizos.dat", "Hechizo" & Hechizo, "MaxFU"))
    
    SubeCarisma ' 'val('(datpath & "Hechizos.dat", "Hechizo" & Hechizo, "SubeCA"))
    MinCarisma ' 'val('(datpath & "Hechizos.dat", "Hechizo" & Hechizo, "MinCA"))
    MaxCarisma ' 'val('(datpath & "Hechizos.dat", "Hechizo" & Hechizo, "MaxCA"))
    
    Invisibilidad ' 'val('(datpath & "Hechizos.dat", "Hechizo" & Hechizo, "Invisibilidad"))
    Paraliza ' 'val('(datpath & "Hechizos.dat", "Hechizo" & Hechizo, "Paraliza"))
    
    Transforma ' 'val('(datpath & "Hechizos.dat", "Hechizo" & Hechizo, "Transforma"))
    Envenena ' 'val('(datpath & "Hechizos.dat", "Hechizo" & Hechizo, "Envenena"))
    Ceguera ' 'val('(datpath & "Hechizos.dat", "Hechizo" & Hechizo, "Ceguera"))
    Estupidez ' 'val('(datpath & "Hechizos.dat", "Hechizo" & Hechizo, "Estupidez"))

    Revivir ' 'val('(datpath & "Hechizos.dat", "Hechizo" & Hechizo, "Revivir"))
    Flecha ' 'val('(datpath & "Hechizos.dat", "Hechizo" & Hechizo, "Flecha"))
    
    Metamorfosis ' 'val('(datpath & "Hechizos.dat", "Hechizo" & Hechizo, "Metamorfosis"))
    Maldicion ' 'val('(datpath & "Hechizos.dat", "Hechizo" & Hechizo, "Maldicion"))
    Bendicion ' 'val('(datpath & "Hechizos.dat", "Hechizo" & Hechizo, "Bendicion"))
 
    RemoverParalisis ' 'val('(datpath & "Hechizos.dat", "Hechizo" & Hechizo, "RemoverParalisis"))
    CuraVeneno ' 'val('(datpath & "Hechizos.dat", "Hechizo" & Hechizo, "CuraVeneno"))
    RemoverMaldicion ' 'val('(datpath & "Hechizos.dat", "Hechizo" & Hechizo, "RemoverMaldicion"))
    
    Invoca ' 'val('(datpath & "Hechizos.dat", "Hechizo" & Hechizo, "Invoca"))
    NumNPC ' 'val('(datpath & "Hechizos.dat", "Hechizo" & Hechizo, "NumNpc"))
    cant ' 'val('(datpath & "Hechizos.dat", "Hechizo" & Hechizo, "Cant"))
    
    Materializa ' 'val('(datpath & "Hechizos.dat", "Hechizo" & Hechizo, "Materializa"))
    CuraArea ' 'val('(datpath & "Hechizos.dat", "Hechizo" & Hechizo, "CuraArea"))
    ItemIndex ' 'val('(datpath & "Hechizos.dat", "Hechizo" & Hechizo, "ItemIndex"))
    
    Nivel ' 'val('(datpath & "Hechizos.dat", "Hechizo" & Hechizo, "Nivel"))
    MinSkill ' 'val('(datpath & "Hechizos.dat", "Hechizo" & Hechizo, "MinSkill"))
    ManaRequerido ' 'val('(datpath & "Hechizos.dat", "Hechizo" & Hechizo, "ManaRequerido"))
    StaRequerido ' 'val('(datpath & "Hechizos.dat", "Hechizo" & Hechizo, "StaRequerido"))
    
    Target ' 'val('(datpath & "Hechizos.dat", "Hechizo" & Hechizo, "Target"))
    
    NoAtacar
End Enum
Public Enum eObjtype
    UseOnce = 1
    Arma = 2
    Armadura = 3
    Arboles = 4
    Dinero = 5
    Puertas = 6
    Contenedores = 7
    Carteles = 8
    Llaves = 9
    Foros = 10
    pociones = 11
    Bebidas = 13
    Libros = 12
    Leña = 14
    Fogata = 15
    Herramientas = 18
    Yacimiento = 22
    Pergaminos = 24
    Teleports = 19
    Yunque = 27
    Fragua = 28
    Minerales = 23
    Instrumentos = 26
    Barcos = 31
    Flechas = 32
    Botellavacia = 33
    Botellallena = 34
    Manchas = 35
    Raiz = 36
    Piel = 30
    Warp = 37
    Amuleto = 44
    Montura = 45
End Enum
