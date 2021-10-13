Attribute VB_Name = "mod_Subasta"
Option Explicit

Private Type tOferta
    uName As String
    uID As Long
    Oro As Long
End Type

Private Type tSubasta
    Item As obj
    MinutosRestantes As Byte
    UltimaOferta As tOferta
    Activo As Boolean
    Subastador As String
    subastadorID As Long
End Type

'IMPORTANTE::: La configuracion del servidor esta en modo CHARFILES, si se quiere usar la DB se tendra que modificar las funciones de este MODULO
'AUTHOR: El_Santo43

Public Subasta As tSubasta

Public Function itemValido(ByVal ObjIndex As Integer) As Boolean
    If ObjIndex > UBound(ObjData) Or ObjIndex < 1 Then itemValido = False: Exit Function
    
    'Aqui se harian los chequeos para ver cuales items se pueden subastar y cuales no.
End Function


Public Function amountValida(ByVal ui As Integer, ByVal Slot As Byte, ByVal amount As Integer) As Boolean
    Dim objamount As Integer
    objamount = UserList(ui).Invent.Object(Slot).amount
    If amount > objamount Then amountValida = False: Exit Function
End Function

Function meterItemEnBoveda(ByVal Userindex As Integer, _
                ByRef tmpObj As obj) As Boolean

    Dim Slot As Integer
    If tmpObj.amount < 1 Then Exit Function
    
    With UserList(Userindex)

        'Ya tiene un objeto de este tipo?
        Slot = 1

        Do Until .BancoInvent.Object(Slot).ObjIndex = tmpObj.ObjIndex And .BancoInvent.Object(Slot).amount + tmpObj.amount <= MAX_INVENTORY_OBJS
            Slot = Slot + 1
            
            If Slot > MAX_BANCOINVENTORY_SLOTS Then
                Exit Do
            End If
        Loop
        
        'Sino se fija por un slot vacio antes del slot devuelto
        If Slot > MAX_BANCOINVENTORY_SLOTS Then
            Slot = 1
            Do Until .BancoInvent.Object(Slot).ObjIndex = 0
                Slot = Slot + 1
                If Slot > MAX_BANCOINVENTORY_SLOTS Then
                    'Call WriteConsoleMsg(Userindex, "No tienes mas espacio en el banco!!", FontTypeNames.FONTTYPE_INFO)
                    meterItemEnBoveda = False
                    Exit Function
                End If
            Loop
            
            .BancoInvent.NroItems = .BancoInvent.NroItems + 1
        End If
        
        If Slot <= MAX_BANCOINVENTORY_SLOTS Then 'Slot valido
            'Mete el obj en el slot
            If .BancoInvent.Object(Slot).amount + tmpObj.amount <= MAX_INVENTORY_OBJS Then
                'Menor que MAX_INV_OBJS
                .BancoInvent.Object(Slot).ObjIndex = tmpObj.ObjIndex
                .BancoInvent.Object(Slot).amount = .BancoInvent.Object(Slot).amount + tmpObj.amount
                meterItemEnBoveda = True
            End If
        End If
    End With
    
End Function

Public Sub darItemaUserOffline(ByVal Name As String, ByVal ObjIndex As Integer, ByVal amount As Integer)
    Dim mifile As String, i As Long
    Dim objStr As String
    Dim Slot As Byte
    Dim Reader As clsIniManager
    
    
    If Not Database_Enabled Then
        Set Reader = New clsIniManager
        'charfiles
        mifile = CharPath & UCase$(Name) & ".chr"
        Reader.Initialize mifile
        
        'aqui deberiamos intentar meterlo en la boveda, pero es muy improbable que no tenga espacio en el inventario.
        For i = 1 To MAX_BANCOINVENTORY_SLOTS
            objStr = Reader.GetValue("BANCOINVENTORY", "OBJ" & i)
            If ReadField(1, objStr, 45) <= 0 Then
                Slot = i
                Exit For
            End If
        Next i
        
        If Slot > 0 Then
            Call Reader.ChangeValue("BANCOINVENTORY", "OBJ" & Slot, ObjIndex & "-" & amount)
            Call Reader.DumpFile(mifile)
            
        End If
        Set Reader = Nothing
    Else
        'lo guardamos directo a la DB
        
    End If
End Sub


Public Sub subasta_PasarMinuto()
    Dim ui As Integer, tmpOro As Long
    With Subasta
        If .Activo = True Then
        
            If .MinutosRestantes >= 1 Then
                .MinutosRestantes = .MinutosRestantes - 1
                Call MensajeGlobal("Subasta> Quedan " & .MinutosRestantes & " minutos, escribe '/OFRECER Cantidad' para participar de la subasta", FontTypeNames.FONTTYPE_GUILD)
        
            ElseIf .MinutosRestantes = 0 Then
                .Activo = False
                
                If LenB(.UltimaOferta.uName) > 0 Then
                    Call MensajeGlobal("Subasta> Vendido a " & .UltimaOferta.uName & " por " & .UltimaOferta.Oro, FontTypeNames.FONTTYPE_GUILD)
                    'primero verificamos si esta online
                    ui = NameIndex(.UltimaOferta.uName)
                    If ui > 0 Then

                        Call meterItemEnBoveda(ui, .Item)

                        Call WriteConsoleMsg(ui, "Subasta> Has ganado la subasta. El item fue guardado en tu boveda.", FontTypeNames.FONTTYPE_INFO)

                    Else
                        Call darItemaUserOffline(.UltimaOferta.uName, .Item.ObjIndex, .Item.amount)
                    End If
                    'ahora le damos el oro al subastador:
                    ui = NameIndex(Subasta.Subastador)
                    
                    If ui > 0 Then
                        UserList(ui).Stats.Gld = UserList(ui).Stats.Gld + Subasta.UltimaOferta.Oro
                        Call WriteUpdateUserStats(ui)
                    Else
                        ' el usuario esta offline, se lo tendremos que agregar al charfile/db
                        If Not Database_Enabled Then
                            tmpOro = CLng(GetVar(CharPath & UCase$(Subasta.Subastador) & ".chr", "STATS", "GLD"))
                            Call WriteVar(CharPath & UCase$(Subasta.Subastador) & ".chr", "STATS", "GLD", tmpOro + Subasta.UltimaOferta.Oro)
                        Else
                            SumGoldtoUserDB Subasta.Subastador, Subasta.UltimaOferta.Oro
                        End If
                    End If
                    
                Else ' Nadie oferto, le devolvemos el item al subastador.
                    ui = NameIndex(.Subastador)
                    If ui > 0 Then
                        'esta ON
                        Call meterItemEnBoveda(ui, .Item)
                        Call WriteConsoleMsg(ui, "Subasta> Nadie ofrecio en tu subasta. El item fue guardado en tu boveda.", FontTypeNames.FONTTYPE_INFO)
                    Else
                        Call darItemaUserOffline(.Subastador, .Item.ObjIndex, .Item.amount)
                    End If
                End If
                
            End If
            
        End If
    End With
End Sub

Public Sub UsuarioOfrece(ByVal ui As Integer, ByVal valor As Long)
    Dim tmpUI As Integer, tmpOro As Long
    
    If UserList(ui).Stats.Gld < valor Then
        Call WriteConsoleMsg(ui, "Subasta> No tenes esa cantidad", FontTypeNames.FONTTYPE_INFO)
        Exit Sub
    End If
    
    If (Subasta.UltimaOferta.Oro * 1.1) < valor Then
        Call WriteConsoleMsg(ui, "Subasta> Tu oferta debe ser al menos 10% mayor que la anterior.", FontTypeNames.FONTTYPE_INFO)
        Exit Sub
    End If
    
    UserList(ui).Stats.Gld = UserList(ui).Stats.Gld - valor
    Call WriteUpdateUserStats(ui)
    
    'le devolvemos al oro al user que ofrecio el oro antes.
    If LenB(Subasta.UltimaOferta.uName) > 0 Then
        tmpUI = NameIndex(Subasta.UltimaOferta.uName)
        If tmpUI > 0 Then
            UserList(tmpUI).Stats.Gld = UserList(tmpUI).Stats.Gld + Subasta.UltimaOferta.Oro
            Call WriteUpdateUserStats(tmpUI)
        Else
            ' el usuario esta offline, se lo tendremos que agregar al charfile/db
            If Not Database_Enabled Then
                tmpOro = CLng(GetVar(CharPath & UCase$(Subasta.UltimaOferta.uName) & ".chr", "STATS", "GLD"))
                Call WriteVar(CharPath & UCase$(Subasta.UltimaOferta.uName) & ".chr", "STATS", "GLD", tmpOro + Subasta.UltimaOferta.Oro)
            Else
                SumGoldtoUserDB Subasta.UltimaOferta.uName, Subasta.UltimaOferta.Oro
            End If
        End If
    End If
    
    Subasta.UltimaOferta.Oro = valor
    Subasta.UltimaOferta.uID = UserList(ui).ID
    Subasta.UltimaOferta.uName = UserList(ui).Name
    If Subasta.MinutosRestantes <= 3 Then
        Subasta.MinutosRestantes = Subasta.MinutosRestantes + 3
        Call MensajeGlobal("Subasta> Se posterga 3 minutos la finalizacion de la subasta.", FontTypeNames.FONTTYPE_GUILD)
    End If
    
    Call MensajeGlobal("Subasta> " & UserList(ui).Name & " ha ofrecido " & valor, FontTypeNames.FONTTYPE_GUILD)
    
End Sub

Public Sub IniciaSubasta(ByVal Userindex As Integer, ByVal Slot As Byte, ByVal amount As Integer, ByVal PrecioInicial As Long)
    Dim iObj As Integer
    iObj = UserList(Userindex).Invent.Object(Slot).ObjIndex
        
    If Subasta.Activo = True Then
        Call WriteConsoleMsg(Userindex, "Subasta> Ya hay una subasta en curso", FontTypeNames.FONTTYPE_INFO)
        Exit Sub
    End If
    
    If itemValido(iObj) = False Then
        Call WriteConsoleMsg(Userindex, "Subasta> El objeto no se puede subastar", FontTypeNames.FONTTYPE_INFO)
        Exit Sub
    End If
    
    If amountValida(Userindex, Slot, amount) = False Then
        Call WriteConsoleMsg(Userindex, "Subasta> No tienes esa cantidad", FontTypeNames.FONTTYPE_INFO)
        Exit Sub
    End If
    
    Call QuitarObjetos(iObj, amount, Userindex)
    
    Subasta.MinutosRestantes = 5
    Subasta.Subastador = UserList(Userindex).Name
    Subasta.subastadorID = UserList(Userindex).ID
    Subasta.Item.ObjIndex = iObj
    Subasta.Item.amount = amount
    Subasta.UltimaOferta.Oro = PrecioInicial
    Subasta.Activo = True
    
    Call MensajeGlobal("Subasta> " & UserList(Userindex).Name & " esta subastando " & amount & " - " & ObjData(iObj).Name & _
                                    " con una oferta inicial de: " & PrecioInicial & ". Escribe '/OFRECER Cantidad ' para participar de la subasta", FontTypeNames.FONTTYPE_GUILD)

    
End Sub






    
