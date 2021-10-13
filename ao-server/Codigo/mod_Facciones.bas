Attribute VB_Name = "mod_Facciones"
Option Explicit

Private Type tPremios
    ItemsAltos() As obj
    ItemsBajos() As obj
End Type

Private Type tJerarquia
    
    iQuestTerminada As Integer
    
    LegionesMatados As Integer
    ArmadasMatados As Integer
    
    Premios() As obj
    PremiosClase() As tPremios
    numPremios As Byte
    
    Titulo As String
    
End Type

Private Type tConfigFacciones
    Jerarquia() As tJerarquia
    CiudadInicial As WorldPos
    TituloNoEnlistado As String
    NumJerarquias As Byte
End Type


Public Faccion(1 To 2) As tConfigFacciones

Public Function esArmada(ByVal ui As Integer) As Boolean
    esArmada = (UserList(ui).Faccion.bando = eFaccion.Armada)
End Function

Public Function esLegion(ByVal ui As Integer) As Boolean
    esLegion = (UserList(ui).Faccion.bando = eFaccion.Legion)
End Function

Public Function esNeutral(ByVal ui As Integer) As Boolean
    esNeutral = (UserList(ui).Faccion.bando = eFaccion.Neutral)
End Function
Public Sub CargarFacciones()
    Dim path As String, leer As clsIniManager
    Dim Main As String
    path = DatPath & "Facciones.dat"
    
    Set leer = New clsIniManager
    
    leer.Initialize path
    Dim i As Long, X As Long, Y As Long, tmpStr As String
    Dim z As Long
    For i = 1 To 2
        With Faccion(i)
            If i = 1 Then Main = "ARMADA" Else Main = "LEGION"
            
            tmpStr = leer.GetValue(Main, "CiudadInicial")
            .CiudadInicial.Map = ReadField(1, tmpStr, 45)
            .CiudadInicial.X = ReadField(1, tmpStr, 45)
            .CiudadInicial.Y = ReadField(1, tmpStr, 45)
            .TituloNoEnlistado = leer.GetValue(Main, "TituloNoEnlistado")
            .NumJerarquias = leer.GetValue(Main, "NumJerarquias")
            
            ReDim .Jerarquia(1 To .NumJerarquias)
            
            For X = 1 To .NumJerarquias
                With .Jerarquia(X)
                    .ArmadasMatados = leer.GetValue(Main & "JERARQUIA" & X, "ARMADASMATADOS")
                    .LegionesMatados = leer.GetValue(Main & "JERARQUIA" & X, "LEGIONESMATADOS")
                    .numPremios = CByte(leer.GetValue(Main & "JERARQUIA" & X, "NumPremios"))
                    If .numPremios > 0 Then
                        ReDim .Premios(1 To .numPremios)
                        For Y = 1 To .numPremios
                            tmpStr = leer.GetValue(Main & "JERARQUIA" & X, "PREMIO" & Y)
                            .Premios(Y).ObjIndex = CInt(ReadField(1, tmpStr, 45))
                            .Premios(Y).amount = CInt(ReadField(2, tmpStr, 45))
                        Next Y
                    End If
                    
                    For Y = 1 To NUMCLASES
                        tmpStr = leer.GetValue(Main & "JERARQUIA" & X, "ObjsClase" & Y & "Altos")
                        .PremiosClase(Y).ItemsAltos = Split(tmpStr)
                        For z = LBound(.PremiosClase(Y).ItemsAltos) To UBound(.PremiosClase(Y).ItemsAltos)
                            .PremiosClase(Y).ItemsAltos(z).amount = 1
                        Next z
                        
                        tmpStr = leer.GetValue(Main & "JERARQUIA" & X, "ObjsClase" & Y & "Bajos")
                        .PremiosClase(Y).ItemsBajos = Split(tmpStr)
                        For z = LBound(.PremiosClase(Y).ItemsBajos) To UBound(.PremiosClase(Y).ItemsBajos)
                            .PremiosClase(Y).ItemsBajos(z).amount = 1
                        Next z
                    Next Y
                    
                    .Titulo = leer.GetValue(Main & "JERARQUIA" & X, "Titulo")
                    .iQuestTerminada = leer.GetValue(Main & "JERARQUIA" & X, "QuestTerminada")
                    
                End With
            Next X
        End With
    Next i
    
    Set leer = Nothing
End Sub











