Attribute VB_Name = "mDx8_Dibujado"
Option Explicit

' Dano en Render
Private Const DAMAGE_TIME As Integer = 1000
Private Const DAMAGE_OFFSET As Integer = 20
Private Const DAMAGE_FONT_S As Byte = 12
 
Private Enum EDType
     edPunal = 1    'Apunalo.
     edNormal = 2   'Hechizo o golpe comï¿½n.
     edCritico = 3  'Golpe Critico
     edFallo = 4    'Fallo el ataque
     edCurar = 5    'Curacion a usuario
     edTrabajo = 6  'Cantidad de items obtenidas a partir del trabajo realizado
End Enum
 
Private DNormalFont    As New StdFont
 
Type DList
     DamageVal      As Integer      'Cantidad de daï¿½o.
     ColorRGB       As Long         'Color.
     DamageType     As EDType       'Tipo, se usa para saber si es apu o no.
     DamageFont     As New StdFont  'Efecto del apu.
     StartedTime    As Long         'Cuando fue creado.
     Downloading    As Byte         'Contador para la posicion Y.
     Activated      As Boolean      'Si esta activado..
End Type

Private DrawBuffer As cDIBSection

Sub DrawGrhtoHdc(ByRef Pic As PictureBox, _
                 ByVal GrhIndex As Long, _
                 ByRef DestRect As RECT)

    '*****************************************************************
    'Draws a Grh's portion to the given area of any Device Context
    '*****************************************************************
         
    DoEvents
    
    Pic.AutoRedraw = False
        
    'Clear the inventory window
    Call Engine_BeginScene
        
    Call Draw_GrhIndex(GrhIndex, 0, 0, 0, Normal_RGBList())
        
    Call Engine_EndScene(DestRect, Pic.hwnd)
    
    Call DrawBuffer.LoadPictureBlt(Pic.hdc)

    Pic.AutoRedraw = True

    Call DrawBuffer.PaintPicture(Pic.hdc, 0, 0, Pic.Width, Pic.Height, 0, 0, vbSrcCopy)

    Pic.Picture = Pic.Image
        
End Sub

Public Sub PrepareDrawBuffer()
    Set DrawBuffer = New cDIBSection
    'El tamanio del buffer es arbitrario = 1024 x 1024
    Call DrawBuffer.Create(1024, 1024)
End Sub

Public Sub CleanDrawBuffer()
    Set DrawBuffer = Nothing
End Sub

Public Sub DrawPJ(ByVal Index As Byte)

    If LenB(cPJ(Index).Nombre) = 0 Then Exit Sub
    DoEvents
    
    Dim cColor       As Long
    Dim Head_OffSet  As Integer
    Dim PixelOffsetX As Integer
    Dim PixelOffsetY As Integer
    Dim RE           As RECT
    
    If cPJ(Index).GameMaster Then
        cColor = 2004510
    Else
        cColor = IIf(cPJ(Index).Criminal, 255, 16744448)
    End If
    
    With frmPanelAccount.lblAccData(Index)
        .Caption = cPJ(Index).Nombre
        .ForeColor = cColor
    End With
    
    With frmPanelAccount.picChar(Index - 1)
        RE.Left = 0
        RE.Top = 0
        RE.Bottom = .Height
        RE.Right = .Width
    End With

    PixelOffsetX = RE.Right \ 2 - 16
    PixelOffsetY = RE.Bottom \ 2
    
    Call Engine_BeginScene
    
    With cPJ(Index)
    
        If .Body <> 0 Then

            Call Draw_Grh(BodyData(.Body).Walk(3), PixelOffsetX, PixelOffsetY, 1, Normal_RGBList(), 0)

            If .Head <> 0 Then
                Call Draw_Grh(HeadData(.Head).Head(3), PixelOffsetX + BodyData(.Body).HeadOffset.X, PixelOffsetY + BodyData(.Body).HeadOffset.Y, 1, Normal_RGBList(), 0)
            End If

            If .helmet <> 0 Then
                Call Draw_Grh(CascoAnimData(.helmet).Head(3), PixelOffsetX + BodyData(.Body).HeadOffset.X, PixelOffsetY + BodyData(.Body).HeadOffset.Y + OFFSET_HEAD, 1, Normal_RGBList(), 0)
            End If

            If .weapon <> 0 Then
                Call Draw_Grh(WeaponAnimData(.weapon).WeaponWalk(3), PixelOffsetX, PixelOffsetY, 1, Normal_RGBList(), 0)
            End If

            If .shield <> 0 Then
                Call Draw_Grh(ShieldAnimData(.shield).ShieldWalk(3), PixelOffsetX, PixelOffsetY, 1, Normal_RGBList(), 0)
            End If
        
        End If
    
    End With

    Call Engine_EndScene(RE, frmPanelAccount.picChar(Index - 1).hwnd)

    Call DrawBuffer.LoadPictureBlt(frmPanelAccount.picChar(Index - 1).hdc)

    frmPanelAccount.picChar(Index - 1).AutoRedraw = True

    Call DrawBuffer.PaintPicture(frmPanelAccount.picChar(Index - 1).hdc, 0, 0, RE.Right, RE.Bottom, 0, 0, vbSrcCopy)

    frmPanelAccount.picChar(Index - 1).Picture = frmPanelAccount.picChar(Index - 1).Image
    
End Sub

Sub Damage_Initialize()

    ' Inicializamos el dano en render
    With DNormalFont
        .Size = 20
        .italic = False
        .bold = False
        .name = "Tahoma"
    End With

End Sub

Sub Damage_Create(ByVal X As Byte, _
                  ByVal Y As Byte, _
                  ByVal ColorRGB As Long, _
                  ByVal DamageValue As Integer, _
                  ByVal edMode As Byte)
 
    ' @ Agrega un nuevo dano.
 
    With MapData(X, Y).Damage
     
        .Activated = True
        .ColorRGB = ColorRGB
        .DamageType = edMode
        .DamageVal = DamageValue
        .StartedTime = GetTickCount
        .Downloading = 0
     
        Select Case .DamageType
        
            Case EDType.edPunal

                With .DamageFont
                    .Size = Val(DAMAGE_FONT_S)
                    .name = "Tahoma"
                    .bold = False
                    Exit Sub

                End With
            
        End Select
     
        .DamageFont = DNormalFont
        .DamageFont.Size = 14
     
    End With
 
End Sub

Private Function EaseOutCubic(Time As Double)
    Time = Time - 1
    EaseOutCubic = Time * Time * Time + 1
End Function
 
Sub Damage_Draw(ByVal X As Byte, _
                ByVal Y As Byte, _
                ByVal PixelX As Integer, _
                ByVal PixelY As Integer)
 
    ' @ Dibuja un dano
 
    With MapData(X, Y).Damage
     
        If (Not .Activated) Or (Not .DamageVal <> 0) Then Exit Sub
        
        Dim ElapsedTime As Long
        ElapsedTime = GetTickCount - .StartedTime
        
        If ElapsedTime < DAMAGE_TIME Then
           
            .Downloading = EaseOutCubic(ElapsedTime / DAMAGE_TIME) * DAMAGE_OFFSET
           
            .ColorRGB = Damage_ModifyColour(.DamageType)
           
            'Efectito para el apu
            If .DamageType = EDType.edPunal Then
                .DamageFont.Size = Damage_NewSize(ElapsedTime)

            End If
               
            'Dibujo
            Select Case .DamageType
            
                Case EDType.edCritico
                    Call DrawText(PixelX, PixelY - .Downloading, .DamageVal & "!!", .ColorRGB)
                
                Case EDType.edCurar
                    Call DrawText(PixelX, PixelY - .Downloading, "+" & .DamageVal, .ColorRGB)
                
                Case EDType.edTrabajo
                    Call DrawText(PixelX, PixelY - .Downloading, "+" & .DamageVal, .ColorRGB)
                    
                Case EDType.edFallo
                    Call DrawText(PixelX, PixelY - .Downloading, "Fallo", .ColorRGB)
                    
                Case Else 'EDType.edNormal
                    Call DrawText(PixelX, PixelY - .Downloading, "-" & .DamageVal, .ColorRGB)
                    
            End Select
            
        'Si llego al tiempo lo limpio
        Else
            Damage_Clear X, Y
           
        End If
       
    End With
 
End Sub
 
Sub Damage_Clear(ByVal X As Byte, ByVal Y As Byte)
 
    ' @ Limpia todo.
 
    With MapData(X, Y).Damage
        .Activated = False
        .ColorRGB = 0
        .DamageVal = 0
        .StartedTime = 0

    End With
 
End Sub
 
Function Damage_ModifyColour(ByVal DamageType As Byte) As Long
 
    ' @ Se usa para el "efecto" de desvanecimiento.
 
    Select Case DamageType
                   
        Case EDType.edPunal
            Damage_ModifyColour = ColoresDano(52)
            
        Case EDType.edFallo
            Damage_ModifyColour = ColoresDano(54)
            
        Case EDType.edCurar
            Damage_ModifyColour = ColoresDano(55)
        
        Case EDType.edTrabajo
            Damage_ModifyColour = ColoresDano(56)
            
        Case Else 'EDType.edNormal
            Damage_ModifyColour = ColoresDano(51)
            
    End Select
 
End Function
 
Function Damage_NewSize(ByVal ElapsedTime As Long) As Byte
 
    ' @ Se usa para el "efecto" del apu.

    ' Nos basamos en la constante DAMAGE_TIME
    Select Case ElapsedTime
 
        Case Is <= DAMAGE_TIME / 5
            Damage_NewSize = 14
       
        Case Is <= DAMAGE_TIME * 2 / 5
            Damage_NewSize = 13
           
        Case Is <= DAMAGE_TIME * 3 / 5
            Damage_NewSize = 12
           
        Case Else
            Damage_NewSize = 11
       
    End Select
 
End Function


Public Function Geometry_Create_TLVertex(ByVal X As Single, ByVal Y As Single, ByVal Z As Single, _
                                            ByVal rhw As Single, ByVal color As Long, ByVal Specular As Long, tu As Single, _
                                            ByVal tv As Single) As TLVERTEX
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 10/07/2002
'**************************************************************
    Geometry_Create_TLVertex.X = X
    Geometry_Create_TLVertex.Y = Y
    Geometry_Create_TLVertex.Z = Z
    Geometry_Create_TLVertex.rhw = rhw
    Geometry_Create_TLVertex.color = color
    Geometry_Create_TLVertex.Specular = Specular
    Geometry_Create_TLVertex.tu = tu
    Geometry_Create_TLVertex.tv = tv
End Function

Public Sub Draw_FilledBox(ByVal X As Integer, ByVal Y As Integer, ByVal Width As Integer, ByVal Height As Integer, color As Long, outlinecolor As Long)
 
    Static box_rect As RECT
    Static Outline As RECT
    Static rgb_list(3) As Long
    Static rgb_list2(3) As Long
    Static Vertex(3) As TLVERTEX
    Static Vertex2(3) As TLVERTEX
    
    rgb_list(0) = color
    rgb_list(1) = color
    rgb_list(2) = color
    rgb_list(3) = color
    
    rgb_list2(0) = outlinecolor
    rgb_list2(1) = outlinecolor
    rgb_list2(2) = outlinecolor
    rgb_list2(3) = outlinecolor
    
    With box_rect
        .Bottom = Y + Height - 1
        .Left = X
        .Right = X + Width - 1
        .Top = Y
    End With
    
    With Outline
        .Bottom = Y + Height
        .Left = X
        .Right = X + Width
        .Top = Y
    End With
    
    Geometry_Create_Box Vertex2(), Outline, Outline, rgb_list2(), 0, 0
    Geometry_Create_Box Vertex(), box_rect, box_rect, rgb_list(), 0, 0
    
    DirectDevice.SetTexture 0, Nothing
    DirectDevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, Vertex2(0), Len(Vertex2(0))
    DirectDevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, Vertex(0), Len(Vertex(0))
End Sub


Public Sub Geometry_Create_Box(ByRef verts() As TLVERTEX, ByRef dest As RECT, ByRef src As RECT, ByRef rgb_list() As Long, _
                                Optional ByRef Textures_Width As Long, Optional ByRef Textures_Height As Long, Optional ByVal Angle As Single)
'**************************************************************
'Author: Aaron Perkins
'Modified by Juan Martín Sotuyo Dodero
'Last Modify Date: 11/17/2002
'
' * v1      * v3
' |\        |
' |  \      |
' |    \    |
' |      \  |
' |        \|
' * v0      * v2
'**************************************************************
    Dim x_center As Single
    Dim y_center As Single
    Dim radius As Single
    Dim x_Cor As Single
    Dim y_Cor As Single
    Dim left_point As Single
    Dim right_point As Single
    Dim temp As Single
   
    If Angle > 0 Then
        'Center coordinates on screen of the square
        x_center = dest.Left + (dest.Right - dest.Left) / 2
        y_center = dest.Top + (dest.Bottom - dest.Top) / 2
       
        'Calculate radius
        radius = Sqr((dest.Right - x_center) ^ 2 + (dest.Bottom - y_center) ^ 2)
       
        'Calculate left and right points
        temp = (dest.Right - x_center) / radius
        right_point = Atn(temp / Sqr(-temp * temp + 1))
        left_point = 3.1459 - right_point
    End If
   
    'Calculate screen coordinates of sprite, and only rotate if necessary
    If Angle = 0 Then
        x_Cor = dest.Left
        y_Cor = dest.Bottom
    Else
        x_Cor = x_center + Cos(-left_point - Angle) * radius
        y_Cor = y_center - Sin(-left_point - Angle) * radius
    End If
   
   
    '0 - Bottom left vertex
    If Textures_Width Or Textures_Height Then
        verts(2) = Geometry_Create_TLVertex(x_Cor, y_Cor, 0, 1, rgb_list(0), 0, src.Left / Textures_Width + 0.001, (src.Bottom + 1) / Textures_Height)
    Else
        verts(2) = Geometry_Create_TLVertex(x_Cor, y_Cor, 0, 1, rgb_list(0), 0, 0, 0)
    End If
    'Calculate screen coordinates of sprite, and only rotate if necessary
    If Angle = 0 Then
        x_Cor = dest.Left
        y_Cor = dest.Top
    Else
        x_Cor = x_center + Cos(left_point - Angle) * radius
        y_Cor = y_center - Sin(left_point - Angle) * radius
    End If
   
   
    '1 - Top left vertex
    If Textures_Width Or Textures_Height Then
        verts(0) = Geometry_Create_TLVertex(x_Cor, y_Cor, 0, 1, rgb_list(1), 0, src.Left / Textures_Width + 0.001, src.Top / Textures_Height + 0.001)
    Else
        verts(0) = Geometry_Create_TLVertex(x_Cor, y_Cor, 0, 1, rgb_list(1), 0, 0, 1)
    End If
    'Calculate screen coordinates of sprite, and only rotate if necessary
    If Angle = 0 Then
        x_Cor = dest.Right
        y_Cor = dest.Bottom
    Else
        x_Cor = x_center + Cos(-right_point - Angle) * radius
        y_Cor = y_center - Sin(-right_point - Angle) * radius
    End If
   
   
    '2 - Bottom right vertex
    If Textures_Width Or Textures_Height Then
        verts(3) = Geometry_Create_TLVertex(x_Cor, y_Cor, 0, 1, rgb_list(2), 0, (src.Right + 1) / Textures_Width, (src.Bottom + 1) / Textures_Height)
    Else
        verts(3) = Geometry_Create_TLVertex(x_Cor, y_Cor, 0, 1, rgb_list(2), 0, 1, 0)
    End If
    'Calculate screen coordinates of sprite, and only rotate if necessary
    If Angle = 0 Then
        x_Cor = dest.Right
        y_Cor = dest.Top
    Else
        x_Cor = x_center + Cos(right_point - Angle) * radius
        y_Cor = y_center - Sin(right_point - Angle) * radius
    End If
   
   
    '3 - Top right vertex
    If Textures_Width Or Textures_Height Then
        verts(1) = Geometry_Create_TLVertex(x_Cor, y_Cor, 0, 1, rgb_list(3), 0, (src.Right + 1) / Textures_Width, src.Top / Textures_Height + 0.001)
    Else
        verts(1) = Geometry_Create_TLVertex(x_Cor, y_Cor, 0, 1, rgb_list(3), 0, 1, 1)
    End If

End Sub

