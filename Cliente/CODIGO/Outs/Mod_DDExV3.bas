Attribute VB_Name = "Mod_DDExV3"
Option Explicit

Public DDExV3 As New cls_Motor

Public Type tDDEXRGBA
    r As Byte
    g As Byte
    b As Byte
    a As Byte
End Type

Public Fuente1 As Long
Public Fuente2 As Long
Public Fuente3 As Long

Public Function DDEXRGBA(a As Byte, r As Byte, g As Byte, b As Byte) As tDDEXRGBA
    DDEXRGBA.a = a
    DDEXRGBA.b = b
    DDEXRGBA.g = g
    DDEXRGBA.r = r
End Function

Public Sub InitDDExV3()
    DDExV3.Iniciar frmMain.RenderDDExV3.hwnd, "Graficos", True
    Fuente1 = DDExV3.CrearFuente("Verdana", 20)
    Fuente2 = DDExV3.CrearFuente("Arial", 20, 1)
    Fuente3 = DDExV3.CrearFuente("MS Sans Serif", 8, 1)
End Sub

Private Sub DDExV3_GrhRender(ByRef Grh As Grh, ByVal X As Integer, ByVal Y As Integer, ByVal Center As Byte, ByVal Animate As Byte, Optional ByVal Angle As Long)
Dim CurrentGrhIndex As Integer
Dim SourceRect As RECT
        
On Error GoTo Finish
    If Animate Then
        If Grh.Started = 1 Then
            Grh.FrameCounter = Grh.FrameCounter + (timerElapsedTime * GrhData(Grh.GrhIndex).NumFrames / Grh.Speed)
            If Grh.FrameCounter > GrhData(Grh.GrhIndex).NumFrames Then
                Grh.FrameCounter = (Grh.FrameCounter Mod GrhData(Grh.GrhIndex).NumFrames) + 1
                
                If Grh.Loops <> -1 Then
                    If Grh.Loops > 0 Then
                        Grh.Loops = Grh.Loops - 1
                    Else
                        Grh.Started = 0
                    End If
                End If
            End If
        End If
    End If
    
    CurrentGrhIndex = GrhData(Grh.GrhIndex).Frames(Grh.FrameCounter)
    
    With GrhData(CurrentGrhIndex)
        If Center Then
            If .TileWidth <> 1 Then
                X = X - Int(.TileWidth * TilePixelWidth / 2) + TilePixelWidth \ 2
            End If
            
            If .TileHeight <> 1 Then
                Y = Y - Int(.TileHeight * TilePixelHeight) + TilePixelHeight
            End If
        End If
        
        SourceRect.Left = .sX
        SourceRect.Top = .sY
        SourceRect.Right = SourceRect.Left + .pixelWidth
        SourceRect.bottom = SourceRect.Top + .pixelHeight
        
        'Draw
        If Angle <> 0 Then
            Dim color As tDDEXRGBA
            color = DDEXRGBA(200, 255, 255, 255)
            DDExV3.DBEx .FileNum, SourceRect, X, Y, color, Angle
        Else
            Dim Color2 As tDDEXRGBA
            Color2.a = 255
            Color2.r = Estado_Actual.r
            Color2.g = Estado_Actual.g
            Color2.b = Estado_Actual.b

            DDExV3.DBAlfa .FileNum, SourceRect, X, Y, Color2
        End If
        
    End With
    
Finish:
    If Err.Number = 9 And Grh.FrameCounter < 1 Then
        Grh.FrameCounter = 1
        Resume
    End If
End Sub

Sub DDExV3_GrhIndexRender(ByVal GrhIndex As Integer, ByVal X As Integer, ByVal Y As Integer, ByVal Center As Byte, ByRef Color_List() As Long, Optional ByVal Angle As Single = 0, Optional ByVal Alpha As Boolean = False)
    Dim SourceRect As RECT
    
    With GrhData(GrhIndex)
        If Center Then
            If .TileWidth <> 1 Then
                X = X - Int(.TileWidth * TilePixelWidth / 2) + TilePixelWidth \ 2
            End If
            
            If .TileHeight <> 1 Then
                Y = Y - Int(.TileHeight * TilePixelHeight) + TilePixelHeight
            End If
        End If
        
        SourceRect.Left = .sX
        SourceRect.Top = .sY
        SourceRect.Right = SourceRect.Left + .pixelWidth
        SourceRect.bottom = SourceRect.Top + .pixelHeight
        
        'Draw
        DDExV3.DBGrafico .FileNum, X, Y, SourceRect, True
        
    End With
End Sub

Sub DDExV3_RenderScreen(ByVal TileX As Integer, ByVal TileY As Integer, ByVal PixelOffsetX As Integer, ByVal PixelOffsetY As Integer, ByVal timerTicksPerFrame As Single)
    Dim Y           As Long     'Keeps track of where on map we are
    Dim X           As Long     'Keeps track of where on map we are
    Dim screenminY  As Integer  'Start Y pos on current screen
    Dim screenmaxY  As Integer  'End Y pos on current screen
    Dim screenminX  As Integer  'Start X pos on current screen
    Dim screenmaxX  As Integer  'End X pos on current screen
    Dim minY        As Integer  'Start Y pos on current map
    Dim maxY        As Integer  'End Y pos on current map
    Dim minX        As Integer  'Start X pos on current map
    Dim maxX        As Integer  'End X pos on current map
    Dim ScreenX     As Integer  'Keeps track of where to place tile on screen
    Dim ScreenY     As Integer  'Keeps track of where to place tile on screen
    Dim minXOffset  As Integer
    Dim minYOffset  As Integer
    Dim PixelOffsetXTemp As Integer 'For centering grhs
    Dim PixelOffsetYTemp As Integer 'For centering grhs
    Dim ColorTechos(3) As Long

    screenminY = TileY - HalfWindowTileHeight
    screenmaxY = TileY + HalfWindowTileHeight
    screenminX = TileX - HalfWindowTileWidth
    screenmaxX = TileX + HalfWindowTileWidth
    
    minY = screenminY - TileBufferSize
    maxY = screenmaxY + TileBufferSize
    minX = screenminX - TileBufferSize
    maxX = screenmaxX + TileBufferSize
    
    If minY < XMinMapSize Then
        minYOffset = YMinMapSize - minY
        minY = YMinMapSize
    End If
    
    If maxY > YMaxMapSize Then maxY = YMaxMapSize
    
    If minX < XMinMapSize Then
        minXOffset = XMinMapSize - minX
        minX = XMinMapSize
    End If
    
    If maxX > XMaxMapSize Then maxX = XMaxMapSize
    
    If screenminY > YMinMapSize Then
        screenminY = screenminY - 1
    Else
        screenminY = 1
        ScreenY = 1
    End If
    
    If screenmaxY < YMaxMapSize Then screenmaxY = screenmaxY + 1
    
    If screenminX > XMinMapSize Then
        screenminX = screenminX - 1
    Else
        screenminX = 1
        ScreenX = 1
    End If
    
    If screenmaxX < XMaxMapSize Then screenmaxX = screenmaxX + 1
    
    DDExV3.LimpiarPantalla
        
    For Y = screenminY To screenmaxY
        For X = screenminX To screenmaxX
            'Layer 1
            Call DDExV3_GrhRender(MapData(X, Y).Graphic(1), (ScreenX - 1) * TilePixelWidth + PixelOffsetX, (ScreenY - 1) * TilePixelHeight + PixelOffsetY, 0, 1)

            'Layer 2
            If MapData(X, Y).Graphic(2).GrhIndex <> 0 Then
                Call DDExV3_GrhRender(MapData(X, Y).Graphic(2), (ScreenX - 1) * TilePixelWidth + PixelOffsetX, (ScreenY - 1) * TilePixelHeight + PixelOffsetY, 0, 0)
            End If
            
            ScreenX = ScreenX + 1
        Next X
        ScreenX = ScreenX - X + screenminX
        ScreenY = ScreenY + 1
    Next Y
    
    '<----- Layer Obj, Char, 3 ----->
    ScreenY = minYOffset - TileBufferSize
    For Y = minY To maxY
        ScreenX = minXOffset - TileBufferSize
        For X = minX To maxX
            PixelOffsetXTemp = ScreenX * TilePixelWidth + PixelOffsetX
            PixelOffsetYTemp = ScreenY * TilePixelHeight + PixelOffsetY
            
            With MapData(X, Y)
                'Object Layer
                If .ObjGrh.GrhIndex <> 0 Then
                    Call DDExV3_GrhRender(.ObjGrh, PixelOffsetXTemp, PixelOffsetYTemp, 1, 0)
                End If
                
                'Char layer
                If .CharIndex <> 0 Then
                    Call DDExV3_CharRender(.CharIndex, PixelOffsetXTemp, PixelOffsetYTemp, timerTicksPerFrame)
                End If
                
                'Layer 3
                If .Graphic(3).GrhIndex <> 0 Then
                    Call DDExV3_GrhRender(.Graphic(3), PixelOffsetXTemp, PixelOffsetYTemp, 1, 1)
                End If
            End With
            
            ScreenX = ScreenX + 1
        Next X
        ScreenY = ScreenY + 1
    Next Y
    
    '<----- Layer 4 ----->
        ScreenY = minYOffset - TileBufferSize
        For Y = minY To maxY
            ScreenX = minXOffset - TileBufferSize
            For X = minX To maxX
                If Abs(MouseTileX - X) < 1 And (Abs(MouseTileY - Y)) < 1 And Settings.NombreItems And MapData(X, Y).OBJInfo.Name <> "" Then
                    Engine_Draw_FillBox ScreenX * TilePixelWidth + PixelOffsetX, ScreenY * TilePixelHeight + PixelOffsetY, Fonts_Render_String_Width(MapData(X, Y).OBJInfo.Name, 3), 14, D3DColorARGB(0, 0, 0, 0), D3DColorARGB(50, 0, 0, 0)

                    Fonts_Render_String MapData(X, Y).OBJInfo.Name, ScreenX * TilePixelWidth + PixelOffsetX, ScreenY * TilePixelHeight + PixelOffsetY, D3DColorARGB(100, 255, 255, 255), 3
                End If
                
                'Layer 4
                If Not bTecho Then
                    If MapData(X, Y).Graphic(4).GrhIndex Then
                        Call DDExV3_GrhRender(MapData(X, Y).Graphic(4), ScreenX * TilePixelWidth + PixelOffsetX, ScreenY * TilePixelHeight + PixelOffsetY, 1, 0)
                    End If
                End If
                ScreenX = ScreenX + 1
            Next X
            ScreenY = ScreenY + 1
        Next Y

    DDExV3.MostrarPantalla
End Sub

Private Sub DDExV3_CharRender(ByVal CharIndex As Long, ByVal PixelOffsetX As Integer, ByVal PixelOffsetY As Integer, ByVal timerTicksPerFrame As Single)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 12/03/04
'Draw char's to screen without offcentering them
'***************************************************
    Dim moved As Boolean

    With CharList(CharIndex)
        If .Moving Then
            'If needed, move left and right
            If .scrollDirectionX <> 0 Then
                .MoveOffsetX = .MoveOffsetX + ScrollPixelsPerFrameX * Sgn(.scrollDirectionX) * timerTicksPerFrame
                
                'Start animations
'TODO : Este parche es para evita los uncornos exploten al moverse!! REVER!!!
                If .Body.Walk(.Heading).Speed > 0 Then _
                    .Body.Walk(.Heading).Started = 1
                .Arma.WeaponWalk(.Heading).Started = 1
                .Escudo.ShieldWalk(.Heading).Started = 1
                
                'Char moved
                moved = True
                
                'Check if we already got there
                If (Sgn(.scrollDirectionX) = 1 And .MoveOffsetX >= 0) Or _
                        (Sgn(.scrollDirectionX) = -1 And .MoveOffsetX <= 0) Then
                    .MoveOffsetX = 0
                    .scrollDirectionX = 0
                End If
            End If
            
            'If needed, move up and down
            If .scrollDirectionY <> 0 Then
                .MoveOffsetY = .MoveOffsetY + ScrollPixelsPerFrameY * Sgn(.scrollDirectionY) * timerTicksPerFrame
                
                'Start animations
'TODO : Este parche es para evita los uncornos exploten al moverse!! REVER!!!
                If .Body.Walk(.Heading).Speed > 0 Then _
                    .Body.Walk(.Heading).Started = 1
                .Arma.WeaponWalk(.Heading).Started = 1
                .Escudo.ShieldWalk(.Heading).Started = 1
                
                'Char moved
                moved = True
                
                'Check if we already got there
                If (Sgn(.scrollDirectionY) = 1 And .MoveOffsetY >= 0) Or _
                        (Sgn(.scrollDirectionY) = -1 And .MoveOffsetY <= 0) Then
                    .MoveOffsetY = 0
                    .scrollDirectionY = 0
                End If
            End If
        End If
        
        'If done moving stop animation
        If Not moved Then
            'Stop animations
            .Body.Walk(.Heading).Started = 0
            .Body.Walk(.Heading).FrameCounter = 1
            
            .Arma.WeaponWalk(.Heading).Started = 0
            .Arma.WeaponWalk(.Heading).FrameCounter = 1
            
            .Escudo.ShieldWalk(.Heading).Started = 0
            .Escudo.ShieldWalk(.Heading).FrameCounter = 1
            
            .Moving = False
        End If
        
        PixelOffsetX = PixelOffsetX + .MoveOffsetX
        PixelOffsetY = PixelOffsetY + .MoveOffsetY
            
        If .Head.Head(.Heading).GrhIndex Then
            If Not .invisible Then
                
                'Draw Body
                If .Body.Walk(.Heading).GrhIndex Then _
                    DDExV3_GrhRender .Body.Walk(.Heading), PixelOffsetX, PixelOffsetY, 1, 1
                    
                
                'Draw Head
                If .Head.Head(.Heading).GrhIndex Then
                    Call DDExV3_GrhRender(.Head.Head(.Heading), PixelOffsetX + .Body.HeadOffset.X, PixelOffsetY + .Body.HeadOffset.Y, 1, 0)
                    
                    'Draw Helmet
                    'If .Casco.Head(.Heading).GrhIndex Then _
                        Call DDrawTransGrhtoSurface(.Casco.Head(.Heading), PixelOffsetX + .Body.HeadOffset.X, PixelOffsetY + .Body.HeadOffset.Y + OFFSET_HEAD, 1, ColorFinal(), 0, .Pos.X, .Pos.Y)
                    
                    'Draw Weapon
                    'If .Arma.WeaponWalk(.Heading).GrhIndex Then _
                        Call DDrawTransGrhtoSurface(.Arma.WeaponWalk(.Heading), PixelOffsetX, PixelOffsetY, 1, ColorFinal(), 1, .Pos.X, .Pos.Y)
                    
                    'Draw Shield
                    'If .Escudo.ShieldWalk(.Heading).GrhIndex Then _
                        Call DDrawTransGrhtoSurface(.Escudo.ShieldWalk(.Heading), PixelOffsetX, PixelOffsetY, 1, ColorFinal(), 1, .Pos.X, .Pos.Y)
                        
                    'Draw name over head
                    If LenB(.Nombre) > 0 Then
                             Call DDExV3_RenderName(CharIndex, PixelOffsetX, PixelOffsetY)
                    End If
                    
                End If
            End If
        Else
            'Draw Body
            If .Body.Walk(.Heading).GrhIndex Then _
                Call DDExV3_GrhRender(.Body.Walk(.Heading), PixelOffsetX, PixelOffsetY, 1, 1, 0)
        End If
        
        'Update dialogs
        Call Dialogos.UpdateDialogPos(PixelOffsetX + .Body.HeadOffset.X - 10, PixelOffsetY + .Body.HeadOffset.Y + OFFSET_HEAD, CharIndex) '34 son los pixeles del grh de la cabeza que quedan superpuestos al cuerpo
        
        'Draw FX
        If .FxIndex <> 0 Then
            DDExV3_GrhRender .fX, PixelOffsetX + FxData(.FxIndex).OffsetX, PixelOffsetY + FxData(.FxIndex).OffsetY, 1, 1
            
            'Check if animation is over
            If .fX.Started = 0 Then _
                .FxIndex = 0
        End If
    End With

End Sub

Private Sub DDExV3_RenderName(ByVal CharIndex As Long, ByVal X As Integer, ByVal Y As Integer)
    Dim Pos As Integer
    Dim line As String
    Dim color As tDDEXRGBA
   
With CharList(CharIndex)
        Pos = getTagPosition(.Nombre)
        
        If .priv = 0 Then
            If .Atacable Then
                color = DDEXRGBA(255, ColoresPJ(48).r, ColoresPJ(48).g, ColoresPJ(48).b)

            Else
                If .muerto Then
                        color = DDEXRGBA(255, 220, 220, 255)
                Else
                    If .Criminal Then
                        color = DDEXRGBA(255, ColoresPJ(50).r, ColoresPJ(50).g, ColoresPJ(50).b)
                    Else
                        color = DDEXRGBA(255, ColoresPJ(49).r, ColoresPJ(49).g, ColoresPJ(49).b)
                    End If
                End If
            End If
        Else
            color = DDEXRGBA(255, ColoresPJ(.priv).r, ColoresPJ(.priv).g, ColoresPJ(.priv).b)
        End If
                            
          'Nick
          line = Left$(.Nombre, Pos - 2)
        '  Fonts_Render_String line, (X + 16) - FormatNumber((Fonts_Render_String_Width(line, Settings.Engine_Font) / 2), 0), Y + 30, Color, Settings.Engine_Font
        DDExV3.DBTexto line, X, Y + 30, color, Fuente3
          'Clan
          line = mid$(.Nombre, Pos)
         ' Fonts_Render_String line, (X + 16) - (Fonts_Render_String_Width(line, Settings.Engine_Font) / 2), Y + 45, D3DColorXRGB(255, 230, 130), Settings.Engine_Font
        DDExV3.DBTexto line, X, Y + 45, color, Fuente3
End With
End Sub
