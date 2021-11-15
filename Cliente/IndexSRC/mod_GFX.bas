Attribute VB_Name = "mod_GFX"
Option Explicit

Public DX8 As dx_GFX_Class
Public DX8_FONT As Long
Public Texture As Long
Public MAPS() As Long

Public BUFFER_TO_DELETE As Long
Public FPSLastCheck As Long
Public FramesPerSecCounter As Long
Public FPSx As Long

Public Function DELETE_BUFFER()
    If BUFFER_TO_DELETE = 0 Then
        DX8.MAP_Unload Texture
    Else
        Dim i As Long
            For i = 1 To BUFFER_TO_DELETE
                DX8.MAP_Unload MAPS(i)
            Next i
    End If
    
    DoEvents
End Function

Public Function RENDER_GRH(ByVal GRH As Long)
    If GRH = 0 Then Exit Function
    DELETE_BUFFER
    
    Dim RECT As GFX_Rect

    If FileExist(BMP_DIRE & GrhData(GRH).FileNum & ".bmp", vbNormal) Then
        Texture = DX8.MAP_Load(BMP_DIRE & GrhData(GRH).FileNum & ".bmp", 0)
    ElseIf FileExist(BMP_DIRE & GrhData(GRH).FileNum & ".png", vbNormal) Then
        Texture = DX8.MAP_Load(BMP_DIRE & GrhData(GRH).FileNum & ".png", 0)
    Else
        MsgBox "LA TEXTURA NO EXISTE"
        Exit Function
    End If
    BUFFER_TO_DELETE = 0
    
    Do While frmMain.Index = GRH And APP_RUN = False
    
        With GrhData(GRH)
                RECT.X = .sX
                RECT.Y = .sY
                RECT.Height = .pixelHeight
                RECT.Width = .pixelWidth
                
            DX8.MAP_SetRegion Texture, RECT
            
            DX8.DRAW_Map Texture, 0, 0, 0, CLng(.pixelWidth), CLng(.pixelHeight)
                    
            If (frmMain.MouseY Or frmMain.MouseX) <> -1 Then
                DX8.DRAW_Line frmMain.MouseX, 0, frmMain.MouseX, 384, 0, DX8.ARGB_Set(255, 255, 0, 0)
                DX8.DRAW_Line 0, frmMain.MouseY, 384, frmMain.MouseY, 0, DX8.ARGB_Set(255, 255, 0, 0)
                DX8.DRAW_Pixel frmMain.MouseX, frmMain.MouseY, 0, DX8.ARGB_Set(255, 255, 255, 255)
                
                DX8.DRAW_Text DX8_FONT, "X: " & frmMain.MouseX & " Y: " & frmMain.MouseY, 300, 350, 0, DX8.ARGB_Set(255, 255, 0, 0), Align_Left
            End If
                DX8.DRAW_Text DX8_FONT, "FPS: " & FPSx, 10, 350, 0, DX8.ARGB_Set(255, 255, 0, 0), Align_Left

            DX8.Frame
         End With

        Call FPS
    Loop

End Function



Public Function RENDER_GRH_ONCE(ByVal GRH As Long)
    If GRH = 0 Then Exit Function
    DELETE_BUFFER
    
    Dim RECT As GFX_Rect

    If FileExist(BMP_DIRE & GrhData(GRH).FileNum & ".bmp", vbNormal) Then
        Texture = DX8.MAP_Load(BMP_DIRE & GrhData(GRH).FileNum & ".bmp", 0)
    ElseIf FileExist(BMP_DIRE & GrhData(GRH).FileNum & ".png", vbNormal) Then
        Texture = DX8.MAP_Load(BMP_DIRE & GrhData(GRH).FileNum & ".png", 0)
    Else
        MsgBox "LA TEXTURA NO EXISTE"
        Exit Function
    End If
    BUFFER_TO_DELETE = 0
    
   ' Do While frmMain.Index = GRH And APP_RUN = False
    
        With GrhData(GRH)
                RECT.X = .sX
                RECT.Y = .sY
                RECT.Height = .pixelHeight
                RECT.Width = .pixelWidth
                
            DX8.MAP_SetRegion Texture, RECT
            
            DX8.DRAW_Map Texture, 0, 0, 0, CLng(.pixelWidth), CLng(.pixelHeight)
                    
           ' If (frmMain.MouseY Or frmMain.MouseX) <> -1 Then
           '     DX8.DRAW_Line frmMain.MouseX, 0, frmMain.MouseX, 384, 0, DX8.ARGB_Set(255, 255, 0, 0)
           '   '''  DX8.DRAW_Line 0, frmMain.MouseY, 384, frmMain.MouseY, 0, DX8.ARGB_Set(255, 255, 0, 0)
           '    '' DX8.DRAW_Pixel frmMain.MouseX, frmMain.MouseY, 0, DX8.ARGB_Set(255, 255, 255, 255)
           '
           '   '  DX8.DRAW_Text DX8_FONT, "X: " & frmMain.MouseX & " Y: " & frmMain.MouseY, 300, 350, 0, DX8.ARGB_Set(255, 255, 0, 0), Align_Left
           ' End If
           '   '  DX8.DRAW_Text DX8_FONT, "FPS: " & FPSx, 10, 350, 0, DX8.ARGB_Set(255, 255, 0, 0), Align_Left'

            DX8.Frame
         End With

        Call FPS
   ' Loop

End Function



Public Function RENDER_ANIM(ByVal GRH As Long)
    If GRH = 0 Then Exit Function
    DELETE_BUFFER
    
    Dim RECT As GFX_Rect

    
    Dim i As Long
    Dim tmp As Long, hWnd_DEST As Long
        
    ReDim MAPS(1 To GrhData(GRH).NumFrames)
    
    For i = 1 To UBound(MAPS())
            If FileExist(BMP_DIRE & GrhData(GrhData(GRH).Frames(1)).FileNum & ".bmp", vbNormal) Then
                MAPS(i) = DX8.MAP_Load(BMP_DIRE & GrhData(GrhData(GRH).Frames(i)).FileNum & ".bmp", 0)
            ElseIf FileExist(BMP_DIRE & GrhData(GrhData(GRH).Frames(1)).FileNum & ".png", vbNormal) Then
                MAPS(i) = DX8.MAP_Load(BMP_DIRE & GrhData(GrhData(GRH).Frames(i)).FileNum & ".png", 0)
            Else
                MsgBox "LA TEXTURA NO EXISTE"
                Exit Function
            End If
    Next i
    
    With GrhData(GRH)

 
    tmp = 0
    BUFFER_TO_DELETE = .NumFrames
    
    Do While frmMain.Index = GRH And APP_RUN = False
        tmp = tmp + 1
        If tmp > .NumFrames Then tmp = 1


            RECT.X = GrhData(.Frames(tmp)).sX
            RECT.Y = GrhData(.Frames(tmp)).sY
            RECT.Height = GrhData(.Frames(tmp)).pixelHeight
            RECT.Width = GrhData(.Frames(tmp)).pixelWidth
            
            DX8.MAP_SetRegion MAPS(tmp), RECT
            
        DX8.DRAW_Map MAPS(tmp), 0, 0, 0, CLng(GrhData(.Frames(tmp)).pixelWidth), CLng(GrhData(.Frames(tmp)).pixelHeight)
        DX8.DRAW_Text DX8_FONT, "FPS: " & FPSx, 10, 350, 0, DX8.ARGB_Set(255, 255, 0, 0), Align_Left
        DX8.Frame ' Renderizamos la escena.
        FPS
    Loop
    
    End With
End Function


Public Function FPS()
        While (GetTickCount - FPSLastCheck) \ 10 < FramesPerSecCounter
            Sleep 5
        Wend
    
    If FPSLastCheck + 1000 < GetTickCount Then
        FPSx = FramesPerSecCounter
        FramesPerSecCounter = 1
        FPSLastCheck = GetTickCount
    Else
        FramesPerSecCounter = FramesPerSecCounter + 1
    End If
    
End Function
