Attribute VB_Name = "Mod_DX8_MiniMap"
Public Sub RenderMiniMap2()
    Dim destRect As RECT
    With destRect
        .Top = 0
        .Left = 0
        .bottom = 100
        .Right = 100
    End With
    Dim map_x As Long, map_y As Long
    
    DirectDevice.Clear 0, ByVal 0, D3DCLEAR_TARGET, 0, 0#, 0
    DirectDevice.BeginScene
        For map_y = YMinMapSize To YMaxMapSize
            For map_x = XMinMapSize To XMaxMapSize
                If MapData(map_x, map_y).Graphic(1).GrhIndex <> 0 Then
                    Call DDrawTransGrhtoSurface2(MapData(map_x, map_y).Graphic(1), map_x, map_y, 1, MapData(map_x, map_y).Engine_Light(), 1, map_x, map_y, False, False, 0)
                    Call DDrawTransGrhtoSurface2(MapData(map_x, map_y).Graphic(1), map_x, map_y, 1, MapData(map_x, map_y).Engine_Light(), 1, map_x, map_y, False, False, 0)
                    Call DDrawTransGrhtoSurface2(MapData(map_x, map_y).Graphic(1), map_x, map_y, 1, MapData(map_x, map_y).Engine_Light(), 1, map_x, map_y, False, False, 0)
                    Call DDrawTransGrhtoSurface2(MapData(map_x, map_y).Graphic(1), map_x, map_y, 1, MapData(map_x, map_y).Engine_Light(), 1, map_x, map_y, False, False, 0)
                End If
            Next map_x
        Next map_y
        
    DirectDevice.EndScene
    DirectDevice.Present destRect, ByVal 0, frmMain.newminimap.hWnd, ByVal 0
End Sub
