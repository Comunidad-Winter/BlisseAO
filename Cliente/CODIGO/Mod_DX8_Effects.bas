Attribute VB_Name = "Mod_DX8_Effects"
'***************************************************
'Author: Ezequiel Juárez (Standelf)
'Last Modification: 23/12/10
'Blisse-AO | Sistema de Cuenta en Screen
'***************************************************
Public Enum DX8_Effects_GRH
    WALK_NORTH = 27295
    WALK_SOUTH = 27296
    WALK_WEST = 27298
    WALK_EAST = 27297
End Enum

Private Type x_Effects
    File As Long
    PosX As Integer
    PosY As Integer
    Time As Long
    IniTime As Long
    Alpha As Integer
    Color(0 To 3) As Long
End Type

Public DX8_Effects(1 To 100) As x_Effects

Public Function DX8_Effects_Create(ByVal File As Long, X As Integer, Y As Integer, Time As Long)
    Dim index As Integer
    index = DX8_Effects_Find_Free
    
    If index = 0 Then Exit Function
    
    With DX8_Effects(index)
        .Alpha = 255
        .PosX = X
        .PosY = Y
        .File = File
        .IniTime = FPS * Time
        .Time = .IniTime
        Call Engine_Long_To_RGB_List(.Color(), D3DColorARGB(.Alpha, 255, 255, 255))
        
        MapData(.PosX, .PosY).Effect = index
    End With
End Function

Public Function DX8_Effects_Kill(ByVal index As Integer)
    MapData(DX8_Effects(index).PosX, DX8_Effects(index).PosY).Effect = 0
    DX8_Effects(index).File = 0
End Function

Public Function DX8_Effects_Find_Free() As Integer
    Dim i As Long
        For i = 1 To 100
            If DX8_Effects(i).File = 0 Then
                DX8_Effects_Find_Free = i
                Exit Function
            End If
        Next i
        DX8_Effects_Find_Free = 0
End Function

Public Function DX8_Effects_Update(ByVal index As Long)
Dim tmpColor(0 To 3) As D3DCOLORVALUE
        With DX8_Effects(index)
            If .Time = 0 And .File <> 0 Then
                DX8_Effects_Kill (index)
            ElseIf .File <> 0 And .Time <> 0 Then
                .Time = .Time - 1
                .Alpha = (((.Time / 100) / (.IniTime / 100)) * 255)
                
                Call Engine_Get_ARGB(MapData(.PosX, .PosY).Engine_Light(0), tmpColor(0))
                Call Engine_Get_ARGB(MapData(.PosX, .PosY).Engine_Light(1), tmpColor(1))
                Call Engine_Get_ARGB(MapData(.PosX, .PosY).Engine_Light(2), tmpColor(2))
                Call Engine_Get_ARGB(MapData(.PosX, .PosY).Engine_Light(3), tmpColor(3))
                
                .Color(0) = D3DColorARGB(.Alpha, tmpColor(0).r, tmpColor(0).g, tmpColor(0).b)
                .Color(1) = D3DColorARGB(.Alpha, tmpColor(1).r, tmpColor(1).g, tmpColor(1).b)
                .Color(2) = D3DColorARGB(.Alpha, tmpColor(2).r, tmpColor(2).g, tmpColor(2).b)
                .Color(3) = D3DColorARGB(.Alpha, tmpColor(3).r, tmpColor(3).g, tmpColor(3).b)
            End If
        End With
End Function

Public Function DX8_Effects_Walk_Create(ByVal Heading As E_Heading, ByVal X As Integer, ByVal Y As Integer)
    Select Case Heading
        Case 1 'NORTH
            DX8_Effects_Create DX8_Effects_GRH.WALK_NORTH, X, Y, 4
        Case 2 'EAST
            DX8_Effects_Create DX8_Effects_GRH.WALK_EAST, X, Y, 4
        Case 3 'SOUTH
            DX8_Effects_Create DX8_Effects_GRH.WALK_SOUTH, X, Y, 4
        Case 4 'WEST
            DX8_Effects_Create DX8_Effects_GRH.WALK_WEST, X, Y, 4
    End Select
End Function
