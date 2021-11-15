Attribute VB_Name = "Mod_DX8_AuraEngine"
'***************************************************
'Author: Ezequiel Juárez (Standelf)
'Last Modification: 07/11/12
'Blisse-AO | Sistema de Auras
'Modification: not overflow is Aura not exist
'***************************************************

Option Explicit

Public Type Aura
    Grh             As Integer  '   GrhIndex
    Rotation        As Byte     '   Rotate or Not
    Angle           As Single   '   Angle
    Speed           As Single   '   Speed
    TickCount       As Long     '   TickCount from Speed Controls
    Color(0 To 3)   As Long     '   Color
    OffsetX         As Integer  '   PixelOffset X
    OffsetY         As Integer  '   PixelOffset Y
End Type

Public Auras()      As Aura     '   List of Aura's
Private ForManager As Byte

Public Sub SetCharacterAura(ByVal CharIndex As Integer, ByVal AuraIndex As Byte, ByVal Slot As Byte)
'***************************************************
'Author: Standelf
'Last Modify Date: 27/05/2010
'***************************************************
    If Slot <= 0 Or Slot >= 4 Then Exit Sub
    Set_Aura CharIndex, Slot, AuraIndex
End Sub

Public Sub Load_Auras()
'***************************************************
'Author: Ezequiel Juárez (Standelf)
'Last Modification: 26/05/10
'Load Auras
'***************************************************
    Dim i As Integer, AurasTotales As Integer, Leer As New clsIniManager
    Leer.Initialize Resources.Bin & "auras.ini"

    AurasTotales = Val(Leer.GetValue("Auras", "NumAuras"))
    
    ReDim Preserve Auras(1 To AurasTotales)
    
            For i = 1 To AurasTotales
                Auras(i).Grh = Val(Leer.GetValue(i, "GrhIndex"))
                
                Auras(i).Rotation = Val(Leer.GetValue(i, "Rotate"))
                Auras(i).Angle = 0
                Auras(i).Speed = Leer.GetValue(i, "Speed")
                
                Auras(i).OffsetX = Val(Leer.GetValue(i, "OffsetX"))
                Auras(i).OffsetY = Val(Leer.GetValue(i, "OffsetY"))

            Dim ColorSet As Byte, TempSet As String
            
            For ColorSet = 0 To 3
                TempSet = Leer.GetValue(Val(i), "Color" & ColorSet)
                Auras(i).Color(ColorSet) = D3DColorXRGB(General_Get_ReadField(1, TempSet, Asc(",")), General_Get_ReadField(2, TempSet, Asc(",")), General_Get_ReadField(3, TempSet, Asc(",")))
            Next ColorSet
                
                Auras(i).TickCount = 0
            Next i
                                                         
    Set Leer = Nothing
End Sub

Public Sub DeInit_Auras()
'***************************************************
'Author: Ezequiel Juárez (Standelf)
'Last Modification: 26/05/10
'DeInit Auras
'***************************************************
    '   Erase Data
    Erase Auras()
    
    '   Finish
    Exit Sub
End Sub

Public Sub Set_Aura(ByVal CharIndex As Integer, Slot As Byte, Aura As Byte)
'***************************************************
'Author: Ezequiel Juárez (Standelf)
'Last Modification: 26/05/10
'Set Aura to Char
'***************************************************
    If Slot <= 0 Or Slot >= 4 Then Exit Sub
    If Aura <= 0 Or Aura >= UBound(Auras()) Then Exit Sub
    
    With CharList(CharIndex).Aura(Slot)
        .Grh = Auras(Aura).Grh
            
        .Angle = Auras(Aura).Angle
        .Rotation = Auras(Aura).Rotation
        .Speed = Auras(Aura).Speed
        
        .OffsetX = Auras(Aura).OffsetX
        .OffsetY = Auras(Aura).OffsetY
        
        .Color(0) = Auras(Aura).Color(0)
        .Color(1) = Auras(Aura).Color(1)
        .Color(2) = Auras(Aura).Color(2)
        .Color(3) = Auras(Aura).Color(3)
        
        .TickCount = GetTickCount
    End With
End Sub

Public Sub Delete_All_Auras(ByVal CharIndex As Integer)
'***************************************************
'Author: Ezequiel Juárez (Standelf)
'Last Modification: 26/05/10
'Kill all of aura´s from Char
'***************************************************
    For ForManager = 1 To 4
        Delete_Aura CharIndex, ForManager
    Next ForManager
End Sub
    
Public Sub Delete_Aura(ByVal CharIndex As Integer, Slot As Byte)
'***************************************************
'Author: Ezequiel Juárez (Standelf)
'Last Modification: 26/05/10
'Kill Aura from Char
'***************************************************
    If Slot <= 0 Or Slot >= 4 Then Exit Sub
    
    CharList(CharIndex).Aura(Slot).Grh = 0
    
End Sub

Public Sub Update_Aura(ByVal CharIndex As Integer, Slot As Byte)
'***************************************************
'Author: Ezequiel Juárez (Standelf)
'Last Modification: 26/05/10
'Update Angle of Aura
'***************************************************
    If Slot <= 0 Or Slot >= 4 Then Exit Sub
    
    With CharList(CharIndex).Aura(Slot)
        If GetTickCount - .TickCount > FramesPerSecond Then
            .Angle = .Angle + .Speed
            If .Angle >= 360 Then .Angle = 0
            .TickCount = GetTickCount
        End If
    End With
End Sub

Public Sub Render_Auras(ByVal CharIndex As Integer, X As Integer, Y As Integer)
'***************************************************
'Author: Ezequiel Juárez (Standelf)
'Last Modification: 26/05/10
'Render the Auras from a Char
'***************************************************
On Error GoTo handle
    Dim i As Byte
        For i = 1 To 4
            With CharList(CharIndex).Aura(i)
                If .Grh <> 0 Then
                    If .Rotation = 1 Then Update_Aura CharIndex, i
                    Call TileEngine_Render_GrhIndex(.Grh, X + .OffsetX, Y + .OffsetY, 1, .Color(), Blit_Alpha.Blendop_Aditive, .Angle)
                End If
            End With
        Next i
handle:
    Exit Sub
End Sub
