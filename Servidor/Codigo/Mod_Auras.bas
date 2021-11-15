Attribute VB_Name = "Mod_Auras"
Public Sub SetAura(ByVal Userindex As Integer, ByVal Aura As Byte, ByVal Slot As Byte, Optional ByVal Refresh As Boolean = False)
    If Slot <= 0 Or Slot >= 5 Then Exit Sub
    UserList(Userindex).Char.Aura(Slot) = Aura
    If Refresh Then
        SendSpecificAura Userindex, Slot
    End If
End Sub

Public Sub ResetAuras(ByVal Userindex As Integer)
    Dim i As Byte
        For i = 1 To 4
            UserList(Userindex).Char.Aura(i) = Aura
        Next i
End Sub

Public Sub SendSpecificAura(ByVal Userindex As Integer, ByVal Slot As Byte)
    If UserList(Userindex).Char.Aura(Slot) <> 0 Then
        Call modSendData.SendToMap(UserList(Userindex).Pos.Map, PrepareAuraSet(Userindex, Slot))
    End If
End Sub

Public Sub SendAuras(ByVal Userindex As Integer)
    Dim i As Byte
        For i = 1 To 4
            If UserList(Userindex).Char.Aura(i) <> 0 Then
                Call modSendData.SendToMap(UserList(Userindex).Pos.Map, PrepareAuraSet(Userindex, i))
            End If
        Next i
End Sub

Public Sub KickAuras(ByVal Userindex As Integer)
    Dim i As Byte
        For i = 1 To 4
            UserList(Userindex).Char.Aura(i) = 0
                Call modSendData.SendToMap(UserList(Userindex).Pos.Map, PrepareAuraSet(Userindex, i))
        Next i
End Sub

Public Function FindSlotFreeAura(ByVal Userindex As Integer) As Byte
    Dim i As Byte
        For i = 1 To 4
            If UserList(Userindex).Char.Aura(i) = 0 Or UserList(Userindex).Char.Aura(i) = 1 Then
                FindSlotFreeAura = i
                Exit Function
            End If
        Next i
End Function

Public Function TieneEstaAura(ByVal Userindex As Integer, Aura As Byte) As Byte
    Dim i As Byte
        For i = 1 To 4
            If UserList(Userindex).Char.Aura(i) = Aura Then
                TieneEstaAura = i
                Exit Function
            End If
        Next i
    TieneEstaAura = 0
End Function
