Attribute VB_Name = "mod_PrestigioCanjes"
Type CanjeItem
    ID As Integer
    Valor As Integer
End Type

Public ItemsDisponibles() As CanjeItem
Public CantidadCanjes As Integer

'Cargamos la lista de objetos disponibles

Public Sub LoadCanjeList()
Dim Leer As New clsIniReader
    Call Leer.Initialize(DatPath & "centrocanjes.dat")
    
    CantidadCanjes = val(Leer.GetValue("Canjes", "Items"))

    ReDim ItemsDisponibles(1 To CantidadCanjes)
    Dim i As Integer

        For i = 1 To CantidadCanjes
            ItemsDisponibles(i).ID = val(Leer.GetValue("Item" & val(i), "ID"))
            ItemsDisponibles(i).Valor = val(Leer.GetValue("Item" & val(i), "Valor"))
        Next i
    
    Set Leer = Nothing
End Sub

Public Sub DarPrestigioCanje(ByVal Userindex As Integer, ReceptorIndex As Integer, Puntos As Integer)
    'Esta función da prestigio de canjes no de reputacion
    'Donaciones
    UserList(ReceptorIndex).PrestigioC = UserList(ReceptorIndex).PrestigioC + Puntos
    
    Call WriteConsoleMsg(Userindex, "El GameMaster " & UserList(Userindex).name & "te otorgó " & Puntos & " puntos de Canjes. Con esto sumas " & UserList(ReceptorIndex).PrestigioR & " de Reputación y " & UserList(ReceptorIndex).PrestigioC & " de Canje.", FontTypeNames.FONTTYPE_INFO)
    Call WriteConsoleMsg(Userindex, "Otorgaste " & Puntos & " de Canje a " & UserList(ReceptorIndex).name, FontTypeNames.FONTTYPE_INFO)
End Sub

Public Sub DarPrestigioRepu(ByVal Userindex As Integer, ReceptorIndex As Integer, Puntos As Integer)
    'Esta función da prestigio de Reputacion y de canjes.
    'Torneos - Deaths - Etc
    UserList(ReceptorIndex).PrestigioC = UserList(ReceptorIndex).PrestigioC + Puntos
    UserList(ReceptorIndex).PrestigioR = UserList(ReceptorIndex).PrestigioR + Puntos
    
    Call WriteConsoleMsg(ReceptorIndex, "El GameMaster " & UserList(Userindex).name & "te otogó " & Puntos & " puntos de Prestigio. Con esto sumas " & UserList(ReceptorIndex).PrestigioR & " de Reputación y " & UserList(ReceptorIndex).PrestigioC & " de Canje.", FontTypeNames.FONTTYPE_INFO)
    Call WriteConsoleMsg(Userindex, "Otorgaste " & Puntos & " de Reputación/Canje a " & UserList(ReceptorIndex).name, FontTypeNames.FONTTYPE_INFO)
   
End Sub

Public Sub QuitarPrestigioRepu(ByVal Userindex As Integer, Puntos As Integer, Motivo As String)
    'Antifacciones - Faltas de respeto a los GMs- Encarcelamiento.
    UserList(Userindex).PrestigioR = UserList(Userindex).PrestigioR - Puntos
    
    If UserList(Userindex).PrestigioR > 0 Then UserList(Userindex).PrestigioR = 0
    Call WriteConsoleMsg(Userindex, "Acabas de perder " & Puntos & " de Prestigio por " & Motivo & ". Tienes en total " & UserList(Userindex).PrestigioR & " puntos de Prestigio.", FontTypeNames.FONTTYPE_INFO)
  
End Sub

Public Sub QuitarPrestigioCanje(ByVal Userindex As Integer, Puntos As Integer)
    UserList(Userindex).PrestigioC = UserList(Userindex).PrestigioC - Puntos
    
    If UserList(Userindex).PrestigioC > 0 Then UserList(Userindex).PrestigioC = 0
End Sub

Public Sub DamePrestigio(ByVal Userindex As Integer)
    Call WriteConsoleMsg(Userindex, "Tienes " & UserList(Userindex).PrestigioR & " puntos de Prestigio.", FontTypeNames.FONTTYPE_INFO)
    Call WriteConsoleMsg(Userindex, "Tienes " & UserList(Userindex).PrestigioC & " puntos de Canje.", FontTypeNames.FONTTYPE_INFO)
End Sub

Public Sub CambiarItem(ByVal Userindex As Integer, ItemIndex As Integer)
    ' Nos fijamos si tienen los puntos
    If UserList(Userindex).PrestigioC < ItemsDisponibles(ItemIndex).Valor Then
        Call WriteConsoleMsg(Userindex, "No tienes el prestigio de canje suficiente para adquirir este item.", FontTypeNames.FONTTYPE_INFO)
        Exit Sub
    End If
    
    ' Nos fijamos que tenga slots en el inventario
    Dim ObjetoCambio As Obj
        ObjetoCambio.Amount = 1
        ObjetoCambio.ObjIndex = ItemsDisponibles(ItemIndex).ID
    
    If MeterItemEnInventario(Userindex, ObjetoCambio) = False Then
        Exit Sub
    Else
        ' Quitamos La cantidad de prestigio
        Call QuitarPrestigioCanje(Userindex, ItemsDisponibles(ItemIndex).Valor)
        Call WriteConsoleMsg(Userindex, "Felicidades!, Has intercambiado " & ItemsDisponibles(ItemIndex).Valor & " puntos de Canje por 1 -" & ObjData(ItemsDisponibles(ItemIndex).ID).name & ".", FontTypeNames.FONTTYPE_INFO)
    End If
    
    Exit Sub
End Sub
