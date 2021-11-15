Attribute VB_Name = "Mod_Cuentas"
Option Explicit
Public Type BovCuenta
    name As String
    Cantidad As Long
    GrhIndex As Integer
    obInd As UserOBJ
End Type
Public Type PJS
    NamePJ As String
    LvlPJ As Byte
    ClasePJ As eClass
End Type
Public Type Acc
    name As String
    Pass As String
    
    Premium As Boolean
    
    CantPjs As Byte
    PJ(1 To 8) As PJS
    
    CantItemBoV As Byte
    
    ItemBov(1 To 10) As BovCuenta
End Type
Public Cuenta As Acc

Public Sub CrearCuenta(ByVal UserIndex As Integer, ByVal name As String, ByVal Pass As String, ByVal email As String)
Dim ciclo As Byte
'¿Posee caracteres invalidos?
If Not AsciiValidos(name) Or LenB(name) = 0 Then
    Call WriteErrorMsg(UserIndex, "Nombre invalido.")
    Exit Sub
End If

'Si ya existe la cuenta
If FileExist(App.Path & "\Cuentas\" & name & ".bgao", vbNormal) Then
    Call WriteErrorMsg(UserIndex, "El nombre de la cuenta ya existe, por favor ingresa otro.")
    Exit Sub
End If

Call WriteVar(App.Path & "\Cuentas\" & name & ".bgao", "CUENTA", "NOMBRE", name)
Call WriteVar(App.Path & "\Cuentas\" & name & ".bgao", "CUENTA", "PASSWORD", Pass)
Call WriteVar(App.Path & "\Cuentas\" & name & ".bgao", "CUENTA", "MAIL", email)
Call WriteVar(App.Path & "\Cuentas\" & name & ".bgao", "CUENTA", "PREMIUM", "FALSE")
Call WriteVar(App.Path & "\Cuentas\" & name & ".bgao", "CUENTA", "FECHA_CREACION", Now)
Call WriteVar(App.Path & "\Cuentas\" & name & ".bgao", "CUENTA", "FECHA_ULTIMA_VISITA", Now)
Call WriteVar(App.Path & "\Cuentas\" & name & ".bgao", "CUENTA", "ADVERTENCIAS", "0")
Call WriteVar(App.Path & "\Cuentas\" & name & ".bgao", "CUENTA", "SUSPENCIONES", "0")
Call WriteVar(App.Path & "\Cuentas\" & name & ".bgao", "CUENTA", "BAN", "0")

'************************RELLENO LOS PJs************************'
Call WriteVar(App.Path & "\Cuentas\" & name & ".bgao", "PERSONAJES", "CANTIDAD_PJS", "0")
For ciclo = 1 To 8
    Call WriteVar(App.Path & "\Cuentas\" & name & ".bgao", "PERSONAJES", "PJ" & ciclo, "")
Next ciclo
'************************************************************'

'************************RELLENO BOVEDA************************'
Call WriteVar(App.Path & "\Cuentas\" & name & ".bgao", "BOVEDA", "CANTIDAD_ITEMS", "0")
For ciclo = 1 To 10
    Call WriteVar(App.Path & "\Cuentas\" & name & ".bgao", "BOVEDA", "ITM" & ciclo, "0-0")
Next ciclo
'*************************************************************'

Call EnviarCuenta(UserIndex, "", "", "", "", "", "", "", "", False, "0", "1")

DoEvents

'Standelf
Dim LastCuent As String, Cuentas As Integer
LastCuent = name
Cuentas = GetVar(App.Path & "\Server.ini", "ESTADISTICAS", "cantidadcuentas")

Call WriteVar(App.Path & "\Server.ini", "ESTADISTICAS", "ultimacuenta", LastCuent)
Call WriteVar(App.Path & "\Server.ini", "ESTADISTICAS", "cantidadcuentas", (Cuentas + 1))
'Standelf

End Sub

Public Sub ConectarCuenta(ByVal UserIndex As Integer, ByVal name As String, ByVal Pass As String)
'Si NO existe la cuenta
If Not FileExist(App.Path & "\Cuentas\" & name & ".bgao", vbNormal) Then
    Call WriteErrorMsg(UserIndex, "El nombre de la cuenta es inexistente.")
    Exit Sub
End If

With Cuenta
'Si la contraseña es correcta
If Pass = GetVar(App.Path & "\Cuentas\" & name & ".bgao", "CUENTA", "PASSWORD") Then
    If GetVar(App.Path & "\Cuentas\" & name & ".bgao", "CUENTA", "BAN") <> "0" Then
        Call WriteErrorMsg(UserIndex, "Se ha denegado el acceso a tu cuenta por mal comportamiento en el servidor. Por favor comunicate con los administradores del juego para más información.")
    Else
        .CantPjs = GetVar(App.Path & "\Cuentas\" & name & ".bgao", "PERSONAJES", "CANTIDAD_PJS")
        
        .PJ(1).NamePJ = GetVar(App.Path & "\Cuentas\" & name & ".bgao", "PERSONAJES", "PJ1")
        .PJ(2).NamePJ = GetVar(App.Path & "\Cuentas\" & name & ".bgao", "PERSONAJES", "PJ2")
        .PJ(3).NamePJ = GetVar(App.Path & "\Cuentas\" & name & ".bgao", "PERSONAJES", "PJ3")
        .PJ(4).NamePJ = GetVar(App.Path & "\Cuentas\" & name & ".bgao", "PERSONAJES", "PJ4")
        .PJ(5).NamePJ = GetVar(App.Path & "\Cuentas\" & name & ".bgao", "PERSONAJES", "PJ5")
        .PJ(6).NamePJ = GetVar(App.Path & "\Cuentas\" & name & ".bgao", "PERSONAJES", "PJ6")
        .PJ(7).NamePJ = GetVar(App.Path & "\Cuentas\" & name & ".bgao", "PERSONAJES", "PJ7")
        .PJ(8).NamePJ = GetVar(App.Path & "\Cuentas\" & name & ".bgao", "PERSONAJES", "PJ8")
        
        .Premium = GetVar(App.Path & "\Cuentas\" & name & ".bgao", "CUENTA", "PREMIUM")
        
        Call EnviarCuenta(UserIndex, .PJ(1).NamePJ, .PJ(2).NamePJ, .PJ(3).NamePJ, .PJ(4).NamePJ, _
        .PJ(5).NamePJ, .PJ(6).NamePJ, .PJ(7).NamePJ, .PJ(8).NamePJ, .Premium, .CantPjs, "1")
        
        Call WriteVar(App.Path & "\Cuentas\" & name & ".bgao", "CUENTA", "FECHA_ULTIMA_VISITA", Now)
    End If
Else
    Call WriteErrorMsg(UserIndex, "La contraseña es incorrecta. Por favor intentalo nuevamente.")
    Exit Sub
End If
End With
End Sub

Public Function PuedeAgregar(ByVal CuentaName As String)
Dim CantidadPJs As Byte, Premium As Boolean, FreeSlot As Byte, temp As String
CantidadPJs = GetVar(App.Path & "\Cuentas\" & CuentaName & ".bgao", "PERSONAJES", "CANTIDAD_PJS")
Premium = GetVar(App.Path & "\Cuentas\" & CuentaName & ".bgao", "CUENTA", "PREMIUM")

If CantidadPJs = 5 And Not Premium Then
    PuedeAgregar = False
    Exit Function
ElseIf CantidadPJs = 8 Then
    PuedeAgregar = False
    Exit Function
Else
    PuedeAgregar = True
    Exit Function
End If

End Function

Public Function AgregarPersonaje(ByVal CuentaName As String, ByVal UserName As String)
Dim CantidadPJs As Byte, Premium As Boolean, FreeSlot As Byte, temp As String
CantidadPJs = GetVar(App.Path & "\Cuentas\" & CuentaName & ".bgao", "PERSONAJES", "CANTIDAD_PJS")
Premium = GetVar(App.Path & "\Cuentas\" & CuentaName & ".bgao", "CUENTA", "PREMIUM")

If Premium = True Then
    For FreeSlot = 1 To 8
        temp = GetVar(App.Path & "\Cuentas\" & CuentaName & ".bgao", "PERSONAJES", "PJ" & FreeSlot)
            If temp = "" Then
                WriteVar App.Path & "\Cuentas\" & CuentaName & ".bgao", "PERSONAJES", "CANTIDAD_PJS", CantidadPJs + 1
                WriteVar App.Path & "\Cuentas\" & CuentaName & ".bgao", "PERSONAJES", "PJ" & FreeSlot, UserName
                
                WriteVar App.Path & "\Charfile\" & UserName & ".CHR", "INIT", "CUENTA", UCase(CuentaName)
                Exit For
            End If
    Next FreeSlot
Else
    For FreeSlot = 1 To 5
        temp = GetVar(App.Path & "\Cuentas\" & CuentaName & ".bgao", "PERSONAJES", "PJ" & FreeSlot)
            If temp = "" Then
                WriteVar App.Path & "\Cuentas\" & CuentaName & ".bgao", "PERSONAJES", "CANTIDAD_PJS", CantidadPJs + 1
                WriteVar App.Path & "\Cuentas\" & CuentaName & ".bgao", "PERSONAJES", "PJ" & FreeSlot, UserName
                
                WriteVar App.Path & "\Charfile\" & UserName & ".CHR", "INIT", "CUENTA", UCase(CuentaName)
                Exit For
            End If
    Next FreeSlot
End If

End Function

Public Sub BorrarPersonaje(ByVal UserIndex As Integer, ByVal CuentaName As String, ByVal IndiceUser As String)
Dim CantidadPJs As Byte
Dim NamePJ As String

'Consulto el nombre del PJ a eliminar
NamePJ = GetVar(App.Path & "\Cuentas\" & CuentaName & ".bgao", "PERSONAJES", "PJ" & IndiceUser)

CantidadPJs = GetVar(App.Path & "\Cuentas\" & CuentaName & ".bgao", "PERSONAJES", "CANTIDAD_PJS")

WriteVar App.Path & "\Cuentas\" & CuentaName & ".bgao", "PERSONAJES", "CANTIDAD_PJS", CantidadPJs - 1
WriteVar App.Path & "\Cuentas\" & CuentaName & ".bgao", "PERSONAJES", "PJ" & IndiceUser, ""

'Muevo el PJ borrado de la cuenta por si lo quiero recuperar
Call FileCopy(App.Path & "\Charfile\" & NamePJ & ".CHR", App.Path & "\CharfileDel\" & NamePJ & ".CHR")
Call Kill(App.Path & "\Charfile\" & NamePJ & ".CHR")
'Ingreso la fecha y hora de borrado del PJ
WriteVar App.Path & "\CharfileDel\" & NamePJ & ".CHR", "INIT", "BORRADO", Now

Call WriteErrorMsg(UserIndex, "Personaje eliminado con éxito.")
End Sub

Public Sub IniciarBoveda(ByVal UserIndex As Integer, ByVal CuentaName As String)
Call LeerBoveda(UserIndex, CuentaName)

Call EnviarBoveda(UserIndex)
End Sub

Public Sub LeerBoveda(ByVal UserIndex As Integer, ByVal CuentaName As String)
Dim i As Byte
Dim Tmp() As String
Dim ObjInd, ObjCant As Long
Dim uOBJ As UserOBJ

For i = 1 To 10
    Tmp() = Split(GetVar(App.Path & "\Cuentas\" & CuentaName & ".bgao", "BOVEDA", "ITM" & i), "-")
    
    ObjInd = Tmp(0)
    ObjCant = Tmp(1)
    
    With Cuenta
        If ObjInd = 0 Then
            .ItemBov(i).GrhIndex = 0
            .ItemBov(i).name = "(VACIO)"
            .ItemBov(i).Cantidad = 0
            .ItemBov(i).obInd.ObjIndex = 0
        Else
            .ItemBov(i).GrhIndex = ObjData(ObjInd).GrhIndex
            .ItemBov(i).name = ObjData(ObjInd).name
            .ItemBov(i).Cantidad = ObjCant
            .ItemBov(i).obInd.ObjIndex = ObjInd
        End If
    End With
Next i
End Sub

Public Sub DepositarBoveda(ByVal UserIndex As Integer, ByVal NameCuenta As String, ByVal Slot As Byte, ByVal Cantidad As Integer)
If UserList(UserIndex).Invent.Object(Slot).Amount > 0 And Cantidad > 0 Then
    If Cantidad > UserList(UserIndex).Invent.Object(Slot).Amount Then Cantidad = UserList(UserIndex).Invent.Object(Slot).Amount
    If Cantidad < 1 Then Exit Sub
       
       Call LeerBoveda(UserIndex, NameCuenta)
       
    Dim Tmp As Byte
    
    '¿Ya tiene un objeto de este tipo?
    Tmp = 1
    Do Until (Cuenta.ItemBov(Tmp).obInd.ObjIndex) = UserList(UserIndex).Invent.Object(Slot).ObjIndex And _
                Cuenta.ItemBov(Tmp).Cantidad + Cantidad <= MAX_INVENTORY_OBJS
            Tmp = Tmp + 1
            If Tmp > 10 Then Exit Do
    Loop
    
        'Sino se fija por un slot vacio antes del slot devuelto
    If Tmp > 10 Then
        Tmp = 1
        Do Until Cuenta.ItemBov(Tmp).obInd.ObjIndex = 0
            Tmp = Tmp + 1
            
            If Tmp > 10 Then
                Call WriteConsoleMsg(UserIndex, "No tienes mas espacio en el banco!!", FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
        Loop
    End If
        
        'Slot valido
    If Tmp <= 10 Then
        'Mete el obj en el slot
        If Cuenta.ItemBov(Tmp).Cantidad + Cantidad <= MAX_INVENTORY_OBJS Then
            
            'Menor que MAX_INV_OBJS
            Cuenta.ItemBov(Tmp).obInd.ObjIndex = UserList(UserIndex).Invent.Object(Slot).ObjIndex
            Cuenta.ItemBov(Tmp).Cantidad = Cuenta.ItemBov(Tmp).Cantidad + Cantidad
            
            WriteVar App.Path & "\Cuentas\" & NameCuenta & ".bgao", "BOVEDA", "ITM" & Tmp, Cuenta.ItemBov(Tmp).obInd.ObjIndex & "-" & Cuenta.ItemBov(Tmp).Cantidad
            
            Call QuitarUserInvItem(UserIndex, CByte(Slot), Cantidad)
            
            'Actualizamos el inventario del usuario
            Call UpdateUserInv(True, UserIndex, 0)
            
            'actualizamos la boveda de la cuenta
            Call IniciarBoveda(UserIndex, NameCuenta)
        Else
            Call WriteConsoleMsg(UserIndex, "El banco no puede cargar tantos objetos.", FontTypeNames.FONTTYPE_INFO)
        End If
    End If
End If
End Sub

Public Sub RetirarBoveda(ByVal UserIndex As Integer, ByVal NameCuenta As String, ByVal Slot As Byte, ByVal Cantidad As Integer)
If Cantidad < 1 Then Exit Sub
Call LeerBoveda(UserIndex, NameCuenta)

If Cuenta.ItemBov(Slot).Cantidad > 0 Then
    If Cantidad > Cuenta.ItemBov(Slot).Cantidad Then Cantidad = Cuenta.ItemBov(Slot).Cantidad
        If Cuenta.ItemBov(Slot).Cantidad <= 0 Then Exit Sub
        Dim Tmp As Byte
       
       '¿Ya tiene un objeto de este tipo?
        Tmp = 1
        Do Until UserList(UserIndex).Invent.Object(Tmp).ObjIndex = Cuenta.ItemBov(Slot).obInd.ObjIndex And _
           UserList(UserIndex).Invent.Object(Tmp).Amount + Cantidad <= MAX_INVENTORY_OBJS
            Tmp = Tmp + 1
            If Tmp > MAX_INVENTORY_SLOTS Then Exit Do
        Loop
        
        'Sino se fija por un slot vacio
        If Tmp > MAX_INVENTORY_SLOTS Then
                Tmp = 1
                Do Until UserList(UserIndex).Invent.Object(Tmp).ObjIndex = 0
                    Tmp = Tmp + 1
                    If Tmp > MAX_INVENTORY_SLOTS Then
                        Call WriteConsoleMsg(UserIndex, "No podés tener mas objetos.", FontTypeNames.FONTTYPE_INFO)
                        Exit Sub
                    End If
                Loop
                UserList(UserIndex).Invent.NroItems = UserList(UserIndex).Invent.NroItems + 1
        End If
        
        'Mete el obj en el slot
        If UserList(UserIndex).Invent.Object(Tmp).Amount + Cantidad <= MAX_INVENTORY_OBJS Then
            
            'Menor que MAX_INV_OBJS
            Cuenta.ItemBov(Slot).obInd.ObjIndex = Cuenta.ItemBov(Slot).obInd.ObjIndex
            Cuenta.ItemBov(Slot).Cantidad = Cuenta.ItemBov(Slot).Cantidad - Cantidad
                           
            If Cuenta.ItemBov(Slot).Cantidad <= 0 Then
                WriteVar App.Path & "\Cuentas\" & NameCuenta & ".bgao", "BOVEDA", "ITM" & Slot, "0-0"
            Else
                WriteVar App.Path & "\Cuentas\" & NameCuenta & ".bgao", "BOVEDA", "ITM" & Slot, Cuenta.ItemBov(Slot).obInd.ObjIndex & "-" & Cuenta.ItemBov(Slot).Cantidad
            End If
            
            UserList(UserIndex).Invent.Object(Tmp).ObjIndex = Cuenta.ItemBov(Slot).obInd.ObjIndex
            UserList(UserIndex).Invent.Object(Tmp).Amount = UserList(UserIndex).Invent.Object(Tmp).Amount + Cantidad
            
            Call LeerBoveda(UserIndex, NameCuenta)
            
            'Actualizamos el inventario del usuario
            Call UpdateUserInv(True, UserIndex, 0)
            
            'actualizamos la boveda de la cuenta
            Call IniciarBoveda(UserIndex, NameCuenta)
            
        Else
            Call WriteConsoleMsg(UserIndex, "No podés tener mas objetos.", FontTypeNames.FONTTYPE_INFO)
        End If

End If
End Sub
