Attribute VB_Name = "Mod_BGSubastas"
'***************************************************
'Author: Ezequiel Ju�rez (Standelf)
'Last Modification: 25/05/10
'Blisse-AO | Black And White AO | Sistema de Subastas 0.13.x
'***************************************************

Public Type c_subasta
    Actual As Boolean ' Sabemos si hay una subasta actualmente
    
    Userindex As Integer ' UserIndex del Usuario Subastando
    OfertaIndex As Integer ' UserIndex del usuario con mayor oferta
    
    OfertaMayor As Long ' Oferta que vale la pena
    ValorBase As Long ' Valor base del item
    
    Objeto As Obj ' Objeto
    
    Tiempo As Byte ' Tiempo de Subasta
End Type
    
Public Subasta As c_subasta

Public Sub Init_Subastas()
    With Subasta
        .Actual = False
        .Userindex = 0
        .OfertaIndex = 0
        .ValorBase = 0
        .OfertaMayor = 0
        .Tiempo = 0
    End With
End Sub

Public Sub Consultar_Subasta(ByVal Userindex As Integer)
    With Subasta
        If .Actual = True Then
                If .Userindex <> -1 Then
                    Call WriteConsoleMsg(Userindex, "[Subasta] El usuario " & UserList(.Userindex).name & " est� subastando " & .Objeto.Amount & " " & ObjData(.Objeto.ObjIndex).name & ". La oferta actual es de " & .OfertaMayor & ". Esta subasta seguir� por " & .Tiempo & " minutos.", FontTypeNames.FONTTYPE_INFO)
                Else
                    Call WriteConsoleMsg(Userindex, "[Subasta] Se est� subastando " & .Objeto.Amount & " " & ObjData(.Objeto.ObjIndex).name & ". La oferta actual es de " & .OfertaMayor & ". Esta subasta seguir� por " & .Tiempo & " minutos.", FontTypeNames.FONTTYPE_INFO)
                End If
            Exit Sub
        Else
            Call WriteIniciarSubastasOrConsulta(Userindex)
        End If
    End With
End Sub

Public Sub Iniciar_Subasta(ByVal Userindex As Integer, Slot As Integer, Amount As Integer, ValorBase As Long)
    With Subasta
        ' Si ya hay una subasta le informamos que debe esperar
        If .Actual = True Then
            Call WriteConsoleMsg(Userindex, "Ya hay una subasta actualmente, deber�s esperar " & .Tiempo & " minutos para inciar una nueva subasta.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        Else
            ' Comprobamos que el usuario tenga lo que intenta ofertar
            If UserList(Userindex).Invent.Object(Slot).Amount < Amount Then
                Call WriteConsoleMsg(Userindex, "No tienes la cantidad de �tems que deseas subastar.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
            
            If ItemNewbie(UserList(Userindex).Invent.Object(Slot).ObjIndex) = True Then
                Call WriteConsoleMsg(Userindex, "No puedes subastar �tems de Newbie.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
            
            ' Actualizamos los datos
            .Actual = True
            .Userindex = Userindex
            
            .OfertaIndex = 0 ' Mientras quede 0 es por que no hay ofertas ;)
            .ValorBase = val(ValorBase)
            .OfertaMayor = val(.ValorBase) ' La oferta mayor es igual al valor inicial ;)
            
            ' Creamos el Objeto
            .Objeto.Amount = Amount
            .Objeto.ObjIndex = UserList(Userindex).Invent.Object(Slot).ObjIndex
            
            .Tiempo = 3
            
            ' Quitamos el Objeto del usuario
            Call QuitarObjetos(.Objeto.ObjIndex, .Objeto.Amount, .Userindex)
            
            ' Ahora podemos informar:
            Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("[Subasta] El usuario " & UserList(.Userindex).name & " est� subastando " & .Objeto.Amount & " " & ObjData(.Objeto.ObjIndex).name & " con un valor inicial de " & .ValorBase, FontTypeNames.FONTTYPE_INFO))

            Exit Sub
        End If
    End With
End Sub

Public Sub Revisar_Subasta(ByVal Userindex As Integer)
    With Subasta
        If Userindex = .OfertaIndex Then
            .OfertaIndex = -1
        End If
        
        If Userindex = .Userindex Then
            .Userindex = -1
        End If
    End With
End Sub

Public Sub Ofertar_Subasta(ByVal Userindex As Integer, Oferta As Long)
    With Subasta
        ' Nos fijamos si existe la subasta
        If .Actual = False Then
            Call WriteConsoleMsg(Userindex, "No hay ninguna subasta actualmente.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        Else
            ' �Tiene la cantidad de oro?
            If UserList(Userindex).Stats.GLD < Oferta Then
                Call WriteConsoleMsg(Userindex, "No tienes la cantidad de oro que intentas ofrecer.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
            
            ' �Ya hay una oferta mayor? �Existia alguna oferta?
            If Oferta <= .OfertaMayor And .OfertaIndex <> 0 Then
                Call WriteConsoleMsg(Userindex, "Tu oferta es menor a la oferta de " & val(.OfertaMayor) & " de " & UserList(.OfertaIndex).name, FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
            
            ' Si no hay oferta, revisamos la nueva oferta para que no sea menor al precio base
            If .OfertaIndex = 0 And Oferta <= .ValorBase Then
                Call WriteConsoleMsg(Userindex, "Tu oferta es menor a la oferta del valor inicial de " & val(.ValorBase), FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
            
            ' Si no hay oferta previa esto no sirve ;)
            If .OfertaIndex <> 0 Then
                ' Antes de actualizar devolvemos las cosas al flaco anterior
                UserList(.OfertaIndex).Stats.GLD = UserList(.OfertaIndex).Stats.GLD + val(.OfertaMayor)
                Call WriteUpdateGold(.OfertaIndex)
            End If
            
            ' Ahora podemos actualizar tranquilos
            .OfertaIndex = Userindex
            .OfertaMayor = Oferta
            
            ' Restamos el oro:
            UserList(.OfertaIndex).Stats.GLD = UserList(.OfertaIndex).Stats.GLD - val(.OfertaMayor)
            Call WriteUpdateGold(.OfertaIndex)
            
            ' Informamos a los usuarios sobre la nueva oferta;
            Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("[Subasta] El usuario " & UserList(.OfertaIndex).name & " aument� la oferta a " & .OfertaMayor, FontTypeNames.FONTTYPE_INFO))
        End If
    End With
End Sub

Public Sub Actualizar_Subasta()
    With Subasta
        ' Revisamos
        If .Actual = False Then
            Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("[Subasta] No hay subastas actualmente, para iniciar una nueva subasta utilice el comando /Subasta", FontTypeNames.FONTTYPE_INFO))
            Exit Sub
        Else
            ' Restamos tiempo
            .Tiempo = .Tiempo - 1
            
            ' Terminamos si es necesario, sino solo recordamos
            If .Tiempo <= 0 Then
                Call Termina_Subasta
            Else
                If .Userindex <> -1 Then
                    Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("[Subasta] El usuario " & UserList(.Userindex).name & " est� subastando " & .Objeto.Amount & " " & ObjData(.Objeto.ObjIndex).name & ". La oferta actual es de " & .OfertaMayor & ". Esta subasta seguir� por " & .Tiempo & " minutos.", FontTypeNames.FONTTYPE_INFO))
                Else
                    Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("[Subasta] Se est� subastando " & .Objeto.Amount & " " & ObjData(.Objeto.ObjIndex).name & ". La oferta actual es de " & .OfertaMayor & ". Esta subasta seguir� por " & .Tiempo & " minutos.", FontTypeNames.FONTTYPE_INFO))
                End If
            End If
        End If
    End With
End Sub

Public Sub Termina_Subasta()
    With Subasta
        If .OfertaIndex = 0 Then
            ' Informamos que la subasta termino, y que nadie ofert�
            Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("[Subasta] La subasta de " & .Objeto.Amount & " " & ObjData(.Objeto.ObjIndex).name & " termin� sin ninguna oferta.", FontTypeNames.FONTTYPE_INFO))
       
            If .Userindex <> -1 Then
                Call MeterItemEnInventario(.Userindex, .Objeto)
            End If
            
            'Reseteamos los datos
            .Actual = False
            .Userindex = 0
            .OfertaIndex = 0
            .ValorBase = 0
            .OfertaMayor = 0
            .Tiempo = 0
        Else
            ' Informamos que la subasta termino, y quien se lleva las cosas.
            
            ' Entregamos el Item, y el Oro
            If .OfertaIndex <> -1 Then
                Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("[Subasta] El usuario " & UserList(.OfertaIndex).name & " gan� la subasta de " & .Objeto.Amount & " " & ObjData(.Objeto.ObjIndex).name & " por la cantidad de " & .OfertaMayor, FontTypeNames.FONTTYPE_INFO))
    
                If MeterItemEnInventario(.OfertaIndex, .Objeto) Then
                    Call WriteConsoleMsg(.OfertaIndex, "Felicitaciones, has ganado la subasta de " & .Objeto.Amount & " " & ObjData(.Objeto.ObjIndex).name & " por la cantidad de " & .OfertaMayor, FontTypeNames.FONTTYPE_INFO)
                End If
            End If
            
            ' Enviamos el Oro
            If .Userindex <> -1 Then
                UserList(.Userindex).Stats.GLD = UserList(.Userindex).Stats.GLD + val(.OfertaMayor)
                Call WriteUpdateGold(.Userindex)
            End If
            
            'Reseteamos los datos
            .Actual = False
            .Userindex = 0
            .OfertaIndex = 0
            .ValorBase = 0
            .OfertaMayor = 0
            .Tiempo = 0
        End If
    End With
End Sub
