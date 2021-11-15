Attribute VB_Name = "Mod_Soporte"
Type Soporte
    Pregunta As String
    Respuesta As String
    GameMaster As String
    UserName As String
    DateFormulate As String
    DateResponse As String
    Leido As Byte
End Type

Public Soportes(1 To 250) As Soporte

Public Function Free_Soporte() As Byte
    Dim i As Byte
        For i = 1 To 250
            If Soportes(i).Pregunta = "" Then
                Free_Soporte = i
                Exit Function
            End If
        Next i
    
    Free_Soporte = 255
End Function

Public Sub Eliminar_soporte(ByVal UserName As String)
    Dim i As Byte
    
        For i = 1 To 250
            If UCase$(Soportes(i).UserName) = UCase$(UserName) Then
                Soportes(i).DateFormulate = ""
                Soportes(i).DateResponse = ""
                Soportes(i).GameMaster = ""
                Soportes(i).Leido = 0
                Soportes(i).Pregunta = ""
                Soportes(i).Respuesta = ""
                Soportes(i).UserName = ""
                Exit Sub
            End If
        Next i
End Sub

Public Sub Change_Soporte(ByVal Userindex As Integer, ByVal message As String)
    Dim Free_Slot As Byte
    Free_Slot = Free_Soporte
        
        If Free_Soporte = 255 Then
        
            Exit Sub
        Else
            With Soportes(Free_Soporte)
                .Pregunta = message
                .UserName = UserList(Userindex).name
                .DateFormulate = Date & " - " & time
                
                .Respuesta = ""
                .GameMaster = ""
                .DateResponse = ""
                .Leido = 0
            End With
        End If
End Sub

Public Sub Save_Soportes()
    Dim i As Byte
        For i = 1 To 250
            If Soportes(i).Respuesta <> "" Or Soportes(i).Pregunta <> "" Then
                Call WriteVar(App.Path & "\Soporte.ini", "Soporte" & i, "Pregunta", Soportes(i).Pregunta)
                Call WriteVar(App.Path & "\Soporte.ini", "Soporte" & i, "DateResponse", Soportes(i).DateResponse)
                Call WriteVar(App.Path & "\Soporte.ini", "Soporte" & i, "DateFormulate", Soportes(i).DateFormulate)
                Call WriteVar(App.Path & "\Soporte.ini", "Soporte" & i, "GameMaster", Soportes(i).GameMaster)
                Call WriteVar(App.Path & "\Soporte.ini", "Soporte" & i, "Leido", Soportes(i).Leido)
                Call WriteVar(App.Path & "\Soporte.ini", "Soporte" & i, "Respuesta", Soportes(i).Respuesta)
                Call WriteVar(App.Path & "\Soporte.ini", "Soporte" & i, "UserName", Soportes(i).UserName)
            End If
        Next i
End Sub

Public Sub Response_Soporte(ByVal Userindex As String, ByVal UserName As String, ByVal Response As String)
    Dim UserOnline As Byte, i As Integer
        UserOnline = 0
        
        For i = 1 To NumUsers
            If UCase$(UserList(i).name) = UCase$(UserName) Then
                UserOnline = i
                Exit For
            End If
        Next i
        
    If UserOnline = 0 Then ' Está off
        ' Actualizar
        If Save_Offline(UserName, UserList(Userindex).name, Date & " - " & time, Response) = True Then
            Call WriteConsoleMsg(Userindex, "El Usuario " & UserName & " está offline, su respuesta será guardada y será mostrada al usuario cuando este ingrese.", FontTypeNames.FONTTYPE_INFO)
        Else
            Call WriteConsoleMsg(Userindex, "El Usuario " & UserName & " No existe.", FontTypeNames.FONTTYPE_INFO)
        End If
    Else
        ' Actualizar
        With UserList(UserOnline)
            .Soporte.Respuesta = Response
            .Soporte.GameMaster = UserList(Userindex).name
            .Soporte.DateResponse = Date & " - " & time
            .Soporte.Leido = 0
        End With
        
        ' Informar
        Call WriteConsoleMsg(Userindex, "El soporte de " & UserName & " fué respondido exitosamente.", FontTypeNames.FONTTYPE_INFO)
        Save_Soporte UserOnline
        Eliminar_soporte UserName
        DoEvents
        
        Call WriteConsoleMsg(UserOnline, UserList(Userindex).name & " respondió tu consulta.", FontTypeNames.FONTTYPE_INFO)
    End If
        
End Sub

Public Sub Save_Soporte(ByVal Userindex As String)
        With UserList(Userindex)
            If .Soporte.Respuesta <> "" Then
                Call WriteVar(CharPath & UCase$(UserName) & ".chr", "Soporte", "TieneRespuesta", "1")
                Call WriteVar(CharPath & UCase$(UserName) & ".chr", "Soporte", "GameMaster", .Soporte.GameMaster)
                Call WriteVar(CharPath & UCase$(UserName) & ".chr", "Soporte", "DateResponse", .Soporte.DateResponse)
                Call WriteVar(CharPath & UCase$(UserName) & ".chr", "Soporte", "Response", .Soporte.Respuesta)
                
                Call WriteVar(CharPath & UCase$(UserName) & ".chr", "Soporte", "DateResponse", .Soporte.DateFormulate)
                Call WriteVar(CharPath & UCase$(UserName) & ".chr", "Soporte", "Response", .Soporte.Pregunta)
            Else
                Call WriteVar(CharPath & UCase$(UserName) & ".chr", "Soporte", "TieneRespuesta", "0")
                Call WriteVar(CharPath & UCase$(UserName) & ".chr", "Soporte", "GameMaster", "-")
                Call WriteVar(CharPath & UCase$(UserName) & ".chr", "Soporte", "DateResponse", "-")
                Call WriteVar(CharPath & UCase$(UserName) & ".chr", "Soporte", "Response", "-")
                

            End If
        End With
End Sub

Public Function Save_Offline(ByVal UserName As String, ByVal GameMaster As String, ByVal DateResponse As String, ByVal Response As String) As Boolean
        If FileExist(CharPath & UCase$(UserName) & ".chr", vbNormal) Then
            Call WriteVar(CharPath & UCase$(UserName) & ".chr", "Soporte", "TieneRespuesta", "1")
            Call WriteVar(CharPath & UCase$(UserName) & ".chr", "Soporte", "GameMaster", GameMaster)
            Call WriteVar(CharPath & UCase$(UserName) & ".chr", "Soporte", "DateResponse", DateResponse)
            Call WriteVar(CharPath & UCase$(UserName) & ".chr", "Soporte", "Response", Response)
            Save_Offline = True
        Else
            Save_Offline = False
        End If
End Function

Public Sub Read_Soporte(ByVal Userindex As Integer)
        With UserList(Userindex).Soporte
            '.DateFormulate
            '.DateResponse
            '.GameMaster
            '.Leido
            '.Pregunta
            '.Respuesta
            '.UserName
        End With
End Sub

Public Sub Read_Response(ByVal Userindex As Integer)
    
End Sub

