Attribute VB_Name = "Mod_ControlAO"
Public MultExp As Byte 'Multiplicador Exp
Public MultOro As Byte 'Multiplicador Oro
Public MultTrabajo As Byte 'Mutiplicador de Trabajador, cuando queremos que el trabajador extraiga más materiales usamos esto. _
                                        Tener en cuenta que multiplica la cantidad obtenida en el random, no aumentar mucho.
Public EstGlobal As Boolean 'Chat Global
Public EstFaccionario As Boolean 'Chat Faccionario

'TonchitoZ: Happy Hour for Account (o como se escriba en inglés) _
                Cambio a cargar desde un INI así no tenemos que subir todo un Exe para cambiar datos _
                y solamente modificamos el INI en 1 segundo ;)
Public IntHappyHour, DuracionHHour As Integer
Public HHOro, HHExp As Byte
Public HappyHour As Boolean

Public Invasion(1 To 20) As Integer

Public Const SND_EQUIPARTIRAR_CASCO As Integer = 215
Public Const SND_EQUIPARTIRAR_ARMADURA As Integer = 216
Public Const SND_EQUIPARTIRAR_ARMA As Integer = 217
Public Const SND_EQUIPARTIRAR_ANILLO As Integer = 218
Public Const SND_EQUIPARTIRAR_SPELL As Integer = 219
Public Const SND_EQUIPARTIRAR_SPELL2 As Integer = 223
Public Const SND_EQUIPARTIRAR_ORO As Integer = 222
Public Const SND_EQUIPARTIRAR_POTAS As Integer = 220
Public Const SND_EQUIPARTIRAR_ESCUDO As Integer = 221

Public Function Check_Hay_Invasion() As Boolean
    Dim i  As Byte
        
        For i = 1 To 20
            If Invasion(i) <> 0 Then
                Check_Hay_Invasion = True
                Exit Function
            End If
        Next i
        
        Check_Hay_Invasion = False
End Function

Public Sub Detener_Invasion(ByVal UserIndex As Integer)
    Dim i  As Byte
        
        For i = 1 To 20
            If Invasion(i) <> 0 Then
                Call QuitarNPC(Invasion(i))
            End If
        Next i
        
    Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("[Invasión] El GameMaster " & UserList(UserIndex).name & " detuvo la invasión.", FontTypeNames.FONTTYPE_INFO))

End Sub
Public Sub Check_Termina_Invasion()
    Dim i  As Byte
        
        For i = 1 To 20
            If Invasion(i) <> 0 Then
                Exit Sub
            End If
        Next i
        
    Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("[Invasión] La invasión ha terminado.", FontTypeNames.FONTTYPE_INFO))

End Sub

Public Sub Check_NPC_Invasion(ByVal NpcIndex As Integer)
    Dim i  As Byte
        For i = 1 To 20
            If Invasion(i) = NpcIndex Then
                Invasion(i) = 0
                Call Check_Termina_Invasion
                Exit Sub
            End If
        Next i
End Sub

Public Sub Check_Estado()
Dim Cambio As Boolean
    Cambio = False
    
    Select Case Hour(Now)
        Case 5, 6, 7
            If EstadoActual <> 1 Then
                EstadoActual = 1
                Cambio = True
            End If
        Case 8, 9, 10, 11, 12
                    If Lloviendo And EstadoActual <> 4 Then
                        EstadoActual = 4
                        Cambio = True
                    ElseIf EstadoActual <> 2 And Not Lloviendo Then
                        EstadoActual = 2
                        Cambio = True
                    End If
            
        Case 13, 14, 15, 16, 17
                    If Lloviendo And EstadoActual <> 4 Then
                        EstadoActual = 4
                        Cambio = True
                    ElseIf EstadoActual <> 3 And Not Lloviendo Then
                        EstadoActual = 3
                        Cambio = True
                    End If
            
        Case 18, 19, 20, 21
            If EstadoActual <> 4 Then
                EstadoActual = 4
                Cambio = True
            End If
        Case 22, 23, 24, 0, 1, 2, 3, 4
            If EstadoActual <> 5 Then
                EstadoActual = 5
                Cambio = True
            End If
    End Select
    
    If Cambio = True Then
        Call SendData(SendTarget.ToAll, 0, PrepareUpdateClima(EstadoActual))
    End If
End Sub

Public Sub Send_Estado(ByVal UserIndex As Integer)
    Call WriteSendClima(UserIndex, CByte(EstadoActual))
End Sub

Public Sub InitServerSettings()
    MultExp = val(GetVar(IniPath & "Server.ini", "INIT", "EXPSV"))
    MultOro = val(GetVar(IniPath & "Server.ini", "INIT", "GOLDSV"))
    IntHappyHour = val(GetVar(IniPath & "Server.ini", "INIT", "INTERVALOHAPPYHOUR"))
    DuracionHHour = val(GetVar(IniPath & "Server.ini", "INIT", "DURACIONHAPPYHOUR"))
    
    MultTrabajo = 8
    EstGlobal = True
    EstFaccionario = True
End Sub

Public Function CaracterInvalido(ByVal Text As String, Character As Byte) As Boolean
Dim i As Byte
    For i = 1 To Len(Text)
        If Asc(mid(Text, i, 1)) = Character Then
            CaracterInvalido = True
            Exit Function
        End If
    Next i
    CaracterInvalido = False
End Function

Public Function ExpAndGoldHH(ByVal HayHH As Boolean) As Byte
If HayHH = True Then
    HHExp = val(GetVar(IniPath & "Server.ini", "INIT", "EXPHappyHour"))
    HHOro = val(GetVar(IniPath & "Server.ini", "INIT", "GOLDHappyHour"))
Else
    HHExp = val(GetVar(IniPath & "Server.ini", "INIT", "EXPSV"))
    HHOro = val(GetVar(IniPath & "Server.ini", "INIT", "GOLDSV"))
End If
End Function


Public Sub Update_Statics()
Dim LastPJ As String, LastCuent As String, Cuentas As Integer, PJS As Integer, Premiums As Integer
LastPJ = GetVar(App.Path & "\Server.ini", "Estadisticas", "ultimopj")
LastCuent = GetVar(App.Path & "\Server.ini", "Estadisticas", "ultimacuenta")
Cuentas = GetVar(App.Path & "\Server.ini", "Estadisticas", "cantidadcuentas")
PJS = GetVar(App.Path & "\Server.ini", "Estadisticas", "cantidadpjs")
Premiums = GetVar(App.Path & "\Server.ini", "Estadisticas", "cuentaspremium")
    
    Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("[Estadisticas] Estadisticas Web Actualizadas.", FontTypeNames.FONTTYPE_INFO))
End Sub


Public Sub UserItemsSound(ByVal UserIndex As Integer, OBJType As eOBJType)
    Select Case OBJType
        Case eOBJType.otAnillo, eOBJType.otLlaves
            Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(SND_EQUIPARTIRAR_ANILLO, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y))
        Case eOBJType.otArmadura
            Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(SND_EQUIPARTIRAR_ARMADURA, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y))
        Case eOBJType.otCASCO
            Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(SND_EQUIPARTIRAR_CASCO, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y))
        Case eOBJType.otESCUDO
            Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(SND_EQUIPARTIRAR_ESCUDO, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y))
        Case eOBJType.otGuita
            Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(SND_EQUIPARTIRAR_ORO, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y))
        Case eOBJType.otPociones
            Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(SND_EQUIPARTIRAR_POTAS, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y))
        Case eOBJType.otWeapon
            Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(SND_EQUIPARTIRAR_ARMA, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y))
        Case eOBJType.otPergaminos
            Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(SND_EQUIPARTIRAR_SPELL, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y))
    End Select
End Sub

Public Sub NpcItemsSound(ByVal NpcIndex As Integer, OBJType As eOBJType)
    Select Case OBJType
        Case eOBJType.otAnillo, eOBJType.otLlaves
            Call SendData(SendTarget.ToNPCArea, NpcIndex, PrepareMessagePlayWave(SND_EQUIPARTIRAR_ANILLO, Npclist(NpcIndex).Pos.X, Npclist(NpcIndex).Pos.Y))
        Case eOBJType.otArmadura
            Call SendData(SendTarget.ToNPCArea, NpcIndex, PrepareMessagePlayWave(SND_EQUIPARTIRAR_ARMADURA, Npclist(NpcIndex).Pos.X, Npclist(NpcIndex).Pos.Y))
        Case eOBJType.otCASCO
            Call SendData(SendTarget.ToNPCArea, NpcIndex, PrepareMessagePlayWave(SND_EQUIPARTIRAR_CASCO, Npclist(NpcIndex).Pos.X, Npclist(NpcIndex).Pos.Y))
        Case eOBJType.otESCUDO
            Call SendData(SendTarget.ToNPCArea, NpcIndex, PrepareMessagePlayWave(SND_EQUIPARTIRAR_ESCUDO, Npclist(NpcIndex).Pos.X, Npclist(NpcIndex).Pos.Y))
        Case eOBJType.otGuita
            Call SendData(SendTarget.ToNPCArea, NpcIndex, PrepareMessagePlayWave(SND_EQUIPARTIRAR_ORO, Npclist(NpcIndex).Pos.X, Npclist(NpcIndex).Pos.Y))
        Case eOBJType.otPociones
            Call SendData(SendTarget.ToNPCArea, NpcIndex, PrepareMessagePlayWave(SND_EQUIPARTIRAR_POTAS, Npclist(NpcIndex).Pos.X, Npclist(NpcIndex).Pos.Y))
        Case eOBJType.otWeapon
            Call SendData(SendTarget.ToNPCArea, NpcIndex, PrepareMessagePlayWave(SND_EQUIPARTIRAR_ARMA, Npclist(NpcIndex).Pos.X, Npclist(NpcIndex).Pos.Y))
        Case eOBJType.otPergaminos
            Call SendData(SendTarget.ToNPCArea, NpcIndex, PrepareMessagePlayWave(SND_EQUIPARTIRAR_SPELL, Npclist(NpcIndex).Pos.X, Npclist(NpcIndex).Pos.Y))
    End Select
End Sub

Public Function NoEsAmbiente(ByVal ObjIndex As Integer) As Boolean

    NoEsAmbiente = Not (ObjData(ObjIndex).OBJType = eOBJType.otArbolElfico Or _
                        ObjData(ObjIndex).OBJType = eOBJType.otArboles Or _
                        ObjData(ObjIndex).OBJType = eOBJType.otCarteles Or _
                        ObjData(ObjIndex).OBJType = eOBJType.otFogata Or _
                        ObjData(ObjIndex).OBJType = eOBJType.otForos Or _
                        ObjData(ObjIndex).OBJType = eOBJType.otManchas Or _
                        ObjData(ObjIndex).OBJType = eOBJType.otPuertas Or _
                        ObjData(ObjIndex).OBJType = eOBJType.otCualquiera Or _
                        ObjData(ObjIndex).OBJType = eOBJType.otAmbiente Or _
                        ObjData(ObjIndex).OBJType = eOBJType.otTeleport)

End Function
