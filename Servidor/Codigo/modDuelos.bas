Attribute VB_Name = "modDuelos"
Public Type Duelo
    Usuario1 As Integer 'Index 1
    Usuario2 As Integer 'Index 2
    
    Pos1 As WorldPos 'BackUP Pos1
    Pos2 As WorldPos 'BackUP Pos2
    
    map As Integer 'Mapa de Duelo
    
    InitPos1 As WorldPos 'Pos Inicial1
    InitPos2 As WorldPos 'Pos Inicial2
    
    Cuenta As Boolean 'Cuenta
    Contador As Byte '3,2,1 YA!
End Type

Public Duelos As Duelo

Public Sub Init_Duelos()
    With Duelos
        .map = 60
        
        .InitPos1.map = .map
        .InitPos1.X = 35
        .InitPos1.Y = 42
        
        .InitPos2.map = .map
        .InitPos2.X = 62
        .InitPos2.Y = 58
        
        .Contador = 4
        .Cuenta = False
    End With
End Sub

Public Sub CuentaRegresiva()
    With Duelos
        If .Cuenta = True Then
            .Contador = .Contador - 1
            
            If .Contador = 0 Then
                Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("El Duelo entre " & UserList(.Usuario1).name & " y " & UserList(.Usuario2).name & " ha comenzado!", FontTypeNames.FONTTYPE_FIGHT))
                .Cuenta = False
                .Contador = 4
                Call SendData(SendTarget.toMap, Duelos.map, PrepareMessageDuelStart(0))
            End If
        End If
    End With
End Sub

Public Sub Iniciar_Duelo()
    With Duelos
        .Cuenta = True
        .Contador = 4
        
        Call WarpUserChar(.Usuario1, .InitPos1.map, .InitPos1.X, .InitPos1.Y, True)
        Call WarpUserChar(.Usuario2, .InitPos2.map, .InitPos2.X, .InitPos2.Y, True)
        
        DoEvents
        Call SendData(SendTarget.toMap, Duelos.map, PrepareMessageDuelStart(1))
        Call SendData(SendTarget.toMap, Duelos.map, PrepareMessageCountStart(3))
        
    End With
End Sub

Public Sub Terminar_Duelo(ByVal UserIndex As Integer)
    With Duelos
        .Cuenta = False
        .Contador = 4
        
        If .Usuario1 = UserIndex Then
            Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg(UserList(.Usuario1).name & " perdió el duelo contra " & UserList(.Usuario2).name, FontTypeNames.FONTTYPE_FIGHT))
            UserList(.Usuario2).Stats.DuelosGanados = UserList(.Usuario2).Stats.DuelosGanados + 1
            UserList(.Usuario1).Stats.DuelosPerdidos = UserList(.Usuario1).Stats.DuelosPerdidos + 1
        ElseIf .Usuario2 = UserIndex Then
            Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg(UserList(.Usuario2).name & " perdió el duelo contra " & UserList(.Usuario1).name, FontTypeNames.FONTTYPE_FIGHT))
            UserList(.Usuario1).Stats.DuelosGanados = UserList(.Usuario1).Stats.DuelosGanados + 1
            UserList(.Usuario2).Stats.DuelosPerdidos = UserList(.Usuario2).Stats.DuelosPerdidos + 1
        End If
        
        UserList(.Usuario1).flags.EnDuelo = False
        UserList(.Usuario2).flags.EnDuelo = False
        
        Call SendData(SendTarget.toMap, Duelos.map, PrepareMessageDuelStart(0))
        
        DoEvents
        
        Call WarpUserChar(.Usuario1, .Pos1.map, .Pos1.X, .Pos1.Y, True)
        Call WarpUserChar(.Usuario2, .Pos2.map, .Pos2.X, .Pos2.Y, True)
        
        .Pos1.map = 0
        .Pos1.X = 0
        .Pos1.Y = 0
        
        .Pos2.map = 0
        .Pos2.X = 0
        .Pos2.Y = 0
        
        .Usuario1 = 0
        .Usuario2 = 0
    End With
End Sub

Public Sub Rendir_Duelo(ByVal UserIndex As Integer)
    With Duelos
        .Cuenta = False
        .Contador = 4
        
        If MapInfo(Duelos.map).NumUsers = 1 Then
        
            UserList(UserIndex).flags.EnDuelo = False
            Call SendData(SendTarget.toMap, Duelos.map, PrepareMessageDuelStart(0))
            
            Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg(UserList(UserIndex).name & " se retiró de la arena de duelos.", FontTypeNames.FONTTYPE_INFO))


            If .Usuario1 = UserIndex Then
                Call WarpUserChar(.Usuario1, .Pos1.map, .Pos1.X, .Pos1.Y, True)
                .Pos1.map = 0
                .Pos1.X = 0
                .Pos1.Y = 0
                .Usuario1 = 0
            ElseIf .Usuario2 = UserIndex Then
                Call WarpUserChar(.Usuario2, .Pos2.map, .Pos2.X, .Pos2.Y, True)
                .Pos2.map = 0
                .Pos2.X = 0
                .Pos2.Y = 0
                .Usuario2 = 0
            End If

            
        ElseIf MapInfo(Duelos.map).NumUsers = 2 Then
            If .Usuario1 = UserIndex Then
                Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg(UserList(.Usuario1).name & " se rindió en el duelo contra " & UserList(.Usuario2).name, FontTypeNames.FONTTYPE_FIGHT))
                UserList(.Usuario2).Stats.DuelosGanados = UserList(.Usuario2).Stats.DuelosGanados + 1
                UserList(.Usuario1).Stats.DuelosPerdidos = UserList(.Usuario1).Stats.DuelosPerdidos + 1
            ElseIf .Usuario2 = UserIndex Then
                Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg(UserList(.Usuario2).name & " se rindió en el duelo contra " & UserList(.Usuario1).name, FontTypeNames.FONTTYPE_FIGHT))
                UserList(.Usuario1).Stats.DuelosGanados = UserList(.Usuario1).Stats.DuelosGanados + 1
                UserList(.Usuario2).Stats.DuelosPerdidos = UserList(.Usuario2).Stats.DuelosPerdidos + 1
            End If
            
            UserList(.Usuario1).flags.EnDuelo = False
            UserList(.Usuario2).flags.EnDuelo = False
            
            Call SendData(SendTarget.toMap, Duelos.map, PrepareMessageDuelStart(0))
            
            DoEvents
            
            Call WarpUserChar(.Usuario1, .Pos1.map, .Pos1.X, .Pos1.Y, True)
            Call WarpUserChar(.Usuario2, .Pos2.map, .Pos2.X, .Pos2.Y, True)
            
            .Pos1.map = 0
            .Pos1.X = 0
            .Pos1.Y = 0
            
            .Pos2.map = 0
            .Pos2.X = 0
            .Pos2.Y = 0
            
            .Usuario1 = 0
            .Usuario2 = 0
        End If
    End With
End Sub

