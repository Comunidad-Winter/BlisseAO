Attribute VB_Name = "Mod_Duelos"
Option Explicit

Dim Minutos, Segundos As Byte

Public Type Counter
    Iniciar As Boolean
    Valor As Byte
End Type
Public Contador As Counter

Public Type EventUsers
    IndUser As Byte 'Indice del usuario en AO
   
    Pos As WorldPos 'Posición en la que se encuentra el PJ antes de aceptar el duelo.
End Type

Public Type tDeath
    estado As Boolean 'Abierto o Cerrado
    MinUser As Byte 'Cantidad de Usuarios que aceptaron el evento
    MaxUser As Byte 'Cantidad máxima de usuarios que puede haber en el evento
    
    AtacableMap As WorldPos 'Area de ataque del evento
    EsperaMap As WorldPos
    
    Pago As Byte 'Dinero necesario para el evento
    User(1 To 255) As EventUsers
End Type
Public DeathM As tDeath

Public Sub CuentaRegresiva()
If Contador.Iniciar = True Then
        Contador.Valor = Contador.Valor - 1
                
        Call SendData(SendTarget.toMap, DeathM.AtacableMap.Map, PrepareMsgCountScreen(Contador.Valor))
        
        If Contador.Valor = -2 Then
            Contador.Iniciar = False
        End If
End If
End Sub

Public Sub AlmacenarEstadoPreDeath(ByVal EventUserIndex As Integer, ByVal Userindex As Integer)
With DeathM
    'Coordenadas donde me encontraba anteriormente
    .User(EventUserIndex).Pos.Map = UserList(Userindex).Pos.Map
    .User(EventUserIndex).Pos.X = UserList(Userindex).Pos.X
    .User(EventUserIndex).Pos.Y = UserList(Userindex).Pos.Y

    'Guardamos el indice
    .User(EventUserIndex).IndUser = Userindex
    UserList(Userindex).IndDeath = EventUserIndex
'    UserList (Userindex)
End With
End Sub

Public Sub TimeToDeathM()
If Segundos = 60 Then
    Minutos = Minutos + 1
    Segundos = 0
    If Minutos < 2 Then
        'Acá aviso: "Quedan X minutos para el comienzo del DeathM."
        Call WriteConsoleMsg(SendTarget.ToAll, "El deathMatch comenzará en " & (3 - Minutos) & " minutos.", FontTypeNames.FONTTYPE_DEATHMATCH)
    
    ElseIf Minutos = 2 Then
        If DeathM.MinUser <> 0 Then
            'Teletrasporto a todos a la sala de espera
            Dim X As Byte
            For X = 1 To DeathM.MinUser
                Call WarpUserChar(DeathM.User(X).IndUser, DeathM.EsperaMap.Map, DeathM.EsperaMap.X, DeathM.EsperaMap.Y, True)
            Next X
        End If
            
        Call WriteConsoleMsg(SendTarget.ToAll, "El deathMatch comenzará en 1 minuto.", FontTypeNames.FONTTYPE_DEATHMATCH)
        
    ElseIf Minutos = 3 Then
        If DeathM.MinUser <> 0 Then
            'Teletrasporto a todos a la arena
            Dim X1 As Byte
            For X1 = 1 To DeathM.MinUser
                Call WarpUserChar(DeathM.User(X1).IndUser, DeathM.AtacableMap.Map, DeathM.AtacableMap.X, DeathM.AtacableMap.Y, True)
            Next X1
        End If
        'Comienzo la cuenta regresiva
        Contador.Iniciar = True
        Contador.Valor = 5
        
        'Desactivo el evento (cierro suscripcion inclusive)
        DeathM.estado = False
        Minutos = 0
    End If
Else
    Segundos = Segundos + 1
    Debug.Print Segundos
End If
End Sub
