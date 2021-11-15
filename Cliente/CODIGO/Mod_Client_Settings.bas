Attribute VB_Name = "Mod_Client_Settings"
'***************************************************
'Author: Ezequiel Juárez (Standelf)
'Last Modification: 05/02/2011
'Blisse-AO | Guardado Binario y Lectura de _
    Opciones del juego (General y Engine) _
    Variables Generales
'***************************************************
Option Explicit

'   Constantes de Links
Public Const Client_Web As String = "http://www.blisse-ao.com.ar/updater/"
Public Const Client_Forum As String = "http://blisse-games.com.ar/blisse-ao-f121.html"
Public Const Game4Fun As String = "http://game4fun.net"

'   Datos del servidor
Public Const Server_IP As String = "localhost"
Public Const Server_Port As String = "10200"

'   Variables Temporales
Public MostrarExp As Boolean
Public EnDuelo As Boolean
Public Items_Seg As Boolean

'   General
Type cl_Settings
    'Graphical Engine
    
    Aceleracion As Byte
    vSync As Boolean
    LimiteFPS As Boolean
    Ventana As Boolean
    Luces As Byte
    BufferSize As Byte
    PartyMembers As Boolean
    DinamicInventory As Boolean
    Nombres As Byte
    NombreItems As Boolean
    UsarSombras As Boolean
    ProyectileEngine As Boolean
    
    ParticleMeditation As Boolean
    MostrarFPS As Boolean
    
    
    Dialog_Align As Byte
    
    
    ParticleEngine As Boolean
    TonalidadPJ As Boolean
    useVideoMemory As Boolean
    MemoryVideoMax As Long
    
    Musica As Boolean
    Sonido As Boolean
    SoundVolume As Byte
    MusicVolume As Byte
    Sonido3D As Boolean
    
    DialogosEnConsola As Boolean
    Mouse_Effect As Boolean
    Water_Effect As Boolean
    
    Shadow_Effect As Boolean
    Walk_Effect As Boolean
    Text_Effect As Boolean
    Reflect_Effect As Boolean
    
    DragerDrop As Boolean
    UltimaCuenta As String
    MiniMap As Boolean
    SeguroItems As Boolean
    FirstTime As Boolean
    GuildNews As Boolean
    DialogoClanesActivo As Boolean
    DialogoClanesCant As Byte
    Recordar As Boolean

End Type

Public Settings As cl_Settings

Public Sub Settings_Init()
'***************************************************
'Author: Standelf
'Last Modification: 14/09/10
'Load The Settings, if file not exist then create a new default config
'***************************************************
    If Not General_File_Exist(Resources.Bin & "Settings.CFG", vbNormal) Then
        Call Settings_Set_Default
        Call Settings_Save
    Else
        Dim N As Integer
        N = FreeFile
            Open Resources.Bin & "Settings.CFG" For Binary As #N
                Get #N, , Settings
            Close #N
    
            Items_Seg = Settings.SeguroItems
    End If
End Sub

Public Sub Settings_Save()
'***************************************************
'Author: Standelf
'Last Modification: 14/09/10
'Save the settings into a file.
'***************************************************
Dim File
File = FreeFile
    Open Resources.Bin & "Settings.CFG" For Binary Access Write As File
        Put File, , Settings
    Close #File
End Sub

Public Sub Settings_Set_Default()
'***************************************************
'Author: Standelf
'Last Modification: 17/05/10
'Load Default settings
'***************************************************
        With Settings
            .Aceleracion = 0
            .vSync = False
            .LimiteFPS = True
            .Ventana = False
            .Luces = 1
            .BufferSize = 9
            .DinamicInventory = True
            .PartyMembers = True
            .DialogosEnConsola = True
            .Nombres = 1
            .NombreItems = True
            .DragerDrop = True
            .UltimaCuenta = ""
            .UsarSombras = True
            .ParticleMeditation = True
            .Musica = True
            .Sonido = True
            .Sonido3D = True
            .SoundVolume = 99
            .MusicVolume = 99
            .Dialog_Align = 1
            .MiniMap = True
            .ParticleEngine = True
            .ProyectileEngine = True
            .TonalidadPJ = True
            .SeguroItems = True
            .FirstTime = True
            .GuildNews = True
            .DialogoClanesActivo = True
            .DialogoClanesCant = 5
            .useVideoMemory = True
            .MemoryVideoMax = 512
            .MostrarFPS = True
            
            .Water_Effect = True
            .Mouse_Effect = True
            .Shadow_Effect = True
            .Walk_Effect = True
            .Text_Effect = True
            .Reflect_Effect = True
            .Recordar = False
            
        End With
End Sub


Public Sub Settings_Set_Max()
'***************************************************
'Author: Standelf
'Last Modification: 24/01/13
'Load Default settings
'***************************************************
        With Settings
            .Aceleracion = 1
            .vSync = False
            .LimiteFPS = True
            .Ventana = False
            .Luces = 1
            .BufferSize = 12
            .DinamicInventory = True
            .PartyMembers = True
            .DialogosEnConsola = True
            .Nombres = 1
            .NombreItems = True
            .DragerDrop = True
            .UltimaCuenta = ""
            .UsarSombras = True
            .ParticleMeditation = True
            .Musica = True
            .Sonido = True
            .Sonido3D = True
            .SoundVolume = 99
            .MusicVolume = 99
            .Dialog_Align = 1
            .MiniMap = True
            .ParticleEngine = True
            .ProyectileEngine = True
            .TonalidadPJ = True
            .SeguroItems = True
            .FirstTime = True
            .GuildNews = True
            .DialogoClanesActivo = True
            .DialogoClanesCant = 5
            .useVideoMemory = True
            .MemoryVideoMax = 512
            .MostrarFPS = True
            
            .Water_Effect = True
            .Mouse_Effect = True
            .Shadow_Effect = True
            .Walk_Effect = True
            .Text_Effect = True
            .Reflect_Effect = True
            .Recordar = False
            
        End With
End Sub

