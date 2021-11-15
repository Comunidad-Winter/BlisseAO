Attribute VB_Name = "mod_Index"
Option Explicit

Public Type GrhData
    sX As Integer
    sY As Integer
    FileNum As Long
    pixelWidth As Integer
    pixelHeight As Integer
    TileWidth As Single
    TileHeight As Single
    NumFrames As Integer
    Frames() As Long
    Speed As Single
    Active As Boolean
End Type

'   Estructura del GRH
Public Type GRH
    GrhIndex As Integer
    FrameCounter As Single
    Speed As Single
    Started As Byte
    Loops As Integer
End Type

Public Enum E_Heading
    NORTH = 1
    EAST = 2
    SOUTH = 3
    WEST = 4
End Enum

Public Type tIndiceFx
    Animacion As Integer
    OffsetX As Integer
    OffsetY As Integer
End Type
Public Type tIndiceCabeza
    Head(1 To 4) As Integer
End Type

Public Type tIndiceCuerpo
    Body(1 To 4) As Integer
    HeadOffsetX As Integer
    HeadOffsetY As Integer
End Type

Public Type tCabecera
    Desc As String * 255
    CRC As Long
    MagicWord As Long
End Type
Public Type Position
    X As Integer
    Y As Integer
End Type
'   Lista de cuerpos
Public Type BodyData
    Walk(E_Heading.NORTH To E_Heading.WEST) As GRH
    HeadOffset As Position
End Type

'   Lista de cabezas
Public Type HeadData
    Head(E_Heading.NORTH To E_Heading.WEST) As GRH
End Type

'   Lista de las animaciones de las armas
Type WeaponAnimData
    WeaponWalk(E_Heading.NORTH To E_Heading.WEST) As GRH
End Type

'   Lista de las animaciones de los escudos
Type ShieldAnimData
    ShieldWalk(E_Heading.NORTH To E_Heading.WEST) As GRH
End Type


Public MiCabecera As tCabecera
Public NUM_GRHS As Long

Public NUM_FXS As Integer

Public NUM_ARM As Integer
Public NUM_BOD As Integer
Public NUM_CAS As Integer
Public NUM_ESC As Integer
Public NUM_HEA As Integer

Public GrhData() As GrhData
Public FxData() As tIndiceFx

Public BodyData() As BodyData
Public HeadData() As HeadData
Public WeaponAnimData() As WeaponAnimData
Public ShieldAnimData() As ShieldAnimData
Public CascoAnimData() As HeadData


Public TEMP_GRH As GRH
Public TEMP_ANIM As GrhData

Public Function LoadGrhData() As Boolean
On Error Resume Next
    Dim GRH As Long
    Dim Frame As Long
    Dim grhCount As Long
    Dim Handle As Integer
    Dim fileVersion As Long
    Dim TEMP_NAME As String
    'Open files
    Handle = FreeFile()
    Open GRH_FILE For Binary Access Read As Handle
    Get Handle, , fileVersion
    
    Get Handle, , grhCount
    NUM_GRHS = grhCount
    
    ReDim GrhData(0 To grhCount) As GrhData
    
    While Not EOF(Handle)
        Get Handle, , GRH
        
        With GrhData(GRH)
            GrhData(GRH).Active = True
            

            
            Get Handle, , .NumFrames
            If .NumFrames <= 0 Then Resume Next
            
            ReDim .Frames(1 To GrhData(GRH).NumFrames)
            
            If .NumFrames > 1 Then
                TEMP_NAME = ", [ANIM]"
                For Frame = 1 To .NumFrames
                    Get Handle, , .Frames(Frame)
                    If .Frames(Frame) <= 0 Or .Frames(Frame) > grhCount Then
                        Resume Next
                    End If
                Next Frame
                
                Get Handle, , .Speed
                
                If .Speed <= 0 Then Resume Next
                
                .pixelHeight = GrhData(.Frames(1)).pixelHeight
                If .pixelHeight <= 0 Then Resume Next
                
                .pixelWidth = GrhData(.Frames(1)).pixelWidth
                If .pixelWidth <= 0 Then Resume Next
                
                .TileWidth = GrhData(.Frames(1)).TileWidth
                If .TileWidth <= 0 Then Resume Next
                
                .TileHeight = GrhData(.Frames(1)).TileHeight
                If .TileHeight <= 0 Then Resume Next
            Else
                Get Handle, , .FileNum
                If .FileNum <= 0 Then Resume Next
                
                If .FileNum <= 0 Then
                    TEMP_NAME = ", [LIBRE]"
                Else
                    TEMP_NAME = ""
                End If
                
                Get Handle, , GrhData(GRH).sX
                If .sX < 0 Then Resume Next
                
                Get Handle, , .sY
                If .sY < 0 Then Resume Next
                
                Get Handle, , .pixelWidth
                If .pixelWidth <= 0 Then Resume Next
                
                Get Handle, , .pixelHeight
                If .pixelHeight <= 0 Then Resume Next
                
                .TileWidth = .pixelWidth / 32
                .TileHeight = .pixelHeight / 32
                
                .Frames(1) = GRH
            End If
            
            If GRH <> 0 Then frmMain.GRH_LIST.AddItem GRH & TEMP_NAME
        End With
        
        frmCargando.Label2.Caption = "Cargando GRH " & GRH & " / " & grhCount
        frmCargando.Shape2.Width = (((GRH / 100) / (grhCount / 100)) * 578)
    Wend
    
    Close Handle
    
    LoadGrhData = True
Exit Function

ErrorHandler:
    LoadGrhData = False
    MsgBox "ERROR"
    End
End Function

Public Sub CargarFxs()
    Dim N As Integer
    Dim i As Long
    Dim NumFxs As Integer
        
    N = FreeFile()
    Open FXS_FILE For Binary Access Read As #N
    
    'cabecera
    Get #N, , MiCabecera
    
    'num de cabezas
    Get #N, , NumFxs
    
    'Resize array
    ReDim FxData(1 To NumFxs) As tIndiceFx
    NUM_FXS = NumFxs
    
    For i = 1 To NumFxs
        Get #N, , FxData(i)
        
        frmMain.FXS_LIST.AddItem i
        
        
        frmCargando.Label2.Caption = "Cargando FX " & i & " / " & NumFxs
        frmCargando.Shape2.Width = (((i / 100) / (NumFxs / 100)) * 578)
    Next i
    
    Close #N
End Sub

Public Sub CargarAnimArmas()
On Error Resume Next

    Dim LoopC As Long
    NUM_ARM = Val(GetVar(WEA_FILE, "INIT", "NumArmas"))
    
    ReDim WeaponAnimData(1 To NUM_ARM) As WeaponAnimData
    
    For LoopC = 1 To NUM_ARM
        'InitGrh WeaponAnimData(LoopC).WeaponWalk(1), Val(GetVar(WEA_FILE, "ARMA" & LoopC, "Dir1")), 0
        'InitGrh WeaponAnimData(LoopC).WeaponWalk(2), Val(GetVar(WEA_FILE, "ARMA" & LoopC, "Dir2")), 0
        'InitGrh WeaponAnimData(LoopC).WeaponWalk(3), Val(GetVar(WEA_FILE, "ARMA" & LoopC, "Dir3")), 0
        'InitGrh WeaponAnimData(LoopC).WeaponWalk(4), Val(GetVar(WEA_FILE, "ARMA" & LoopC, "Dir4")), 0
    Next LoopC
End Sub

Public Sub CargarAnimEscudos()
On Error Resume Next

    Dim LoopC As Long

    
    ReDim ShieldAnimData(1 To Val(GetVar(SHI_FILE, "INIT", "NumEscudos"))) As ShieldAnimData
    
    For LoopC = 1 To UBound(ShieldAnimData())
        'InitGrh ShieldAnimData(LoopC).ShieldWalk(1), Val(GetVar(SHI_FILE, "ESC" & LoopC, "Dir1")), 0
        'InitGrh ShieldAnimData(LoopC).ShieldWalk(2), Val(GetVar(SHI_FILE, "ESC" & LoopC, "Dir2")), 0
        'InitGrh ShieldAnimData(LoopC).ShieldWalk(3), Val(GetVar(SHI_FILE, "ESC" & LoopC, "Dir3")), 0
        'InitGrh ShieldAnimData(LoopC).ShieldWalk(4), Val(GetVar(SHI_FILE, "ESC" & LoopC, "Dir4")), 0
    Next LoopC
End Sub

Public Sub CargarCabezas()
    Dim N As Integer
    Dim i As Long
    Dim Numheads As Integer
    Dim Miscabezas() As tIndiceCabeza
    
    N = FreeFile()
    Open HEA_FILE For Binary Access Read As #N
    
    'cabecera
    Get #N, , MiCabecera
    
    'num de cabezas
    Get #N, , Numheads
    NUM_HEA = Numheads
    
    'Resize array
    ReDim HeadData(0 To Numheads) As HeadData
    ReDim Miscabezas(0 To Numheads) As tIndiceCabeza
    
    For i = 1 To Numheads
        Get #N, , Miscabezas(i)
        
        If Miscabezas(i).Head(1) Then
            Call InitGrh(HeadData(i).Head(1), Miscabezas(i).Head(1), 0)
            Call InitGrh(HeadData(i).Head(2), Miscabezas(i).Head(2), 0)
            Call InitGrh(HeadData(i).Head(3), Miscabezas(i).Head(3), 0)
            Call InitGrh(HeadData(i).Head(4), Miscabezas(i).Head(4), 0)
        End If
    
        frmMain.HHH_LIST.AddItem i
        
        
        frmCargando.Label2.Caption = "Cargando Cabeza " & i & " / " & Numheads
        frmCargando.Shape2.Width = (((i / 100) / (Numheads / 100)) * 578)
        
    Next i
    
    Close #N

End Sub

Public Sub CargarCascos()
    Dim N As Integer
    Dim i As Long
    Dim NumCascos As Integer

    Dim Miscabezas() As tIndiceCabeza
    
    N = FreeFile()
    Open HEL_FILE For Binary Access Read As #N
    
    'cabecera
    Get #N, , MiCabecera
    
    'num de cabezas
    Get #N, , NumCascos
    NUM_CAS = NumCascos
    'Resize array
    ReDim CascoAnimData(0 To NumCascos) As HeadData
    ReDim Miscabezas(0 To NumCascos) As tIndiceCabeza
    
    For i = 1 To NumCascos
        Get #N, , Miscabezas(i)
        
        If Miscabezas(i).Head(1) Then
            Call InitGrh(CascoAnimData(i).Head(1), Miscabezas(i).Head(1), 0)
            Call InitGrh(CascoAnimData(i).Head(2), Miscabezas(i).Head(2), 0)
            Call InitGrh(CascoAnimData(i).Head(3), Miscabezas(i).Head(3), 0)
            Call InitGrh(CascoAnimData(i).Head(4), Miscabezas(i).Head(4), 0)
        End If
        
        frmCargando.Label2.Caption = "Cargando Casco " & i & " / " & NumCascos
        frmCargando.Shape2.Width = (((i / 100) / (NumCascos / 100)) * 578)
    Next i
    
    Close #N
End Sub

Public Sub CargarCuerpos()
    Dim N As Integer
    Dim i As Long
    Dim NumCuerpos As Integer
    Dim MisCuerpos() As tIndiceCuerpo

    N = FreeFile()
    Open BOD_FILE For Binary Access Read As #N
    
    'cabecera
    Get #N, , MiCabecera
    
    'num de cabezas
    Get #N, , NumCuerpos
    
    'Resize array
    ReDim BodyData(0 To NumCuerpos) As BodyData
    ReDim MisCuerpos(0 To NumCuerpos) As tIndiceCuerpo
    
    For i = 1 To NumCuerpos
        Get #N, , MisCuerpos(i)
        
        If MisCuerpos(i).Body(1) Then
            'InitGrh BodyData(i).Walk(1), MisCuerpos(i).Body(1), 0
            'InitGrh BodyData(i).Walk(2), MisCuerpos(i).Body(2), 0
            'InitGrh BodyData(i).Walk(3), MisCuerpos(i).Body(3), 0
            'InitGrh BodyData(i).Walk(4), MisCuerpos(i).Body(4), 0
            
            BodyData(i).HeadOffset.X = MisCuerpos(i).HeadOffsetX
            BodyData(i).HeadOffset.Y = MisCuerpos(i).HeadOffsetY
        End If
    Next i
    
    Close #N
    
End Sub

