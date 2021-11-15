Attribute VB_Name = "Mod_GameFile"
'***************************************************
'Author: Ezequiel Juárez (Standelf)
'Last Modification: 11/01/2011
'Blisse-AO | GameFile Mod
'This module was created to put all the burdens of _
    Argentum Online resources.
'***************************************************

Option Explicit

Public Type tCabecera 'Cabecera de los con
    Desc As String * 255
    CRC As Long
    MagicWord As Long
End Type

Public MiCabecera As tCabecera

Public Sub IniciarCabecera(ByRef Cabecera As tCabecera)
    Cabecera.Desc = "Argentum Online by Noland Studios. Copyright Noland-Studios 2001, pablomarquez@noland-studios.com.ar"
    Cabecera.CRC = Rnd * 100
    Cabecera.MagicWord = Rnd * 10
End Sub

Public Function LoadGrhData() As Boolean
On Error Resume Next
    Dim Grh As Long
    Dim Frame As Long
    Dim grhCount As Long
    Dim handle As Integer
    Dim fileVersion As Long
    'Open files
    handle = FreeFile()
    Open Resources.Bin & "Graficos.ind" For Binary Access Read As handle
    Get handle, , fileVersion
    
    Get handle, , grhCount
    
    ReDim GrhData(0 To grhCount) As GrhData
    
    While Not EOF(handle)
        Get handle, , Grh
        
        With GrhData(Grh)
            GrhData(Grh).active = True
            Get handle, , .NumFrames
            If .NumFrames <= 0 Then Resume Next
            
            ReDim .Frames(1 To GrhData(Grh).NumFrames)
            
            If .NumFrames > 1 Then
                For Frame = 1 To .NumFrames
                    Get handle, , .Frames(Frame)
                    If .Frames(Frame) <= 0 Or .Frames(Frame) > grhCount Then
                        Resume Next
                    End If
                Next Frame
                
                Get handle, , .Speed
                
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
                Get handle, , .FileNum
                If .FileNum <= 0 Then Resume Next
                
                Get handle, , GrhData(Grh).sX
                If .sX < 0 Then Resume Next
                
                Get handle, , .sY
                If .sY < 0 Then Resume Next
                
                Get handle, , .pixelWidth
                If .pixelWidth <= 0 Then Resume Next
                
                Get handle, , .pixelHeight
                If .pixelHeight <= 0 Then Resume Next
                
                .TileWidth = .pixelWidth / TilePixelHeight
                .TileHeight = .pixelHeight / TilePixelWidth
                
                .Frames(1) = Grh
            End If
        End With
    Wend
    
    Close handle
    
    Dim Count As Long
    
    Open Resources.Bin & "minimap.dat" For Binary As #1
        Seek #1, 1
        For Count = 1 To UBound(GrhData())
            If GrhData(Count).active Then
                Get #1, , GrhData(Count).MiniMap_color
            End If
        Next Count
    Close #1
    
    LoadGrhData = True
Exit Function

ErrorHandler:
    LoadGrhData = False
End Function

Public Sub CargarAnimArmas()
On Error Resume Next

    Dim LoopC As Long
    Dim arch As String
    
    arch = Resources.Bin & "armas.dat"
    NumWeaponAnims = Val(General_Get_Var(arch, "INIT", "NumArmas"))
    
    ReDim WeaponAnimData(1 To NumWeaponAnims) As WeaponAnimData
    
    For LoopC = 1 To NumWeaponAnims
        InitGrh WeaponAnimData(LoopC).WeaponWalk(1), Val(General_Get_Var(arch, "ARMA" & LoopC, "Dir1")), 0
        InitGrh WeaponAnimData(LoopC).WeaponWalk(2), Val(General_Get_Var(arch, "ARMA" & LoopC, "Dir2")), 0
        InitGrh WeaponAnimData(LoopC).WeaponWalk(3), Val(General_Get_Var(arch, "ARMA" & LoopC, "Dir3")), 0
        InitGrh WeaponAnimData(LoopC).WeaponWalk(4), Val(General_Get_Var(arch, "ARMA" & LoopC, "Dir4")), 0
    Next LoopC
End Sub

Public Sub CargarColores()
On Error Resume Next
    Dim archivoC As String
    archivoC = Resources.Bin & "colores.dat"
    
    If Not General_File_Exist(archivoC, vbArchive) Then
        Call MsgBox("ERROR: no se ha podido cargar los colores. Falta el archivo colores.dat, reinstale el juego", vbCritical + vbOKOnly)
        Exit Sub
    End If
    
    Dim i As Long
    
    For i = 0 To 48 '49 y 50 reservados para ciudadano y criminal
        ColoresPJ(i) = D3DColorXRGB(General_Get_Var(archivoC, CStr(i), "R"), General_Get_Var(archivoC, CStr(i), "G"), General_Get_Var(archivoC, CStr(i), "B"))
    Next i
    
    '   Crimi
    ColoresPJ(50) = D3DColorXRGB(General_Get_Var(archivoC, "CR", "R"), General_Get_Var(archivoC, "CR", "G"), General_Get_Var(archivoC, "CR", "B"))

    '   Ciuda
    ColoresPJ(49) = D3DColorXRGB(General_Get_Var(archivoC, "CI", "R"), General_Get_Var(archivoC, "CI", "G"), General_Get_Var(archivoC, "CI", "B"))
    
    '   Atacable
    ColoresPJ(50) = D3DColorXRGB(General_Get_Var(archivoC, "AT", "R"), General_Get_Var(archivoC, "AT", "G"), General_Get_Var(archivoC, "AT", "B"))
End Sub

Public Sub CargarAnimEscudos()
On Error Resume Next

    Dim LoopC As Long
    Dim arch As String
    
    arch = Resources.Bin & "escudos.dat"
    
    ReDim ShieldAnimData(1 To Val(General_Get_Var(arch, "INIT", "NumEscudos"))) As ShieldAnimData
    
    For LoopC = 1 To UBound(ShieldAnimData())
        InitGrh ShieldAnimData(LoopC).ShieldWalk(1), Val(General_Get_Var(arch, "ESC" & LoopC, "Dir1")), 0
        InitGrh ShieldAnimData(LoopC).ShieldWalk(2), Val(General_Get_Var(arch, "ESC" & LoopC, "Dir2")), 0
        InitGrh ShieldAnimData(LoopC).ShieldWalk(3), Val(General_Get_Var(arch, "ESC" & LoopC, "Dir3")), 0
        InitGrh ShieldAnimData(LoopC).ShieldWalk(4), Val(General_Get_Var(arch, "ESC" & LoopC, "Dir4")), 0
    Next LoopC
End Sub

Public Sub CargarCabezas()
    Dim N As Integer
    Dim i As Long
    Dim Numheads As Integer
    Dim Miscabezas() As tIndiceCabeza
    
    N = FreeFile()
    Open Resources.Bin & "cabezas.ind" For Binary Access Read As #N
    
    'cabecera
    Get #N, , MiCabecera
    
    'num de cabezas
    Get #N, , Numheads
    
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
    Next i
    
    Close #N

End Sub

Public Sub CargarCascos()
    Dim N As Integer
    Dim i As Long
    Dim NumCascos As Integer

    Dim Miscabezas() As tIndiceCabeza
    
    N = FreeFile()
    Open Resources.Bin & "cascos.ind" For Binary Access Read As #N
    
    'cabecera
    Get #N, , MiCabecera
    
    'num de cabezas
    Get #N, , NumCascos
    
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
    Next i
    
    Close #N
End Sub

Public Sub CargarCuerpos()
    Dim N As Integer
    Dim i As Long
    Dim NumCuerpos As Integer
    Dim MisCuerpos() As tIndiceCuerpo

    N = FreeFile()
    Open Resources.Bin & "Cuerpos.ind" For Binary Access Read As #N
    
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
            InitGrh BodyData(i).Walk(1), MisCuerpos(i).Body(1), 0
            InitGrh BodyData(i).Walk(2), MisCuerpos(i).Body(2), 0
            InitGrh BodyData(i).Walk(3), MisCuerpos(i).Body(3), 0
            InitGrh BodyData(i).Walk(4), MisCuerpos(i).Body(4), 0
            
            BodyData(i).HeadOffset.X = MisCuerpos(i).HeadOffsetX
            BodyData(i).HeadOffset.Y = MisCuerpos(i).HeadOffsetY
        End If
    Next i
    
    Close #N
    
End Sub

Public Sub CargarFxs()
    Dim N As Integer
    Dim i As Long
    Dim NumFxs As Integer
        
    N = FreeFile()
    Open Resources.Bin & "Anims.ind" For Binary Access Read As #N
    
    'cabecera
    Get #N, , MiCabecera
    
    'num de cabezas
    Get #N, , NumFxs
    
    'Resize array
    ReDim FxData(1 To NumFxs) As tIndiceFx
    
    For i = 1 To NumFxs
        Get #N, , FxData(i)
    Next i
    
    Close #N
End Sub

