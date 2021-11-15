Attribute VB_Name = "Mod_Map_System"
Option Explicit

Public Enum eAreas
    Iglesia = 1
    Banco = 2
    Herreria = 3
    Hechizeria = 4
    Alquimia = 5
    Sastreria = 6
    Quest = 7
    Entrenador = 8
    EntrenadorSkills = 9
    EntrenadorSpells = 10
    Curandero = 11
    Identificador = 12
    Bar = 13
    Wall = 20
    House1 = 30
    House2 = 31
    House3 = 32
    House4 = 33
    House5 = 34
    House6 = 35
End Enum
    
Public MapArea() As eAreas
Public Setting_Map_Areas As Boolean

Public LastArea As eAreas

Public Sub Init_Map_System()
    ReDim MapArea(XMinMapSize To XMaxMapSize, YMinMapSize To YMaxMapSize) As eAreas
    
End Sub

Public Function Get_Area_Color(ByVal Area As eAreas) As Long
    Select Case Area
        Case eAreas.Alquimia
            Get_Area_Color = ColorData.AmarilloAB(1)
        Case eAreas.Banco
            Get_Area_Color = ColorData.DoradoAB(1)
        Case eAreas.House1 To eAreas.House6
            Get_Area_Color = ColorData.RojoAB(1)
        Case eAreas.Wall
            Get_Area_Color = ColorData.GrisAB(1)
        Case eAreas.Iglesia
            Get_Area_Color = ColorData.CelesteAB(1)
        Case eAreas.Sastreria
            Get_Area_Color = ColorData.BlancoAB(1)
        Case eAreas.Identificador
            Get_Area_Color = ColorData.VerdeAB(1)
        Case eAreas.Bar
            Get_Area_Color = ColorData.AzulAB(1)
        Case Else
            Get_Area_Color = D3DColorARGB(0, 0, 0, 0)
    End Select
End Function

Public Sub Set_Map_Area(ByVal X As Byte, Y As Byte, ByVal Area As eAreas)
    MapArea(X, Y) = Area
    MapArea(X + 1, Y) = Area
    MapArea(X, Y + 1) = Area
    MapArea(X + 1, Y + 1) = Area
End Sub

Public Sub Map_Area_Read(ByVal Map As Integer)
Dim Leer As New clsIniReader, Temp As Integer
    If Not General_File_Exist(Dir_Resources & "Ambiente\" & Map & ".area", vbNormal) Then
        Dim X As Long
        Dim Y As Long
            For X = 1 To 100
            For Y = 1 To 100
                MapArea(X, Y) = 0
            Next Y
            Next X
    Else
        Dim N As Integer
        N = FreeFile
            Open Dir_Resources & "Ambiente\" & Map & ".area" For Binary As #N
                Get #N, , MapArea()
            Close #N
    End If
End Sub

Public Sub Map_Area_Save(ByVal Map As Integer)
'***************************************************
'Author: Standelf
'Last Modification: 15/10/10
'***************************************************
Dim File
File = FreeFile
    Open Dir_Resources & "Ambiente\" & Map & ".area" For Binary Access Write As File
        Put File, , MapArea()
    Close #File
End Sub

Public Sub CheckAreaMusic(ByVal X As Byte, Y As Byte)
If Settings.Musica = False Then Exit Sub

    If MapArea(X, Y) = eAreas.Bar Then
        Play_MP3 eMP3.Bar
    End If
    
    
    If LastArea = Bar And MapArea(X, Y) <> Bar Then
        Stop_MP3
        
        Play_MP3 CurMapAmbient.Music
    End If
        
    LastArea = MapArea(X, Y)
End Sub

