Attribute VB_Name = "Mod_DX8_Ambient"
'***************************************************
'Author: Ezequiel Juárez (Standelf)
'Last Modification: ??/??/10
'Blisse-AO | Sistema de Ambientes
'***************************************************

Option Explicit

Enum eAreas
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
    Armaduras = 14
    Mercado = 15
    Provisiones = 16
    Carpinteria = 17
    House1 = 18
    House2 = 19
    House3 = 20
    House4 = 21
    House5 = 22
    House6 = 23
End Enum

Type A_Light
    Range                   As Byte
    r                       As Integer
    g                       As Integer
    b                       As Integer
End Type

Type MapAmbientBlock
    Light                   As A_Light
    Particle                As Byte
    Vertex_Offset(0 To 3)   As Long
    Area                    As Byte
End Type

Type MapAmbient
    MapBlocks()             As MapAmbientBlock
    UseDayAmbient           As Boolean
    OwnAmbientLight         As D3DCOLORVALUE
    Fog                     As Integer
    Snow                    As Boolean
    Rain                    As Boolean
    Music                   As Byte
End Type

Public Setting_Map_Areas As Boolean
Public LastArea As eAreas
Public CurMapAmbient As MapAmbient
    
Public Sub Ambient_Aply_OwnAmbient()
        If CurMapAmbient.UseDayAmbient = False Then
            Estado_Actual = CurMapAmbient.OwnAmbientLight
        Else
            Call Actualizar_Estado(Estado_Actual_Date)
        End If
    
        Dim xx As Integer, yy As Integer
            For xx = XMinMapSize To XMaxMapSize
                For yy = YMinMapSize To YMaxMapSize
                    If CurMapAmbient.UseDayAmbient = False Then
                        Call Engine_D3DColor_To_RGB_List(MapData(xx, yy).Engine_Light(), CurMapAmbient.OwnAmbientLight)
                    End If
                Next yy
            Next xx
            
        Call LightRenderAll
End Sub

Public Sub Ambient_Init(ByVal Map As Integer)
'***************************************************
'Author: Standelf
'Last Modification: 15/10/10
'***************************************************
    With CurMapAmbient
        ' #### Set Default Ambient
        .Fog = -1
        .UseDayAmbient = True
        .OwnAmbientLight.a = 255
        .OwnAmbientLight.r = 0
        .OwnAmbientLight.g = 0
        .OwnAmbientLight.b = 0
        
        .Rain = True
        .Snow = False
        
        
        ' #### Redim Blocks
        ReDim .MapBlocks(XMinMapSize To XMaxMapSize, YMinMapSize To YMaxMapSize) As MapAmbientBlock
        
        
        ' #### Load File
        If General_File_Exist(Resources.Ambient & Map & ".amb", vbNormal) Then
            Dim N As Integer
            N = FreeFile
                Open Resources.Ambient & Map & ".amb" For Binary As #N
                    Get #N, , CurMapAmbient
                Close #N
        End If
        
        If .Music <> 0 Then
            Play_MP3 .Music
            frmAmbientEditor.Text5.Text = .Music
            frmAmbientEditor.List1.Selected(.Music - 1) = True
        Else
            frmAmbientEditor.Text5.Text = 0
            frmAmbientEditor.List1.Selected(0) = True
        End If
        
        If .UseDayAmbient = False Then
            Estado_Actual = .OwnAmbientLight
        Else
            Call Actualizar_Estado(Estado_Actual_Date)
        End If
                    
        Dim xx As Integer, yy As Integer
        
            For xx = XMinMapSize To XMaxMapSize
                For yy = YMinMapSize To YMaxMapSize
                    If .UseDayAmbient = False Then
                        Call Engine_D3DColor_To_RGB_List(MapData(xx, yy).Engine_Light(), .OwnAmbientLight)
                    End If
                    
                    Dim Vertex As Long
                    For Vertex = 0 To 3
                        If .MapBlocks(xx, yy).Vertex_Offset(Vertex) <> 0 Then
                            MapData(xx, yy).Vertex_Offset(Vertex) = .MapBlocks(xx, yy).Vertex_Offset(Vertex)
                        End If
                    Next Vertex
                    
                    If .MapBlocks(xx, yy).Light.Range <> 0 Then
                        Create_Light_To_Map xx, yy, .MapBlocks(xx, yy).Light.Range, .MapBlocks(xx, yy).Light.r, .MapBlocks(xx, yy).Light.g, .MapBlocks(xx, yy).Light.b
                    End If
                Next yy
            Next xx
            
        Call LightRenderAll
            
            If .UseDayAmbient = True Then
                frmAmbientEditor.Option1(0).value = True
            Else
                frmAmbientEditor.Option1(1).value = True
                frmAmbientEditor.Text1(0).Text = .OwnAmbientLight.r
                frmAmbientEditor.Text1(1).Text = .OwnAmbientLight.g
                frmAmbientEditor.Text1(2).Text = .OwnAmbientLight.b
            End If
                                        
            If .Fog <> -1 Then
                frmAmbientEditor.Check1.value = Checked
                frmAmbientEditor.HScroll1.value = .Fog
            Else
                frmAmbientEditor.Check1.value = Unchecked
            End If
            
            
            
            If .Rain = True Then frmAmbientEditor.Check3.value = Checked
            If .Snow = True Then frmAmbientEditor.Check2.value = Checked
            
            
    End With
End Sub

Public Sub Ambient_Save(ByVal Map As Integer)
'***************************************************
'Author: Standelf
'Last Modification: 15/10/10
'***************************************************
Dim File
File = FreeFile
    Open Resources.Ambient & Map & ".amb" For Binary Access Write As File
        Put File, , CurMapAmbient
    Close #File
End Sub



Public Function Get_Area_Color(ByVal Area As eAreas) As Long
'***************************************************
'Author: Standelf
'Last Modification: 16/11/12
'***************************************************
    Select Case Area
        Case eAreas.Alquimia
            Get_Area_Color = ColorData.AmarilloAB(1)
        Case eAreas.Banco
            Get_Area_Color = ColorData.DoradoAB(1)
        Case eAreas.House1 To eAreas.House6
            Get_Area_Color = ColorData.RojoAB(1)
        Case eAreas.Carpinteria
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

Public Sub Ambient_Set_Area(ByVal X As Byte, Y As Byte, ByVal Area As eAreas, ByVal Rango As Byte)
'***************************************************
'Author: Standelf
'Last Modification: 16/11/12
'***************************************************

    If Rango = 1 Or Rango = 2 Or Rango = 3 Then
        CurMapAmbient.MapBlocks(X, Y).Area = Area
    End If
    
    If Rango = 2 Or Rango = 3 Then
        CurMapAmbient.MapBlocks(X + 1, Y).Area = Area
        CurMapAmbient.MapBlocks(X, Y + 1).Area = Area
        CurMapAmbient.MapBlocks(X + 1, Y + 1).Area = Area
    End If
    
    If Rango = 3 Then
        CurMapAmbient.MapBlocks(X, Y - 1).Area = Area
        CurMapAmbient.MapBlocks(X + 1, Y - 1).Area = Area
        CurMapAmbient.MapBlocks(X - 1, Y - 1).Area = Area
    
        CurMapAmbient.MapBlocks(X - 1, Y).Area = Area
        CurMapAmbient.MapBlocks(X - 1, Y + 1).Area = Area
    End If
End Sub


Public Sub Ambient_Check_Music(ByVal X As Byte, Y As Byte)
'***************************************************
'Author: Standelf
'Last Modification: 16/11/1/20
'***************************************************

If Settings.Musica = False Then Exit Sub

    If CurMapAmbient.MapBlocks(X, Y).Area = eAreas.Bar Then
        Play_MP3 eMP3.Bar
    End If
    
    If LastArea = Bar And CurMapAmbient.MapBlocks(X, Y).Area <> Bar Then
        Play_MP3 CurMapAmbient.Music
    End If
        
    LastArea = CurMapAmbient.MapBlocks(X, Y).Area
End Sub
