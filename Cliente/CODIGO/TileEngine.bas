Attribute VB_Name = "Mod_DX8_TileEngine"
Option Explicit
Public AlphaPres As Byte


Private Fade_GUI(1 To 12) As Integer
Private Fade_GUIe(1 To 12) As Boolean

'Caminata fluida
Public Engine_Movement_Speed As Single

'Quad Draw
Public indexList(0 To 5) As Integer
Public ibQuad As DxVBLibA.Direct3DIndexBuffer8
Public vbQuadIdx As DxVBLibA.Direct3DVertexBuffer8
Dim Temp_Verts(3) As D3DTLVERTEX

'Status del user
Public UserIndex As Integer
Public UserMoving As Byte
Public UserBody As Integer
Public UserHead As Integer
Public UserPos As Position
Public AddtoUserPos As Position 'Si se mueve
Public UserCharIndex As Integer

Private OffsetCounterX As Single
Private OffsetCounterY As Single
    
Public WeatherFogX1 As Single
Public WeatherFogY1 As Single
Public WeatherFogX2 As Single
Public WeatherFogY2 As Single
Public WeatherFogCount As Byte

Public ParticleOffsetX As Long
Public ParticleOffsetY As Long
Public LastOffsetX As Integer
Public LastOffsetY As Integer

Public EngineRun As Boolean

Public FramesPerSecond As Single
Public FramesPerSecondCounter As Long
Public FramesPerSecondLastTime As Long

'Tamaño del la vista en Tiles
Private WindowTileWidth As Integer
Private WindowTileHeight As Integer
Public HalfWindowTileWidth As Integer
Public HalfWindowTileHeight As Integer

'Tamaño de los tiles en pixels
Public TilePixelHeight As Integer
Public TilePixelWidth As Integer

'Number of pixels the engine scrolls per frame. MUST divide evenly into pixels per tile
Public ScrollPixelsPerFrameX As Integer
Public ScrollPixelsPerFrameY As Integer

Public timerElapsedTime As Single
Dim timerTicksPerFrame As Single


'   Data
Public NumBodies As Integer
Public Numheads As Integer
Public NumFxs As Integer

Public NumChars As Integer
Public LastChar As Integer
Public NumWeaponAnims As Integer
Public NumShieldAnims As Integer

Private MouseTileX As Byte
Private MouseTileY As Byte

'*******************************************************************
'   Grh Data & Config
'*******************************************************************

'   Fogatas
Private Const GrhFogata As Integer = 1521

'   Animaciones infinitas
Private Const INFINITE_LOOPS As Integer = -1

'   Contiene info acerca de donde se puede encontrar un grh tamaño y animacion
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
    MiniMap_color As Long
End Type

'   Estructura del GRH
Public Type Grh
    GrhIndex As Integer
    FrameCounter As Single
    Speed As Single
    Started As Byte
    Loops As Integer
End Type

'   Posicion
Public Type Position
    X As Long
    Y As Long
End Type

'   Lista de cuerpos
Public Type BodyData
    Walk(E_Heading.NORTH To E_Heading.WEST) As Grh
    HeadOffset As Position
End Type

'   Lista de cabezas
Public Type HeadData
    Head(E_Heading.NORTH To E_Heading.WEST) As Grh
End Type

'   Lista de las animaciones de las armas
Type WeaponAnimData
    WeaponWalk(E_Heading.NORTH To E_Heading.WEST) As Grh
End Type

'   Lista de las animaciones de los escudos
Type ShieldAnimData
    ShieldWalk(E_Heading.NORTH To E_Heading.WEST) As Grh
End Type

'   Info de un objeto
Public Type obj
    OBJIndex As Integer
    Amount As Integer
    name As String
End Type

Public GrhData() As GrhData
Public BodyData() As BodyData
Public HeadData() As HeadData
Public FxData() As tIndiceFx
Public WeaponAnimData() As WeaponAnimData
Public ShieldAnimData() As ShieldAnimData
Public CascoAnimData() As HeadData


'*******************************************************************
'   Map Data & Config
'*******************************************************************

'   Tamaño de los mapas
Public Const XMaxMapSize As Byte = 100
Public Const XMinMapSize As Byte = 1
Public Const YMaxMapSize As Byte = 100
Public Const YMinMapSize As Byte = 1

'   Posicion en el Mundo
Public Type WorldPos
    Map As Integer
    X As Integer
    Y As Integer
End Type

'   Límites del mapa
Public MinXBorder As Byte
Public MaxXBorder As Byte
Public MinYBorder As Byte
Public MaxYBorder As Byte




'   Tipo de las celdas del mapa
Public Type MapBlock
    Graphic(1 To 4) As Grh
    CharIndex As Integer
    ObjGrh As Grh
    
    NPCIndex As Integer
    OBJInfo As obj
    TileExit As WorldPos
    Blocked As Byte
    
    Trigger As Integer
    
    Vertex_Offset(0 To 3) As Long
    Effect As Long
    
    Engine_Light(0 To 3) As Long 'Standelf, Light Engine.
End Type

'Info de cada mapa
Public Type MapInfo
    Music As String
    name As String
    StartPos As WorldPos
    MapVersion As Integer
End Type

Public MapData() As MapBlock
Public MapInfo As MapInfo

'*******************************************************************
'   Char Data & Config
'*******************************************************************

'Apariencia del personaje
Public Type Char
    Active As Byte
    Heading As E_Heading
    Pos As Position
    Vertex_Offset As Integer

    iHead As Integer
    iBody As Integer
    Body As BodyData
    Head As HeadData
    Casco As HeadData
    Arma As WeaponAnimData
    Escudo As ShieldAnimData
    UsandoArma As Boolean
    
    fX As Grh
    FxIndex As Integer
    
    Criminal As Byte
    Atacable As Boolean
    
    Nombre As String
    
    scrollDirectionX As Integer
    scrollDirectionY As Integer
    
    Moving As Byte
    MoveOffsetX As Single
    MoveOffsetY As Single
    
    Pie As Boolean
    muerto As Boolean
    invisible As Boolean
    priv As Byte
    
    'Standelf
    Aura(1 To 4) As Aura
    ParticleIndex As Integer
    
    minHP As Long
    maxHP As Long
    ShowBar As Integer
    tmpName As String
    
End Type

Public CharList(1 To 10000) As Char
'   Control de Lluvia
Public bRain As Boolean

Public Enum PlayLoop
    plNone = 0
    plLluviain = 1
    plLluviaout = 2
End Enum

'   Control de Techos
Public bTecho As Boolean

'   Control de Fogatas
Public bFogata As Boolean
'   Api
Private Declare Function QueryPerformanceFrequency Lib "kernel32" (lpFrequency As Currency) As Long
Private Declare Function QueryPerformanceCounter Lib "kernel32" (lpPerformanceCount As Currency) As Long

Sub ConvertCPtoTP(ByVal viewPortX As Integer, ByVal viewPortY As Integer, ByRef tX As Byte, ByRef tY As Byte)
'******************************************
'Converts where the mouse is in the main window to a tile position. MUST be called eveytime the mouse moves.
'******************************************

    If MainScreenRect.Right - ScreenWidth <> 0 Then
        tX = UserPos.X + viewPortX \ TilePixelWidth - Round(MainScreenRect.Right / 32, 0) \ 2 + IIf((MainScreenRect.Right - ScreenWidth) > 0, 1, 0)
        tY = UserPos.Y + viewPortY \ TilePixelHeight - Round(MainScreenRect.bottom / 32, 0) \ 2
    Else
        tX = UserPos.X + viewPortX \ TilePixelWidth - WindowTileWidth \ 2
        tY = UserPos.Y + viewPortY \ TilePixelHeight - WindowTileHeight \ 2
    End If
End Sub

Sub MakeChar(ByVal CharIndex As Integer, ByVal Body As Integer, ByVal Head As Integer, ByVal Heading As Byte, ByVal X As Integer, ByVal Y As Integer, ByVal Arma As Integer, ByVal Escudo As Integer, ByVal Casco As Integer)
    'Apuntamos al ultimo Char
    If CharIndex > LastChar Then LastChar = CharIndex
    
    With CharList(CharIndex)
        If .Active = 0 Then NumChars = NumChars + 1
        
        If Arma = 0 Then Arma = 2
        If Escudo = 0 Then Escudo = 2
        If Casco = 0 Then Casco = 2
        
        .iHead = Head
        .iBody = Body
        .Head = HeadData(Head)

        .Body = BodyData(Body)
        .Arma = WeaponAnimData(Arma)
        
        .Escudo = ShieldAnimData(Escudo)
        .Casco = CascoAnimData(Casco)
        
        .Heading = Heading
        
        .Moving = 0
        .MoveOffsetX = 0
        .MoveOffsetY = 0
        
        .muerto = (Head = CASPER_HEAD)
        .Pos.X = X
        .Pos.Y = Y
        
        .Active = 1
        
    End With
    
    'Plot on map
    MapData(X, Y).CharIndex = CharIndex
End Sub

Sub ResetCharInfo(ByVal CharIndex As Integer)
    Delete_All_Auras CharIndex
    
    With CharList(CharIndex)
        .Active = 0
        .Criminal = 0
        .Atacable = False
        .FxIndex = 0
        .invisible = False
        .Moving = 0
        .muerto = False
        .Nombre = ""
        .Pie = False
        .Pos.X = 0
        .Pos.Y = 0
        .UsandoArma = False
        
        .minHP = 0
        .maxHP = 0
        .ShowBar = 0
        
        If .ParticleIndex <> 0 Then
            Call Effect_Kill(.ParticleIndex, False)
            .ParticleIndex = 0
        End If
        
    End With
End Sub

Sub EraseChar(ByVal CharIndex As Integer)
'*****************************************************************
'Erases a character from CharList and map
'*****************************************************************
    CharList(CharIndex).Active = 0
    
    'Update lastchar
    If CharIndex = LastChar Then
        Do Until CharList(LastChar).Active = 1
            LastChar = LastChar - 1
            If LastChar = 0 Then Exit Do
        Loop
    End If
    
    
    
    If CharList(CharIndex).ParticleIndex <> 0 Then
        Call Effect_Kill(CharList(CharIndex).ParticleIndex, False)
    End If
    
    MapData(CharList(CharIndex).Pos.X, CharList(CharIndex).Pos.Y).CharIndex = 0
    
    'Remove char's dialog
    Call Dialogos.RemoveDialog(CharIndex)
    
    Call ResetCharInfo(CharIndex)
    
    'Update NumChars
    NumChars = NumChars - 1
End Sub

Public Sub InitGrh(ByRef Grh As Grh, ByVal GrhIndex As Integer, Optional ByVal Started As Byte = 2)
'*****************************************************************
'Sets up a grh. MUST be done before rendering
'*****************************************************************
    Grh.GrhIndex = GrhIndex
    
    If Started = 2 Then
        If GrhData(Grh.GrhIndex).NumFrames > 1 Then
            Grh.Started = 1
        Else
            Grh.Started = 0
        End If
    Else
        If GrhData(Grh.GrhIndex).NumFrames = 1 Then Started = 0
        Grh.Started = Started
    End If
    
    
    If Grh.Started Then
        Grh.Loops = INFINITE_LOOPS
    Else
        Grh.Loops = 0
    End If
    
    Grh.FrameCounter = 1
    Grh.Speed = GrhData(Grh.GrhIndex).Speed

End Sub

Sub MoveCharbyHead(ByVal CharIndex As Integer, ByVal nHeading As E_Heading)
'*****************************************************************
'Starts the movement of a character in nHeading direction
'*****************************************************************
    Dim addX As Integer
    Dim addY As Integer
    Dim X As Integer
    Dim Y As Integer
    Dim nX As Integer
    Dim nY As Integer
    
    With CharList(CharIndex)
        X = .Pos.X
        Y = .Pos.Y
        
        If Settings.Walk_Effect And (TileEngine_Is_Dessert(X, Y) Or TileEngine_Is_Snow(X, Y)) Then
            DX8_Effects_Walk_Create .Heading, X, Y
        End If
        
        'Figure out which way to move
        Select Case nHeading
            Case E_Heading.NORTH
                addY = -1
        
            Case E_Heading.EAST
                addX = 1
        
            Case E_Heading.SOUTH
                addY = 1
            
            Case E_Heading.WEST
                addX = -1
        End Select

        
        nX = X + addX
        nY = Y + addY
        
        If nX = 0 Then nX = 1
        If nY = 0 Then nY = 1
        MapData(nX, nY).CharIndex = CharIndex
        .Pos.X = nX
        .Pos.Y = nY
        MapData(X, Y).CharIndex = 0
        
        .MoveOffsetX = -1 * (TilePixelWidth * addX)
        .MoveOffsetY = -1 * (TilePixelHeight * addY)
        
        .Moving = 1
        .Heading = nHeading
        
        .scrollDirectionX = addX
        .scrollDirectionY = addY
    End With
    
    If UserEstado = 0 Then Call DoPasosFx(CharIndex)
    
    'areas viejos
    If (nY < MinLimiteY) Or (nY > MaxLimiteY) Or (nX < MinLimiteX) Or (nX > MaxLimiteX) Then
        If CharIndex <> UserCharIndex Then
            Call EraseChar(CharIndex)
        End If
    End If
End Sub

Public Sub DoFogataFx()
    Dim location As Position
    
    If bFogata Then
        bFogata = HayFogata(location)
        If Not bFogata Then
            Call Audio.StopWave(FogataBufferIndex)
            FogataBufferIndex = 0
        End If
    Else
        bFogata = HayFogata(location)
        If bFogata And FogataBufferIndex = 0 Then FogataBufferIndex = Audio.PlayWave("fuego.wav", location.X, location.Y, LoopStyle.Enabled)
    End If
End Sub

Private Function EstaPCarea(ByVal CharIndex As Integer) As Boolean
    With CharList(CharIndex).Pos
        EstaPCarea = .X > UserPos.X - MinXBorder And .X < UserPos.X + MinXBorder And .Y > UserPos.Y - MinYBorder And .Y < UserPos.Y + MinYBorder
    End With
End Function

Private Sub SonidoPasos(ByVal PosX As Integer, PosY As Integer, Pie As Boolean)
    If TileEngine_Is_Grass(PosX, PosY) Then
        If Pie Then
            Call Audio.PlayWave(SND_PASTO1, PosX, PosY)
        Else
            Call Audio.PlayWave(SND_PASTO2, PosX, PosY)
        End If
    ElseIf TileEngine_Is_Snow(PosX, PosY) Then
        If Pie Then
            Call Audio.PlayWave(SND_NIEVE1, PosX, PosY)
        Else
            Call Audio.PlayWave(SND_NIEVE2, PosX, PosY)
        End If
    ElseIf TileEngine_Is_Dessert(PosX, PosY) Then
        If Pie Then
            Call Audio.PlayWave(SND_NIEVE1, PosX, PosY)
        Else
            Call Audio.PlayWave(SND_NIEVE2, PosX, PosY)
        End If
    Else
        If Pie Then
            Call Audio.PlayWave(SND_PASOS1, PosX, PosY)
        Else
            Call Audio.PlayWave(SND_PASOS2, PosX, PosY)
        End If
    End If
End Sub

Public Function TileEngine_Is_Grass(ByVal X As Integer, ByVal Y As Integer) As Boolean
    If X > 0 And X < 101 And Y > 0 And Y < 101 Then
    
    If MapData(X, Y).Graphic(1).GrhIndex >= 6000 And MapData(X, Y).Graphic(1).GrhIndex <= 6559 Then _
        TileEngine_Is_Grass = True
    Else
        TileEngine_Is_Grass = False
    End If
End Function

Public Function TileEngine_Is_Snow(ByVal X As Integer, ByVal Y As Integer) As Boolean
    If X > 0 And X < 101 And Y > 0 And Y < 101 Then
    
    If MapData(X, Y).Graphic(1).GrhIndex >= 26168 And MapData(X, Y).Graphic(1).GrhIndex <= 26711 Then _
        TileEngine_Is_Snow = True
    Else
        TileEngine_Is_Snow = False
    End If
End Function

Public Function TileEngine_Is_Dessert(ByVal X As Integer, ByVal Y As Integer) As Boolean
    If X > 0 And X < 101 And Y > 0 And Y < 101 Then
    
    If MapData(X, Y).Graphic(1).GrhIndex >= 7704 And MapData(X, Y).Graphic(1).GrhIndex <= 7719 Then _
        TileEngine_Is_Dessert = True
    Else
        TileEngine_Is_Dessert = False
    End If
End Function

Function TileEngine_Is_Water(ByVal X As Integer, ByVal Y As Integer) As Boolean
    If X > 0 And X < 101 And Y > 0 And Y < 101 Then
    
    TileEngine_Is_Water = ((MapData(X, Y).Graphic(1).GrhIndex >= 1505 And MapData(X, Y).Graphic(1).GrhIndex <= 1520) Or _
            (MapData(X, Y).Graphic(1).GrhIndex >= 5665 And MapData(X, Y).Graphic(1).GrhIndex <= 5680) Or _
            (MapData(X, Y).Graphic(1).GrhIndex >= 13547 And MapData(X, Y).Graphic(1).GrhIndex <= 13562)) And _
                MapData(X, Y).Graphic(2).GrhIndex = 0
    Else
      TileEngine_Is_Water = False
    End If
End Function

Function TileEngine_Is_Magma(ByVal X As Integer, ByVal Y As Integer) As Boolean
    If X > 0 And X < 101 And Y > 0 And Y < 101 Then
 
    TileEngine_Is_Magma = MapData(X, Y).Graphic(1).GrhIndex >= 5837 And MapData(X, Y).Graphic(1).GrhIndex <= 5852
    Else
      TileEngine_Is_Magma = False
    End If
End Function

 
Sub DoPasosFx(ByVal CharIndex As Integer)
    If Not UserNavegando Then
        With CharList(CharIndex)
            If Not .muerto And EstaPCarea(CharIndex) Then
                .Pie = Not .Pie
                
                SonidoPasos .Pos.X, CInt(.Pos.Y), .Pie
            End If
        End With
    Else
        Call Audio.PlayWave(SND_NAVEGANDO, CharList(CharIndex).Pos.X, CharList(CharIndex).Pos.Y)
    End If
    
    If CharIndex = UserCharIndex Then
        Call Ambient_Check_Music(CByte(UserPos.X), CByte(UserPos.Y))
    End If
End Sub

Sub MoveCharbyPos(ByVal CharIndex As Integer, ByVal nX As Integer, ByVal nY As Integer)
    Dim X As Integer
    Dim Y As Integer
    Dim addX As Integer
    Dim addY As Integer
    Dim nHeading As E_Heading
    
    With CharList(CharIndex)
        X = .Pos.X
        Y = .Pos.Y
        
        MapData(X, Y).CharIndex = 0
        
        addX = nX - X
        addY = nY - Y
        
        If Sgn(addX) = 1 Then
            nHeading = E_Heading.EAST
        ElseIf Sgn(addX) = -1 Then
            nHeading = E_Heading.WEST
        ElseIf Sgn(addY) = -1 Then
            nHeading = E_Heading.NORTH
        ElseIf Sgn(addY) = 1 Then
            nHeading = E_Heading.SOUTH
        End If
        
        MapData(nX, nY).CharIndex = CharIndex
        
        .Pos.X = nX
        .Pos.Y = nY
        
        .MoveOffsetX = -1 * (TilePixelWidth * addX)
        .MoveOffsetY = -1 * (TilePixelHeight * addY)
        
        .Moving = 1
        .Heading = nHeading
        
        .scrollDirectionX = Sgn(addX)
        .scrollDirectionY = Sgn(addY)
        
        'parche para que no medite cuando camina
        If .FxIndex = FxMeditar.CHICO Or .FxIndex = FxMeditar.GRANDE Or .FxIndex = FxMeditar.MEDIANO Or .FxIndex = FxMeditar.XGRANDE Or .FxIndex = FxMeditar.XXGRANDE Then
            .FxIndex = 0
        End If
    End With
    
    If Not EstaPCarea(CharIndex) Then Call Dialogos.RemoveDialog(CharIndex)
    
    If (nY < MinLimiteY) Or (nY > MaxLimiteY) Or (nX < MinLimiteX) Or (nX > MaxLimiteX) Then
        Call EraseChar(CharIndex)
    End If
End Sub

Sub MoveScreen(ByVal nHeading As E_Heading)
'******************************************
'Starts the screen moving in a direction
'******************************************
    Dim X As Integer
    Dim Y As Integer
    Dim tX As Integer
    Dim tY As Integer
    
    'Figure out which way to move
    Select Case nHeading
        Case E_Heading.NORTH
            Y = -1
        
        Case E_Heading.EAST
            X = 1
        
        Case E_Heading.SOUTH
            Y = 1
        
        Case E_Heading.WEST
            X = -1
    End Select
    
    'Fill temp pos
    tX = UserPos.X + X
    tY = UserPos.Y + Y
    
    'Check to see if its out of bounds
    If tX < MinXBorder Or tX > MaxXBorder Or tY < MinYBorder Or tY > MaxYBorder Then
        Exit Sub
    Else
        'Start moving... MainLoop does the rest
        AddtoUserPos.X = X
        UserPos.X = tX
        AddtoUserPos.Y = Y
        UserPos.Y = tY
        UserMoving = 1
        
        If Settings.MiniMap Then Call General_Pixel_Map_Set_Area
        
        bTecho = IIf(MapData(UserPos.X, UserPos.Y).Trigger = 1 Or _
                MapData(UserPos.X, UserPos.Y).Trigger = 2 Or _
                MapData(UserPos.X, UserPos.Y).Trigger = 6 Or _
                MapData(UserPos.X, UserPos.Y).Trigger = 7 Or _
                MapData(UserPos.X, UserPos.Y).Trigger = 4, True, False)
    End If
End Sub

Private Function HayFogata(ByRef location As Position) As Boolean
    Dim J As Long
    Dim K As Long
    
    For J = UserPos.X - 8 To UserPos.X + 8
        For K = UserPos.Y - 6 To UserPos.Y + 6
            If InMapBounds(J, K) Then
                If MapData(J, K).ObjGrh.GrhIndex = GrhFogata Then
                    location.X = J
                    location.Y = K
                    
                    HayFogata = True
                    Exit Function
                End If
            End If
        Next K
    Next J
End Function

Function NextOpenChar() As Integer
'*****************************************************************
'Finds next open char slot in CharList
'*****************************************************************
    Dim LoopC As Long
    Dim Dale As Boolean
    
    LoopC = 1
    Do While CharList(LoopC).Active And Dale
        LoopC = LoopC + 1
        Dale = (LoopC <= UBound(CharList))
    Loop
    
    NextOpenChar = LoopC
End Function


Function LegalPos(ByVal X As Integer, ByVal Y As Integer) As Boolean
'*****************************************************************
'Checks to see if a tile position is legal
'*****************************************************************
    'Limites del mapa
    If X < MinXBorder Or X > MaxXBorder Or Y < MinYBorder Or Y > MaxYBorder Then
        Exit Function
    End If
    
    'Tile Bloqueado?
    If MapData(X, Y).Blocked = 1 Then
        Exit Function
    End If
    
    '¿Hay un personaje?
    If MapData(X, Y).CharIndex > 0 Then
        Exit Function
    End If
   
    If UserNavegando <> TileEngine_Is_Water(X, Y) Then
        Exit Function
    End If
    
    LegalPos = True
End Function

Function MoveToLegalPos(ByVal X As Integer, ByVal Y As Integer) As Boolean
'*****************************************************************
'Author: ZaMa
'Last Modify Date: 01/08/2009
'Checks to see if a tile position is legal, including if there is a casper in the tile
'10/05/2009: ZaMa - Now you can't change position with a casper which is in the shore.
'01/08/2009: ZaMa - Now invisible admins can't change position with caspers.
'*****************************************************************
    Dim CharIndex As Integer
    
    'Limites del mapa
    If X < MinXBorder Or X > MaxXBorder Or Y < MinYBorder Or Y > MaxYBorder Then
        Exit Function
    End If
    
    'Tile Bloqueado?
    If MapData(X, Y).Blocked = 1 Then
        Exit Function
    End If
    
    CharIndex = MapData(X, Y).CharIndex
    '¿Hay un personaje?
    If CharIndex > 0 Then
    
        If MapData(UserPos.X, UserPos.Y).Blocked = 1 Then
            Exit Function
        End If
        
        With CharList(CharIndex)
            ' Si no es casper, no puede pasar
            If .iHead <> CASPER_HEAD And .iBody <> FRAGATA_FANTASMAL Then
                Exit Function
            Else
                ' No puedo intercambiar con un casper que este en la orilla (Lado tierra)
                If TileEngine_Is_Water(UserPos.X, UserPos.Y) Then
                    If Not TileEngine_Is_Water(X, Y) Then Exit Function
                Else
                    ' No puedo intercambiar con un casper que este en la orilla (Lado agua)
                    If TileEngine_Is_Water(X, Y) Then Exit Function
                End If
                
                ' Los admins no pueden intercambiar pos con caspers cuando estan invisibles
                If CharList(UserCharIndex).priv > 0 And CharList(UserCharIndex).priv < 6 Then
                    If CharList(UserCharIndex).invisible = True Then Exit Function
                End If
            End If
        End With
    End If
   
    If UserNavegando <> TileEngine_Is_Water(X, Y) Then
        Exit Function
    End If
    
    MoveToLegalPos = True
End Function

Function InMapBounds(ByVal X As Integer, ByVal Y As Integer) As Boolean
'*****************************************************************
'Checks to see if a tile position is in the maps bounds
'*****************************************************************
    If X < XMinMapSize Or X > XMaxMapSize Or Y < YMinMapSize Or Y > YMaxMapSize Then
        Exit Function
    End If
    
    InMapBounds = True
End Function

Sub RenderScreen(ByVal TileX As Integer, ByVal TileY As Integer, ByVal PixelOffsetX As Integer, ByVal PixelOffsetY As Integer)
    Dim Y           As Long     'Keeps track of where on map we are
    Dim X           As Long     'Keeps track of where on map we are
    Dim screenminY  As Integer  'Start Y pos on current screen
    Dim screenmaxY  As Integer  'End Y pos on current screen
    Dim screenminX  As Integer  'Start X pos on current screen
    Dim screenmaxX  As Integer  'End X pos on current screen
    Dim MinY        As Integer  'Start Y pos on current map
    Dim MaxY        As Integer  'End Y pos on current map
    Dim MinX        As Integer  'Start X pos on current map
    Dim MaxX        As Integer  'End X pos on current map
    Dim screenX     As Integer  'Keeps track of where to place tile on screen
    Dim screenY     As Integer  'Keeps track of where to place tile on screen
    Dim minXOffset  As Integer
    Dim minYOffset  As Integer
    Dim PixelOffsetXTemp As Integer 'For centering grhs
    Dim PixelOffsetYTemp As Integer 'For centering grhs
    Dim ElapsedTime As Single

    ElapsedTime = Engine_ElapsedTime()
    
    'Figure out Ends and Starts of screen
    screenminY = TileY - HalfWindowTileHeight
    screenmaxY = TileY + HalfWindowTileHeight
    screenminX = TileX - HalfWindowTileWidth
    screenmaxX = TileX + HalfWindowTileWidth
    
    MinY = screenminY - Engine_Get_TileBuffer
    MaxY = screenmaxY + Engine_Get_TileBuffer
    MinX = screenminX - Engine_Get_TileBuffer
    MaxX = screenmaxX + Engine_Get_TileBuffer
    
    'Make sure mins and maxs are allways in map bounds
    If MinY < XMinMapSize Then
        minYOffset = YMinMapSize - MinY
        MinY = YMinMapSize
    End If
    
    If MaxY > YMaxMapSize Then MaxY = YMaxMapSize
    
    If MinX < XMinMapSize Then
        minXOffset = XMinMapSize - MinX
        MinX = XMinMapSize
    End If
    
    If MaxX > XMaxMapSize Then MaxX = XMaxMapSize
    
    'If we can, we render around the view area to make it smoother
    If screenminY > YMinMapSize Then
        screenminY = screenminY - 1
    Else
        screenminY = 1
        screenY = 1
    End If
    
    If screenmaxY < YMaxMapSize Then screenmaxY = screenmaxY + 1
    
    If screenminX > XMinMapSize Then
        screenminX = screenminX - 1
    Else
        screenminX = 1
        screenX = 1
    End If
    
    If screenmaxX < XMaxMapSize Then screenmaxX = screenmaxX + 1
    
    ParticleOffsetX = (Engine_PixelPosX(screenminX) - PixelOffsetX)
    ParticleOffsetY = (Engine_PixelPosY(screenminY) - PixelOffsetY)
    
    
    For Y = screenminY To screenmaxY
        For X = screenminX To screenmaxX
            'Layer 1
            Call Engine_Render_Layer1(X, Y, screenX, screenY, PixelOffsetX, PixelOffsetY)

            'Layer 2
            If MapData(X, Y).Graphic(2).GrhIndex <> 0 Then
                Call TileEngine_Render_GRH(MapData(X, Y).Graphic(2), _
                    (screenX - 1) * TilePixelWidth + PixelOffsetX, _
                    (screenY - 1) * TilePixelHeight + PixelOffsetY, _
                    0, 0, X, Y, MapData(X, Y).Engine_Light(), 0, 0, False)
            End If
                
            'Walk Effect
            If Settings.Walk_Effect Then
                If MapData(X, Y).Effect <> 0 Then
                    Call TileEngine_Render_GrhIndex(DX8_Effects(MapData(X, Y).Effect).File, _
                        (screenX - 1) * TilePixelWidth + PixelOffsetX, _
                        (screenY - 1) * TilePixelHeight + PixelOffsetY, 0, DX8_Effects(MapData(X, Y).Effect).Color(), 0, 0, False)
                        DX8_Effects_Update MapData(X, Y).Effect
                End If
            End If
            screenX = screenX + 1
        Next X
        screenX = screenX - X + screenminX
        screenY = screenY + 1
    Next Y
    
    '<----- Layer Obj, Char, 3 ----->
    screenY = minYOffset - Engine_Get_TileBuffer
    For Y = MinY To MaxY
        screenX = minXOffset - Engine_Get_TileBuffer
        For X = MinX To MaxX
            PixelOffsetXTemp = screenX * TilePixelWidth + PixelOffsetX
            PixelOffsetYTemp = screenY * TilePixelHeight + PixelOffsetY

            
            With MapData(X, Y)
                'Object Layer
                If .ObjGrh.GrhIndex <> 0 Then
                    Call TileEngine_Render_GRH(.ObjGrh, PixelOffsetXTemp, PixelOffsetYTemp - MapData(X, Y).Vertex_Offset(0), 1, 1, X, Y, MapData(X, Y).Engine_Light(), 0, 0, False)
                End If
                    
                'Char layer
                If .CharIndex <> 0 Then
                    Call CharRender(.CharIndex, PixelOffsetXTemp, PixelOffsetYTemp)
                End If
                    
                'Layer 3
                If .Graphic(3).GrhIndex <> 0 Then
                    Call TileEngine_Render_GRH(.Graphic(3), PixelOffsetXTemp, PixelOffsetYTemp, 1, 1, X, Y, MapData(X, Y).Engine_Light(), 0, 0, False)
                End If
            End With
            
            screenX = screenX + 1
        Next X
        
        screenY = screenY + 1
    Next Y
    
    '   #### OBJNAMES
    If Settings.NombreItems Then
        If MapData(MouseTileX, MouseTileY).OBJInfo.name <> "" Then
            X = Engine_TPtoSPX(MouseTileX)
            Y = Engine_TPtoSPY(MouseTileY)
            
                Engine_Render_Fill_Box X, Y, Engine_GetTextWidth(1, MapData(MouseTileX, MouseTileY).OBJInfo.name) + 10, 13, ColorData.NegroAB(1)
                Fonts_Render_String MapData(MouseTileX, MouseTileY).OBJInfo.name, X + 5, Y
        End If
    End If
    
    
    '<----- Layer 4 ----->
        screenY = minYOffset - Engine_Get_TileBuffer
        For Y = MinY To MaxY
            screenX = minXOffset - Engine_Get_TileBuffer
            For X = MinX To MaxX
                'Layer 4
                If Not bTecho Then
                    If MapData(X, Y).Graphic(4).GrhIndex Then
                        Call TileEngine_Render_GRH(MapData(X, Y).Graphic(4), _
                            screenX * TilePixelWidth + PixelOffsetX, _
                            screenY * TilePixelHeight + PixelOffsetY, _
                            1, 0, X, Y, MapData(X, Y).Engine_Light(), 0, 0, False)
                    End If
                End If
                    
                If CurMapAmbient.MapBlocks(X, Y).Area <> 0 And Setting_Map_Areas Then
                    Engine_Render_Fill_Box screenX * TilePixelWidth + PixelOffsetX, screenY * TilePixelHeight + PixelOffsetY, 32, 32, Get_Area_Color(CurMapAmbient.MapBlocks(X, Y).Area)
                    Fonts_Render_String CurMapAmbient.MapBlocks(X, Y).Area, screenX * TilePixelWidth + PixelOffsetX, screenY * TilePixelHeight + PixelOffsetY, ColorData.BlancoAB(1), False, 2
                End If

                screenX = screenX + 1
            Next X
            screenY = screenY + 1
        Next Y

    'Weather Update & Render
    Call Engine_Weather_Update
    Call Effect_UpdateAll
    Call Engine_Water_Effect_Update
    
    If Settings.ProyectileEngine = True Then
        Dim J As Integer
        
        If LastProjectile > 0 Then
            For J = 1 To LastProjectile
                If ProjectileList(J).Grh.GrhIndex Then
                    Dim Angle As Single
                    'Update the position
                    Angle = DegreeToRadian * Engine_GetAngle(ProjectileList(J).X, ProjectileList(J).Y, ProjectileList(J).tX, ProjectileList(J).tY)
                    ProjectileList(J).X = ProjectileList(J).X + (Sin(Angle) * ElapsedTime * 0.63)
                    ProjectileList(J).Y = ProjectileList(J).Y - (Cos(Angle) * ElapsedTime * 0.63)
                    
                    'Update the rotation
                    'If ProjectileList(J).RotateSpeed > 0 Then
                    '    ProjectileList(J).Rotate = ProjectileList(J).Rotate + (ProjectileList(J).RotateSpeed * ElapsedTime * 0.01)
                    '    Do While ProjectileList(J).Rotate > 360
                    '        ProjectileList(J).Rotate = ProjectileList(J).Rotate - 360
                    '    Loop
                    'End If
    
                    'Draw if within range
                    X = ((-MinX - 1) * 32) + ProjectileList(J).X + PixelOffsetX + ((10 - Settings.BufferSize) * 32) - 288 + ProjectileList(J).OffsetX
                    Y = ((-MinY - 1) * 32) + ProjectileList(J).Y + PixelOffsetY + ((10 - Settings.BufferSize) * 32) - 288 + ProjectileList(J).OffsetY
                    
                    
                    If Y >= -32 Then
                        If Y <= (ScreenHeight + 32) Then
                            If X >= -32 Then
                                If X <= (ScreenWidth + 32) Then
                                    If ProjectileList(J).Rotate = 0 Then
                                        TileEngine_Render_GRH ProjectileList(J).Grh, X, Y, 1, 0, 50, 50, ColorData.Blanco
                                    Else
                                    
                                    
                                          If (ProjectileList(J).X > ProjectileList(J).tX) And (ProjectileList(J).Y > ProjectileList(J).tY) Then
                                            ProjectileList(J).Rotate = 23.5
                                          ElseIf (ProjectileList(J).X > ProjectileList(J).tX) And (ProjectileList(J).Y < ProjectileList(J).tY) Then
                                             ProjectileList(J).Rotate = 110
                                          ElseIf (ProjectileList(J).X < ProjectileList(J).tX) And (ProjectileList(J).Y < ProjectileList(J).tY) Then
                                              ProjectileList(J).Rotate = 1.3
                                          ElseIf (ProjectileList(J).X = ProjectileList(J).tX) And (ProjectileList(J).Y < ProjectileList(J).tY) Then
                                              ProjectileList(J).Rotate = 1.92
                                          ElseIf (ProjectileList(J).X = ProjectileList(J).tX) And (ProjectileList(J).Y > ProjectileList(J).tY) Then
                                              ProjectileList(J).Rotate = 200
                                          Else
                                                ProjectileList(J).Rotate = 0
                                          End If

                                        TileEngine_Render_GRH ProjectileList(J).Grh, X, Y, 1, 0, 50, 50, ColorData.Blanco, 0, ProjectileList(J).Rotate, False
                                    End If
                                End If
                            End If
                        End If
                    End If
                    
                End If
            Next J
            
            'Check if it is close enough to the target to remove
            For J = 1 To LastProjectile
                If ProjectileList(J).Grh.GrhIndex Then
                    If Abs(ProjectileList(J).X - ProjectileList(J).tX) < 20 Then
                        If Abs(ProjectileList(J).Y - ProjectileList(J).tY) < 20 Then
                            Engine_Projectile_Erase J
                        End If
                    End If
                End If
            Next J
            
        End If
    End If
    
    '   Set Offsets
    LastOffsetX = ParticleOffsetX
    LastOffsetY = ParticleOffsetY
    

    If Settings.PartyMembers Then Call Draw_Party_Members
    If DX_Count.DoIt Then Call RenderCount
    
    If AlphaPres <> 0 Then
        AlphaPres = AlphaPres - 1
        Call Engine_Render_Fill_Box(0, 0, frmMain.MainViewPic.Width, frmMain.MainViewPic.Height, D3DColorARGB(AlphaPres, 0, 0, 0))
    End If

End Sub

Sub RenderConnect()
Dim tmp(0 To 3) As Long

Dim i As Long
    For i = 1 To 12
        If Fade_GUIe(i) = False Then
            Fade_GUI(i) = Fade_GUI(i) + 1
        End If
        
        If Fade_GUI(i) = 240 Then Fade_GUIe(i) = True
        
        If Fade_GUIe(i) = True Then
            Fade_GUI(i) = Fade_GUI(i) - 1
        End If
        
        If Fade_GUI(i) = 0 Then Fade_GUIe(i) = False
        
    Next i
        
    

    Engine_BeginScene

    ' #### Background
    TileEngine_Render_GrhIndex 31773, 0, 0, 0, ColorData.Blanco()

    Engine_Long_To_RGB_List tmp(), D3DColorARGB(Fade_GUI(1), 255, 255, 255)
    TileEngine_Render_GrhIndex 23673, 241, 406, 0, tmp()
    Engine_Long_To_RGB_List tmp(), D3DColorARGB(Fade_GUI(2), 255, 255, 255)
    TileEngine_Render_GrhIndex 23674, 241, 419, 0, tmp()
    Engine_Long_To_RGB_List tmp(), D3DColorARGB(Fade_GUI(3), 255, 255, 255)
    TileEngine_Render_GrhIndex 23675, 241, 444, 0, tmp()
    Engine_Long_To_RGB_List tmp(), D3DColorARGB(Fade_GUI(4), 255, 255, 255)
    TileEngine_Render_GrhIndex 23676, 349, 406, 0, tmp()
    Engine_Long_To_RGB_List tmp(), D3DColorARGB(Fade_GUI(5), 255, 255, 255)
    TileEngine_Render_GrhIndex 23677, 339, 426, 0, tmp()
    Engine_Long_To_RGB_List tmp(), D3DColorARGB(Fade_GUI(6), 255, 255, 255)
    TileEngine_Render_GrhIndex 23678, 327, 447, 0, tmp()
    Engine_Long_To_RGB_List tmp(), D3DColorARGB(Fade_GUI(7), 255, 255, 255)
    TileEngine_Render_GrhIndex 23679, 323, 475, 0, tmp()
    Engine_Long_To_RGB_List tmp(), D3DColorARGB(Fade_GUI(8), 255, 255, 255)
    TileEngine_Render_GrhIndex 23680, 318, 499, 0, tmp()
    Engine_Long_To_RGB_List tmp(), D3DColorARGB(Fade_GUI(9), 255, 255, 255)
    TileEngine_Render_GrhIndex 23681, 465, 447, 0, tmp()
    Engine_Long_To_RGB_List tmp(), D3DColorARGB(Fade_GUI(10), 255, 255, 255)
    TileEngine_Render_GrhIndex 23682, 468, 475, 0, tmp()
    Engine_Long_To_RGB_List tmp(), D3DColorARGB(Fade_GUI(11), 255, 255, 255)
    TileEngine_Render_GrhIndex 23683, 468, 505, 0, tmp()
    Engine_Long_To_RGB_List tmp(), D3DColorARGB(Fade_GUI(12), 255, 255, 255)
    TileEngine_Render_GrhIndex 23684, 520, 412, 0, tmp()
    
    
    ' #### Connect GUI
    TileEngine_Render_GrhIndex 23669, 0, 0, 0, ColorData.Blanco()
    If Settings.Recordar Then TileEngine_Render_GrhIndex 23672, 342, 397, 0, ColorData.Blanco()
    
    
    ' #### Render FPS
    Engine_Render_FPS
    
    Fonts_Render_String frmConnect.tmpName, 400, 276, ColorData.Blanco(1), True, 1
    Fonts_Render_String frmConnect.tmpPassF, 400, 331, ColorData.Blanco(1), True, 1
    
    
    Static TextBoxMode As Long
    Static ShowTextBoxMode As Boolean
    
    If GetTickCount - TextBoxMode > 500 Then
        
        ShowTextBoxMode = Not ShowTextBoxMode
        TextBoxMode = GetTickCount
    End If
    
        If frmConnect.Focus = 1 Then
            If ShowTextBoxMode Then Fonts_Render_String "|", 400 + (Engine_GetTextWidth(1, frmConnect.tmpName) / 2), 276, ColorData.Blanco(1), False, 1
        End If
    
        If frmConnect.Focus = 2 Then
            If ShowTextBoxMode Then Fonts_Render_String "|", 400 + (Engine_GetTextWidth(1, frmConnect.tmpPassF) / 2), 328, ColorData.Blanco(1), False, 1
        End If
    
    Engine_EndScene ConnectScreenRect, frmConnect.hWnd
    
    timerElapsedTime = GetElapsedTime()
End Sub

Public Function RenderSounds()
'**************************************************************
'Author: Juan Martín Sotuyo Dodero
'Last Modify Date: 3/30/2008
'Actualiza todos los sonidos del mapa.
'**************************************************************
        If bRain And CurMapAmbient.Rain = True Then
        
            'Standelf
            If General_Get_Random_Number(1, 20000) < 10 Then
                Call Audio.PlayWave("105.wav", UserPos.X - 5, UserPos.Y)
                Start_Rampage
                OnRampage = GetTickCount
                OnRampageImg = OnRampage
                OnRampageImgGrh = 24653
            End If
            
            If OnRampageImg <> 0 Then
                If GetTickCount - OnRampageImg > 36 Then
                
                    OnRampageImgGrh = OnRampageImgGrh + 1
                    If OnRampageImgGrh = 24664 Then OnRampageImgGrh = 24653
        
                    OnRampageImg = GetTickCount
                End If
            End If
            
            If OnRampage <> 0 Then 'Hay Uno en curso
                If GetTickCount - OnRampage > 400 Then
                    End_Rampage
                    OnRampage = 0
                End If
            End If
        
            If bTecho Then
                If frmMain.IsPlaying <> PlayLoop.plLluviain Then
                    If RainBufferIndex Then _
                        Call Audio.StopWave(RainBufferIndex)
                    RainBufferIndex = Audio.PlayWave("lluviain.wav", 0, 0, LoopStyle.Enabled)
                    frmMain.IsPlaying = PlayLoop.plLluviain
                End If
            Else
                If frmMain.IsPlaying <> PlayLoop.plLluviaout Then
                    If RainBufferIndex Then _
                        Call Audio.StopWave(RainBufferIndex)
                    RainBufferIndex = Audio.PlayWave("lluviaout.wav", 0, 0, LoopStyle.Enabled)
                    frmMain.IsPlaying = PlayLoop.plLluviaout
                End If
            End If
        End If
    
    DoFogataFx
End Function

Function HayUserAbajo(ByVal X As Integer, ByVal Y As Integer, ByVal GrhIndex As Integer) As Boolean
    If GrhIndex > 0 Then
        HayUserAbajo = _
            CharList(UserCharIndex).Pos.X >= X - (GrhData(GrhIndex).TileWidth \ 2) _
                And CharList(UserCharIndex).Pos.X <= X + (GrhData(GrhIndex).TileWidth \ 2) _
                And CharList(UserCharIndex).Pos.Y >= Y - (GrhData(GrhIndex).TileHeight - 1) _
                And CharList(UserCharIndex).Pos.Y <= Y
    End If
End Function

Public Function InitTileEngine(ByVal setDisplayFormhWnd As Long, ByVal setTilePixelHeight As Integer, ByVal setTilePixelWidth As Integer, ByVal pixelsToScrollPerFrameX As Integer, pixelsToScrollPerFrameY As Integer) As Boolean
'***************************************************
'Author: Aaron Perkins
'Last Modification: 08/14/07
'Last modified by: Juan Martín Sotuyo Dodero (Maraxus)
'Configures the engine to start running.
'***************************************************
    TilePixelWidth = setTilePixelWidth
    TilePixelHeight = setTilePixelHeight
    WindowTileHeight = Round(frmMain.MainViewPic.Height / 32, 0)
    WindowTileWidth = Round(frmMain.MainViewPic.Width / 32, 0)
    
    HalfWindowTileHeight = WindowTileHeight \ 2
    HalfWindowTileWidth = WindowTileWidth \ 2

    MinXBorder = XMinMapSize + (WindowTileWidth \ 2)
    MaxXBorder = XMaxMapSize - (WindowTileWidth \ 2)
    MinYBorder = YMinMapSize + (WindowTileHeight \ 2)
    MaxYBorder = YMaxMapSize - (WindowTileHeight \ 2)
    
    'Resize mapdata array
    ReDim MapData(XMinMapSize To XMaxMapSize, YMinMapSize To YMaxMapSize) As MapBlock
    
    'Set intial user position
    UserPos.X = MinXBorder
    UserPos.Y = MinYBorder
    
    'Set scroll pixels per frame
    ScrollPixelsPerFrameX = pixelsToScrollPerFrameX
    ScrollPixelsPerFrameY = pixelsToScrollPerFrameY

On Error GoTo 0

    Call LoadGrhData
    Call CargarCuerpos
    Call CargarCabezas
    Call CargarCascos
    Call CargarFxs
    Call Load_Auras
    Call CargarAnimArmas
    Call CargarAnimEscudos
    Call CargarColores
    
    
    'Index Buffer. Dunkan
    indexList(0) = 0: indexList(1) = 1: indexList(2) = 2
    indexList(3) = 3: indexList(4) = 4: indexList(5) = 5
    
    Set ibQuad = DirectDevice.CreateIndexBuffer(Len(indexList(0)) * 4, 0, D3DFMT_INDEX16, D3DPOOL_MANAGED)
    D3DIndexBuffer8SetData ibQuad, 0, Len(indexList(0)) * 4, 0, indexList(0)
    
    Set vbQuadIdx = DirectDevice.CreateVertexBuffer(Len(Temp_Verts(0)) * 4, 0, D3DFVF_XYZ Or D3DFVF_DIFFUSE Or D3DFVF_SPECULAR Or D3DFVF_TEX1, D3DPOOL_MANAGED)
    
    InitTileEngine = True
End Function

Sub ShowNextFrame(ByVal DisplayFormTop As Integer, ByVal DisplayFormLeft As Integer, ByVal MouseViewX As Integer, ByVal MouseViewY As Integer)

    If EngineRun Then
        If UserMoving Then
            '****** Move screen Left and Right if needed ******
            If AddtoUserPos.X <> 0 Then
                OffsetCounterX = OffsetCounterX - ScrollPixelsPerFrameX * AddtoUserPos.X * timerTicksPerFrame
                If Abs(OffsetCounterX) >= Abs(TilePixelWidth * AddtoUserPos.X) Then
                    OffsetCounterX = 0
                    AddtoUserPos.X = 0
                    UserMoving = False
                End If
            End If
            
            '****** Move screen Up and Down if needed ******
            If AddtoUserPos.Y <> 0 Then
                OffsetCounterY = OffsetCounterY - ScrollPixelsPerFrameY * AddtoUserPos.Y * timerTicksPerFrame
                If Abs(OffsetCounterY) >= Abs(TilePixelHeight * AddtoUserPos.Y) Then
                    OffsetCounterY = 0
                    AddtoUserPos.Y = 0
                    UserMoving = False
                End If
            End If
        End If
        
        'Update mouse position within view area
        Call ConvertCPtoTP(MouseViewX, MouseViewY, MouseTileX, MouseTileY)

        '****** Update screen ******
        If UserCiego Then
            DirectDevice.Clear 0, ByVal 0, D3DCLEAR_TARGET, 0, 1#, 0
        Else
            Call RenderScreen(UserPos.X - AddtoUserPos.X, UserPos.Y - AddtoUserPos.Y, OffsetCounterX, OffsetCounterY)
        End If
        
        Call Dialogos.Render
        Call DibujarCartel
        Call DialogosClanes.Draw
        
        'Get timing info
        timerElapsedTime = GetElapsedTime()
        timerTicksPerFrame = timerElapsedTime * Engine_Get_BaseSpeed
    End If

End Sub

Public Function GetElapsedTime() As Single
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 10/07/2002
'Gets the time that past since the last call
'**************************************************************
    Dim Start_Time As Currency
    Static end_time As Currency
    Static timer_freq As Currency

    'Get the timer frequency
    If timer_freq = 0 Then
        QueryPerformanceFrequency timer_freq
    End If
    
    'Get current time
    Call QueryPerformanceCounter(Start_Time)
    
    'Calculate elapsed time
    GetElapsedTime = (Start_Time - end_time) / timer_freq * 1000
    
    'Get next end time
    Call QueryPerformanceCounter(end_time)
End Function

Private Sub CharRender(ByVal CharIndex As Long, ByVal PixelOffsetX As Integer, ByVal PixelOffsetY As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 12/03/04
'Draw char's to screen without offcentering them
'***************************************************
    Dim Moved As Boolean
    Dim MouseOverChar As Boolean
    Dim Offset As Integer
        
    With CharList(CharIndex)
        If .Moving Then
            'If needed, move left and right
            If .scrollDirectionX <> 0 Then
                .MoveOffsetX = .MoveOffsetX + ScrollPixelsPerFrameX * Sgn(.scrollDirectionX) * timerTicksPerFrame
                
                'Start animations
'TODO : Este parche es para evita los uncornos exploten al moverse!! REVER!!!
                If .Body.Walk(.Heading).Speed > 0 Then _
                    .Body.Walk(.Heading).Started = 1
                .Arma.WeaponWalk(.Heading).Started = 1
                .Escudo.ShieldWalk(.Heading).Started = 1
                
                'Char moved
                Moved = True
                
                'Check if we already got there
                If (Sgn(.scrollDirectionX) = 1 And .MoveOffsetX >= 0) Or _
                        (Sgn(.scrollDirectionX) = -1 And .MoveOffsetX <= 0) Then
                    .MoveOffsetX = 0
                    .scrollDirectionX = 0
                End If
            End If
            
            'If needed, move up and down
            If .scrollDirectionY <> 0 Then
                .MoveOffsetY = .MoveOffsetY + ScrollPixelsPerFrameY * Sgn(.scrollDirectionY) * timerTicksPerFrame
                
                'Start animations
'TODO : Este parche es para evita los uncornos exploten al moverse!! REVER!!!
                If .Body.Walk(.Heading).Speed > 0 Then _
                    .Body.Walk(.Heading).Started = 1
                .Arma.WeaponWalk(.Heading).Started = 1
                .Escudo.ShieldWalk(.Heading).Started = 1
                
                'Char moved
                Moved = True
                
                'Check if we already got there
                If (Sgn(.scrollDirectionY) = 1 And .MoveOffsetY >= 0) Or _
                        (Sgn(.scrollDirectionY) = -1 And .MoveOffsetY <= 0) Then
                    .MoveOffsetY = 0
                    .scrollDirectionY = 0
                End If
            End If
        End If
        
        'If done moving stop animation
        If Not Moved Then
            'Stop animations
            .Body.Walk(.Heading).Started = 0
            .Body.Walk(.Heading).FrameCounter = 1
         
            If IsAttacking = False Then
                .Arma.WeaponWalk(.Heading).Started = 0
                .Arma.WeaponWalk(.Heading).FrameCounter = 1
            End If
         
            .Escudo.ShieldWalk(.Heading).Started = 0
            .Escudo.ShieldWalk(.Heading).FrameCounter = 1
         
            .Moving = False
        End If
        
        PixelOffsetX = PixelOffsetX + .MoveOffsetX
        PixelOffsetY = PixelOffsetY + .MoveOffsetY
        
        ' #### Mountain Offset
        Offset = MapData(.Pos.X, .Pos.Y).Vertex_Offset(0)
        
        If TileEngine_Is_Water(.Pos.X, .Pos.Y) Then Offset = polygonCount(0)

        If .Vertex_Offset > Offset Then
            .Vertex_Offset = .Vertex_Offset - 1
        ElseIf .Vertex_Offset < Offset Then
            .Vertex_Offset = .Vertex_Offset + 1
        End If
            
        PixelOffsetY = PixelOffsetY - .Vertex_Offset
        '  #### Mountain Offset

        ' #### Check If Mouse is Over Char
        If Settings.TonalidadPJ Then
            If Abs(MouseTileX - .Pos.X) < 1 And (Abs(MouseTileY - .Pos.Y)) < 1 And CharIndex <> UserCharIndex Then _
                MouseOverChar = True
        End If
        ' #### Is Mouse Over Char?
        
        
        ' #### Render Char
        If .Head.Head(.Heading).GrhIndex Then
            If Not .invisible Then
            
            Engine_Movement_Speed = 0.5
                
                ' #### Render Shadow
                If Settings.UsarSombras = True And .invisible = False And TileEngine_Is_Water(.Pos.X, .Pos.Y) = False Then
                    Dim DegreesShadow As Single
                    DegreesShadow = Engine_Get_2_Points_Angle(PixelOffsetX, -PixelOffsetY, frmMain.MouseX, -frmMain.MouseY)
                    
                    If .Body.Walk(.Heading).GrhIndex Then
                        Call TileEngine_Render_GRH(.Body.Walk(.Heading), PixelOffsetX, PixelOffsetY, 1, 1, .Pos.X, .Pos.Y, ColorData.SombraAB, 0, DegreesShadow, True)
                    End If
                End If
            
                ' #### Render Wills
                Call Render_Auras(CharIndex, PixelOffsetX, PixelOffsetY)
                
                    If Settings.TonalidadPJ And MouseOverChar Then
                        ' #### Render Body
                        If .Body.Walk(.Heading).GrhIndex Then _
                            Call TileEngine_Render_GRH(.Body.Walk(.Heading), PixelOffsetX, PixelOffsetY, 1, 1, .Pos.X, .Pos.Y, ColorData.Verde, 0, 0, False)
                        
                        ' #### Render  Head
                        If .Head.Head(.Heading).GrhIndex Then
                            Call TileEngine_Render_GRH(.Head.Head(.Heading), PixelOffsetX + .Body.HeadOffset.X, PixelOffsetY + .Body.HeadOffset.Y, 1, 1, .Pos.X, .Pos.Y, ColorData.Verde)
                            
                            ' #### Render  Helmet
                            If .Casco.Head(.Heading).GrhIndex Then _
                                Call TileEngine_Render_GRH(.Casco.Head(.Heading), PixelOffsetX + .Body.HeadOffset.X, PixelOffsetY + .Body.HeadOffset.Y + OFFSET_HEAD, 1, 1, .Pos.X, .Pos.Y, ColorData.Verde)
                                
                            ' #### Render  Weapon
                            If .Arma.WeaponWalk(.Heading).GrhIndex Then _
                                Call TileEngine_Render_GRH(.Arma.WeaponWalk(.Heading), PixelOffsetX, PixelOffsetY, 1, 1, .Pos.X, .Pos.Y, ColorData.Verde)
                                
                            ' #### Render  Shield
                            If .Escudo.ShieldWalk(.Heading).GrhIndex Then _
                                Call TileEngine_Render_GRH(.Escudo.ShieldWalk(.Heading), PixelOffsetX, PixelOffsetY, 1, 1, .Pos.X, .Pos.Y, ColorData.Verde)
                        End If
                    Else
                        ' #### Render Body
                        If .Body.Walk(.Heading).GrhIndex Then _
                            Call TileEngine_Render_GRH(.Body.Walk(.Heading), PixelOffsetX, PixelOffsetY, 1, 1, .Pos.X, .Pos.Y, MapData(.Pos.X, .Pos.Y).Engine_Light(), 0, 0, False)
                        
                        ' #### Render  Head
                        If .Head.Head(.Heading).GrhIndex Then
                            Call TileEngine_Render_GRH(.Head.Head(.Heading), PixelOffsetX + .Body.HeadOffset.X, PixelOffsetY + .Body.HeadOffset.Y, 1, 1, .Pos.X, .Pos.Y, MapData(.Pos.X, .Pos.Y).Engine_Light())
                            
                            ' #### Render  Helmet
                            If .Casco.Head(.Heading).GrhIndex Then _
                                Call TileEngine_Render_GRH(.Casco.Head(.Heading), PixelOffsetX + .Body.HeadOffset.X, PixelOffsetY + .Body.HeadOffset.Y + OFFSET_HEAD, 1, 1, .Pos.X, .Pos.Y, MapData(.Pos.X, .Pos.Y).Engine_Light())
                                
                            ' #### Render  Weapon
                            If .Arma.WeaponWalk(.Heading).GrhIndex Then _
                                Call TileEngine_Render_GRH(.Arma.WeaponWalk(.Heading), PixelOffsetX, PixelOffsetY, 1, 1, .Pos.X, .Pos.Y, MapData(.Pos.X, .Pos.Y).Engine_Light())
                                
                            ' #### Render  Shield
                            If .Escudo.ShieldWalk(.Heading).GrhIndex Then _
                                Call TileEngine_Render_GRH(.Escudo.ShieldWalk(.Heading), PixelOffsetX, PixelOffsetY, 1, 1, .Pos.X, .Pos.Y, MapData(.Pos.X, .Pos.Y).Engine_Light())
                        End If
                    End If
                    
                    
                ' #### REFLECTION IN WATER
                If Settings.Reflect_Effect Then
                    If TileEngine_Is_Water(.Pos.X, .Pos.Y + 1) Then
                        ' #### Render Body
                        If .Body.Walk(.Heading).GrhIndex Then
                            Call TileEngine_Render_GRH(.Body.Walk(.Heading), PixelOffsetX, PixelOffsetY + 40, 1, 1, .Pos.X, .Pos.Y, ColorData.ReflectionBody(), 0, 0, , , IIf(.Heading = WEST Or .Heading = EAST, False, True), True)
                        End If
                        
                        Dim OffsetY As Integer, EsEnano As Boolean
                        EsEnano = .iHead >= ENANO_H_PRIMER_CABEZA And .iHead <= ENANO_H_ULTIMA_CABEZA
                        
                        OffsetY = GrhData(.Body.Walk(.Heading).GrhIndex).pixelHeight + IIf(EsEnano = False, 8, -8)
                        If .Head.Head(.Heading).GrhIndex Then
                            Call TileEngine_Render_GRH(.Head.Head(.Heading), PixelOffsetX + .Body.HeadOffset.X, PixelOffsetY + .Body.HeadOffset.Y + OffsetY, 1, 1, .Pos.X, .Pos.Y, ColorData.ReflectionHead(), , , , , , True)
                        End If
                        
                    End If
                End If
                
                    'Draw name over head
                    If LenB(.Nombre) > 0 Then
                        If Nombres = 1 And (General_Get_GM(UserCharIndex) Or Abs(MouseTileX - .Pos.X) < 2 And (Abs(MouseTileY - .Pos.Y)) < 2) Then
                             Call RenderName(CharIndex, PixelOffsetX, PixelOffsetY)
                        ElseIf Nombres = 2 Then
                             Call RenderName(CharIndex, PixelOffsetX, PixelOffsetY)
                        End If
                    End If
            End If
        Else

            If .Body.Walk(.Heading).GrhIndex Then
                Call TileEngine_Render_GRH(.Body.Walk(.Heading), PixelOffsetX, PixelOffsetY, 1, 1, .Pos.X, .Pos.Y, MapData(.Pos.X, .Pos.Y).Engine_Light())
            End If
            
            ' #### REFLECTION IN WATER
            If Settings.Reflect_Effect Then
                If TileEngine_Is_Water(.Pos.X, .Pos.Y + 1) Then
                    If .Body.Walk(.Heading).GrhIndex Then
                        Call TileEngine_Render_GRH(.Body.Walk(.Heading), PixelOffsetX, PixelOffsetY + 40, 1, 1, .Pos.X, .Pos.Y, ColorData.Reflection(), 0, 0, , , IIf(.Heading = WEST Or .Heading = EAST, False, True), True)
                    End If
                End If
            End If
        End If
        
        'Update dialogs
        If Settings.Dialog_Align = 1 Then
            Call Dialogos.UpdateDialogPos(PixelOffsetX + .Body.HeadOffset.X - 10, PixelOffsetY + .Body.HeadOffset.Y + OFFSET_HEAD, CharIndex) '34 son los pixeles del grh de la cabeza que quedan superpuestos al cuerpo
        Else
            Call Dialogos.UpdateDialogPos(PixelOffsetX + 16, PixelOffsetY + .Body.HeadOffset.Y + OFFSET_HEAD, CharIndex) '34 son los pixeles del grh de la cabeza que quedan superpuestos al cuerpo
        End If
        
        Engine_Movement_Speed = 1
        
        'Draw FX
        If .FxIndex <> 0 Then
            Call TileEngine_Render_GRH(.fX, PixelOffsetX + FxData(.FxIndex).OffsetX, PixelOffsetY + FxData(.FxIndex).OffsetY, 1, 1, .Pos.X, .Pos.Y, MapData(.Pos.X, .Pos.Y).Engine_Light(), Blit_Alpha.Blendop_Sustrative)
            
            'Check if animation is over
            If .fX.Started = 0 Then _
                .FxIndex = 0
        End If
    
    If MouseOverChar And SpellGrhIndex <> 0 Then
        TileEngine_Render_GrhIndex SpellGrhIndex, PixelOffsetX + 16, PixelOffsetY + 16, 1, ColorData.Blanco()
    End If
    
    
    ' #### LIFE BAR
    
        If .minHP <> 0 And .maxHP <> 0 And .ShowBar <> 0 Then
            .ShowBar = .ShowBar - 1
            Fonts_Render_String .tmpName, PixelOffsetX + 16, PixelOffsetY - (35 + 12), ColorData.BlancoAB(1), True
            Engine_Render_Fill_Box PixelOffsetX - 13, PixelOffsetY - (35 + 1), 62, 6, ColorData.NegroAB(1)
            Engine_Render_Fill_Box PixelOffsetX - 12, PixelOffsetY - (35), (((.minHP / 100) / (.maxHP / 100)) * 60), 4, ColorData.RojoAB(1)
        End If
    End With
End Sub

Private Sub RenderName(ByVal CharIndex As Long, ByVal X As Integer, ByVal Y As Integer)
    Dim Pos As Integer
    Dim line As String
    Dim Color As Long
   
    With CharList(CharIndex)
            Pos = General_Get_TagPosition(.Nombre)
    
            If .priv = 0 Then
                If .Atacable Then
                    Color = ColoresPJ(48)
                Else
                    If .muerto Then
                        Color = D3DColorARGB(255, 220, 220, 255)
                    Else
                        If .Criminal Then
                            Color = ColoresPJ(50)
                        Else
                            Color = ColoresPJ(49)
                        End If
                    End If
                End If
            Else
                Color = ColoresPJ(.priv)
            End If
    
            'Nick
            line = Left$(.Nombre, Pos - 2)
            Fonts_Render_String line, (X + 16), Y + 30, Color, True
            
            'Clan
            line = mid$(.Nombre, Pos)
            Fonts_Render_String line, (X + 16), Y + 40, D3DColorXRGB(255, 230, 130), True
              
    End With
End Sub
 
Public Sub SetCharacterFx(ByVal CharIndex As Integer, ByVal fX As Integer, ByVal Loops As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 12/03/04
'Sets an FX to the character.
'***************************************************
On Error GoTo handle
    
    With CharList(CharIndex)
        .FxIndex = fX
        
        If .FxIndex > 0 Then
            Call InitGrh(.fX, FxData(fX).Animacion)
        
            .fX.Loops = Loops
        End If
    End With
    
handle:
    Exit Sub
End Sub





Public Sub RenderItem(ByVal Pic As PictureBox, ByVal GrhIndex As Long)
    With GrhData(GrhIndex)
        Pic.Cls
        Pic.Picture = LoadPicture(Resources.Graphics & .FileNum & ".bmp")
    End With
End Sub

Public Sub TileEngine_Render_GRH(ByRef Grh As Grh, ByVal X As Integer, ByVal Y As Integer, ByVal Center As Byte, ByVal Animate As Byte, ByVal PosX As Byte, ByVal PosY As Byte, ByRef Color() As Long, Optional ByVal Alpha As Byte = 0, Optional ByVal Angle As Single = 0, Optional ByVal Shadow As Boolean = False, Optional ByVal ReSize As Integer = 0, Optional ByVal InvertX As Boolean = False, Optional ByVal InvertY As Boolean = False)
    Dim CurrentGrhIndex As Integer
    Dim SourceRect As RECT
    
    If Grh.GrhIndex = 0 Then Exit Sub
On Error GoTo Error
        
    If Animate Then
        If Grh.Started = 1 Then
            Grh.FrameCounter = Grh.FrameCounter + (timerElapsedTime * GrhData(Grh.GrhIndex).NumFrames / Grh.Speed) * Engine_Movement_Speed
            If Grh.FrameCounter > GrhData(Grh.GrhIndex).NumFrames Then
                Grh.FrameCounter = (Grh.FrameCounter Mod GrhData(Grh.GrhIndex).NumFrames) + 1
                
                If Grh.Loops <> INFINITE_LOOPS Then
                    If Grh.Loops > 0 Then
                        Grh.Loops = Grh.Loops - 1
                    Else
                        Grh.Started = 0
                    End If
                End If
            End If
        End If
    End If
    
    'Figure out what frame to draw (always 1 if not animated)
    CurrentGrhIndex = GrhData(Grh.GrhIndex).Frames(Grh.FrameCounter)
    
    With GrhData(CurrentGrhIndex)
        'Center Grh over X,Y pos
        If Center And ReSize = 0 Then
            If .TileWidth <> 1 Then
                X = X - Int(.TileWidth * TilePixelWidth / 2) + TilePixelWidth \ 2
            End If
            
            If .TileHeight <> 1 Then
                Y = Y - Int(.TileHeight * TilePixelHeight) + TilePixelHeight
            End If
        ElseIf Center And ReSize <> 0 Then
            If .TileWidth <> 1 Then
                X = X - (Int(.TileWidth * TilePixelWidth / 2) + TilePixelWidth \ 2) / ReSize + (TilePixelWidth / ReSize)
            End If
            
            If .TileHeight <> 1 Then
                Y = Y - (Int(.TileHeight * TilePixelHeight / 2) + TilePixelHeight \ 2) / ReSize
            End If
        End If
        
        SourceRect.Left = .sX
        SourceRect.Top = .sY
        SourceRect.Right = SourceRect.Left + .pixelWidth
        SourceRect.bottom = SourceRect.Top + .pixelHeight
        
        'Draw
        Call Device_Textured_Render(X, Y, Grh.GrhIndex, SurfaceDB.Surface(.FileNum), SourceRect, Color(), Alpha, Angle, Shadow, ReSize, InvertX, InvertY)
    End With
Exit Sub

Error:
    If Err.Number = 9 And Grh.FrameCounter < 1 Then
        Grh.FrameCounter = 1
        Resume
    Else
        Call Log_Engine("Error in TileEngine_Render_GRH, " & Err.Description & ", (" & Err.Number & ")")
        MsgBox "Error en el Engine Gráfico, Por favor contacte a los adminsitradores enviandoles el archivo Errors.Log que se encuentra el la carpeta del cliente.", vbCritical
        Call General_Close_Client
    End If
End Sub

Public Sub TileEngine_Render_GrhIndex(ByVal GrhIndex As Integer, ByVal X As Integer, ByVal Y As Integer, ByVal Center As Byte, ByRef Color_List() As Long, Optional ByVal Alpha As Byte = 0, Optional ByVal Angle As Single = 0, Optional ByVal Shadow As Boolean = False)
    Dim SourceRect As RECT
    
    With GrhData(GrhIndex)
        'Center Grh over X,Y pos
        If Center Then
            If .TileWidth <> 1 Then
                X = X - Int(.TileWidth * TilePixelWidth / 2) + TilePixelWidth \ 2
            End If
            
            If .TileHeight <> 1 Then
                Y = Y - Int(.TileHeight * TilePixelHeight) + TilePixelHeight
            End If
        End If
        
        SourceRect.Left = .sX
        SourceRect.Top = .sY
        SourceRect.Right = SourceRect.Left + .pixelWidth
        SourceRect.bottom = SourceRect.Top + .pixelHeight
        
        'Draw
        Call Device_Textured_Render(X, Y, GrhIndex, SurfaceDB.Surface(.FileNum), SourceRect, Color_List(), Alpha, Angle, Shadow)
    End With
End Sub


Public Sub RefreshAllChars()
'*****************************************************************
'Goes through the charlist and replots all the characters on the map
'Used to make sure everyone is visible
'*****************************************************************
    Dim LoopC As Long
    
    For LoopC = 1 To LastChar
        If CharList(LoopC).Active = 1 Then
            MapData(CharList(LoopC).Pos.X, CharList(LoopC).Pos.Y).CharIndex = LoopC
        End If
    Next LoopC
End Sub

Sub DrawGrhtoHdc(ByVal hWndDest As Long, ByVal GrhIndex As Integer, ByRef DestRect As RECT)
    Engine_BeginScene
    
        Call TileEngine_Render_GrhIndex(GrhIndex, 0, 0, 0, ColorData.Blanco())
        
    Engine_EndScene DestRect, hWndDest
End Sub























Public Sub Device_Textured_Render(ByVal X As Integer, ByVal Y As Integer, ByVal GrhIndex As Integer, ByVal Texture As Direct3DTexture8, ByRef src_rect As RECT, ByRef Color_List() As Long, Optional Alpha As Byte = 0, Optional ByVal Angle As Single = 0, Optional ByVal Shadow As Boolean = False, Optional ByVal ReSize As Integer = 0, Optional ByVal InvertX As Boolean = False, Optional ByVal InvertY As Boolean = False)

    If Alpha <> 0 Then
        Select Case Alpha
            Case Blit_Alpha.Blendop_Color '1
                DirectDevice.SetRenderState D3DRS_SRCBLEND, D3DBLEND_ONE
                DirectDevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_ONE
            Case Blit_Alpha.Blendop_Aditive '2
                DirectDevice.SetRenderState D3DRS_SRCBLEND, D3DBLEND_ONE
                DirectDevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_ONE
            Case Blit_Alpha.Blendop_Sustrative '3
                DirectDevice.SetRenderState D3DRS_SRCBLEND, D3DBLEND_SRCCOLOR
                DirectDevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_INVSRCCOLOR
            Case Blit_Alpha.Blendop_Inverse  '4
                DirectDevice.SetTextureStageState 0, D3DTSS_ALPHAOP, D3DTOP_MODULATE
                DirectDevice.SetTextureStageState 0, D3DTSS_COLOROP, D3DTOP_SELECTARG1
                DirectDevice.SetTextureStageState 0, D3DTSS_COLORARG1, (D3DTA_TEXTURE Or D3DTA_COMPLEMENT)
                DirectDevice.SetTextureStageState 0, D3DTSS_COLORARG2, D3DTA_CURRENT
                                    
                DirectDevice.SetTextureStageState 1, D3DTSS_COLORARG1, D3DTA_TEXTURE
                DirectDevice.SetTextureStageState 1, D3DTSS_COLOROP, D3DTOP_MODULATE
                DirectDevice.SetTextureStageState 1, D3DTSS_COLORARG2, D3DTA_CURRENT
                                    
            Case Blit_Alpha.Blendop_XOR '5
                DirectDevice.SetRenderState D3DRS_SRCBLEND, D3DBLEND_INVDESTCOLOR
                DirectDevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_INVSRCCOLOR
                
            Case Blit_Alpha.Blendop_Crystaline '6
                DirectDevice.SetRenderState D3DRS_SRCBLEND, D3DBLEND_DESTCOLOR
                DirectDevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_SRCCOLOR
                
            Case Blit_Alpha.Blendop_GreyScale '7
                If Not DX8_GreyScale Then _
                    pixelShaderSet ps1Desat
                    
            Case Else ' SET NORMAL
                DirectDevice.SetRenderState D3DRS_SRCBLEND, D3DBLEND_ONE
                DirectDevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_ONE
        End Select
    End If

        
        
    Dim Dest_Rect As RECT
    Dim Temp_Verts(3) As D3DTLVERTEX
    Dim SRDesc As D3DSURFACE_DESC
    Dim TexWidth As Long
    Dim TexHeight As Long
    
    If ReSize = 0 Then
        With Dest_Rect
            .bottom = Y + (src_rect.bottom - src_rect.Top)
            .Left = X
            .Right = X + (src_rect.Right - src_rect.Left)
            .Top = Y
        End With
    Else
        With Dest_Rect
            .bottom = Y + (src_rect.bottom - src_rect.Top) / ReSize
            .Left = X
            .Right = X + (src_rect.Right - src_rect.Left) / ReSize
            .Top = Y
        End With
    End If
    

    Texture.GetLevelDesc 0, SRDesc

    TexWidth = SRDesc.Width
    TexHeight = SRDesc.Height
    
    
    If Shadow = False Then
        If Alpha <> 0 Then
            Engine_Geometry_Create_Box Temp_Verts(), Dest_Rect, src_rect, ColorData.Blanco, TexWidth, TexHeight, Angle, InvertX, InvertY
        Else
            Engine_Geometry_Create_Box Temp_Verts(), Dest_Rect, src_rect, Color_List(), TexWidth, TexHeight, Angle, InvertX, InvertY
        End If
    Else
        Dim RadAngle As Single
        Dim CenterX As Single
        Dim CenterY As Single
        Dim index As Integer
        Dim NewX As Single
        Dim NewY As Single
        Dim SinRad As Single
        Dim CosRad As Single
        Dim Width As Single
        Dim Height As Single
    
        Width = Int(GrhData(GrhIndex).TileWidth * TilePixelWidth * 0.4) + TilePixelWidth \ 2
        Height = Int(GrhData(GrhIndex).TileHeight * TilePixelHeight) + TilePixelHeight
    
    
        Engine_Geometry_Create_Box Temp_Verts(), Dest_Rect, src_rect, ColorData.SombraAB(), TexWidth, TexHeight, 0, False, False
    
        If Angle <> 0 And Angle <> 360 Then
            Dim DistanceVertex As Single
            DistanceVertex = (Engine_Get_Distance(X, Y, frmMain.MouseX, frmMain.MouseY) * 0.01745329251994) * 0.5


                Temp_Verts(1).sX = Temp_Verts(1).sX - (Width * DistanceVertex) - 5
                Temp_Verts(1).sY = Temp_Verts(1).sY - (Width * DistanceVertex)
                
                Temp_Verts(3).sX = Temp_Verts(3).sX - (Width * DistanceVertex) + 5
                Temp_Verts(3).sY = Temp_Verts(3).sY - (Width * DistanceVertex)
           
                RadAngle = Angle * 0.01745329251994
     
                CenterX = X + (Width * 0.5)
                CenterY = Y + (Height * 0.5)
    
            SinRad = Sin(RadAngle)
            CosRad = Cos(RadAngle)
     
            For index = 0 To 3
     
                NewX = CenterX + (Temp_Verts(index).sX - CenterX) * -CosRad - (Temp_Verts(index).sY - CenterY) * -SinRad
                NewY = CenterY + (Temp_Verts(index).sY - CenterY) * -CosRad + (Temp_Verts(index).sX - CenterX) * -SinRad
     
                Temp_Verts(index).sX = NewX
                Temp_Verts(index).sY = NewY
            Next index
     
        End If
    End If
    
    DirectDevice.SetTexture 0, Texture
    
    ' Faster load.
    DirectDevice.DrawIndexedPrimitiveUP D3DPT_TRIANGLESTRIP, 0, 4, 2, _
                indexList(0), D3DFMT_INDEX16, _
                Temp_Verts(0), Len(Temp_Verts(0))
                
    If Alpha <> 0 Then SetDefaultBlendingStages
    
    If Alpha = Blit_Alpha.Blendop_GreyScale And Not DX8_GreyScale Then
        pixelShaderSet ps1Normal
    ElseIf Alpha = Blit_Alpha.Blendop_Inverse Then
        SetDefaultTextureStages
    End If
End Sub

Public Sub SetDefaultBlendingStages()
    DirectDevice.SetRenderState D3DRS_SRCBLEND, D3DBLEND_SRCALPHA
    DirectDevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_INVSRCALPHA
End Sub

Public Sub SetDefaultTextureStages()
    
    DirectDevice.SetRenderState D3DRS_LIGHTING, False
    DirectDevice.SetRenderState D3DRS_SRCBLEND, D3DBLEND_SRCALPHA
    DirectDevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_INVSRCALPHA
    DirectDevice.SetRenderState D3DRS_ALPHABLENDENABLE, True
    DirectDevice.SetRenderState D3DRS_POINTSIZE, Engine_FToDW(2)
    DirectDevice.SetTextureStageState 0, D3DTSS_ALPHAOP, D3DTOP_MODULATE
    DirectDevice.SetRenderState D3DRS_POINTSPRITE_ENABLE, 1
    DirectDevice.SetRenderState D3DRS_POINTSCALE_ENABLE, 0
    
End Sub

Sub MoveTo(ByVal Direccion As E_Heading)
'***************************************************
'Author: Alejandro Santos (AlejoLp)
'Last Modify Date: 06/28/2008
'Last Modified By: Lucas Tavolaro Ortiz (Tavo)
' 06/03/2006: AlejoLp - Elimine las funciones Move[NSWE] y las converti a esta
' 12/08/2007: Tavo    - Si el usuario esta paralizado no se puede mover.
' 06/28/2008: NicoNZ - Saqué lo que impedía que si el usuario estaba paralizado se ejecute el sub.
'***************************************************
    Dim LegalOk As Boolean
    
    If Cartel Then Cartel = False
    
    Select Case Direccion
        Case E_Heading.NORTH
            LegalOk = MoveToLegalPos(UserPos.X, UserPos.Y - 1)
        Case E_Heading.EAST
            LegalOk = MoveToLegalPos(UserPos.X + 1, UserPos.Y)
        Case E_Heading.SOUTH
            LegalOk = MoveToLegalPos(UserPos.X, UserPos.Y + 1)
        Case E_Heading.WEST
            LegalOk = MoveToLegalPos(UserPos.X - 1, UserPos.Y)
    End Select
    
    If LegalOk And Not UserParalizado Then
        Call WriteWalk(Direccion)
        If Not UserDescansar And Not UserMeditar Then
            MoveCharbyHead UserCharIndex, Direccion
            MoveScreen Direccion
        End If
    Else
        If CharList(UserCharIndex).Heading <> Direccion Then
            Call WriteChangeHeading(Direccion)
        End If
    End If
    
    If frmMain.MacroTrabajo.Enabled Then Call frmMain.DesactivarMacroTrabajo
    
    ' Update 3D sounds!
    Call Audio.MoveListener(UserPos.X, UserPos.Y)
End Sub

Sub RandomMove()
'***************************************************
'Author: Alejandro Santos (AlejoLp)
'Last Modify Date: 06/03/2006
'   06/03/2006: AlejoLp - Ahora utiliza la funcion MoveTo
'***************************************************
    Call MoveTo(General_Get_Random_Number(NORTH, WEST))
End Sub


Sub set_GUI_Efect()

    Dim i As Long
        For i = 1 To 12
            Fade_GUI(i) = General_Get_Random_Number(10, 200)
            Fade_GUIe(i) = General_Get_Random_Number(0, 1)
        Next i
End Sub

