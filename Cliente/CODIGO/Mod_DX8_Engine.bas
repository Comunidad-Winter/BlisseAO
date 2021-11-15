Attribute VB_Name = "Mod_DX8_Engine"
'DX8 Objects
Public DirectX As New DirectX8
Public DirectD3D8 As D3DX8
Public DirectD3D As Direct3D8
Public DirectDevice As Direct3DDevice8
Public DirectCaps As D3DCAPS8

Public Enum Blit_Alpha
    Blendop_None = 0
    Blendop_Color = 1
    Blendop_Aditive = 2
    Blendop_Sustrative = 3
    Blendop_Inverse = 4
    Blendop_XOR = 5
    Blendop_Crystaline = 6
    Blendop_GreyScale = 7
End Enum

Public Enum Texture_Filtering
  TexFilter_None
  TexFilter_Bilinear
  TexFilter_Trilinear
  TexFilter_Anisotropic
End Enum

Public Enum Texture_Inversion
    Invert_None
    Invert_Vertical
    Invert_Horizontal
End Enum

Public DX8_GreyScale As Boolean
Public DX8_HaveWater As Boolean


Public SurfaceDB As New clsSurfaceManDyn

Public Audio As New clsAudio

Public Type TEXTURE_STATISTICS
    Texture As Direct3DTexture8
    TextureWidth As Integer
    TextureHeight As Integer
End Type

Public polygonCount(1) As Single
Public WATER_TICKCOUNT As Long


Public Engine_BaseSpeed As Single
Public Engine_TileBuffer As Integer

Public Const ScreenWidth As Long = 544
Public Const ScreenHeight As Long = 416

Public MainScreenRect As RECT
Public ConnectScreenRect As RECT

Private endtime As Long

Public Type TColores
    Blanco(3) As Long
    BlancoAB(3) As Long
    Rojo(3) As Long
    RojoAB(3) As Long
    Negro(3) As Long
    NegroAB(3) As Long
    Verde(3) As Long
    VerdeAB(3) As Long
    Amarillo(3) As Long
    AmarilloAB(3) As Long
    Azul(3) As Long
    AzulAB(3) As Long
    Dorado(3) As Long
    DoradoAB(3) As Long
    Gris(3) As Long
    GrisAB(3) As Long
    Celeste(3) As Long
    CelesteAB(3) As Long
    SombraAB(3) As Long
    Reflection(3) As Long
    ReflectionBody(3) As Long
    ReflectionHead(3) As Long
End Type

Public ColorData As TColores
    
Public Filtering_Mode As Texture_Filtering

Private Declare Function QueryPerformanceFrequency Lib "kernel32" (lpFrequency As Currency) As Long

Public Sub Engine_BeginScene(Optional ByVal Color As Long = 0)
'***************************************************
'Author: Ezequiel Juárez (Standelf)
'Last Modification: 29/12/10
'Blisse-AO | DD Clear & BeginScene
'***************************************************
    DirectDevice.Clear 0, ByVal 0, D3DCLEAR_TARGET, Color, 1#, 0
    DirectDevice.BeginScene

End Sub

 
Function Engine_Calculate_Polygon(mx As Integer, my As Integer, Xn As Integer, Yn As Integer, Radio As Byte, am As Integer) As Integer
'**************************************************************
'Author: Unknown
'Last Modify Date: 18/10/2012
'**************************************************************
On Error GoTo Error

    Dim Dp As Integer, Dm As Integer
    Dp = Abs(mx - Xn) + Abs(my - Yn)
    Dm = Radio * 32
   
    Engine_Calculate_Polygon = Val(am * (1 - (Dp / Dm)))
    If Engine_Calculate_Polygon < 0 Then Engine_Calculate_Polygon = 0

Error:
    Call Log_Engine("Error in Engine_Calculate_Polygon, " & Err.Description & " (" & Err.Number & ")")
End Function

Private Function Engine_Collision_Between(ByVal value As Single, ByVal Bound1 As Single, ByVal Bound2 As Single) As Byte

    'Checks if a value lies between two bounds
    If Bound1 > Bound2 Then
        If value >= Bound2 Then
            If value <= Bound1 Then Engine_Collision_Between = 1
        End If
    Else
        If value >= Bound1 Then
            If value <= Bound2 Then Engine_Collision_Between = 1
        End If
    End If
    
End Function

Public Function Engine_Collision_Line(ByVal L1X1 As Long, ByVal L1Y1 As Long, ByVal L1X2 As Long, ByVal L1Y2 As Long, ByVal L2X1 As Long, ByVal L2Y1 As Long, ByVal L2X2 As Long, ByVal L2Y2 As Long) As Byte

Dim m1 As Single
Dim M2 As Single
Dim b1 As Single
Dim b2 As Single
Dim IX As Single

    'This will fix problems with vertical lines
    If L1X1 = L1X2 Then L1X1 = L1X1 + 1
    If L2X1 = L2X2 Then L2X1 = L2X1 + 1

    'Find the first slope
    m1 = (L1Y2 - L1Y1) / (L1X2 - L1X1)
    b1 = L1Y2 - m1 * L1X2

    'Find the second slope
    M2 = (L2Y2 - L2Y1) / (L2X2 - L2X1)
    b2 = L2Y2 - M2 * L2X2
    
    'Check if the slopes are the same
    If M2 - m1 = 0 Then
    
        If b2 = b1 Then
            'The lines are the same
            Engine_Collision_Line = 1
        Else
            'The lines are parallel (can never intersect)
            Engine_Collision_Line = 0
        End If
        
    Else
        
        'An intersection is a point that lies on both lines. To find this, we set the Y equations equal and solve for X.
        'M1X+B1 = M2X+B2 -> M1X-M2X = -B1+B2 -> X = B1+B2/(M1-M2)
        IX = ((b2 - b1) / (m1 - M2))
        
        'Check for the collision
        If Engine_Collision_Between(IX, L1X1, L1X2) Then
            If Engine_Collision_Between(IX, L2X1, L2X2) Then Engine_Collision_Line = 1
        End If
        
    End If
    
End Function

Public Function Engine_Collision_LineRect(ByVal sX As Long, ByVal sY As Long, ByVal SW As Long, ByVal SH As Long, ByVal x1 As Long, ByVal Y1 As Long, ByVal x2 As Long, ByVal Y2 As Long) As Byte

    'Top line
    If Engine_Collision_Line(sX, sY, sX + SW, sY, x1, Y1, x2, Y2) Then
        Engine_Collision_LineRect = 1
        Exit Function
    End If
    
    'Right line
    If Engine_Collision_Line(sX + SW, sY, sX + SW, sY + SH, x1, Y1, x2, Y2) Then
        Engine_Collision_LineRect = 1
        Exit Function
    End If

    'Bottom line
    If Engine_Collision_Line(sX, sY + SH, sX + SW, sY + SH, x1, Y1, x2, Y2) Then
        Engine_Collision_LineRect = 1
        Exit Function
    End If

    'Left line
    If Engine_Collision_Line(sX, sY, sX, sY + SW, x1, Y1, x2, Y2) Then
        Engine_Collision_LineRect = 1
        Exit Function
    End If

End Function

Function Engine_Collision_Rect(ByVal x1 As Integer, ByVal Y1 As Integer, ByVal Width1 As Integer, ByVal Height1 As Integer, ByVal x2 As Integer, ByVal Y2 As Integer, ByVal Width2 As Integer, ByVal Height2 As Integer) As Boolean

    If x1 + Width1 >= x2 Then
        If x1 <= x2 + Width2 Then
            If Y1 + Height1 >= Y2 Then
                If Y1 <= Y2 + Height2 Then
                    Engine_Collision_Rect = True
                End If
            End If
        End If
    End If

End Function

Public Function Engine_Convert_Degrees_To_Radians(ByVal s_degree As Double) As Single
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 5/12/2003
'Converts a degree to a radian
'**************************************************************

    Engine_Convert_Degrees_To_Radians = (s_degree * 3.14159265358979) / 180
    
End Function

 
Public Function Engine_Convert_Radians_To_Degrees(ByVal s_radians As Double) As Integer
'**************************************************************
'Author: Juan Martín Sotuyo Dodero
'Last Modify Date: 8/25/2004
'Converts a radian to degrees
'**************************************************************

      Engine_Convert_Radians_To_Degrees = (s_radians * 180) / 3.14159265358979
 
End Function

Public Function Engine_Create_D3DTLVertex(ByVal X As Single, ByVal Y As Single, ByVal Z As Single, _
                                            ByVal rhw As Single, ByVal Color As Long, ByVal Specular As Long, Tu As Single, _
                                            ByVal Tv As Single) As D3DTLVERTEX
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 10/07/2002
'**************************************************************
    Engine_Create_D3DTLVertex.sX = X
    Engine_Create_D3DTLVertex.sY = Y
    Engine_Create_D3DTLVertex.rhw = rhw
    Engine_Create_D3DTLVertex.Color = Color
    Engine_Create_D3DTLVertex.Specular = Specular
    Engine_Create_D3DTLVertex.Tu = Tu
    Engine_Create_D3DTLVertex.Tv = Tv
End Function

Public Sub Engine_Create_Elevation(ByVal Altura As Integer, ByVal Radio_X As Integer, ByVal Radio_Y As Integer, _
                            ByVal Pos_X As Integer, ByVal Pos_Y As Integer)
'**************************************************************
'Author: Unknown
'Last Modify Date: 18/10/2012
'**************************************************************
On Error GoTo Error
    Dim X           As Byte
    Dim Y           As Byte
    Dim MinX        As Integer
    Dim MaxX        As Integer
   
    For Y = Pos_Y To Pos_Y + Radio_Y / 2
        MinX = Pos_X - Radio_X / 2 + (Y - Pos_Y - 1)
        MaxX = Pos_X + Radio_X / 2 - (Y - Pos_Y - 1)

        For X = MinX To MaxX
                MapData(X, Y).Vertex_Offset(0) = Altura * Sin(DegreeToRadian * 180 * (X - (Pos_X - Radio_X / 2 + Y - Pos_Y)) / (MaxX - MinX))
                MapData(X, Y).Vertex_Offset(1) = Altura * Sin(DegreeToRadian * 180 * (X + 1 - (Pos_X - Radio_X / 2 + Y - Pos_Y)) / (MaxX - MinX))
               
                If MapData(X, Y).Vertex_Offset(0) < 0 Then MapData(X, Y).Vertex_Offset(0) = 0
               
                MapData(X, Y - 1).Vertex_Offset(2) = MapData(X, Y).Vertex_Offset(0)
                MapData(X, Y - 1).Vertex_Offset(3) = MapData(X, Y).Vertex_Offset(1)

                'Set Ambient Vertex
                CurMapAmbient.MapBlocks(X, Y).Vertex_Offset(0) = MapData(X, Y).Vertex_Offset(0)
                CurMapAmbient.MapBlocks(X, Y).Vertex_Offset(1) = MapData(X, Y).Vertex_Offset(1)
                CurMapAmbient.MapBlocks(X, Y).Vertex_Offset(2) = MapData(X, Y).Vertex_Offset(2)
                CurMapAmbient.MapBlocks(X, Y).Vertex_Offset(3) = MapData(X, Y).Vertex_Offset(3)
                CurMapAmbient.MapBlocks(X, Y - 1).Vertex_Offset(2) = MapData(X, Y - 1).Vertex_Offset(2)
                CurMapAmbient.MapBlocks(X, Y - 1).Vertex_Offset(3) = MapData(X, Y - 1).Vertex_Offset(3)
        Next X
    Next Y
Error:
    Call Log_Engine("Error in Engine_Create_Elevation, " & Err.Description & " (" & Err.Number & ")")
End Sub

Sub Engine_Create_Polygon(X As Integer, Y As Integer, Radio As Byte, alturamaxima As Integer, Optional Sube As Boolean = True)
'**************************************************************
'Author: Unknown
'Last Modify Date: 18/10/2012
'**************************************************************
On Error GoTo Error

Dim xb As Integer, yb As Integer
 
For xb = X - Radio To X + Radio
    For yb = Y - Radio To Y + Radio
        If Sube Then
            MapData(xb, yb).Vertex_Offset(0) = Engine_Calculate_Polygon(xb * 32, yb * 32, X * 32, Y * 32, Radio, alturamaxima)
            MapData(xb, yb).Vertex_Offset(1) = Engine_Calculate_Polygon(xb * 32 + 32, yb * 32, X * 32, Y * 32, Radio, alturamaxima)
            MapData(xb, yb).Vertex_Offset(2) = Engine_Calculate_Polygon(xb * 32, yb * 32 + 32, X * 32, Y * 32, Radio, alturamaxima)
            MapData(xb, yb).Vertex_Offset(3) = Engine_Calculate_Polygon(xb * 32 + 32, yb * 32 + 32, X * 32, Y * 32, Radio, alturamaxima)
        Else
            MapData(xb, yb).Vertex_Offset(0) = -Engine_Calculate_Polygon(xb * 32, yb * 32, X * 32, Y * 32, Radio, alturamaxima)
            MapData(xb, yb).Vertex_Offset(1) = -Engine_Calculate_Polygon(xb * 32 + 32, yb * 32, X * 32, Y * 32, Radio, alturamaxima)
            MapData(xb, yb).Vertex_Offset(2) = -Engine_Calculate_Polygon(xb * 32, yb * 32 + 32, X * 32, Y * 32, Radio, alturamaxima)
            MapData(xb, yb).Vertex_Offset(3) = -Engine_Calculate_Polygon(xb * 32 + 32, yb * 32 + 32, X * 32, Y * 32, Radio, alturamaxima)
        End If
        
            'Set Ambient Vertex
            CurMapAmbient.MapBlocks(xb, yb).Vertex_Offset(0) = MapData(xb, yb).Vertex_Offset(0)
            CurMapAmbient.MapBlocks(xb, yb).Vertex_Offset(1) = MapData(xb, yb).Vertex_Offset(1)
            CurMapAmbient.MapBlocks(xb, yb).Vertex_Offset(2) = MapData(xb, yb).Vertex_Offset(2)
            CurMapAmbient.MapBlocks(xb, yb).Vertex_Offset(3) = MapData(xb, yb).Vertex_Offset(3)

    Next yb
Next xb

Error:
    Call Log_Engine("Error in Engine_Create_Polygon, " & Err.Description & " (" & Err.Number & ")")
End Sub

Public Sub Engine_D3DColor_To_RGB_List(rgb_list() As Long, Color As D3DCOLORVALUE)
'***************************************************
'Author: Ezequiel Juárez (Standelf)
'Last Modification: 14/05/10
'Blisse-AO | Set a D3DColorValue to a RGB List
'***************************************************
    rgb_list(0) = D3DColorARGB(Color.a, Color.r, Color.g, Color.b)
    rgb_list(1) = rgb_list(0)
    rgb_list(2) = rgb_list(0)
    rgb_list(3) = rgb_list(0)
End Sub

Public Sub Engine_DirectX8_End()
'***************************************************
'Author: Standelf
'Last Modification: 26/05/2010
'Destroys all DX objects
'***************************************************
On Error Resume Next
    Dim i As Byte
    
    '   DeInit Lights
    Call DeInit_LightEngine
    
    '   DeInit Auras
    Call DeInit_Auras
    
    '   Clean Particles
    For i = 1 To UBound(ParticleTexture)
        If Not ParticleTexture(i) Is Nothing Then Set ParticleTexture(i) = Nothing
    Next i
    
    '   Clean Texture
    DirectDevice.SetTexture 0, Nothing

    '   Erase Data
    Erase MapData()
    Erase CharList()
    
    Set DirectD3D8 = Nothing
    Set DirectD3D = Nothing
    Set DirectX = Nothing
    Set DirectDevice = Nothing
     
    Audio.MP3_Destroy
 
End Sub

Public Function Engine_DirectX8_Init() As Boolean

    Dim DispMode As D3DDISPLAYMODE
    Dim D3DWindow As D3DPRESENT_PARAMETERS
    
    Set DirectX = New DirectX8
    Set DirectD3D = DirectX.Direct3DCreate
    Set DirectD3D8 = New D3DX8

    DirectD3D.GetAdapterDisplayMode D3DADAPTER_DEFAULT, DispMode
    
    With D3DWindow
        .Windowed = True
        .SwapEffect = IIf((Settings.vSync) = True, D3DSWAPEFFECT_COPY_VSYNC, D3DSWAPEFFECT_COPY)
        .BackBufferFormat = DispMode.Format
        
        DispMode.Height = 600
        DispMode.Width = 800
        
        .BackBufferWidth = IIf(Settings.Ventana, 0, DispMode.Width)
        .BackBufferHeight = IIf(Settings.Ventana, 0, DispMode.Height)
        .EnableAutoDepthStencil = 1
        .AutoDepthStencilFormat = D3DFMT_D16
        .hDeviceWindow = frmMain.MainViewPic.hWnd
    End With

    Dim Aceleration As CONST_D3DCREATEFLAGS

    Select Case Settings.Aceleracion
        Case 0: Aceleration = D3DCREATE_SOFTWARE_VERTEXPROCESSING
        Case 1: Aceleration = D3DCREATE_HARDWARE_VERTEXPROCESSING
        Case 2: Aceleration = D3DCREATE_MIXED_VERTEXPROCESSING
        Case Else: Aceleration = D3DCREATE_SOFTWARE_VERTEXPROCESSING
    End Select
    
        Set DirectDevice = DirectD3D.CreateDevice(D3DADAPTER_DEFAULT, D3DDEVTYPE_HAL, _
                                frmMain.MainViewPic.hWnd, Aceleration, _
                                D3DWindow)

    DirectDevice.SetVertexShader D3DFVF_XYZRHW Or D3DFVF_TEX1 Or D3DFVF_DIFFUSE Or D3DFVF_SPECULAR
    
    ' Set Default Stages
    SetDefaultTextureStages
   
    endtime = GetTickCount
    'FPS = 101
    'FramesPerSecCounter = 100
    Engine_Set_TileBuffer Settings.BufferSize
    Engine_Set_BaseSpeed 0.016

    With MainScreenRect
        .bottom = frmMain.MainViewPic.ScaleHeight
        .Right = frmMain.MainViewPic.ScaleWidth
    End With

    With ConnectScreenRect
        .bottom = frmConnect.ScaleHeight
        .Right = frmConnect.ScaleWidth
    End With

    Engine_Init_FontSettings 1
    Engine_Init_FontTextures 1
    Engine_Init_FontSettings 2
    Engine_Init_FontTextures 2
    Engine_Init_FontSettings 3
    Engine_Init_FontTextures 3
    
    Engine_Make_Color_Data
    
    Init_MeteoEngine
    Engine_Init_ParticleEngine
    alturaAgua = 4
    
    Call SurfaceDB.Initialize(DirectD3D8, Settings.useVideoMemory, Resources.Graphics, Settings.MemoryVideoMax)

'   ERROR CONTROL
    If Err Then
        MsgBox "No se puede iniciar DirectX. Por favor asegurese de tener la última versión correctamente instalada. Puede descargarla desde: " & Client_Web & "support/directx.zip"
        Engine_DirectX8_Init = False
        Exit Function
    End If
    
    If Err Then
        MsgBox "No se puede iniciar DirectD3D. Por favor asegurese de tener la última versión correctamente instalada. Puede descargarla desde: " & Client_Web & "support/directx.zip"
        Engine_DirectX8_Init = False
        Exit Function
    End If
    
    If DirectDevice Is Nothing Then
        MsgBox "No se puede inicializar DirectDevice. Por favor asegurese de tener la última versión correctamente instalada. Puede descargarla desde: " & Client_Web & "support/directx.zip"
        Engine_DirectX8_Init = False
        Exit Function
    End If
    
    Engine_DirectX8_Init = True
End Function

Function Engine_Distance(ByVal x1 As Integer, ByVal Y1 As Integer, ByVal x2 As Integer, ByVal Y2 As Integer) As Long
'***************************************************
'Author: Standelf
'Last Modification: -
'***************************************************

    Engine_Distance = Abs(x1 - x2) + Abs(Y1 - Y2)
    
End Function

Public Function Engine_ElapsedTime() As Long
'**************************************************************
'Gets the time that past since the last call
'More info: http://www.vbgore.com/GameClient.TileEngine.Engine_ElapsedTime
'**************************************************************
Dim Start_Time As Long

    'Get current time
    Start_Time = GetTickCount

    'Calculate elapsed time
    Engine_ElapsedTime = Start_Time - endtime

    'Get next end time
    endtime = Start_Time

End Function

Public Sub Engine_EndScene(ByRef DestRect As RECT, Optional ByVal hWndDest As Long = 0)
'***************************************************
'Author: Ezequiel Juárez (Standelf)
'Last Modification: 29/12/10
'Blisse-AO | DD EndScene & Present
'***************************************************
    
    If hWndDest = 0 Then
        DirectDevice.EndScene
        DirectDevice.Present DestRect, ByVal 0, ByVal 0, ByVal 0
    Else
        DirectDevice.EndScene
        DirectDevice.Present DestRect, ByVal 0, hWndDest, ByVal 0
    End If
    
End Sub

Public Function Engine_FToDW(f As Single) As Long
'*****************************************************************
'Converts a float to a D-Word, or in Visual Basic terms, a Single to a Long
'More info: http://www.vbgore.com/CommonCode.Particles.Effect_FToDW
'*****************************************************************
Dim buf As D3DXBuffer

    'Converts a single into a long (Float to DWORD)
    Set buf = DirectD3D8.CreateBuffer(4)
    DirectD3D8.BufferSetData buf, 0, 4, 1, f
    DirectD3D8.BufferGetData buf, 0, 4, 1, Effect_FToDW

End Function

Public Sub Engine_Geometry_Create_Box(ByRef verts() As D3DTLVERTEX, ByRef dest As RECT, ByRef src As RECT, ByRef rgb_list() As Long, _
                                Optional ByRef Textures_Width As Long, Optional ByRef Textures_Height As Long, Optional ByVal Angle As Single, Optional ByVal InvertX As Boolean = False, Optional ByVal InvertY As Boolean = False)
    Dim x_center As Single
    Dim y_center As Single
    Dim radius As Single
    Dim x_Cor As Single
    Dim y_Cor As Single
    Dim left_point As Single
    Dim right_point As Single
    Dim Temp As Single
    Dim auxr As RECT
    If Angle > 0 Then
        'Center coordinates on screen of the square
        x_center = dest.Left + (dest.Right - dest.Left) / 2
        y_center = dest.Top + (dest.bottom - dest.Top) / 2

        radius = Sqr((dest.Right - x_center) ^ 2 + (dest.bottom - y_center) ^ 2)
       
        Temp = (dest.Right - x_center) / radius
        right_point = Atn(Temp / Sqr(-Temp * Temp + 1))
        left_point = 3.1459 - right_point
    End If

    If Angle = 0 Then
        x_Cor = dest.Left
        y_Cor = dest.bottom
    Else
        x_Cor = x_center + Cos(-left_point - Angle) * radius
        y_Cor = y_center - Sin(-left_point - Angle) * radius
    End If
   
    auxr = src
    If InvertX Then
        src.Left = auxr.Right
        src.Right = auxr.Left
    End If
    
    If InvertY Then
        src.Top = auxr.bottom
        src.bottom = auxr.Top
    End If
    
    
    '0 - Bottom left vertex
    If Textures_Width Or Textures_Height Then
        verts(0) = Engine_Create_D3DTLVertex(x_Cor, y_Cor, 0, 1, rgb_list(2), 0, src.Left / Textures_Width + 0.001, (src.bottom + 1) / Textures_Height)
    Else
        verts(0) = Engine_Create_D3DTLVertex(x_Cor, y_Cor, 0, 1, rgb_list(2), 0, 0, 0)
    End If
    'Calculate screen coordinates of sprite, and only rotate if necessary
    If Angle = 0 Then
        x_Cor = dest.Left
        y_Cor = dest.Top
    Else
        x_Cor = x_center + Cos(left_point - Angle) * radius
        y_Cor = y_center - Sin(left_point - Angle) * radius
    End If
   
   
    '1 - Top left vertex
    If Textures_Width Or Textures_Height Then
        verts(1) = Engine_Create_D3DTLVertex(x_Cor, y_Cor, 0, 1, rgb_list(0), 0, src.Left / Textures_Width + 0.001, src.Top / Textures_Height + 0.001)
    Else
        verts(1) = Engine_Create_D3DTLVertex(x_Cor, y_Cor, 0, 1, rgb_list(0), 0, 0, 1)
    End If
    'Calculate screen coordinates of sprite, and only rotate if necessary
    If Angle = 0 Then
        x_Cor = dest.Right
        y_Cor = dest.bottom
    Else
        x_Cor = x_center + Cos(-right_point - Angle) * radius
        y_Cor = y_center - Sin(-right_point - Angle) * radius
    End If
   
   
    '2 - Bottom right vertex
    If Textures_Width Or Textures_Height Then
        verts(2) = Engine_Create_D3DTLVertex(x_Cor, y_Cor, 0, 1, rgb_list(3), 0, (src.Right + 1) / Textures_Width, (src.bottom + 1) / Textures_Height)
    Else
        verts(2) = Engine_Create_D3DTLVertex(x_Cor, y_Cor, 0, 1, rgb_list(3), 0, 1, 0)
    End If
    'Calculate screen coordinates of sprite, and only rotate if necessary
    If Angle = 0 Then
        x_Cor = dest.Right
        y_Cor = dest.Top
    Else
        x_Cor = x_center + Cos(right_point - Angle) * radius
        y_Cor = y_center - Sin(right_point - Angle) * radius
    End If
   
   
    '3 - Top right vertex
    If Textures_Width Or Textures_Height Then
        verts(3) = Engine_Create_D3DTLVertex(x_Cor, y_Cor, 0, 1, rgb_list(1), 0, (src.Right + 1) / Textures_Width, src.Top / Textures_Height + 0.001)
    Else
        verts(3) = Engine_Create_D3DTLVertex(x_Cor, y_Cor, 0, 1, rgb_list(1), 0, 1, 1)
    End If
 
End Sub

Public Function Engine_Get_2_Points_Angle(ByVal x1 As Long, ByVal Y1 As Long, ByVal x2 As Long, ByVal Y2 As Long) As Double
'**************************************************************
'Author: Unknown
'Last Modify Date: 18/10/2012
'**************************************************************

    Engine_Get_2_Points_Angle = Engine_Get_X_Y_Angle((x2 - x1), (Y2 - Y1))
   
End Function

Public Sub Engine_Get_ARGB(Color As Long, data As D3DCOLORVALUE)
'**************************************************************
'Author: Standelf
'Last Modify Date: 18/10/2012
'**************************************************************
    
    Dim a As Long, r As Long, g As Long, b As Long
        
    If Color < 0 Then
        a = ((Color And (&H7F000000)) / (2 ^ 24)) Or &H80&
    Else
        a = Color / (2 ^ 24)
    End If
    
    r = (Color And &HFF0000) / (2 ^ 16)
    g = (Color And &HFF00&) / (2 ^ 8)
    b = (Color And &HFF&)
    
    With data
        .a = a
        .r = r
        .g = g
        .b = b
    End With
        
End Sub

Public Function Engine_Get_BaseSpeed() As Single
'**************************************************************
'Author: Standelf
'Last Modify Date: 29/12/2010
'**************************************************************

    Engine_Get_BaseSpeed = Engine_BaseSpeed
    
End Function

Function Engine_Get_Distance(ByVal x1 As Integer, ByVal Y1 As Integer, ByVal x2 As Integer, ByVal Y2 As Integer) As Single
 
    Engine_Get_Distance = (Abs(x2 - x1) + Abs(Y2 - Y1)) * 0.5
   
End Function

Public Function Engine_Get_TileBuffer() As Single
'**************************************************************
'Author: Standelf
'Last Modify Date: 30/12/2010
'**************************************************************

    Engine_Get_TileBuffer = Engine_TileBuffer
    
End Function

' ####################################################################################
' ####################################################################################
' ########          OTRAS FUNCIONES
' ####################################################################################
' ####################################################################################

Public Function Engine_Get_Vertex_Offset() As Single
'**************************************************************
'Author: Standelf
'Last Modify Date: 18/10/2012
'**************************************************************

    Engine_Get_Vertex_Offset = MapData(UserPos.X, UserPos.Y).Vertex_Offset(0)

End Function

 
Public Function Engine_Get_X_Y_Angle(ByVal X As Double, ByVal Y As Double) As Double
'**************************************************************
'Author: Unknown
'Last Modify Date: 18/10/2012
'**************************************************************

Dim dblres              As Double
 
    dblres = 0
   
    If (Y <> 0) Then
        dblres = Engine_Convert_Radians_To_Degrees(Atn(X / Y))
        If (X <= 0 And Y < 0) Then
            dblres = dblres + 180
        ElseIf (X > 0 And Y < 0) Then
            dblres = dblres + 180
        ElseIf (X < 0 And Y > 0) Then
            dblres = dblres + 360
        End If
    Else
        If (X > 0) Then
            dblres = 90
        ElseIf (X < 0) Then
            dblres = 270
        End If
    End If
   
    Engine_Get_X_Y_Angle = dblres
   
End Function

Public Function Engine_Invert(ByVal Number As Variant) As Long
    Number = Number * -1
End Function

Public Sub Engine_Long_To_RGB_List(rgb_list() As Long, Long_Color As Long)
'***************************************************
'Author: Ezequiel Juárez (Standelf)
'Last Modification: 16/05/10
'Blisse-AO | Set a Long Color to a RGB List
'***************************************************
    rgb_list(0) = Long_Color
    rgb_list(1) = rgb_list(0)
    rgb_list(2) = rgb_list(0)
    rgb_list(3) = rgb_list(0)
End Sub

Public Sub Engine_Make_Color_Data()
    Call Engine_Long_To_RGB_List(ColorData.Blanco, D3DColorXRGB(255, 255, 255))
    Call Engine_Long_To_RGB_List(ColorData.Rojo, D3DColorXRGB(255, 0, 0))
    Call Engine_Long_To_RGB_List(ColorData.Negro, D3DColorXRGB(0, 0, 0))
    Call Engine_Long_To_RGB_List(ColorData.Gris, D3DColorXRGB(100, 100, 100))
    Call Engine_Long_To_RGB_List(ColorData.Amarillo, D3DColorXRGB(255, 255, 0))
    Call Engine_Long_To_RGB_List(ColorData.Dorado, D3DColorXRGB(200, 220, 0))
    Call Engine_Long_To_RGB_List(ColorData.Azul, D3DColorXRGB(0, 0, 255))
    Call Engine_Long_To_RGB_List(ColorData.Verde, D3DColorXRGB(0, 255, 0))
    Call Engine_Long_To_RGB_List(ColorData.Celeste, D3DColorXRGB(153, 217, 234))
    
    Call Engine_Long_To_RGB_List(ColorData.BlancoAB, D3DColorARGB(100, 255, 255, 255))
    Call Engine_Long_To_RGB_List(ColorData.RojoAB, D3DColorARGB(100, 255, 0, 0))
    Call Engine_Long_To_RGB_List(ColorData.NegroAB, D3DColorARGB(100, 0, 0, 0))
    Call Engine_Long_To_RGB_List(ColorData.GrisAB, D3DColorARGB(100, 100, 100, 100))
    Call Engine_Long_To_RGB_List(ColorData.AmarilloAB, D3DColorARGB(100, 255, 255, 0))
    Call Engine_Long_To_RGB_List(ColorData.DoradoAB, D3DColorARGB(100, 200, 220, 0))
    Call Engine_Long_To_RGB_List(ColorData.AzulAB, D3DColorARGB(100, 0, 0, 255))
    Call Engine_Long_To_RGB_List(ColorData.VerdeAB, D3DColorARGB(100, 0, 255, 0))
    Call Engine_Long_To_RGB_List(ColorData.CelesteAB, D3DColorARGB(100, 153, 217, 234))
    
    ColorData.SombraAB(1) = D3DColorARGB(0, 0, 0, 0)
    ColorData.SombraAB(0) = D3DColorARGB(0, 0, 0, 0)
    ColorData.SombraAB(3) = D3DColorARGB(100, 0, 0, 0)
    ColorData.SombraAB(2) = D3DColorARGB(100, 0, 0, 0)
    
    ColorData.Reflection(3) = D3DColorARGB(20, 255, 255, 255)
    ColorData.Reflection(2) = D3DColorARGB(20, 255, 255, 255)
    ColorData.Reflection(1) = D3DColorARGB(100, 255, 255, 255)
    ColorData.Reflection(0) = D3DColorARGB(100, 255, 255, 255)
    
    
    ColorData.ReflectionBody(3) = D3DColorARGB(50, 255, 255, 255)
    ColorData.ReflectionBody(2) = D3DColorARGB(50, 255, 255, 255)
    ColorData.ReflectionBody(1) = D3DColorARGB(100, 255, 255, 255)
    ColorData.ReflectionBody(0) = D3DColorARGB(100, 255, 255, 255)
    
    ColorData.ReflectionHead(3) = D3DColorARGB(0, 255, 255, 255)
    ColorData.ReflectionHead(2) = D3DColorARGB(0, 255, 255, 255)
    ColorData.ReflectionHead(1) = D3DColorARGB(150, 255, 255, 255)
    ColorData.ReflectionHead(0) = D3DColorARGB(150, 255, 255, 255)
     
End Sub

Public Function Engine_PixelPosX(ByVal X As Integer) As Integer
'*****************************************************************
'Converts a tile position to a screen position
'More info: http://www.vbgore.com/GameClient.TileEngine.Engine_PixelPosX
'*****************************************************************

    Engine_PixelPosX = (X - 1) * 32
    
End Function

Public Function Engine_PixelPosY(ByVal Y As Integer) As Integer
'*****************************************************************
'Converts a tile position to a screen position
'More info: http://www.vbgore.com/GameClient.TileEngine.Engine_PixelPosY
'*****************************************************************

    Engine_PixelPosY = (Y - 1) * 32
    
End Function

Public Sub Engine_Render_Fill_Box(ByVal X As Integer, ByVal Y As Integer, ByVal Width As Integer, ByVal Height As Integer, Color As Long)
'***************************************************
'Author: Ezequiel Juárez (Standelf)
'Last Modification: 29/12/10
'Blisse-AO | Render Box
'***************************************************
    Dim b_Rect As RECT
    Dim b_Color(0 To 3) As Long
    Dim b_Vertex(0 To 3) As D3DTLVERTEX
    
    Engine_Long_To_RGB_List b_Color(), Color

    With b_Rect
        .bottom = Y + Height
        .Left = X
        .Right = X + Width
        .Top = Y
    End With

    Engine_Geometry_Create_Box b_Vertex(), b_Rect, b_Rect, b_Color(), 0, 0
    
    DirectDevice.SetTexture 0, Nothing
    DirectDevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, b_Vertex(0), Len(b_Vertex(0))
End Sub

Public Sub Engine_Render_FPS()
'***************************************************
'Author: Standelf
'Last Modification: 28/02/2013
'Limit FPS & Calculate later, The new mod calculate more exactly
'***************************************************

If Settings.MostrarFPS = False Then Exit Sub

    Dim ActualTime As Long
    ActualTime = GetTickCount
    FramesPerSecondCounter = FramesPerSecondCounter + 1
    
    If Settings.LimiteFPS And Not Settings.vSync Then
        While (GetTickCount - FramesPerSecondLastTime) \ 10 < FramesPerSecondCounter
            Sleep 1
        Wend
    End If
    
    If (ActualTime - FramesPerSecondLastTime) > 1000 Then
        FramesPerSecond = (FramesPerSecondCounter / ((ActualTime - FramesPerSecondLastTime) / 1000))
        FramesPerSecondLastTime = ActualTime
        FramesPerSecondCounter = 0
    End If
    
    Fonts_Render_String Round(FramesPerSecond, 2), 2, 2, ColorData.BlancoAB(1), False, 1
End Sub

Public Sub Engine_Render_Layer1(ByVal X As Long, ByVal Y As Long, _
                                ByVal screenX As Integer, ByVal screenY As Integer, _
                                ByVal PixelOffsetX As Integer, ByVal PixelOffsetY As Integer)

' / - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
' / Author: Emanuel Matías 'Dunkan'
' / Note: Efectos de la capa 1, movimiento de polígonos y dibujado
' / - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -


    'SetDefaultTextureStages
        
Dim VertexArray(0 To 3) As D3DTLVERTEX
Dim D3DTextures         As TEXTURE_STATISTICS
Dim SRDesc              As D3DSURFACE_DESC
Dim SrcWidth            As Integer
Dim Width               As Integer
Dim SrcHeight           As Integer
Dim Height              As Integer
Dim new_x               As Integer
Dim new_y               As Integer
Dim SrcBitmapWidth      As Long
Dim SrcBitmapHeight     As Long
            
If MapData(X, Y).Graphic(1).GrhIndex Then

    new_x = (screenX - 1) * 32 + PixelOffsetX
    new_y = (screenY - 1) * 32 + PixelOffsetY
           
    If MapData(X, Y).Graphic(1).Started = 1 Then
    
        MapData(X, Y).Graphic(1).FrameCounter = MapData(X, Y).Graphic(1).FrameCounter + ((timerElapsedTime * 0.1) * GrhData(MapData(X, Y).Graphic(1).GrhIndex).NumFrames / MapData(X, Y).Graphic(1).Speed)
            
            If MapData(X, Y).Graphic(1).FrameCounter > GrhData(MapData(X, Y).Graphic(1).GrhIndex).NumFrames Then
                MapData(X, Y).Graphic(1).FrameCounter = (MapData(X, Y).Graphic(1).FrameCounter Mod GrhData(MapData(X, Y).Graphic(1).GrhIndex).NumFrames) + 1
            End If
            
    End If
                        
    Dim iGrhIndex   As Integer
    
    iGrhIndex = GrhData(MapData(X, Y).Graphic(1).GrhIndex).Frames(MapData(X, Y).Graphic(1).FrameCounter)
            
    With GrhData(iGrhIndex)
    
        Set D3DTextures.Texture = SurfaceDB.Surface(.FileNum) 'Cargamos la textura
                
        D3DTextures.Texture.GetLevelDesc 0, SRDesc ' Medimos las texturas
        
        D3DTextures.TextureWidth = SRDesc.Width
        D3DTextures.TextureHeight = SRDesc.Height
                
        SrcWidth = 32
        Width = 32
                   
        Height = 32
        SrcHeight = 32
                
        SrcBitmapWidth = D3DTextures.TextureWidth
        SrcBitmapHeight = D3DTextures.TextureHeight
               
        'Seteamos los RHW a 1
        VertexArray(0).rhw = 1
        VertexArray(1).rhw = 1
        VertexArray(2).rhw = 1
        VertexArray(3).rhw = 1
             
        'Find the left side of the rectangle
        VertexArray(0).sX = new_x
        VertexArray(0).Tu = (.sX / SrcBitmapWidth)
             
        'Find the top side of the rectangle
        VertexArray(0).sY = new_y
        VertexArray(0).Tv = (.sY / SrcBitmapHeight)
               
        'Find the right side of the rectangle
        VertexArray(1).sX = new_x + Width
        VertexArray(1).Tu = (.sX + SrcWidth) / SrcBitmapWidth
             
        'These values will only equal each other when not a shadow
        VertexArray(2).sX = VertexArray(0).sX
        VertexArray(3).sX = VertexArray(1).sX
               
        'Find the bottom of the rectangle
        VertexArray(2).sY = new_y + Height
        VertexArray(2).Tv = (.sY + SrcHeight) / SrcBitmapHeight
             
        'Because this is a perfect rectangle, all of the values below will equal one of the values we already got
        VertexArray(1).sY = VertexArray(0).sY
        VertexArray(1).Tv = VertexArray(0).Tv
        VertexArray(2).Tu = VertexArray(0).Tu
        VertexArray(3).sY = VertexArray(2).sY
        VertexArray(3).Tu = VertexArray(1).Tu
        VertexArray(3).Tv = VertexArray(2).Tv
                            
        VertexArray(0).sY = VertexArray(0).sY - MapData(X, Y).Vertex_Offset(0)
        VertexArray(1).sY = VertexArray(1).sY - MapData(X, Y).Vertex_Offset(1)
        VertexArray(2).sY = VertexArray(2).sY - MapData(X, Y).Vertex_Offset(2)
        VertexArray(3).sY = VertexArray(3).sY - MapData(X, Y).Vertex_Offset(3)
        
        'Static Polygon As Long
        'Dim tmp_Color As D3DCOLORVALUE
       
        'Dim PolygonColor(0 To 2) As Long
        'For Polygon = 0 To 3
        '    If MapData(X, Y).Vertex_Offset(Polygon) <> 0 Then
        '
        '    Call Engine_Get_ARGB(MapData(X, Y).Engine_Light(Polygon), tmp_Color)
        '
        '    PolygonColor(0) = tmp_Color.r - (MapData(X, Y).Vertex_Offset(Polygon))
        '    PolygonColor(1) = tmp_Color.g - (MapData(X, Y).Vertex_Offset(Polygon))
        '    PolygonColor(2) = tmp_Color.b - (MapData(X, Y).Vertex_Offset(Polygon))
        '
        '    If PolygonColor(0) < 0 Then PolygonColor(0) = MapData(X, Y).Vertex_Offset(Polygon) / 6
        '        If PolygonColor(1) < 0 Then PolygonColor(1) = MapData(X, Y).Vertex_Offset(Polygon) / 6
        '                If PolygonColor(2) < 0 Then PolygonColor(2) = MapData(X, Y).Vertex_Offset(Polygon) / 6
        '
        '       VertexArray(Polygon).Color = D3DColorARGB(255, PolygonColor(0), PolygonColor(1), PolygonColor(2))
        '    Else
        '        VertexArray(Polygon).Color = MapData(X, Y).Engine_Light(Polygon)
        '    End If
        'Next Polygon
        VertexArray(0).Color = MapData(X, Y).Engine_Light(0)
        VertexArray(1).Color = MapData(X, Y).Engine_Light(1)
        VertexArray(2).Color = MapData(X, Y).Engine_Light(2)
        VertexArray(3).Color = MapData(X, Y).Engine_Light(3)
        
        ' #### WATER
        If DX8_HaveWater And Settings.Water_Effect = True Then
        
            If (TileEngine_Is_Water(X, Y) Or TileEngine_Is_Magma(X, Y)) Then
        
            Dim POLYGON_IGNORE_TOP      As Byte
            Dim POLYGON_IGNORE_LOWER    As Byte
                    
            POLYGON_IGNORE_LOWER = 0
            POLYGON_IGNORE_TOP = 0
                    
            If (TileEngine_Is_Water(X, Y - 1) Or TileEngine_Is_Magma(X, Y - 1)) = False Then POLYGON_IGNORE_TOP = 1
            If (TileEngine_Is_Water(X, Y + 1) Or TileEngine_Is_Magma(X, Y + 1)) = False Then POLYGON_IGNORE_LOWER = 1
                
                If X Mod 2 = 0 Then
                    
                    If Y Mod 2 = 0 Then
                        If POLYGON_IGNORE_TOP <> 1 Then
                            VertexArray(0).sY = VertexArray(0).sY - Val(polygonCount(0))
                            VertexArray(1).sY = VertexArray(1).sY + Val(polygonCount(0))
                        End If
                        If POLYGON_IGNORE_LOWER <> 1 Then
                            VertexArray(2).sY = VertexArray(2).sY + Val(polygonCount(1))
                            VertexArray(3).sY = VertexArray(3).sY - Val(polygonCount(1))
                        End If
                    Else
                        If POLYGON_IGNORE_TOP <> 1 Then
                            VertexArray(0).sY = VertexArray(0).sY + Val(polygonCount(1))
                            VertexArray(1).sY = VertexArray(1).sY - Val(polygonCount(1))
                        End If
                        If POLYGON_IGNORE_LOWER <> 1 Then
                            VertexArray(2).sY = VertexArray(2).sY - Val(polygonCount(0))
                            VertexArray(3).sY = VertexArray(3).sY + Val(polygonCount(0))
                        End If
                               
                    End If
                           
                ElseIf X Mod 2 = 1 Then
                       
                    If Y Mod 2 = 0 Then
                        If POLYGON_IGNORE_TOP <> 1 Then
                            VertexArray(0).sY = VertexArray(0).sY + Val(polygonCount(0))
                            VertexArray(1).sY = VertexArray(1).sY - Val(polygonCount(0))
                        End If
                        If POLYGON_IGNORE_LOWER <> 1 Then
                            VertexArray(2).sY = VertexArray(2).sY - Val(polygonCount(1))
                            VertexArray(3).sY = VertexArray(3).sY + Val(polygonCount(1))
                        End If
                    Else
                        If POLYGON_IGNORE_TOP <> 1 Then
                            VertexArray(0).sY = VertexArray(0).sY - Val(polygonCount(1))
                            VertexArray(1).sY = VertexArray(1).sY + Val(polygonCount(1))
                        End If
                               
                        If POLYGON_IGNORE_LOWER <> 1 Then
                            VertexArray(2).sY = VertexArray(2).sY + Val(polygonCount(0))
                            VertexArray(3).sY = VertexArray(3).sY - Val(polygonCount(0))
                        End If
                    End If
                    
                End If
                
            End If
        End If
        
        DirectDevice.SetTexture 0, D3DTextures.Texture
        
        DirectDevice.DrawIndexedPrimitiveUP D3DPT_TRIANGLESTRIP, 0, 4, 2, _
                indexList(0), D3DFMT_INDEX16, _
                VertexArray(0), Len(VertexArray(0))
    End With
    
End If

End Sub

Public Sub Engine_Render_Line(x1 As Single, Y1 As Single, x2 As Single, Y2 As Single, Optional Color As Long = -1, Optional Color2 As Long = -1)
'**************************************************************
'Author: Standelf
'Last Modify Date: 18/10/2012
'**************************************************************
    
On Error GoTo Error
Dim Vertex(1)   As D3DTLVERTEX

    Vertex(0) = Engine_Create_D3DTLVertex(x1, Y1, 0, 1, Color, 0, 0, 0)
    Vertex(1) = Engine_Create_D3DTLVertex(x2, Y2, 0, 1, Color2, 0, 0, 0)

    DirectDevice.SetVertexShader D3DFVF_XYZRHW Or D3DFVF_TEX1 Or D3DFVF_DIFFUSE Or D3DFVF_SPECULAR
    DirectDevice.SetTexture 0, Nothing
    DirectDevice.DrawPrimitiveUP D3DPT_LINELIST, 1, Vertex(0), Len(Vertex(0))
Exit Sub

Error:
    Call Log_Engine("Error in Engine_Render_Line, " & Err.Description & " (" & Err.Number & ")")
End Sub

Public Sub Engine_Render_Outline_Box(ByVal X As Single, ByVal Y As Single, ByVal Width As Integer, ByVal Height As Integer, Color As Long, Optional ByVal Border As Integer = 1)
'***************************************************
'Author: Ezequiel Juárez (Standelf)
'Last Modification: 24/10/2012
'Proyecto-AO | Draw Outline Box
'***************************************************

    Call Engine_Render_Line(X, Y, X + Width, Y, Color, Color)
    Call Engine_Render_Line(X, Y + Height, X + Width, Y + Height, Color, Color)
    
    Call Engine_Render_Line(X, Y, X, Y + Height, Color, Color)
    Call Engine_Render_Line(X + Width, Y, X + Width, Y + Height, Color, Color)
    
End Sub

Public Sub Engine_Render_Point(x1 As Single, Y1 As Single, Optional Color As Long = -1)
'**************************************************************
'Author: Standelf
'Last Modify Date: 18/10/2012
'**************************************************************
    
On Error GoTo Error
Dim Vertex(0)   As D3DTLVERTEX

    Vertex(0) = Engine_Create_D3DTLVertex(x1, Y1, 0, 1, Color, 0, 0, 0)


    DirectDevice.SetVertexShader D3DFVF_XYZRHW Or D3DFVF_TEX1 Or D3DFVF_DIFFUSE Or D3DFVF_SPECULAR
    DirectDevice.SetTexture 0, Nothing
    DirectDevice.DrawPrimitiveUP D3DPT_POINTLIST, 1, Vertex(0), Len(Vertex(0))
Exit Sub

Error:
    Call Log_Engine("Error in Engine_Render_Point, " & Err.Description & " (" & Err.Number & ")")
End Sub

Public Function Engine_Render_Radial_Map(ByVal MinY As Integer, MaxY As Integer, MinX As Integer, MaxX As Integer)
'**************************************************************
'Author: Standelf
'Last Modify Date: 18/10/2012
'**************************************************************

On Error GoTo Error
    Dim sX      As Byte
    Dim sY      As Byte
    
        For sX = MinX To MaxX
            For sY = MinY To MaxY
                If MapData(sX, sY).Blocked Then Engine_Render_Fill_Box sX, sY, 1, 1, D3DColorARGB(150, 0, 0, 0)
            Next sY
        Next sX
        
        Engine_Render_Fill_Box UserPos.X - 2, UserPos.Y - 2, 4, 4, D3DColorARGB(150, 150, 0, 0)

Error:
    Call Log_Engine("Error in Engine_Render_Radial_Map, " & Err.Description & " (" & Err.Number & ")")
End Function

Sub Engine_Reset_Tile_Vertex()
'**************************************************************
'Author: Standelf
'Last Modify Date: 30/01/2011
'**************************************************************

On Error GoTo Error
Dim X As Integer, Y As Integer, i As Byte

For X = MinXBorder To MaxXBorder
    For Y = MinYBorder To MaxYBorder
        For i = 0 To 3
            MapData(X, Y).Vertex_Offset(i) = 0
            
            'Set Ambient Vertex
            CurMapAmbient.MapBlocks(X, Y).Vertex_Offset(i) = MapData(X, Y).Vertex_Offset(i)

        Next i
    Next Y
Next X

Error:
    Call Log_Engine("Error in Engine_Reset_Tile_Vertex, " & Err.Description & " (" & Err.Number & ")")
End Sub

Public Function Engine_Save_BackBuffer()
'**************************************************************
'Author: Standelf
'Last Modify Date: 18/10/2012
'**************************************************************

Dim PAL As PALETTEENTRY
Dim FileName As String

    FileName = App.Path & "\screen1.bmp"
    
    PAL.Blue = 255
    PAL.Green = 255
    PAL.Red = 255

    DirectD3D8.SaveSurfaceToFile FileName, D3DXIFF_BMP, DirectDevice.GetBackBuffer(0, D3DBACKBUFFER_TYPE_MONO), PAL, MainScreenRect
End Function

Public Function Engine_Set_ARGB(Alpha As Integer, Red As Integer, Green As Integer, Blue As Integer) As D3DCOLORVALUE
'**************************************************************
'Author: Standelf
'Last Modify Date: 18/10/2012
'**************************************************************

If Alpha < 0 Then Alpha = 0 Else If Alpha > 255 Then Alpha = 255
If Red < 0 Then Red = 0 Else If Red > 255 Then Red = 255
If Green < 0 Then Green = 0 Else If Green > 255 Then Green = 255
If Blue < 0 Then Blue = 0 Else If Blue > 255 Then Blue = 255

Engine_Set_ARGB.a = Alpha
Engine_Set_ARGB.r = Red
Engine_Set_ARGB.g = Green
Engine_Set_ARGB.b = Blue

End Function

Public Sub Engine_Set_BaseSpeed(ByVal BaseSpeed As Single)
'**************************************************************
'Author: Standelf
'Last Modify Date: 29/12/2010
'**************************************************************

    Engine_BaseSpeed = BaseSpeed
    
End Sub

Public Sub Engine_Set_Texture_Filter(DirectDevice As Direct3DDevice8, Caps As D3DCAPS8, Stage As Long, Filter As Texture_Filtering, Optional MaxAnisotropy As Long = 2)
'**************************************************************
'Author: •Parra
'Last Modify Date: 18/10/2012
'Modify by Standelf, Fix MIPFILTER
'**************************************************************

  Select Case Filter
    
    Case TexFilter_None
      DirectDevice.SetTextureStageState Stage, D3DTSS_MAGFILTER, D3DTEXF_NONE
      DirectDevice.SetTextureStageState Stage, D3DTSS_MINFILTER, D3DTEXF_NONE
    Case TexFilter_Bilinear
      DirectDevice.SetTextureStageState Stage, D3DTSS_MAGFILTER, D3DTEXF_POINT
      DirectDevice.SetTextureStageState Stage, D3DTSS_MINFILTER, D3DTEXF_LINEAR
    Case TexFilter_Trilinear
      DirectDevice.SetTextureStageState Stage, D3DTSS_MAGFILTER, D3DTEXF_LINEAR
      DirectDevice.SetTextureStageState Stage, D3DTSS_MINFILTER, D3DTEXF_LINEAR
    Case TexFilter_Anisotropic
      DirectDevice.SetTextureStageState Stage, D3DTSS_MAGFILTER, D3DTEXF_LINEAR
      
      If Caps.MaxAnisotropy >= MaxAnisotropy Then
        DirectDevice.SetTextureStageState Stage, D3DTSS_MINFILTER, D3DTEXF_ANISOTROPIC
        DirectDevice.SetTextureStageState Stage, D3DTSS_MAXANISOTROPY, MaxAnisotropy
      ElseIf Caps.MaxAnisotropy >= 2 Then
        DirectDevice.SetTextureStageState Stage, D3DTSS_MINFILTER, D3DTEXF_ANISOTROPIC
        DirectDevice.SetTextureStageState Stage, D3DTSS_MAXANISOTROPY, Caps.MaxAnisotropy
      Else
        DirectDevice.SetTextureStageState Stage, D3DTSS_MINFILTER, D3DTEXF_LINEAR
      End If
      
  End Select
End Sub

Public Sub Engine_Set_TileBuffer(ByVal setEngine_TileBuffer As Single)
'**************************************************************
'Author: Standelf
'Last Modify Date: 30/12/2010
'**************************************************************

    Engine_TileBuffer = setEngine_TileBuffer

End Sub

Public Function Engine_Set_WireFrame(ByVal value As Boolean)
'**************************************************************
'Author: Standelf
'Last Modify Date: 18/10/2012
'**************************************************************

    If value Then
        DirectDevice.SetRenderState D3DRS_FILLMODE, D3DFILL_WIREFRAME
    Else
        DirectDevice.SetRenderState D3DRS_FILLMODE, D3DFILL_SOLID
    End If
End Function

Public Function Engine_TPtoSPX(ByVal X As Byte) As Long
'************************************************************
'Tile Position to Screen Position
'Takes the tile position and returns the pixel location on the screen
'More info: http://www.vbgore.com/GameClient.TileEngine.Engine_TPtoSPX
'************************************************************

    Engine_TPtoSPX = Engine_PixelPosX(X - ((UserPos.X - HalfWindowTileWidth) - Engine_Get_TileBuffer)) + OffsetCounterX - 272 + ((10 - Settings.BufferSize) * 32)
    
End Function

Public Function Engine_TPtoSPY(ByVal Y As Byte) As Long
'************************************************************
'Tile Position to Screen Position
'Takes the tile position and returns the pixel location on the screen
'More info: http://www.vbgore.com/GameClient.TileEngine.Engine_TPtoSPY
'************************************************************

    Engine_TPtoSPY = Engine_PixelPosY(Y - ((UserPos.Y - HalfWindowTileHeight) - Engine_Get_TileBuffer)) + OffsetCounterY - 272 + ((10 - Settings.BufferSize) * 32)
    
End Function

' ####################################################################################
' ####################################################################################
' ########          CÁLCULOS
' ####################################################################################
' ####################################################################################

Public Sub Engine_Water_Effect_Update()

    If Settings.Water_Effect And DX8_HaveWater Then
        'Set Timer for gradual effect
        If GetTickCount - WATER_TICKCOUNT > 8 Then
            Dim waterHeight As Integer: waterHeight = 4
            Static polygon_one_down As Single
        
            WATER_TICKCOUNT = GetTickCount
        
            If polygon_one_down = 0 Then
                polygonCount(0) = polygonCount(0) + (4 * 0.042)
                If polygonCount(0) >= waterHeight Then
                    polygonCount(0) = waterHeight
                    polygon_one_down = 1
                End If
            Else
                polygonCount(0) = polygonCount(0) - (4 * 0.042)
                If polygonCount(0) <= -waterHeight Then
                    polygonCount(0) = -waterHeight
                    polygon_one_down = 0
                End If
            End If
                      
            polygonCount(1) = polygonCount(0)
               
            If polygon_one_down = 0 Then
                polygonCount(1) = polygonCount(1) + (waterHeight * 0.5)
                If polygonCount(1) >= waterHeight Then polygonCount(1) = waterHeight - (polygonCount(1) - waterHeight)
            Else
                polygonCount(1) = polygonCount(1) - (waterHeight * 0.5)
                If polygonCount(1) <= -waterHeight Then polygonCount(1) = -waterHeight + Abs(polygonCount(1) + waterHeight)
            End If
               
            polygonCount(1) = -polygonCount(1)
        End If
    End If

End Sub

