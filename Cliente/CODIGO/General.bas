Attribute VB_Name = "Mod_General"
'***************************************************
'Author: Ezequiel Juárez (Standelf)
'Last Modification: 11/01/2011
'Blisse-AO | This module contains the general functions _
    of the application.
'***************************************************

Option Explicit

' Set Directory
Public Type DIRS
    Graphics    As String
    Sounds      As String
    MP3         As String
    GUI         As String
    Maps        As String
    Ambient     As String
    Bin         As String
    Game        As String
End Type

Public Resources As DIRS

Private Declare Function GetActiveWindow Lib "user32" () As Long
Private Declare Sub keybd_event Lib "user32" (ByVal bVk As Byte, ByVal bScan As Byte, ByVal dwFlags As Long, ByVal dwExtraInfo As Long)
Private Const KEYEVENTF_KEYUP = &H2
Private Const KEYEVENTF_EXTENDEDKEY = &H1
     
Public Function General_Press_Key(ByVal Key As Byte, Optional ByVal Repeat As Byte = 1)
For Repeat = 1 To Repeat
    Call keybd_event(Key, 0, 0, 0)
    Call keybd_event(Key, 0, KEYEVENTF_KEYUP, 0)
    DoEvents
Next Repeat
End Function

Public Function General_Is_App_Active() As Boolean
    General_Is_App_Active = (GetActiveWindow <> 0)
End Function

Public Function General_Get_Graphics_Path() As String
    General_Get_Graphics_Path = App.Path & "\Graficos\"
End Function

Public Function General_Get_Random_Number(ByVal LowerBound As Long, ByVal UpperBound As Long) As Long
    General_Get_Random_Number = (UpperBound - LowerBound) * Rnd + LowerBound
End Function

Public Function General_Get_RawName(ByRef sName As String) As String
'***************************************************
'Author: ZaMa
'Last Modify Date: 13/01/2010
'Last Modified By: -
'Returns the char name without the clan name (if it has it).
'***************************************************

    Dim Pos As Integer
    
    Pos = InStr(1, sName, "<")
    
    If Pos > 0 Then
        General_Get_RawName = Trim(Left(sName, Pos - 1))
    Else
        General_Get_RawName = sName
    End If

End Function

Sub General_Add_to_RichTextBox(ByRef RichTextBox As RichTextBox, ByVal Text As String, Optional ByVal Red As Integer = -1, Optional ByVal Green As Integer, Optional ByVal Blue As Integer, Optional ByVal bold As Boolean = False, Optional ByVal italic As Boolean = False, Optional ByVal bCrLf As Boolean = True)
'***************************************************
'Author: Unknown
'Last Modify Date: 13/01/2010
'Last Modified By: -
'Returns the char name without the clan name (if it has it).
'***************************************************
    With RichTextBox
        If Len(.Text) > 1000 Then
            .SelStart = InStr(1, .Text, vbCrLf) + 1
            .SelLength = Len(.Text) - .SelStart + 2
            .TextRTF = .SelRTF
        End If
        
        .SelStart = Len(.Text)
        .SelLength = 0
        .SelBold = bold
        .SelItalic = italic
        
        If Not Red = -1 Then .SelColor = RGB(Red, Green, Blue)
        
        If bCrLf And Len(.Text) > 0 Then Text = vbCrLf & Text
        .SelText = Text
    End With
End Sub

Function General_Is_Ascii_Valid(ByVal cad As String) As Boolean
    Dim car As Byte
    Dim i As Long
    
    cad = LCase$(cad)
    
    For i = 1 To Len(cad)
        car = Asc(mid$(cad, i, 1))
        
        If ((car < 97 Or car > 122) Or car = Asc("º")) And (car <> 255) And (car <> 32) Then
            Exit Function
        End If
    Next i
    
    General_Is_Ascii_Valid = True
End Function

Function General_Check_AccountData(ByVal CheckEMail As Boolean, ByVal CheckPass As Boolean) As Boolean
'Validamos los datos de la cuenta (TonchitoZ)
    Dim LoopC As Long
    Dim CharAscii As Integer
    
If CheckEMail And Cuenta.Email = "" Then
    MsgBox ("Debes introducir una cuenta de E-Mail.")
    Exit Function
    
    If Not General_Check_Mail_String(Cuenta.Email) Then
        MsgBox "Direccion de mail invalida."
        Exit Function
    End If
End If

If Cuenta.Pass = "" Then
    MsgBox ("Debes introducir una contraseña válida.")
    Exit Function
End If

For LoopC = 1 To Len(Cuenta.Pass)
    CharAscii = Asc(mid$(Cuenta.Pass, LoopC, 1))
    If Not General_Is_Legal_CharacterPass(CharAscii) Then
        MsgBox ("Password inválido. El caractér " & Chr$(CharAscii) & " no está permitido.")
        Exit Function
    End If
Next LoopC

If CheckPass Then
    If Cuenta.Pass <> FrmNewCuenta.TRepPass Then
        MsgBox ("Las contraseñas no coinciden.")
        Exit Function
    End If
End If

If Cuenta.name = "" Then
    MsgBox ("Debes introducir un nombre para tu cuenta.")
    Exit Function
End If

If Len(Cuenta.name) < 6 Then
    MsgBox ("El nombre de la cuenta debe contener como mínimo 6 carácteres.")
    Exit Function
End If

If Right$(Cuenta.name, 1) = " " Then
    UserName = RTrim$(Cuenta.name)
    MsgBox "Nombre de cuenta invalido, se han removido los espacios al final del nombre."
End If
For LoopC = 1 To Len(Cuenta.name)
    CharAscii = Asc(mid$(Cuenta.name, LoopC, 1))
    If Not General_Is_Legal_Character(CharAscii) Then
        MsgBox ("Nombre inválido. El caractér " & Chr$(CharAscii) & " no está permitido.")
        Exit Function
    End If
Next LoopC
    
General_Check_AccountData = True

End Function

Function General_Check_User_Data(ByVal CheckEMail As Boolean) As Boolean
    'Validamos los datos del user
    Dim LoopC As Long
    Dim CharAscii As Integer
    
    For LoopC = 1 To Len(UserPassword)
        CharAscii = Asc(mid$(UserPassword, LoopC, 1))
        If Not General_Is_Legal_Character(CharAscii) Then
            MsgBox ("Password inválido. El caractér " & Chr$(CharAscii) & " no está permitido.")
            Exit Function
        End If
    Next LoopC
    
    If UserName = "" Then
        MsgBox ("Ingrese un nombre de personaje.")
        Exit Function
    End If
    
    If Len(UserName) > 30 Then
        MsgBox ("El nombre debe tener menos de 30 letras.")
        Exit Function
    End If
    
    For LoopC = 1 To Len(UserName)
        CharAscii = Asc(mid$(UserName, LoopC, 1))
        If Not General_Is_Legal_Character(CharAscii) Then
            MsgBox ("Nombre inválido. El caractér " & Chr$(CharAscii) & " no está permitido.")
            Exit Function
        End If
    Next LoopC
    
    General_Check_User_Data = True
End Function

Sub General_Unload_Forms()
On Error Resume Next

    Dim MyForm As Form
    
    For Each MyForm In Forms
        Unload MyForm
    Next
End Sub

Function General_Is_Legal_CharacterPass(ByVal KeyAscii As Integer) As Boolean
    'if backspace allow
    If KeyAscii = 8 Then
        General_Is_Legal_CharacterPass = True
        Exit Function
    End If
    
    'Only allow space, numbers, letters and special characters
    If KeyAscii >= 48 And KeyAscii <= 57 Or KeyAscii >= 65 And KeyAscii <= 90 Or KeyAscii >= 97 And KeyAscii <= 122 Then
        General_Is_Legal_CharacterPass = True
        Exit Function
    End If

    'else everything is cool
    General_Is_Legal_CharacterPass = False
End Function

Function General_Is_Legal_Character(ByVal KeyAscii As Integer) As Boolean
'*****************************************************************
'Only allow characters that are Win 95 filename compatible
'*****************************************************************
    'if backspace allow
    If KeyAscii = 8 Then
        General_Is_Legal_Character = True
        Exit Function
    End If
    
    'Only allow space, numbers, letters and special characters
    If KeyAscii < 32 Or KeyAscii = 44 Then
        Exit Function
    End If
    
    If KeyAscii > 126 Then
        Exit Function
    End If
    
    'Check for bad special characters in between
    If KeyAscii = 34 Or KeyAscii = 42 Or KeyAscii = 47 Or KeyAscii = 58 Or KeyAscii = 60 Or KeyAscii = 62 Or KeyAscii = 63 Or KeyAscii = 92 Or KeyAscii = 124 Then
        Exit Function
    End If
    
    'else everything is cool
    General_Is_Legal_Character = True
End Function

Sub General_Set_Connected()
'*****************************************************************
'Sets the client to "Connect" mode
'*****************************************************************
    'Set Connected
    Connected = True
    
    'Unload the connect form
    Unload frmCrearPersonaje
    Unload frmConnect
    
    frmMain.lblName.Caption = UserName
    frmMain.Label5.Caption = UserLvl
    Call Reset_Party

    'Load main form
    frmMain.Visible = True
    AlphaPres = 255
    
    Call frmMain.ControlSM(eSMType.mSpells, False)
    Call frmMain.ControlSM(eSMType.mWork, False)
    Call frmMain.ControlSM(eSMType.sItem, Items_Seg)
    
    If Settings.MiniMap Then
        Call General_Pixel_Map_Render
    End If
    
    frmMain.SetFocus

End Sub


Private Sub General_Check_Keys()
'*****************************************************************
'Checks keys and respond
'*****************************************************************
    'No input allowed while Argentum is not the active window
    If Not General_Is_App_Active() Then Exit Sub
    
    'No walking when in commerce or banking.
    If Comerciando Then Exit Sub
    
    If EnDuelo Then Exit Sub
    
    If UserMeditar Then Exit Sub
    
    'No walking while writting in the forum.
    If MirandoForo Then Exit Sub
    
    'If game is paused, abort movement.
    If pausa Then Exit Sub
    
    'TODO: Debería informarle por consola?
    If Traveling Then Exit Sub

    'Don't allow any these keys during movement..
    If UserMoving = 0 Then
        If Not UserEstupido Then
            'Move Up
            If GetKeyState(CustomKeys.BindedKey(eKeyType.mKeyUp)) < 0 Then
                If frmMain.TrainingMacro.Enabled Then frmMain.DesactivarMacroHechizos
                Call MoveTo(NORTH)
                
                frmMain.lblMapPosX.Caption = UserPos.X
                frmMain.lblMapPosY.Caption = UserPos.Y
                Exit Sub
            End If
            
            'Move Right
            If GetKeyState(CustomKeys.BindedKey(eKeyType.mKeyRight)) < 0 Then
                If frmMain.TrainingMacro.Enabled Then frmMain.DesactivarMacroHechizos
                Call MoveTo(EAST)
                frmMain.lblMapPosX.Caption = UserPos.X
                frmMain.lblMapPosY.Caption = UserPos.Y
                Exit Sub
            End If
        
            'Move down
            If GetKeyState(CustomKeys.BindedKey(eKeyType.mKeyDown)) < 0 Then
                If frmMain.TrainingMacro.Enabled Then frmMain.DesactivarMacroHechizos
                Call MoveTo(SOUTH)
                frmMain.lblMapPosX.Caption = UserPos.X
                frmMain.lblMapPosY.Caption = UserPos.Y
                Exit Sub
            End If
        
            'Move left
            If GetKeyState(CustomKeys.BindedKey(eKeyType.mKeyLeft)) < 0 Then
                If frmMain.TrainingMacro.Enabled Then frmMain.DesactivarMacroHechizos
                Call MoveTo(WEST)
                frmMain.lblMapPosX.Caption = UserPos.X
                frmMain.lblMapPosY.Caption = UserPos.Y
                Exit Sub
            End If
            
            '   We haven't moved - Update 3D sounds!
            Call Audio.MoveListener(UserPos.X, UserPos.Y)
        Else
            Dim kp As Boolean
            kp = (GetKeyState(CustomKeys.BindedKey(eKeyType.mKeyUp)) < 0) Or _
                GetKeyState(CustomKeys.BindedKey(eKeyType.mKeyRight)) < 0 Or _
                GetKeyState(CustomKeys.BindedKey(eKeyType.mKeyDown)) < 0 Or _
                GetKeyState(CustomKeys.BindedKey(eKeyType.mKeyLeft)) < 0
            
            If kp Then
                Call RandomMove
            Else
                '   We haven't moved - Update 3D sounds!
                Call Audio.MoveListener(UserPos.X, UserPos.Y)
            End If
            
            If frmMain.TrainingMacro.Enabled Then frmMain.DesactivarMacroHechizos
                frmMain.lblMapPosX.Caption = UserPos.X
                frmMain.lblMapPosY.Caption = UserPos.Y
            If Settings.MiniMap Then Call General_Pixel_Map_Set_Area
        End If
    End If
    

End Sub

Sub General_Load_MapData(ByVal Map As Integer)
    Dim Y As Long
    Dim X As Long
    Dim tempint As Integer
    Dim ByFlags As Byte
    Dim handle As Integer
    Dim LargestTileSize As Long
    handle = FreeFile()

    DX8_HaveWater = False
    
    Open Resources.Maps & "mapa" & Map & ".map" For Binary As handle
    Seek handle, 1
            
    'map Header
    Get handle, , MapInfo.MapVersion
    Get handle, , MiCabecera
    Get handle, , tempint
    Get handle, , tempint
    Get handle, , tempint
    Get handle, , tempint

    'Load arrays
    For Y = YMinMapSize To YMaxMapSize
        For X = XMinMapSize To XMaxMapSize
            Get handle, , ByFlags
            
            MapData(X, Y).Blocked = (ByFlags And 1)
            
            Get handle, , MapData(X, Y).Graphic(1).GrhIndex
            InitGrh MapData(X, Y).Graphic(1), MapData(X, Y).Graphic(1).GrhIndex
            
            'Layer 2 used?
            If ByFlags And 2 Then
                Get handle, , MapData(X, Y).Graphic(2).GrhIndex
                InitGrh MapData(X, Y).Graphic(2), MapData(X, Y).Graphic(2).GrhIndex
            Else
                MapData(X, Y).Graphic(2).GrhIndex = 0
            End If
                
            'Layer 3 used?
            If ByFlags And 4 Then
                Get handle, , MapData(X, Y).Graphic(3).GrhIndex
                InitGrh MapData(X, Y).Graphic(3), MapData(X, Y).Graphic(3).GrhIndex
            Else
                MapData(X, Y).Graphic(3).GrhIndex = 0
            End If
                
            'Layer 4 used?
            If ByFlags And 8 Then
                Get handle, , MapData(X, Y).Graphic(4).GrhIndex
                InitGrh MapData(X, Y).Graphic(4), MapData(X, Y).Graphic(4).GrhIndex
            Else
                MapData(X, Y).Graphic(4).GrhIndex = 0
            End If
            
            'Trigger used?
            If ByFlags And 16 Then
                Get handle, , MapData(X, Y).Trigger
            Else
                MapData(X, Y).Trigger = 0
            End If
            
            'Erase NPCs
            If MapData(X, Y).CharIndex > 0 Then
                Call EraseChar(MapData(X, Y).CharIndex)
            End If
            
            'Erase OBJs
            MapData(X, Y).ObjGrh.GrhIndex = 0
            MapData(X, Y).OBJInfo.name = vbNullString

            'Reset the Vertex
            Dim Vertex As Long
            
            For Vertex = 0 To 3
                MapData(X, Y).Vertex_Offset(Vertex) = 0
            Next Vertex
            
            'Erase Lights
            Call Engine_D3DColor_To_RGB_List(MapData(X, Y).Engine_Light(), Estado_Actual) 'Standelf, Light & Meteo Engine
            
            
            'Set Water
            If TileEngine_Is_Water(X, Y) = True Then
                DX8_HaveWater = True
            End If
            If TileEngine_Is_Magma(X, Y) = True Then
                DX8_HaveWater = True
            End If
        Next X
    Next Y
    
    Close handle
    
    Call LightRemoveAll
    
    '   Erase particle effects
    ReDim Effect(1 To NumEffects)
    
    MapInfo.name = UserMapName
    MapInfo.Music = ""
    Ambient_Init Map
    
    If Settings.MiniMap Then Call General_Pixel_Map_Render
    If Settings.MiniMap Then Call General_Pixel_Map_Set_Area
    

    If UserMap = 120 Then Effect_Waterfall_Begin Engine_TPtoSPX(8), Engine_TPtoSPY(3), 1, 800
    
End Sub

Function General_Get_ReadField(ByVal Pos As Integer, ByRef Text As String, ByVal SepASCII As Byte) As String
'*****************************************************************
'Gets a field from a delimited string
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 11/15/2004
'*****************************************************************
    Dim i As Long
    Dim lastPos As Long
    Dim CurrentPos As Long
    Dim delimiter As String * 1
    
    delimiter = Chr$(SepASCII)
    
    For i = 1 To Pos
        lastPos = CurrentPos
        CurrentPos = InStr(lastPos + 1, Text, delimiter, vbBinaryCompare)
    Next i
    
    If CurrentPos = 0 Then
        General_Get_ReadField = mid$(Text, lastPos + 1, Len(Text) - lastPos)
    Else
        General_Get_ReadField = mid$(Text, lastPos + 1, CurrentPos - lastPos - 1)
    End If
End Function

Function General_Get_Field_Count(ByRef Text As String, ByVal SepASCII As Byte) As Long
'*****************************************************************
'Gets the number of fields in a delimited string
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 07/29/2007
'*****************************************************************
    Dim Count As Long
    Dim curPos As Long
    Dim delimiter As String * 1
    
    If LenB(Text) = 0 Then Exit Function
    
    delimiter = Chr$(SepASCII)
    
    curPos = 0
    
    Do
        curPos = InStr(curPos + 1, Text, delimiter)
        Count = Count + 1
    Loop While curPos <> 0
    
    General_Get_Field_Count = Count
End Function

Function General_File_Exist(ByVal File As String, ByVal FileType As VbFileAttribute) As Boolean
    General_File_Exist = (Dir$(File, FileType) <> "")
End Function

Sub Main()
    InitCommonControls
    Randomize GetTickCount
    
    With Resources
        .Game = App.Path & "\"
        .Ambient = .Game & "Resources\Ambient\"
        .Bin = .Game & "Resources\Bin\"
        .Graphics = .Game & "Resources\Graficos\"
        .GUI = .Game & "Resources\GUI\"
        .Maps = .Game & "Resources\Mapas\"
        .MP3 = .Game & "Resources\MP3\"
        .Sounds = .Game & "Resources\wav\"
    End With
    
    '     Logs System Init
    Call Init_Logs
    
    '     Cargando
    frmCargando.Show
    frmCargando.Refresh
    
    DoEvents
    
    '     Set Resolution; Cambiar esto por DX8
    Call SetResolution
    
    '     Blisse-Security
    #If SeguridadBlisse = 1 Then
        Call Init_Security
    #End If
    
    '     Cargamos la configuración del usuario
    Call Settings_Init
    
    '     Cargamos Variables feas
    Call General_Init_Names
    Call Mod_Protocol.InitFonts
    
    '     Iniciamos el Engine de DirectX 8
    If Not Engine_DirectX8_Init Then
        Call General_Close_Client
    End If
          
    
    '     Tile Engine
    If Not InitTileEngine(frmMain.hWnd, 32, 32, 8, 8) Then
        Call General_Close_Client
    End If
    
    '     Inventario
    Call Inventario.Initialize(DirectD3D8, frmMain.PicInv, MAX_INVENTORY_SLOTS)

    '     Iniciamos el Engine de Sonido
    Call Audio.Initialize(DirectX, frmMain.hWnd, Resources.Sounds, Resources.Sounds)
    
    '     Seteamos las variables de sonido
    Audio.MusicActivated = Settings.Musica
    Audio.SoundActivated = Settings.Sonido
    Audio.SoundVolume = Settings.SoundVolume
    Audio.MusicVolume = Settings.MusicVolume
    Audio.SoundEffectsActivated = Settings.Sonido3D
    
    If Settings.Musica Then _
        Play_MP3 eMP3.MainMenu
    '     Quitamos el Cargando y mostramos el conectar
    Unload frmCargando
    frmConnect.Visible = True
    AlphaPres = 255
    set_GUI_Efect
    CurMapAmbient.Fog = 100
    
    '     Inicialización de variables globales
    prgRun = True
    pausa = False
    
    '     Set the intervals of timers
    Call MainTimer.SetInterval(TimersIndex.Attack, INT_ATTACK)
    Call MainTimer.SetInterval(TimersIndex.Work, INT_WORK)
    Call MainTimer.SetInterval(TimersIndex.UseItemWithU, INT_USEITEMU)
    Call MainTimer.SetInterval(TimersIndex.UseItemWithDblClick, INT_USEITEMDCK)
    Call MainTimer.SetInterval(TimersIndex.SendRPU, INT_SENTRPU)
    Call MainTimer.SetInterval(TimersIndex.CastSpell, INT_CAST_SPELL)
    Call MainTimer.SetInterval(TimersIndex.Arrows, INT_ARROWS)
    Call MainTimer.SetInterval(TimersIndex.CastAttack, INT_CAST_ATTACK)
    Call MainTimer.SetInterval(TimersIndex.ClandDialog, 1000)
    frmMain.MacroTrabajo.Interval = INT_MACRO_TRABAJO
    frmMain.MacroTrabajo.Enabled = False
    
    '     Init timers
    Call MainTimer.Start(TimersIndex.Attack)
    Call MainTimer.Start(TimersIndex.Work)
    Call MainTimer.Start(TimersIndex.UseItemWithU)
    Call MainTimer.Start(TimersIndex.UseItemWithDblClick)
    Call MainTimer.Start(TimersIndex.SendRPU)
    Call MainTimer.Start(TimersIndex.CastSpell)
    Call MainTimer.Start(TimersIndex.Arrows)
    Call MainTimer.Start(TimersIndex.CastAttack)
    Call MainTimer.Start(TimersIndex.ClandDialog)
    
    Engine_Movement_Speed = 1

    Do While prgRun
        If frmMain.WindowState <> 1 And frmMain.Visible Then
            '   ####    Init Scene
            Engine_BeginScene
                    
            '   ####    Render Scene
            Call ShowNextFrame(frmMain.Top, frmMain.Left, frmMain.MouseX, frmMain.MouseY)
            
            Engine_Render_FPS
            '   ####    End Scene
            Engine_EndScene MainScreenRect, 0
            
            
            If MainTimer.Check(TimersIndex.ClandDialog) Then
                If Not DialogosClanes Is Nothing Then DialogosClanes.PassTimer
            End If
            
            Call RenderSounds
            Call General_Check_Keys
              
            #If SeguridadBlisse = 1 Then
                Updates_Intervals
            #End If
   
            'Render Others
             If frmMain.PicInv.Visible Then
                 Call Inventario.DrawInv
             End If
             
             If frmBancoObj.PicBancoInv.Visible Then
                   Call InvBanco(0).DrawInv
             End If
             If frmBancoObj.PicInv.Visible Then
                 Call InvBanco(1).DrawInv
             End If
             
             If frmComerciar.Visible = True Then
                 Call InvComUsu.DrawInv
                 Call InvComNpc.DrawInv
             End If
            
             If frmComerciarUsu.Visible = True Then
                Call InvOfferComUsu(0).DrawInv
                 Call InvOfferComUsu(1).DrawInv
                 Call InvComUsu.DrawInv
                 Call InvOroComUsu(0).DrawInv
                 Call InvOroComUsu(1).DrawInv
                 Call InvOroComUsu(2).DrawInv
             End If
    
        'FRM CONNECT VISIBLE
        ElseIf frmConnect.WindowState <> 1 And frmConnect.Visible Then
            Call RenderConnect
        End If
        
        
        '   If there is anything to be sent, we send it
        Call FlushBuffer
        
        DoEvents
    Loop
    
    Call General_Close_Client
End Sub




Sub General_Write_Var(ByVal File As String, ByVal Main As String, ByVal Var As String, ByVal value As String)
'*****************************************************************
'Writes a var to a text file
'*****************************************************************
    writeprivateprofilestring Main, Var, value, File
End Sub

Function General_Get_Var(ByVal File As String, ByVal Main As String, ByVal Var As String) As String
'*****************************************************************
'Gets a Var from a text file
'*****************************************************************
    Dim sSpaces As String '   This will hold the input that the program will retrieve
    
    sSpaces = Space$(500) '   This tells the computer how long the longest string can be. If you want, you can change the number 100 to any number you wish
    
    getprivateprofilestring Main, Var, vbNullString, sSpaces, Len(sSpaces), File
    
    General_Get_Var = RTrim$(sSpaces)
    General_Get_Var = Left$(General_Get_Var, Len(General_Get_Var) - 1)
End Function

'[CODE 002]:MatuX
'
'    Función para chequear el email
'
'    Corregida por Maraxus para que reconozca como válidas casillas con puntos antes de la arroba y evitar un chequeo innecesario
Public Function General_Check_Mail_String(ByVal sString As String) As Boolean
On Error GoTo errHnd
    Dim lPos  As Long
    Dim Lx    As Long
    Dim iAsc  As Integer
    
    '1er test: Busca un simbolo @
    lPos = InStr(sString, "@")
    If (lPos <> 0) Then
        '2do test: Busca un simbolo . después de @ + 1
        If Not (InStr(lPos, sString, ".", vbBinaryCompare) > lPos + 1) Then _
            Exit Function
        
        '3er test: Recorre todos los caracteres y los valída
        For Lx = 0 To Len(sString) - 1
            If Not (Lx = (lPos - 1)) Then   'No chequeamos la '@'
                iAsc = Asc(mid$(sString, (Lx + 1), 1))
                If Not General_Is_CMS_Validate_Char(iAsc) Then _
                    Exit Function
            End If
        Next Lx
        
        'Finale
        General_Check_Mail_String = True
    End If
errHnd:
    Call Log_Unknown("Error in function General_Check_Mail_String, " & sString)
End Function

'    Corregida por Maraxus para que reconozca como válidas casillas con puntos antes de la arroba
Private Function General_Is_CMS_Validate_Char(ByVal iAsc As Integer) As Boolean
    General_Is_CMS_Validate_Char = (iAsc >= 48 And iAsc <= 57) Or _
                        (iAsc >= 65 And iAsc <= 90) Or _
                        (iAsc >= 97 And iAsc <= 122) Or _
                        (iAsc = 95) Or (iAsc = 45) Or (iAsc = 46)
End Function

Private Sub General_Init_Names()
'**************************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 11/27/2005
'Inicializa los nombres de razas, ciudades, clases, skills, atributos, etc.
'**************************************************************
    
    ListaRazas(eRaza.Humano) = "Humano"
    ListaRazas(eRaza.Elfo) = "Elfo"
    ListaRazas(eRaza.ElfoOscuro) = "Elfo Oscuro"
    ListaRazas(eRaza.Gnomo) = "Gnomo"
    ListaRazas(eRaza.Enano) = "Enano"

    ListaClases(eClass.Mage) = "Mago"
    ListaClases(eClass.Cleric) = "Clerigo"
    ListaClases(eClass.Warrior) = "Guerrero"
    ListaClases(eClass.Assasin) = "Asesino"
    ListaClases(eClass.Thief) = "Ladron"
    ListaClases(eClass.Bard) = "Bardo"
    ListaClases(eClass.Druid) = "Druida"
    ListaClases(eClass.Bandit) = "Bandido"
    ListaClases(eClass.Paladin) = "Paladin"
    ListaClases(eClass.Hunter) = "Cazador"
    ListaClases(eClass.Worker) = "Trabajador"
    ListaClases(eClass.Pirat) = "Pirata"
    
    SkillsNames(eSkill.Magia) = "Magia"
    SkillsNames(eSkill.Robar) = "Robar"
    SkillsNames(eSkill.Tacticas) = "Evasión en combate"
    SkillsNames(eSkill.Armas) = "Combate cuerpo a cuerpo"
    SkillsNames(eSkill.Meditar) = "Meditar"
    SkillsNames(eSkill.Apuñalar) = "Apuñalar"
    SkillsNames(eSkill.Ocultarse) = "Ocultarse"
    SkillsNames(eSkill.Supervivencia) = "Supervivencia"
    SkillsNames(eSkill.Talar) = "Talar árboles"
    SkillsNames(eSkill.Comerciar) = "Comercio"
    SkillsNames(eSkill.Defensa) = "Defensa con escudos"
    SkillsNames(eSkill.Pesca) = "Pesca"
    SkillsNames(eSkill.Mineria) = "Mineria"
    SkillsNames(eSkill.Carpinteria) = "Carpinteria"
    SkillsNames(eSkill.Herreria) = "Herreria"
    SkillsNames(eSkill.Liderazgo) = "Liderazgo"
    SkillsNames(eSkill.Domar) = "Domar animales"
    SkillsNames(eSkill.Proyectiles) = "Combate a distancia"
    SkillsNames(eSkill.Wrestling) = "Combate sin armas"
    SkillsNames(eSkill.Navegacion) = "Navegacion"

    AtributosNames(eAtributos.Fuerza) = "Fuerza"
    AtributosNames(eAtributos.Agilidad) = "Agilidad"
    AtributosNames(eAtributos.Inteligencia) = "Inteligencia"
    AtributosNames(eAtributos.Carisma) = "Carisma"
    AtributosNames(eAtributos.Constitucion) = "Constitucion"
End Sub

Public Sub General_Clean_Dialogs()
'**************************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 11/27/2005
'Removes all text from the console and dialogs
'**************************************************************
    'Clean console and dialogs
    frmMain.RecTxt.Text = vbNullString
    Call DialogosClanes.RemoveDialogs
    Call Dialogos.RemoveAllDialogs
End Sub

Public Sub General_Close_Client()
'**************************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 11/01/2011
'Frees all used resources, cleans up and leaves
'11/1/2011: Standelf - Order and add features of Security, Settings system and _
    Graphical Engine.
'**************************************************************
    #If SeguridadBlisse = 1 Then
        Call DeInit_Security
        Call Settings_Save
    #End If
    
    EngineRun = False
    
    '     Put the old resolution
    Call Mod_DX8_Resolution.ResetResolution
    
    '     Stop The Graphical Engine
    Call Engine_DirectX8_End
    
    '     Destruimos los objetos públicos creados
    Set CustomKeys = Nothing
    Set SurfaceDB = Nothing
    Set Dialogos = Nothing
    Set DialogosClanes = Nothing
    Set Audio = Nothing
    Set Inventario = Nothing
    Set MainTimer = Nothing
    Set incomingData = Nothing
    Set outgoingData = Nothing
    Erase MapData()
    
    Call General_Unload_Forms
    End
End Sub

Public Function General_Get_GM(CharIndex As Integer) As Boolean
General_Get_GM = False
If CharList(CharIndex).priv >= 1 And CharList(CharIndex).priv <= 5 Or CharList(CharIndex).priv = 25 Then _
    General_Get_GM = True

End Function

Public Function General_Get_TagPosition(ByVal Nick As String) As Integer
Dim buf As Integer
    buf = InStr(Nick, "<")
        If buf > 0 Then
            General_Get_TagPosition = buf
            Exit Function
        End If
    buf = InStr(Nick, "[")
        If buf > 0 Then
            General_Get_TagPosition = buf
            Exit Function
        End If
    General_Get_TagPosition = Len(Nick) + 2
End Function


Public Function General_Get_StrenghtColor() As Long
Dim m As Long
m = 255 / MAXATRIBUTOS
General_Get_StrenghtColor = RGB(255 - (m * UserFuerza), (m * UserFuerza), 0)
End Function

Public Function General_Get_DexterityColor() As Long
Dim m As Long
m = 255 / MAXATRIBUTOS
General_Get_DexterityColor = RGB(255, m * UserAgilidad, 0)
End Function

Public Function General_Is_Anounce(ByVal ForumType As Byte) As Boolean
'***************************************************
'Author: ZaMa
'Last Modification: 22/02/2010
'Returns true if the post is sticky.
'***************************************************
    Select Case ForumType
        Case eForumMsgType.ieCAOS_STICKY
            General_Is_Anounce = True
            
        Case eForumMsgType.ieGENERAL_STICKY
            General_Is_Anounce = True
            
        Case eForumMsgType.ieREAL_STICKY
            General_Is_Anounce = True
            
    End Select
    
End Function

Public Function General_Get_Forum_Alignment(ByVal yForumType As Byte) As Byte
'***************************************************
'Author: ZaMa
'Last Modification: 01/03/2010
'Returns the forum alignment.
'***************************************************
    Select Case yForumType
        Case eForumMsgType.ieCAOS, eForumMsgType.ieCAOS_STICKY
            General_Get_Forum_Alignment = eForumType.ieCAOS
            
        Case eForumMsgType.ieGeneral, eForumMsgType.ieGENERAL_STICKY
            General_Get_Forum_Alignment = eForumType.ieGeneral
            
        Case eForumMsgType.ieREAL, eForumMsgType.ieREAL_STICKY
            General_Get_Forum_Alignment = eForumType.ieREAL
            
    End Select
    
End Function

Public Sub General_Write_Login()
    If EstadoLogin = E_MODO.Normal Then
        Call WriteLoginExistingChar
    ElseIf EstadoLogin = E_MODO.CrearNuevoPj Then
        Call WriteLoginNewChar
    ElseIf EstadoLogin = E_MODO.LoginCuenta Then
        Call WriteLoginCuenta
    ElseIf EstadoLogin = E_MODO.CrearCuenta Then
        Call WriteCrearCuenta
    ElseIf EstadoLogin = E_MODO.BorrarPJ Then
        Call WriteBorrPJCuenta
    End If

    DoEvents

    Call FlushBuffer
End Sub

Public Function General_Valid_String(ByVal tString As String) As Boolean
    If tString = vbNullString Then
        General_Valid_String = False
        Exit Function
    End If
    
    Dim i As Long
        For i = 1 To Len(tString)
            If Asc(mid(tString, i, 1)) <> 32 Then
                General_Valid_String = True
            End If
        Next i
        
        General_Valid_String = False
        Exit Function
End Function

Public Sub General_Drop_X_Y(ByVal X As Byte, ByVal Y As Byte)
'**************************************************************
'Author: Standelf
'Last Modify Date: 11/01/2011
'**************************************************************
    If Items_Seg = True Then
        With FontTypes(FontTypeNames.FONTTYPE_WARNING)
            Call ShowConsoleMsg("Para tirar algún objeto desactiva previamente el seguro de ítems.", .Red, .Green, .Blue, .bold, .italic)
        End With
        Exit Sub
    End If
    
    If (Inventario.SelectedItem > 0 And Inventario.SelectedItem < MAX_INVENTORY_SLOTS + 1) Then
        If GetKeyState(vbKeyShift) < 0 Then
            frmCantidad.IniciarDD X, Y
        Else
            WriteDrop Inventario.SelectedItem, 1, X, Y
        End If
    End If
    
End Sub

Public Function General_Set_GUI(ByVal Interface As String) As IPicture
'**************************************************************
'Author: Standelf
'Last Modify Date: 04/01/2010
'**************************************************************

    If General_File_Exist(Resources.GUI & Interface & ".gif", vbNormal) = True Then
        Set General_Set_GUI = LoadPicture(Resources.GUI & Interface & ".gif")
    ElseIf General_File_Exist(Resources.GUI & Interface & ".jpg", vbNormal) = True Then
        Set General_Set_GUI = LoadPicture(Resources.GUI & Interface & ".jpg")
    Else
        Set General_Set_GUI = Nothing
    End If
    
End Function


Public Sub General_Pixel_Map_Set_Area()
'**************************************************************
'Author: Standelf
'Last Modify Date: ??/??/2010
'**************************************************************
    
    frmMain.UserPosition.Top = UserPos.Y
    frmMain.UserPosition.Left = UserPos.X
    
End Sub

Public Sub General_Pixel_Map_Render()
'**************************************************************
'Author: Rubio93
'Last Modify Date: 11/01/2010
'Last Modify Author: Standelf
'**************************************************************

    frmMain.MiniMap.Cls
    Dim map_x As Byte, map_y As Byte

    For map_y = YMinMapSize To YMaxMapSize
        For map_x = XMinMapSize To XMaxMapSize
        
            If MapData(map_x, map_y).Graphic(1).GrhIndex > 0 Then
                SetPixel frmMain.MiniMap.hDC, map_x, map_y, GrhData(MapData(map_x, map_y).Graphic(1).GrhIndex).MiniMap_color
            End If
            If MapData(map_x, map_y).Graphic(2).GrhIndex > 0 Then
                SetPixel frmMain.MiniMap.hDC, map_x, map_y, GrhData(MapData(map_x, map_y).Graphic(2).GrhIndex).MiniMap_color
            End If
            
        Next map_x
    Next map_y

    frmMain.MiniMap.Refresh
End Sub
