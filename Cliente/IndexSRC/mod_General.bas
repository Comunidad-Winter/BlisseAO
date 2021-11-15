Attribute VB_Name = "mod_General"
Option Explicit

Public GRH_FILE As String
Public BMP_DIRE As String
Public FXS_FILE As String

Public HEA_FILE As String
Public WEA_FILE As String
Public SHI_FILE As String
Public HEL_FILE As String
Public BOD_FILE As String


Public APP_RUN As Boolean

'Color Finder
Public Declare Function GetPixel Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long) As Long
Public Declare Function SetPixel Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long) As Long
    
    Public Declare Function BitBlt Lib "gdi32" ( _
    ByVal hDestDC As Long, _
    ByVal X As Long, _
    ByVal Y As Long, _
    ByVal nWidth As Long, _
    ByVal nHeight As Long, _
    ByVal hSrcDC As Long, _
    ByVal xSrc As Long, _
    ByVal ySrc As Long, _
    ByVal dwRop As Long) As Long


Public Type GRHColor
    Color As Long
End Type

Public GRH_COLORS() As Long
    

Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Private Declare Sub InitCommonControls Lib "comctl32.dll" ()
Public Declare Function GetTickCount Lib "kernel32" () As Long
Private Declare Function writeprivateprofilestring Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationname As String, ByVal lpKeyname As Any, ByVal lpString As String, ByVal lpFileName As String) As Long
Private Declare Function getprivateprofilestring Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationname As String, ByVal lpKeyname As Any, ByVal lpdefault As String, ByVal lpreturnedstring As String, ByVal nSize As Long, ByVal lpFileName As String) As Long



Sub Main()
    GRH_FILE = App.Path & "\Game\Resources\Init\Graficos.ind"
    BMP_DIRE = App.Path & "\Game\Resources\Graficos\"
    FXS_FILE = App.Path & "\Game\Resources\Init\Anims.ind"
    
    HEA_FILE = App.Path & "\Game\Resources\Init\Cabezas.ind"
    WEA_FILE = App.Path & "\Game\Resources\Init\Armas.dat"
    HEL_FILE = App.Path & "\Game\Resources\Init\Cascos.ind"
    SHI_FILE = App.Path & "\Game\Resources\Init\Escudos.dat"
    BOD_FILE = App.Path & "\Game\Resources\Init\Cuerpos.ind"
    
    
    InitCommonControls
    DoEvents
    
    frmCargando.Show
    DoEvents
    Set DX8 = New dx_GFX_Class
    Call DX8.Init(frmMain.Render.hWnd, 384, 384, 32, True, True, False)
    frmMain.Render.Top = 170
    frmMain.Render.Left = 192
    
    DX8_FONT = DX8.FONT_LoadSystemFont("Verdana", 8, False, False, False, False)
    
    LoadGrhData
    CargarFxs
    CargarAnimArmas
    CargarAnimEscudos
    CargarCabezas
    CargarCascos
    CargarCuerpos

    frmMain.Show
    Unload frmCargando
End Sub

Function FileExist(ByVal File As String, ByVal FileType As VbFileAttribute) As Boolean
    FileExist = (Dir$(File, FileType) <> "")
End Function

Public Function ReadField(pos As Integer, Text As String, SepASCII As Integer) As String
'*****************************************************************
'Gets a field from a string
'*****************************************************************

Dim i As Integer
Dim LastPos As Integer
Dim CurChar As String * 1
Dim FieldNum As Integer
Dim Seperator As String

Seperator = Chr(SepASCII)
LastPos = 0
FieldNum = 0

For i = 1 To Len(Text)
    CurChar = Mid(Text, i, 1)
    If CurChar = Seperator Then
        FieldNum = FieldNum + 1
        If FieldNum = pos Then
            ReadField = Mid(Text, LastPos + 1, (InStr(LastPos + 1, Text, Seperator, vbTextCompare) - 1) - (LastPos))
            Exit Function
        End If
        LastPos = i
    End If
Next i
FieldNum = FieldNum + 1

If FieldNum = pos Then
    ReadField = Mid(Text, LastPos + 1)
End If


End Function


Sub WriteVar(ByVal File As String, ByVal Main As String, ByVal Var As String, ByVal value As String)
'*****************************************************************
'Writes a var to a text file
'*****************************************************************
    writeprivateprofilestring Main, Var, value, File
End Sub

Function GetVar(ByVal File As String, ByVal Main As String, ByVal Var As String) As String
'*****************************************************************
'Gets a Var from a text file
'*****************************************************************
    Dim sSpaces As String '   This will hold the input that the program will retrieve
    
    sSpaces = Space$(500) '   This tells the computer how long the longest string can be. If you want, you can change the number 100 to any number you wish
    
    getprivateprofilestring Main, Var, vbNullString, sSpaces, Len(sSpaces), File
    
    GetVar = RTrim$(sSpaces)
    GetVar = Left$(GetVar, Len(GetVar) - 1)
End Function


Public Sub InitGrh(ByRef GRH As GRH, ByVal GrhIndex As Integer, Optional ByVal Started As Byte = 2)
'*****************************************************************
'Sets up a grh. MUST be done before rendering
'*****************************************************************
    GRH.GrhIndex = GrhIndex
    
    If Started = 2 Then
        If GrhData(GRH.GrhIndex).NumFrames > 1 Then
            GRH.Started = 1
        Else
            GRH.Started = 0
        End If
    Else
        If GrhData(GRH.GrhIndex).NumFrames = 1 Then Started = 0
        GRH.Started = Started
    End If
    
    
    If GRH.Started Then
        GRH.Loops = -1
    Else
        GRH.Loops = 0
    End If
    
    GRH.FrameCounter = 1
    GRH.Speed = GrhData(GRH.GrhIndex).Speed

End Sub
