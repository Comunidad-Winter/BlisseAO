VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Cls_DDEx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'*******************************************************
'*Esta Clase fue Creada por Loopzer                    *
'*para facilitar el uso de                             *
'*Directx9 Directx8 y OGL                              *
'*******************************************************
'Pagina=Ddex.com.ar
'Mail=Loopzer@gmail.com
'DDEX Ver. 4.1.0.8

'futura funcion que iniciara la prueba de rendimiento
Private Declare Function IniciarPrueba Lib "DDEXv4.dll" () As Long

Private Declare Function DibujarVertices Lib "DDEXv4.dll" (ByRef puntero As Any, ByRef v As Any, ByVal cantidad As Long, ByVal alphaf As Long) As Long

Private Declare Function DibujarCajaEx Lib "DDEXv4.dll" (ByRef puntero As Any, ByRef r As Any, ByRef color As Any, ByVal angulo As Long) As Long

Private Declare Function DibujarCaja Lib "DDEXv4.dll" (ByRef puntero As Any, ByRef r As Any, ByRef color As Any, ByVal angulo As Long) As Long

Private Declare Function DibujarLineaEx Lib "DDEXv4.dll" (ByRef puntero As Any, ByRef r As Any, ByRef color As Any) As Long

Private Declare Function PreCargarGrafico Lib "DDEXv4.dll" (ByRef puntero As Any, ByVal archivo As Long) As Long

Private Declare Function ZoomPantalla Lib "DDEXv4.dll" (ByRef puntero As Any, ByVal zoom As Long) As Long

Private Declare Function IniciarMotor Lib "DDEXv4.dll" (ByRef puntero As Any, ByVal ventana As Long, ByRef carga As Any, ByRef configuracion As Any) As Long

Private Declare Function RotarPantalla Lib "DDEXv4.dll" (ByRef puntero As Any, ByVal angulo As Long) As Long

Private Declare Function TerminarMotor Lib "DDEXv4.dll" (ByRef puntero As Any) As Long

Private Declare Function DibujarRGrafico Lib "DDEXv4.dll" (ByRef puntero As Any, ByVal archivo As Long, ByRef r As Any, ByVal X As Long, ByVal Y As Long) As Long

Private Declare Function DibujarRGraficoAlpha Lib "DDEXv4.dll" (ByRef puntero As Any, ByVal archivo As Long, ByRef r As Any, ByVal X As Long, ByVal Y As Long, ByRef color As Any) As Long

Private Declare Function DibujarRGraficoEx Lib "DDEXv4.dll" (ByRef puntero As Any, ByVal archivo As Long, ByRef r As Any, ByVal X As Long, ByVal Y As Long, ByRef color As Any, ByVal angulo As Long) As Long

Private Declare Function DibujarTexto Lib "DDEXv4.dll" (ByRef puntero As Any, ByRef Texto As Any, ByVal X As Long, ByVal Y As Long, ByRef color As Any, ByVal fuente As Long) As Long

Private Declare Function CargarFuenteX Lib "DDEXv4.dll" (ByRef puntero As Any, ByRef fuente As Any, ByVal tam As Long, ByVal negrita As Byte, ByVal cursiva As Byte) As Long

Private Declare Function LimpiarMotor Lib "DDEXv4.dll" (ByRef puntero As Any) As Long

Private Declare Function DibujarMotor Lib "DDEXv4.dll" (ByRef puntero As Any) As Long

Private Declare Function DibujarPantallaEx Lib "DDEXv4.dll" (ByRef puntero As Any, ByVal ventana As Long, ByRef entrada As Any, ByRef salida As Any) As Long

Private Type DDEXCFG
    vsync As Byte
    api As Byte
    MODO As Byte
    modo2 As Byte
End Type

Public Enum API_grafica
    DX9 = 0
    DX92 = 1
    OGL = 2
    DX8 = 3
End Enum

Public Enum MODO
    DX9_SOF = 0
    DX9_HARD = 1
    OGL_ = 2 'solo es una guia
End Enum


Private Declare Function VB6_ArrayPointer Lib "msvbvm60" Alias "VarPtr" (arr As Any) As Long

Private puntero As Long
Public Sub PreCargar(archivo As Long)
On Error Resume Next
    PreCargarGrafico ByVal VB6_ArrayPointer(puntero), archivo
End Sub
'vertices x,y,z,rhw,color
Public Sub DibujarVerticesEx(ByRef v, ByVal cantidad As Long, ByVal alphaf As Long)
    DibujarVertices puntero, v, cantidad, alphaf
End Sub
'carpeta = la carpeta donde se encuentren los graficos
'soporta BMP y PNG  automaticamente se carga el BMP numero y si no existe lo busca en PNG

Public Function Iniciar(ByVal hwnd As Long, carpeta As String, Optional api As API_grafica = DX9, Optional m As MODO = DX9_HARD, Optional SVertical As Boolean = False) As Boolean
    On Error Resume Next
    
    Dim cc As DDEXCFG
    puntero = 0
    Dim ll As Long
    If SVertical = True Then
        cc.vsync = 1
    Else
        cc.vsync = 0
    End If
    cc.api = api
    cc.MODO = m
    cc.modo2 = 0
    IniciarMotor puntero, hwnd, ByVal StrPtr(StrConv(carpeta, vbFromUnicode)), cc
    'IniciarMotor ByVal VB6_ArrayPointer(puntero), hwnd, ByVal StrPtr(StrConv(carpeta, vbFromUnicode)), cc
End Function
Public Sub SetearZoom(zoom_ As Long)
On Error Resume Next
   ZoomPantalla ByVal VB6_ArrayPointer(puntero), zoom_
End Sub
Public Sub Termina()
    TerminarMotor ByVal VB6_ArrayPointer(puntero)
End Sub

Public Function CrearFuente(Nombre As String, tam As Long, Optional negrita As Boolean = False, Optional cursiva As Boolean = False)
    Dim t As Integer
    t = 0
    If negrita Then t = 255
    CrearFuente = CargarFuenteX(ByVal VB6_ArrayPointer(puntero), ByVal StrPtr(StrConv(Nombre, vbFromUnicode)), tam, t, cursiva)
End Function
Friend Sub DBTexto(ByVal Str As String, ByVal X As Integer, ByVal Y As Integer, lColor As tDDEXRGBA, Optional fuente As Long = 1)
    If LenB(Str) = 0 Then Exit Sub
    DibujarTexto puntero, ByVal StrPtr(StrConv(Str, vbFromUnicode)), X, Y, lColor, fuente
    'DibujarTexto puntero, ByVal StrPtr(StrConv(Str, vbFromUnicode)), X, Y, lColor, fuente
    'DibujarTexto ByVal VB6_ArrayPointer(puntero), ByVal StrPtr(StrConv(Str, vbFromUnicode)), X, Y, lColor, fuente
End Sub
Friend Sub DbCaja(X As Long, Y As Long, ancho As Long, alto As Long, lColor As tDDEXRGBA, Optional angulo As Long = 0)
    Dim r As RECT
    r.Left = X
    r.Top = Y
    r.bottom = Y + alto
    r.Right = X + ancho
    DibujarCaja puntero, r, lColor, angulo
    'DibujarCajaEx ByVal VB6_ArrayPointer(puntero), r, lColor, angulo
End Sub
Friend Sub DbCajaEx(X As Long, Y As Long, ancho As Long, alto As Long, lColor As tDDEXRGBA, Optional angulo As Long = 0)
    Dim r As RECT
    r.Left = X
    r.Top = Y
    r.bottom = Y + alto
    r.Right = X + ancho
    DibujarCajaEx puntero, r, lColor, angulo
    'DibujarCajaEx ByVal VB6_ArrayPointer(puntero), r, lColor, angulo
End Sub

Friend Sub DbLinea(X As Long, Y As Long, x2 As Long, y2 As Long, lColor As tDDEXRGBA)
    Dim r As RECT
    r.Left = X
    r.Top = Y
    r.bottom = y2
    r.Right = x2
    DibujarLineaEx ByVal VB6_ArrayPointer(puntero), r, lColor
End Sub

Friend Sub MostrarPantallaEx(entrada As RECT, salida As RECT, ventana As Long)
    DibujarPantallaEx ByVal VB6_ArrayPointer(puntero), ventana, entrada, salida
End Sub
Public Sub MostrarPantalla()
    DibujarMotor puntero
    'DibujarMotor ByVal VB6_ArrayPointer(puntero)
End Sub
Public Sub GirarPantalla(angulo As Long)
    RotarPantalla ByVal VB6_ArrayPointer(puntero), angulo
End Sub
Public Sub LimpiarPantalla()
    LimpiarMotor puntero
    'LimpiarMotor ByVal VB6_ArrayPointer(puntero)
End Sub
Friend Sub DBEx(archivo As Long, Rec As RECT, ByVal X As Integer, ByVal Y As Integer, c As tDDEXRGBA, Optional angulo As Long = 0)
    DibujarRGraficoEx puntero, archivo, Rec, X, Y, c, angulo
    'DibujarRGraficoEx ByVal VB6_ArrayPointer(puntero), archivo, Rec, X, Y, c, angulo
End Sub

Friend Sub DBAlfa(archivo As Long, Rec As RECT, ByVal X As Integer, ByVal Y As Integer, c As tDDEXRGBA)
    DibujarRGraficoAlpha puntero, archivo, Rec, X, Y, c
    'DibujarRGraficoAlpha ByVal VB6_ArrayPointer(puntero), archivo, Rec, X, Y, c
End Sub
Friend Sub DBGrafico(archivo As Long, ByVal X As Long, ByVal Y As Long, RRect As RECT, Optional ColorKey As Boolean = True)
    DibujarRGrafico puntero, archivo, RRect, X, Y
    'DibujarRGrafico ByVal VB6_ArrayPointer(puntero), archivo, RRect, X, Y
End Sub



