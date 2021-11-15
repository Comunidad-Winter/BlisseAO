VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cls_Motor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'*******************************************************
'*Esta Clase fue Creada por Loopzer                       *
'*para facilitar el uso de                                           *
'*Directx9 y otros efectos                                   *
'*******************************************************
'Mail=Loopzer@gmail.com
'DDEX Ver. 3.0

Private Declare Function DibujarCajaEx Lib "DDEXv3.dll" (ByRef r As Any, ByRef color As Any, ByVal angulo As Long) As Long

Private Declare Function DibujarLineaEx Lib "DDEXv3.dll" (ByRef r As Any, ByRef color As Any) As Long

Private Declare Function PreCargarGrafico Lib "DDEXv3.dll" (ByVal archivo As Long) As Long


Private Declare Function IniciarMotor Lib "DDEXv3.dll" (ByVal ventana As Long, ByRef carga As Any, ByVal configuracion As Long) As Long


Private Declare Function RotarPantalla Lib "DDEXv3.dll" (ByVal angulo As Long) As Long

Private Declare Function TerminarMotor Lib "DDEXv3.dll" () As Long

Private Declare Function DibujarRGrafico Lib "DDEXv3.dll" (ByVal archivo As Long, ByRef r As Any, ByVal X As Long, ByVal Y As Long) As Long

Private Declare Function DibujarRGraficoAlpha Lib "DDEXv3.dll" (ByVal archivo As Long, ByRef r As Any, ByVal X As Long, ByVal Y As Long, ByRef color As Any) As Long

Private Declare Function DibujarRGraficoEx Lib "DDEXv3.dll" (ByVal archivo As Long, ByRef r As Any, ByVal X As Long, ByVal Y As Long, ByRef color As Any, ByVal angulo As Long) As Long

Private Declare Function DibujarTexto Lib "DDEXv3.dll" (ByRef texto As Any, ByVal X As Long, ByVal Y As Long, ByRef color As Any, ByVal fuente As Long) As Long

Private Declare Function CargarFuenteX Lib "DDEXv3.dll" (ByRef fuente As Any, ByVal tam As Long, ByVal negrita As Byte, ByVal cursiva As Byte) As Long

Private Declare Function LimpiarMotor Lib "DDEXv3.dll" () As Long

Private Declare Function DibujarMotor Lib "DDEXv3.dll" () As Long

Public Sub PreCargar(archivo As Long)
    PreCargarGrafico archivo
End Sub

Public Function Iniciar(ByVal hwnd As Long, carpeta As String, Optional SVertical As Boolean = False) As Boolean
    Dim ll As Long
    
    If SVertical = True Then
        ll = 2 Or 1
    Else
        ll = 2
    End If
    
    IniciarMotor hwnd, ByVal StrPtr(StrConv(carpeta, vbFromUnicode)), ll
End Function


Public Sub Termina()
    TerminarMotor
End Sub

Public Function CrearFuente(Nombre As String, tam As Long, Optional negrita As Boolean = False, Optional cursiva As Boolean = False)
    Dim t As Integer
    t = 0
    If negrita Then t = 255
    CrearFuente = CargarFuenteX(ByVal StrPtr(StrConv(Nombre, vbFromUnicode)), tam, t, cursiva)
End Function

Friend Sub DibujarCaja(X As Long, Y As Long, ancho As Long, alto As Long, lColor As tDDEXRGBA, Optional angulo As Long = 0)
    Dim r As RECT
    r.Left = X
    r.Top = Y
    r.bottom = Y + alto
    r.Right = X + ancho
    DibujarCajaEx r, lColor, angulo
End Sub

Friend Sub DibujarLinea(X As Long, Y As Long, x2 As Long, y2 As Long, lColor As tDDEXRGBA)
    Dim r As RECT
    r.Left = X
    r.Top = Y
    r.bottom = y2
    r.Right = x2
    DibujarLineaEx r, lColor
End Sub

Friend Sub DBTexto(ByVal Str As String, ByVal X As Integer, ByVal Y As Integer, lColor As tDDEXRGBA, Optional fuente As Long = 0)
    DibujarTexto ByVal StrPtr(StrConv(Str, vbFromUnicode)), X, Y, lColor, fuente
End Sub

Public Sub MostrarPantalla()
    DibujarMotor
End Sub

Public Sub GirarPantalla(angulo As Long)
    RotarPantalla angulo
End Sub

Public Sub LimpiarPantalla()
    LimpiarMotor
End Sub

Friend Sub DBEx(archivo As Long, Rec As RECT, ByVal X As Integer, ByVal Y As Integer, c As tDDEXRGBA, Optional angulo As Long = 0)
    DibujarRGraficoEx archivo, Rec, X, Y, c, angulo
End Sub

Friend Sub DBAlfa(archivo As Long, Rec As RECT, ByVal X As Integer, ByVal Y As Integer, c As tDDEXRGBA)
    DibujarRGraficoAlpha archivo, Rec, X, Y, c
End Sub

Friend Sub DBGrafico(archivo As Long, ByVal X As Long, ByVal Y As Long, RRect As RECT, Optional ColorKey As Boolean = True)
    DibujarRGrafico archivo, RRect, X, Y
End Sub
