Attribute VB_Name = "Mod_RoundControls"
'***************************************************
'Author: Ezequiel Juárez (Standelf)
'Last Modification: 22/05/10
'Blisse-AO | Round Controls and Forms
'***************************************************

Option Explicit

Private Declare Function CreateRoundRectRgn Lib "gdi32" ( _
    ByVal x1 As Long, _
    ByVal Y1 As Long, _
    ByVal x2 As Long, _
    ByVal Y2 As Long, _
    ByVal X3 As Long, _
    ByVal Y3 As Long) As Long

Private Declare Function SetWindowRgn Lib "user32" ( _
    ByVal hWnd As Long, _
    ByVal hRgn As Long, _
    ByVal bRedraw As Boolean) As Long

Public Sub Round_Form(Formulario As Form, Radio As Long)
'***************************************************
'Author: Unknown
'Last Modification: 22/05/10 by Standelf
'Edit sub to can round forms
'***************************************************
Dim Region As Long
Dim Ret As Long
Dim Ancho As Long
Dim Alto As Long
Dim old_Scale As Integer
    
    old_Scale = Formulario.ScaleMode
    Formulario.ScaleMode = vbPixels
    Ancho = Formulario.ScaleWidth
    Alto = Formulario.ScaleHeight
    Region = CreateRoundRectRgn(0, 0, Ancho, Alto, Radio, Radio)
    Ret = SetWindowRgn(Formulario.hWnd, Region, True)
    Formulario.ScaleMode = old_Scale
End Sub

Public Sub Round_Picture(PictureBox As PictureBox, Radio As Long)
'***************************************************
'Author: Unknown
'Last Modification: ??/??/????
'***************************************************
Dim Region As Long
Dim Ret As Long
Dim Ancho As Long
Dim Alto As Long
Dim old_Scale As Integer
    
    old_Scale = PictureBox.ScaleMode
    PictureBox.ScaleMode = vbPixels
    Ancho = PictureBox.ScaleWidth
    Alto = PictureBox.ScaleHeight
    Region = CreateRoundRectRgn(0, 0, Ancho, Alto, Radio, Radio)
    Ret = SetWindowRgn(PictureBox.hWnd, Region, True)
    PictureBox.ScaleMode = old_Scale
End Sub
