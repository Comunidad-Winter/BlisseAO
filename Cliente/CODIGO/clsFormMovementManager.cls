VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsFormMovementManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim WithEvents Formulario As Form
Attribute Formulario.VB_VarHelpID = -1
Dim WithEvents Picture As PictureBox
Attribute Picture.VB_VarHelpID = -1

Private DragX As Single
Private Dragy As Single

Private LimitY As Integer

Public Sub Initialize(ByRef tForm As Form, Optional ByRef tPict As PictureBox, Optional ByVal VerticalLimit As Integer = 0)
'***************************************************
'Author: ZaMa
'Last Modify Date: 29/12/09
'Sets the reference to the form
'***************************************************
    Set Formulario = tForm
    Set Picture = tPict
    
    LimitY = VerticalLimit
End Sub

Private Sub Formulario_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
'***************************************************
'Author: ZaMa
'Last Modify Date: 29/12/09
'Implements the mouse down event of the form.
'***************************************************
    If LimitY > 0 Then
        If Y > LimitY Then Exit Sub
    End If

    DragX = x
    Dragy = Y
End Sub

Private Sub Picture_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
'***************************************************
'Author: ZaMa
'Last Modify Date: 29/12/09
'Implements the mouse down event of the form.
'Last Modify Date: 29/05/10 by TonchitoZ
'***************************************************
    If LimitY > 0 Then
        If Y > LimitY Then Exit Sub
    End If

    DragX = x
    Dragy = Y
End Sub

Private Sub Picture_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
'***************************************************
'Author: ZaMa
'Last Modify Date: 01/07/10
'Implements the mouse move event of the form.
'Last Modify Date: 29/05/10 by TonchitoZ
'***************************************************
    If LimitY > 0 Then
        If Y > LimitY Then Exit Sub
    End If
    
    If Button > 0 Then
        Formulario.Left = Formulario.Left + (x - DragX)
        Formulario.Top = Formulario.Top + (Y - Dragy)
    End If
End Sub

Private Sub Formulario_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
'***************************************************
'Author: ZaMa
'Last Modify Date: 01/07/10
'Implements the mouse move event of the form.
'***************************************************
    If LimitY > 0 Then
        If Y > LimitY Then Exit Sub
    End If
    
    If Button > 0 Then
        Formulario.Left = Formulario.Left + (x - DragX)
        Formulario.Top = Formulario.Top + (Y - Dragy)
    End If
End Sub
