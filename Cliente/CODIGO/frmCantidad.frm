VERSION 5.00
Begin VB.Form frmCantidad 
   BackColor       =   &H80000000&
   BorderStyle     =   0  'None
   ClientHeight    =   1470
   ClientLeft      =   1635
   ClientTop       =   4410
   ClientWidth     =   3240
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   98
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   216
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin CLBLISSEAO.BGAOButton imgTirar 
      Height          =   465
      Left            =   120
      TabIndex        =   2
      Top             =   900
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   820
      Caption         =   "Tirar"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FontName        =   "Verdana"
      FontSize        =   8,25
   End
   Begin CLBLISSEAO.BGAOButton imgTirarTodo 
      Height          =   465
      Left            =   1650
      TabIndex        =   1
      Top             =   900
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   820
      Caption         =   "Tirar Todo"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FontName        =   "Verdana"
      FontSize        =   8,25
   End
   Begin VB.TextBox txtCantidad 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   390
      Left            =   435
      MaxLength       =   5
      TabIndex        =   0
      Top             =   435
      Width           =   2325
   End
End
Attribute VB_Name = "frmCantidad"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private clsFormulario As clsFormMovementManager
Private DDy As Byte
Private DDx As Byte

Private Sub Form_Load()
    Set clsFormulario = New clsFormMovementManager
    clsFormulario.Initialize Me
    Me.Picture = General_Set_GUI("VentanaTirarOro")
    imgTirar.Init b_Normal
    imgTirarTodo.Init b_Normal
End Sub

Public Sub IniciarDD(ByVal dX As Byte, ByVal Dy As Byte)
    DDx = dX
    DDy = Dy
    Me.Show vbModal
End Sub

Private Sub imgTirar_Click()
    If LenB(txtCantidad.Text) > 0 Then
        If Not IsNumeric(txtCantidad.Text) Then Exit Sub
        If frmMain.UsabaDrag Or (DDx <> 0 Or DDy <> 0) Then
            Call WriteDrop(Inventario.SelectedItem, frmCantidad.txtCantidad.Text, DDx, DDy)
        Else
            Call WriteDrop(Inventario.SelectedItem, frmCantidad.txtCantidad.Text, CByte(UserPos.X), CByte(UserPos.Y))
        End If
        frmCantidad.txtCantidad.Text = ""
    End If
    
    DDx = 0: DDy = 0
    Unload Me
End Sub
Private Sub imgTirarTodo_Click()
    If Inventario.SelectedItem = 0 Then Exit Sub
    If Inventario.SelectedItem <> FLAGORO Then
        If DDx <> 0 Or DDy <> 0 Then
            Call WriteDrop(Inventario.SelectedItem, Inventario.Amount(Inventario.SelectedItem), DDx, DDy)
        Else
            Call WriteDrop(Inventario.SelectedItem, Inventario.Amount(Inventario.SelectedItem), CByte(UserPos.X), CByte(UserPos.Y))
        End If
        
        DDx = 0: DDy = 0
        Unload Me
    Else
        If UserGLD > 10000 Then
            Call WriteDrop(Inventario.SelectedItem, 10000, UserPos.X, Val(UserPos.Y))
            DDx = 0: DDy = 0
            Unload Me
        Else
            Call WriteDrop(Inventario.SelectedItem, UserGLD, UserPos.X, Val(UserPos.Y))
            DDx = 0: DDy = 0
            Unload Me
        End If
    End If
    frmCantidad.txtCantidad.Text = ""
End Sub
Private Sub txtCantidad_Change()
On Error GoTo ErrHandler
    If Val(txtCantidad.Text) < 0 Then
        txtCantidad.Text = "1"
    End If
    If Val(txtCantidad.Text) > MAX_INVENTORY_OBJS Then
        txtCantidad.Text = "10000"
    End If
    
    Exit Sub
ErrHandler:
    txtCantidad.Text = "1"
End Sub
Private Sub txtCantidad_KeyPress(KeyAscii As Integer)
    If (KeyAscii <> 8) Then
        If (KeyAscii < 48 Or KeyAscii > 57) Then
            KeyAscii = 0
        End If
    End If
End Sub
