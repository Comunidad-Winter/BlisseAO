VERSION 5.00
Begin VB.Form FrmBancoCuenta 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Form1"
   ClientHeight    =   2280
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   8580
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2280
   ScaleWidth      =   8580
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CmdRetirar 
      Caption         =   "Retirar"
      Height          =   375
      Left            =   6345
      TabIndex        =   5
      Top             =   855
      Width           =   2130
   End
   Begin VB.CommandButton cmdDepositar 
      Caption         =   "Depositar"
      Height          =   375
      Left            =   6345
      TabIndex        =   4
      Top             =   1395
      Width           =   2130
   End
   Begin VB.TextBox cantidad 
      Height          =   285
      Left            =   7320
      TabIndex        =   2
      Text            =   "1"
      Top             =   1890
      Width           =   1080
   End
   Begin VB.ListBox List1 
      Height          =   2010
      Index           =   1
      ItemData        =   "FrmBancoCuenta.frx":0000
      Left            =   3240
      List            =   "FrmBancoCuenta.frx":0002
      TabIndex        =   1
      Top             =   135
      Width           =   2985
   End
   Begin VB.ListBox List1 
      Height          =   2010
      Index           =   0
      ItemData        =   "FrmBancoCuenta.frx":0004
      Left            =   45
      List            =   "FrmBancoCuenta.frx":0006
      TabIndex        =   0
      Top             =   135
      Width           =   2985
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Cantidad"
      ForeColor       =   &H80000009&
      Height          =   195
      Left            =   6480
      TabIndex        =   3
      Top             =   1950
      Width           =   765
   End
End
Attribute VB_Name = "FrmBancoCuenta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' ProyectoAO Documentación ***********************************
' Autor: TonchitoZ
' Descripción: Formulario de Bóveda Compartida entre Cuenta _
                de Donador
' Última modificación: 07/11/2012
'*************************************************************

Option Explicit

Private Sub cmdDepositar_Click()
    Call Audio.PlayWave(SND_CLICK)
    If List1(1).List(List1(1).ListIndex) = "" Or List1(1).ListIndex < 0 Then Exit Sub
    If Not IsNumeric(cantidad.Text) Then Exit Sub
    Call WriteBankDepositCuenta(List1(1).ListIndex + 1, cantidad.Text)
End Sub
Private Sub CmdRetirar_Click()
    Call Audio.PlayWave(SND_CLICK)
    If List1(0).List(List1(0).ListIndex) = "" Or List1(0).ListIndex < 0 Then Exit Sub
    If Not IsNumeric(cantidad.Text) Then Exit Sub
    Call WriteBankRetirarCuenta(List1(0).ListIndex + 1, cantidad.Text)
End Sub
Private Sub Form_Load()
Dim i As Byte
    Call FrmBancoCuenta.List1(1).Clear
        For i = 1 To MAX_INVENTORY_SLOTS
            If Inventario.OBJIndex(i) <> 0 Then
                FrmBancoCuenta.List1(1).AddItem Inventario.ItemName(i) & "(" & Inventario.Amount(i) & ")"
            Else
                FrmBancoCuenta.List1(1).AddItem "(VACIO)"
            End If
        Next i
End Sub
