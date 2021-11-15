VERSION 5.00
Begin VB.Form frmBancoObj 
   BackColor       =   &H80000000&
   BorderStyle     =   0  'None
   ClientHeight    =   5775
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6630
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   ForeColor       =   &H8000000F&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   385
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   442
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin CLBLISSEAO.BGAOButton BGAOButton1 
      Height          =   465
      Index           =   0
      Left            =   120
      TabIndex        =   11
      Top             =   4500
      Width           =   3000
      _ExtentX        =   5292
      _ExtentY        =   820
      Caption         =   "Retirar ítem"
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
   Begin CLBLISSEAO.BGAOButton imgRetirarOro 
      Height          =   285
      Left            =   4785
      TabIndex        =   8
      Top             =   1290
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   503
      Caption         =   "Retirar"
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
   Begin VB.PictureBox PicInv 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      CausesValidation=   0   'False
      ClipControls    =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   2700
      Left            =   3450
      ScaleHeight     =   180
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   216
      TabIndex        =   7
      Top             =   1830
      Width           =   3240
   End
   Begin VB.PictureBox PicBancoInv 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      CausesValidation=   0   'False
      ClipControls    =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   2700
      Left            =   0
      ScaleHeight     =   180
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   216
      TabIndex        =   6
      Top             =   1830
      Width           =   3240
   End
   Begin VB.TextBox CantidadOro 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   270
      Left            =   3090
      MaxLength       =   7
      TabIndex        =   5
      Text            =   "1"
      Top             =   1275
      Width           =   1035
   End
   Begin VB.TextBox cantidad 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   270
      Left            =   3045
      MaxLength       =   5
      TabIndex        =   4
      Text            =   "1"
      Top             =   5400
      Width           =   915
   End
   Begin CLBLISSEAO.BGAOButton imgDepositarOro 
      Height          =   285
      Left            =   240
      TabIndex        =   9
      Top             =   1290
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   503
      Caption         =   "Depositar"
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
   Begin CLBLISSEAO.BGAOButton imgCerrar 
      Height          =   255
      Left            =   6270
      TabIndex        =   10
      Top             =   105
      Width           =   255
      _ExtentX        =   2778
      _ExtentY        =   661
      Caption         =   "X"
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
   Begin CLBLISSEAO.BGAOButton BGAOButton1 
      Height          =   465
      Index           =   1
      Left            =   3525
      TabIndex        =   12
      Top             =   4500
      Width           =   3000
      _ExtentX        =   3201
      _ExtentY        =   661
      Caption         =   "Depositar ítem"
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
   Begin VB.Label lblUserGld 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   3075
      TabIndex        =   3
      Top             =   855
      Width           =   105
   End
   Begin VB.Image CmdMoverBov 
      Height          =   240
      Index           =   1
      Left            =   3195
      Top             =   1980
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image CmdMoverBov 
      Height          =   240
      Index           =   0
      Left            =   3195
      Top             =   2340
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   1
      Left            =   3720
      TabIndex        =   2
      Top             =   5100
      Visible         =   0   'False
      Width           =   60
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   2
      Left            =   3720
      TabIndex        =   1
      Top             =   4920
      Visible         =   0   'False
      Width           =   60
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   4950
      Width           =   3000
   End
End
Attribute VB_Name = "frmBancoObj"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' ProyectoAO Documentación ***********************************
' Autor: Standelf
' Descripción: Bóveda de usuarios
' Última modificación: 07/11/2012
'*************************************************************

Option Explicit

Private clsFormulario As clsFormMovementManager
Public LasActionBuy As Boolean
Public LastIndex1 As Integer
Public LastIndex2 As Integer
Public NoPuedeMover As Boolean

Private Sub BGAOButton1_Click(Index As Integer)
    Call Audio.PlayWave(SND_CLICK)
    If InvBanco(Index).SelectedItem = 0 Then Exit Sub
    If Not IsNumeric(cantidad.Text) Then Exit Sub

    Select Case Index
        Case 0
            LastIndex1 = InvBanco(0).SelectedItem
            LasActionBuy = True
            Call WriteBankExtractItem(InvBanco(0).SelectedItem, cantidad.Text)
            
       Case 1
            LastIndex2 = InvBanco(1).SelectedItem
            LasActionBuy = False
            Call WriteBankDeposit(InvBanco(1).SelectedItem, cantidad.Text)
    End Select
End Sub

Private Sub cantidad_Change()
    If Val(cantidad.Text) < 1 Then
        cantidad.Text = 1
    End If
    If Val(cantidad.Text) > MAX_INVENTORY_OBJS Then
        cantidad.Text = MAX_INVENTORY_OBJS
    End If
End Sub
Private Sub cantidad_KeyPress(KeyAscii As Integer)
    If (KeyAscii <> 8) Then
        If (KeyAscii <> 6) And (KeyAscii < 48 Or KeyAscii > 57) Then
            KeyAscii = 0
        End If
    End If
End Sub
Private Sub CantidadOro_Change()
    If Val(CantidadOro.Text) < 1 Then
        cantidad.Text = 1
    End If
End Sub
Private Sub CantidadOro_KeyPress(KeyAscii As Integer)
    If (KeyAscii <> 8) Then
        If (KeyAscii <> 6) And (KeyAscii < 48 Or KeyAscii > 57) Then
            KeyAscii = 0
        End If
    End If
End Sub
Private Sub Form_Load()
    Set clsFormulario = New clsFormMovementManager
    clsFormulario.Initialize Me
    Me.Picture = General_Set_GUI("Boveda")
    
    imgCerrar.Init s_Small
    imgDepositarOro.Init s_Normal
    imgRetirarOro.Init s_Normal
    BGAOButton1(0).Init b_Large
    BGAOButton1(1).Init b_Large
End Sub

Private Sub Image1_Click(Index As Integer)

End Sub

Private Sub imgDepositarOro_Click()
    Call WriteBankDepositGold(Val(CantidadOro.Text))
End Sub
Private Sub imgRetirarOro_Click()
    Call WriteBankExtractGold(Val(CantidadOro.Text))
End Sub
Private Sub PicBancoInv_Click()
    If InvBanco(0).SelectedItem <> 0 Then
        With UserBancoInventory(InvBanco(0).SelectedItem)
            Label1(0).Caption = .name
            Select Case .OBJType
                Case 2, 32
                    Label1(1).Caption = "Máx Golpe:" & .MaxHit
                    Label1(2).Caption = "Mín Golpe:" & .MinHit
                    Label1(1).Visible = True
                    Label1(2).Visible = True
                Case 3, 16, 17
                    Label1(1).Caption = "Máx Defensa:" & .MaxDef
                    Label1(2).Caption = "Mín Defensa:" & .MinDef
                    Label1(1).Visible = True
                    Label1(2).Visible = True
                Case Else
                    Label1(1).Visible = False
                    Label1(2).Visible = False
            End Select
        End With
    Else
        Label1(0).Caption = ""
        Label1(1).Visible = False
        Label1(2).Visible = False
    End If
End Sub
Private Sub PicInv_Click()
    If InvBanco(1).SelectedItem <> 0 Then
        With Inventario
            Label1(0).Caption = .ItemName(InvBanco(1).SelectedItem)
            Select Case .OBJType(InvBanco(1).SelectedItem)
                Case eObjType.otWeapon, eObjType.otFlechas
                    Label1(1).Caption = "Máx Golpe:" & .MaxHit(InvBanco(1).SelectedItem)
                    Label1(2).Caption = "Mín Golpe:" & .MinHit(InvBanco(1).SelectedItem)
                    Label1(1).Visible = True
                    Label1(2).Visible = True
                Case eObjType.otcasco, eObjType.otArmadura, eObjType.otescudo '   3, 16, 17
                    Label1(1).Caption = "Máx Defensa:" & .MaxDef(InvBanco(1).SelectedItem)
                    Label1(2).Caption = "Mín Defensa:" & .MinDef(InvBanco(1).SelectedItem)
                    Label1(1).Visible = True
                    Label1(2).Visible = True
                Case Else
                    Label1(1).Visible = False
                    Label1(2).Visible = False
            End Select
        End With
    Else
        Label1(0).Caption = ""
        Label1(1).Visible = False
        Label1(2).Visible = False
    End If
End Sub
Private Sub imgCerrar_Click()
    Call WriteBankEnd
    NoPuedeMover = False
End Sub
