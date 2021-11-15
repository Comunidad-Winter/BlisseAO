VERSION 5.00
Begin VB.Form FrmNewCuenta 
   BackColor       =   &H0000FF00&
   BorderStyle     =   0  'None
   Caption         =   "  Crear cuenta."
   ClientHeight    =   4650
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6000
   ControlBox      =   0   'False
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
   ScaleHeight     =   310
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   400
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin CLBLISSEAO.BGAOButton Image1 
      Height          =   285
      Index           =   0
      Left            =   1500
      TabIndex        =   5
      Top             =   4200
      Width           =   3000
      _ExtentX        =   5292
      _ExtentY        =   503
      Caption         =   "Crear Cuenta"
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
   Begin VB.TextBox TMail 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   1440
      TabIndex        =   1
      Top             =   2370
      Width           =   3075
   End
   Begin VB.TextBox TRepPass 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      IMEMode         =   3  'DISABLE
      Left            =   3135
      PasswordChar    =   "•"
      TabIndex        =   4
      Top             =   3780
      Width           =   2475
   End
   Begin VB.TextBox TPass 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      IMEMode         =   3  'DISABLE
      Left            =   375
      PasswordChar    =   "•"
      TabIndex        =   3
      Top             =   3780
      Width           =   2475
   End
   Begin VB.TextBox TRepMail 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   1440
      TabIndex        =   2
      Top             =   2955
      Width           =   3075
   End
   Begin VB.TextBox TName 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   1440
      TabIndex        =   0
      Top             =   825
      Width           =   3075
   End
   Begin CLBLISSEAO.BGAOButton Image1 
      Height          =   285
      Index           =   1
      Left            =   5520
      TabIndex        =   6
      Top             =   270
      Width           =   285
      _ExtentX        =   503
      _ExtentY        =   503
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
End
Attribute VB_Name = "FrmNewCuenta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
Me.Picture = General_Set_GUI("NewAcount")
Image1(0).Init s_Large
Image1(1).Init s_Small
End Sub

Private Sub Image1_Click(Index As Integer)
Call Audio.PlayWave(SND_CLICK)

Select Case Index
    Case 0
        If TRepMail.Text <> TMail.Text Then
            MsgBox "La confirmación de tu cuenta electrónica no coincide con la misma."
            Exit Sub
        End If
        
        Cuenta.name = UCase(LTrim(TName.Text))
        Cuenta.Pass = TPass.Text
        Cuenta.Email = TMail.Text

        If frmMain.Winsock1.State <> sckClosed Then
            frmMain.Winsock1.Close
            DoEvents
        End If
    
        If General_Check_AccountData(True, True) = True Then
            #If SeguridadBlisse = 1 Then
                If RevisarCodigo = False Then Exit Sub
            #End If
            EstadoLogin = CrearCuenta
            frmMain.Winsock1.Connect Server_IP, Server_Port
        End If
        
        Exit Sub
        
    Case 1
        Unload Me
End Select
End Sub
