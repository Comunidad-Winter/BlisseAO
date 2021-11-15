VERSION 5.00
Begin VB.Form frmUserRequest 
   BorderStyle     =   0  'None
   ClientHeight    =   2430
   ClientLeft      =   0
   ClientTop       =   -75
   ClientWidth     =   4650
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   162
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   310
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin CLBLISSEAO.BGAOButton imgCerrar 
      Height          =   285
      Left            =   120
      TabIndex        =   1
      Top             =   1920
      Width           =   3000
      _ExtentX        =   5292
      _ExtentY        =   450
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
   Begin VB.TextBox Text1 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1395
      Left            =   225
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   405
      Width           =   4185
   End
End
Attribute VB_Name = "frmUserRequest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private clsFormulario As New clsFormMovementManager

Public Sub recievePeticion(ByVal p As String)

    Text1 = Replace$(p, "º", vbCrLf)
    Me.Show vbModeless, frmMain

End Sub

Private Sub Form_Load()
    '   Handles Form movement (drag and drop).
    clsFormulario.Initialize Me
    
    Me.Picture = General_Set_GUI("VentanaPeticion")
    imgCerrar.Init s_Large
End Sub

Private Sub imgCerrar_Click()
    Unload Me
End Sub
