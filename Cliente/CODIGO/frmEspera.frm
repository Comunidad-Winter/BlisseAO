VERSION 5.00
Begin VB.Form frmEspera 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   1500
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4500
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmEspera.frx":0000
   ScaleHeight     =   100
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   300
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   250
      Left            =   3840
      Top             =   120
   End
   Begin CLBLISSEAO.ctrGIF ctrGIF1 
      Height          =   480
      Left            =   360
      TabIndex        =   0
      Top             =   600
      Width           =   480
      _ExtentX        =   847
      _ExtentY        =   847
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Por favor espere."
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   255
      Left            =   960
      TabIndex        =   1
      Top             =   720
      Width           =   3255
   End
End
Attribute VB_Name = "frmEspera"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private clsFormulario As clsFormMovementManager

Private Sub Form_Load()
    ' Handles Form movement (drag and drop).
    Set clsFormulario = New clsFormMovementManager
    clsFormulario.Initialize Me
    
    ctrGIF1.GifPath = App.path & "\Graficos\ajax-loader.gif"
    ctrGIF1.StartGif
    Label1.Caption = "Por favor espere"
    Timer1.Enabled = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ctrGIF1.StopGif
    Timer1.Enabled = False
End Sub

Private Sub Timer1_Timer()
    If Label1.Caption = "Por favor espere" Then
        Label1.Caption = "Por favor espere."
    ElseIf Label1.Caption = "Por favor espere." Then
        Label1.Caption = "Por favor espere.."
    ElseIf Label1.Caption = "Por favor espere.." Then
        Label1.Caption = "Por favor espere..."
    ElseIf Label1.Caption = "Por favor espere..." Then
        Label1.Caption = "Por favor espere"
    End If
End Sub
