VERSION 5.00
Begin VB.Form frmEligeAlineacion 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   5265
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   6720
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5265
   ScaleWidth      =   6720
   StartUpPosition =   1  'CenterOwner
   Begin CLBLISSEAO.BGAOButton imgSalir 
      Height          =   285
      Left            =   5040
      TabIndex        =   0
      Top             =   4800
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   503
      Caption         =   "Salir"
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
   Begin VB.Image imgReal 
      Height          =   765
      Left            =   795
      Tag             =   "1"
      Top             =   300
      Width           =   5745
   End
   Begin VB.Image imgNeutral 
      Height          =   435
      Left            =   810
      Tag             =   "1"
      Top             =   2220
      Width           =   5730
   End
   Begin VB.Image imgLegal 
      Height          =   585
      Left            =   810
      Tag             =   "1"
      Top             =   1320
      Width           =   5715
   End
   Begin VB.Image imgCaos 
      Height          =   555
      Left            =   825
      Tag             =   "1"
      Top             =   4110
      Width           =   5700
   End
   Begin VB.Image imgCriminal 
      Height          =   585
      Left            =   825
      Tag             =   "1"
      Top             =   3150
      Width           =   5865
   End
End
Attribute VB_Name = "frmEligeAlineacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private clsFormulario As clsFormMovementManager

Private Enum eAlineacion
    ieREAL = 0
    ieCAOS = 1
    ieNeutral = 2
    ieLegal = 4
    ieCriminal = 5
End Enum

Private Sub Form_Load()
    '   Handles Form movement (drag and drop).
    Set clsFormulario = New clsFormMovementManager
    clsFormulario.Initialize Me
    
    Me.Picture = General_Set_GUI("VentanaFundarClan")
    
    imgSalir.Init s_Normal
End Sub

Private Sub imgCaos_Click()
    Call WriteGuildFundation(eAlineacion.ieCAOS)
    Unload Me
End Sub

Private Sub imgCriminal_Click()
    Call WriteGuildFundation(eAlineacion.ieCriminal)
    Unload Me
End Sub

Private Sub imgLegal_Click()
    Call WriteGuildFundation(eAlineacion.ieLegal)
    Unload Me
End Sub

Private Sub imgNeutral_Click()
    Call WriteGuildFundation(eAlineacion.ieNeutral)
    Unload Me
End Sub

Private Sub imgReal_Click()
    Call WriteGuildFundation(eAlineacion.ieREAL)
    Unload Me
End Sub

Private Sub imgSalir_Click()
    Unload Me
End Sub
