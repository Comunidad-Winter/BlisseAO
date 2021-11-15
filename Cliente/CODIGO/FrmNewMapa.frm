VERSION 5.00
Begin VB.Form FrmNewMapa 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   ClientHeight    =   7920
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8595
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
   Picture         =   "FrmNewMapa.frx":0000
   ScaleHeight     =   528
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   573
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "FrmNewMapa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim comMapaX, comMapaY As Byte

Private Sub Form_Load()
comMapaX = 36
comMapaY = 45
End Sub
