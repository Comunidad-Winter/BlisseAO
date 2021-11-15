VERSION 5.00
Begin VB.Form FrmRetorno 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   2115
   ClientLeft      =   5520
   ClientTop       =   4875
   ClientWidth     =   4410
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2115
   ScaleWidth      =   4410
   ShowInTaskbar   =   0   'False
   Begin VB.Label Label1 
      Caption         =   "Retornar a la ciudad inicial"
      Height          =   495
      Left            =   240
      TabIndex        =   1
      Top             =   360
      Width           =   3975
   End
   Begin VB.Label Label2 
      Caption         =   "Quedarse boludeando"
      Height          =   735
      Left            =   240
      TabIndex        =   0
      Top             =   1080
      Width           =   3855
   End
End
Attribute VB_Name = "FrmRetorno"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Khalem

Private Sub Label1_Click()
Call WriteWarpChar(UserName, 1, 60, 50)
End Sub

Private Sub Label2_Click()
Unload Me
End Sub

