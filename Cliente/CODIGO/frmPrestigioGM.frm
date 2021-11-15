VERSION 5.00
Begin VB.Form frmPrestigioGM 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Prestigio"
   ClientHeight    =   2505
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4935
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
   ScaleHeight     =   2505
   ScaleWidth      =   4935
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      Caption         =   "Otorgar Prestigio (Solo Canjes):"
      Height          =   975
      Left            =   120
      TabIndex        =   6
      Top             =   1440
      Width           =   4695
      Begin VB.TextBox Text4 
         Height          =   285
         Left            =   120
         TabIndex        =   9
         Text            =   "Standelf"
         Top             =   480
         Width           =   1455
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Left            =   1680
         TabIndex        =   8
         Text            =   "1"
         Top             =   480
         Width           =   1455
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Dar Prestigio"
         Height          =   495
         Left            =   3240
         TabIndex        =   7
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Usuario:"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Cantidad:"
         Height          =   255
         Left            =   1680
         TabIndex        =   10
         Top             =   240
         Width           =   1455
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Otorgar Prestigio(Rep/Canje):"
      Height          =   1215
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4695
      Begin VB.CommandButton Command1 
         Caption         =   "Dar Prestigio"
         Height          =   495
         Left            =   3240
         TabIndex        =   5
         Top             =   240
         Width           =   1335
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   1680
         TabIndex        =   4
         Text            =   "1"
         Top             =   480
         Width           =   1455
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   120
         TabIndex        =   1
         Text            =   "Standelf"
         Top             =   480
         Width           =   1455
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Nota: Para restar prestigio de Reputación poner -"
         ForeColor       =   &H00404040&
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   840
         Width           =   4455
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Cantidad:"
         Height          =   255
         Left            =   1680
         TabIndex        =   3
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Usuario:"
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   1455
      End
   End
End
Attribute VB_Name = "frmPrestigioGM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    If ValidNumber(Val(Text2), eNumber_Types.ent_Integer) Then
        Call WritePrestigio(Text1.Text, Text2.Text, 1)
    Else
        'No es numerico
        Call ShowConsoleMsg("Valor inválido.")
    End If
End Sub


Private Sub Command2_Click()
    If ValidNumber(Val(Text3), eNumber_Types.ent_Integer) Then
        Call WritePrestigio(Text4.Text, Text3.Text, 2)
    Else
        'No es numerico
        Call ShowConsoleMsg("Valor inválido.")
    End If
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
    Case 13
        KeyAscii = 0
    Case 8, 45, 48 To 57
    Case Else
        KeyAscii = 0
        Beep
    End Select
End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
    Case 13
        KeyAscii = 0
    Case 8, 48 To 57
    Case Else
        KeyAscii = 0
        Beep
    End Select
End Sub
