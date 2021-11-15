VERSION 5.00
Begin VB.Form frmEditANIM 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Editar Animación"
   ClientHeight    =   1890
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   3165
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   126
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   211
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   2400
      Top             =   600
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   9
      Left            =   1320
      TabIndex        =   3
      Text            =   "0"
      Top             =   960
      Width           =   855
   End
   Begin VB.TextBox Text1 
      Enabled         =   0   'False
      Height          =   375
      Index           =   8
      Left            =   1320
      TabIndex        =   2
      Text            =   "0"
      Top             =   600
      Width           =   855
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   1
      Left            =   1320
      TabIndex        =   1
      Text            =   "0"
      Top             =   120
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Actualizar"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   1440
      Width           =   2895
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "FRAMES:"
      Height          =   255
      Index           =   7
      Left            =   120
      TabIndex        =   6
      Top             =   960
      Width           =   1095
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "SPEED:"
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   5
      Top             =   600
      Width           =   1095
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "GRHs:"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   1095
   End
End
Attribute VB_Name = "frmEditANIM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
frmMain.Index = 0


        With GrhData(frmMain.Text1(0))
            .NumFrames = Text1(9).Text
            .Speed = (.NumFrames * 1000) / 18
            ReDim .Frames(1 To .NumFrames)
            
            Dim i As Long
            For i = 1 To .NumFrames
                .Frames(i) = ReadField(CInt(i), Text1(1).Text, Asc(" "))
            Next i
        End With
        
        MsgBox "ACTUALIZADO"
        
End Sub

Private Sub Timer1_Timer()
    Text1(8).Text = (Text1(9).Text * 1000) / 18
End Sub
