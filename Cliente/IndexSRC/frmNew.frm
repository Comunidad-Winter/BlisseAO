VERSION 5.00
Begin VB.Form frmNew 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "AGREGAR GRH"
   ClientHeight    =   2295
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5595
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
   ScaleHeight     =   2295
   ScaleWidth      =   5595
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer2 
      Interval        =   100
      Left            =   4560
      Top             =   120
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   5040
      Top             =   120
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Agregar"
      Height          =   375
      Left            =   3840
      TabIndex        =   16
      Top             =   1800
      Width           =   1575
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Animación"
      Height          =   255
      Index           =   1
      Left            =   1800
      TabIndex        =   15
      Top             =   120
      Width           =   1215
   End
   Begin VB.OptionButton Option1 
      Caption         =   "GRH"
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   14
      Top             =   120
      Value           =   -1  'True
      Width           =   1815
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   1
      Left            =   1320
      TabIndex        =   6
      Text            =   "0"
      Top             =   600
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   2
      Left            =   1440
      TabIndex        =   5
      Text            =   "0"
      Top             =   1440
      Width           =   615
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   3
      Left            =   1440
      TabIndex        =   4
      Text            =   "0"
      Top             =   1800
      Width           =   615
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   4
      Left            =   3120
      TabIndex        =   3
      Text            =   "0"
      Top             =   1440
      Width           =   615
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   5
      Left            =   3120
      TabIndex        =   2
      Text            =   "0"
      Top             =   1800
      Width           =   615
   End
   Begin VB.TextBox Text1 
      Enabled         =   0   'False
      Height          =   375
      Index           =   8
      Left            =   4440
      TabIndex        =   1
      Text            =   "0"
      Top             =   600
      Width           =   855
   End
   Begin VB.TextBox Text1 
      Enabled         =   0   'False
      Height          =   375
      Index           =   9
      Left            =   4440
      TabIndex        =   0
      Text            =   "0"
      Top             =   960
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "(Separe los GRHs con espacios ej: 1 2 3 4)"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000A&
      Height          =   495
      Left            =   360
      TabIndex        =   17
      Top             =   960
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "Gráfico:"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   13
      Top             =   600
      Width           =   1095
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "SPEED:"
      Height          =   255
      Index           =   2
      Left            =   3240
      TabIndex        =   12
      Top             =   600
      Width           =   1095
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "SRC Y:"
      Height          =   255
      Index           =   3
      Left            =   240
      TabIndex        =   11
      Top             =   1800
      Width           =   1095
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "SRC X:"
      Height          =   255
      Index           =   4
      Left            =   240
      TabIndex        =   10
      Top             =   1440
      Width           =   1095
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "DEST X:"
      Height          =   255
      Index           =   5
      Left            =   2160
      TabIndex        =   9
      Top             =   1440
      Width           =   855
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "DEST Y:"
      Height          =   255
      Index           =   6
      Left            =   2160
      TabIndex        =   8
      Top             =   1800
      Width           =   855
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "FRAMES:"
      Height          =   255
      Index           =   7
      Left            =   3240
      TabIndex        =   7
      Top             =   960
      Width           =   1095
   End
End
Attribute VB_Name = "frmNew"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Timer1.Enabled = True
End Sub

Private Sub Option1_Click(Index As Integer)
    Select Case Index
        Case 0
            Label3(1).Caption = "Gráfico:"
            Text1(9).Enabled = False
            Label1.Visible = False
        Case 1
            Label3(1).Caption = "GRHs:"
            Text1(9).Enabled = True
            Label1.Visible = True
    End Select
End Sub

Private Sub Timer1_Timer()
    If frmMain.Index <> 0 Then
        frmMain.Index = 0
        Exit Sub
    End If
    
    If Option1(0).value = True Then
        NUM_GRHS = NUM_GRHS + 1
    
        ReDim Preserve GrhData(0 To NUM_GRHS) As GrhData
        
        frmMain.GRH_LIST.AddItem NUM_GRHS
        
        With GrhData(NUM_GRHS)
            .FileNum = Text1(1).Text
            .sX = Text1(2).Text
            .sY = Text1(3).Text
            .pixelHeight = Text1(5).Text
            .pixelWidth = Text1(4).Text
            .NumFrames = 1
        End With
        
        MsgBox "GRH " & NUM_GRHS & " Creado con Éxito!"
    Else
        NUM_GRHS = NUM_GRHS + 1
    
        ReDim Preserve GrhData(0 To NUM_GRHS) As GrhData
        
        frmMain.GRH_LIST.AddItem NUM_GRHS
        
        With GrhData(NUM_GRHS)
            .NumFrames = Text1(9).Text
            .Speed = (.NumFrames * 1000) / 18
            ReDim .Frames(1 To .NumFrames)
            
            Dim i As Long
            For i = 1 To .NumFrames
                .Frames(i) = ReadField(CInt(i), Text1(1).Text, Asc(" "))
            Next i
        End With
        
        MsgBox "ANIM " & NUM_GRHS & " Creada con Éxito!"
    End If
    
    Timer1.Enabled = False
End Sub

Private Sub Timer2_Timer()
    Text1(8).Text = (Text1(9).Text * 1000) / 18
End Sub
