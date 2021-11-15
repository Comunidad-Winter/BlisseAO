VERSION 5.00
Begin VB.Form frmNewCAS 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Creador de Cascos/Cabezas"
   ClientHeight    =   2310
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4575
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
   ScaleHeight     =   154
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   305
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   3960
      Top             =   240
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Crear"
      Height          =   375
      Left            =   2640
      TabIndex        =   11
      Top             =   1800
      Width           =   1815
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Cabezas"
      Height          =   255
      Index           =   0
      Left            =   1560
      TabIndex        =   5
      Top             =   240
      Value           =   -1  'True
      Width           =   1695
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Cascos"
      Height          =   255
      Index           =   1
      Left            =   1560
      TabIndex        =   4
      Top             =   480
      Width           =   1695
   End
   Begin VB.TextBox Text4 
      Height          =   360
      Index           =   0
      Left            =   1320
      TabIndex        =   3
      Text            =   "0"
      Top             =   960
      Width           =   855
   End
   Begin VB.TextBox Text4 
      Height          =   360
      Index           =   1
      Left            =   1320
      TabIndex        =   2
      Text            =   "0"
      Top             =   1320
      Width           =   855
   End
   Begin VB.TextBox Text4 
      Height          =   360
      Index           =   2
      Left            =   3600
      TabIndex        =   1
      Text            =   "0"
      Top             =   960
      Width           =   855
   End
   Begin VB.TextBox Text4 
      Height          =   360
      Index           =   3
      Left            =   3600
      TabIndex        =   0
      Text            =   "0"
      Top             =   1320
      Width           =   855
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "Frente:"
      Height          =   255
      Index           =   13
      Left            =   2280
      TabIndex        =   10
      Top             =   960
      Width           =   1215
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "Izquierda:"
      Height          =   255
      Index           =   14
      Left            =   120
      TabIndex        =   9
      Top             =   1320
      Width           =   1095
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "Derecha:"
      Height          =   255
      Index           =   15
      Left            =   120
      TabIndex        =   8
      Top             =   960
      Width           =   1095
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "Espalda:"
      Height          =   255
      Index           =   16
      Left            =   2280
      TabIndex        =   7
      Top             =   1320
      Width           =   1215
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "Crear:"
      Height          =   255
      Index           =   17
      Left            =   120
      TabIndex        =   6
      Top             =   240
      Width           =   1335
   End
End
Attribute VB_Name = "frmNewCas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
    Timer1.Enabled = True
End Sub

Private Sub Timer1_Timer()
    If frmMain.Index <> 0 Then
        frmMain.Index = 0
        Exit Sub
    End If
    
    
    If Option1(0).value = True Then
        NUM_HEA = NUM_HEA + 1
        ReDim Preserve HeadData(0 To NUM_HEA) As HeadData
                HeadData(NUM_HEA).Head(1).GrhIndex = Text4(3).Text
                HeadData(NUM_HEA).Head(2).GrhIndex = Text4(1).Text
                HeadData(NUM_HEA).Head(3).GrhIndex = Text4(2).Text
                HeadData(NUM_HEA).Head(4).GrhIndex = Text4(0).Text
                
             MsgBox "CABEZA " & NUM_HEA & " Creada con Éxito!"
             
             If frmMain.Option1(0).value = True Then frmMain.HHH_LIST.AddItem NUM_HEA
    Else
        NUM_CAS = NUM_CAS + 1
        ReDim Preserve CascoAnimData(0 To NUM_HEA) As HeadData
                CascoAnimData(NUM_CAS).Head(1).GrhIndex = Text4(3).Text
                CascoAnimData(NUM_CAS).Head(2).GrhIndex = Text4(1).Text
                CascoAnimData(NUM_CAS).Head(3).GrhIndex = Text4(2).Text
                CascoAnimData(NUM_CAS).Head(4).GrhIndex = Text4(0).Text
                
             MsgBox "CASCO " & NUM_CAS & " Creado con Éxito!"
             
             If frmMain.Option1(1).value = True Then frmMain.HHH_LIST.AddItem NUM_CAS
    End If
    
    Timer1.Enabled = False
End Sub
