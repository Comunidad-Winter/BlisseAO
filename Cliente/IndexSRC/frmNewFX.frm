VERSION 5.00
Begin VB.Form frmNewFX 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Creador de FXs"
   ClientHeight    =   2055
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   3015
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
   ScaleHeight     =   137
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   201
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Crear FX"
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   1560
      Width           =   2775
   End
   Begin VB.TextBox Text1 
      Height          =   405
      Index           =   2
      Left            =   1440
      TabIndex        =   5
      Text            =   "0"
      Top             =   120
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Height          =   405
      Index           =   1
      Left            =   1440
      TabIndex        =   4
      Text            =   "0"
      Top             =   1080
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Height          =   405
      Index           =   0
      Left            =   1440
      TabIndex        =   3
      Text            =   "0"
      Top             =   600
      Width           =   1455
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "OffSet Y:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   2
      Top             =   1080
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "OffSet X:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Animación:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "frmNewFX"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
NUM_FXS = NUM_FXS + 1
    ReDim Preserve FxData(1 To NUM_FXS)
    
    With FxData(NUM_FXS)
        .Animacion = Text1(2).Text
        .OffsetX = Text1(0).Text
        .OffsetY = Text1(1).Text
    End With
    
    MsgBox "FX " & NUM_FXS & " Creada con éxito."
    frmMain.FXS_LIST.AddItem NUM_FXS
    
    Unload Me
End Sub
