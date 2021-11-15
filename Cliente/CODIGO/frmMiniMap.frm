VERSION 5.00
Begin VB.Form FrmMiniMap 
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   2055
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2220
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MousePointer    =   15  'Size All
   Picture         =   "frmMiniMap.frx":0000
   ScaleHeight     =   137
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   148
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      CausesValidation=   0   'False
      Height          =   450
      Index           =   2
      Left            =   1770
      MousePointer    =   1  'Arrow
      Picture         =   "frmMiniMap.frx":1973
      ScaleHeight     =   30
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   30
      TabIndex        =   5
      Top             =   480
      Width           =   450
   End
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      CausesValidation=   0   'False
      Height          =   330
      Index           =   1
      Left            =   1485
      MousePointer    =   1  'Arrow
      Picture         =   "frmMiniMap.frx":1E9F
      ScaleHeight     =   22
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   22
      TabIndex        =   4
      Top             =   1710
      Width           =   330
   End
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      CausesValidation=   0   'False
      Height          =   330
      Index           =   0
      Left            =   1755
      MousePointer    =   1  'Arrow
      Picture         =   "frmMiniMap.frx":2284
      ScaleHeight     =   22
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   22
      TabIndex        =   3
      Top             =   1455
      Width           =   330
   End
   Begin VB.PictureBox MiniMap 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   1515
      Left            =   495
      MousePointer    =   15  'Size All
      ScaleHeight     =   101
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   101
      TabIndex        =   0
      Top             =   405
      Width           =   1515
      Begin VB.Image UserPosition 
         Height          =   135
         Left            =   75
         Picture         =   "frmMiniMap.frx":2641
         Top             =   45
         Width           =   135
      End
      Begin VB.Shape UserArea 
         BorderColor     =   &H00808080&
         FillColor       =   &H00808080&
         Height          =   270
         Left            =   0
         Shape           =   3  'Circle
         Top             =   0
         Width           =   270
      End
      Begin VB.Label Label1 
         Caption         =   "Label1"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   135
         Left            =   1560
         TabIndex        =   1
         Top             =   1560
         Visible         =   0   'False
         Width           =   735
      End
   End
   Begin VB.Image Image2 
      Height          =   135
      Index           =   3
      Left            =   120
      Picture         =   "frmMiniMap.frx":2767
      Top             =   1320
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Image Image2 
      Height          =   135
      Index           =   2
      Left            =   120
      Picture         =   "frmMiniMap.frx":288D
      Top             =   1080
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Image Image2 
      Height          =   135
      Index           =   1
      Left            =   120
      Picture         =   "frmMiniMap.frx":29B3
      Top             =   840
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Image Image2 
      Height          =   135
      Index           =   0
      Left            =   120
      Picture         =   "frmMiniMap.frx":2ACF
      Top             =   600
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Image Image1 
      Height          =   180
      Left            =   1965
      Top             =   90
      Width           =   195
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Nombre Mapa"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   195
      Left            =   120
      TabIndex        =   2
      Top             =   75
      Width           =   1815
   End
End
Attribute VB_Name = "FrmMiniMap"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private clsFormulario As clsFormMovementManager

Private Sub Form_Load()
    '   Handles Form movement (drag and drop).
    Set clsFormulario = New clsFormMovementManager
    clsFormulario.Initialize Me, MiniMap, 150
    
    Round_Picture MiniMap, 100
    Round_Picture Picture1(0), 110
    Round_Picture Picture1(1), 110
    Round_Picture Picture1(2), 110
    
    InitializeSurfaceCapture Me
    CreateSurfacefromMask Me
    ReleaseSurfaceCapture Me
    
    If Not frmMain.Visible Then Me.Hide
End Sub

Private Sub Image1_Click()
    Settings.MiniMap = False
    If frmOpciones.Visible = True Then frmOpciones.ChkMap.value = Unchecked
    Unload Me
End Sub

Private Sub MiniMap_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbRightButton And esGM(UserCharIndex) And Not UserMeditar Then
        Call WriteWarpChar("YO", UserMap, CByte(X), CByte(Y))
    End If
End Sub

Private Sub MiniMap_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    frmMain.SetFocus
End Sub

Private Sub Picture1_Click(index As Integer)
    Select Case index
        Case 2
            Call FrmMiniMap.Show(vbModeless, frmMain)
        Case 1
       '     Engine_Set_Zoom_Out
        Case 0
       '     Engine_Set_Zoom_In
    End Select
    
End Sub
