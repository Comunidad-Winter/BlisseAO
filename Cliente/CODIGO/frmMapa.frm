VERSION 5.00
Begin VB.Form frmMapa 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   5790
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5025
   ClipControls    =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   386
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   335
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin CLBLISSEAO.BGAOButton imgToogleMap 
      Height          =   285
      Index           =   0
      Left            =   60
      TabIndex        =   1
      Top             =   2340
      Width           =   285
      _ExtentX        =   503
      _ExtentY        =   503
      Caption         =   "<"
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
   Begin CLBLISSEAO.BGAOButton imgToogleMap 
      Height          =   285
      Index           =   1
      Left            =   4680
      TabIndex        =   2
      Top             =   2340
      Visible         =   0   'False
      Width           =   285
      _ExtentX        =   503
      _ExtentY        =   503
      Caption         =   ">"
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
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00000080&
      BorderStyle     =   0  'None
      Height          =   4965
      Left            =   0
      ScaleHeight     =   331
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   335
      TabIndex        =   3
      Top             =   0
      Width           =   5025
      Begin VB.Shape Shape1 
         BorderColor     =   &H000080FF&
         Height          =   495
         Left            =   510
         Top             =   975
         Visible         =   0   'False
         Width           =   510
      End
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Mapa: 1"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   255
      Left            =   0
      TabIndex        =   4
      Top             =   4920
      Width           =   5055
   End
   Begin VB.Image imgCerrar 
      Height          =   255
      Left            =   8040
      MousePointer    =   99  'Custom
      Top             =   240
      Width           =   255
   End
   Begin VB.Label lblTexto 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmMapa.frx":0000
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   615
      Left            =   0
      TabIndex        =   0
      Top             =   5160
      Width           =   5055
   End
End
Attribute VB_Name = "frmMapa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private clsFormulario As clsFormMovementManager

Private Enum eMaps
    ieGeneral
    ieDungeon
End Enum

Private picMaps(1) As Picture

Private CurrentMap As eMaps

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyRight, vbKeyLeft
            ToggleImgMaps
        Case Else
            Unload Me
    End Select
End Sub

Private Sub ToggleImgMaps()
    imgToogleMap(CurrentMap).Visible = False
    
    If CurrentMap = eMaps.ieGeneral Then
        imgCerrar.Visible = False
        CurrentMap = eMaps.ieDungeon
    Else
        imgCerrar.Visible = True
        CurrentMap = eMaps.ieGeneral
    End If
    
    imgToogleMap(CurrentMap).Visible = True
    Picture1.Picture = picMaps(CurrentMap)
End Sub

Private Sub Form_Load()
On Error GoTo Error
    
    '   Handles Form movement (drag and drop).
    Set clsFormulario = New clsFormMovementManager
    clsFormulario.Initialize Me
        
    'Cargamos las imagenes de los mapas
    Set picMaps(eMaps.ieGeneral) = General_Set_GUI("mapageneral")
    Set picMaps(eMaps.ieDungeon) = General_Set_GUI("mapageneral")
    
    imgToogleMap(0).Init s_Small
    imgToogleMap(1).Init s_Small
    '   Imagen de fondo
    CurrentMap = eMaps.ieGeneral
    Picture1.Picture = picMaps(CurrentMap)

    Exit Sub
Error:
    MsgBox Err.Description, vbInformation, "Error: " & Err.Number
    Unload Me
End Sub

Private Sub imgCerrar_Click()
    Unload Me
End Sub

Private Sub imgToogleMap_Click(Index As Integer)
    ToggleImgMaps
End Sub

Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim temp_x As Single
    Dim temp_y As Single
    
    temp_x = X \ 34
    temp_y = Y \ 33
    
    Label1.Caption = "Mapa: " & temp_x + (temp_y) * (Picture1.ScaleWidth \ 33) + 1

End Sub
