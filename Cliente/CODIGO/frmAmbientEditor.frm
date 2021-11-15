VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmAmbientEditor 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Editor de Ambiente"
   ClientHeight    =   6525
   ClientLeft      =   150
   ClientTop       =   720
   ClientWidth     =   4440
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
   ScaleHeight     =   435
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   296
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture2 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4455
      Left            =   960
      ScaleHeight     =   297
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   217
      TabIndex        =   9
      Top             =   120
      Visible         =   0   'False
      Width           =   3255
      Begin VB.CheckBox Check5 
         Caption         =   "Tiene Tormentas de arena"
         Enabled         =   0   'False
         Height          =   495
         Left            =   120
         TabIndex        =   47
         Top             =   2880
         Width           =   2415
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Usar Niebla en el Mapa"
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   120
         Width           =   2655
      End
      Begin VB.HScrollBar HScroll1 
         Enabled         =   0   'False
         Height          =   615
         Left            =   120
         Max             =   150
         Min             =   25
         TabIndex        =   13
         Top             =   720
         Value           =   25
         Width           =   2655
      End
      Begin VB.CommandButton Command9 
         Caption         =   "Aplicar"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   12
         Top             =   1440
         Width           =   2655
      End
      Begin VB.CheckBox Check2 
         Caption         =   "Tiene Nevadas"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   2520
         Width           =   2415
      End
      Begin VB.CheckBox Check3 
         Caption         =   "Tiene Lluvias"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   2160
         Width           =   2415
      End
      Begin VB.Label Label2 
         Caption         =   "Grado de Niebla"
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   480
         Width           =   2415
      End
   End
   Begin VB.PictureBox Picture4 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5655
      Left            =   960
      ScaleHeight     =   377
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   217
      TabIndex        =   16
      Top             =   240
      Visible         =   0   'False
      Width           =   3255
      Begin VB.CommandButton Command4 
         Caption         =   "Remover"
         Height          =   375
         Index           =   0
         Left            =   1560
         TabIndex        =   20
         Top             =   480
         Width           =   1095
      End
      Begin VB.TextBox Text5 
         Alignment       =   2  'Center
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   23
         Text            =   "0"
         Top             =   120
         Width           =   1095
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Aplicar"
         Height          =   375
         Index           =   1
         Left            =   480
         TabIndex        =   22
         Top             =   480
         Width           =   1095
      End
      Begin VB.ListBox List1 
         Appearance      =   0  'Flat
         Height          =   4125
         ItemData        =   "frmAmbientEditor.frx":0000
         Left            =   120
         List            =   "frmAmbientEditor.frx":0043
         TabIndex        =   18
         Top             =   1320
         Width           =   3015
      End
      Begin VB.Label Label6 
         Caption         =   "Theme Actual:"
         Height          =   255
         Left            =   120
         TabIndex        =   21
         Top             =   120
         Width           =   1335
      End
      Begin VB.Label Label5 
         Caption         =   "Themes:"
         Height          =   255
         Left            =   120
         TabIndex        =   19
         Top             =   960
         Width           =   1455
      End
   End
   Begin VB.PictureBox Picture5 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5655
      Left            =   960
      ScaleHeight     =   377
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   217
      TabIndex        =   17
      Top             =   240
      Width           =   3255
      Begin VB.OptionButton Option2 
         Caption         =   "3x3"
         Height          =   255
         Index           =   3
         Left            =   2160
         TabIndex        =   28
         Top             =   480
         Width           =   855
      End
      Begin VB.CheckBox Check4 
         Caption         =   "Activar editor de áreas"
         Height          =   255
         Left            =   120
         TabIndex        =   27
         Top             =   120
         Width           =   2535
      End
      Begin VB.OptionButton Option2 
         Caption         =   "2x2"
         Height          =   255
         Index           =   2
         Left            =   1200
         TabIndex        =   26
         Top             =   480
         Width           =   735
      End
      Begin VB.OptionButton Option2 
         Caption         =   "1 Tile"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   25
         Top             =   480
         Width           =   855
      End
      Begin VB.ListBox List2 
         Appearance      =   0  'Flat
         Height          =   4125
         ItemData        =   "frmAmbientEditor.frx":0148
         Left            =   120
         List            =   "frmAmbientEditor.frx":0191
         TabIndex        =   0
         Top             =   1320
         Width           =   3015
      End
      Begin VB.Label Label7 
         Caption         =   "Areas:"
         Height          =   255
         Left            =   120
         TabIndex        =   24
         Top             =   960
         Width           =   1455
      End
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4935
      Left            =   960
      ScaleHeight     =   329
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   217
      TabIndex        =   1
      Top             =   240
      Visible         =   0   'False
      Width           =   3255
      Begin VB.PictureBox Picture9 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   2655
         Left            =   2820
         ScaleHeight     =   2655
         ScaleWidth      =   315
         TabIndex        =   43
         Top             =   2100
         Width           =   315
      End
      Begin VB.PictureBox Picture11 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2790
         Left            =   120
         Picture         =   "frmAmbientEditor.frx":0286
         ScaleHeight     =   186
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   174
         TabIndex        =   45
         Top             =   2040
         Width           =   2610
      End
      Begin VB.PictureBox Picture10 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   2775
         Left            =   2760
         ScaleHeight     =   2775
         ScaleWidth      =   435
         TabIndex        =   44
         Top             =   2040
         Width           =   435
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Usar Luz del Ambiente"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   7
         Top             =   120
         Value           =   -1  'True
         Width           =   2295
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Usar ambiente propio"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   6
         Top             =   480
         Width           =   2175
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         BorderStyle     =   0  'None
         Height          =   285
         Index           =   0
         Left            =   360
         TabIndex        =   5
         Text            =   "255"
         Top             =   840
         Width           =   495
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         BorderStyle     =   0  'None
         Height          =   285
         Index           =   1
         Left            =   1320
         TabIndex        =   4
         Text            =   "255"
         Top             =   840
         Width           =   495
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         BorderStyle     =   0  'None
         Height          =   285
         Index           =   2
         Left            =   2160
         TabIndex        =   3
         Text            =   "255"
         Top             =   840
         Width           =   495
      End
      Begin VB.CommandButton Command7 
         Caption         =   "Aplicar"
         Height          =   375
         Left            =   120
         TabIndex        =   2
         Top             =   1320
         Width           =   2895
      End
      Begin VB.Label Label9 
         Caption         =   "Tomar Color:"
         Height          =   255
         Left            =   120
         TabIndex        =   46
         Top             =   1800
         Width           =   1455
      End
      Begin VB.Label Label1 
         Caption         =   "R:           G:           B:"
         Height          =   375
         Left            =   120
         TabIndex        =   8
         Top             =   840
         Width           =   2535
      End
   End
   Begin VB.PictureBox Picture3 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5655
      Left            =   960
      ScaleHeight     =   377
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   217
      TabIndex        =   30
      Top             =   240
      Visible         =   0   'False
      Width           =   3255
      Begin VB.PictureBox Picture8 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   2655
         Left            =   2820
         ScaleHeight     =   2655
         ScaleWidth      =   315
         TabIndex        =   42
         Top             =   2820
         Width           =   315
      End
      Begin VB.PictureBox Picture7 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   2775
         Left            =   2760
         ScaleHeight     =   2775
         ScaleWidth      =   435
         TabIndex        =   41
         Top             =   2760
         Width           =   435
      End
      Begin VB.CommandButton Command10 
         Caption         =   "Borrar Luz Actual"
         Height          =   375
         Left            =   120
         TabIndex        =   31
         Top             =   1920
         Width           =   2895
      End
      Begin VB.PictureBox Picture6 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2790
         Left            =   120
         Picture         =   "frmAmbientEditor.frx":7A61
         ScaleHeight     =   186
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   174
         TabIndex        =   39
         Top             =   2760
         Width           =   2610
      End
      Begin VB.HScrollBar HScroll2 
         Height          =   495
         Left            =   960
         Max             =   10
         Min             =   1
         TabIndex        =   36
         Top             =   840
         Value           =   1
         Width           =   2055
      End
      Begin VB.TextBox Text4 
         Alignment       =   2  'Center
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   2160
         TabIndex        =   35
         Text            =   "255"
         Top             =   120
         Width           =   495
      End
      Begin VB.TextBox Text3 
         Alignment       =   2  'Center
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   1320
         TabIndex        =   34
         Text            =   "255"
         Top             =   120
         Width           =   495
      End
      Begin VB.TextBox Text2 
         Alignment       =   2  'Center
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   465
         TabIndex        =   33
         Text            =   "255"
         Top             =   120
         Width           =   495
      End
      Begin VB.CommandButton Command8 
         Caption         =   "Crear Luz en Posición Actual"
         Height          =   375
         Left            =   120
         TabIndex        =   32
         Top             =   1560
         Width           =   2895
      End
      Begin VB.Label Label8 
         Caption         =   "Tomar Color:"
         Height          =   255
         Left            =   120
         TabIndex        =   40
         Top             =   2520
         Width           =   1455
      End
      Begin VB.Label Label4 
         Caption         =   "Radio de Luz:"
         Height          =   495
         Left            =   120
         TabIndex        =   38
         Top             =   840
         Width           =   735
      End
      Begin VB.Label Label3 
         Caption         =   "R:           G:           B:"
         Height          =   375
         Left            =   120
         TabIndex        =   37
         Top             =   120
         Width           =   2535
      End
   End
   Begin MSComctlLib.TabStrip TabStrip1 
      Height          =   6375
      Left            =   120
      TabIndex        =   29
      Top             =   0
      Width           =   4215
      _ExtentX        =   7435
      _ExtentY        =   11245
      MultiRow        =   -1  'True
      Placement       =   2
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   6
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Luz de Ambiente"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Luces del Mapa"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Música de Ambiente"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab4 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Meteorología"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab5 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Áreas del Mapa"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab6 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Editor de Terreno"
            ImageVarType    =   2
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Menu mnuReload 
      Caption         =   "Recargar Ambiente"
   End
   Begin VB.Menu mnuSave 
      Caption         =   "Guardar Ambiente"
   End
End
Attribute VB_Name = "frmAmbientEditor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' ProyectoAO Documentación ***********************************
' Autor: Standelf
' Descripción: Formulario para editar el Ambiente
' Última modificación: 07/11/2012
'*********************************************************


Option Explicit
Public Area_Range As Byte
Private MouX As Single
Private MouY As Single

Private Sub Check1_Click()
    If Check1.value = Checked Then
        HScroll1.Enabled = True
    Else
        HScroll1.Enabled = False
    End If
End Sub

Private Sub Check2_Click()
    If Check2.value = Checked Then
        CurMapAmbient.Snow = True
    Else
        CurMapAmbient.Snow = False
        If Effect(WeatherEffectIndex).EffectNum = EffectNum_Snow Then Effect_Kill WeatherEffectIndex
    End If
End Sub

Private Sub Check3_Click()
    If Check3.value = Checked Then
        CurMapAmbient.Rain = True
    Else
        CurMapAmbient.Rain = False
    End If
End Sub

Private Sub Check4_Click()
    If Check4.value = Checked Then
        Setting_Map_Areas = True
    Else
        Setting_Map_Areas = False
    End If
End Sub

Private Sub Command4_Click(Index As Integer)
    Select Case Index
        Case 0
            CurMapAmbient.Music = 0
            Audio.MP3_Stop
        Case 1
            CurMapAmbient.Music = Val(List1.ListIndex + 1)
            Play_MP3 CurMapAmbient.Music
    End Select
End Sub

Private Sub Command7_Click()
    If Option1(0).value = True Then
        CurMapAmbient.UseDayAmbient = True
            CurMapAmbient.OwnAmbientLight.a = 255
            CurMapAmbient.OwnAmbientLight.r = 0
            CurMapAmbient.OwnAmbientLight.g = 0
            CurMapAmbient.OwnAmbientLight.b = 0
    Else
            CurMapAmbient.UseDayAmbient = False
            CurMapAmbient.OwnAmbientLight.a = 255
            CurMapAmbient.OwnAmbientLight.r = Val(Text1(0).Text)
            CurMapAmbient.OwnAmbientLight.g = Val(Text1(1).Text)
            CurMapAmbient.OwnAmbientLight.b = Val(Text1(2).Text)
    End If
    
    DoEvents
    
    Call Ambient_Aply_OwnAmbient
End Sub

Private Sub Command8_Click()
    CurMapAmbient.MapBlocks(UserPos.X, UserPos.Y).Light.b = Val(Text4.Text)
    CurMapAmbient.MapBlocks(UserPos.X, UserPos.Y).Light.g = Val(Text3.Text)
    CurMapAmbient.MapBlocks(UserPos.X, UserPos.Y).Light.r = Val(Text2.Text)
    CurMapAmbient.MapBlocks(UserPos.X, UserPos.Y).Light.Range = Val(HScroll2.value)
    
    Create_Light_To_Map UserPos.X, UserPos.Y, Val(HScroll2.value), Val(Text2.Text), Val(Text3.Text), Val(Text4.Text)
End Sub

Private Sub Command9_Click()
    If Check1.value = Unchecked Then
        CurMapAmbient.Fog = -1
    Else
        CurMapAmbient.Fog = Val(HScroll1.value)
    End If
End Sub

Private Sub Command10_Click()
    CurMapAmbient.MapBlocks(UserPos.X, UserPos.Y).Light.b = 0
    CurMapAmbient.MapBlocks(UserPos.X, UserPos.Y).Light.g = 0
    CurMapAmbient.MapBlocks(UserPos.X, UserPos.Y).Light.r = 0
    CurMapAmbient.MapBlocks(UserPos.X, UserPos.Y).Light.Range = 0
    
    Call Delete_Light_To_Map(UserPos.X, UserPos.Y)
End Sub

Private Sub Form_Load()
    TabStrip1.Tabs(5).Selected = True
End Sub

Private Sub mnuReload_Click()
    Ambient_Init UserMap
End Sub

Private Sub mnusave_Click()
    Ambient_Save UserMap
    DoEvents
    
    Ambient_Init UserMap
End Sub

Private Sub Option2_Click(Index As Integer)
    Area_Range = CByte(Index)
End Sub

Private Sub Picture11_Click()
    Dim Color As Long, Color2 As D3DCOLORVALUE
        Color = GetPixel(Picture11.hDC, MouX, MouY)

        Color2.r = Color Mod 256
        Color2.g = (Color \ 256) Mod 256
        Color2.b = (Color \ 256 \ 256) Mod 256
        
        Text1(0).Text = Color2.r
        Text1(1).Text = Color2.g
        Text1(2).Text = Color2.b
        
        Picture9.BackColor = RGB(Color2.r, Color2.g, Color2.b)
End Sub

Private Sub Picture11_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    MouX = X
    MouY = Y
End Sub

Private Sub Picture6_Click()
    Dim Color As Long, Color2 As D3DCOLORVALUE
        Color = GetPixel(Picture6.hDC, MouX, MouY)

        Color2.r = Color Mod 256
        Color2.g = (Color \ 256) Mod 256
        Color2.b = (Color \ 256 \ 256) Mod 256
        
        Text2.Text = Color2.r
        Text3.Text = Color2.g
        Text4.Text = Color2.b
        
        Picture8.BackColor = RGB(Color2.r, Color2.g, Color2.b)
End Sub

Private Sub Picture6_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    MouX = X
    MouY = Y
End Sub

Private Sub TabStrip1_Click()
    Picture3.Visible = False    '  Creador de luces
    Picture1.Visible = False    '  Luz de Ambiente
    Picture5.Visible = False    '  Areas
    Picture4.Visible = False    '  MP3
    Picture2.Visible = False    '  Meteorologia
    
    Select Case TabStrip1.SelectedItem.Index
        Case 1
            Picture1.Visible = True
        Case 2
            Picture3.Visible = True
        Case 3
            Picture4.Visible = True    '  MP3
        Case 4
            Picture2.Visible = True    '  Meteorologia
        Case 5
            Picture5.Visible = True    '  Areas
        Case 6
            MsgBox "No disponible"
    End Select
End Sub

Private Sub Text1_Change(Index As Integer)
    Picture9.BackColor = RGB(Text1(0), Text1(2), Text1(1))
End Sub

Private Sub Text2_Change()
    Picture8.BackColor = RGB(Text2, Text3, Text4)
End Sub
