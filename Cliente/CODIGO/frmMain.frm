VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.ocx"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.ocx"
Begin VB.Form frmMain 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000040&
   BorderStyle     =   0  'None
   ClientHeight    =   9510
   ClientLeft      =   360
   ClientTop       =   300
   ClientWidth     =   16425
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00000000&
   Icon            =   "frmMain.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   Picture         =   "frmMain.frx":0CCA
   ScaleHeight     =   634
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1095
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Left            =   4320
      TabIndex        =   46
      Top             =   240
      Width           =   735
   End
   Begin VB.TextBox SendTxt 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   240
      Left            =   390
      MaxLength       =   160
      MultiLine       =   -1  'True
      TabIndex        =   2
      TabStop         =   0   'False
      ToolTipText     =   "Chat"
      Top             =   2145
      Visible         =   0   'False
      Width           =   7590
   End
   Begin VB.PictureBox REspada 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      CausesValidation=   0   'False
      ClipControls    =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   360
      Left            =   2400
      ScaleHeight     =   24
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   24
      TabIndex        =   37
      ToolTipText     =   "Espada"
      Top             =   8400
      Width           =   360
   End
   Begin VB.PictureBox REscudo 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      CausesValidation=   0   'False
      ClipControls    =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   360
      Left            =   1950
      ScaleHeight     =   24
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   24
      TabIndex        =   36
      ToolTipText     =   "Escudo"
      Top             =   8400
      Width           =   360
   End
   Begin VB.PictureBox RCasco 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      CausesValidation=   0   'False
      ClipControls    =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   360
      Left            =   2850
      ScaleHeight     =   24
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   24
      TabIndex        =   35
      ToolTipText     =   "Casco"
      Top             =   8400
      Width           =   360
   End
   Begin VB.PictureBox RTunica 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      CausesValidation=   0   'False
      ClipControls    =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   360
      Left            =   1500
      ScaleHeight     =   24
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   24
      TabIndex        =   34
      ToolTipText     =   "Armadura"
      Top             =   8400
      Width           =   360
   End
   Begin VB.ListBox hlst 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1980
      Left            =   8520
      TabIndex        =   33
      TabStop         =   0   'False
      Top             =   2760
      Visible         =   0   'False
      Width           =   3240
   End
   Begin VB.PictureBox invInfo 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1050
      Left            =   6300
      ScaleHeight     =   70
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   121
      TabIndex        =   30
      Top             =   2400
      Visible         =   0   'False
      Width           =   1815
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Datos del Item"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000080FF&
         Height          =   675
         Index           =   1
         Left            =   90
         TabIndex        =   32
         Top             =   300
         Width           =   1635
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Nombre Item"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000080FF&
         Height          =   195
         Index           =   0
         Left            =   90
         TabIndex        =   31
         Top             =   75
         Width           =   1635
      End
   End
   Begin VB.PictureBox picInv 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      CausesValidation=   0   'False
      ClipControls    =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   2760
      Left            =   8520
      ScaleHeight     =   184
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   184
      TabIndex        =   29
      Top             =   2760
      Width           =   2760
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
      Height          =   1500
      Left            =   10230
      MousePointer    =   15  'Size All
      ScaleHeight     =   100
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   100
      TabIndex        =   25
      Top             =   7020
      Width           =   1500
      Begin VB.Shape UserPosition 
         BackColor       =   &H000000FF&
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         FillColor       =   &H000000FF&
         Height          =   60
         Left            =   120
         Shape           =   3  'Circle
         Top             =   120
         Width           =   60
      End
      Begin VB.Label Label7 
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
         TabIndex        =   26
         Top             =   1560
         Visible         =   0   'False
         Width           =   735
      End
   End
   Begin VB.Timer TrainingMacro 
      Enabled         =   0   'False
      Interval        =   3121
      Left            =   12120
      Top             =   600
   End
   Begin VB.Timer MacroTrabajo 
      Enabled         =   0   'False
      Left            =   12120
      Top             =   1080
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   12120
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   12120
      Top             =   1560
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.PictureBox picSM 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   360
      Index           =   4
      Left            =   11370
      MousePointer    =   99  'Custom
      ScaleHeight     =   24
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   24
      TabIndex        =   21
      Top             =   5955
      Width           =   360
   End
   Begin RichTextLib.RichTextBox RecTxt 
      Height          =   1365
      Left            =   135
      TabIndex        =   0
      TabStop         =   0   'False
      ToolTipText     =   "Mensajes del servidor"
      Top             =   525
      Width           =   8160
      _ExtentX        =   14393
      _ExtentY        =   2408
      _Version        =   393217
      BackColor       =   0
      BorderStyle     =   0
      HideSelection   =   0   'False
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      DisableNoScroll =   -1  'True
      Appearance      =   0
      AutoVerbMenu    =   -1  'True
      TextRTF         =   $"frmMain.frx":490F9
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.PictureBox shpPara 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   0  'None
      Height          =   75
      Left            =   7635
      ScaleHeight     =   75
      ScaleWidth      =   480
      TabIndex        =   16
      Top             =   8130
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.PictureBox imgPara 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   480
      Left            =   7650
      Picture         =   "frmMain.frx":49176
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   15
      Top             =   7635
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.PictureBox picSM 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   360
      Index           =   3
      Left            =   10860
      MousePointer    =   99  'Custom
      ScaleHeight     =   24
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   24
      TabIndex        =   12
      Top             =   5955
      Width           =   360
   End
   Begin VB.PictureBox picSM 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   360
      Index           =   2
      Left            =   10860
      MousePointer    =   99  'Custom
      ScaleHeight     =   24
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   24
      TabIndex        =   11
      Top             =   5955
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.PictureBox picSM 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   360
      Index           =   1
      Left            =   11370
      MousePointer    =   99  'Custom
      ScaleHeight     =   24
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   24
      TabIndex        =   10
      Top             =   6465
      Width           =   360
   End
   Begin VB.PictureBox picSM 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   360
      Index           =   0
      Left            =   10860
      MousePointer    =   99  'Custom
      ScaleHeight     =   24
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   24
      TabIndex        =   9
      Top             =   6465
      Width           =   360
   End
   Begin VB.PictureBox MainViewPic 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00400040&
      BorderStyle     =   0  'None
      Height          =   6240
      Left            =   135
      MousePointer    =   99  'Custom
      ScaleHeight     =   412
      ScaleMode       =   0  'User
      ScaleWidth      =   544
      TabIndex        =   14
      Top             =   2025
      Width           =   8160
   End
   Begin VB.TextBox SendCMSTXT 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C0C0&
      Height          =   240
      Left            =   390
      MaxLength       =   160
      MultiLine       =   -1  'True
      TabIndex        =   1
      TabStop         =   0   'False
      ToolTipText     =   "Chat"
      Top             =   2145
      Visible         =   0   'False
      Width           =   7590
   End
   Begin RichTextLib.RichTextBox GuildTxt 
      Height          =   1365
      Left            =   135
      TabIndex        =   17
      TabStop         =   0   'False
      ToolTipText     =   "Mensajes del Clan"
      Top             =   525
      Visible         =   0   'False
      Width           =   8160
      _ExtentX        =   14393
      _ExtentY        =   2408
      _Version        =   393217
      BackColor       =   0
      BorderStyle     =   0
      HideSelection   =   0   'False
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      DisableNoScroll =   -1  'True
      Appearance      =   0
      AutoVerbMenu    =   -1  'True
      TextRTF         =   $"frmMain.frx":494FD
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Quest"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   330
      Left            =   2775
      TabIndex        =   45
      Top             =   135
      Width           =   855
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "D: N/A"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   165
      Left            =   435
      TabIndex        =   44
      ToolTipText     =   "Fuerza"
      Top             =   8775
      Width           =   990
   End
   Begin VB.Label lblStrg 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "00"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   165
      Left            =   435
      TabIndex        =   43
      ToolTipText     =   "Fuerza"
      Top             =   8415
      Width           =   990
   End
   Begin VB.Label lblDext 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "00"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   165
      Left            =   495
      TabIndex        =   42
      ToolTipText     =   "Agilidad"
      Top             =   8595
      Width           =   855
   End
   Begin VB.Label lblWeapon 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0/0"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   165
      Left            =   2385
      TabIndex        =   41
      ToolTipText     =   "Arma"
      Top             =   8790
      Width           =   390
   End
   Begin VB.Label lblShielder 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0/0"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   165
      Left            =   1935
      TabIndex        =   40
      ToolTipText     =   "Escudo"
      Top             =   8790
      Width           =   390
   End
   Begin VB.Label lblHelm 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0/0"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   165
      Left            =   2835
      TabIndex        =   39
      ToolTipText     =   "Casco"
      Top             =   8790
      Width           =   390
   End
   Begin VB.Label lblArmor 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0/0"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   165
      Left            =   1485
      TabIndex        =   38
      ToolTipText     =   "Armadura"
      Top             =   8790
      Width           =   390
   End
   Begin VB.Image cmdMoverHechi 
      Height          =   330
      Index           =   0
      Left            =   11475
      MouseIcon       =   "frmMain.frx":4957A
      MousePointer    =   99  'Custom
      Top             =   5160
      Width           =   240
   End
   Begin VB.Image cmdMoverHechi 
      Height          =   270
      Index           =   1
      Left            =   11475
      MouseIcon       =   "frmMain.frx":496CC
      MousePointer    =   99  'Custom
      Top             =   4845
      Width           =   225
   End
   Begin VB.Image CmdLanzar 
      Height          =   435
      Left            =   8610
      MouseIcon       =   "frmMain.frx":4981E
      MousePointer    =   99  'Custom
      Top             =   5010
      Width           =   1635
   End
   Begin VB.Image cmdInfo 
      Height          =   435
      Left            =   10440
      MouseIcon       =   "frmMain.frx":49970
      MousePointer    =   99  'Custom
      Top             =   5010
      Width           =   900
   End
   Begin VB.Image imgSpells 
      Height          =   810
      Left            =   10320
      Top             =   1680
      Width           =   780
   End
   Begin VB.Image imgInventory 
      Height          =   810
      Left            =   9180
      Top             =   1680
      Width           =   780
   End
   Begin VB.Label lblMapPosY 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "50"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Left            =   11520
      TabIndex        =   28
      Top             =   8640
      Width           =   300
      WordWrap        =   -1  'True
   End
   Begin VB.Label imgAsignarSkill 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "999"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   11460
      TabIndex        =   27
      Top             =   1200
      Width           =   300
   End
   Begin VB.Image imgRanking 
      Height          =   330
      Left            =   8520
      Top             =   8130
      Width           =   1440
   End
   Begin VB.Image expBar 
      Height          =   150
      Left            =   8880
      Top             =   750
      Width           =   2400
   End
   Begin VB.Label lblExp 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "999999999/99999999"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Left            =   8880
      TabIndex        =   24
      Top             =   720
      Width           =   2400
   End
   Begin VB.Label lblDropGold 
      BackStyle       =   0  'Transparent
      Height          =   345
      Left            =   8520
      MousePointer    =   99  'Custom
      TabIndex        =   23
      Top             =   6960
      Width           =   285
   End
   Begin VB.Label GldLbl 
      BackStyle       =   0  'Transparent
      Caption         =   "1000000000000"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Left            =   8880
      TabIndex        =   22
      Top             =   7005
      Width           =   1245
   End
   Begin VB.Label lblMana 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C000&
      BackStyle       =   0  'Transparent
      Caption         =   "999/999"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   180
      Left            =   8580
      TabIndex        =   4
      ToolTipText     =   "Maná"
      Top             =   6555
      Width           =   2025
   End
   Begin VB.Label lblEnergia 
      Alignment       =   2  'Center
      BackColor       =   &H0000C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "999/999"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   180
      Left            =   8580
      TabIndex        =   3
      ToolTipText     =   "Energía"
      Top             =   6015
      Width           =   2025
   End
   Begin VB.Label lblVida 
      Alignment       =   2  'Center
      BackColor       =   &H000000FF&
      BackStyle       =   0  'Transparent
      Caption         =   "999/999"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   180
      Left            =   8580
      TabIndex        =   5
      ToolTipText     =   "Salud"
      Top             =   6285
      Width           =   2025
   End
   Begin VB.Label lblMapName 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "-"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   5.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   240
      Left            =   10200
      TabIndex        =   20
      Top             =   8520
      Width           =   1605
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblMapPosX 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "50"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Left            =   10080
      TabIndex        =   19
      Top             =   8640
      Width           =   300
      WordWrap        =   -1  'True
   End
   Begin VB.Image hpBar 
      Height          =   225
      Left            =   8565
      Top             =   6255
      Width           =   2040
   End
   Begin VB.Image manBar 
      Height          =   225
      Left            =   8565
      Top             =   6525
      Width           =   2040
   End
   Begin VB.Image staBar 
      Height          =   225
      Left            =   8565
      Top             =   5985
      Width           =   2040
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "44"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   300
      Left            =   11040
      TabIndex        =   18
      Top             =   1380
      Width           =   480
   End
   Begin VB.Image imgOpciones 
      Height          =   330
      Left            =   8520
      Top             =   8490
      Width           =   1440
   End
   Begin VB.Image imgPremium 
      Height          =   240
      Left            =   120
      Picture         =   "frmMain.frx":49AC2
      Top             =   150
      Visible         =   0   'False
      Width           =   960
   End
   Begin VB.Image Image3 
      Height          =   300
      Left            =   8160
      Top             =   1800
      Width           =   315
   End
   Begin VB.Image imgEstadisticas 
      Height          =   330
      Left            =   8520
      Top             =   7410
      Width           =   1440
   End
   Begin VB.Image imgClanes 
      Height          =   330
      Left            =   8520
      Top             =   7770
      Width           =   1440
   End
   Begin VB.Label lblCerrar 
      BackStyle       =   0  'Transparent
      Height          =   315
      Left            =   11610
      MousePointer    =   99  'Custom
      TabIndex        =   13
      Top             =   150
      Width           =   315
   End
   Begin VB.Label lblName 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Standelf"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   270
      Left            =   8865
      TabIndex        =   8
      Top             =   345
      Width           =   2445
   End
   Begin VB.Label lblHambre 
      BackColor       =   &H00004000&
      BackStyle       =   0  'Transparent
      Caption         =   "100%"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   180
      Left            =   10800
      TabIndex        =   6
      ToolTipText     =   "Hambre"
      Top             =   5760
      Width           =   525
   End
   Begin VB.Label lblSed 
      BackColor       =   &H00400000&
      BackStyle       =   0  'Transparent
      Caption         =   "100%"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   180
      Left            =   11280
      TabIndex        =   7
      ToolTipText     =   "Sed"
      Top             =   5760
      Width           =   555
   End
   Begin VB.Menu mnuconsol 
      Caption         =   "Consola"
      Visible         =   0   'False
      Begin VB.Menu mnu_fake 
         Caption         =   "General:"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnu_clanC 
         Caption         =   "Consola de Clan"
      End
      Begin VB.Menu mnu_NormalC 
         Caption         =   "Consola General"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnu_CleanConsole 
         Caption         =   "Limpiar Consola"
      End
      Begin VB.Menu mnu_fake2 
         Caption         =   "Modos de Habla:"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnu_talked 
         Caption         =   "Normal"
         Checked         =   -1  'True
         Index           =   0
      End
      Begin VB.Menu mnu_talked 
         Caption         =   "Gritar"
         Index           =   1
      End
      Begin VB.Menu mnu_talked 
         Caption         =   "Susurrar"
         Index           =   2
      End
      Begin VB.Menu mnu_talked 
         Caption         =   "Global"
         Index           =   3
      End
      Begin VB.Menu mnu_talked 
         Caption         =   "Facción"
         Index           =   4
      End
      Begin VB.Menu mnu_talked 
         Caption         =   "Clan"
         Index           =   5
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' ProyectoAO Documentación ***********************************
' Autor: Standelf
' Descripción: Formulario Principal
' Última modificación: 07/11/2012
' Modificación: Reordenamiento y Limpieza de Líneas innecesarias
'*************************************************************

Option Explicit

Private TalkMode As Byte
Private tmpNamePrivate As String
Private last_i As Long
Public UsandoDrag As Boolean
Public UsabaDrag As Boolean
Public tX As Byte
Public tY As Byte
Public MouseX As Long
Public MouseY As Long
Public MouseBoton As Long
Public MouseShift As Long
Private clicX As Long
Private clicY As Long
Public IsPlaying As Byte
Private clsFormulario As clsFormMovementManager
Private tmpLastCommand As String ' Save the last Text


Public tester As Boolean




Private Sub Check1_Click()
tester = Not tester

End Sub

Private Sub Command1_Click()
    DX8_GreyScale = Not DX8_GreyScale
    SurfaceDB.Reset_Surfaces
End Sub

'   Call WriteRequestPartyForm
'   Call WriteQuestListRequest

'**************************************************************************************************
' ProyectoAO Menú de Hechizos ******************************************************************
'**************************************************************************************************

Private Sub imgSpells_Click()
Call Audio.PlayWave(SND_CLICK)
    PicInv.Visible = False
    hlst.Visible = True
End Sub


Private Sub cmdINFO_Click()
    Call Audio.PlayWave(SND_CLICK)
    If hlst.ListIndex <> -1 Then
        Call WriteSpellInfo(hlst.ListIndex + 1)
    End If
End Sub

Private Sub cmdMoverHechi_Click(index As Integer)
    Call Audio.PlayWave(SND_CLICK)
    
    If hlst.Visible = True Then
        If hlst.ListIndex = -1 Then Exit Sub
        Dim sTemp As String
    
        Select Case index
            Case 1
                If hlst.ListIndex = 0 Then Exit Sub
            Case 0
                If hlst.ListIndex = hlst.ListCount - 1 Then Exit Sub
        End Select
    
        Call WriteMoveSpell(index = 1, hlst.ListIndex + 1)
        
        Select Case index
            Case 1 'subir
                sTemp = hlst.List(hlst.ListIndex - 1)
                hlst.List(hlst.ListIndex - 1) = hlst.List(hlst.ListIndex)
                hlst.List(hlst.ListIndex) = sTemp
                hlst.ListIndex = hlst.ListIndex - 1
            Case 0 'bajar
                sTemp = hlst.List(hlst.ListIndex + 1)
                hlst.List(hlst.ListIndex + 1) = hlst.List(hlst.ListIndex)
                hlst.List(hlst.ListIndex) = sTemp
                hlst.ListIndex = hlst.ListIndex + 1
        End Select
    End If
End Sub

Private Sub hlst_KeyDown(KeyCode As Integer, Shift As Integer)
       KeyCode = 0
End Sub

Private Sub hlst_KeyPress(KeyAscii As Integer)
       KeyAscii = 0
End Sub

Private Sub hlst_KeyUp(KeyCode As Integer, Shift As Integer)
        KeyCode = 0
End Sub


'**************************************************************************************************
' ProyectoAO Menú del Inventario ******************************************************************
'**************************************************************************************************

Private Sub imgInventory_Click()
Call Audio.PlayWave(SND_CLICK)
    PicInv.Visible = True
    hlst.Visible = False
End Sub

Private Sub Label1_Click()
    WriteQuestListRequest
End Sub

Private Sub picInv_DblClick()
    If frmCarp.Visible Or frmHerrero.Visible Then Exit Sub
    If Not MainTimer.Check(TimersIndex.UseItemWithDblClick) Then Exit Sub
    If MacroTrabajo.Enabled Then Call DesactivarMacroTrabajo
    Call UsarItem
    UsandoDrag = False
End Sub

Private Sub picInv_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call Audio.PlayWave(SND_CLICK)
End Sub

Private Sub picInv_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Not UsandoDrag And Not UsabaDrag Then
        PicInv.MousePointer = vbDefault
    End If
End Sub

Private Sub picInv_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbRightButton Then
        If Settings.DragerDrop = False Then Exit Sub
        
        #If SeguridadBlisse Then
            If CanUse_Interval(dInter.DragDrop) Then
        #End If
        
        Dim i As Integer
        
        If Inventario.GrhIndex(Inventario.SelectedItem) > 0 Then
            last_i = Inventario.SelectedItem
                If last_i > 0 And last_i <= MAX_INVENTORY_SLOTS Then
                    Dim poss As Integer
                    poss = Index_Seek(Inventario.GrhIndex(Inventario.SelectedItem))
                        If poss = 0 Then
                            i = GrhData(Inventario.GrhIndex(Inventario.SelectedItem)).FileNum
            
                             Dim File As String
                             File = Resources.Graphics & GrhData(Inventario.GrhIndex(Inventario.SelectedItem)).FileNum & ".bmp"
                             
                             frmMain.ImageList1.ListImages.Add , CStr("g" & Inventario.GrhIndex(Inventario.SelectedItem)), Picture:=LoadPicture(File)
                             poss = frmMain.ImageList1.ListImages.Count
                        End If
                    
                    UsandoDrag = True
                        
                    Set PicInv.MouseIcon = frmMain.ImageList1.ListImages(poss).ExtractIcon
                   frmMain.PicInv.MousePointer = vbCustom
        
                    Exit Sub
                End If
            End If
        #If SeguridadBlisse Then
        End If
        #End If
    End If
End Sub

'**************************************************************************************************
' ProyectoAO Botones en general *******************************************************************
'**************************************************************************************************

Private Sub imgAsignarSkill_Click()
    Dim i As Integer
    
    LlegaronSkills = False
    Call WriteRequestSkills
    Call FlushBuffer
    
    Do While Not LlegaronSkills
        DoEvents
    Loop
    LlegaronSkills = False
    
    For i = 1 To NUMSKILLS
        frmSkills3.Text1(i).Caption = UserSkills(i)
    Next i
    
    Alocados = SkillPoints
    frmSkills3.puntos.Caption = SkillPoints
    frmSkills3.Show , frmMain

End Sub

Private Sub imgClanes_Click()
    Call Audio.PlayWave(SND_CLICK)
    If frmGuildLeader.Visible Then Unload frmGuildLeader
    Call WriteRequestGuildLeaderInfo
End Sub

Private Sub imgEstadisticas_Click()
    Call Audio.PlayWave(SND_CLICK)
    LlegaronAtrib = False
    LlegaronSkills = False
    LlegoFama = False
    Call WriteRequestAtributes
    Call WriteRequestSkills
    Call WriteRequestMiniStats
    Call WriteRequestFame
    Call FlushBuffer
    Do While Not LlegaronSkills Or Not LlegaronAtrib Or Not LlegoFama
        DoEvents
    Loop
    frmEstadisticas.Iniciar_Labels
    frmEstadisticas.Show , frmMain
    LlegaronAtrib = False
    LlegaronSkills = False
    LlegoFama = False
    
End Sub

Private Sub imgOpciones_Click()
    Call Audio.PlayWave(SND_CLICK)
    Call frmOpciones.Show(vbModeless, frmMain)
End Sub

Private Sub imgRanking_Click()
    Call Audio.PlayWave(SND_CLICK)
End Sub























Private Sub Form_Load()
    If Not Settings.Ventana Then
        '   Handles Form movement (drag and drop).
        Set clsFormulario = New clsFormMovementManager
        clsFormulario.Initialize Me, , 120
    End If

    'Load static GUI
    'frmMain.Picture = General_Set_GUI("GUI_main")
    expBar.Picture = General_Set_GUI("GUI_exp")
    staBar.Picture = General_Set_GUI("GUI_est")
    hpBar.Picture = General_Set_GUI("GUI_life")
    manBar.Picture = General_Set_GUI("GUI_man")
    invInfo.Picture = General_Set_GUI("GUI_infoinv")
    
    'imgOpciones.Picture = General_Set_GUI("GUI_op0")
    'imgEstadisticas.Picture = General_Set_GUI("GUI_est0")
    'imgClanes.Picture = General_Set_GUI("GUI_com0")
    'imgRanking.Picture = General_Set_GUI("GUI_rank0")
    
    'imgSpells.Picture = General_Set_GUI("GUI_spell0")
    'imgInventory.Picture = General_Set_GUI("GUI_inv1")
    
    'CmdLanzar.Picture = General_Set_GUI("GUI_shut0")
    'cmdINFO.Picture = General_Set_GUI("GUI_inf0")
    
    
    'Round_Picture MiniMap, 100

    Call SetWindowLong(RecTxt.hWnd, GWL_EXSTYLE, WS_EX_TRANSPARENT)
    Call SetWindowLong(GuildTxt.hWnd, GWL_EXSTYLE, WS_EX_TRANSPARENT)
    
    Me.Height = 9000
    Me.Width = 12000
End Sub





Public Sub ControlSM(ByVal index As Byte, ByVal Mostrar As Boolean)
                 
Select Case index
    Case eSMType.sResucitation
        If Mostrar Then
            Call General_Add_to_RichTextBox(frmMain.RecTxt, MENSAJE_SEGURO_RESU_ON, 0, 255, 0, True, False, True)
            picSM(index).ToolTipText = "Seguro de resucitación activado."
            picSM(index).Picture = General_Set_GUI("GUI_segr0")
        Else
            Call General_Add_to_RichTextBox(frmMain.RecTxt, MENSAJE_SEGURO_RESU_OFF, 255, 0, 0, True, False, True)
            picSM(index).ToolTipText = "Seguro de resucitación desactivado."
            picSM(index).Picture = General_Set_GUI("GUI_segr1")
        End If
        
    Case eSMType.sSafemode
        If Mostrar Then
            Call General_Add_to_RichTextBox(frmMain.RecTxt, MENSAJE_SEGURO_ACTIVADO, 0, 255, 0, True, False, True)
            picSM(index).ToolTipText = "Seguro activado."
            picSM(index).Picture = General_Set_GUI("GUI_segc0")
        Else
            Call General_Add_to_RichTextBox(frmMain.RecTxt, MENSAJE_SEGURO_DESACTIVADO, 255, 0, 0, True, False, True)
            picSM(index).Picture = General_Set_GUI("GUI_segc1")
        End If
        
    Case eSMType.mSpells
        If Mostrar Then
            picSM(index).ToolTipText = "Macro de hechizos activado."
            picSM(index).Picture = General_Set_GUI("GUI_mact00")
        Else
            picSM(index).ToolTipText = "Macro de hechizos desactivado."
            picSM(index).Picture = General_Set_GUI("GUI_mact00")
        End If
        
    Case eSMType.mWork
        If Mostrar Then
            picSM(index).ToolTipText = "Macro de trabajo activado."
            picSM(index).Picture = General_Set_GUI("GUI_mact0")
        Else
            picSM(index).ToolTipText = "Macro de trabajo desactivado."
            picSM(index).Picture = General_Set_GUI("GUI_mact1")
        End If
        
    Case eSMType.sItem
        If Mostrar Then
            Call General_Add_to_RichTextBox(frmMain.RecTxt, MENSAJE_SEGUROi_ACTIVADO, 0, 255, 0, True, False, True)
            picSM(index).ToolTipText = "Seguro de Items activado."
            picSM(index).Picture = General_Set_GUI("GUI_segi0")
        Else
            Call General_Add_to_RichTextBox(frmMain.RecTxt, MENSAJE_SEGUROi_DESACTIVADO, 255, 0, 0, True, False, True)
            picSM(index).ToolTipText = "Seguro de Items desactivado."
            picSM(index).Picture = General_Set_GUI("GUI_segi1")
        End If
End Select

End Sub






Private Sub MiniMap_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbRightButton And General_Get_GM(UserCharIndex) And Not UserMeditar Then
        Call WriteWarpChar("YO", UserMap, CByte(X), CByte(Y))
    End If
End Sub


Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
'***************************************************
'Autor: Unknown
'Last Modification: 18/11/2009
'18/11/2009: ZaMa - Ahora se pueden poner comandos en los mensajes personalizados (execpto guildchat y privados)
'***************************************************

    If (Not SendTxt.Visible) And (Not SendCMSTXT.Visible) Then
        
        'Checks if the key is valid
        If LenB(CustomKeys.ReadableName(KeyCode)) > 0 Then
            Select Case KeyCode
                Case CustomKeys.BindedKey(eKeyType.mKeyToggleMusic)
                    Audio.MusicActivated = Not Audio.MusicActivated
                    
                Case CustomKeys.BindedKey(eKeyType.mKeyToggleSound)
                    Audio.SoundActivated = Not Audio.SoundActivated
                    
                Case CustomKeys.BindedKey(eKeyType.mKeyToggleFxs)
                    Audio.SoundEffectsActivated = Not Audio.SoundEffectsActivated
                
                Case CustomKeys.BindedKey(eKeyType.mKeyGetObject)
                    Call AgarrarItem
                
                Case CustomKeys.BindedKey(eKeyType.mKeyEquipObject)
                    Call EquiparItem
                
                Case CustomKeys.BindedKey(eKeyType.mKeyToggleNames)
                    Nombres = Nombres + 1
                    If Nombres = 4 Then Nombres = 1
                
                Case CustomKeys.BindedKey(eKeyType.mKeyTamAnimal)
                    If UserEstado = 1 Then
                        With FontTypes(FontTypeNames.FONTTYPE_INFO)
                            Call ShowConsoleMsg("¡¡Estás muerto!!", .Red, .Green, .Blue, .bold, .italic)
                        End With
                    Else
                        Call WriteWork(eSkill.Domar)
                    End If
                    
                Case CustomKeys.BindedKey(eKeyType.mKeySteal)
                    If UserEstado = 1 Then
                        With FontTypes(FontTypeNames.FONTTYPE_INFO)
                            Call ShowConsoleMsg("¡¡Estás muerto!!", .Red, .Green, .Blue, .bold, .italic)
                        End With
                    Else
                        Call WriteWork(eSkill.Robar)
                    End If
                    
                Case CustomKeys.BindedKey(eKeyType.mKeyHide)
                    If UserEstado = 1 Then
                        With FontTypes(FontTypeNames.FONTTYPE_INFO)
                            Call ShowConsoleMsg("¡¡Estás muerto!!", .Red, .Green, .Blue, .bold, .italic)
                        End With
                    Else
                        Call WriteWork(eSkill.Ocultarse)
                    End If
                                    
                Case CustomKeys.BindedKey(eKeyType.mKeyDropObject)
                   #If SeguridadBlisse = 1 Then
                        If CanUse_Interval(dInter.Drop) Then Call TirarItem
                    #Else
                        Call TirarItem
                    #End If

                
                Case CustomKeys.BindedKey(eKeyType.mKeyUseObject)
                    If MacroTrabajo.Enabled Then Call DesactivarMacroTrabajo
                        
                    If MainTimer.Check(TimersIndex.UseItemWithU) Then
                        Call UsarItem
                    End If
                
                Case CustomKeys.BindedKey(eKeyType.mKeyRequestRefresh)
                    If MainTimer.Check(TimersIndex.SendRPU) Then
                        Call WriteRequestPositionUpdate
                        Beep
                    End If
                Case CustomKeys.BindedKey(eKeyType.mKeyToggleSafeMode)
                    Call WriteSafeToggle

                Case CustomKeys.BindedKey(eKeyType.mKeyToggleResuscitationSafe)
                    Call WriteResuscitationToggle
            End Select
        End If
    End If
    
    Select Case KeyCode
        Case CustomKeys.BindedKey(eKeyType.mKeyTalkWithGuild)
            If SendTxt.Visible Then Exit Sub
            
            If (Not Comerciando) And (Not MirandoAsignarSkills) And _
              (Not frmMSG.Visible) And (Not MirandoForo) And _
              (Not frmEstadisticas.Visible) And (Not frmCantidad.Visible) Then
                SendCMSTXT.Visible = True
                SendCMSTXT.SetFocus
            End If
        
        Case CustomKeys.BindedKey(eKeyType.mKeyTakeScreenShot)
            Call ShowConsoleMsg("ScreenShoot no tomada", 255, 0, 0, False, False)
                
        Case CustomKeys.BindedKey(eKeyType.mKeyShowOptions)
            Call frmOpciones.Show(vbModeless, frmMain)
        
        Case CustomKeys.BindedKey(eKeyType.mKeyMeditate)
            If UserMinMAN = UserMaxMAN Then Exit Sub
            
            If UserEstado = 1 Then
                With FontTypes(FontTypeNames.FONTTYPE_INFO)
                    Call ShowConsoleMsg("¡¡Estás muerto!!", .Red, .Green, .Blue, .bold, .italic)
                End With
                Exit Sub
            End If
            
            #If SeguridadBlisse = 1 Then
                If CanUse_Interval(dInter.Drop) Then Call WriteMeditate
            #Else
                Call WriteMeditate
            #End If
            
        
        Case CustomKeys.BindedKey(eKeyType.mKeyCastSpellMacro)
            If UserEstado = 1 Then
                With FontTypes(FontTypeNames.FONTTYPE_INFO)
                    Call ShowConsoleMsg("¡¡Estás muerto!!", .Red, .Green, .Blue, .bold, .italic)
                End With
                Exit Sub
            End If
            
            If TrainingMacro.Enabled Then
                DesactivarMacroHechizos
            Else
                ActivarMacroHechizos
            End If
        
        Case CustomKeys.BindedKey(eKeyType.mKeyWorkMacro)
            If UserEstado = 1 Then
                With FontTypes(FontTypeNames.FONTTYPE_INFO)
                    Call ShowConsoleMsg("¡¡Estás muerto!!", .Red, .Green, .Blue, .bold, .italic)
                End With
                Exit Sub
            End If
            
            If MacroTrabajo.Enabled Then
                Call DesactivarMacroTrabajo
            Else
                Call ActivarMacroTrabajo
            End If
        
        Case CustomKeys.BindedKey(eKeyType.mKeyExitGame)
            If frmMain.MacroTrabajo.Enabled Then Call DesactivarMacroTrabajo
            Call WriteQuit
            
        Case CustomKeys.BindedKey(eKeyType.mKeyAttack)
            If Shift <> 0 Then Exit Sub
            
            If Not MainTimer.Check(TimersIndex.Arrows, False) Then Exit Sub 'Check if arrows interval has finished.
            If Not MainTimer.Check(TimersIndex.CastSpell, False) Then 'Check if spells interval has finished.
                If Not MainTimer.Check(TimersIndex.CastAttack) Then Exit Sub 'Corto intervalo Golpe-Hechizo
            Else
                If Not MainTimer.Check(TimersIndex.Attack) Or UserDescansar Or UserMeditar Then Exit Sub
            End If
            
            If TrainingMacro.Enabled Then Call DesactivarMacroHechizos
            If MacroTrabajo.Enabled Then Call DesactivarMacroTrabajo
            Call WriteAttack
        
        Case CustomKeys.BindedKey(eKeyType.mKeyTalk)
            If SendCMSTXT.Visible Then Exit Sub
            
            If (Not Comerciando) And (Not MirandoAsignarSkills) And _
              (Not frmMSG.Visible) And (Not MirandoForo) And _
              (Not frmEstadisticas.Visible) And (Not frmCantidad.Visible) Then
                SendTxt.Visible = True

                SendTxt.SetFocus
            End If
            
    End Select
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    MouseBoton = Button
    MouseShift = Shift
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    clicX = X
    clicY = Y
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If prgRun = True Then
        prgRun = False
        Cancel = 1
    End If
End Sub

Private Sub Image3_Click()
    frmMain.PopupMenu frmMain.mnuconsol
End Sub

Private Sub lblCerrar_Click()
Call Audio.PlayWave(SND_CLICK)
If UserParalizado Then 'Inmo
    With FontTypes(FontTypeNames.FONTTYPE_WARNING)
        Call ShowConsoleMsg("No puedes salir estando paralizado.", .Red, .Green, .Blue, .bold, .italic)
    End With
    Exit Sub
End If
If frmMain.MacroTrabajo.Enabled Then Call frmMain.DesactivarMacroTrabajo
Call WriteQuit
End Sub

Private Sub lblExp_Click()
MostrarExp = Not MostrarExp
If MostrarExp = True Then
    frmMain.lblExp.Caption = UserExp & "/" & UserPasarNivel
Else
    If UserPasarNivel > 0 Then
        frmMain.lblExp.Caption = "[" & Round(CDbl(UserExp) * CDbl(100) / CDbl(UserPasarNivel), 2) & "%]"
    Else
        frmMain.lblExp.Caption = "[N/A]"
    End If
End If
End Sub

Private Sub macrotrabajo_Timer()
    If Inventario.SelectedItem = 0 Then
        Call DesactivarMacroTrabajo
        Exit Sub
    End If
    
    'Macros are disabled if not using Argentum!
    If Not General_Is_App_Active() Then
        Call DesactivarMacroTrabajo
        Exit Sub
    End If
    
    If UsingSkill = eSkill.Pesca Or UsingSkill = eSkill.Talar Or UsingSkill = eSkill.Mineria Or _
                UsingSkill = FundirMetal Or (UsingSkill = eSkill.Herreria And Not frmHerrero.Visible) Then
        Call WriteWorkLeftClick(tX, tY, UsingSkill)
        UsingSkill = 0
    End If
    
    'If Inventario.OBJType(Inventario.SelectedItem) = eObjType.otWeapon Then
     If Not (frmCarp.Visible = True) Then Call UsarItem
End Sub

Public Sub ActivarMacroTrabajo()
    MacroTrabajo.Interval = INT_MACRO_TRABAJO
    MacroTrabajo.Enabled = True
    Call General_Add_to_RichTextBox(frmMain.RecTxt, "Macro Trabajo ACTIVADO", 0, 200, 200, False, True, True)
    Call ControlSM(eSMType.mWork, True)
End Sub

Public Sub DesactivarMacroTrabajo()
    MacroTrabajo.Enabled = False
    MacroBltIndex = 0
    UsingSkill = 0
    MousePointer = vbDefault
    Call General_Add_to_RichTextBox(frmMain.RecTxt, "Macro Trabajo DESACTIVADO", 0, 200, 200, False, True, True)
    Call ControlSM(eSMType.mWork, False)
End Sub


Private Sub MainViewPic_Click()
    If Cartel Then Cartel = False
        
    If Not Comerciando Then
        Call ConvertCPtoTP(MouseX, MouseY, tX, tY)
        
        If Setting_Map_Areas = True Then
            If MouseBoton = vbLeftButton Then
                Ambient_Set_Area tX, tY, frmAmbientEditor.List2.ListIndex + 1, frmAmbientEditor.Area_Range
            Else
                Ambient_Set_Area tX, tY, 0, frmAmbientEditor.Area_Range
            End If
        End If
        
        If MouseShift = 0 Then
            If MouseBoton <> vbRightButton Then
                '[ybarra]
                If UsaMacro Then
                    CnTd = CnTd + 1
                    If CnTd = 3 Then
                        Call WriteUseSpellMacro
                        CnTd = 0
                    End If
                    UsaMacro = False
                End If
                '[/ybarra]
                If UsingSkill = 0 Then
                    Call WriteLeftClick(tX, tY)
                Else
                
                    If TrainingMacro.Enabled Then Call DesactivarMacroHechizos
                    If MacroTrabajo.Enabled Then Call DesactivarMacroTrabajo
                    
                    If Not MainTimer.Check(TimersIndex.Arrows, False) Then 'Check if arrows interval has finished.
                        frmMain.MousePointer = vbDefault
                        UsingSkill = 0
                        With FontTypes(FontTypeNames.FONTTYPE_TALK)
                            Call General_Add_to_RichTextBox(frmMain.RecTxt, "No puedes lanzar proyectiles tan rápido.", .Red, .Green, .Blue, .bold, .italic)
                        End With
                        Exit Sub
                    End If
                    
                    'Splitted because VB isn't lazy!
                    If UsingSkill = Proyectiles Then
                        If Not MainTimer.Check(TimersIndex.Arrows) Then
                            frmMain.MousePointer = vbDefault
                            UsingSkill = 0
                            With FontTypes(FontTypeNames.FONTTYPE_TALK)
                                Call General_Add_to_RichTextBox(frmMain.RecTxt, "No puedes lanzar proyectiles tan rápido.", .Red, .Green, .Blue, .bold, .italic)
                            End With
                            Exit Sub
                        End If
                    End If
                    
                    'Splitted because VB isn't lazy!
                    If UsingSkill = Magia Then
                        If Not MainTimer.Check(TimersIndex.Attack, False) Then 'Check if attack interval has finished.
                            If Not MainTimer.Check(TimersIndex.CastAttack) Then 'Corto intervalo de Golpe-Magia
                               frmMain.MousePointer = vbDefault
                                UsingSkill = 0
                                With FontTypes(FontTypeNames.FONTTYPE_TALK)
                                    Call General_Add_to_RichTextBox(frmMain.RecTxt, "No puedes lanzar hechizos tan rápido.", .Red, .Green, .Blue, .bold, .italic)
                                End With
                                Exit Sub
                            End If
                        Else
                            If Not MainTimer.Check(TimersIndex.CastSpell) Then 'Check if spells interval has finished.
                                frmMain.MousePointer = vbDefault
                                UsingSkill = 0
                                With FontTypes(FontTypeNames.FONTTYPE_TALK)
                                    Call General_Add_to_RichTextBox(frmMain.RecTxt, "No puedes lanzar hechizos tan rapido.", .Red, .Green, .Blue, .bold, .italic)
                                End With
                                Exit Sub
                            End If
                        End If
                    End If
                    
                    'Splitted because VB isn't lazy!
                    If (UsingSkill = Pesca Or UsingSkill = Robar Or UsingSkill = Talar Or UsingSkill = Mineria Or UsingSkill = FundirMetal) Then
                        If Not MainTimer.Check(TimersIndex.Work) Then
                            frmMain.MousePointer = vbDefault
                            UsingSkill = 0
                            Exit Sub
                        End If
                    End If
                    
                    If frmMain.MousePointer <> 2 Then Exit Sub 'Parcheo porque a veces tira el hechizo sin tener el cursor (NicoNZ)
                    
                    frmMain.MousePointer = vbDefault
                    Call WriteWorkLeftClick(tX, tY, UsingSkill)
                    UsingSkill = 0
                End If
            Else
                'Call AbrirMenuViewPort
            End If
        ElseIf (MouseShift And 1) = 1 Then
            If Not CustomKeys.KeyAssigned(KeyCodeConstants.vbKeyShift) Then
                If MouseBoton = vbLeftButton Then
                    Call WriteWarpChar("YO", UserMap, tX, tY)
                End If
            End If
        End If
    End If

    SpellGrhIndex = 0
End Sub

Private Sub MainViewPic_DblClick()
'**************************************************************
'Author: Unknown
'Last Modify Date: 12/27/2007
'12/28/2007: ByVal - Chequea que la ventana de comercio y boveda no este abierta al hacer doble clic a un comerciante, sobrecarga la lista de items.
'**************************************************************
    If Not MirandoForo And Not Comerciando Then 'frmComerciar.Visible And Not frmBancoObj.Visible Then
        Call WriteDoubleClick(tX, tY)
    End If
End Sub

Private Sub MainViewPic_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    MouseBoton = Button
    MouseShift = Shift
    
    If tX <> 0 And tY <> 0 Then
        If MapData(tX, tY).CharIndex <> 0 And Button = vbRightButton Then
            Call WriteGiveMePower(CInt(tX), CInt(tY))
        End If
    End If
    
End Sub

Private Sub MainViewPic_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    MouseX = X
    MouseY = Y

    If UsabaDrag Then
        Call ConvertCPtoTP(MouseX, MouseY, tX, tY)
        Call General_Drop_X_Y(tX, tY)
        UsabaDrag = False
    End If
    
    
End Sub

Private Sub MainViewPic_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    clicX = X
    clicY = Y
End Sub

Private Sub mnu_changeuser_Click()
Call Audio.PlayWave(SND_CLICK)
    
EstadoLogin = E_MODO.LoginCuenta

If frmMain.Winsock1.State <> sckClosed Then
    frmMain.Winsock1.Close
    DoEvents
End If
    
If General_Check_AccountData(False, False) = True Then
    frmMain.Winsock1.Connect Server_IP, Server_Port
End If
       
Me.Hide
End Sub

Private Sub mnu_clanC_Click()
    mnu_NormalC.Checked = False
    mnu_clanC.Checked = True
    
    RecTxt.Visible = False
    GuildTxt.Visible = True
End Sub

Private Sub mnu_CleanConsole_Click()
    RecTxt.Text = vbNullString
    GuildTxt.Text = vbNullString
End Sub

Private Sub mnu_NormalC_Click()
    mnu_NormalC.Checked = True
    mnu_clanC.Checked = False
    
    RecTxt.Visible = True
    GuildTxt.Visible = False
End Sub

Private Sub mnu_talked_Click(index As Integer)
    Dim i As Integer
        For i = 0 To 5
            mnu_talked(i).Checked = False
        Next i
    mnu_talked(index).Checked = True
    
    TalkMode = index
End Sub

Private Sub PicMH_Click()
    Call General_Add_to_RichTextBox(frmMain.RecTxt, "Auto lanzar hechizos. Utiliza esta habilidad para entrenar únicamente. Para activarlo/desactivarlo utiliza F7.", 255, 255, 255, False, False, True)
End Sub

Private Sub Coord_Click()
    Call General_Add_to_RichTextBox(frmMain.RecTxt, "Estas coordenadas son tu ubicación en el mapa. Utiliza la letra L para corregirla si esta no se corresponde con la del servidor por efecto del Lag.", 255, 255, 255, False, False, True)
End Sub





Private Sub picSM_DblClick(index As Integer)
Select Case index
    Case eSMType.sResucitation
        Call WriteResuscitationToggle
        
    Case eSMType.sSafemode
        Call WriteSafeToggle
        
    Case eSMType.mSpells
        If UserEstado = 1 Then
            With FontTypes(FontTypeNames.FONTTYPE_INFO)
                Call ShowConsoleMsg("¡¡Estás muerto!!", .Red, .Green, .Blue, .bold, .italic)
            End With
            Exit Sub
        End If
        
        If TrainingMacro.Enabled Then
            Call DesactivarMacroHechizos
        Else
            Call ActivarMacroHechizos
        End If
        
    Case eSMType.mWork
        If UserEstado = 1 Then
            With FontTypes(FontTypeNames.FONTTYPE_INFO)
                Call ShowConsoleMsg("¡¡Estás muerto!!", .Red, .Green, .Blue, .bold, .italic)
            End With
            Exit Sub
        End If
        
        If MacroTrabajo.Enabled Then
            Call DesactivarMacroTrabajo
        Else
            Call ActivarMacroTrabajo
        End If
        
    Case eSMType.sItem
        If Items_Seg = False Then
            Items_Seg = True
        Else
            Items_Seg = False
        End If
        
        Call ControlSM(eSMType.sItem, Items_Seg)
End Select
End Sub


Private Sub RCasco_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If UsabaDrag Then
        If Inventario.OBJType(Inventario.SelectedItem) = 17 Then
            Call EquiparItem
            UsabaDrag = False
        Else
            UsabaDrag = False
        End If
    End If
End Sub


Private Sub REscudo_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If UsabaDrag Then
        If Inventario.OBJType(Inventario.SelectedItem) = 16 Then
            Call EquiparItem
            UsabaDrag = False
        Else
            UsabaDrag = False
        End If
    End If
End Sub

Private Sub REspada_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If UsabaDrag Then
        If Inventario.OBJType(Inventario.SelectedItem) = 2 Then
            Call EquiparItem
            UsabaDrag = False
        Else
            UsabaDrag = False
        End If
    End If
End Sub

Private Sub Rtunica_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If UsabaDrag Then
        If Inventario.OBJType(Inventario.SelectedItem) = 3 Then
            Call EquiparItem
            UsabaDrag = False
        Else
            UsabaDrag = False
        End If
    End If
End Sub

Private Sub SendTxt_KeyUp(KeyCode As Integer, Shift As Integer)
    If SendTxt.Text = vbNullString And KeyCode = vbKeyUp Then SendTxt.Text = tmpLastCommand
    'Send text
    If KeyCode = vbKeyReturn Then
    
        'Esto no me gusta pero es lo más rápido
        If LenB(stxtbuffer) <> 0 And TalkMode <> 0 Then
            Select Case TalkMode
                Case 1 '   Gritar
                    FixText stxtbuffer, "-"
                Case 2 '   Susurrar
                    FixText stxtbuffer, "\" & tmpNamePrivate & " "
                Case 3 '   Global
                    If CanUse_Interval(dInter.GlobalChat) Then FixText stxtbuffer, "."
                Case 4 '   Facción
                    FixText stxtbuffer, ","
                Case 5 '   Clan
                    FixText stxtbuffer, "/CMSG "
            End Select
        End If
        
        If LenB(stxtbuffer) <> 0 Then
            tmpLastCommand = stxtbuffer
            Call ParseUserCommand(stxtbuffer)
        End If
        
        stxtbuffer = ""
        SendTxt.Text = ""
        KeyCode = 0
        SendTxt.Visible = False
        
        If PicInv.Visible Then
            PicInv.SetFocus
        Else
            hlst.SetFocus
        End If
    End If
End Sub



'[END]'

''''''''''''''''''''''''''''''''''''''
'       ITEM CONTROL                   '
''''''''''''''''''''''''''''''''''''''

Private Sub TirarItem()
    If Items_Seg = True Then
        With FontTypes(FontTypeNames.FONTTYPE_WARNING)
            Call ShowConsoleMsg("Para tirar algún objeto desactiva previamente el seguro de ítems.", .Red, .Green, .Blue, .bold, .italic)
        End With
        Exit Sub
    End If
    
    If UserEstado = 1 Then
        With FontTypes(FontTypeNames.FONTTYPE_INFO)
            Call ShowConsoleMsg("¡¡Estás muerto!!", .Red, .Green, .Blue, .bold, .italic)
        End With
    Else
        If (Inventario.SelectedItem > 0 And Inventario.SelectedItem < MAX_INVENTORY_SLOTS + 1) Or (Inventario.SelectedItem = FLAGORO) Then
            If Inventario.Amount(Inventario.SelectedItem) = 1 Then
                Call WriteDrop(Inventario.SelectedItem, 1, UserPos.X, Val(UserPos.Y))
            Else
                If Inventario.Amount(Inventario.SelectedItem) > 1 Then
                    If Not Comerciando Then frmCantidad.Show , frmMain
                End If
            End If
        End If
    End If
End Sub

Private Sub AgarrarItem()
    If UserEstado = 1 Then
        With FontTypes(FontTypeNames.FONTTYPE_INFO)
            Call ShowConsoleMsg("¡¡Estás muerto!!", .Red, .Green, .Blue, .bold, .italic)
        End With
    Else
        Call WritePickUp
    End If
End Sub

Private Sub UsarItem()
    If pausa Then Exit Sub
    
    If Comerciando Then Exit Sub
    
    If TrainingMacro.Enabled Then DesactivarMacroHechizos
    
    If (Inventario.SelectedItem > 0) And (Inventario.SelectedItem < MAX_INVENTORY_SLOTS + 1) Then _
        Call WriteUseItem(Inventario.SelectedItem)
End Sub

Private Sub EquiparItem()
    If UserEstado = 1 Then
        With FontTypes(FontTypeNames.FONTTYPE_INFO)
                Call ShowConsoleMsg("¡¡Estás muerto!!", .Red, .Green, .Blue, .bold, .italic)
        End With
    Else
        If Comerciando Then Exit Sub
        
        If (Inventario.SelectedItem > 0) And (Inventario.SelectedItem < MAX_INVENTORY_SLOTS + 1) Then _
        Call WriteEquipItem(Inventario.SelectedItem)
    End If
End Sub


''''''''''''''''''''''''''''''''''''''
'       HECHIZOS CONTROL               '
''''''''''''''''''''''''''''''''''''''

Private Sub TrainingMacro_Timer()
    If Not hlst.Visible Then
        DesactivarMacroHechizos
        Exit Sub
    End If
    
    'Macros are disabled if focus is not on Argentum!
    If Not General_Is_App_Active() Then
        DesactivarMacroHechizos
        Exit Sub
    End If
    
    If Comerciando Then Exit Sub
    
    If hlst.List(hlst.ListIndex) <> "(None)" And MainTimer.Check(TimersIndex.CastSpell, False) Then
        Call WriteCastSpell(hlst.ListIndex + 1)
        Call WriteWork(eSkill.Magia)
    End If
    
    Call ConvertCPtoTP(MouseX, MouseY, tX, tY)
    
    If UsingSkill = Magia And Not MainTimer.Check(TimersIndex.CastSpell) Then Exit Sub
    
    If UsingSkill = Proyectiles And Not MainTimer.Check(TimersIndex.Attack) Then Exit Sub
    
    Call WriteWorkLeftClick(tX, tY, UsingSkill)
    UsingSkill = 0
End Sub

Private Sub cmdLanzar_Click()
    Call Audio.PlayWave(SND_CLICK)
    If hlst.List(hlst.ListIndex) <> "(None)" And MainTimer.Check(TimersIndex.Work, False) Then
        If UserEstado = 1 Then
            With FontTypes(FontTypeNames.FONTTYPE_INFO)
                Call ShowConsoleMsg("¡¡Estás muerto!!", .Red, .Green, .Blue, .bold, .italic)
            End With
        Else
            Call WriteCastSpell(hlst.ListIndex + 1)
            Call WriteWork(eSkill.Magia)
            Call SetSpellCast
            UsaMacro = True
        End If
    End If
End Sub

Private Sub CmdLanzar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    UsaMacro = False
    CnTd = 0
End Sub



Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    MouseX = X - MainViewPic.Left
    MouseY = Y - MainViewPic.Top
    
    'Trim to fit screen
    If MouseX < 0 Then
        MouseX = 0
    ElseIf MouseX > MainViewPic.Width Then
        MouseX = MainViewPic.Width
    End If
    
    'Trim to fit screen
    If MouseY < 0 Then
        MouseY = 0
    ElseIf MouseY > MainViewPic.Height Then
        MouseY = MainViewPic.Height
    End If
    
End Sub


Private Sub lblDropGold_Click()
    #If SeguridadBlisse = 1 Then 'Standelf
        If CanUse_Interval(dInter.Drop) Then
    #End If
    
    Inventario.SelectGold
    If UserGLD > 0 Then
        If Not Comerciando Then frmCantidad.Show , frmMain
    End If
    
    #If SeguridadBlisse = 1 Then
        End If
    #End If
End Sub




Private Sub RecTxt_Change()
On Error Resume Next  'el .SetFocus causaba errores al salir y volver a entrar
    If Not General_Is_App_Active() Then Exit Sub
    
    If SendTxt.Visible Then
        SendTxt.SetFocus
    ElseIf Me.SendCMSTXT.Visible Then
        SendCMSTXT.SetFocus
    ElseIf (Not Comerciando) And (Not MirandoAsignarSkills) And _
        (Not frmMSG.Visible) And (Not MirandoForo) And _
        (Not frmEstadisticas.Visible) And (Not frmCantidad.Visible) Then
         
        If PicInv.Visible Then
            PicInv.SetFocus
        ElseIf hlst.Visible Then
            hlst.SetFocus
        End If
    End If
End Sub

Private Sub RecTxt_KeyDown(KeyCode As Integer, Shift As Integer)
    If PicInv.Visible Then
        PicInv.SetFocus
    Else
        hlst.SetFocus
    End If
End Sub

Private Sub SendTxt_Change()
'**************************************************************
'Author: Unknown
'Last Modify Date: 3/06/2006
'3/06/2006: Maraxus - impedí se inserten caractéres no imprimibles
'**************************************************************
    If Len(SendTxt.Text) > 160 Then
        stxtbuffer = "Soy un cheater, avisenle a un gm"
    Else
        'Make sure only valid chars are inserted (with Shift + Insert they can paste illegal chars)
        Dim i As Long
        Dim tempstr As String
        Dim CharAscii As Integer
        
        For i = 1 To Len(SendTxt.Text)
            CharAscii = Asc(mid$(SendTxt.Text, i, 1))
            If CharAscii >= vbKeySpace And CharAscii <= 250 Then
                tempstr = tempstr & Chr$(CharAscii)
            End If
        Next i
        
        If tempstr <> SendTxt.Text Then
            'We only set it if it's different, otherwise the event will be raised
            'constantly and the client will crush
            SendTxt.Text = tempstr
        End If
        
        stxtbuffer = SendTxt.Text
    End If
End Sub

Private Sub SendTxt_KeyPress(KeyAscii As Integer)
    If Not (KeyAscii = vbKeyBack) And _
       Not (KeyAscii >= vbKeySpace And KeyAscii <= 250) Then _
        KeyAscii = 0
End Sub

Private Sub SendCMSTXT_KeyUp(KeyCode As Integer, Shift As Integer)

    'Send text
    If KeyCode = vbKeyReturn Then
        'Say
        If stxtbuffercmsg <> "" Then
            Call ParseUserCommand("/CMSG " & stxtbuffercmsg)
        End If

        stxtbuffercmsg = ""
        SendCMSTXT.Text = ""
        KeyCode = 0
        Me.SendCMSTXT.Visible = False
        
        If PicInv.Visible Then
            PicInv.SetFocus
        Else
            hlst.SetFocus
        End If
    End If
End Sub

Private Sub SendCMSTXT_KeyPress(KeyAscii As Integer)
    If Not (KeyAscii = vbKeyBack) And _
       Not (KeyAscii >= vbKeySpace And KeyAscii <= 250) Then _
        KeyAscii = 0
End Sub

Private Sub SendCMSTXT_Change()
    If Len(SendCMSTXT.Text) > 160 Then
        stxtbuffercmsg = "Soy un cheater, avisenle a un GM"
    Else
        'Make sure only valid chars are inserted (with Shift + Insert they can paste illegal chars)
        Dim i As Long
        Dim tempstr As String
        Dim CharAscii As Integer
        
        For i = 1 To Len(SendCMSTXT.Text)
            CharAscii = Asc(mid$(SendCMSTXT.Text, i, 1))
            If CharAscii >= vbKeySpace And CharAscii <= 250 Then
                tempstr = tempstr & Chr$(CharAscii)
            End If
        Next i
        
        If tempstr <> SendCMSTXT.Text Then
            'We only set it if it's different, otherwise the event will be raised
            'constantly and the client will crush
            SendCMSTXT.Text = tempstr
        End If
        
        stxtbuffercmsg = SendCMSTXT.Text
    End If
End Sub

'   -------------------
'      W I N S O C K
'   -------------------
Private Sub Winsock1_Close()
    Dim i As Long
    
    Connected = False
    
    If Winsock1.State <> sckClosed Then _
        Winsock1.Close
    
    frmConnect.MousePointer = vbNormal
    
    Do While i < Forms.Count - 1
        i = i + 1
        
        If Forms(i).name <> Me.name And Forms(i).name <> frmConnect.name And Forms(i).name <> frmCrearPersonaje.name Then
            Unload Forms(i)
        End If
    Loop
    
    On Local Error GoTo 0
    If Not frmCrearPersonaje.Visible Then
        frmConnect.Visible = True
    End If
    
    frmMain.Visible = False
    Call Reset_Party
    
    pausa = False
    UserMeditar = False

    UserClase = 0
    UserSexo = 0
    UserRaza = 0
    UserHogar = 0
    
    For i = 1 To NUMSKILLS
        UserSkills(i) = 0
    Next i

    For i = 1 To NUMATRIBUTOS
        UserAtributos(i) = 0
    Next i

    SkillPoints = 0
    imgAsignarSkill.Caption = 0
    Alocados = 0
End Sub

Private Sub Winsock1_Connect()
    
    'Clean input and output buffers
    Call incomingData.ReadASCIIStringFixed(incomingData.Length)
    Call outgoingData.ReadASCIIStringFixed(outgoingData.Length)

    
    Select Case EstadoLogin
        Case E_MODO.CrearNuevoPj
            General_Write_Login

        Case E_MODO.Normal
            General_Write_Login

        Case E_MODO.Dados
            Call Audio.PlayMIDI("7.mid")
            frmCrearPersonaje.Show vbModal
        
        Case E_MODO.LoginCuenta
            General_Write_Login
        
        Case E_MODO.CrearCuenta
            General_Write_Login
            
        Case E_MODO.BorrarPJ
            General_Write_Login
    End Select
End Sub

Private Sub Winsock1_DataArrival(ByVal BytesTotal As Long)
    Dim RD As String
    Dim data() As Byte

    Winsock1.GetData RD
    
    data = StrConv(RD, vbFromUnicode)
    
    'Set data in the buffer
    Call incomingData.WriteBlock(data)
    
    'Send buffer to Handle data
    Call HandleIncomingData
End Sub

Private Sub Winsock1_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    '*********************************************
    'Handle socket errors
    '*********************************************
    
    Call MsgBox(Description, vbApplicationModal + vbInformation + vbOKOnly + vbDefaultButton1, "Error")
    frmConnect.MousePointer = 1

    If Winsock1.State <> sckClosed Then _
        Winsock1.Close

    If Not frmCrearPersonaje.Visible Then
        frmConnect.Show
        AlphaPres = 255
        set_GUI_Efect
        CurMapAmbient.Fog = 100
    Else
        frmCrearPersonaje.MousePointer = 0
    End If
End Sub





















Private Function Index_Seek(ByVal GH As Integer) As Integer
Dim i As Integer
For i = 1 To frmMain.ImageList1.ListImages.Count
    If frmMain.ImageList1.ListImages(i).Key = "g" & CStr(GH) Then
        Index_Seek = i
        Exit For
    End If
Next i
End Function

Private Function FixText(ByRef Text As String, Character As String)
'***************************************************
'Author: Standelf
'Last Modification: 24/01/10
'Talks Mods.
'***************************************************
Dim firstCharacter As String
    firstCharacter = mid(Text, 1, 1)
    If firstCharacter = "." Or firstCharacter = "/" Or firstCharacter = "-" Or firstCharacter = "\" Then Exit Function

    Text = Character & Text
End Function

Public Sub ActivarMacroHechizos()
    If Not hlst.Visible Then
        Call General_Add_to_RichTextBox(frmMain.RecTxt, "Debes tener seleccionado el hechizo para activar el auto-lanzar", 0, 200, 200, False, True, True)
        Exit Sub
    End If
    
    TrainingMacro.Interval = INT_MACRO_HECHIS
    TrainingMacro.Enabled = True
    Call General_Add_to_RichTextBox(frmMain.RecTxt, "Auto lanzar hechizos activado", 0, 200, 200, False, True, True)
    Call ControlSM(eSMType.mSpells, True)
End Sub

Public Sub DesactivarMacroHechizos()
    TrainingMacro.Enabled = False
    Call General_Add_to_RichTextBox(frmMain.RecTxt, "Auto lanzar hechizos desactivado", 0, 150, 150, False, True, True)
    Call ControlSM(eSMType.mSpells, False)
End Sub
