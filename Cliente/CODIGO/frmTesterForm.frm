VERSION 5.00
Begin VB.Form frmTesterForm 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Beta Tester Panel"
   ClientHeight    =   6495
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   7800
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
   ScaleHeight     =   6495
   ScaleWidth      =   7800
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command22 
      Caption         =   "Command22"
      Height          =   615
      Left            =   2160
      TabIndex        =   66
      Top             =   4080
      Width           =   1455
   End
   Begin VB.CommandButton Command8 
      Caption         =   "Command8"
      Height          =   615
      Left            =   1680
      TabIndex        =   65
      Top             =   5640
      Width           =   1815
   End
   Begin VB.TextBox Text99 
      Height          =   615
      Left            =   2040
      TabIndex        =   64
      Text            =   "Text5"
      Top             =   4920
      Width           =   1095
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Command5"
      Height          =   495
      Left            =   960
      TabIndex        =   63
      Top             =   4800
      Width           =   735
   End
   Begin VB.Frame Frame7 
      Caption         =   "Frame7"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3375
      Left            =   4440
      TabIndex        =   31
      Top             =   2760
      Width           =   3255
      Begin VB.TextBox Text11 
         Alignment       =   2  'Center
         BorderStyle     =   0  'None
         Height          =   285
         Index           =   3
         Left            =   1920
         TabIndex        =   62
         Text            =   "0"
         Top             =   1680
         Width           =   495
      End
      Begin VB.TextBox Text11 
         Alignment       =   2  'Center
         BorderStyle     =   0  'None
         Height          =   285
         Index           =   2
         Left            =   120
         TabIndex        =   61
         Text            =   "0"
         Top             =   1680
         Width           =   495
      End
      Begin VB.TextBox Text11 
         Alignment       =   2  'Center
         BorderStyle     =   0  'None
         Height          =   285
         Index           =   1
         Left            =   1920
         TabIndex        =   60
         Text            =   "0"
         Top             =   840
         Width           =   495
      End
      Begin VB.TextBox Text11 
         Alignment       =   2  'Center
         BorderStyle     =   0  'None
         Height          =   285
         Index           =   0
         Left            =   120
         TabIndex        =   59
         Text            =   "0"
         Top             =   840
         Width           =   495
      End
      Begin VB.TextBox Text10 
         Height          =   285
         Left            =   2880
         TabIndex        =   58
         Text            =   "5"
         Top             =   3000
         Width           =   255
      End
      Begin VB.TextBox Text9 
         Height          =   285
         Left            =   2400
         TabIndex        =   57
         Text            =   "5"
         Top             =   3000
         Width           =   255
      End
      Begin VB.TextBox Text8 
         Height          =   285
         Left            =   1680
         TabIndex        =   56
         Text            =   "100"
         Top             =   3000
         Width           =   495
      End
      Begin VB.TextBox Text7 
         Height          =   285
         Left            =   2400
         TabIndex        =   53
         Text            =   "5"
         Top             =   2400
         Width           =   495
      End
      Begin VB.TextBox Text6 
         Height          =   285
         Left            =   1680
         TabIndex        =   52
         Text            =   "100"
         Top             =   2400
         Width           =   495
      End
      Begin VB.CommandButton Command21 
         Caption         =   "Crear Montaña"
         Height          =   375
         Left            =   120
         TabIndex        =   51
         Top             =   2880
         Width           =   1455
      End
      Begin VB.CommandButton Command12 
         Caption         =   "Crear Polígono"
         Height          =   375
         Left            =   120
         TabIndex        =   50
         Top             =   2280
         Width           =   1455
      End
      Begin VB.CommandButton Command11 
         Caption         =   "Reset All Tiles"
         Height          =   735
         Left            =   2520
         TabIndex        =   49
         Top             =   720
         Width           =   615
      End
      Begin VB.CommandButton Command10 
         Caption         =   "Reset Tile"
         Height          =   495
         Left            =   2520
         TabIndex        =   48
         Top             =   240
         Width           =   615
      End
      Begin VB.Timer Timer1 
         Interval        =   500
         Left            =   120
         Top             =   960
      End
      Begin VB.CommandButton Command9 
         Caption         =   "+"
         Height          =   255
         Index           =   7
         Left            =   1800
         TabIndex        =   43
         Top             =   1800
         Width           =   255
      End
      Begin VB.CommandButton Command9 
         Caption         =   "-"
         Height          =   255
         Index           =   6
         Left            =   1560
         TabIndex        =   42
         Top             =   1800
         Width           =   255
      End
      Begin VB.CommandButton Command9 
         Caption         =   "+"
         Height          =   255
         Index           =   5
         Left            =   600
         TabIndex        =   41
         Top             =   1800
         Width           =   255
      End
      Begin VB.CommandButton Command9 
         Caption         =   "-"
         Height          =   255
         Index           =   4
         Left            =   360
         TabIndex        =   40
         Top             =   1800
         Width           =   255
      End
      Begin VB.CommandButton Command9 
         Caption         =   "+"
         Height          =   255
         Index           =   3
         Left            =   1800
         TabIndex        =   39
         Top             =   255
         Width           =   255
      End
      Begin VB.CommandButton Command9 
         Caption         =   "-"
         Height          =   255
         Index           =   2
         Left            =   1560
         TabIndex        =   38
         Top             =   255
         Width           =   255
      End
      Begin VB.CommandButton Command9 
         Caption         =   "+"
         Height          =   255
         Index           =   1
         Left            =   600
         TabIndex        =   37
         Top             =   255
         Width           =   255
      End
      Begin VB.CommandButton Command9 
         Caption         =   "-"
         Height          =   255
         Index           =   0
         Left            =   360
         TabIndex        =   36
         Top             =   255
         Width           =   255
      End
      Begin VB.Label Label6 
         Caption         =   "Altura    X      Y"
         Height          =   255
         Left            =   1680
         TabIndex        =   55
         Top             =   2760
         Width           =   1455
      End
      Begin VB.Label Label5 
         Caption         =   "Altura    Radio"
         Height          =   255
         Left            =   1680
         TabIndex        =   54
         Top             =   2160
         Width           =   1455
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000C&
         Height          =   255
         Index           =   3
         Left            =   1920
         TabIndex        =   47
         Top             =   1440
         Width           =   495
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000C&
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   46
         Top             =   1440
         Width           =   495
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000C&
         Height          =   255
         Index           =   1
         Left            =   1920
         TabIndex        =   45
         Top             =   600
         Width           =   495
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000C&
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   44
         Top             =   600
         Width           =   495
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "3"
         Height          =   255
         Index           =   3
         Left            =   1750
         TabIndex        =   35
         Top             =   1560
         Width           =   135
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "1"
         Height          =   255
         Index           =   2
         Left            =   600
         TabIndex        =   34
         Top             =   1560
         Width           =   135
      End
      Begin VB.Line Line1 
         BorderColor     =   &H000000FF&
         BorderStyle     =   3  'Dot
         X1              =   1710
         X2              =   720
         Y1              =   650
         Y2              =   1635
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "1"
         Height          =   255
         Index           =   1
         Left            =   1750
         TabIndex        =   33
         Top             =   510
         Width           =   135
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         Height          =   255
         Index           =   0
         Left            =   585
         TabIndex        =   32
         Top             =   510
         Width           =   135
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H000000FF&
         Height          =   1095
         Left            =   720
         Shape           =   1  'Square
         Top             =   600
         Width           =   1000
      End
   End
   Begin VB.Frame Frame6 
      Caption         =   "Otros"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   120
      TabIndex        =   29
      Top             =   1920
      Width           =   2295
      Begin VB.CommandButton Command6 
         Caption         =   "Bóveda Premium"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   120
         TabIndex        =   30
         Top             =   240
         Width           =   2010
      End
   End
   Begin VB.Frame Frame5 
      Caption         =   "Particle Engine"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2535
      Left            =   4440
      TabIndex        =   18
      Top             =   120
      Width           =   3255
      Begin VB.CommandButton Command20 
         Caption         =   "Effect Create"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1920
         TabIndex        =   28
         Top             =   1680
         Width           =   1095
      End
      Begin VB.CommandButton Command17 
         Caption         =   "Create Particle:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   23
         Top             =   1680
         Width           =   1215
      End
      Begin VB.CommandButton Command19 
         Caption         =   "Partícula a Usuario Index (2)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   26
         Top             =   1320
         Width           =   2895
      End
      Begin VB.CommandButton Command18 
         Caption         =   "Partícula a Usuario Index (1)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   25
         Top             =   960
         Width           =   2895
      End
      Begin VB.CommandButton Command15 
         Caption         =   "Remove All"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1560
         TabIndex        =   21
         Top             =   600
         Width           =   1455
      End
      Begin VB.CommandButton Command16 
         Caption         =   "Create Protection"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1560
         TabIndex        =   22
         Top             =   240
         Width           =   1455
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Projectile Engine"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   27
         Top             =   2160
         Width           =   1575
      End
      Begin VB.TextBox Text4 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1350
         TabIndex        =   24
         Text            =   "1"
         Top             =   1725
         Width           =   495
      End
      Begin VB.CommandButton Command14 
         Caption         =   "Create Mouse P."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   20
         Top             =   600
         Width           =   1455
      End
      Begin VB.CommandButton Command13 
         Caption         =   "Create Equation"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   19
         Top             =   240
         Width           =   1455
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Aura Engine"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   120
      TabIndex        =   10
      Top             =   2760
      Width           =   4185
      Begin VB.CommandButton Command7 
         Caption         =   "Recargar las auras"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   2760
         TabIndex        =   17
         Top             =   360
         Width           =   1335
      End
      Begin VB.TextBox Text3 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2040
         TabIndex        =   16
         Text            =   "0"
         Top             =   360
         Width           =   615
      End
      Begin VB.TextBox Text2 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   720
         TabIndex        =   15
         Text            =   "0"
         Top             =   360
         Width           =   615
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Quitar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   2040
         TabIndex        =   12
         Top             =   720
         Width           =   735
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Agregar Aura"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   11
         Top             =   720
         Width           =   1935
      End
      Begin VB.Label Label2 
         Caption         =   "Slot:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1680
         TabIndex        =   14
         Top             =   360
         Width           =   615
      End
      Begin VB.Label Label1 
         Caption         =   "Aura:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   13
         Top             =   360
         Width           =   615
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Meteo Engine"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   825
      Left            =   60
      TabIndex        =   4
      Top             =   990
      Width           =   2415
      Begin VB.CommandButton Command3 
         Caption         =   "5"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Index           =   5
         Left            =   1875
         TabIndex        =   9
         Top             =   255
         Width           =   450
      End
      Begin VB.CommandButton Command3 
         Caption         =   "4"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Index           =   4
         Left            =   1440
         TabIndex        =   8
         Top             =   255
         Width           =   450
      End
      Begin VB.CommandButton Command3 
         Caption         =   "3"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Index           =   3
         Left            =   1005
         TabIndex        =   7
         Top             =   255
         Width           =   450
      End
      Begin VB.CommandButton Command3 
         Caption         =   "2"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Index           =   2
         Left            =   570
         TabIndex        =   6
         Top             =   255
         Width           =   450
      End
      Begin VB.CommandButton Command3 
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Index           =   1
         Left            =   135
         TabIndex        =   5
         Top             =   255
         Width           =   450
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Light_Engine"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   885
      Left            =   75
      TabIndex        =   0
      Top             =   90
      Width           =   4185
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   2895
         TabIndex        =   3
         Text            =   "255 255 255 5"
         Top             =   390
         Width           =   1185
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Crear Luz (Pos Actual)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1560
         TabIndex        =   2
         Top             =   300
         Width           =   1245
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Borrar todas las luces"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   105
         TabIndex        =   1
         Top             =   300
         Width           =   1215
      End
   End
End
Attribute VB_Name = "frmTesterForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************
'Author: Ezequiel Juárez (Standelf)
'Last Modification: 14/05/10
'Blisse-AO | Tester Form, Used to Test _
    more features of the Game or engine.
'***************************************************

Private Sub Command1_Click()
    Call LightRemoveAll
End Sub






Private Sub Command10_Click()
MapData(UserPos.X, UserPos.Y).Vertex_Offset(0) = 0
MapData(UserPos.X, UserPos.Y).Vertex_Offset(1) = 0
MapData(UserPos.X, UserPos.Y).Vertex_Offset(2) = 0
MapData(UserPos.X, UserPos.Y).Vertex_Offset(3) = 0

    'Set Ambient Vertex
    CurMapAmbient.MapBlocks(UserPos.X, UserPos.Y).Vertex_Offset(0) = MapData(UserPos.X, UserPos.Y).Vertex_Offset(0)
    CurMapAmbient.MapBlocks(UserPos.X, UserPos.Y).Vertex_Offset(1) = MapData(UserPos.X, UserPos.Y).Vertex_Offset(1)
    CurMapAmbient.MapBlocks(UserPos.X, UserPos.Y).Vertex_Offset(2) = MapData(UserPos.X, UserPos.Y).Vertex_Offset(2)
    CurMapAmbient.MapBlocks(UserPos.X, UserPos.Y).Vertex_Offset(3) = MapData(UserPos.X, UserPos.Y).Vertex_Offset(3)
End Sub

Private Sub Command11_Click()
    Engine_Reset_Tile_Vertex
End Sub

Private Sub Command12_Click()
Engine_Create_Polygon CInt(UserPos.X), CInt(UserPos.Y), Text7.Text, Text6.Text, True
End Sub

Private Sub Command13_Click()
    Effect_EquationTemplate_Begin 200, 200, 1, 500, 1
End Sub

Private Sub Command15_Click()
    Effect_Kill 0, True
End Sub

Private Sub Command16_Click()
    Effect_Protection_Begin 272, 210, 1, 400, 30, 5
End Sub

Private Sub Command17_Click()
    Select Case Val(Text4.Text)
        Case 1
            Call Effect_Fire_Begin(200, 200, 1, 200, 180, 1)
        Case 2
            Call Effect_Snow_Begin(4, 200)
        Case 3
            Call Effect_Heal_Begin(200, 200, 3, 200, 1)
        Case 4
            Call Effect_Bless_Begin(200, 200, 1, 200, 30, 30)
        Case 5
            Call Effect_Protection_Begin(200, 200, 11, 200, 30, 30)
        Case 6
            Call Effect_Strengthen_Begin(200, 200, 1, 200, 30, 1)
        Case 7
            Call Effect_Rain_Begin(9, 200)
        Case 8
            Call Effect_EquationTemplate_Begin(200, 200, 1, 200, 1)
        Case 9
            Call Effect_Waterfall_Begin(200, 200, 1, 200)
        Case 10
            Call Effect_Summon_Begin(200, 200, 1, 300, 0)
    End Select
End Sub

Private Sub Command18_Click()
        Dim TempIndex As Integer
        TempIndex = Effect_Heal_Begin(1, 1, 3, 120, 1)
        Effect(TempIndex).BindToChar = UserCharIndex
        Effect(TempIndex).BindSpeed = 8
End Sub

Private Sub Command19_Click()
            Dim TempIndex As Integer
            TempIndex = Effect_Heal_Begin(700, 700, 1, 120, 1)
            Effect(TempIndex).BindToChar = UserCharIndex
            Effect(TempIndex).BindSpeed = 8
End Sub

Private Sub Command2_Click()
If Text1.Text <> "" Then
    #If LightEngine = 1 Then
        Call Create_Light_To_Map(UserPos.X, UserPos.Y, General_Get_ReadField(4, Text1.Text, Asc(" ")), General_Get_ReadField(1, Text1.Text, Asc(" ")), General_Get_ReadField(2, Text1.Text, Asc(" ")), General_Get_ReadField(3, Text1.Text, Asc(" ")))
    #Else
        Call Create_Light_To_Map(UserPos.X, UserPos.Y, General_Get_ReadField(4, Text1.Text, Asc(" ")), General_Get_ReadField(1, Text1.Text, Asc(" ")), General_Get_ReadField(2, Text1.Text, Asc(" ")), General_Get_ReadField(3, Text1.Text, Asc(" ")))
    #End If
End If
End Sub

Private Sub Command21_Click()
Engine_Create_Elevation Text8.Text, Text9.Text, Text10.Text, UserPos.X, UserPos.Y
End Sub

Private Sub Command22_Click()
Static wire As Boolean
wire = Not wire
Engine_Set_WireFrame wire
End Sub

Private Sub Command3_Click(index As Integer)
    Call Actualizar_Estado(index)
End Sub

Private Sub Command4_Click(index As Integer)
    Select Case index
        Case 0
            Set_Aura UserCharIndex, Text3.Text, Text2.Text
        Case 1
            Delete_Aura UserCharIndex, Text3.Text
    End Select
End Sub

Private Sub Command5_Click()
CharList(UserCharIndex).Aura(1).Speed = Text99.Text
End Sub

Private Sub Command6_Click()
If UserEstado = 1 Then 'Muerto
    With FontTypes(FontTypeNames.FONTTYPE_INFO)
        Call ShowConsoleMsg("¡¡Estás muerto!!", .Red, .Green, .Blue, .bold, .italic)
    End With
    Exit Sub
End If

If Cuenta.Premium = True Then
    Call WriteBankStartCuenta
Else
    MsgBox "Necesitas ser usuario PREMIUM para poder obtener una boveda de cuenta de 10 slots."
End If
End Sub

Private Sub Command7_Click()
    Load_Auras
End Sub

Private Sub Command8_Click()
Dim X As Long

    For X = 0 To 6
        MapData(UserPos.X + X, UserPos.Y).Vertex_Offset(3) = MapData(UserPos.X + X, UserPos.Y).Vertex_Offset(3) + 1
    Next X
    
        MapData(UserPos.X + 7, UserPos.Y).Vertex_Offset(2) = MapData(UserPos.X + X, UserPos.Y).Vertex_Offset(2) + 1
        
    For X = 0 To 6
        MapData(UserPos.X + X, UserPos.Y + 1).Vertex_Offset(1) = MapData(UserPos.X + X, UserPos.Y + 1).Vertex_Offset(1) + 1
    Next X
    
        MapData(UserPos.X + 7, UserPos.Y + 1).Vertex_Offset(0) = MapData(UserPos.X + X, UserPos.Y + 1).Vertex_Offset(0) + 1
                
        
        
End Sub

Private Sub Command9_Click(index As Integer)
    With MapData(UserPos.X, UserPos.Y)
    Select Case index
        Case 0
            .Vertex_Offset(0) = .Vertex_Offset(0) - 5
        Case 1
            .Vertex_Offset(0) = .Vertex_Offset(0) + 5
            
        Case 2
            .Vertex_Offset(1) = .Vertex_Offset(1) - 5
        Case 3
            .Vertex_Offset(1) = .Vertex_Offset(1) + 5
            
        Case 4
            .Vertex_Offset(2) = .Vertex_Offset(2) - 5
        Case 5
            .Vertex_Offset(2) = .Vertex_Offset(2) + 5
            
        Case 6
            .Vertex_Offset(3) = .Vertex_Offset(3) - 5
        Case 7
            .Vertex_Offset(3) = .Vertex_Offset(3) + 5
    End Select
    

    'Set Ambient Vertex
    CurMapAmbient.MapBlocks(UserPos.X, UserPos.Y).Vertex_Offset(0) = .Vertex_Offset(0)
    CurMapAmbient.MapBlocks(UserPos.X, UserPos.Y).Vertex_Offset(1) = .Vertex_Offset(1)
    CurMapAmbient.MapBlocks(UserPos.X, UserPos.Y).Vertex_Offset(2) = .Vertex_Offset(2)
    CurMapAmbient.MapBlocks(UserPos.X, UserPos.Y).Vertex_Offset(3) = .Vertex_Offset(3)
    End With
End Sub

Private Sub Text11_Change(index As Integer)
    MapData(UserPos.X, UserPos.Y).Vertex_Offset(index) = Val(Text11(index).Text)
    'Set Ambient Vertex
    CurMapAmbient.MapBlocks(UserPos.X, UserPos.Y).Vertex_Offset(index) = MapData(UserPos.X, UserPos.Y).Vertex_Offset(index)
End Sub

Private Sub Timer1_Timer()
    If Me.Visible = True Then
        Dim i As Integer
            For i = 0 To 3
                Label4(i).Caption = MapData(UserPos.X, UserPos.Y).Vertex_Offset(i)
                Text11(i).Text = MapData(UserPos.X, UserPos.Y).Vertex_Offset(i)
            Next i
    End If
End Sub
