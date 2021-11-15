VERSION 5.00
Begin VB.Form frmRank 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   7500
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9750
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00FFFFFF&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   500
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   650
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Label LabelInfo 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   $"frmRank.frx":0000
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   855
      Left            =   360
      TabIndex        =   66
      Top             =   5760
      Width           =   9015
   End
   Begin VB.Label Jugador10 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   5
      Left            =   7560
      TabIndex        =   65
      Top             =   6240
      Width           =   1695
   End
   Begin VB.Label Jugador9 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   5
      Left            =   7560
      TabIndex        =   64
      Top             =   5760
      Width           =   1695
   End
   Begin VB.Label Jugador8 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   5
      Left            =   7560
      TabIndex        =   63
      Top             =   5280
      Width           =   1695
   End
   Begin VB.Label Jugador7 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   5
      Left            =   7560
      TabIndex        =   62
      Top             =   4800
      Width           =   1695
   End
   Begin VB.Label Jugador6 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   5
      Left            =   7560
      TabIndex        =   61
      Top             =   4320
      Width           =   1695
   End
   Begin VB.Label Jugador5 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   5
      Left            =   7560
      TabIndex        =   60
      Top             =   3840
      Width           =   1695
   End
   Begin VB.Label Jugador4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   5
      Left            =   7560
      TabIndex        =   59
      Top             =   3360
      Width           =   1695
   End
   Begin VB.Label Jugador3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
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
      Height          =   255
      Index           =   5
      Left            =   7560
      TabIndex        =   58
      Top             =   2880
      Width           =   1695
   End
   Begin VB.Label Jugador2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   5
      Left            =   7560
      TabIndex        =   57
      Top             =   2400
      Width           =   1695
   End
   Begin VB.Label Jugador1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
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
      Height          =   255
      Index           =   5
      Left            =   7560
      TabIndex        =   56
      Top             =   1920
      Width           =   1695
   End
   Begin VB.Label Jugador10 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   4
      Left            =   6480
      TabIndex        =   55
      Top             =   6240
      Width           =   975
   End
   Begin VB.Label Jugador9 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   4
      Left            =   6480
      TabIndex        =   54
      Top             =   5760
      Width           =   975
   End
   Begin VB.Label Jugador8 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   4
      Left            =   6480
      TabIndex        =   53
      Top             =   5280
      Width           =   975
   End
   Begin VB.Label Jugador7 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   4
      Left            =   6480
      TabIndex        =   52
      Top             =   4800
      Width           =   975
   End
   Begin VB.Label Jugador6 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   4
      Left            =   6480
      TabIndex        =   51
      Top             =   4320
      Width           =   975
   End
   Begin VB.Label Jugador5 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   4
      Left            =   6480
      TabIndex        =   50
      Top             =   3840
      Width           =   975
   End
   Begin VB.Label Jugador4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   4
      Left            =   6480
      TabIndex        =   49
      Top             =   3360
      Width           =   975
   End
   Begin VB.Label Jugador3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
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
      Height          =   255
      Index           =   4
      Left            =   6480
      TabIndex        =   48
      Top             =   2880
      Width           =   975
   End
   Begin VB.Label Jugador2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   4
      Left            =   6480
      TabIndex        =   47
      Top             =   2400
      Width           =   975
   End
   Begin VB.Label Jugador1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
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
      Height          =   255
      Index           =   4
      Left            =   6480
      TabIndex        =   46
      Top             =   1920
      Width           =   975
   End
   Begin VB.Label Jugador10 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   3
      Left            =   5280
      TabIndex        =   45
      Top             =   6240
      Width           =   975
   End
   Begin VB.Label Jugador9 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   3
      Left            =   5280
      TabIndex        =   44
      Top             =   5760
      Width           =   975
   End
   Begin VB.Label Jugador8 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   3
      Left            =   5280
      TabIndex        =   43
      Top             =   5280
      Width           =   975
   End
   Begin VB.Label Jugador7 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   3
      Left            =   5280
      TabIndex        =   42
      Top             =   4800
      Width           =   975
   End
   Begin VB.Label Jugador6 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   3
      Left            =   5280
      TabIndex        =   41
      Top             =   4320
      Width           =   975
   End
   Begin VB.Label Jugador5 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   3
      Left            =   5280
      TabIndex        =   40
      Top             =   3840
      Width           =   975
   End
   Begin VB.Label Jugador4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   3
      Left            =   5280
      TabIndex        =   39
      Top             =   3360
      Width           =   975
   End
   Begin VB.Label Jugador3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
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
      Height          =   255
      Index           =   3
      Left            =   5280
      TabIndex        =   38
      Top             =   2880
      Width           =   975
   End
   Begin VB.Label Jugador2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   3
      Left            =   5280
      TabIndex        =   37
      Top             =   2400
      Width           =   975
   End
   Begin VB.Label Jugador1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
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
      Height          =   255
      Index           =   3
      Left            =   5280
      TabIndex        =   36
      Top             =   1920
      Width           =   975
   End
   Begin VB.Label Jugador10 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   2
      Left            =   3240
      TabIndex        =   35
      Top             =   6240
      Width           =   2055
   End
   Begin VB.Label Jugador9 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   2
      Left            =   3240
      TabIndex        =   34
      Top             =   5760
      Width           =   2055
   End
   Begin VB.Label Jugador8 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   2
      Left            =   3240
      TabIndex        =   33
      Top             =   5280
      Width           =   2055
   End
   Begin VB.Label Jugador7 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   2
      Left            =   3240
      TabIndex        =   32
      Top             =   4800
      Width           =   2055
   End
   Begin VB.Label Jugador6 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   2
      Left            =   3240
      TabIndex        =   31
      Top             =   4320
      Width           =   2055
   End
   Begin VB.Label Jugador5 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   2
      Left            =   3240
      TabIndex        =   30
      Top             =   3840
      Width           =   2055
   End
   Begin VB.Label Jugador4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   2
      Left            =   3240
      TabIndex        =   29
      Top             =   3360
      Width           =   2055
   End
   Begin VB.Label Jugador3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
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
      Height          =   255
      Index           =   2
      Left            =   3240
      TabIndex        =   28
      Top             =   2880
      Width           =   2055
   End
   Begin VB.Label Jugador2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   2
      Left            =   3240
      TabIndex        =   27
      Top             =   2400
      Width           =   2055
   End
   Begin VB.Label Jugador1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
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
      Height          =   255
      Index           =   2
      Left            =   3240
      TabIndex        =   26
      Top             =   1920
      Width           =   2055
   End
   Begin VB.Label Jugador10 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   1
      Left            =   1200
      TabIndex        =   25
      Top             =   6240
      Width           =   2055
   End
   Begin VB.Label Jugador9 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   1
      Left            =   1200
      TabIndex        =   24
      Top             =   5760
      Width           =   2055
   End
   Begin VB.Label Jugador8 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   1
      Left            =   1200
      TabIndex        =   23
      Top             =   5280
      Width           =   2055
   End
   Begin VB.Label Jugador7 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   1
      Left            =   1200
      TabIndex        =   22
      Top             =   4800
      Width           =   2055
   End
   Begin VB.Label Jugador6 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   1
      Left            =   1200
      TabIndex        =   21
      Top             =   4320
      Width           =   2055
   End
   Begin VB.Label Jugador5 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   1
      Left            =   1200
      TabIndex        =   20
      Top             =   3840
      Width           =   2055
   End
   Begin VB.Label Jugador4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   1
      Left            =   1200
      TabIndex        =   19
      Top             =   3360
      Width           =   2055
   End
   Begin VB.Label Jugador3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
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
      Height          =   255
      Index           =   1
      Left            =   1200
      TabIndex        =   18
      Top             =   2880
      Width           =   2055
   End
   Begin VB.Label Jugador2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   1
      Left            =   1200
      TabIndex        =   17
      Top             =   2400
      Width           =   2055
   End
   Begin VB.Label Jugador1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
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
      Height          =   255
      Index           =   1
      Left            =   1200
      TabIndex        =   16
      Top             =   1920
      Width           =   2055
   End
   Begin VB.Label Jugador10 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "10"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   0
      Left            =   480
      TabIndex        =   15
      Top             =   6240
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Jugador9 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "9"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   0
      Left            =   480
      TabIndex        =   14
      Top             =   5760
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Jugador8 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "8"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   0
      Left            =   480
      TabIndex        =   13
      Top             =   5280
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Jugador7 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "7"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   0
      Left            =   480
      TabIndex        =   12
      Top             =   4800
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Jugador6 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "6"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   0
      Left            =   480
      TabIndex        =   11
      Top             =   4320
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Jugador5 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "5"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   0
      Left            =   480
      TabIndex        =   10
      Top             =   3840
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Jugador4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "4"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   0
      Left            =   480
      TabIndex        =   9
      Top             =   3360
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Jugador3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "3"
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
      Height          =   255
      Index           =   0
      Left            =   480
      TabIndex        =   8
      Top             =   2880
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Jugador2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "2"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   0
      Left            =   480
      TabIndex        =   7
      Top             =   2400
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Jugador1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "1"
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
      Height          =   255
      Index           =   0
      Left            =   480
      TabIndex        =   6
      Top             =   1920
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "FRAGS"
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
      Height          =   255
      Index           =   5
      Left            =   6480
      TabIndex        =   5
      Top             =   1440
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "PRESTIGIO"
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
      Height          =   255
      Index           =   4
      Left            =   7560
      TabIndex        =   4
      Top             =   1440
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "NIVEL"
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
      Height          =   255
      Index           =   3
      Left            =   5280
      TabIndex        =   3
      Top             =   1440
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "CLAN"
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
      Height          =   255
      Index           =   2
      Left            =   3240
      TabIndex        =   2
      Top             =   1440
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "NICK"
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
      Height          =   255
      Index           =   1
      Left            =   1200
      TabIndex        =   1
      Top             =   1440
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "POS"
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
      Height          =   255
      Index           =   0
      Left            =   480
      TabIndex        =   0
      Top             =   1440
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image2 
      Height          =   525
      Index           =   3
      Left            =   6705
      Top             =   6840
      Width           =   1830
   End
   Begin VB.Image Image2 
      Height          =   525
      Index           =   2
      Left            =   4875
      Top             =   6840
      Width           =   1830
   End
   Begin VB.Image Image2 
      Height          =   525
      Index           =   1
      Left            =   3060
      Top             =   6840
      Width           =   1830
   End
   Begin VB.Image Image2 
      Height          =   525
      Index           =   0
      Left            =   1230
      Top             =   6840
      Width           =   1830
   End
   Begin VB.Image Image1 
      Height          =   240
      Left            =   9450
      Top             =   45
      Width           =   255
   End
End
Attribute VB_Name = "frmRank"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private clsFormulario As clsFormMovementManager
Private RankWeb As String

Private Sub Form_Load()
    '   Handles Form movement (drag and drop).
    Set clsFormulario = New clsFormMovementManager
    clsFormulario.Initialize Me
    
    Me.Picture = Set_Interface("VentanaRanking")
End Sub

Private Sub Image1_Click()
    Unload Me
End Sub

Private Sub Image2_Click(Index As Integer)
    LabelInfo.Visible = False
    Select Case Index
        Case 0, 1, 3
            Dim i As Byte, AllVisible As Boolean
                AllVisible = (Index = 0)
                For i = 0 To 5
                    Label1(i).Visible = True
                    Jugador1(i).Visible = True
                    Jugador2(i).Visible = True
                    Jugador3(i).Visible = True
                    Jugador4(i).Visible = True
                    Jugador5(i).Visible = True
                    Jugador6(i).Visible = AllVisible
                    Jugador7(i).Visible = AllVisible
                    Jugador8(i).Visible = AllVisible
                    Jugador9(i).Visible = AllVisible
                    Jugador10(i).Visible = AllVisible
                Next i
                
            If Index = 0 Then Call ObtenerRank(1)
            If Index = 1 Then Call ObtenerRank(2)
            If Index = 3 Then Call ObtenerRank(3)
            
        Case 2
            MsgBox "No disponible en Beta.", vbOKOnly
    End Select
End Sub

Private Sub ObtenerRank(ByVal Modo As Byte)
Dim i As Byte
    Select Case Modo
        Case 1
            RankWeb = frmMain.InetBG.OpenURL(Client_Web & "clrank.php?accion=rankgeneral")
        Case 2
            RankWeb = frmMain.InetBG.OpenURL(Client_Web & "clrank.php?accion=rank&orderby=nivel")
        Case 3
            RankWeb = frmMain.InetBG.OpenURL(Client_Web & "clrank.php?accion=rank&orderby=Prestigio")
    End Select
'    ReadField

    For i = 1 To 5
        Jugador1(i).Caption = ReadField(i + 1, RankWeb, Asc(","))
        Jugador2(i).Caption = ReadField(i + 8, RankWeb, Asc(","))
        Jugador3(i).Caption = ReadField(i + 15, RankWeb, Asc(","))
        Jugador4(i).Caption = ReadField(i + 22, RankWeb, Asc(","))
        Jugador5(i).Caption = ReadField(i + 29, RankWeb, Asc(","))

        If Modo = 1 Then
            Jugador6(i).Caption = ReadField(i + 36, RankWeb, Asc(","))
            Jugador7(i).Caption = ReadField(i + 43, RankWeb, Asc(","))
            Jugador8(i).Caption = ReadField(i + 50, RankWeb, Asc(","))
            Jugador9(i).Caption = ReadField(i + 57, RankWeb, Asc(","))
            Jugador10(i).Caption = ReadField(i + 64, RankWeb, Asc(","))
        End If
    Next i
End Sub

Private Sub Jugador1_Click(Index As Integer)
    If Index = 1 Then
        If MsgBox("¿Desea ver las estadisticas del usuario en la web oficial?", vbYesNo) = vbYes Then
            Call ShellExecute(0, "Open", Client_Web & "estadisticaspj.php?user=" & LCase$(Jugador1(1).Caption), "", App.path, SW_SHOWNORMAL)
        End If
    End If
End Sub

Private Sub Jugador2_Click(Index As Integer)
    If Index = 1 Then
        If MsgBox("¿Desea ver las estadisticas del usuario en la web oficial?", vbYesNo) = vbYes Then
            Call ShellExecute(0, "Open", Client_Web & "estadisticaspj.php?user=" & LCase$(Jugador2(1).Caption), "", App.path, SW_SHOWNORMAL)
        End If
    End If
End Sub

Private Sub Jugador3_Click(Index As Integer)
    If Index = 1 Then
        If MsgBox("¿Desea ver las estadisticas del usuario en la web oficial?", vbYesNo) = vbYes Then
            Call ShellExecute(0, "Open", Client_Web & "estadisticaspj.php?user=" & LCase$(Jugador3(1).Caption), "", App.path, SW_SHOWNORMAL)
        End If
    End If
End Sub

Private Sub Jugador4_Click(Index As Integer)
    If Index = 1 Then
        If MsgBox("¿Desea ver las estadisticas del usuario en la web oficial?", vbYesNo) = vbYes Then
            Call ShellExecute(0, "Open", Client_Web & "estadisticaspj.php?user=" & LCase$(Jugador4(1).Caption), "", App.path, SW_SHOWNORMAL)
        End If
    End If
End Sub

Private Sub Jugador5_Click(Index As Integer)
    If Index = 1 Then
        If MsgBox("¿Desea ver las estadisticas del usuario en la web oficial?", vbYesNo) = vbYes Then
            Call ShellExecute(0, "Open", Client_Web & "estadisticaspj.php?user=" & LCase$(Jugador5(1).Caption), "", App.path, SW_SHOWNORMAL)
        End If
    End If
End Sub

Private Sub Jugador6_Click(Index As Integer)
    If Index = 1 Then
        If MsgBox("¿Desea ver las estadisticas del usuario en la web oficial?", vbYesNo) = vbYes Then
            Call ShellExecute(0, "Open", Client_Web & "estadisticaspj.php?user=" & LCase$(Jugador6(1).Caption), "", App.path, SW_SHOWNORMAL)
        End If
    End If
End Sub

Private Sub Jugador7_Click(Index As Integer)
    If Index = 1 Then
        If MsgBox("¿Desea ver las estadisticas del usuario en la web oficial?", vbYesNo) = vbYes Then
            Call ShellExecute(0, "Open", Client_Web & "estadisticaspj.php?user=" & LCase$(Jugador7(1).Caption), "", App.path, SW_SHOWNORMAL)
        End If
    End If
End Sub

Private Sub Jugador8_Click(Index As Integer)
    If Index = 1 Then
        If MsgBox("¿Desea ver las estadisticas del usuario en la web oficial?", vbYesNo) = vbYes Then
            Call ShellExecute(0, "Open", Client_Web & "estadisticaspj.php?user=" & LCase$(Jugador8(1).Caption), "", App.path, SW_SHOWNORMAL)
        End If
    End If
End Sub

Private Sub Jugador9_Click(Index As Integer)
    If Index = 1 Then
        If MsgBox("¿Desea ver las estadisticas del usuario en la web oficial?", vbYesNo) = vbYes Then
            Call ShellExecute(0, "Open", Client_Web & "estadisticaspj.php?user=" & LCase$(Jugador9(1).Caption), "", App.path, SW_SHOWNORMAL)
        End If
    End If
End Sub

Private Sub Jugador10_Click(Index As Integer)
    If Index = 1 Then
        If MsgBox("¿Desea ver las estadisticas del usuario en la web oficial?", vbYesNo) = vbYes Then
            Call ShellExecute(0, "Open", Client_Web & "estadisticaspj.php?user=" & LCase$(Jugador10(1).Caption), "", App.path, SW_SHOWNORMAL)
        End If
    End If
End Sub

