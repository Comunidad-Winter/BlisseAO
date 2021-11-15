VERSION 5.00
Begin VB.Form frmInvasion 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Crear Invasión"
   ClientHeight    =   2370
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   3615
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2370
   ScaleWidth      =   3615
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Detener Invasión actual"
      Height          =   375
      Left            =   120
      TabIndex        =   9
      Top             =   1920
      Width           =   3375
   End
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      Height          =   810
      Index           =   1
      Left            =   1695
      Picture         =   "frmInvasion.frx":0000
      ScaleHeight     =   750
      ScaleWidth      =   315
      TabIndex        =   2
      Top             =   360
      Width           =   375
   End
   Begin VB.Frame Frame1 
      Caption         =   "Lista de Invasiones"
      Height          =   1815
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3375
      Begin VB.PictureBox Picture1 
         AutoSize        =   -1  'True
         Height          =   660
         Index           =   5
         Left            =   1680
         Picture         =   "frmInvasion.frx":016F
         ScaleHeight     =   600
         ScaleWidth      =   720
         TabIndex        =   6
         Top             =   1080
         Width           =   780
      End
      Begin VB.PictureBox Picture1 
         AutoSize        =   -1  'True
         Height          =   825
         Index           =   6
         Left            =   1140
         Picture         =   "frmInvasion.frx":0518
         ScaleHeight     =   765
         ScaleWidth      =   330
         TabIndex        =   7
         Top             =   240
         Width           =   390
      End
      Begin VB.PictureBox Picture1 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Height          =   1365
         Index           =   4
         Left            =   2520
         Picture         =   "frmInvasion.frx":06D9
         ScaleHeight     =   1305
         ScaleWidth      =   735
         TabIndex        =   5
         Top             =   225
         Width           =   795
      End
      Begin VB.PictureBox Picture1 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Height          =   810
         Index           =   3
         Left            =   1995
         Picture         =   "frmInvasion.frx":0F1A
         ScaleHeight     =   750
         ScaleWidth      =   375
         TabIndex        =   4
         Top             =   240
         Width           =   435
      End
      Begin VB.PictureBox Picture1 
         AutoSize        =   -1  'True
         Height          =   855
         Index           =   2
         Left            =   615
         Picture         =   "frmInvasion.frx":1474
         ScaleHeight     =   795
         ScaleWidth      =   435
         TabIndex        =   3
         Top             =   240
         Width           =   495
      End
      Begin VB.PictureBox Picture1 
         AutoSize        =   -1  'True
         Height          =   885
         Index           =   0
         Left            =   120
         Picture         =   "frmInvasion.frx":1A29
         ScaleHeight     =   825
         ScaleWidth      =   405
         TabIndex        =   1
         Top             =   240
         Width           =   465
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Click en una imágen para iniciar invasión en el mapa actual."
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   120
         TabIndex        =   8
         Top             =   1155
         Width           =   1455
      End
   End
End
Attribute VB_Name = "frmInvasion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    WriteInitInvasion -1
End Sub
Private Sub Picture1_Click(Index As Integer)
    Select Case Index
        Case 0 'asesino
            WriteInitInvasion 531
        Case 1 'esqueleto :)
            WriteInitInvasion 503
        Case 2 'licantropo
            WriteInitInvasion 545
        Case 3 'liche
            WriteInitInvasion 554
        Case 4 'orco?
            WriteInitInvasion 541
        Case 5 'tarantula
            WriteInitInvasion 547
        Case 6 'zombie
            WriteInitInvasion 502
    End Select
    
    Unload Me
End Sub
