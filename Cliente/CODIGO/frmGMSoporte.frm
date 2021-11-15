VERSION 5.00
Begin VB.Form frmGMSoporte 
   BorderStyle     =   0  'None
   ClientHeight    =   4800
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4620
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
   ScaleHeight     =   320
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   308
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin CLBLISSEAO.BGAOButton Command1 
      Height          =   285
      Index           =   0
      Left            =   600
      TabIndex        =   8
      Top             =   4440
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   503
      Caption         =   "Cerrar"
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
   Begin VB.OptionButton Option1 
      Caption         =   "Consulta Regular"
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   5
      Top             =   2400
      Width           =   1815
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Reporte de BUG"
      Height          =   255
      Index           =   1
      Left            =   2280
      TabIndex        =   4
      Top             =   2400
      Width           =   1815
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Denunciar"
      Height          =   255
      Index           =   2
      Left            =   240
      TabIndex        =   3
      Top             =   2760
      Width           =   1815
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   855
      Left            =   330
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   3300
      Width           =   3945
   End
   Begin CLBLISSEAO.BGAOButton Command1 
      Height          =   285
      Index           =   1
      Left            =   2520
      TabIndex        =   9
      Top             =   4440
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   503
      Caption         =   "Enviar"
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
   Begin VB.Label Label2 
      BackColor       =   &H8000000D&
      BackStyle       =   0  'Transparent
      Caption         =   "Página Oficial."
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   225
      Left            =   1770
      TabIndex        =   7
      Top             =   1635
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmGMSoporte.frx":0000
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   975
      Left            =   240
      TabIndex        =   6
      Top             =   855
      Width           =   4215
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Información:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   210
      TabIndex        =   2
      Top             =   600
      Width           =   1935
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Motivo de Consulta:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   225
      TabIndex        =   1
      Top             =   2010
      Width           =   1935
   End
End
Attribute VB_Name = "frmGMSoporte"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private clsFormulario As clsFormMovementManager

Private Sub Command1_Click(Index As Integer)
    Select Case Index
        Case 0
            Unload Me
        Case 1
            If Option1(0).value = True Then
                    Call WriteGMRequest
                    Unload Me
                    
            ElseIf Option1(1).value = True Then
                    If LenB(Trim$(Text1.Text)) Then
                        Call WriteBugReport(Text1.Text)
                        Call ShowConsoleMsg("Bug enviado correctamente.")
                        Unload Me
                    Else
                        Call ShowConsoleMsg("Ingrese la descripción del BUG.")
                    End If
                    
            ElseIf Option1(2).value = True Then
                    If LenB(Trim$(Text1.Text)) Then
                        Call WriteDenounce(Text1.Text)
                        Unload Me
                    Else
                        Call ShowConsoleMsg("Ingrese la descripción de su Denuncia.")
                    End If
            End If
    End Select
End Sub

Private Sub Form_Load()

    Set clsFormulario = New clsFormMovementManager
    clsFormulario.Initialize Me
    
    Me.Picture = General_Set_GUI("VentanaGMSoporte")
    Command1(0).Init s_Normal
    Command1(1).Init s_Normal
    
End Sub

Private Sub Label2_Click()
    Call ShellExecute(0, "Open", Client_Web, "", App.path, SW_SHOWNORMAL)
End Sub
