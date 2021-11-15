VERSION 5.00
Begin VB.Form FrmPenas 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "  Penalidades a cuentas de los usuarios."
   ClientHeight    =   3180
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   3450
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
   ScaleHeight     =   3180
   ScaleWidth      =   3450
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdBAN 
      Caption         =   "BANEAR USUARIO"
      Height          =   285
      Left            =   90
      TabIndex        =   8
      Top             =   2835
      Width           =   3300
   End
   Begin VB.Frame Frame1 
      Caption         =   "  Suspender  "
      Height          =   1095
      Left            =   90
      TabIndex        =   4
      Top             =   1530
      Width           =   3300
      Begin VB.CommandButton Command2 
         Caption         =   "Aplicar suspención"
         Height          =   285
         Left            =   180
         TabIndex        =   7
         Top             =   675
         Width           =   2940
      End
      Begin VB.ComboBox CPenas 
         Height          =   315
         Left            =   2025
         TabIndex        =   6
         Text            =   "Combo1"
         Top             =   225
         Width           =   735
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Cantidad de dias:"
         Height          =   195
         Left            =   450
         TabIndex        =   5
         Top             =   270
         Width           =   1515
      End
   End
   Begin VB.CommandButton cmdADVERTENCIA 
      Caption         =   "Aplicar advertencia"
      Height          =   285
      Left            =   90
      TabIndex        =   3
      Top             =   1125
      Width           =   3300
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   3465
      Y1              =   990
      Y2              =   990
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "¿BANEADO?:"
      Height          =   195
      Left            =   90
      TabIndex        =   2
      Top             =   675
      Width           =   1110
   End
   Begin VB.Label LSuspenciones 
      AutoSize        =   -1  'True
      Caption         =   "Cantidad de Suspenciones:"
      Height          =   195
      Left            =   90
      TabIndex        =   1
      Top             =   405
      Width           =   2340
   End
   Begin VB.Label LAdvertencia 
      AutoSize        =   -1  'True
      Caption         =   "Cantidad de advertencias:"
      Height          =   195
      Left            =   90
      TabIndex        =   0
      Top             =   135
      Width           =   2265
   End
End
Attribute VB_Name = "FrmPenas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Ciclo As Byte
Dim tStr As String
Dim Nick As String

Private Sub cmdADVERTENCIA_Click()
    If LenB(Nick) <> 0 Then
        tStr = InputBox("Escriba el motivo de la advertencia.", "Advertir a " & Nick)
                
        If LenB(tStr) <> 0 Then
            Call ParseUserCommand("/ADVERTENCIA " & Nick & "@" & tStr)
        End If
    End If
End Sub

Private Sub cmdBAN_Click()
    If LenB(Nick) <> 0 Then
        tStr = InputBox("Escriba el motivo del ban.", "BAN a " & Nick)
                
        If LenB(tStr) <> 0 Then _
            If MsgBox("¿Seguro desea banear a " & Nick & "?", vbYesNo, "Atencion!") = vbYes Then _
                Call WriteBanChar(Nick, tStr)
    End If
End Sub

Private Sub Form_Load()
For Ciclo = 1 To 10
    CPenas.AddItem Ciclo
Next Ciclo

Nick = frmPanelGm.cboListaUsus.Text
End Sub
