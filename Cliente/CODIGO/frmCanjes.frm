VERSION 5.00
Begin VB.Form frmCanjes 
   Caption         =   "Centro de canjes"
   ClientHeight    =   3870
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   3165
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
   ScaleHeight     =   3870
   ScaleWidth      =   3165
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Cancelar"
      Height          =   375
      Left            =   1680
      TabIndex        =   1
      Top             =   3360
      Width           =   1335
   End
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
      Height          =   2565
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   2895
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Cambiar Item"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   3360
      Width           =   1575
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00000000&
      FillColor       =   &H000000FF&
      Height          =   480
      Left            =   120
      ScaleHeight     =   420
      ScaleWidth      =   420
      TabIndex        =   0
      Top             =   2760
      Width           =   480
   End
   Begin VB.Label lblName 
      Caption         =   "Nada"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   495
      Left            =   3720
      TabIndex        =   5
      Top             =   120
      Width           =   1665
   End
   Begin VB.Label lblInfo 
      Caption         =   "Slot vacio, No hay ítem para canjear."
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   540
      Left            =   720
      TabIndex        =   4
      Top             =   2760
      Width           =   2265
   End
End
Attribute VB_Name = "frmCanjes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' ProyectoAO Documentación ***********************************
' Autor: Standelf
' Descripción: Formulario de Canjes
' Última modificación: 07/11/2012
'*************************************************************


Private Sub list1_Click()
    If List1.Text = "Nada" Then
        lblName.Caption = "Nada"
        lblInfo.Caption = "Slot vacio, No hay ítem para canjear."
        lblName.ForeColor = &HC0&
        lblInfo.ForeColor = &HC0&
        Command1.Enabled = False
        Picture1.Cls
    Else
        lblName.Caption = Inventario.ItemName(Val(List1.ListIndex + 1))
        lblInfo.Caption = "Valor Inicial: " & FormatNumber(Inventario.Valor(Val(List1.ListIndex + 1)), 0) & vbCrLf & _
                                "Cantidad Disponible: " & Inventario.Amount(Val(List1.ListIndex + 1)) & vbCrLf & _
                                "MinHit/MaxHit: " & Inventario.MinHit(Val(List1.ListIndex + 1)) & "/" & Inventario.MaxHit(Val(List1.ListIndex + 1)) & vbCrLf & _
                                "MinDef/MaxDef: " & Inventario.MinDef(Val(List1.ListIndex + 1)) & "/" & Inventario.MaxDef(Val(List1.ListIndex + 1))
        
        lblName.ForeColor = &H8000000C
        lblInfo.ForeColor = &H8000000C
        Command1.Enabled = True
        
        Call RenderItem(Picture1, Inventario.GrhIndex(Val(List1.ListIndex + 1)))
    End If

End Sub

