VERSION 5.00
Begin VB.Form frmParty 
   BorderStyle     =   0  'None
   ClientHeight    =   6420
   ClientLeft      =   0
   ClientTop       =   -75
   ClientWidth     =   5640
   LinkTopic       =   "Form1"
   ScaleHeight     =   428
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   376
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin CLBLISSEAO.BGAOButton imgExpulsar 
      Height          =   285
      Left            =   1290
      TabIndex        =   4
      Top             =   3480
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   503
      Caption         =   "Expulsar"
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
   Begin VB.TextBox SendTxt 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   555
      MaxLength       =   160
      MultiLine       =   -1  'True
      TabIndex        =   2
      TabStop         =   0   'False
      ToolTipText     =   "Chat"
      Top             =   720
      Width           =   4530
   End
   Begin VB.TextBox txtToAdd 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   1530
      MaxLength       =   20
      TabIndex        =   1
      Top             =   4365
      Width           =   2580
   End
   Begin VB.ListBox lstMembers 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1395
      Left            =   1530
      TabIndex        =   0
      Top             =   1590
      Width           =   2595
   End
   Begin CLBLISSEAO.BGAOButton imgLiderGrupo 
      Height          =   285
      Left            =   2850
      TabIndex        =   5
      Top             =   3480
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   503
      Caption         =   "Hacer Lider"
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
   Begin CLBLISSEAO.BGAOButton imgAgregar 
      Height          =   285
      Left            =   2070
      TabIndex        =   6
      Top             =   4800
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   503
      Caption         =   "Agregar"
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
   Begin CLBLISSEAO.BGAOButton imgDisolver 
      Height          =   285
      Left            =   330
      TabIndex        =   7
      Top             =   5400
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   503
      Caption         =   "Disolver Party"
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
   Begin CLBLISSEAO.BGAOButton imgCerrar 
      Height          =   285
      Left            =   3810
      TabIndex        =   8
      Top             =   5400
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
   Begin CLBLISSEAO.BGAOButton imgSalirParty 
      Height          =   285
      Left            =   330
      TabIndex        =   9
      Top             =   5400
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   503
      Caption         =   "SalirParty"
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
   Begin VB.Label lblTotalExp 
      BackStyle       =   0  'Transparent
      Caption         =   "000000"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   3075
      TabIndex        =   3
      Top             =   3150
      Width           =   1335
   End
End
Attribute VB_Name = "frmParty"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private clsFormulario As clsFormMovementManager

Private sPartyChat As String
Private Const LEADER_FORM_HEIGHT As Integer = 6015
Private Const NORMAL_FORM_HEIGHT As Integer = 4455
Private Const OFFSET_BUTTONS As Integer = 43 '   pixels


Private Sub Form_Load()
    '   Handles Form movement (drag and drop).
    Set clsFormulario = New clsFormMovementManager
    clsFormulario.Initialize Me
    
    lstMembers.Clear
        
    If EsPartyLeader Then
        Me.Picture = General_Set_GUI("VentanaPartyLider")
        Me.Height = LEADER_FORM_HEIGHT
    Else
        Me.Picture = General_Set_GUI("VentanaPartyMiembro")
        Me.Height = NORMAL_FORM_HEIGHT
    End If
    
    imgAgregar.Init s_Normal
    imgCerrar.Init s_Normal
    imgDisolver.Init s_Normal
    imgLiderGrupo.Init s_Normal
    imgExpulsar.Init s_Normal
    imgSalirParty.Init s_Normal

    '   Botones visibles solo para el lider
    imgExpulsar.Visible = EsPartyLeader
    imgLiderGrupo.Visible = EsPartyLeader
    txtToAdd.Visible = EsPartyLeader
    imgAgregar.Visible = EsPartyLeader
    
    imgDisolver.Visible = EsPartyLeader
    imgSalirParty.Visible = Not EsPartyLeader
    
    imgSalirParty.Top = Me.ScaleHeight - OFFSET_BUTTONS
    imgDisolver.Top = Me.ScaleHeight - OFFSET_BUTTONS
    imgCerrar.Top = Me.ScaleHeight - OFFSET_BUTTONS

    MirandoParty = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
    MirandoParty = False
End Sub

Private Sub imgAgregar_Click()
    If Len(txtToAdd) > 0 Then
        If Not IsNumeric(txtToAdd) Then
            Call WritePartyAcceptMember(Trim(txtToAdd.Text))
            Unload Me
            Call WriteRequestPartyForm
        End If
    End If
End Sub

Private Sub imgCerrar_Click()
    Unload Me
End Sub

Private Sub imgDisolver_Click()
    Call WritePartyLeave
    Unload Me
End Sub

Private Sub imgExpulsar_Click()
   
    If lstMembers.ListIndex < 0 Then Exit Sub
    
    Dim fname As String
    fname = GetName
    
    If fname <> "" Then
        Call WritePartyKick(fname)
        Unload Me
        
        '   Para que no llame al form si disolvió la party
        If UCase$(fname) <> UCase$(UserName) Then Call WriteRequestPartyForm
    End If

End Sub

Private Function GetName() As String
'**************************************************************
'Author: ZaMa
'Last Modify Date: 27/12/2009
'**************************************************************
    Dim sName As String
    
    sName = Trim(mid(lstMembers.List(lstMembers.ListIndex), 1, InStr(lstMembers.List(lstMembers.ListIndex), " (")))
    If Len(sName) > 0 Then GetName = sName
        
End Function

Private Sub imgLiderGrupo_Click()
    
    If lstMembers.ListIndex < 0 Then Exit Sub
    
    Dim sName As String
    sName = GetName
    
    If sName <> "" Then
        Call WritePartySetLeader(sName)
        Unload Me
        Call WriteRequestPartyForm
    End If
End Sub

Private Sub imgSalirParty_Click()
    Call WritePartyLeave
    Unload Me
End Sub

Private Sub SendTxt_Change()
'**************************************************************
'Author: Unknown
'Last Modify Date: 03/10/2009
'**************************************************************
    If Len(SendTxt.Text) > 160 Then
        sPartyChat = "Soy un cheater, avisenle a un gm"
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
        
        sPartyChat = SendTxt.Text
    End If
End Sub

Private Sub SendTxt_KeyPress(KeyAscii As Integer)
    If Not (KeyAscii = vbKeyBack) And _
       Not (KeyAscii >= vbKeySpace And KeyAscii <= 250) Then _
        KeyAscii = 0
End Sub

Private Sub SendTxt_KeyUp(KeyCode As Integer, Shift As Integer)
    'Send text
    If KeyCode = vbKeyReturn Then
        If LenB(sPartyChat) <> 0 Then Call WritePartyMessage(sPartyChat)
        
        sPartyChat = ""
        SendTxt.Text = ""
        KeyCode = 0
        SendTxt.SetFocus
    End If
End Sub


Private Sub txtToAdd_KeyPress(KeyAscii As Integer)
    If Not (KeyAscii = vbKeyBack) And _
       Not (KeyAscii >= vbKeySpace And KeyAscii <= 250) Then _
        KeyAscii = 0
End Sub

Private Sub txtToAdd_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then imgAgregar_Click
End Sub


