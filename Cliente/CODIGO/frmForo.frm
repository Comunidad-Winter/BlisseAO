VERSION 5.00
Begin VB.Form frmForo 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   0  'None
   ClientHeight    =   6855
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6210
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   457
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   414
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin CLBLISSEAO.BGAOButton imgDejarMsg 
      Height          =   285
      Left            =   840
      TabIndex        =   7
      Top             =   6000
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   503
      Caption         =   "Dejar Mensaje"
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
   Begin CLBLISSEAO.BGAOButton imgListaMsg 
      Height          =   285
      Left            =   2430
      TabIndex        =   6
      Top             =   6000
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   503
      Caption         =   "Ver Mensajes"
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
      Left            =   4080
      TabIndex        =   5
      Top             =   6000
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
   Begin VB.TextBox txtTitulo 
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
      Height          =   315
      Left            =   1140
      MaxLength       =   35
      TabIndex        =   2
      Top             =   900
      Visible         =   0   'False
      Width           =   4620
   End
   Begin VB.TextBox txtPost 
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
      Height          =   3960
      Left            =   780
      MultiLine       =   -1  'True
      TabIndex        =   1
      Text            =   "frmForo.frx":0000
      Top             =   1935
      Visible         =   0   'False
      Width           =   4770
   End
   Begin VB.ListBox lstTitulos 
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
      Height          =   5100
      Left            =   765
      TabIndex        =   0
      Top             =   825
      Width           =   4785
   End
   Begin CLBLISSEAO.BGAOButton imgDejarAnuncio 
      Height          =   285
      Left            =   2430
      TabIndex        =   8
      Top             =   6000
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   503
      Caption         =   "Dejar anuncio"
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
   Begin VB.Label lblTitulo 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
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
      Left            =   1125
      TabIndex        =   4
      Top             =   960
      Width           =   4695
   End
   Begin VB.Label lblAutor 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Nombre"
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
      Left            =   1125
      TabIndex        =   3
      Top             =   1455
      Width           =   4650
   End
   Begin VB.Image imgTab 
      Height          =   255
      Index           =   2
      Left            =   4320
      Top             =   360
      Width           =   1575
   End
   Begin VB.Image imgTab 
      Height          =   255
      Index           =   1
      Left            =   2520
      Top             =   360
      Width           =   1575
   End
   Begin VB.Image imgTab 
      Height          =   255
      Index           =   0
      Left            =   960
      Top             =   360
      Width           =   1575
   End
End
Attribute VB_Name = "frmForo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private clsFormulario As clsFormMovementManager

'   Para controlar las imagenes de fondo y el envio de posteos
Private ForoActual As eForumType
Private VerListaMsg As Boolean
Private Lectura As Boolean

Public ForoLimpio As Boolean
Private Sticky As Boolean

'   Para restringir la visibilidad de los foros
Public Privilegios As Byte
Public ForosVisibles As eForumType
Public CanPostSticky As Byte

'   Imagenes de fondo
Private FondosDejarMsg(0 To 2) As Picture
Private FondosListaMsg(0 To 2) As Picture

Private PuedeDejar As Boolean

Private Sub Form_Unload(Cancel As Integer)
    MirandoForo = False
    Privilegios = 0
End Sub

Private Sub imgDejarAnuncio_Click()
    Lectura = False
    VerListaMsg = False
    Sticky = True
    
    'Switch to proper background
    ToogleScreen
End Sub

Private Sub imgDejarMsg_Click()
    If Not PuedeDejar Then Exit Sub
    
    Dim PostStyle As Byte
    
    If Not VerListaMsg Then
        If Not Lectura Then
        
            If Sticky Then
                PostStyle = GetStickyPost
            Else
                PostStyle = GetNormalPost
            End If

            Call WriteForumPost(txtTitulo.Text, txtPost.Text, PostStyle)
            
            '   Actualizo localmente
            Call clsForos.AddPost(ForoActual, txtTitulo.Text, UserName, txtPost.Text, Sticky)
            Call UpdateList
            
            VerListaMsg = True
        End If
    Else
        VerListaMsg = False
        Sticky = False
    End If
    
    Lectura = False
    
    'Switch to proper background
    ToogleScreen
End Sub

Private Sub imgCerrar_Click()
    Unload Me
End Sub

Private Sub imgListaMsg_Click()
    VerListaMsg = True
    ToogleScreen
End Sub


Private Sub Form_Load()
    '   Handles Form movement (drag and drop).
    Set clsFormulario = New clsFormMovementManager
    clsFormulario.Initialize Me
    
    LoadButtons
    
    '   Initial config
    ForoActual = eForumType.ieGeneral
    VerListaMsg = True
    UpdateList
    
    '   Default background
    ToogleScreen
    
    ForoLimpio = False
    MirandoForo = True
    
    '   Si no es caos o gms, no puede ver el tab de caos.
    If (Privilegios And eForumVisibility.ieCAOS_MEMBER) = 0 Then imgTab(2).Visible = False
    
    '   Si no es armada o gm, no puede ver el tab de armadas.
    If (Privilegios And eForumVisibility.ieREAL_MEMBER) = 0 Then imgTab(1).Visible = False
    
End Sub

Private Sub LoadButtons()
    Set FondosListaMsg(eForumType.ieGeneral) = General_Set_GUI("ForoGeneral")
    Set FondosListaMsg(eForumType.ieREAL) = General_Set_GUI("ForoReal")
    Set FondosListaMsg(eForumType.ieCAOS) = General_Set_GUI("ForoCaos")
    
    Set FondosDejarMsg(eForumType.ieGeneral) = General_Set_GUI("ForoMsjGeneral")
    Set FondosDejarMsg(eForumType.ieREAL) = General_Set_GUI("ForoMsjReal")
    Set FondosDejarMsg(eForumType.ieCAOS) = General_Set_GUI("ForoMsjCaos")
    
    imgDejarMsg.Init s_Normal
    imgListaMsg.Init s_Normal
    imgCerrar.Init s_Normal
    imgDejarAnuncio.Init s_Normal
    
End Sub

Private Sub imgTab_Click(Index As Integer)

    Call Audio.PlayWave(SND_CLICK)
    
    If Index <> ForoActual Then
        ForoActual = Index
        VerListaMsg = True
        Lectura = False
        UpdateList
        ToogleScreen
    Else
        If Not VerListaMsg Then
            VerListaMsg = True
            Lectura = False
            ToogleScreen
        End If
    End If
End Sub

Private Sub ToogleScreen()
    
    Dim PostOffset As Integer
    
    txtTitulo.Visible = Not VerListaMsg And Not Lectura
    lblTitulo.Visible = Not VerListaMsg And Lectura
    PuedeDejar = VerListaMsg Or Lectura
    
    txtPost.Visible = Not VerListaMsg
    
    imgDejarAnuncio.Visible = VerListaMsg And PuedeDejarAnuncios
    imgListaMsg.Visible = Not VerListaMsg
    lstTitulos.Visible = VerListaMsg
    
    If VerListaMsg Then
        Me.Picture = FondosListaMsg(ForoActual)
    Else
        If Lectura Then
            With lstTitulos
                PostOffset = .ItemData(.ListIndex)
                
                '   Normal post?
                If PostOffset < STICKY_FORUM_OFFSET Then
                    lblTitulo.Caption = Foros(ForoActual).GeneralTitle(PostOffset)
                    txtPost.Text = Foros(ForoActual).GeneralPost(PostOffset)
                    lblAutor.Caption = Foros(ForoActual).GeneralAuthor(PostOffset)
                
                '   Sticky post
                Else
                    PostOffset = PostOffset - STICKY_FORUM_OFFSET
                    
                    lblTitulo.Caption = Foros(ForoActual).StickyTitle(PostOffset)
                    txtPost.Text = Foros(ForoActual).StickyPost(PostOffset)
                    lblAutor.Caption = Foros(ForoActual).StickyAuthor(PostOffset)
                End If
            End With
        Else
            lblAutor.Caption = UserName
            txtTitulo.Text = vbNullString
            txtPost.Text = vbNullString
            
            txtTitulo.SetFocus
        End If
        
        txtPost.Locked = Lectura
        Me.Picture = FondosDejarMsg(ForoActual)
    End If
    
End Sub

Private Function PuedeDejarAnuncios() As Boolean
    
    '   No puede
    If CanPostSticky = 0 Then Exit Function

    If ForoActual = eForumType.ieGeneral Then
        '   Solo puede dejar en el general si es gm
        If CanPostSticky <> 2 Then Exit Function
    End If
    
    PuedeDejarAnuncios = True
    
End Function

Private Sub lstTitulos_Click()
    VerListaMsg = False
    Lectura = True
    ToogleScreen
End Sub

Private Sub txtPost_Change()
    If Lectura Then Exit Sub
    
    PuedeDejar = Len(txtTitulo.Text) <> 0 And Len(txtPost.Text) <> 0
End Sub

Private Sub txtTitulo_Change()
    If Lectura Then Exit Sub
    
    PuedeDejar = Len(txtTitulo.Text) <> 0 And Len(txtPost.Text) <> 0
End Sub

Private Sub UpdateList()
    Dim PostIndex As Long
    
    lstTitulos.Clear
    
    With lstTitulos
        '   Sticky first
        For PostIndex = 1 To clsForos.GetNroSticky(ForoActual)
            .AddItem "[ANUNCIO] " & Foros(ForoActual).StickyTitle(PostIndex) & " (" & Foros(ForoActual).StickyAuthor(PostIndex) & ")"
            .ItemData(.NewIndex) = STICKY_FORUM_OFFSET + PostIndex
        Next PostIndex
    
        '   Then normal posts
        For PostIndex = 1 To clsForos.GetNroPost(ForoActual)
            .AddItem Foros(ForoActual).GeneralTitle(PostIndex) & " (" & Foros(ForoActual).GeneralAuthor(PostIndex) & ")"
            .ItemData(.NewIndex) = PostIndex
        Next PostIndex
    End With
    
End Sub

Private Function GetStickyPost() As Byte
    Select Case ForoActual
        Case 0
            GetStickyPost = eForumMsgType.ieGENERAL_STICKY
            
        Case 1
            GetStickyPost = eForumMsgType.ieREAL_STICKY
            
        Case 2
            GetStickyPost = eForumMsgType.ieCAOS_STICKY
            
    End Select
    
End Function

Private Function GetNormalPost() As Byte
    Select Case ForoActual
        Case 0
            GetNormalPost = eForumMsgType.ieGeneral
            
        Case 1
            GetNormalPost = eForumMsgType.ieREAL
            
        Case 2
            GetNormalPost = eForumMsgType.ieCAOS
            
    End Select
    
End Function
