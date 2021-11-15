VERSION 5.00
Begin VB.Form frmHerrero 
   BackColor       =   &H8000000A&
   BorderStyle     =   0  'None
   Caption         =   "Herrero"
   ClientHeight    =   5385
   ClientLeft      =   0
   ClientTop       =   -75
   ClientWidth     =   6675
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   359
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   445
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtCantItems 
      Alignment       =   2  'Center
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
      Left            =   5175
      MaxLength       =   5
      TabIndex        =   14
      Text            =   "1"
      Top             =   2940
      Width           =   1050
   End
   Begin VB.PictureBox picUpgradeItem 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      Height          =   465
      Index           =   1
      Left            =   5430
      ScaleHeight     =   465
      ScaleWidth      =   480
      TabIndex        =   13
      Top             =   1560
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.PictureBox picUpgradeItem 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      Height          =   465
      Index           =   2
      Left            =   5400
      ScaleHeight     =   465
      ScaleWidth      =   480
      TabIndex        =   12
      Top             =   2355
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.PictureBox picUpgradeItem 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      Height          =   465
      Index           =   3
      Left            =   5430
      ScaleHeight     =   465
      ScaleWidth      =   480
      TabIndex        =   11
      Top             =   3150
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.PictureBox picUpgradeItem 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      Height          =   465
      Index           =   4
      Left            =   5430
      ScaleHeight     =   465
      ScaleWidth      =   480
      TabIndex        =   10
      Top             =   3945
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.ComboBox cboItemsCiclo 
      BackColor       =   &H80000006&
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
      Left            =   5325
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   4080
      Width           =   735
   End
   Begin VB.PictureBox picItem 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      Height          =   465
      Index           =   4
      Left            =   870
      ScaleHeight     =   465
      ScaleWidth      =   480
      TabIndex        =   9
      Top             =   3945
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.PictureBox picLingotes3 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      Height          =   465
      Left            =   1710
      ScaleHeight     =   465
      ScaleWidth      =   1440
      TabIndex        =   8
      Top             =   3945
      Visible         =   0   'False
      Width           =   1440
   End
   Begin VB.PictureBox picItem 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      Height          =   465
      Index           =   3
      Left            =   870
      ScaleHeight     =   465
      ScaleWidth      =   480
      TabIndex        =   7
      Top             =   3150
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.PictureBox picLingotes2 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      Height          =   465
      Left            =   1710
      ScaleHeight     =   465
      ScaleWidth      =   1440
      TabIndex        =   6
      Top             =   3150
      Visible         =   0   'False
      Width           =   1440
   End
   Begin VB.PictureBox picItem 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      Height          =   465
      Index           =   2
      Left            =   870
      ScaleHeight     =   465
      ScaleWidth      =   480
      TabIndex        =   5
      Top             =   2355
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.PictureBox picLingotes1 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      Height          =   465
      Left            =   1710
      ScaleHeight     =   465
      ScaleWidth      =   1440
      TabIndex        =   4
      Top             =   2355
      Visible         =   0   'False
      Width           =   1440
   End
   Begin VB.VScrollBar Scroll 
      Height          =   3135
      Left            =   450
      TabIndex        =   0
      Top             =   1410
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.PictureBox picLingotes0 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      Height          =   465
      Left            =   1710
      ScaleHeight     =   465
      ScaleWidth      =   1440
      TabIndex        =   3
      Top             =   1560
      Visible         =   0   'False
      Width           =   1440
   End
   Begin VB.PictureBox picItem 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      Height          =   465
      Index           =   1
      Left            =   870
      ScaleHeight     =   465
      ScaleWidth      =   480
      TabIndex        =   2
      Top             =   1560
      Visible         =   0   'False
      Width           =   480
   End
   Begin CLBLISSEAO.BGAOButton picConstruir0 
      Height          =   465
      Left            =   3360
      TabIndex        =   15
      Top             =   1560
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   820
      Caption         =   "Construir"
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
   Begin CLBLISSEAO.BGAOButton picMejorar0 
      Height          =   465
      Left            =   3360
      TabIndex        =   16
      Top             =   1560
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   820
      Caption         =   "Mejorar"
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
   Begin CLBLISSEAO.BGAOButton picConstruir1 
      Height          =   465
      Left            =   3360
      TabIndex        =   17
      Top             =   2400
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   820
      Caption         =   "Construir"
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
   Begin CLBLISSEAO.BGAOButton picMejorar1 
      Height          =   465
      Left            =   3360
      TabIndex        =   18
      Top             =   2355
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   820
      Caption         =   "Mejorar"
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
   Begin CLBLISSEAO.BGAOButton picConstruir3 
      Height          =   465
      Left            =   3360
      TabIndex        =   19
      Top             =   3960
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   820
      Caption         =   "Construir"
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
   Begin CLBLISSEAO.BGAOButton picMejorar3 
      Height          =   465
      Left            =   3360
      TabIndex        =   20
      Top             =   3945
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   820
      Caption         =   "Mejorar"
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
   Begin CLBLISSEAO.BGAOButton picConstruir2 
      Height          =   465
      Left            =   3360
      TabIndex        =   21
      Top             =   3150
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   820
      Caption         =   "Construir"
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
   Begin CLBLISSEAO.BGAOButton picMejorar2 
      Height          =   465
      Left            =   3360
      TabIndex        =   22
      Top             =   3150
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   820
      Caption         =   "Mejorar"
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
      Height          =   465
      Left            =   3000
      TabIndex        =   23
      Top             =   4680
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   820
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
   Begin VB.Image picCheckBox 
      Height          =   420
      Left            =   5415
      MousePointer    =   99  'Custom
      Top             =   1860
      Width           =   435
   End
   Begin VB.Image picPestania 
      Height          =   255
      Index           =   2
      Left            =   3240
      MousePointer    =   99  'Custom
      Top             =   480
      Width           =   1095
   End
   Begin VB.Image picPestania 
      Height          =   255
      Index           =   1
      Left            =   1680
      MousePointer    =   99  'Custom
      Top             =   480
      Width           =   1455
   End
   Begin VB.Image picPestania 
      Height          =   255
      Index           =   0
      Left            =   720
      MousePointer    =   99  'Custom
      Top             =   480
      Width           =   975
   End
End
Attribute VB_Name = "frmHerrero"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Enum ePestania
    ieArmas
    ieArmaduras
    ieMejorar
End Enum

Private Pestanias(0 To 2) As Picture
Private UltimaPestania As Byte


Private Cargando As Boolean

Private UsarMacro As Boolean
Private Armas As Boolean

Private clsFormulario As clsFormMovementManager

Private Sub CargarImagenes()
    'Set Pestanias(ePestania.ieArmas) = LoadPicture(ImgPath & "VentanaHerreriaArmas.gif")
    'Set Pestanias(ePestania.ieArmaduras) = LoadPicture(ImgPath & "VentanaHerreriaArmaduras.gif")
    'Set Pestanias(ePestania.ieMejorar) = LoadPicture(ImgPath & "VentanaHerreriaMejorar.gif")
    
    'Set picCheck = LoadPicture(ImgPath & "CheckBoxHerreria.jpg")
    
    'Set picRecuadroItem = LoadPicture(ImgPath & "RecuadroItemsHerreria.jpg")
    'Set picRecuadroLingotes = LoadPicture(ImgPath & "RecuadroLingotes.jpg")
    
    'For Index = 1 To MAX_LIST_ITEMS
    '    imgMarcoItem(Index).Picture = picRecuadroItem
    '    imgMarcoUpgrade(Index).Picture = picRecuadroItem
    '    imgMarcoLingotes(Index).Picture = picRecuadroLingotes
    'Next Index
    
    'Set cPicCerrar = New clsGraphicalButton
    'Set cPicConstruir(0) = New clsGraphicalButton
    'Set cPicConstruir(1) = New clsGraphicalButton
    'Set cPicConstruir(2) = New clsGraphicalButton
    'Set cPicConstruir(3) = New clsGraphicalButton
    'Set cPicMejorar(0) = New clsGraphicalButton
    'Set cPicMejorar(1) = New clsGraphicalButton
    'Set cPicMejorar(2) = New clsGraphicalButton
    'Set cPicMejorar(3) = New clsGraphicalButton
'
'    Set LastPressed = New clsGraphicalButton
'
'    Call cPicCerrar.Initialize(imgCerrar, ImgPath & "BotonCerrarHerreria.jpg", ImgPath & "BotonCerrarRolloverHerreria.jpg", ImgPath & "BotonCerrarClickHerreria.jpg", Me)
    'Call cPicConstruir(0).Initialize(picConstruir0, ImgPath & "BotonConstruirHerreria.jpg", ImgPath & "BotonConstruirRolloverHerreria.jpg", ImgPath & "BotonConstruirClickHerreria.jpg", Me)
    'Call cPicConstruir(1).Initialize(picConstruir1, ImgPath & "BotonConstruirHerreria.jpg", ImgPath & "BotonConstruirRolloverHerreria.jpg", ImgPath & "BotonConstruirClickHerreria.jpg", Me)
    'Call cPicConstruir(2).Initialize(picConstruir2, ImgPath & "BotonConstruirHerreria.jpg", ImgPath & "BotonConstruirRolloverHerreria.jpg", ImgPath & "BotonConstruirClickHerreria.jpg", Me)
    'Call cPicConstruir(3).Initialize(picConstruir3, ImgPath & "BotonConstruirHerreria.jpg", ImgPath & "BotonConstruirRolloverHerreria.jpg", ImgPath & "BotonConstruirClickHerreria.jpg", Me)
    'Call cPicMejorar(0).Initialize(picMejorar0, ImgPath & "BotonMejorarHerreria.jpg", ImgPath & "BotonMejorarRolloverHerreria.jpg", ImgPath & "BotonMejorarClickHerreria.jpg", Me)
    'Call cPicMejorar(1).Initialize(picMejorar1, ImgPath & "BotonMejorarHerreria.jpg", ImgPath & "BotonMejorarRolloverHerreria.jpg", ImgPath & "BotonMejorarClickHerreria.jpg", Me)
    'Call cPicMejorar(2).Initialize(picMejorar2, ImgPath & "BotonMejorarHerreria.jpg", ImgPath & "BotonMejorarRolloverHerreria.jpg", ImgPath & "BotonMejorarClickHerreria.jpg", Me)
    'Call cPicMejorar(3).Initialize(picMejorar3, ImgPath & "BotonMejorarHerreria.jpg", ImgPath & "BotonMejorarRolloverHerreria.jpg", ImgPath & "BotonMejorarClickHerreria.jpg", Me)
'
'    imgCantidadCiclo.Picture = LoadPicture(ImgPath & "ConstruirPorCiclo.jpg")
'
End Sub

Private Sub ConstruirItem(ByVal Index As Integer)
    Dim ItemIndex As Integer
    Dim CantItemsCiclo As Integer
    
    If Scroll.Visible = True Then ItemIndex = Scroll.value
    ItemIndex = ItemIndex + Index
    
    Select Case UltimaPestania
        Case ePestania.ieArmas
        
            If UsarMacro Then
                CantItemsCiclo = Val(cboItemsCiclo.Text)
                MacroBltIndex = ArmasHerrero(ItemIndex).OBJIndex
                frmMain.ActivarMacroTrabajo
            Else
                '   Que cosntruya el maximo, total si sobra no importa, valida el server
                CantItemsCiclo = Val(cboItemsCiclo.List(cboItemsCiclo.ListCount - 1))
            End If
            
            Call WriteInitCrafting(Val(txtCantItems.Text), CantItemsCiclo)
            Call WriteCraftBlacksmith(ArmasHerrero(ItemIndex).OBJIndex)
            
        Case ePestania.ieArmaduras
        
            If UsarMacro Then
                CantItemsCiclo = Val(cboItemsCiclo.Text)
                MacroBltIndex = ArmadurasHerrero(ItemIndex).OBJIndex
                frmMain.ActivarMacroTrabajo
             Else
                '   Que cosntruya el maximo, total si sobra no importa, valida el server
                CantItemsCiclo = Val(cboItemsCiclo.List(cboItemsCiclo.ListCount - 1))
            End If
            
            Call WriteInitCrafting(Val(txtCantItems.Text), CantItemsCiclo)
            Call WriteCraftBlacksmith(ArmadurasHerrero(ItemIndex).OBJIndex)
        
        Case ePestania.ieMejorar
            Call WriteItemUpgrade(HerreroMejorar(ItemIndex).OBJIndex)
    End Select
    
    Unload Me

End Sub

Private Sub Form_Load()
    Dim MaxConstItem As Integer
    Dim i As Integer
    
    '   Handles Form movement (drag and drop).
    Set clsFormulario = New clsFormMovementManager
    clsFormulario.Initialize Me
    
    CargarImagenes
    
    '   Cargar imagenes
    Set Me.Picture = Pestanias(ePestania.ieArmas)
    'picCheckBox.Picture = picCheck
    
    Cargando = True
    
    MaxConstItem = CInt((UserLvl - 4) / 5)
    MaxConstItem = IIf(MaxConstItem < 1, 1, MaxConstItem)
    
    For i = 1 To MaxConstItem
        cboItemsCiclo.AddItem i
    Next i
    
    cboItemsCiclo.ListIndex = 0
    
    Cargando = False
    
    UsarMacro = True
    Armas = True
    UltimaPestania = 0
End Sub

Public Sub HideExtraControls(ByVal NumItems As Integer, Optional ByVal Upgrading As Boolean = False)
    Dim i As Integer
    
    picLingotes0.Visible = (NumItems >= 1)
    picLingotes1.Visible = (NumItems >= 2)
    picLingotes2.Visible = (NumItems >= 3)
    picLingotes3.Visible = (NumItems >= 4)
    
    For i = 1 To MAX_LIST_ITEMS
        picItem(i).Visible = (NumItems >= i)
        picUpgradeItem(i).Visible = (NumItems >= i And Upgrading)
    Next i
    
    picConstruir0.Visible = (NumItems >= 1 And Not Upgrading)
    picConstruir1.Visible = (NumItems >= 2 And Not Upgrading)
    picConstruir2.Visible = (NumItems >= 3 And Not Upgrading)
    picConstruir3.Visible = (NumItems >= 4 And Not Upgrading)
    
    picMejorar0.Visible = (NumItems >= 1 And Upgrading)
    picMejorar1.Visible = (NumItems >= 2 And Upgrading)
    picMejorar2.Visible = (NumItems >= 3 And Upgrading)
    picMejorar3.Visible = (NumItems >= 4 And Upgrading)
    
    picCheckBox.Visible = Not Upgrading
    cboItemsCiclo.Visible = Not Upgrading And UsarMacro

    txtCantItems.Visible = Not Upgrading
    picCheckBox.Visible = Not Upgrading
    
    If NumItems > MAX_LIST_ITEMS Then
        Scroll.Visible = True
        Cargando = True
        Scroll.max = NumItems - MAX_LIST_ITEMS
        Cargando = False
    Else
        Scroll.Visible = False
    End If
End Sub

Public Sub RenderList(ByVal Inicio As Integer, ByVal Armas As Boolean)
Dim i As Long
Dim NumItems As Integer
Dim ObjHerrero() As tItemsConstruibles

If Armas Then
    ObjHerrero = ArmasHerrero
Else
    ObjHerrero = ArmadurasHerrero
End If

NumItems = UBound(ObjHerrero)
Inicio = Inicio - 1

For i = 1 To MAX_LIST_ITEMS
    If i + Inicio <= NumItems Then
        With ObjHerrero(i + Inicio)
            '   Agrego el item
            Call RenderItem(picItem(i), .GrhIndex)
            picItem(i).ToolTipText = .name
            
             '   Inventariode lingotes
            Call InvLingosHerreria(i).SetItem(1, 0, .LinH, 0, LH_GRH, 0, 0, 0, 0, 0, 0, "Lingotes de Hierro")
            Call InvLingosHerreria(i).SetItem(2, 0, .LinP, 0, LP_GRH, 0, 0, 0, 0, 0, 0, "Lingotes de Plata")
            Call InvLingosHerreria(i).SetItem(3, 0, .LinO, 0, LO_GRH, 0, 0, 0, 0, 0, 0, "Lingotes de Oro")
        End With
    End If
Next i
End Sub

Public Sub RenderUpgradeList(ByVal Inicio As Integer)
Dim i As Long
Dim NumItems As Integer

NumItems = UBound(HerreroMejorar)
Inicio = Inicio - 1

For i = 1 To MAX_LIST_ITEMS
    If i + Inicio <= NumItems Then
        With HerreroMejorar(i + Inicio)
            '   Agrego el item
            Call RenderItem(picItem(i), .GrhIndex)
            picItem(i).ToolTipText = .name
            
            Call RenderItem(picUpgradeItem(i), .UpgradeGrhIndex)
            picUpgradeItem(i).ToolTipText = .UpgradeName
            
             '   Inventariode lingotes
            Call InvLingosHerreria(i).SetItem(1, 0, .LinH, 0, LH_GRH, 0, 0, 0, 0, 0, 0, "Lingotes de Hierro")
            Call InvLingosHerreria(i).SetItem(2, 0, .LinP, 0, LP_GRH, 0, 0, 0, 0, 0, 0, "Lingotes de Plata")
            Call InvLingosHerreria(i).SetItem(3, 0, .LinO, 0, LO_GRH, 0, 0, 0, 0, 0, 0, "Lingotes de Oro")
        End With
    End If
Next i
End Sub


Private Sub imgCantidadCiclo_Click()

End Sub

Private Sub imgCerrar_Click()
    Unload Me
End Sub

Private Sub picCheckBox_Click()
    
    UsarMacro = Not UsarMacro

    If UsarMacro Then
        'picCheckBox.Picture = picCheck
    Else
        'picCheckBox.Picture = Nothing
    End If
    
    cboItemsCiclo.Visible = UsarMacro
'    imgCantidadCiclo.Visible = UsarMacro
End Sub

Private Sub picConstruir0_Click()
    Call ConstruirItem(1)
End Sub

Private Sub picConstruir1_Click()
    Call ConstruirItem(2)
End Sub

Private Sub picConstruir2_Click()
    Call ConstruirItem(3)
End Sub

Private Sub picConstruir3_Click()
    Call ConstruirItem(4)
End Sub


Private Sub picMejorar0_Click()
    Call ConstruirItem(1)
End Sub

Private Sub picMejorar1_Click()
    Call ConstruirItem(2)
End Sub

Private Sub picMejorar2_Click()
    Call ConstruirItem(3)
End Sub

Private Sub picMejorar3_Click()
    Call ConstruirItem(4)
End Sub

Private Sub picPestania_Click(Index As Integer)
    Dim i As Integer
    Dim NumItems As Integer
    
    If Cargando Then Exit Sub
    If UltimaPestania = Index Then Exit Sub
    
    Scroll.value = 0
    
    Select Case Index
        Case ePestania.ieArmas
            '   Background
            Me.Picture = Pestanias(ePestania.ieArmas)
            
            NumItems = UBound(ArmasHerrero)
        
            Call HideExtraControls(NumItems)
            
            '   Cargo inventarios e imagenes
            Call RenderList(1, True)
            
            Armas = True
            
        Case ePestania.ieArmaduras
            '   Background
            Me.Picture = Pestanias(ePestania.ieArmaduras)
            
            NumItems = UBound(ArmadurasHerrero)
        
            Call HideExtraControls(NumItems)
            
            '   Cargo inventarios e imagenes
            Call RenderList(1, False)
            
            Armas = False
            
        Case ePestania.ieMejorar
            '   Background
            Me.Picture = Pestanias(ePestania.ieMejorar)
            
            NumItems = UBound(HerreroMejorar)
            
            Call HideExtraControls(NumItems, True)
            
            Call RenderUpgradeList(1)
    End Select

    UltimaPestania = Index
End Sub

Private Sub Scroll_Change()
    Dim i As Long
    
    If Cargando Then Exit Sub
    
    i = Scroll.value
    '   Cargo inventarios e imagenes
    
    Select Case UltimaPestania
        Case ePestania.ieArmas
            Call RenderList(i + 1, True)
        Case ePestania.ieArmaduras
            Call RenderList(i + 1, False)
        Case ePestania.ieMejorar
            Call RenderUpgradeList(i + 1)
    End Select
End Sub

Private Sub txtCantItems_Change()
On Error GoTo ErrHandler
    If Val(txtCantItems.Text) < 0 Then
        txtCantItems.Text = 1
    End If
    
    If Val(txtCantItems.Text) > MAX_INVENTORY_OBJS Then
        txtCantItems.Text = MAX_INVENTORY_OBJS
    End If
    
    Exit Sub
    
ErrHandler:
    'If we got here the user may have pasted (Shift + Insert) a REALLY large number, causing an overflow, so we set amount back to 1
    txtCantItems.Text = MAX_INVENTORY_OBJS
End Sub

Private Sub txtCantItems_KeyPress(KeyAscii As Integer)
    If (KeyAscii <> 8) Then
        If (KeyAscii < 48 Or KeyAscii > 57) Then
            KeyAscii = 0
        End If
    End If
End Sub

