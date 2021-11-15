VERSION 5.00
Begin VB.Form frmCarp 
   BackColor       =   &H8000000B&
   BorderStyle     =   0  'None
   Caption         =   "Carpintero"
   ClientHeight    =   5430
   ClientLeft      =   0
   ClientTop       =   -105
   ClientWidth     =   6705
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   362
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   447
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin CLBLISSEAO.BGAOButton imgConstruir0 
      Height          =   465
      Left            =   3120
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
   Begin VB.PictureBox picUpgrade 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      Height          =   465
      Index           =   4
      Left            =   5430
      ScaleHeight     =   31
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   14
      Top             =   3945
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.PictureBox picUpgrade 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      Height          =   465
      Index           =   3
      Left            =   5430
      ScaleHeight     =   31
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   13
      Top             =   3150
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.PictureBox picUpgrade 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      Height          =   465
      Index           =   2
      Left            =   5430
      ScaleHeight     =   31
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   12
      Top             =   2355
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.PictureBox picUpgrade 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      Height          =   465
      Index           =   1
      Left            =   5430
      ScaleHeight     =   31
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   11
      Top             =   1560
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.TextBox txtCantItems 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
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
      Height          =   285
      Left            =   5175
      MaxLength       =   6
      TabIndex        =   1
      Text            =   "1"
      ToolTipText     =   "Ingrese la cantidad total de items a construir."
      Top             =   2925
      Width           =   1050
   End
   Begin VB.ComboBox cboItemsCiclo 
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
      Left            =   5400
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   4080
      Width           =   735
   End
   Begin VB.PictureBox picItem 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      Height          =   465
      Index           =   1
      Left            =   870
      ScaleHeight     =   31
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   10
      Top             =   1560
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.PictureBox picMaderas0 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      Height          =   465
      Left            =   1710
      ScaleHeight     =   31
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   64
      TabIndex        =   9
      Top             =   1560
      Visible         =   0   'False
      Width           =   960
   End
   Begin VB.VScrollBar Scroll 
      Height          =   3135
      Left            =   450
      TabIndex        =   0
      Top             =   1410
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.PictureBox picMaderas1 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      Height          =   465
      Left            =   1710
      ScaleHeight     =   31
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   64
      TabIndex        =   8
      Top             =   2355
      Visible         =   0   'False
      Width           =   960
   End
   Begin VB.PictureBox picItem 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      Height          =   465
      Index           =   2
      Left            =   870
      ScaleHeight     =   31
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   7
      Top             =   2355
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.PictureBox picMaderas2 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      Height          =   465
      Left            =   1710
      ScaleHeight     =   31
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   64
      TabIndex        =   6
      Top             =   3150
      Visible         =   0   'False
      Width           =   960
   End
   Begin VB.PictureBox picItem 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      Height          =   465
      Index           =   3
      Left            =   870
      ScaleHeight     =   31
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   5
      Top             =   3150
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.PictureBox picMaderas3 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      Height          =   465
      Left            =   1710
      ScaleHeight     =   31
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   64
      TabIndex        =   4
      Top             =   3945
      Visible         =   0   'False
      Width           =   960
   End
   Begin VB.PictureBox picItem 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      Height          =   465
      Index           =   4
      Left            =   870
      ScaleHeight     =   31
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   3
      Top             =   3945
      Visible         =   0   'False
      Width           =   480
   End
   Begin CLBLISSEAO.BGAOButton imgMejorar0 
      Height          =   465
      Left            =   3120
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
   Begin CLBLISSEAO.BGAOButton imgConstruir1 
      Height          =   465
      Left            =   3120
      TabIndex        =   17
      Top             =   2355
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
   Begin CLBLISSEAO.BGAOButton imgMejorar1 
      Height          =   465
      Left            =   3120
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
   Begin CLBLISSEAO.BGAOButton imgConstruir3 
      Height          =   465
      Left            =   3120
      TabIndex        =   19
      Top             =   3945
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
   Begin CLBLISSEAO.BGAOButton imgMejorar3 
      Height          =   465
      Left            =   3120
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
   Begin CLBLISSEAO.BGAOButton imgConstruir2 
      Height          =   465
      Left            =   3120
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
   Begin CLBLISSEAO.BGAOButton imgMejorar2 
      Height          =   465
      Left            =   3120
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
      Left            =   2760
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
   Begin VB.Image imgPestania 
      Height          =   255
      Index           =   2
      Left            =   2400
      MousePointer    =   99  'Custom
      Top             =   480
      Width           =   975
   End
   Begin VB.Image imgCantidadCiclo 
      Height          =   645
      Left            =   5160
      Top             =   3435
      Width           =   1110
   End
   Begin VB.Image imgPestania 
      Height          =   255
      Index           =   1
      Left            =   1680
      MousePointer    =   99  'Custom
      Top             =   480
      Width           =   1215
   End
   Begin VB.Image imgPestania 
      Height          =   255
      Index           =   0
      Left            =   720
      MousePointer    =   99  'Custom
      Top             =   480
      Width           =   975
   End
   Begin VB.Image imgChkMacro 
      Height          =   420
      Left            =   5415
      MousePointer    =   99  'Custom
      Top             =   1860
      Width           =   435
   End
End
Attribute VB_Name = "frmCarp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Cargando As Boolean
Private clsFormulario As clsFormMovementManager

Private Enum ePestania
    ieItems
    ieMejorar
End Enum

Private Pestanias(1) As Picture
Private UltimaPestania As Byte

Private UsarMacro As Boolean

Private Sub Form_Load()
    
    '   Handles Form movement (drag and drop).
    Set clsFormulario = New clsFormMovementManager
    clsFormulario.Initialize Me
    
    Call LoadDefaultValues
    
    Set Pestanias(ePestania.ieItems) = General_Set_GUI("VentanaCarpinteriaItems")
    Set Pestanias(ePestania.ieMejorar) = General_Set_GUI("VentanaCarpinteriaMejorar")
    
    Me.Picture = General_Set_GUI("VentanaCarpinteriaItems")
End Sub

Private Sub LoadDefaultValues()
    
    Dim MaxConstItem As Integer
    Dim i As Integer

    Cargando = True
    
    MaxConstItem = CInt((UserLvl - 4) / 5)
    MaxConstItem = IIf(MaxConstItem < 1, 1, MaxConstItem)
    
    For i = 1 To MaxConstItem
        cboItemsCiclo.AddItem i
    Next i
    
    cboItemsCiclo.ListIndex = 0
    
    Scroll.value = 0
    
    UsarMacro = True
    
    UltimaPestania = ePestania.ieItems
    
    Cargando = False
End Sub


Private Sub Construir(ByVal Index As Integer)

    Dim ItemIndex As Integer
    Dim CantItemsCiclo As Integer
    
    If Scroll.Visible = True Then ItemIndex = Scroll.value
    ItemIndex = ItemIndex + Index
    
    Select Case UltimaPestania
        Case ePestania.ieItems
        
            If UsarMacro Then
                CantItemsCiclo = Val(cboItemsCiclo.Text)
                MacroBltIndex = ObjCarpintero(ItemIndex).OBJIndex
                frmMain.ActivarMacroTrabajo
            Else
                '   Que cosntruya el maximo, total si sobra no importa, valida el server
                CantItemsCiclo = Val(cboItemsCiclo.List(cboItemsCiclo.ListCount - 1))
            End If
            
            Call WriteInitCrafting(Val(txtCantItems.Text), CantItemsCiclo)
            Call WriteCraftCarpenter(ObjCarpintero(ItemIndex).OBJIndex)
            
        Case ePestania.ieMejorar
            Call WriteItemUpgrade(CarpinteroMejorar(ItemIndex).OBJIndex)
    End Select
        
    Unload Me

End Sub

Public Sub HideExtraControls(ByVal NumItems As Integer, Optional ByVal Upgrading As Boolean = False)
    Dim i As Integer
    
    picMaderas0.Visible = (NumItems >= 1)
    picMaderas1.Visible = (NumItems >= 2)
    picMaderas2.Visible = (NumItems >= 3)
    picMaderas3.Visible = (NumItems >= 4)
    
    imgConstruir0.Visible = (NumItems >= 1 And Not Upgrading)
    imgConstruir1.Visible = (NumItems >= 2 And Not Upgrading)
    imgConstruir2.Visible = (NumItems >= 3 And Not Upgrading)
    imgConstruir3.Visible = (NumItems >= 4 And Not Upgrading)
    
    imgMejorar0.Visible = (NumItems >= 1 And Upgrading)
    imgMejorar1.Visible = (NumItems >= 2 And Upgrading)
    imgMejorar2.Visible = (NumItems >= 3 And Upgrading)
    imgMejorar3.Visible = (NumItems >= 4 And Upgrading)

    
    For i = 1 To MAX_LIST_ITEMS
        picItem(i).Visible = (NumItems >= i)

        '   Upgrade
        picUpgrade(i).Visible = (NumItems >= i And Upgrading)
    Next i
    
    If NumItems > MAX_LIST_ITEMS Then
        Scroll.Visible = True
        Cargando = True
        Scroll.max = NumItems - MAX_LIST_ITEMS
        Cargando = False
    Else
        Scroll.Visible = False
    End If
    
    txtCantItems.Visible = Not Upgrading
    cboItemsCiclo.Visible = Not Upgrading And UsarMacro
    imgChkMacro.Visible = Not Upgrading
    imgCantidadCiclo.Visible = Not Upgrading And UsarMacro
End Sub

Public Sub RenderList(ByVal Inicio As Integer)
Dim i As Long
Dim NumItems As Integer

NumItems = UBound(ObjCarpintero)
Inicio = Inicio - 1

For i = 1 To MAX_LIST_ITEMS
    If i + Inicio <= NumItems Then
        With ObjCarpintero(i + Inicio)
            '   Agrego el item
            Call RenderItem(picItem(i), .GrhIndex)
            picItem(i).ToolTipText = .name
        
            '   Inventario de leños
            Call InvMaderasCarpinteria(i).SetItem(1, 0, .Madera, 0, MADERA_GRH, 0, 0, 0, 0, 0, 0, "Leña")
            Call InvMaderasCarpinteria(i).SetItem(2, 0, .MaderaElfica, 0, MADERA_ELFICA_GRH, 0, 0, 0, 0, 0, 0, "Leña élfica")
        End With
    End If
Next i
End Sub

Public Sub RenderUpgradeList(ByVal Inicio As Integer)
Dim i As Long
Dim NumItems As Integer

NumItems = UBound(CarpinteroMejorar)
Inicio = Inicio - 1

For i = 1 To MAX_LIST_ITEMS
    If i + Inicio <= NumItems Then
        With CarpinteroMejorar(i + Inicio)
            '   Agrego el item
            Call RenderItem(picItem(i), .GrhIndex)
            picItem(i).ToolTipText = .name
            
            Call RenderItem(picUpgrade(i), .UpgradeGrhIndex)
            picUpgrade(i).ToolTipText = .UpgradeName
        
            '   Inventario de leños
            Call InvMaderasCarpinteria(i).SetItem(1, 0, .Madera, 0, MADERA_GRH, 0, 0, 0, 0, 0, 0, "Leña")
            Call InvMaderasCarpinteria(i).SetItem(2, 0, .MaderaElfica, 0, MADERA_ELFICA_GRH, 0, 0, 0, 0, 0, 0, "Leña élfica")
        End With
    End If
Next i
End Sub

Private Sub imgCerrar_Click()
    Unload Me
End Sub

Private Sub imgChkMacro_Click()
    UsarMacro = Not UsarMacro
    
    If UsarMacro Then
        'imgChkMacro.Picture = picCheck
    Else
        'Set imgChkMacro.Picture = Nothing
    End If
    
    cboItemsCiclo.Visible = UsarMacro
    'imgCantidadCiclo.Visible = UsarMacro
End Sub

Private Sub imgConstruir0_Click()
    Call Construir(1)
End Sub

Private Sub imgConstruir1_Click()
    Call Construir(2)
End Sub

Private Sub imgConstruir2_Click()
    Call Construir(3)
End Sub

Private Sub imgConstruir3_Click()
    Call Construir(4)
End Sub

Private Sub imgMejorar0_Click()
    Call Construir(1)
End Sub

Private Sub imgMejorar1_Click()
    Call Construir(2)
End Sub

Private Sub imgMejorar2_Click()
    Call Construir(3)
End Sub

Private Sub imgMejorar3_Click()
    Call Construir(4)
End Sub

Private Sub imgPestania_Click(Index As Integer)
    Dim i As Integer
    Dim NumItems As Integer
    
    If Cargando Then Exit Sub
    If UltimaPestania = Index Then Exit Sub
    
    Scroll.value = 0
    
    Select Case Index
        Case ePestania.ieItems
            '   Background
            Me.Picture = Pestanias(ePestania.ieItems)
            
            NumItems = UBound(ObjCarpintero)
        
            Call HideExtraControls(NumItems)
            
            '   Cargo inventarios e imagenes
            Call RenderList(1)
            

        Case ePestania.ieMejorar
            '   Background
            Me.Picture = Pestanias(ePestania.ieMejorar)
            
            NumItems = UBound(CarpinteroMejorar)
            
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
        Case ePestania.ieItems
            Call RenderList(i + 1)
        Case ePestania.ieMejorar
            Call RenderUpgradeList(i + 1)
    End Select
End Sub
