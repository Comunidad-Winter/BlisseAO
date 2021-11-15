VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "Indexador BETA 1"
   ClientHeight    =   8430
   ClientLeft      =   120
   ClientTop       =   750
   ClientWidth     =   14520
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   562
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   968
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture3 
      AutoRedraw      =   -1  'True
      Height          =   4815
      Left            =   10800
      ScaleHeight     =   4755
      ScaleWidth      =   5235
      TabIndex        =   58
      Top             =   7155
      Width           =   5295
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Generar Color Finder"
      Height          =   615
      Left            =   8745
      TabIndex        =   57
      Top             =   7095
      Width           =   1815
   End
   Begin VB.TextBox Text6 
      Height          =   360
      Left            =   11115
      TabIndex        =   56
      Text            =   "0"
      Top             =   5775
      Width           =   1695
   End
   Begin VB.TextBox Text5 
      Height          =   885
      Left            =   8745
      MultiLine       =   -1  'True
      TabIndex        =   55
      Text            =   "frmMain.frx":0000
      Top             =   6135
      Width           =   4035
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Buscar GRHS del BMP:"
      Height          =   330
      Left            =   8745
      TabIndex        =   54
      Top             =   5760
      Width           =   2340
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Ir"
      Height          =   375
      Left            =   2160
      TabIndex        =   48
      Top             =   7920
      Width           =   615
   End
   Begin VB.TextBox Text3 
      Height          =   375
      Left            =   1080
      TabIndex        =   47
      Text            =   "1"
      Top             =   7920
      Width           =   975
   End
   Begin VB.ListBox GRH_LIST 
      Height          =   7500
      Left            =   120
      TabIndex        =   45
      Top             =   360
      Width           =   2655
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3255
      Left            =   8760
      ScaleHeight     =   3255
      ScaleWidth      =   4335
      TabIndex        =   35
      Top             =   2400
      Width           =   4335
      Begin VB.TextBox Text4 
         Height          =   360
         Index           =   3
         Left            =   3360
         TabIndex        =   53
         Text            =   "0"
         Top             =   2160
         Width           =   855
      End
      Begin VB.TextBox Text4 
         Height          =   360
         Index           =   2
         Left            =   3360
         TabIndex        =   52
         Text            =   "0"
         Top             =   1800
         Width           =   855
      End
      Begin VB.TextBox Text4 
         Height          =   360
         Index           =   1
         Left            =   3360
         TabIndex        =   51
         Text            =   "0"
         Top             =   1440
         Width           =   855
      End
      Begin VB.TextBox Text4 
         Height          =   360
         Index           =   0
         Left            =   3360
         TabIndex        =   50
         Text            =   "0"
         Top             =   1080
         Width           =   855
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Cascos"
         Height          =   255
         Index           =   1
         Left            =   2640
         TabIndex        =   43
         Top             =   720
         Width           =   1695
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Cabezas"
         Height          =   255
         Index           =   0
         Left            =   2640
         TabIndex        =   42
         Top             =   480
         Value           =   -1  'True
         Width           =   1695
      End
      Begin VB.ListBox HHH_LIST 
         Height          =   2940
         Left            =   0
         TabIndex        =   37
         Top             =   240
         Width           =   2055
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Aplicar Cambios"
         Height          =   495
         Index           =   2
         Left            =   2280
         TabIndex        =   36
         Top             =   2640
         Width           =   1935
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "Mostrar:"
         Height          =   255
         Index           =   17
         Left            =   2160
         TabIndex        =   49
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "Espalda:"
         Height          =   255
         Index           =   16
         Left            =   2040
         TabIndex        =   44
         Top             =   2160
         Width           =   1215
      End
      Begin VB.Label Label5 
         Caption         =   "Listado de Cabezas:"
         Height          =   255
         Left            =   0
         TabIndex        =   41
         Top             =   0
         Width           =   2055
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "Derecha:"
         Height          =   255
         Index           =   15
         Left            =   2160
         TabIndex        =   40
         Top             =   1080
         Width           =   1095
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "Izquierda:"
         Height          =   255
         Index           =   14
         Left            =   2160
         TabIndex        =   39
         Top             =   1440
         Width           =   1095
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "Frente:"
         Height          =   255
         Index           =   13
         Left            =   2040
         TabIndex        =   38
         Top             =   1800
         Width           =   1215
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Editar ANIM"
      Enabled         =   0   'False
      Height          =   375
      Left            =   7320
      TabIndex        =   34
      Top             =   2040
      Width           =   1335
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2055
      Left            =   8760
      ScaleHeight     =   2055
      ScaleWidth      =   4335
      TabIndex        =   24
      Top             =   120
      Width           =   4335
      Begin VB.CommandButton Command1 
         Caption         =   "Aplicar Cambios"
         Height          =   495
         Index           =   1
         Left            =   2280
         TabIndex        =   33
         Top             =   1440
         Width           =   1935
      End
      Begin VB.TextBox Text2 
         Height          =   375
         Index           =   2
         Left            =   3480
         TabIndex        =   32
         Text            =   "0"
         Top             =   960
         Width           =   735
      End
      Begin VB.TextBox Text2 
         Height          =   375
         Index           =   1
         Left            =   3360
         TabIndex        =   31
         Text            =   "0"
         Top             =   480
         Width           =   855
      End
      Begin VB.TextBox Text2 
         Height          =   375
         Index           =   0
         Left            =   3360
         TabIndex        =   30
         Text            =   "0"
         Top             =   120
         Width           =   855
      End
      Begin VB.ListBox FXS_LIST 
         Height          =   1740
         Left            =   0
         TabIndex        =   26
         Top             =   240
         Width           =   2055
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "Animacion:"
         Height          =   255
         Index           =   12
         Left            =   2160
         TabIndex        =   29
         Top             =   960
         Width           =   1215
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "OffSet Y:"
         Height          =   255
         Index           =   11
         Left            =   2160
         TabIndex        =   28
         Top             =   480
         Width           =   1095
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "OffSet X:"
         Height          =   255
         Index           =   10
         Left            =   2160
         TabIndex        =   27
         Top             =   120
         Width           =   1095
      End
      Begin VB.Label Label4 
         Caption         =   "Listado de FXs:"
         Height          =   255
         Left            =   0
         TabIndex        =   25
         Top             =   0
         Width           =   1575
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Aplicar Cambios"
      Height          =   375
      Index           =   0
      Left            =   5280
      TabIndex        =   23
      Top             =   2040
      Width           =   1935
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   9
      Left            =   7200
      TabIndex        =   22
      Text            =   "0"
      Top             =   720
      Width           =   855
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   8
      Left            =   7200
      TabIndex        =   21
      Text            =   "0"
      Top             =   360
      Width           =   855
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   7
      Left            =   7920
      TabIndex        =   20
      Text            =   "0"
      Top             =   1560
      Width           =   615
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   6
      Left            =   7920
      TabIndex        =   19
      Text            =   "0"
      Top             =   1200
      Width           =   615
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   5
      Left            =   5760
      TabIndex        =   15
      Text            =   "0"
      Top             =   1560
      Width           =   615
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   4
      Left            =   5760
      TabIndex        =   14
      Text            =   "0"
      Top             =   1200
      Width           =   615
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   3
      Left            =   4080
      TabIndex        =   6
      Text            =   "0"
      Top             =   1560
      Width           =   615
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   2
      Left            =   4080
      TabIndex        =   5
      Text            =   "0"
      Top             =   1200
      Width           =   615
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   1
      Left            =   4080
      TabIndex        =   4
      Text            =   "0"
      Top             =   720
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      Enabled         =   0   'False
      Height          =   375
      Index           =   0
      Left            =   4080
      TabIndex        =   3
      Text            =   "0"
      Top             =   360
      Width           =   1695
   End
   Begin VB.PictureBox Render 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   5760
      Left            =   2880
      MousePointer    =   2  'Cross
      ScaleHeight     =   384
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   384
      TabIndex        =   1
      Top             =   2550
      Width           =   5760
   End
   Begin VB.Label Label6 
      Caption         =   "Ir a GRH:"
      Height          =   375
      Left            =   120
      TabIndex        =   46
      Top             =   7920
      Width           =   975
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "TILE HEIGHT:"
      Height          =   255
      Index           =   9
      Left            =   6360
      TabIndex        =   18
      Top             =   1560
      Width           =   1455
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "TILE WIDTH:"
      Height          =   255
      Index           =   8
      Left            =   6360
      TabIndex        =   17
      Top             =   1200
      Width           =   1455
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "FRAMES:"
      Height          =   255
      Index           =   7
      Left            =   6000
      TabIndex        =   16
      Top             =   720
      Width           =   1095
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "DEST Y:"
      Height          =   255
      Index           =   6
      Left            =   4800
      TabIndex        =   13
      Top             =   1560
      Width           =   855
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "DEST X:"
      Height          =   255
      Index           =   5
      Left            =   4800
      TabIndex        =   12
      Top             =   1200
      Width           =   855
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "SRC X:"
      Height          =   255
      Index           =   4
      Left            =   2880
      TabIndex        =   11
      Top             =   1200
      Width           =   1095
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "SRC Y:"
      Height          =   255
      Index           =   3
      Left            =   2880
      TabIndex        =   10
      Top             =   1560
      Width           =   1095
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "SPEED:"
      Height          =   255
      Index           =   2
      Left            =   6000
      TabIndex        =   9
      Top             =   360
      Width           =   1095
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "Gráfico:"
      Height          =   255
      Index           =   1
      Left            =   2880
      TabIndex        =   8
      Top             =   720
      Width           =   1095
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "Número:"
      Height          =   255
      Index           =   0
      Left            =   2880
      TabIndex        =   7
      Top             =   360
      Width           =   1095
   End
   Begin VB.Label Label2 
      Caption         =   "Vista Prévia:"
      Height          =   375
      Left            =   2880
      TabIndex        =   2
      Top             =   2160
      Width           =   5775
   End
   Begin VB.Label Label1 
      Caption         =   "Listado de GRH's:"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2655
   End
   Begin VB.Menu mnuArchivo 
      Caption         =   "Archivo"
      Begin VB.Menu mnuPAO_EXPORT 
         Caption         =   "Exportar a PAO"
      End
      Begin VB.Menu mnuAOEXPORT 
         Caption         =   "Exportar a AO"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Salir"
      End
   End
   Begin VB.Menu mnuED 
      Caption         =   "Edición"
      Begin VB.Menu mnuEdit 
         Caption         =   "Agregar GRH"
      End
      Begin VB.Menu mnuFX 
         Caption         =   "Agregar FX"
      End
      Begin VB.Menu mnuNewCas 
         Caption         =   "Agregar Casco/Cabeza"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Index As Long
Public MouseX As Single
Public MouseY As Single

Private Sub Command1_Click(Index As Integer)
    
    Select Case Index
        Case 0
            If frmMain.Index = 0 Then Exit Sub
        
            With GrhData(Text1(0).Text)
                .FileNum = Text1(1).Text
                .sX = Text1(2).Text
                .sY = Text1(3).Text
                .pixelWidth = Text1(4).Text
                .pixelHeight = Text1(5).Text
                .TileWidth = Text1(6).Text
                .TileHeight = Text1(7).Text
            End With
            
            MsgBox "ACTUALIZADO"
        Case 1
            If GrhData(Text2(2).Text).NumFrames = 1 Then Exit Sub
            With FxData(2)
                .Animacion = Text2(2).Text
                .OffsetX = Text2(0).Text
                .OffsetY = Text2(1).Text
            End With
            
            MsgBox "ACTUALIZADO"
        Case 2
            Dim tmp As Integer
            tmp = ReadField(1, HHH_LIST.Text, Asc(","))
            If Option1(0).value = True Then
                HeadData(tmp).Head(1).GrhIndex = Text4(3).Text
                HeadData(tmp).Head(2).GrhIndex = Text4(1).Text
                HeadData(tmp).Head(3).GrhIndex = Text4(2).Text
                HeadData(tmp).Head(4).GrhIndex = Text4(0).Text
            Else
                CascoAnimData(tmp).Head(1).GrhIndex = Text4(3).Text
                CascoAnimData(tmp).Head(2).GrhIndex = Text4(1).Text
                CascoAnimData(tmp).Head(3).GrhIndex = Text4(2).Text
                CascoAnimData(tmp).Head(4).GrhIndex = Text4(0).Text
            End If
            
            MsgBox "ACTUALIZADO"
    End Select
    
End Sub

Private Sub Command2_Click()
    If frmMain.Index = 0 Then Exit Sub
    Dim tmp_TEXT As String, i As Long
    
    For i = 1 To GrhData(Index).NumFrames
        tmp_TEXT = tmp_TEXT & GrhData(Index).Frames(i) & " "
    Next i


    
        frmEditANIM.Text1(1).Text = tmp_TEXT
        
        frmEditANIM.Text1(8).Text = GrhData(Index).Speed
        frmEditANIM.Text1(9).Text = GrhData(Index).NumFrames
        
    Index = 0
    frmEditANIM.Show
End Sub

Private Sub Command3_Click()
Dim i As Long
    For i = 1 To GRH_LIST.ListCount
        If ReadField(1, GRH_LIST.List(i), Asc(",")) = Text3.Text Then
            GRH_LIST.Selected(i) = True
            Exit For
        End If
    Next i


End Sub

Private Sub Command4_Click()
Dim i As Long
Text5.Text = ""
    For i = 1 To UBound(GrhData())
        'If GrhData(i).FileNum = 25181 Then _
            Text5.Text = Text5.Text & i & vbNewLine
            
        If GrhData(i).FileNum = Val(Text6.Text) Then _
            Text5.Text = Text5.Text & i & vbNewLine
    Next i
End Sub

Private Sub Command5_Click()

Dim LoopG As Long
    ReDim GRH_COLORS(1 To UBound(GrhData)) As Long
    
    For LoopG = 1 To UBound(GrhData)
        With GrhData(LoopG)
            frmMain.Index = LoopG
            
            If .NumFrames = 1 Then
                If .FileNum <> 0 Then
                    
                 '   RENDER_GRH_ONCE frmMain.Index
                  '  Debug.Print GetPixel(Render.hDC, 16, 16)
                    
                    If FileExist(BMP_DIRE & .FileNum & ".bmp", vbNormal) Then
                        frmMain.Picture3.Picture = LoadPicture(BMP_DIRE & .FileNum & ".bmp")
                        DoEvents
                        
                        Render.Cls
                        BitBlt Render.hDC, 0, 0, .pixelWidth, .pixelHeight, Picture3.hDC, .sX, .sY, vbSrcCopy
                        GRH_COLORS(LoopG) = GetPixel(Render.hDC, .pixelWidth / 2, .pixelHeight / 2)
                        
                        Render.BackColor = GRH_COLORS(LoopG)
                        
                    End If
                End If
            Else
               GRH_COLORS(LoopG) = 0
            End If
        End With
        

    
        frmMain.Index = 0
    Next LoopG
    

Dim File
File = FreeFile
    Open App.Path & "\COLORS.BIN" For Binary Access Write As File
        Put File, , GRH_COLORS()
    Close #File
    
End Sub

Private Sub Form_Load()
MouseX = -1
MouseY = -1

frmMain.Render.MousePointer = vbNoDrop
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    MouseX = -1
    MouseY = -1
End Sub

Private Sub FXS_LIST_Click()
    Index = ReadField(1, FXS_LIST.Text, Asc(","))

    
    With FxData(Index)
        Text2(0).Text = .OffsetX
        Text2(1).Text = .OffsetY
        Text2(2).Text = .Animacion
    End With

    Index = Text2(2).Text
    
    RENDER_ANIM Index
End Sub

Private Sub GRH_LIST_Click()
    Index = ReadField(1, GRH_LIST.Text, Asc(","))

    With GrhData(Index)
    Text1(0).Text = Index
    Text1(1).Text = .FileNum
    Text1(2).Text = .sX
    Text1(3).Text = .sY
    Text1(4).Text = .pixelWidth
    Text1(5).Text = .pixelHeight
    Text1(6).Text = .TileWidth
    Text1(7).Text = .TileHeight
    Text1(8).Text = .Speed
    Text1(9).Text = .NumFrames
    
    If .NumFrames <> 1 Then
        Text1(8).Enabled = True
        Text1(9).Enabled = True
        Command2.Enabled = True
        RENDER_ANIM Index
    Else
        Text1(8).Enabled = False
        Text1(9).Enabled = False
        Command2.Enabled = False
        RENDER_GRH Index
    End If
    
    End With
    
    
End Sub

Private Sub mnuFXS_Click()
frmFXS.Show
End Sub


Private Sub HHH_LIST_Click()
    Index = ReadField(1, HHH_LIST.Text, Asc(","))
    
    If Option1(0).value = True Then
        Text4(0).Text = HeadData(Index).Head(4).GrhIndex
        Text4(1).Text = HeadData(Index).Head(2).GrhIndex
        Text4(2).Text = HeadData(Index).Head(3).GrhIndex
        Text4(3).Text = HeadData(Index).Head(1).GrhIndex
        Index = HeadData(Index).Head(3).GrhIndex
    Else
        Text4(0).Text = CascoAnimData(Index).Head(4).GrhIndex
        Text4(1).Text = CascoAnimData(Index).Head(2).GrhIndex
        Text4(2).Text = CascoAnimData(Index).Head(3).GrhIndex
        Text4(3).Text = CascoAnimData(Index).Head(1).GrhIndex
        Index = CascoAnimData(Index).Head(3).GrhIndex
    End If

    RENDER_GRH Index
End Sub

Private Sub mnuEdit_Click()
    Index = 0
    frmNew.Show
End Sub

Private Sub mnuExit_Click()
    If MsgBox("¿Desea salir sin guardar?", vbYesNo) = vbYes Then
        frmMain.Index = 0
        DELETE_BUFFER
        DoEvents
        
        End
    End If
End Sub

Private Sub mnuFX_Click()
    Index = 0
    frmNewFX.Show
End Sub

Private Sub mnuNewCas_Click()
Index = 0
frmNewCas.Show
End Sub

Private Sub Option1_Click(Index As Integer)
Dim i As Integer
HHH_LIST.Clear
    Select Case Index
        Case 0
            For i = 1 To NUM_HEA
                HHH_LIST.AddItem i
            Next i
        Case 1
            For i = 1 To NUM_CAS
                HHH_LIST.AddItem i
            Next i
    End Select
End Sub

Private Sub Render_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    MouseX = X
    MouseY = Y
End Sub

Private Sub Text1_Change(Index As Integer)
    Select Case Index
    
        Case 9
            Text1(8).Text = (Val(Text1(9).Text) * 1000) / 18
             Text1(8).Enabled = Text1(9).Text <> 1
    End Select
End Sub

