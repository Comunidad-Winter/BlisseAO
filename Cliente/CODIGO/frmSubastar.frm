VERSION 5.00
Begin VB.Form frmSubastar 
   BorderStyle     =   0  'None
   ClientHeight    =   5070
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6180
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   338
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   412
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      BackColor       =   &H80000007&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   180
      Index           =   1
      Left            =   3540
      TabIndex        =   7
      Text            =   "1"
      Top             =   1635
      Width           =   2175
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H80000007&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   180
      Index           =   0
      Left            =   3540
      TabIndex        =   6
      Text            =   "1000"
      Top             =   1005
      Width           =   2175
   End
   Begin CLBLISSEAO.BGAOButton Command2 
      Height          =   465
      Left            =   3480
      TabIndex        =   4
      Top             =   4080
      Width           =   1050
      _ExtentX        =   1852
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
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      FillColor       =   &H000000FF&
      ForeColor       =   &H80000008&
      Height          =   480
      Left            =   3570
      ScaleHeight     =   450
      ScaleWidth      =   450
      TabIndex        =   3
      Top             =   2385
      Width           =   480
   End
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000007&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   3735
      Left            =   435
      TabIndex        =   0
      Top             =   810
      Width           =   2730
   End
   Begin CLBLISSEAO.BGAOButton Command1 
      Height          =   465
      Left            =   4800
      TabIndex        =   5
      Top             =   4080
      Width           =   1050
      _ExtentX        =   1852
      _ExtentY        =   820
      Caption         =   "Subastar"
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
   Begin VB.Label lblName 
      BackStyle       =   0  'Transparent
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
      Left            =   4125
      TabIndex        =   2
      Top             =   2445
      Width           =   1665
   End
   Begin VB.Label lblInfo 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Slot vacio, No hay �tem para subastar."
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
      Height          =   900
      Left            =   3540
      TabIndex        =   1
      Top             =   3000
      Width           =   2250
   End
End
Attribute VB_Name = "frmSubastar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************
'Author: Ezequiel Ju�rez (Standelf)
'Last Modification: 25/05/10
'Blisse-AO | Formulario de Subastas 0.13.x
'***************************************************

Private Sub Command1_Click()
    '   Subastas 1 item o mas o te rajas!
    If Val(Text1(1).Text) <= 0 Then
        MsgBox "Debes subastar una cantidad mayor a 0 de �tems", vbOKOnly
        Exit Sub
    End If
    
    '   Que valga mas de 0 �� no lo regales ��
    If Val(Text1(0).Text) <= 100 Then
        MsgBox "Debes poner un valor mayor a 100 para poder subastar", vbOKOnly
        Exit Sub
    End If
    
    
    '   Si no tiene nada no se puede subastar
    If List1.Text = "Nada" Then
        MsgBox "Debes seleccionar un �tem para poder iniciar una subasta", vbOKOnly
    Else
        '   Si no tiene la cantidad de items que quiere subastar lo rajamos ;)
        If Inventario.Amount(List1.ListIndex + 1) < Val(Text1(1).Text) Then
            MsgBox "No tienes la cantidad de items que intentas subastar", vbOKOnly
        Else
            '   Enviamos los datos para comenzar la subasta
            Call WriteIniciarSubasta(Val(List1.ListIndex + 1), Text1(1).Text, Text1(0).Text)
        End If
    End If
End Sub

Private Sub Command2_Click()
    Unload Me
End Sub

Private Sub Form_Load()
Dim i As Byte

    Me.Picture = General_Set_GUI("VentanaSubasta")
    Command1.Init sb_Normal
    Command2.Init sb_Normal
    

    '   Cargamos la lista, si no tiene nada agregamos un "Nada" para que se distinga.
    For i = 1 To MAX_INVENTORY_SLOTS
        If Inventario.ItemName(i) = "" Then
            List1.AddItem "Nada"
        Else
            List1.AddItem Inventario.ItemName(i) & "(" & Inventario.Amount(i) & ")"
        End If
    Next i
    
        lblName.Caption = "Nada"
        lblInfo.Caption = "Slot vacio, No hay �tem para subastar."
        lblName.ForeColor = &HC0&
        lblInfo.ForeColor = &HC0&
        Command1.Visible = False
        Picture1.Cls
        

       
End Sub

Private Sub Form_Unload(Cancel As Integer)
    List1.Clear
        lblName.Caption = "Nada"
        lblInfo.Caption = "Slot vacio, No hay �tem para subastar."
        lblName.ForeColor = &HC0&
        lblInfo.ForeColor = &HC0&
        Command1.Visible = False
        Picture1.Cls
End Sub

Private Sub list1_Click()
    If List1.Text = "Nada" Then
        lblName.Caption = "Nada"
        lblInfo.Caption = "Slot vacio, No hay �tem para subastar."
        lblName.ForeColor = &HC0&
        lblInfo.ForeColor = &HC0&
        Command1.Visible = False
        Picture1.Cls
    Else
        lblName.Caption = Inventario.ItemName(Val(List1.ListIndex + 1))
        lblInfo.Caption = "Valor Inicial: " & FormatNumber(Inventario.Valor(Val(List1.ListIndex + 1)), 0) & vbCrLf & _
                                "Cantidad Disponible: " & Inventario.Amount(Val(List1.ListIndex + 1)) & vbCrLf & _
                                "MinHit/MaxHit: " & Inventario.MinHit(Val(List1.ListIndex + 1)) & "/" & Inventario.MaxHit(Val(List1.ListIndex + 1)) & vbCrLf & _
                                "MinDef/MaxDef: " & Inventario.MinDef(Val(List1.ListIndex + 1)) & "/" & Inventario.MaxDef(Val(List1.ListIndex + 1))
        
        lblName.ForeColor = &H8000000C
        lblInfo.ForeColor = &H8000000C
        Command1.Visible = True
        
        Call RenderItem(Picture1, Inventario.GrhIndex(Val(List1.ListIndex + 1)))
    End If
End Sub
