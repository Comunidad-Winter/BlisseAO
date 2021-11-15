VERSION 5.00
Begin VB.Form FrmCuenta 
   AutoRedraw      =   -1  'True
   BackColor       =   &H8000000C&
   BorderStyle     =   0  'None
   ClientHeight    =   9000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12000
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
   ScaleHeight     =   600
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   800
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Image imgPremium 
      Height          =   240
      Left            =   10920
      Picture         =   "FrmCuenta.frx":0000
      Top             =   8640
      Visible         =   0   'False
      Width           =   960
   End
   Begin VB.Image Image1 
      Height          =   255
      Left            =   180
      Top             =   8700
      Width           =   2055
   End
   Begin VB.Label LPJ 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   7
      Left            =   3000
      TabIndex        =   11
      Top             =   5970
      Width           =   6000
   End
   Begin VB.Label LPJ 
      BackStyle       =   0  'Transparent
      Caption         =   "ASDASD"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   6
      Left            =   3000
      TabIndex        =   10
      Top             =   5520
      Width           =   6000
   End
   Begin VB.Label LPJ 
      BackStyle       =   0  'Transparent
      Caption         =   "ASDASD"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   5
      Left            =   3000
      TabIndex        =   9
      Top             =   5070
      Width           =   6000
   End
   Begin VB.Label LPJ 
      BackStyle       =   0  'Transparent
      Caption         =   "ASDASD"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   4
      Left            =   3000
      TabIndex        =   8
      Top             =   4620
      Width           =   6000
   End
   Begin VB.Label LPJ 
      BackStyle       =   0  'Transparent
      Caption         =   "ASDASD"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   3
      Left            =   3000
      TabIndex        =   7
      Top             =   4200
      Width           =   6000
   End
   Begin VB.Label LPJ 
      BackStyle       =   0  'Transparent
      Caption         =   "ASDASD"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   2
      Left            =   3000
      TabIndex        =   6
      Top             =   3720
      Width           =   6000
   End
   Begin VB.Label LPJ 
      BackStyle       =   0  'Transparent
      Caption         =   "ASDASD"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   1
      Left            =   3000
      TabIndex        =   5
      Top             =   3270
      Width           =   6000
   End
   Begin VB.Label LPJ 
      BackStyle       =   0  'Transparent
      Caption         =   "ASDASD"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   0
      Left            =   3000
      TabIndex        =   4
      Top             =   2820
      Width           =   6000
   End
   Begin VB.Label boton 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Crear personaje"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   0
      Left            =   4140
      MousePointer    =   10  'Up Arrow
      TabIndex        =   3
      Top             =   7035
      Width           =   1815
   End
   Begin VB.Label boton 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Volver"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   3
      Left            =   8280
      MousePointer    =   10  'Up Arrow
      TabIndex        =   2
      Top             =   7035
      Width           =   1935
   End
   Begin VB.Label boton 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Borrar personaje"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   2
      Left            =   6240
      MousePointer    =   10  'Up Arrow
      TabIndex        =   1
      Top             =   7035
      Width           =   1815
   End
   Begin VB.Label boton 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Conectar"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   1
      Left            =   2040
      MousePointer    =   10  'Up Arrow
      TabIndex        =   0
      Top             =   7035
      Width           =   1815
   End
End
Attribute VB_Name = "FrmCuenta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim PJSeleccionado As Byte
Private clsFormulario As clsFormMovementManager

Private Sub Form_Load()
    If Not Settings.Ventana Then
        '   Handles Form movement (drag and drop).
        Set clsFormulario = New clsFormMovementManager
        clsFormulario.Initialize Me, , 120
    End If
    
Me.Caption = "  Cuenta de " & Cuenta.name
Me.Picture = General_Set_GUI("ConectarCuenta")
End Sub

Private Sub boton_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim i As Byte
For i = 0 To 3
    boton(i).ForeColor = vbWhite
Next i
    
boton(Index).ForeColor = RGB(102, 255, 204)
End Sub

Private Sub Image1_Click()
    Call ShellExecute(0, "Open", Client_Web, "", App.Path, SW_SHOWNORMAL)
End Sub

Private Sub LPJ_Click(Index As Integer)
Call Audio.PlayWave(SND_CLICK)
PJSeleccionado = Index

Dim i As Byte
For i = 0 To 7
    If Cuenta.pjs(i + 1).promedio < 0 Then
        'Criminal
        LPJ(i).ForeColor = &H8080FF
    Else
        'Ciudadano
        LPJ(i).ForeColor = &HFFC0C0
    End If
Next i

If LPJ(PJSeleccionado).ForeColor <> vbWhite Then
   LPJ(PJSeleccionado).ForeColor = vbWhite
End If
End Sub

Private Sub boton_Click(Index As Integer)
Call Audio.PlayWave(SND_CLICK)

If frmMain.Winsock1.State <> sckClosed Then
    frmMain.Winsock1.Close
    DoEvents
End If

Select Case Index
    Case 0 'CREAR PJ
        If Cuenta.Premium = True Then
            If Cuenta.CantPJ = 8 Then
                MsgBox "No tienes más espacio para continuar creando personajes."
                Exit Sub
            End If
        Else
            If Cuenta.CantPJ = 5 Then
                MsgBox "No tienes más espacio para continuar creando personajes." + vbCrLf + "¿Quieres tener más espacio? Haste PREMIUM en este instante y obten beneficios ¡exclusivos!"
                Exit Sub
            End If
        End If
        
        EstadoLogin = Dados
        
        frmMain.Winsock1.Connect Server_IP, Server_Port
        
        Unload Me
        Exit Sub
        
    Case 1  'CONECTAR PJ
        If Cuenta.pjs(PJSeleccionado + 1).NamePJ = "" Then Exit Sub
        
        UserName = Cuenta.pjs(PJSeleccionado + 1).NamePJ
        EstadoLogin = Normal
        
        frmMain.Winsock1.Connect Server_IP, Server_Port
        
        Unload Me
        Exit Sub
    
    Case 2 'BORRAR PJ
        If Cuenta.pjs(PJSeleccionado + 1).NamePJ = "" Then Exit Sub
        If MsgBox("Al borrar un personaje de su cuenta perderá todo lo que hay en él." & vbCrLf & "¿Está totalmente seguro que decea eliminar el mismo?", vbInformation + vbYesNo, "Eliminar Personaje de la cuenta.") = vbYes Then
            #If SeguridadBlisse = 1 Then
                If RevisarCodigo = False Then Exit Sub
            #End If
            If Cuenta.pjs(PJSeleccionado + 1).NamePJ = "" Then Exit Sub
            
            IndexSelectedUSer = PJSeleccionado + 1
            EstadoLogin = BorrarPJ
            
            frmMain.Winsock1.Connect Server_IP, Server_Port
            
            LPJ(PJSeleccionado).Caption = ""
            Cuenta.pjs(PJSeleccionado + 1).NamePJ = ""
            Exit Sub
        Else
            Exit Sub
        End If
    Case 3
        frmConnect.Show
        AlphaPres = 255
        set_GUI_Efect
        CurMapAmbient.Fog = 100
        Unload Me
End Select
End Sub

Private Sub LPJ_DblClick(Index As Integer)
If frmMain.Winsock1.State <> sckClosed Then
    frmMain.Winsock1.Close
    DoEvents
End If

        If Cuenta.pjs(PJSeleccionado + 1).NamePJ = "" Then Exit Sub
        
        UserName = Cuenta.pjs(PJSeleccionado + 1).NamePJ
        EstadoLogin = Normal
        
        frmMain.Winsock1.Connect Server_IP, Server_Port
        
        Unload Me
        Exit Sub
End Sub
