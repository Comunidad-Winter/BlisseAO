VERSION 5.00
Begin VB.Form frmConnect 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Argoth, mod de Argentum Online"
   ClientHeight    =   9000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12000
   FillColor       =   &H00000040&
   Icon            =   "frmConnect.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   600
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   800
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.Image imgServArgentina 
      Height          =   795
      Left            =   360
      MousePointer    =   99  'Custom
      Top             =   9240
      Visible         =   0   'False
      Width           =   2595
   End
End
Attribute VB_Name = "frmConnect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private clsFormulario As clsFormMovementManager

Private MousePosX As Integer
Private MousePosY As Integer

Public tmpName As String
Public tmpPass As String
Public tmpPassF As String
Public Focus As Byte

Private Sub Form_Activate()
    If Settings.UltimaCuenta <> vbNullString And Settings.Recordar Then
        tmpName = Settings.UltimaCuenta
    Else
        tmpName = ""
    End If
    Exit Sub
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 27 Then
        prgRun = False
    End If
End Sub

Private Sub Form_Load()
    
    EngineRun = False

End Sub

Private Sub Form_Click()
    ' GUI
    Focus = 0
    
    ' #### GAME4FUN
    If Engine_Collision_Rect(697, 556, 97, 24, MousePosX, MousePosY, 1, 1) Then
        Call Audio.PlayWave(SND_CLICK)
        Call ShellExecute(0, "Open", Game4Fun, "", App.Path, SW_SHOWNORMAL)
    End If
    

    ' #### PW
    If Engine_Collision_Rect(320, 325, 160, 20, MousePosX, MousePosY, 1, 1) Then
        Call Audio.PlayWave(SND_CLICK)
        Focus = 2
    End If
    ' #### CUENTA
    If Engine_Collision_Rect(320, 273, 160, 20, MousePosX, MousePosY, 1, 1) Then
        Call Audio.PlayWave(SND_CLICK)
        Focus = 1
    End If
    
    ' #### Conectarse
    If Engine_Collision_Rect(349, 371, 102, 15, MousePosX, MousePosY, 1, 1) Then
        Call Audio.PlayWave(SND_CLICK)
        
        If frmMain.Winsock1.State <> sckClosed Then
            frmMain.Winsock1.Close
            DoEvents
        End If
        
        Cuenta.name = tmpName
        Cuenta.Pass = tmpPass
    
        If General_Check_AccountData(False, False) = True Then
            EstadoLogin = LoginCuenta
    
            frmMain.Winsock1.Connect Server_IP, Server_Port
        End If
    End If
    
    ' #### Recordar
    If Engine_Collision_Rect(340, 395, 18, 17, MousePosX, MousePosY, 1, 1) Then
        Call Audio.PlayWave(SND_CLICK)
        Settings.Recordar = Not Settings.Recordar
    End If
    
    ' #### Configuración
    If Engine_Collision_Rect(26, 404, 115, 23, MousePosX, MousePosY, 1, 1) Then
        Call Audio.PlayWave(SND_CLICK)
        frmOpciones.Show , frmConnect
    End If
    
    ' #### Crear Cuenta
    If Engine_Collision_Rect(26, 376, 108, 23, MousePosX, MousePosY, 1, 1) Then
        Call Audio.PlayWave(SND_CLICK)
        FrmNewCuenta.Show , frmConnect
    End If
    
    ' #### Salir
    If Engine_Collision_Rect(26, 482, 108, 23, MousePosX, MousePosY, 1, 1) Then
        Call Audio.PlayWave(SND_CLICK)
        prgRun = False
    End If
End Sub

Private Sub Form_DblClick()
    Focus = 0
    
    ' #### PW
    If Engine_Collision_Rect(320, 325, 160, 20, MousePosX, MousePosY, 1, 1) Then
        Focus = 2
        tmpPass = ""
        tmpPassF = ""
    End If
    
    ' #### CUENTA
    If Engine_Collision_Rect(320, 273, 160, 20, MousePosX, MousePosY, 1, 1) Then
        Focus = 1
        tmpName = ""
    End If
    
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If Focus = 0 Then Exit Sub
Debug.Print KeyAscii

    If Focus = 1 Then
    
        If KeyAscii = 9 Then
            Focus = 2
            Exit Sub
        End If
        
        If Len(tmpName) = 21 And KeyAscii <> 8 Then Exit Sub
        
        If KeyAscii = 8 Then
            If Len(tmpName) <> 0 Then _
            tmpName = mid(tmpName, 1, Len(tmpName) - 1)
        Else
            tmpName = tmpName & Chr(KeyAscii)
        End If
    End If
    
    If Focus = 2 And KeyAscii = 13 Then
        If frmMain.Winsock1.State <> sckClosed Then
            frmMain.Winsock1.Close
            DoEvents
        End If
        
        Cuenta.name = tmpName
        Cuenta.Pass = tmpPass
    
        If General_Check_AccountData(False, False) = True Then
            EstadoLogin = LoginCuenta
    
            frmMain.Winsock1.Connect Server_IP, Server_Port
        End If
        
        Exit Sub
    End If
    If Focus = 2 Then
        If Len(tmpPass) = 21 And KeyAscii <> 8 Then Exit Sub
        
        If KeyAscii = 8 Then
            If Len(tmpPass) <> 0 Then
            tmpPass = mid(tmpPass, 1, Len(tmpPass) - 1)
            tmpPassF = mid(tmpPassF, 1, Len(tmpPassF) - 1)
            End If
        Else
            tmpPass = tmpPass & Chr(KeyAscii)
            tmpPassF = tmpPassF & "*"
        End If
    End If
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    MousePosX = X
    MousePosY = Y
End Sub
