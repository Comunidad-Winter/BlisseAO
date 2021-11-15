VERSION 5.00
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.ocx"
Begin VB.Form frmBGUpdate 
   BorderStyle     =   0  'None
   Caption         =   "BGUpdater"
   ClientHeight    =   1500
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmBGUpdate.frx":0000
   ScaleHeight     =   100
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   200
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   120
      Top             =   600
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   120
      Top             =   120
   End
   Begin InetCtlsObjects.Inet Inet1 
      Left            =   2280
      Top             =   120
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin CLBLISSEAO.BGAOAniGif BGAOAniGif1 
      Height          =   480
      Left            =   180
      TabIndex        =   0
      Top             =   285
      Visible         =   0   'False
      Width           =   480
      _ExtentX        =   847
      _ExtentY        =   847
   End
   Begin VB.Image Image2 
      Height          =   480
      Left            =   180
      Picture         =   "frmBGUpdate.frx":4182
      Top             =   285
      Width           =   480
   End
   Begin VB.Image Image1 
      Height          =   180
      Left            =   2040
      Top             =   1200
      Width           =   765
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H80000002&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   300
      Left            =   195
      Top             =   810
      Width           =   2535
   End
   Begin VB.Label Label2 
      Height          =   375
      Left            =   720
      TabIndex        =   1
      Top             =   360
      Width           =   2055
   End
End
Attribute VB_Name = "frmBGUpdate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public UpdatesInt As Integer
Public UpdatesCli As Integer
Public UpdatesToInstall As Integer
Private clsFormulario As clsFormMovementManager

Public Sub InitBusqueda()
    Timer2.Enabled = False
    Image2.Visible = False
    BGAOAniGif1.Visible = True
    BGAOAniGif1.ContinueGif
    
    Label2.Caption = "Buscando actualizaciones."
    Me.Visible = True
    UpdatesInt = Inet1.OpenURL(Client_Web & "/num.txt")
    UpdatesCli = GetClUpdates
    DoEvents
    
    UpdatesToInstall = UpdatesInt - UpdatesCli
    
    If UpdatesToInstall = 0 Then
        Label2.Caption = "Cliente Actualizado Correctamente."
        Shape1.Width = 100
        
        If FileExist(App.path & "\Game\Blisse-AO.exe", vbNormal) = True Or FileExist(App.path & "\Game\Blisse-AO_Light.exe", vbNormal) = True Then
            Call ShellExecute(0, "Open", App.path & "\BGAO.exe", App.EXEName & ".exe /bgupdate", App.path, SW_SHOWNORMAL)
            End
        End If
        DoEvents
        
        Unload Me
    Else
        Label2.Caption = "Se Instalarán " & UpdatesToInstall & " actualizaciones."
        Call UpdateNow
        DoEvents
        
        If FileExist(App.path & "\Game\Blisse-AO.exe", vbNormal) = True Or FileExist(App.path & "\Game\Blisse-AO_Light.exe", vbNormal) = True Then
            Call ShellExecute(0, "Open", App.path & "\BGAO.exe", App.EXEName & ".exe /bgupdate", App.path, SW_SHOWNORMAL)
            End
        End If

        DoEvents
        
        Unload Me
            
    End If
    
End Sub

Private Sub UpdateNow()
    Dim i As Integer
        BGAOAniGif1.GifPath = App.path & "\Game\Resources\Interface\loader.gif" 'Dir_Resources & "Interface\loader.gif"
        BGAOAniGif1.StartGif
        BGAOAniGif1.StopGif
    
        For i = UpdatesCli + 1 To UpdatesInt
            Label2.Caption = "Descargando actualización n" & i & "."
        
                Inet1.AccessType = icUseDefault
                    Dim b() As Byte
                
                    b() = Inet1.OpenURL("http://www.blisse-ao.com.ar/updater/parche" & i & ".zip", icByteArray)
                
                    Open App.path & "\parche" & i & ".zip" For Binary Access _
                        Write As #1
                            Put #1, , b()
                    Close #1

                DoEvents
            Label2.Caption = "Instalando actualización n" & i & "."
            UnZip App.path & "\parche" & i & ".zip", App.path & "\"
                DoEvents
            Call Kill(App.path & "\parche" & i & ".zip")
                
            
            Shape1.Width = (((i / 100) / (UpdatesInt / 100)) * 169)
        Next i
        
    DoEvents
    
    Label2.Caption = "Cliente Actualizado Correctamente."
    Shape1.Width = 100
    PutClUpdates
End Sub

Private Sub Form_Load()
    '   Handles Form movement (drag and drop).
    Set clsFormulario = New clsFormMovementManager
    clsFormulario.Initialize Me
    
    Label2.ForeColor = RGB(136, 134, 123)
    Shape1.BackColor = RGB(136, 134, 123)
    Shape1.Width = 1
    
    UpdatesInt = 0
    UpdatesCli = 0
    BGAOAniGif1.GifPath = App.path & "\Game\Resources\Interface\loader.gif" 'Dir_Resources & "Interface\loader.gif"
    BGAOAniGif1.StartGif
    BGAOAniGif1.StopGif
    Timer2.Enabled = True
    DoEvents
    Exit Sub
End Sub

Private Sub Image1_Click()
    If MsgBox("El juego se está actualizando, Si sale ahora podrá dañar su cliente. ¿Está seguro que desea salir?", vbYesNo, "Blisse-Updater") = vbYes Then
        End
    End If
End Sub

Private Sub Timer1_Timer()
    If Label2.Caption = "Cliente Actualizado Correctamente." Then
        BGAOAniGif1.StopGif
        BGAOAniGif1.Visible = False
        Image2.Visible = True
        Shape1.Width = 169
    End If
End Sub

Private Function GetClUpdates() As Byte
    Dim N As Integer
        N = FreeFile
            Open App.path & "\Game\System" For Binary As #N
                Get #N, , GetClUpdates
            Close #N
End Function

Private Function PutClUpdates() As Boolean
Dim File
File = FreeFile
    Open App.path & "\Game\System" For Binary Access Write As File
        Put File, , UpdatesInt
    Close #File
End Function

Private Sub Timer2_Timer()
    Call InitBusqueda
    Timer2.Enabled = False
End Sub
