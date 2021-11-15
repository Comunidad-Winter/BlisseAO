VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   6915
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   11985
   LinkTopic       =   "Form1"
   ScaleHeight     =   6915
   ScaleWidth      =   11985
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000001&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   6000
      Left            =   1440
      ScaleHeight     =   400
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   400
      TabIndex        =   1
      Top             =   600
      Width           =   6000
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   735
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   975
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim X As Integer, xx As Integer
Dim Y As Integer, yy As Integer
Dim RECT As RECT
    RECT.bottom = 400
    RECT.Right = 400
    RECT.Left = 0
    RECT.Top = 0
        Engine_BeginScene 0
        
        
    For X = 1 To 100
    For Y = 1 To 100

        
        If MapData(X, Y).Graphic(1).GrhIndex <> 0 Then
            Call DDrawGrhtoSurfaceScale(MapData(X, Y).Graphic(1), X * 4, Y * 4, 1, 1, X, Y)
        End If
        


    Next Y
    Next X
        For X = 1 To 100
    For Y = 1 To 100
        'If MapData(x, y).Graphic(2).GrhIndex <> 0 Then
            'DDrawGrhtoSurface MapData(x, y).Graphic(2), x, y, 0, 1, x, y
        'End If

        
        If MapData(X, Y).Graphic(3).GrhIndex <> 0 Then
            DDrawTransGrhtoSurfaceScale MapData(X, Y).Graphic(3), X * 4, Y * 4, 1, MapData(X, Y).Engine_Light(), 0, X, Y, False, 0, False
        End If
     Next Y
    Next X
           
        Engine_Render_Fill_Box UserPos.X * 4, UserPos.Y * 4, 4, 4, D3DColorARGB(150, 255, 0, 0)
        Engine_EndScene RECT, Picture1.hWnd
End Sub

Private Sub Form_Load()

End Sub
