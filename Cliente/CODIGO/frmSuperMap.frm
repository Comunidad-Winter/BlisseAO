VERSION 5.00
Begin VB.Form frmSuperMap 
   Caption         =   "Form1"
   ClientHeight    =   6420
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7350
   LinkTopic       =   "Form1"
   ScaleHeight     =   428
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   490
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      Height          =   4500
      Left            =   120
      ScaleHeight     =   296
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   336
      TabIndex        =   0
      Top             =   120
      Width           =   5100
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Command2"
      Height          =   735
      Left            =   5280
      TabIndex        =   3
      Top             =   2040
      Width           =   1695
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   735
      Left            =   5280
      TabIndex        =   2
      Top             =   1200
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   855
      Left            =   5280
      TabIndex        =   1
      Top             =   120
      Width           =   1695
   End
End
Attribute VB_Name = "frmSuperMap"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Dim map_x As Long, map_y As Long
    
    Me.Picture1.Cls

    For map_y = YMinMapSize To YMaxMapSize
        For map_x = XMinMapSize To XMaxMapSize
            If MapData(map_x, map_y).Graphic(1).GrhIndex > 0 Then
                Dim X1 As Integer
                Dim Y1 As Integer
                If GrhData(MapData(map_x, map_y).Graphic(1).GrhIndex).MiniMap_color <> 0 Then
                    SetPixel Me.Picture1.hdc, (map_x - 1) * 2, (map_y - 1) * 2, GrhData(MapData(map_x, map_y).Graphic(1).GrhIndex).MiniMap_color
                    SetPixel Me.Picture1.hdc, (map_x - 1) * 2, (map_y) * 2 + 1, GrhData(MapData(map_x, map_y).Graphic(1).GrhIndex).MiniMap_color
                    
                    SetPixel Me.Picture1.hdc, (map_x - 1) * 2 + 1, (map_y - 1) * 2, GrhData(MapData(map_x, map_y).Graphic(1).GrhIndex).MiniMap_color
                    SetPixel Me.Picture1.hdc, (map_x - 1) * 2 + 1, (map_y - 1) * 2 + 1, GrhData(MapData(map_x, map_y).Graphic(1).GrhIndex).MiniMap_color
                End If
            End If
        Next map_x
    Next map_y
    
    Me.Picture1.Refresh
End Sub

Private Sub Command2_Click()
    Dim destRect As RECT
    With destRect
        .Top = 0
        .Left = 0
        .bottom = 200
        .Right = 200
    End With
    Dim map_x As Long, map_y As Long
    
    DirectDevice.Clear 0, ByVal 0, D3DCLEAR_TARGET, 0, 0#, 0
    DirectDevice.BeginScene
        For map_y = YMinMapSize To YMaxMapSize
            For map_x = XMinMapSize To XMaxMapSize
                If MapData(map_x, map_y).Graphic(1).GrhIndex <> 0 Then
                    Call DDrawTransGrhtoSurface2(MapData(map_x, map_y).Graphic(1), (map_x - 1) * 2, (map_y - 1) * 2, 1, MapData(map_x, map_y).Engine_Light(), 1, map_x, map_y, False, False, 0)
                    Call DDrawTransGrhtoSurface2(MapData(map_x, map_y).Graphic(1), (map_x - 1) * 2, (map_y - 1) * 2 + 1, 1, MapData(map_x, map_y).Engine_Light(), 1, map_x, map_y, False, False, 0)
                    Call DDrawTransGrhtoSurface2(MapData(map_x, map_y).Graphic(1), (map_x - 1) * 2 + 1, (map_y - 1) * 2, 1, MapData(map_x, map_y).Engine_Light(), 1, map_x, map_y, False, False, 0)
                    Call DDrawTransGrhtoSurface2(MapData(map_x, map_y).Graphic(1), (map_x - 1) * 2 + 1, (map_y - 1) * 2 + 1, 1, MapData(map_x, map_y).Engine_Light(), 1, map_x, map_y, False, False, 0)
                End If
                
                If MapData(map_x, map_y).Graphic(2).GrhIndex <> 0 Then
                    Call DDrawTransGrhtoSurface2(MapData(map_x, map_y).Graphic(2), (map_x - 1) * 2, (map_y - 1) * 2, 1, MapData(map_x, map_y).Engine_Light(), 1, map_x, map_y, False, False, 0)
                    Call DDrawTransGrhtoSurface2(MapData(map_x, map_y).Graphic(2), (map_x - 1) * 2, (map_y - 1) * 2 + 1, 1, MapData(map_x, map_y).Engine_Light(), 1, map_x, map_y, False, False, 0)
                    Call DDrawTransGrhtoSurface2(MapData(map_x, map_y).Graphic(2), (map_x - 1) * 2 + 1, (map_y - 1) * 2, 1, MapData(map_x, map_y).Engine_Light(), 1, map_x, map_y, False, False, 0)
                    Call DDrawTransGrhtoSurface2(MapData(map_x, map_y).Graphic(2), (map_x - 1) * 2 + 1, (map_y - 1) * 2 + 1, 1, MapData(map_x, map_y).Engine_Light(), 1, map_x, map_y, False, False, 0)
                End If
            Next map_x
        Next map_y
        
        For map_y = YMinMapSize To YMaxMapSize
            For map_x = XMinMapSize To XMaxMapSize
                If MapData(map_x, map_y).ObjGrh.GrhIndex <> 0 Then
                    Call DDrawTransGrhtoSurface2(MapData(map_x, map_y).ObjGrh, (map_x - 1) * 2, (map_y - 1) * 2, 1, MapData(map_x, map_y).Engine_Light(), 1, map_x, map_y, False, False, 0)
                    Call DDrawTransGrhtoSurface2(MapData(map_x, map_y).ObjGrh, (map_x - 1) * 2, (map_y - 1) * 2 + 1, 1, MapData(map_x, map_y).Engine_Light(), 1, map_x, map_y, False, False, 0)
                    Call DDrawTransGrhtoSurface2(MapData(map_x, map_y).ObjGrh, (map_x - 1) * 2 + 1, (map_y - 1) * 2, 1, MapData(map_x, map_y).Engine_Light(), 1, map_x, map_y, False, False, 0)
                    Call DDrawTransGrhtoSurface2(MapData(map_x, map_y).ObjGrh, (map_x - 1) * 2 + 1, (map_y - 1) * 2 + 1, 1, MapData(map_x, map_y).Engine_Light(), 1, map_x, map_y, False, False, 0)
                End If
            Next map_x
        Next map_y
        
    DirectDevice.EndScene
    DirectDevice.Present destRect, ByVal 0, Me.Picture1.hWnd, ByVal 0
    
End Sub
