VERSION 5.00
Begin VB.Form frmRenderMap 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "frmRenderMap"
   ClientHeight    =   11280
   ClientLeft      =   150
   ClientTop       =   780
   ClientWidth     =   12000
   ControlBox      =   0   'False
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
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   752
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   800
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Render 
      BackColor       =   &H00004000&
      BorderStyle     =   0  'None
      Height          =   3000
      Left            =   0
      ScaleHeight     =   200
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   200
      TabIndex        =   0
      Top             =   0
      Width           =   3000
   End
   Begin VB.Menu mnuGenerar 
      Caption         =   "Generar BMP"
   End
   Begin VB.Menu mnuR8 
      Caption         =   "Generar R8"
   End
   Begin VB.Menu mnuSalir 
      Caption         =   "Salir"
   End
End
Attribute VB_Name = "frmRenderMap"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Map_RECT As RECT


Private Sub Form_Load()
    With Map_RECT
        .bottom = 200 '800 / 32 = 25 * 4 = 100
        .Right = 200 '800 / 32 = 25 * 4 = 100
    End With
End Sub

Private Sub mnuGenerar_Click()
Dim X As Long
Dim Y As Long
Dim Xi As Long
Dim Yi As Long
Dim TMP_X As Long
Dim TMP_Y As Long


For Xi = 1 To 4
    For Yi = 1 To 4
                 
        
        Engine_BeginScene
        
        TMP_X = -1
            For X = (Xi * 25) - 24 To (Xi * 25)
            TMP_X = TMP_X + 1
                TMP_Y = -1
                For Y = (Yi * 25) - 24 To (Yi * 25)
                TMP_Y = TMP_Y + 1
                        If MapData(X, Y).Graphic(1).GrhIndex <> 0 Then
                            TileEngine_Render_GRH MapData(X, Y).Graphic(1), (TMP_X - 1) * 32, (TMP_Y - 1) * 32, 0, 0, X, Y, ColorData.Blanco
                        End If
                        If MapData(X, Y).Graphic(2).GrhIndex <> 0 Then
                            TileEngine_Render_GRH MapData(X, Y).Graphic(2), (TMP_X - 1) * 32, (TMP_Y - 1) * 32, 0, 0, X, Y, ColorData.Blanco
                        End If
                Next Y
            Next X
            
        TMP_X = -1
            For X = (Xi * 25) - 24 To (Xi * 25)
            TMP_X = TMP_X + 1
                TMP_Y = -1
                For Y = (Yi * 25) - 24 To (Yi * 25)
                TMP_Y = TMP_Y + 1
                
                    If MapData(X, Y).ObjGrh.GrhIndex <> 0 Then
                        TileEngine_Render_GRH MapData(X, Y).ObjGrh, (TMP_X - 1) * 32, (TMP_Y - 1) * 32, 1, 0, X, Y, ColorData.Blanco
                    End If
                    If MapData(X, Y).Graphic(3).GrhIndex <> 0 Then
                        TileEngine_Render_GRH MapData(X, Y).Graphic(3), (TMP_X - 1) * 32, (TMP_Y - 1) * 32, 1, 0, X, Y, ColorData.Blanco
                    End If
                Next Y
            Next X
          
        Engine_EndScene Map_RECT, Me.Render.hWnd
         
        Sleep 1000
    Next Yi
Next Xi

End Sub

Private Sub SaveBackBuffer(ByVal Format As CONST_D3DXIMAGE_FILEFORMAT, ByVal Name As String)
Dim PAL As PALETTEENTRY
Dim FileName As String

    FileName = App.Path & "\" & Name & ".bmp"
    
    PAL.Blue = 255
    PAL.Green = 255
    PAL.Red = 255

    DirectD3D8.SaveSurfaceToFile FileName, D3DXIFF_BMP, DirectDevice.GetBackBuffer(0, D3DBACKBUFFER_TYPE_MONO), PAL, Map_RECT

End Sub

Private Sub mnuR8_Click()
Dim X As Long
Dim Y As Long
        Engine_BeginScene
        
For X = 1 To 100
For Y = 1 To 100
    If MapData(X, Y).Graphic(1).GrhIndex <> 0 Then
        TileEngine_Render_GRH MapData(X, Y).Graphic(1), (X - 1) * 2, (Y - 1) * 2, 0, 0, X, Y, ColorData.Blanco, , , , 16
    End If
    
    If MapData(X, Y).Graphic(2).GrhIndex <> 0 Then
        TileEngine_Render_GRH MapData(X, Y).Graphic(2), (X - 1) * 2, (Y - 1) * 2, 0, 0, X, Y, ColorData.Blanco, , , , 16
    End If
Next Y
Next X

Dim Banco As Boolean, herreria As Boolean
' herreria = 23662
'banco = 23665
For X = 1 To 100
    For Y = 1 To 100
        If Banco And herreria Then Exit For
                    
        If CurMapAmbient.MapBlocks(X, Y).Area = eAreas.herreria And Not herreria Then
                TileEngine_Render_GrhIndex 23662, (X - 1) * 2, (Y - 1) * 2 - 2, 0, ColorData.Blanco
            herreria = True
        End If
        
        If CurMapAmbient.MapBlocks(X, Y).Area = eAreas.Banco And Not Banco Then
                TileEngine_Render_GrhIndex 23665, (X - 1) * 2 + 8, (Y - 1) * 2 + 4, 0, ColorData.Blanco
            Banco = True
        End If
                    
    Next Y

        If Banco And herreria Then Exit For

Next X

    

        Engine_EndScene Map_RECT, Me.Render.hWnd
        
        SaveBackBuffer D3DXIFF_BMP, "asd"
End Sub
