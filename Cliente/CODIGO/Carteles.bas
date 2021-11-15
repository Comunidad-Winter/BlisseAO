Attribute VB_Name = "Mod_DX8_Carteles"
Option Explicit

Const XPosCartel = 100
Const YPosCartel = 100
Const MAXLONG = 30

Public Cartel               As Boolean
Public Leyenda              As String
Public LeyendaFormateada()  As String
Public textura              As Integer

Sub InitCartel(Ley As String, Grh As Integer)
If Not Cartel Then
    Leyenda = Ley
    textura = Grh
    Cartel = True
    ReDim LeyendaFormateada(0 To (Len(Ley) \ (MAXLONG \ 2)))
                
    Dim i As Integer, K As Integer, anti As Integer
    anti = 1
    K = 0
    i = 0
    Call DarFormato(Leyenda, i, K, anti)
    i = 0
    Do While LeyendaFormateada(i) <> "" And i < UBound(LeyendaFormateada)
        
       i = i + 1
    Loop
    ReDim Preserve LeyendaFormateada(0 To i)
Else
    Exit Sub
End If
End Sub

Private Function DarFormato(s As String, i As Integer, K As Integer, anti As Integer)
If anti + i <= Len(s) + 1 Then
    If ((i >= MAXLONG) And mid$(s, anti + i, 1) = " ") Or (anti + i = Len(s)) Then
        LeyendaFormateada(K) = mid(s, anti, i + 1)
        K = K + 1
        anti = anti + i + 1
        i = 0
    Else
        i = i + 1
    End If
    Call DarFormato(s, i, K, anti)
End If
End Function

Sub DibujarCartel()
Dim X As Integer, Y As Integer
    If Not Cartel Then Exit Sub
    
    X = XPosCartel + 20
    Y = YPosCartel + 20
    
    Call TileEngine_Render_GrhIndex(textura, XPosCartel, YPosCartel, 0, ColorData.Blanco())
    Dim J As Integer, desp As Integer
    
    For J = 0 To UBound(LeyendaFormateada)
        Fonts_Render_String LeyendaFormateada(J), X, Y + desp, -1, False
        desp = desp + (frmMain.Font.Size) + 5
    Next
End Sub

