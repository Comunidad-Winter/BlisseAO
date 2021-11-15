Attribute VB_Name = "Mod_DX8_FontsEngine"
'***************************************************
'Author: Ezequiel Juárez (Standelf)
'Last Modification: 15/05/10
'Blisse-AO | Renderizado de Fuentes en DirectX
'***************************************************

Option Explicit

Dim i As Integer
Dim J As Byte

Type fuente
    Characters(32 To 255) As Long 'ASCII Characters, Below 32 are useless.
    CharactersHeight As Byte
End Type

Public Fuentes() As fuente

Public Sub Fonts_DeInit()
'***************************************************
'Author: Ezequiel Juárez (Standelf)
'Last Modification: 16/05/10
'***************************************************

    'Kill Font's
    Erase Fuentes()
    
End Sub

Public Sub Fonts_Initializate()
'***************************************************
'Author: Ezequiel Juárez (Standelf)
'Last Modification: 16/05/10
'Load the fonts.
'***************************************************
    Dim Leer As New clsIniReader
    Dim Num_Fuentes As Byte
        
    Leer.Initialize Dir_Game & "Fonts.cfg"
    
    Num_Fuentes = Val(Leer.GetValue("Fuentes", "Num_Fuentes"))
    ReDim Fuentes(1 To Num_Fuentes)
    
        For J = 1 To Num_Fuentes
            For i = 32 To 255
                Fuentes(J).Characters(i) = Val(Leer.GetValue("Fuentes", "Fuentes(" & J & ").Caracteres(" & i & ") "))
                If i = 32 Then Fuentes(J).CharactersHeight = CByte(GrhData(Fuentes(J).Characters(i)).pixelHeight)
            Next i
        Next J
        
    Set Leer = Nothing
End Sub

Public Sub Fonts_Render_String(ByVal Text As String, ByVal X As Integer, ByVal Y As Integer, Optional ByVal c As Long = -1, Optional ByVal Font_Num As Byte = 1, Optional ByVal NoShadow As Boolean = False)
'***************************************************
'Author: Ezequiel Juárez (Standelf)
'Last Modification: 16/05/10
'Render the Text in the Screen.
'***************************************************
If Len(Text) = 0 Then Exit Sub

Dim Color(0 To 3) As Long, GrhIndex As Long, Suma As Integer
    Engine_Long_To_RGB_List Color(), c
    
    Suma = 0
    
    If Settings.Engine_Font_Sombras Then
        Dim ColorSombra(0 To 3) As Long
        Engine_Long_To_RGB_List ColorSombra(), D3DColorXRGB(10, 10, 10)
    End If
    
    For i = 1 To Len(Text)
    
        GrhIndex = Fuentes(Font_Num).Characters(Asc(mid(Text, i, 1)))
        If Settings.Engine_Font_Sombras And NoShadow = False Then DDrawTransGrhIndextoSurface GrhIndex, X + Suma + 1, Y + 1, 0, ColorSombra(), 0, False
        DDrawTransGrhIndextoSurface GrhIndex, X + Suma, Y, 0, Color(), 0, False
        Suma = Suma + GrhData(GrhIndex).pixelWidth
    Next i
    
End Sub

Public Sub Fonts_Render_String_RGBA(ByVal Text As String, ByVal X As Integer, ByVal Y As Integer, Optional ByVal Font_Num As Byte = 1, Optional ByVal r As Byte = 255, Optional ByVal g As Byte = 255, Optional ByVal b As Byte = 255, Optional ByVal a As Byte = 255)
'***************************************************
'Author: Ezequiel Juárez (Standelf)
'Last Modification: 27/01/2011
'Render the Text in the Screen with RGBA Colors.
'***************************************************
If Len(Text) = 0 Then Exit Sub

Dim Color(0 To 3) As Long, GrhIndex As Long, Suma As Integer
    Engine_Long_To_RGB_List Color(), D3DColorARGB(a, r, g, b)
    
    Suma = 0
    
    If Settings.Engine_Font_Sombras Then
        Dim ColorSombra(0 To 3) As Long
        Engine_Long_To_RGB_List ColorSombra(), D3DColorARGB(a, 10, 10, 10)
    End If
    
    For i = 1 To Len(Text)
    
        GrhIndex = Fuentes(Font_Num).Characters(Asc(mid(Text, i, 1)))
        If Settings.Engine_Font_Sombras Then DDrawTransGrhIndextoSurface GrhIndex, X + Suma + 1, Y + 1, 0, ColorSombra(), 0, False
        DDrawTransGrhIndextoSurface GrhIndex, X + Suma, Y, 0, Color(), 0, False
        Suma = Suma + GrhData(GrhIndex).pixelWidth
    Next i
    
End Sub

Public Function Fonts_Render_String_Width(ByVal Text As String, Optional ByVal Font_Num As Byte = 1) As Integer 'I dislike this...
'***************************************************
'Author: Ezequiel Juárez (Standelf)
'Last Modification: 16/05/10
'Return the Width of the text.
'***************************************************

If Len(Text) = 0 Then
    Fonts_Render_String_Width = 0
    Exit Function
End If

Dim Suma As Integer
    Suma = 0
    
    For i = 1 To Len(Text)
        Suma = Suma + GrhData(Fuentes(Font_Num).Characters(Asc(mid(Text, i, 1)))).pixelWidth
    Next i
    
    Fonts_Render_String_Width = Suma
End Function
