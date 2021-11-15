Attribute VB_Name = "Mod_DX8_Font_Engine"
Option Explicit
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)

'Describes the return from a texture init
Private Type D3DXIMAGE_INFO_A
    Width As Long
    Height As Long
    Depth As Long
    MipLevels As Long
    Format As CONST_D3DFORMAT
    ResourceType As CONST_D3DRESOURCETYPE
    ImageFileFormat As Long
End Type
Private Type POINTAPI
    X As Long
    Y As Long
End Type
Private Type CharVA
    Vertex(0 To 3) As D3DTLVERTEX
End Type
Private Type VFH
    BitmapWidth As Long         'Size of the bitmap itself
    BitmapHeight As Long
    CellWidth As Long           'Size of the cells (area for each character)
    CellHeight As Long
    BaseCharOffset As Byte      'The character we start from
    CharWidth(0 To 255) As Byte 'The actual factual width of each character
    CharVA(0 To 255) As CharVA
End Type
Private Type CustomFont
    HeaderInfo As VFH           'Holds the header information
    Texture As Direct3DTexture8 'Holds the texture of the text
    RowPitch As Integer         'Number of characters per row
    RowFactor As Single         'Percentage of the texture width each character takes
    ColFactor As Single         'Percentage of the texture height each character takes
    CharHeight As Byte          'Height to use for the text - easiest to start with CellHeight value, and keep lowering until you get a good value
    TextureSize As POINTAPI     'Size of the texture
End Type
 
'Private Const Font_Default_TextureNum As Long = -1   'The texture number used to represent this font - only used for AlternateRendering - keep negative to prevent interfering with game textures
Private cFont(1 To 3) As CustomFont ' _Default2 As CustomFont

Public Function Engine_GetTextWidth(ByRef Font As Byte, ByVal Text As String) As Integer
'***************************************************
'Returns the width of text
'More info: http://www.vbgore.com/GameClient.TileEn ... tTextWidth
'***************************************************
Dim i As Integer
 
    'Make sure we have text
    If LenB(Text) = 0 Then Exit Function
   
    'Loop through the text
    For i = 1 To Len(Text)
       
        'Add up the stored character widths
        Engine_GetTextWidth = Engine_GetTextWidth + cFont(Font).HeaderInfo.CharWidth(Asc(mid$(Text, i, 1)))
       
    Next i
 
End Function

 
Sub Engine_Init_FontSettings(ByVal Font As Byte)
'*********************************************************
'****** Coded by Dunkan (emanuel.m@dunkancorp.com) *******
'*********************************************************
Dim FileNum As Byte
Dim LoopChar As Long
Dim Row As Single
Dim u As Single
Dim v As Single
 
    '*** Default font ***
 
    'Load the header information
    FileNum = FreeFile
    Open Resources.Bin & "font" & Font & ".dat" For Binary As #FileNum
        Get #FileNum, , cFont(Font).HeaderInfo
    Close #FileNum
   
    'Calculate some common values
    cFont(Font).CharHeight = cFont(Font).HeaderInfo.CellHeight - 4
    cFont(Font).RowPitch = cFont(Font).HeaderInfo.BitmapWidth \ cFont(Font).HeaderInfo.CellWidth
    cFont(Font).ColFactor = cFont(Font).HeaderInfo.CellWidth / cFont(Font).HeaderInfo.BitmapWidth
    cFont(Font).RowFactor = cFont(Font).HeaderInfo.CellHeight / cFont(Font).HeaderInfo.BitmapHeight
   
    'Cache the verticies used to draw the character (only requires setting the color and adding to the X/Y values)
    For LoopChar = 0 To 255
       
        'tU and tV value (basically tU = BitmapXPosition / BitmapWidth, and height for tV)
        Row = (LoopChar - cFont(Font).HeaderInfo.BaseCharOffset) \ cFont(Font).RowPitch
        u = ((LoopChar - cFont(Font).HeaderInfo.BaseCharOffset) - (Row * cFont(Font).RowPitch)) * cFont(Font).ColFactor
        v = Row * cFont(Font).RowFactor
 
        'Set the verticies
        With cFont(Font).HeaderInfo.CharVA(LoopChar)
            .Vertex(0).Color = D3DColorARGB(255, 0, 0, 0)   'Black is the most common color
            .Vertex(0).rhw = 1
            .Vertex(0).Tu = u
            .Vertex(0).Tv = v
            .Vertex(0).sX = 0
            .Vertex(0).sY = 0
            .Vertex(0).sz = 0
           
            .Vertex(1).Color = D3DColorARGB(255, 0, 0, 0)
            .Vertex(1).rhw = 1
            .Vertex(1).Tu = u + cFont(Font).ColFactor
            .Vertex(1).Tv = v
            .Vertex(1).sX = cFont(Font).HeaderInfo.CellWidth
            .Vertex(1).sY = 0
            .Vertex(1).sz = 0
           
            .Vertex(2).Color = D3DColorARGB(255, 0, 0, 0)
            .Vertex(2).rhw = 1
            .Vertex(2).Tu = u
            .Vertex(2).Tv = v + cFont(Font).RowFactor
            .Vertex(2).sX = 0
            .Vertex(2).sY = cFont(Font).HeaderInfo.CellHeight
            .Vertex(2).sz = 0
           
            .Vertex(3).Color = D3DColorARGB(255, 0, 0, 0)
            .Vertex(3).rhw = 1
            .Vertex(3).Tu = u + cFont(Font).ColFactor
            .Vertex(3).Tv = v + cFont(Font).RowFactor
            .Vertex(3).sX = cFont(Font).HeaderInfo.CellWidth
            .Vertex(3).sY = cFont(Font).HeaderInfo.CellHeight
            .Vertex(3).sz = 0
        End With
       
    Next LoopChar
 
End Sub

 
Sub Engine_Init_FontTextures(ByVal Font As Byte)
On Error GoTo eDebug:
'*****************************************************************
'Init the custom font textures
'More info: http://www.vbgore.com/GameClient.TileEn ... ntTextures
'*****************************************************************
Dim TexInfo As D3DXIMAGE_INFO_A
 
    'Check if we have the device
    If DirectDevice.TestCooperativeLevel <> D3D_OK Then Exit Sub
 
    '*** Default font ***
   
    'Set the texture
    
    If Font = 1 Or Font = 3 Then
        Set cFont(Font).Texture = DirectD3D8.CreateTextureFromFileEx(DirectDevice, Resources.Bin & "Font" & Font & ".png", D3DX_DEFAULT, D3DX_DEFAULT, D3DX_DEFAULT, 0, D3DFMT_UNKNOWN, D3DPOOL_MANAGED, D3DX_FILTER_POINT, D3DX_FILTER_POINT, &HFFFAFFAC, ByVal 0, ByVal 0)
    Else
        Set cFont(Font).Texture = DirectD3D8.CreateTextureFromFileEx(DirectDevice, Resources.Bin & "Font" & Font & ".bmp", D3DX_DEFAULT, D3DX_DEFAULT, D3DX_DEFAULT, 0, D3DFMT_UNKNOWN, D3DPOOL_MANAGED, D3DX_FILTER_POINT, D3DX_FILTER_POINT, &HFF000000, ByVal 0, ByVal 0)
    End If
    
    'Store the size of the texture
    cFont(Font).TextureSize.X = TexInfo.Width
    cFont(Font).TextureSize.Y = TexInfo.Height
   
    Exit Sub
eDebug:
    If Err.Number = "-2005529767" Then
        MsgBox "Error en la textura utilizada de DirectX 8", vbCritical
        End
    End If
    End
 
End Sub

Private Sub Engine_Render_Text(ByRef Font As Byte, ByVal Text As String, ByVal X As Long, ByVal Y As Long, ByVal Color As Long, Optional ByVal Center As Boolean = False)

Dim TempVA(0 To 3) As D3DTLVERTEX
Dim tempstr() As String
Dim Count As Integer
Dim ascii() As Byte
Dim i As Long
Dim J As Long
Dim tempColor As Long
Dim ResetColor As Byte
Dim YOffset As Single
 
    DirectDevice.SetRenderState D3DRS_SRCBLEND, D3DBLEND_SRCALPHA
    DirectDevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_INVSRCALPHA
    'directdevice.SetRenderState D3DRS_ALPHABLENDENABLE, True
   
    'Check if we have the device
    If DirectDevice.TestCooperativeLevel <> D3D_OK Then Exit Sub
 
    'Check for valid text to render
    If LenB(Text) = 0 Then Exit Sub
   
    'Get the text into arrays (split by vbCrLf)
    tempstr = Split(Text, vbCrLf)
   
    'Set the temp color (or else the first character has no color)
    tempColor = Color
 
    'Set the texture
    DirectDevice.SetTexture 0, cFont(Font).Texture
   
    If Center Then
        X = X - Engine_GetTextWidth(Font, Text) * 0.5
    End If
   
    'Loop through each line if there are line breaks (vbCrLf)
    For i = 0 To UBound(tempstr)
        If Len(tempstr(i)) > 0 Then
            YOffset = i * cFont(Font).CharHeight
            Count = 0
       
            'Convert the characters to the ascii value
            ascii() = StrConv(tempstr(i), vbFromUnicode)
       
            'Loop through the characters
            For J = 1 To Len(tempstr(i))
 
                'Check for a key phrase
                'If ascii(j - 1) = 124 Then 'If Ascii = "|"
                '    KeyPhrase = (Not KeyPhrase)  'TempColor = ARGB 255/255/0/0
                '    If KeyPhrase Then TempColor = ARGB(255, 0, 0, alpha) Else ResetColor = 1
                'Else
 
                    'Render with triangles
                    'If AlternateRender = 0 Then
 
                        'Copy from the cached vertex array to the temp vertex array
                        CopyMemory TempVA(0), cFont(Font).HeaderInfo.CharVA(ascii(J - 1)).Vertex(0), 32 * 4
 
                        'Set up the verticies
                        TempVA(0).sX = X + Count
                        TempVA(0).sY = Y + YOffset
                       
                        TempVA(1).sX = TempVA(1).sX + X + Count
                        TempVA(1).sY = TempVA(0).sY
 
                        TempVA(2).sX = TempVA(0).sX
                        TempVA(2).sY = TempVA(2).sY + TempVA(0).sY
 
                        TempVA(3).sX = TempVA(1).sX
                        TempVA(3).sY = TempVA(2).sY
                       
                        'Set the colors
                        TempVA(0).Color = tempColor
                        TempVA(1).Color = tempColor
                        TempVA(2).Color = tempColor
                        TempVA(3).Color = tempColor
                       
                        'Draw the verticies
                        DirectDevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, TempVA(0), Len(TempVA(0))
                       
                     
                    'Shift over the the position to render the next character
                    Count = Count + cFont(Font).HeaderInfo.CharWidth(ascii(J - 1))
               
                'End If
               
                'Check to reset the color
                If ResetColor Then
                    ResetColor = 0
                    tempColor = Color
                End If
               
            Next J
           
        End If
    Next i
   
End Sub

Public Sub Fonts_Render_String(ByVal Text As String, ByVal X As Integer, ByVal Y As Integer, Optional ByVal c As Long = -1, Optional ByVal Center As Boolean = False, Optional ByVal Font As Byte = 1)
    Engine_Render_Text Font, Text, CLng(X), CLng(Y), c, Center
End Sub

Public Sub Fonts_Render_String_RGBA(ByVal Text As String, ByVal X As Integer, ByVal Y As Integer, Optional ByVal r As Byte = 255, Optional ByVal g As Byte = 255, Optional ByVal b As Byte = 255, Optional ByVal a As Byte = 255, Optional ByVal Center As Boolean = False, Optional ByVal Font As Byte = 1)
    Engine_Render_Text Font, Text, CLng(X), CLng(Y), D3DColorARGB(a, r, g, b), Center
End Sub

