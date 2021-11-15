Attribute VB_Name = "Mod_DX8_CountScreen"
'***************************************************
'Author: Ezequiel Juárez (Standelf)
'Last Modification: 23/12/10
'Blisse-AO | Sistema de Cuenta en Screen
'***************************************************
Option Explicit

Public Type Count
    min                     As Byte
    TickCount               As Long
    DoIt                    As Boolean
End Type

Public DX_Count             As Count
Public Const Count_Font     As Byte = 8


Public Sub RenderCount()
    '   Si no hay cuenta no rompemos mas el render
    If DX_Count.DoIt = False Then Exit Sub

    If DX_Count.min <> 0 Then
        '   Si no es 0 Dibujamos normal
        'Call Fonts_Render_String(DX_Count.min, (ScreenWidth - Fonts_Render_String_Width(DX_Count.min, Count_Font)) / 2, (ScreenHeight - Fuentes(Count_Font).CharactersHeight) / 2, -1, Count_Font)
    Else
        'Si es 0, Dibujamos el "@" que en la fuente está puesto como el "YA!"
        'Call Fonts_Render_String("@", (ScreenWidth - Fonts_Render_String_Width("@", Count_Font)) / 2, (ScreenHeight - Fuentes(Count_Font).CharactersHeight) / 2, -1, Count_Font)
    End If
    
    Fonts_Render_String DX_Count.min, 272, 40, ColorData.Blanco(1), True, 3
    
    'Checkeamos la cuenta, si es necesario restamos valor
    Call CheckCount
End Sub

Public Sub CheckCount()
'***************************************************
'Author: Ezequiel Juárez (Standelf)
'Last Modification: 23/12/10
'Check the count
'***************************************************

    '   Si no hay cuenta no rompemos mas
    If DX_Count.DoIt = False Then Exit Sub
    
        If GetTickCount - DX_Count.TickCount > 1000 Then
        '   Nos fijamos que haya pasado el tiempo
            If DX_Count.min > 0 Then
            '   Si es mayor a 0 restamos
                DX_Count.min = DX_Count.min - 1
                DX_Count.TickCount = GetTickCount
            ElseIf DX_Count.min = 0 Then
                '   Si es 0 quitamos la cuenta
                DX_Count.min = 0
                DX_Count.DoIt = False
            End If
        End If
End Sub

Public Sub InitCount(ByVal max As Byte)
'***************************************************
'Author: Ezequiel Juárez (Standelf)
'Last Modification: 23/12/10
'Check the count
'***************************************************

    '   Si hay cuenta no rompemos a la actual
    If DX_Count.DoIt = True Then Exit Sub
    
    With DX_Count
        '   Seteamos el Min
        .min = max
        
        '   Seteamos el tiempo
        .TickCount = GetTickCount
        
        '   Y entonces... DO IT!
        .DoIt = True
    End With
End Sub
