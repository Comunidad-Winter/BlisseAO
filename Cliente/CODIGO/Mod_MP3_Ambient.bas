Attribute VB_Name = "Mod_MP3_Ambient"
Option Explicit

Private MP3_In_Use As eMP3

Public Enum eMP3
    Magma = 1
    Water = 2
    Rain1 = 3
    City1 = 4
    City2 = 5
    Rain2 = 6
    House = 7
    Dessert = 8
    MainMenu = 9
    Bar = 10
    Dungeon = 11
End Enum

Public Sub Play_MP3(ByVal MP3 As eMP3)
If Settings.Musica = False Then Exit Sub

    If MP3_In_Use <> MP3 Then
        If MP3_In_Use <> 0 Then Audio.MP3_Stop
        
        DoEvents
        
        If General_File_Exist(Resources.MP3 & MP3 & ".mp3", vbNormal) Then
            Audio.MP3_Play CStr(MP3)
            MP3_In_Use = MP3
        End If
    End If
End Sub
