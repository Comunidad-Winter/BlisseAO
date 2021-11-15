Attribute VB_Name = "Mod_Duelos"
Option Explicit
Public Enum num
    Cero = 24674
    Uno = 24675
    Dos = 24676
    Tres = 24677
    Cuatro = 24678
    Cinco = 24679
    Seis = 24680
    Siete = 24681
    Ocho = 24682
    Nueve = 24683
    Figth = 24684
End Enum

Public Type Counter
    Iniciar As Boolean
    Valor As Byte
End Type
Public Contador As Counter

Public Sub DrawCounterScreen()

Dim Cnt_Color(0 To 3) As Long
Call Engine_Long_To_RGB_List(Cnt_Color(), D3DColorARGB(130, 255, 255, 255))

Select Case Contador.Valor
    Case 0 'End conteo
        Contador.Iniciar = False
        Exit Sub
    Case 1 'Figth!!!!!
        Call DDrawTransGrhIndextoSurface(num.Figth, (frmMain.MainViewPic.ScaleWidth / 2) - (GrhData(num.Figth).pixelWidth / 2), 10, 0, Cnt_Color(), 0, False, False)
        Exit Sub
    Case 2 'Uno
        Call DDrawTransGrhIndextoSurface(num.Uno, (frmMain.MainViewPic.ScaleWidth / 2) - (GrhData(num.Uno).pixelWidth / 2), 10, 0, Cnt_Color(), 0, False, False)
        Exit Sub
    Case 3 'dos
        Call DDrawTransGrhIndextoSurface(num.Dos, (frmMain.MainViewPic.ScaleWidth / 2) - (GrhData(num.Dos).pixelWidth / 2), 10, 0, Cnt_Color(), 0, False, False)
        Exit Sub
    Case 4
        Call DDrawTransGrhIndextoSurface(num.Tres, (frmMain.MainViewPic.ScaleWidth / 2) - (GrhData(num.Tres).pixelWidth / 2), 10, 0, Cnt_Color(), 0, False, False)
        Exit Sub
    Case 5
        Call DDrawTransGrhIndextoSurface(num.Cuatro, (frmMain.MainViewPic.ScaleWidth / 2) - (GrhData(num.Cuatro).pixelWidth / 2), 10, 0, Cnt_Color(), 0, False, False)
        Exit Sub
    Case 6
        Call DDrawTransGrhIndextoSurface(num.Cinco, (frmMain.MainViewPic.ScaleWidth / 2) - (GrhData(num.Cinco).pixelWidth / 2), 10, 0, Cnt_Color(), 0, False, False)
        Exit Sub
    Case 7
        Call DDrawTransGrhIndextoSurface(num.Seis, (frmMain.MainViewPic.ScaleWidth / 2) - (GrhData(num.Seis).pixelWidth / 2), 10, 0, Cnt_Color(), 0, False, False)
        Exit Sub
    Case 8
        Call DDrawTransGrhIndextoSurface(num.Siete, (frmMain.MainViewPic.ScaleWidth / 2) - (GrhData(num.Siete).pixelWidth / 2), 10, 0, Cnt_Color(), 0, False, False)
        Exit Sub
    Case 9
        Call DDrawTransGrhIndextoSurface(num.Ocho, (frmMain.MainViewPic.ScaleWidth / 2) - (GrhData(num.Ocho).pixelWidth / 2), 10, 0, Cnt_Color(), 0, False, False)
        Exit Sub
    Case 10
        Call DDrawTransGrhIndextoSurface(num.Nueve, (frmMain.MainViewPic.ScaleWidth / 2) - (GrhData(num.Nueve).pixelWidth / 2), 10, 0, Cnt_Color(), 0, False, False)
        Exit Sub
End Select
End Sub
