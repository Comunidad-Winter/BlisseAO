Attribute VB_Name = "Mod_DDEx"
Public DDEx As New Cls_DDEx

Public Type tDDEXRGBA
    a As Byte
    r As Byte
    g As Byte
    b As Byte
End Type

Public Function DDEXRGBA(a As Byte, r As Byte, g As Byte, b As Byte) As tDDEXRGBA
    DDEXRGBA.a = a
    DDEXRGBA.r = r
    DDEXRGBA.g = g
    DDEXRGBA.b = b
End Function

Public Sub IniciarDDEx()
    Call DDEx.Iniciar(frmMain.Picture1.hwnd, "Game\Resources\Graficos", DX9, DX9_SOF, False)
End Sub
