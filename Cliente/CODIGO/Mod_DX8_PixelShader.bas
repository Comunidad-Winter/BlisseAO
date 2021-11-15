Attribute VB_Name = "Mod_DX8_PixelShader"
Option Explicit

Public ShowShader As Boolean

Public Const dats           As String = "ps.1.1 tex t0 tex t1 add_sat r0, t0, t1 mul r0, t0, v0" 'add_sat r0, t0, t1
Public Const Glowsrc        As String = "vs.1.1 dcl_position v0 dcl_texcoord v3 mov oPos, v0 add oT0, v3, c0 add oT1, v3, c1 add oT2, v3, c2 add oT3, v3, c3"
Public Const Gussian_Blur   As String = "ps.1.4 def c0, 0.2f, 0.2f, 0.2f, 1.0f texld r0, t0 texld r1, t1 texld r2, t2 texld r3, t3 texld r4, t4 add r0, r0, r1 add r2, r2, r3 add r0, r0, r2 add r0, r0, r4 mul r0, r0, c0"
Public Const BrightPasssrc  As String = "ps.1.4 def c0,0.561797752,0.561797752,0.561797752,1 def c1,0.78125,0.78125,0.78125,1 def c2,1,1,1,1 def c3,0.1,0.1,0.1,1 texld r0,t0 mul_x4 r0,r0,c0 mul_x2 r1,r0,c1 add r1,r1,c2 mul r0,r0,r1 mov_x4 r2,c2 add r2,r2,c2 sub r0,r0,r2 mul_sat r0,r0,c3"
Public Const psOriginalColor As String = "ps.1.1 tex t0 tex t3 mul r0, t0, t3"

Public Const psDesaturation As String = "ps.1.1 def c0, 0.3, 0.59, 0.11, 1 def c1, 0.5, 0.5, 0.5, 0 tex t0 dp3 r0, t0, c0 add r0,c1,r0"


Public Const psIlumination As String = "ps.1.0 " & _
    "tex t0 " & _
    "tex t1 " & _
    "add r1,t0,v1 " & _
    "add r1,r1,t1 " & _
    "add r1,r1,v0 " & _
    "mov r0,r1"

Public ps1Normal As Long
Public ps1Desat As Long

Dim DXlngShaderArray() As Long
Dim DXlngShaderSize As Long
Dim DXBufferCode As D3DXBuffer


Public Function PixelShaderInit()

    ps1Normal = pixelShaderMakeFromMemory(dats)
    ps1Desat = pixelShaderMakeFromMemory(psDesaturation)

End Function

Public Function pixelShaderMake(PSFileName As String)
'Assemble a pixel shader from a file, returning its handle
On Error GoTo PSErr
Set DXBufferCode = DirectD3D8.AssembleShaderFromFile(PSFileName$, 0, "", Nothing)
DXlngShaderSize = DXBufferCode.GetBufferSize() / 4
ReDim DXlngShaderArray(DXlngShaderSize - 1)
DirectD3D8.BufferGetData DXBufferCode, 0&, 4&, DXlngShaderSize&, DXlngShaderArray(0)
pixelShaderMake = DirectDevice.CreatePixelShader(DXlngShaderArray(0))
Set DXBufferCode = Nothing
Exit Function
PSErr:
MsgBox "Unable to create pixel shader", vbCritical, ""
Set DXBufferCode = Nothing
End Function
Public Function pixelShaderMakeFromMemory(PSContents As String)
'Assemble a pixel shader from a string, return its handle
On Error GoTo PSErr
Set DXBufferCode = DirectD3D8.AssembleShader(PSContents$, 0, Nothing, "")
DXlngShaderSize = DXBufferCode.GetBufferSize() / 4
ReDim DXlngShaderArray(DXlngShaderSize - 1)
DirectD3D8.BufferGetData DXBufferCode, 0&, 4&, DXlngShaderSize&, DXlngShaderArray(0)
pixelShaderMakeFromMemory = DirectDevice.CreatePixelShader(DXlngShaderArray(0))
Set DXBufferCode = Nothing
Exit Function
PSErr:
MsgBox "Unable to create pixel shader", vbCritical, ""
Set DXBufferCode = Nothing
End Function
Public Sub pixelShaderSet(ByRef lngPixelShaderHandle As Long)
'Enables a pixel shader
DirectDevice.SetPixelShader lngPixelShaderHandle&
End Sub
Public Sub pixelShaderDelete(ByRef lngPixelShaderHandle As Long)
'Deletes a pixel shader
DirectDevice.DeletePixelShader lngPixelShaderHandle&
End Sub
Public Sub pixelShaderSetCRegister(RegIndex As Long, ValR As Single, ValG As Single, ValB As Single, ValA As Single)
Dim SinArr(3)
SinArr(0) = ValR
SinArr(1) = ValG
SinArr(2) = ValB
SinArr(3) = ValA
DirectDevice.SetPixelShaderConstant RegIndex, SinArr(0), 4
End Sub
