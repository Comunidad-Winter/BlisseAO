Attribute VB_Name = "Mod_Formularios"
Option Explicit
 
Public Type POINTAPI
    X As Long
    Y As Long
End Type
 
Private Type SAFEARRAYBOUND
    cElements As Long
    lLbound As Long
End Type
 
Private Type SAFEARRAY2D
    cDims As Integer
    fFeatures As Integer
    cbElements As Long
    cLocks As Long
    pvData As Long
    Bounds(0 To 1) As SAFEARRAYBOUND
End Type
 
Private Type BITMAP
    bmType As Long
    bmWidth As Long
    bmHeight As Long
    bmWidthBytes As Long
    bmPlanes As Integer
    bmBitsPixel As Integer
    bmBits As Long
End Type
 
Const HTCAPTION = 2
Const WM_NCLBUTTONDOWN = &HA1
Const RGN_OR = 2
 
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function CreateRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function CombineRgn Lib "gdi32" (ByVal hDestRgn As Long, ByVal hSrcRgn1 As Long, ByVal hSrcRgn2 As Long, ByVal nCombineMode As Long) As Long
Private Declare Function GetPixel Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function SetWindowRgn Lib "user32" (ByVal hWnd As Long, ByVal hRgn As Long, ByVal bRedraw As Long) As Long
Private Declare Function CreatePolygonRgn Lib "gdi32" (lpPoint As POINTAPI, ByVal nCount As Long, ByVal nPolyFillMode As Long) As Long
Private Declare Function CreateEllipticRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function GetObjectAPI Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long
Private Declare Function VarPtrArray Lib "msvbvm50.dll" Alias "VarPtr" (Ptr() As Any) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDst As Any, pSrc As Any, ByVal ByteLen As Long)
 
 
Public lRegion As Long
Private Declare Function SetLayeredWindowAttributes Lib "user32" _
                (ByVal hWnd As Long, _
                 ByVal crKey As Long, _
                 ByVal bAlpha As Byte, _
                 ByVal dwFlags As Long) As Long

Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" _
                (ByVal hWnd As Long, _
                 ByVal nIndex As Long) As Long

Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" _
               (ByVal hWnd As Long, _
                ByVal nIndex As Long, _
                ByVal dwNewLong As Long) As Long

Private Const GWL_EXSTYLE = (-20)
Private Const LWA_ALPHA = &H2
Private Const WS_EX_LAYERED = &H80000
 
Public Sub InitializeSurfaceCapture(frm As Form)
    lRegion = CreateRectRgn(0, 0, 0, 0)
    frm.Visible = False
End Sub
 
Public Sub ReleaseSurfaceCapture(frm As Form)
    ApplySurfaceTo frm
'    frm.Visible = True
    Call DeleteObject(lRegion)
End Sub
 
Public Sub ApplySurfaceTo(frm As Form)
    Call SetWindowRgn(frm.hWnd, lRegion, True)
End Sub
 
' Create a polygonal region - has to be more than 2 pts (or 4 input values)
Public Sub CreateSurfacefromPoints(ParamArray XY())
    Dim lRegionTemp As Long
    Dim XY2() As POINTAPI
    Dim nIndex As Integer
    Dim nTemp As Integer
    Dim nSize As Integer
    nSize = CInt(UBound(XY) / 2) - 1
    ReDim XY2(nSize + 2)
    nIndex = 0
    For nTemp = 0 To nSize
        XY2(nTemp).X = XY(nIndex)
        nIndex = nIndex + 1
        XY2(nTemp).Y = XY(nIndex)
        nIndex = nIndex + 1
    Next nTemp
    lRegionTemp = CreatePolygonRgn(XY2(0), (UBound(XY2) + 1), 2)
    Call CombineRgn(lRegion, lRegion, lRegionTemp, RGN_OR)
    Call DeleteObject(lRegionTemp)
End Sub
 
' Create a ciruclar/elliptical region
Public Sub CreateSurfacefromEllipse(X1 As Integer, Y1 As Integer, X2 As Integer, Y2 As Integer)
    Dim lRegionTemp As Long
    lRegionTemp = CreateEllipticRgn(X1, Y1, X2, Y2)
    Call CombineRgn(lRegion, lRegion, lRegionTemp, RGN_OR)
    Call DeleteObject(lRegionTemp)
End Sub
 
' Create a rectangular region
Public Sub CreateSurfacefromRect(X1 As Integer, Y1 As Integer, X2 As Integer, Y2 As Integer)
    Dim lRegionTemp As Long
    lRegionTemp = CreateRectRgn(X1, Y1, X2, Y2)
    Call CombineRgn(lRegion, lRegion, lRegionTemp, RGN_OR)
    Call DeleteObject(lRegionTemp)
End Sub
 
' My best creation (more like tweak) yet! Super fast routines qown j00!
Public Sub CreateSurfacefromMask(obj As Object, Optional lBackColor As Long)
    ' Insight: Down with getpixel!!
    Dim lReturn   As Long
    Dim lRgnTmp   As Long
    Dim lSkinRgn  As Long
    Dim lStart    As Long
    Dim lRow      As Long
    Dim lCol      As Long
    Dim glHeight  As Integer
    Dim glWidth   As Integer
    Dim pict() As Byte
    Dim pict2() As Byte
    Dim sa As SAFEARRAY2D
    Dim bmp As BITMAP
    GetObjectAPI obj.Picture, Len(bmp), bmp
    ' Load the bmp into a safearray ptr
    With sa
        .cbElements = 1
        .cDims = 2
        .Bounds(0).lLbound = 0
        .Bounds(0).cElements = bmp.bmHeight
        .Bounds(1).lLbound = 0
        .Bounds(1).cElements = bmp.bmWidthBytes
        .pvData = bmp.bmBits
    End With
    ' Unfortunately this only supports 256 color bmps (damn high bit graphics!!)
    If bmp.bmBitsPixel <> 8 Then
        CreateSurfacefromMask_GetPixel obj
        Exit Sub
    End If
    CopyMemory ByVal VarPtrArray(pict), VarPtr(sa), 4
    ' Get the dimensions for future reference
    glHeight = UBound(pict, 2)
    glWidth = UBound(pict, 1)
    ' Create an identity array to flip the damn inversed regions
    ReDim pict2(glWidth, glHeight)
    ' Flip em!
    Dim nTempX As Integer
    Dim nTempY As Integer
    For nTempX = glWidth To 0 Step -1
        For nTempY = glHeight To 0 Step -1
            pict2(nTempX, nTempY) = pict(nTempX, glHeight - nTempY)
        Next nTempY
    Next nTempX
    ' Clear the original array
    CopyMemory ByVal VarPtrArray(pict), 0&, 4
    ' Let's make our regions!
    lSkinRgn = CreateRectRgn(0, 0, 0, 0)
    With obj
        If lBackColor < 1 Then lBackColor = pict2(0, 0)
        For lRow = 0 To glHeight
            lCol = 0
            Do While lCol < glWidth
                Do While lCol < glWidth
                    If pict2(lCol, lRow) = lBackColor Then
                        lCol = lCol + 1
                    Else
                        Exit Do
                    End If
                Loop
                If lCol < glWidth Then
                    lStart = lCol
                    Do While lCol < glWidth
                        If pict2(lCol, lRow) <> lBackColor Then
                            lCol = lCol + 1
                        Else
                            Exit Do
                        End If
                    Loop
                    If lCol > glWidth Then lCol = glWidth
                    lRgnTmp = CreateRectRgn(lStart, lRow, lCol, (lRow + 1))
                    lReturn = CombineRgn(lSkinRgn, lSkinRgn, lRgnTmp, RGN_OR)
                    Call DeleteObject(lRgnTmp)
                End If
            Loop
        Next
    End With
    ' Clear the identity array
    CopyMemory ByVal VarPtrArray(pict2), 0&, 4
    ' Return the f****** fast generated region!
    lReturn = CombineRgn(lRegion, lRegion, lSkinRgn, RGN_OR)
End Sub
 
' XCopied from The Scarms! Felt like my obligation to leave this code intact w/o
' any changes to variables, etc (cept for the sub's name). Thanks d00d!
Public Sub CreateSurfacefromMask_GetPixel(obj As Object, Optional lBackColor As Long)
    Dim lReturn   As Long
    Dim lRgnTmp   As Long
    Dim lSkinRgn  As Long
    Dim lStart    As Long
    Dim lRow      As Long
    Dim lCol      As Long
    Dim glHeight  As Integer
    Dim glWidth   As Integer
    lSkinRgn = CreateRectRgn(0, 0, 0, 0)
    With obj
        glHeight = .Height / Screen.TwipsPerPixelY
        glWidth = .Width / Screen.TwipsPerPixelX
        If lBackColor < 1 Then lBackColor = GetPixel(.hDC, 0, 0)
        For lRow = 0 To glHeight - 1
            lCol = 0
            Do While lCol < glWidth
                Do While lCol < glWidth And GetPixel(.hDC, lCol, lRow) = lBackColor
                    lCol = lCol + 1
                Loop
                If lCol < glWidth Then
                    lStart = lCol
                    Do While lCol < glWidth And GetPixel(.hDC, lCol, lRow) <> lBackColor
                        lCol = lCol + 1
                    Loop
                    If lCol > glWidth Then lCol = glWidth
                    lRgnTmp = CreateRectRgn(lStart, lRow, lCol, lRow + 1)
                    lReturn = CombineRgn(lSkinRgn, lSkinRgn, lRgnTmp, RGN_OR)
                    Call DeleteObject(lRgnTmp)
                End If
            Loop
        Next
    End With
    lReturn = CombineRgn(lRegion, lRegion, lSkinRgn, RGN_OR)
End Sub


'Función para saber si formulario ya es transparente. _
 Se le pasa el Hwnd del formulario en cuestión

Public Function Is_Transparent(ByVal hWnd As Long) As Boolean
On Error Resume Next

Dim msg As Long

    msg = GetWindowLong(hWnd, GWL_EXSTYLE)
       
       If (msg And WS_EX_LAYERED) = WS_EX_LAYERED Then
          Is_Transparent = True
       Else
          Is_Transparent = False
       End If

    If Err Then
       Is_Transparent = False
    End If

End Function

Public Function Set_Opacity(ByVal hWnd As Long, _
                                      Valor As Integer) As Long

Dim msg As Long

On Error Resume Next

If Valor < 0 Or Valor > 255 Then
   Set_Opacity = 1
Else
   msg = GetWindowLong(hWnd, GWL_EXSTYLE)
   msg = msg Or WS_EX_LAYERED
   
   SetWindowLong hWnd, GWL_EXSTYLE, msg
   
   SetLayeredWindowAttributes hWnd, 0, Valor, LWA_ALPHA

   Set_Opacity = 0

End If


If Err Then
   Set_Opacity = 2
End If

End Function

