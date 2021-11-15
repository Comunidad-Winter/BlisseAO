VERSION 5.00
Begin VB.UserControl BGAOAniGif 
   ClientHeight    =   1290
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1470
   ScaleHeight     =   1290
   ScaleWidth      =   1470
   Begin VB.Timer Timer 
      Left            =   0
      Top             =   0
   End
   Begin VB.Image imgSource 
      Height          =   1035
      Index           =   0
      Left            =   0
      Top             =   0
      Width           =   1155
   End
End
Attribute VB_Name = "BGAOAniGif"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Dim mTotalFrames As Long
Dim mRepeatTimes As Long
Dim mGifPath As String
Dim FrameCount As Long

Private Sub Timer_Timer()
Dim i As Long
    If FrameCount < TotalFrames Then
        imgSource(FrameCount).Visible = False
        FrameCount = FrameCount + 1
        imgSource(FrameCount).Visible = True
        Timer.Interval = CLng(imgSource(FrameCount).Tag)
    Else
        FrameCount = 0
        For i = 1 To imgSource.Count - 1
            imgSource(i).Visible = False
        Next i
        imgSource(FrameCount).Visible = True
        Timer.Interval = CLng(imgSource(FrameCount).Tag)
    End If
End Sub

Private Sub UserControl_Initialize()
imgSource(0).Move 0, 0, ScaleWidth, ScaleHeight
End Sub

Private Sub UserControl_Resize()
imgSource(0).Move 0, 0, ScaleWidth, ScaleHeight
End Sub

Public Property Get TotalFrames() As Long
    TotalFrames = mTotalFrames
End Property

Public Property Let TotalFrames(ByVal vNewValue As Long)
    mTotalFrames = vNewValue
End Property

Public Property Get RepeatTimes() As Long
    RepeatTimes = mRepeatTimes
End Property

Public Property Let RepeatTimes(ByVal vNewValue As Long)
    mRepeatTimes = vNewValue
End Property

Public Property Get GifPath() As String
    GifPath = mGifPath
End Property

Public Property Let GifPath(ByVal vNewValue As String)
    If Dir(vNewValue) = "" Then
        Err.Raise vbObjectError + 1, , "File not found"
        Exit Property
    End If
    If Right(vNewValue, 3) <> "gif" Then
        Err.Raise vbObjectError + 2, , "File format is not supported"
        Exit Property
    End If
    mGifPath = vNewValue
End Property
Private Function LoadGif(sFile As String, aImg As Variant) As Boolean
    LoadGif = False
    If Dir$(sFile) = "" Or sFile = "" Then
       Err.Raise vbObjectError + 1, , "File not found"
       Exit Function
    End If
    On Error GoTo ErrHandler
    Dim fNum As Integer
    Dim imgHeader As String, fileHeader As String
    Dim buf$, picbuf$
    Dim imgCount As Integer
    Dim i&, J&, xOff&, yOff&, TimeWait&
    Dim GifEnd As String
    GifEnd = Chr(0) & Chr(33) & Chr(249)
    For i = 1 To aImg.Count - 1
        Unload aImg(i)
    Next i
    fNum = FreeFile
    Open sFile For Binary Access Read As fNum
        buf = String(LOF(fNum), Chr(0))
        Get #fNum, , buf 'Get GIF File into buffer
    Close fNum
    
    i = 1
    imgCount = 0
    J = InStr(1, buf, GifEnd) + 1
    fileHeader = Left(buf, J)
    If Left$(fileHeader, 3) <> "GIF" Then
       Err.Raise vbObjectError + 2, , "File format is not supported"
       Exit Function
    End If
    LoadGif = True
    i = J + 2
    If Len(fileHeader) >= 127 Then
        mRepeatTimes = Asc(mid(fileHeader, 126, 1)) + (Asc(mid(fileHeader, 127, 1)) * 256&)
    Else
        mRepeatTimes = 0
    End If

    Do ' Split GIF Files at separate pictures
       ' and load them into Image Array
        imgCount = imgCount + 1
        J = InStr(i, buf, GifEnd) + 3
        If J > Len(GifEnd) Then
            fNum = FreeFile
            Open "temp.gif" For Binary As fNum
                picbuf = String(Len(fileHeader) + J - i, Chr(0))
                picbuf = fileHeader & mid(buf, i - 1, J - i)
                Put #fNum, 1, picbuf
                imgHeader = Left(mid(buf, i - 1, J - i), 16)
            Close fNum
            TimeWait = ((Asc(mid(imgHeader, 4, 1))) + (Asc(mid(imgHeader, 5, 1)) * 256&)) * 10&
            If imgCount > 1 Then
                xOff = Asc(mid(imgHeader, 9, 1)) + (Asc(mid(imgHeader, 10, 1)) * 256&)
                yOff = Asc(mid(imgHeader, 11, 1)) + (Asc(mid(imgHeader, 12, 1)) * 256&)
                Load aImg(imgCount - 1)
                aImg(imgCount - 1).Left = aImg(0).Left + (xOff * Screen.TwipsPerPixelX)
                aImg(imgCount - 1).Top = aImg(0).Top + (yOff * Screen.TwipsPerPixelY)
            End If
            ' Use .Tag Property to save TimeWait interval for separate Image
            aImg(imgCount - 1).Tag = TimeWait
            aImg(imgCount - 1).Picture = LoadPicture("temp.gif")
            Kill ("temp.gif")
            i = J
        End If
        DoEvents
    Loop Until J = 3
' If there are one more Image - Load it
    If i < Len(buf) Then
        fNum = FreeFile
        Open "temp.gif" For Binary As fNum
            picbuf = String(Len(fileHeader) + Len(buf) - i, Chr(0))
            picbuf = fileHeader & mid(buf, i - 1, Len(buf) - i)
            Put #fNum, 1, picbuf
            imgHeader = Left(mid(buf, i - 1, Len(buf) - i), 16)
        Close fNum
        TimeWait = ((Asc(mid(imgHeader, 4, 1))) + (Asc(mid(imgHeader, 5, 1)) * 256)) * 10
        If imgCount > 1 Then
            xOff = Asc(mid(imgHeader, 9, 1)) + (Asc(mid(imgHeader, 10, 1)) * 256)
            yOff = Asc(mid(imgHeader, 11, 1)) + (Asc(mid(imgHeader, 12, 1)) * 256)
            Load aImg(imgCount - 1)
            aImg(imgCount - 1).Left = aImg(0).Left + (xOff * Screen.TwipsPerPixelX)
            aImg(imgCount - 1).Top = aImg(0).Top + (yOff * Screen.TwipsPerPixelY)
        End If
        aImg(imgCount - 1).Tag = TimeWait
        aImg(imgCount - 1).Picture = LoadPicture("temp.gif")
        Kill ("temp.gif")
    End If
    TotalFrames = aImg.Count - 1
    Exit Function
ErrHandler:
    Err.Raise Err.Number, Err.Source, Err.Description
    LoadGif = False
    On Error GoTo 0
End Function


Public Sub StartGif()
    Timer.Enabled = False
    If LoadGif(mGifPath, imgSource) Then
       FrameCount = 0
       Timer.Interval = CLng(imgSource(0).Tag)
       Timer.Enabled = True
    End If
End Sub

Public Sub StopGif()
    Timer.Enabled = False
End Sub

Public Sub ContinueGif()
    Timer.Enabled = True
End Sub

