VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Console"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Type Message
    Text As String
    Color As Long
    Bold As Boolean
    Italic As Boolean
End Type

Private Messages(1 To 255) As Message
Private OffsetTT As Byte
Private LastMessage As Byte

Property Get Offset() As Byte
    Offset = OffsetTT
End Property

Property Let Offset(ByVal Offset As Byte)
    OffsetTT = Offset
End Property

Public Function Add_Message(ByVal data As String, ByVal R As Byte, ByVal G As Byte, ByVal B As Byte, Optional ByVal Bold As Boolean = False, Optional ByVal Italic As Boolean = False)
    Dim Index As Byte
        Index = Find_Free_Index()
        
        Messages(Index).Text = data
        Messages(Index).Color = RGB(R, G, B)
        Messages(Index).Bold = Bold
        Messages(Index).Italic = Italic
        Debug.Print Index
    Call Update
End Function

Public Function Clean()
    Dim i As Integer
        For i = 1 To 255
            Messages(i).Text = vbNullString
            Messages(i).Color = -1
        Next i
        
    For i = 1 To 5
        frmMain.Console(i - 1).Caption = ""
    Next i
    
End Function

Private Function Find_Free_Index() As Byte
    Dim i As Integer
        For i = 1 To 255
            If Messages(i).Text = vbNullString Then
                Find_Free_Index = i
                Exit Function
            End If
        Next i
        
    Call Clean
    Find_Free_Index = 1
End Function



Public Function Delete_Index(ByVal Index As Byte)
    Set Messages(Index).Text = vbNullString
    Messages(Index).Color = -1
End Function

Private Function Get_Last() As Byte
Dim Last As Byte, i As Integer

        For i = 1 To 255
            If Messages(i).Text <> vbNullString Then
                Last = i
            End If
        Next i
        
        Get_Last = Last
End Function

Private Function Update()
Dim i As Long, pos As Byte
Dim Last As Byte, R As Byte, G As Byte, B As Byte
    pos = 0
    Last = Get_Last
    If Last = 0 Then Exit Function
    
    
    For i = Last - 4 To Last
        If i >= 1 Then
            If Messages(i).Text <> vbNullString Then
                frmMain.Console(pos).Caption = Messages(i).Text
                
                frmMain.Console(pos).ForeColor = Messages(i).Color
                frmMain.Console(pos).Font.Bold = Messages(i).Bold
                frmMain.Console(pos).Font.Italic = Messages(i).Italic
                
                frmMain.Console(pos).Refresh
                pos = pos + 1
            End If
        End If
    Next i
End Function

