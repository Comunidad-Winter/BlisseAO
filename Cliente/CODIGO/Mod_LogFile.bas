Attribute VB_Name = "Mod_LogFile"
'***************************************************
'Author: Ezequiel Juárez (Standelf)
'Last Modification: 12/06/10
'Blisse-AO | Log Writer.
'***************************************************
'   @ param Argentina 1 - 0 Nigeria
'   @ param Argentina 4 - 1 North Korea
'   @ param Greek 0 - 2 Argentina
'   @ param Argentina 3 - 1 Mexico
'   @ param Argentina 0 - 4 Germany =(

Public GeneralLogFile As String
Public LogBuffer As String 'Save The Log!
Public LogInSession As Long

Public Parameter_EngineDebug As String
Public Parameter_ProtocolDebug As String
Public Parameter_UnknownDebug As String
Public Parameter_InitializateDebug As String

Public Sub Init_Logs()
    GeneralLogFile = Resources.Bin & "General.log"
    Parameter_EngineDebug = "Blisse-AO DirectX Engine Debug:  "
    Parameter_ProtocolDebug = "Blisse-AO Protocol Debug:  "
    Parameter_UnknownDebug = "Blisse-AO Unknown Debug:  "
    Parameter_InitializateDebug = vbCrLf & " .:: Blisse-AO " & Date & " - " & Time & " Cliente Iniciado ::. "
    LogInSession = 0
End Sub

Public Sub Log_Engine(ByVal Desc As String)
    If Desc = vbNullString Then Exit Sub
    Put_Log Parameter_EngineDebug & Desc
End Sub

Public Sub Log_Protocol(ByVal Desc As String)
    If Desc = vbNullString Then Exit Sub
    Put_Log Parameter_ProtocolDebug & Desc
End Sub

Public Sub Log_Unknown(ByVal Desc As String)
    If Desc = vbNullString Then Exit Sub
    Put_Log Parameter_UnknownDebug & Desc
End Sub

Public Sub Put_Log(ByVal Message As String)
    If LogInSession = 0 Then
        LogBuffer = Parameter_InitializateDebug
        Call Save_Log
    End If
    
    LogBuffer = Message & " (" & Date & " - " & Time & ")."
    Call Save_Log
    
    LogInSession = LogInSession + 1
End Sub

Public Sub Save_Log()
On Local Error Resume Next

    Dim NFile As Integer
    NFile = FreeFile
    
    Open GeneralLogFile For Append Shared As #NFile
        Print #NFile, LogBuffer
    Close #NFile
End Sub
