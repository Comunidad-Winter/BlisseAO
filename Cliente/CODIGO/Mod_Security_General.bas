Attribute VB_Name = "Mod_Security_General"
'***************************************************
'Author: Ezequiel Juárez (Standelf)
'Last Modification: 26/05/10
'Blisse-AO | Security General.
'***************************************************

Private Declare Function CreateMutex Lib "kernel32" Alias "CreateMutexA" (ByRef lpMutexAttributes As SECURITY_ATTRIBUTES, ByVal bInitialOwner As Long, ByVal lpName As String) As Long
Private Declare Function ReleaseMutex Lib "kernel32" (ByVal hMutex As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long

Private Type SECURITY_ATTRIBUTES
    nLength As Long
    lpSecurityDescriptor As Long
    bInheritHandle As Long
End Type

Private Const ERROR_ALREADY_EXISTS = 183&
Private mutexHID As Long

Public Sub Init_Security()
'***************************************************
'Author: Standelf
'Last Modification: 26/25/10
'***************************************************
    '   Check Double Client
    #If Codeando = 0 Then
        Call Check_DoubleClient
    #End If

    '   Init Intervals
    Call Init_Intervals
End Sub

Public Sub DeInit_Security()
'***************************************************
'Author: Standelf
'Last Modification: 26/25/10
'Erase And Liberate Security
'***************************************************
    '   Liberate
    Call ReleaseInstance
    
    '   Erase intervals
    Erase Intervalos()
End Sub

Public Sub Check_DoubleClient()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 01/04/07
'***************************************************
    If FindPreviousInstance Then
        Call MsgBox(App.Title & " ya se está ejecutando actualmente.", vbApplicationModal + vbInformation + vbOKOnly, "Error al ejecutar")
        End
    End If
End Sub

Public Function ScrabbleString(ByVal String1 As String, ByVal String2 As String) As String
'***************************************************
'Author: Standelf
'Last Modification: 26/25/2010
'ScrabbleString 2 Strings
'***************************************************
Dim i As Byte
If Len(String1) <> Len(String2) Then Exit Function

    For i = 1 To Len(String1)
        ScrabbleString = ScrabbleString & mid(String1, i, 1)
        ScrabbleString = ScrabbleString & mid(String2, i, 1)
    Next i
End Function

Public Function GenerateKey(ByVal cantidad As Byte) As String
'***************************************************
'Author: Standelf
'Last Modification: 19/06/09
'Generates a security code of selected characters, the code is random with
'      letters (AA - Zz) And numbers(1, 9)
'***************************************************
Dim i As Byte, tempstring As String
    For i = 1 To cantidad
        If General_Get_Random_Number(1, 2) = 1 Then
            tempstring = tempstring & General_Get_Random_Number(1, 9)
        Else
            tempstring = tempstring & IIf(General_Get_Random_Number(1, 2) = 1, LCase$(Chr(97 + Rnd() * 862150000 Mod 26)), UCase$(Chr(97 + Rnd() * 862150000 Mod 26)))
        End If
    Next i
            
    GenerateKey = tempstring
End Function

Public Function RevisarCodigo() As Boolean
'***************************************************
'Author: Standelf
'Last Modification: 18/06/09
'Check a Random Key
'***************************************************
Dim CodigoOriginal As String * 6, CodigoUsuario As String * 6
    CodigoOriginal = GenerateKey(6) 'Generamos el código de 6 Digitos
    
        'Mostramos y solicitamos el codigo al usuario
        CodigoUsuario = InputBox("Ingrese a continuación el siguiente código de seguridad: " & CodigoOriginal, "Confirmación de Seguridad")

        If CodigoUsuario <> CodigoOriginal Then
            MsgBox "Código de seguridad Inválido. Recuerde que el código tiene sensibilidad de mayusculas y minusculas.", vbCritical
            RevisarCodigo = False
        Else
            RevisarCodigo = True
        End If
End Function

Private Function CreateNamedMutex(ByRef mutexName As String) As Boolean
'***************************************************
'Author: Fredy Horacio Treboux (liquid)
'Last Modification: 01/04/07
'Last Modified by: Juan Martín Sotuyo Dodero (Maraxus) - Changed Security Atributes to make it work in all OS
'***************************************************
    Dim sa As SECURITY_ATTRIBUTES
    
    With sa
        .bInheritHandle = 0
        .lpSecurityDescriptor = 0
        .nLength = LenB(sa)
    End With
    
    mutexHID = CreateMutex(sa, False, "Global\" & mutexName)
    
    CreateNamedMutex = Not (Err.LastDllError = ERROR_ALREADY_EXISTS) 'check if the mutex already existed
End Function

Public Function FindPreviousInstance() As Boolean
'***************************************************
'Author: Fredy Horacio Treboux (liquid)
'Last Modification: 01/04/07
'***************************************************
    'We try to create a mutex, the name could be anything, but must contain no backslashes.
    If CreateNamedMutex("UniqueNameThatActuallyCouldBeAnything") Then
        'There's no other instance running
        FindPreviousInstance = False
    Else
        'There's another instance running
        FindPreviousInstance = True
        Call Log_Unknown("Previous Instance: Double client found")
    End If
End Function

Public Sub ReleaseInstance()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 01/04/07
'***************************************************
    Call ReleaseMutex(mutexHID)
    Call CloseHandle(mutexHID)
End Sub
