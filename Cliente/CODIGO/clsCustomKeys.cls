VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCustomKeys"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const NUM_CUSTOM_KEYS As Byte = 27
Private Const ERROR_NO_KEY_FILE As Long = vbObjectError + 16
Private Const KEY_CONFIG_FILE As String = "BindKeys.bin"
Private mappedKeys(1 To NUM_CUSTOM_KEYS) As Integer

Public Enum eKeyType
    mKeyUp = 1
    mKeyDown
    mKeyLeft
    mKeyRight
    mKeyToggleMusic
    mKeyToggleSound
    mKeyToggleFxs
    mKeyRequestRefresh
    mKeyToggleNames
    mKeyGetObject
    mKeyEquipObject
    mKeyTamAnimal
    mKeySteal
    mKeyToggleSafeMode
    mKeyToggleResuscitationSafe
    mKeyHide
    mKeyDropObject
    mKeyUseObject
    mKeyAttack
    mKeyTalk
    mKeyTalkWithGuild
    mKeyTakeScreenShot
    mKeyShowOptions
    mKeyMeditate
    mKeyCastSpellMacro
    mKeyWorkMacro
    mKeyExitGame
End Enum

Public Function ReadableName(ByVal KeyCode As Integer) As String
'***************************************************
'Author: Rapsodius
'Last Modification: 08/04/07
'
'***************************************************
    Select Case KeyCode
        Case KeyCodeConstants.vbKeyA To KeyCodeConstants.vbKeyZ
            ReadableName = Chr$(KeyCode)
        
        Case KeyCodeConstants.vbKeyNumpad0 To KeyCodeConstants.vbKeyNumpad9
            ReadableName = KeyCode - KeyCodeConstants.vbKeyNumpad0 & " (teclado num�rico)"
        
        Case KeyCodeConstants.vbKeyF1 To KeyCodeConstants.vbKeyF16
            ReadableName = "F" & (KeyCode - KeyCodeConstants.vbKeyF1 + 1)
        
        Case KeyCodeConstants.vbKeyMultiply
            ReadableName = "* (teclado numerico)"
        
        Case KeyCodeConstants.vbKeyAdd
            ReadableName = "+ (teclado numerico)"
        
        Case KeyCodeConstants.vbKeySeparator
            'Note: Separator appears in MSDN as the Enter key for
            'the numpad, but MS recognizes it as the normal
            'Enter key
            ReadableName = "Enter (teclado num�rico)"
        
        Case KeyCodeConstants.vbKeySubtract
            ReadableName = "- (teclado numerico)"
        
        Case KeyCodeConstants.vbKeyDecimal
            ReadableName = ". (teclado numerico)"
        
        Case KeyCodeConstants.vbKeyDivide
            ReadableName = "/ (teclado numerico)"
        
        Case KeyCodeConstants.vbKeyShift
            ReadableName = "Shift"
        
        Case KeyCodeConstants.vbKeyControl
            ReadableName = "Control"
        
        Case KeyCodeConstants.vbKeyMenu 'Alt
            ReadableName = "Alt"
        
        Case KeyCodeConstants.vbKeyPause
            ReadableName = "Pausa"
        
        Case KeyCodeConstants.vbKeyCapital
            ReadableName = "Blq Mayus"
        
        Case KeyCodeConstants.vbKeyEscape
            ReadableName = "Esc"
        
        Case KeyCodeConstants.vbKeyPageUp
            ReadableName = "Ret Pag"
        
        Case KeyCodeConstants.vbKeyPageDown
            ReadableName = "Av Pag"
        
        Case KeyCodeConstants.vbKeyEnd
            ReadableName = "Fin"
        
        Case KeyCodeConstants.vbKeyHome
            ReadableName = "Inicio"
        
        Case KeyCodeConstants.vbKeyLeft
            ReadableName = "Izquierda"
        
        Case KeyCodeConstants.vbKeyUp
            ReadableName = "Arriba"
        
        Case KeyCodeConstants.vbKeyRight
            ReadableName = "Derecha"
        
        Case KeyCodeConstants.vbKeyDown
            ReadableName = "Abajo"
        
        Case KeyCodeConstants.vbKeyInsert
            ReadableName = "Insertar"
        
        Case KeyCodeConstants.vbKeyNumlock
            ReadableName = "Blq Num"
        
        Case KeyCodeConstants.vbKeyScrollLock
            ReadableName = "Blq Despl"
        
        Case KeyCodeConstants.vbKeySpace
            ReadableName = "Barra Espaciadora"
        
        Case KeyCodeConstants.vbKeyBack
            ReadableName = "Borrar"
        
        Case KeyCodeConstants.vbKeyReturn
            ReadableName = "Enter"
        
        Case KeyCodeConstants.vbKeyDelete
            ReadableName = "Suprimir"
        
        Case Else
            'In all other cases, key is considered not valid.
            ReadableName = Constants.vbNullString
    End Select
End Function

''
'Reads custom keys from a file.
'
'@remarks   If the custom keys file is not found ERROR_NO_KEY_FILE is raised.

Public Sub LoadCustomKeys()
'***************************************************
'Author: Rapsodius
'Last Modification: 08/21/07
'
'***************************************************
    Dim i As Long
    Dim FileNum As Integer
    FileNum = FreeFile()
    
    'Inits keyboard config with default configurations
    Call LoadDefaults
        
    Open Resources.Bin & KEY_CONFIG_FILE For Binary Access Read Lock Read Write As FileNum
    
    For i = 1 To NUM_CUSTOM_KEYS
        Get FileNum, , mappedKeys(i)
        If LenB(ReadableName(mappedKeys(i))) = 0 Then
            Close FileNum
            Call Err.Raise(ERROR_NO_KEY_FILE)
            Exit Sub
        End If
    Next i
    
    Close FileNum
End Sub

''
'   Saves custom keys to a file

Public Sub SaveCustomKeys()
'***************************************************
'Author: Rapsodius
'Last Modification: 08/04/07
'
'***************************************************
    Dim FileNum As Integer
    Dim i As Long
    FileNum = FreeFile()
    
    Open Resources.Bin & KEY_CONFIG_FILE For Binary Access Write As FileNum
    
    For i = 1 To NUM_CUSTOM_KEYS
        Put FileNum, , mappedKeys(i)
    Next i
    
    Close #FileNum
End Sub

''
'Gets a key code
'
'@param index Index of the key code to retrieve
'@return The Key code that belongs to index

Public Property Get BindedKey(ByVal Index As eKeyType) As Integer
'***************************************************
'Author: Rapsodius
'Last Modification: 08/04/07
'
'***************************************************
    If Index < 1 Or Index > NUM_CUSTOM_KEYS Then Exit Property
    BindedKey = mappedKeys(Index)
End Property

''
'Sets a key code
'
'@param index Index of the key to change
'@param NewVal New key code

Public Property Let BindedKey(ByVal Index As eKeyType, ByVal NewVal As Integer)
'***************************************************
'Author: Rapsodius
'Last Modification: 08/05/07
'
'***************************************************
    If Index < 1 Or Index > NUM_CUSTOM_KEYS Then Exit Property
    
    If LenB(ReadableName(NewVal)) = 0 Then 'If key is not valid...
        Exit Property                      'Rejects the new assignment
    End If
    
    If KeyAssigned(NewVal) Then 'Also reject it in case key is already assigned
        Exit Property
    End If
    
    mappedKeys(Index) = NewVal
End Property

''
'Loads default keys

Public Sub LoadDefaults()
'***************************************************
'Author: Rapsodius
'Last Modification: 08/04/07
'
'***************************************************
    mappedKeys(eKeyType.mKeyUp) = KeyCodeConstants.vbKeyUp
    mappedKeys(eKeyType.mKeyDown) = KeyCodeConstants.vbKeyDown
    mappedKeys(eKeyType.mKeyLeft) = KeyCodeConstants.vbKeyLeft
    mappedKeys(eKeyType.mKeyRight) = KeyCodeConstants.vbKeyRight
    mappedKeys(eKeyType.mKeyToggleMusic) = KeyCodeConstants.vbKeyM
    mappedKeys(eKeyType.mKeyToggleSound) = KeyCodeConstants.vbKeyS
    mappedKeys(eKeyType.mKeyToggleFxs) = KeyCodeConstants.vbKeyF
    mappedKeys(eKeyType.mKeyRequestRefresh) = KeyCodeConstants.vbKeyL
    mappedKeys(eKeyType.mKeyToggleNames) = KeyCodeConstants.vbKeyN
    mappedKeys(eKeyType.mKeyGetObject) = KeyCodeConstants.vbKeyA
    mappedKeys(eKeyType.mKeyEquipObject) = KeyCodeConstants.vbKeyE
    mappedKeys(eKeyType.mKeyTamAnimal) = KeyCodeConstants.vbKeyD
    mappedKeys(eKeyType.mKeySteal) = KeyCodeConstants.vbKeyR
    mappedKeys(eKeyType.mKeyToggleSafeMode) = KeyCodeConstants.vbKeyMultiply
    mappedKeys(eKeyType.mKeyToggleResuscitationSafe) = KeyCodeConstants.vbKeyEnd
    mappedKeys(eKeyType.mKeyHide) = KeyCodeConstants.vbKeyO
    mappedKeys(eKeyType.mKeyDropObject) = KeyCodeConstants.vbKeyT
    mappedKeys(eKeyType.mKeyUseObject) = KeyCodeConstants.vbKeyU
    mappedKeys(eKeyType.mKeyAttack) = KeyCodeConstants.vbKeyControl
    mappedKeys(eKeyType.mKeyTalk) = KeyCodeConstants.vbKeyReturn
    mappedKeys(eKeyType.mKeyTalkWithGuild) = KeyCodeConstants.vbKeyDelete
    mappedKeys(eKeyType.mKeyTakeScreenShot) = KeyCodeConstants.vbKeyF2
    mappedKeys(eKeyType.mKeyShowOptions) = KeyCodeConstants.vbKeyF5
    mappedKeys(eKeyType.mKeyMeditate) = KeyCodeConstants.vbKeyF6
    mappedKeys(eKeyType.mKeyCastSpellMacro) = KeyCodeConstants.vbKeyF7
    mappedKeys(eKeyType.mKeyWorkMacro) = KeyCodeConstants.vbKeyF8
    mappedKeys(eKeyType.mKeyExitGame) = KeyCodeConstants.vbKeyF12
End Sub

''
'Returns whether a key is binded to an action
'
'@param KeyCode Key code to check
'@return True if key is mapped, false otherwise

Public Function KeyAssigned(ByVal KeyCode As Integer) As Boolean
'***************************************************
'Author: Rapsodius
'Last Modification: 08/04/07
'
'***************************************************
    Dim Counter As Long
    
    KeyAssigned = False
    
    For Counter = 1 To NUM_CUSTOM_KEYS
        If mappedKeys(Counter) = KeyCode Then
            KeyAssigned = True
            Exit For
        End If
    Next Counter
End Function

''
'Retrieves the number of customizable keys
'
'@return The number of customizable keys

Public Property Get Count() As Byte
'***************************************************
'Author: Rapsodius
'Last Modification: 08/04/07
'
'***************************************************
    Count = NUM_CUSTOM_KEYS
End Property

''
'Constructor. Loads Keys from file

Private Sub Class_Initialize()
'***************************************************
'Author: Rapsodius
'Last Modification: 08/04/07
'
'***************************************************
On Error GoTo ErrHand:
    Call LoadCustomKeys
    Exit Sub
    
ErrHand:
    If Err.Number = ERROR_NO_KEY_FILE Then
        Call LoadDefaults
        Call SaveCustomKeys
        Resume Next
    End If
End Sub

''
'Destructor. Saves Keys to file

Private Sub Class_Terminate()
'***************************************************
'Author: Rapsodius
'Last Modification: 08/04/07
'
'***************************************************
    Call SaveCustomKeys
End Sub
