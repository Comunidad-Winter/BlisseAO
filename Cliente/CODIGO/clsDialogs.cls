VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDialogs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Option Base 0

Private Const MAX_DIALOGS As Byte = 100
Private Const MAX_LENGTH As Byte = 18

Private Type dialog
    TextLine()  As String
    X           As Integer
    Y           As Integer

    CharIndex   As Integer
    Color       As D3DCOLORVALUE
    renderable  As Boolean
    
    Elevate As Integer
    Inicia As Boolean
    TickCountFadeOut As Long
    TickCountFadeIn As Long
    TickCOuntOutFinal As Long
End Type

Private dialogs(MAX_DIALOGS - 1) As dialog
Private dialogCount As Byte
Private Declare Function GetTickCount Lib "kernel32" () As Long

Private Sub Class_Initialize()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 07/28/07
'Initializes the dialog instance
'***************************************************
    dialogCount = 0
End Sub

''
'   Clean up and release resources

Private Sub Class_Terminate()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 07/28/07
'Clean up
'***************************************************
End Sub

Private Function BinarySearch(ByVal CharIndex As Integer) As Integer
'**************************************************************
'Author: Juan Martín Sotuyo Dodero
'Last Modify Date: 07/28/07
'Returns the index of the dialog in the list, or the negation
'of the position were it should be if not found (for binary insertion)
'**************************************************************
    Dim min As Long
    Dim max As Long
    Dim mid As Long
    
    min = 0
    max = dialogCount - 1
    
    Do While min <= max
        mid = (min + max) \ 2
        
        If dialogs(mid).CharIndex < CharIndex Then
            min = mid + 1
        ElseIf dialogs(mid).CharIndex > CharIndex Then
            max = mid - 1
        Else
            'We found it
            BinarySearch = mid
            Exit Function
        End If
    Loop
    
    'Not found, return the negation of the position where it should be
    '(all higher values are to the right of the list and lower values are to the left)
    BinarySearch = Not min
End Function

''
'   Creates a new dialog. If the user already had a dialog, then it's replaced,
'   otherwise a new slot is used.
'
'   @param    chat        The text of the chat being spoken.
'   @param    charIndex   The character saying the chat.
'   @param    color       The color being used in the chat when rendered.
'
'   @remarks  If there are no more free slots, and the char didn't have a previous dialog,
'             the chat is ignored.

Public Sub CreateDialog(ByRef chat As String, ByVal CharIndex As Integer, ByRef Color As D3DCOLORVALUE)
'**************************************************************
'Author: Juan Martín Sotuyo Dodero
'Last Modification: 08/29/07 by Rapsodius
'Changed life time of the dialogs to a variable amount,
'depending on the number of chars
'**************************************************************
    Dim Slot As Integer
    Dim i As Long
    
    Slot = BinarySearch(CharIndex)
    
    If Slot < 0 Then
        If dialogCount = MAX_DIALOGS Then Exit Sub  'Out of space! Should never happen....
        
        'We need to add it. Get insertion index and move list backwards.
        Slot = Not Slot
        
        For i = dialogCount To Slot + 1 Step -1
            dialogs(i) = dialogs(i - 1)
        Next i
        
        dialogCount = dialogCount + 1
    End If
    
    With dialogs(Slot)
        .CharIndex = CharIndex
        .Color = Color
        .renderable = False
        .TextLine = FormatChat(chat)
        
        .Elevate = 20
        .TickCountFadeIn = GetTickCount
        .TickCountFadeOut = .TickCountFadeIn
        .TickCOuntOutFinal = .TickCountFadeIn
        .Inicia = True
    End With
End Sub

''
'   Splits a text into several lines to make it comply with the MAX_LENGTH unless it's impossible (a single word longer than MAX_LENGTH).
'
'   @param    chat    The text to be formated.
'
'   @return   The array of lines into which the text is splitted.
'
'   @see      MAX_LENGTH

Private Function FormatChat(ByRef chat As String) As String()
'**************************************************************
'Author: Juan Martín Sotuyo Dodero
'Last Modify Date: 07/28/07
'Formats a dialog into different text lines.
'**************************************************************
    Dim word As String
    Dim curPos As Long
    Dim Length As Long
    Dim acumLength As Long
    Dim lineLength As Long
    Dim wordLength As Long
    Dim curLine As Long
    Dim chatLines() As String
    
    'Initialize variables
    curLine = 0
    curPos = 1
    Length = Len(chat)
    acumLength = 0
    lineLength = -1
    ReDim chatLines(General_Get_Field_Count(chat, 32)) As String
    
    'Start formating
    Do While acumLength < Length
        word = General_Get_ReadField(curPos, chat, 32)
        
        wordLength = Len(word)
        
        '   Is the first word of the first line? (it's the only that can start at -1)
        If lineLength = -1 Then
            chatLines(curLine) = word
       
            lineLength = wordLength
            acumLength = wordLength
        Else
            '   Is the word too long to fit in this line?
            If lineLength + wordLength + 1 > MAX_LENGTH Then
                'Put it in the next line
                curLine = curLine + 1
                chatLines(curLine) = word
                
                lineLength = wordLength
            Else
                'Add it to this line
                chatLines(curLine) = chatLines(curLine) & " " & word
                
                lineLength = lineLength + wordLength + 1
            End If
            
            acumLength = acumLength + wordLength + 1
        End If
        
        'Increase to search for next word
        curPos = curPos + 1
    Loop
    
    '   If it's only one line, center text
    If Settings.Dialog_Align = 1 Then
        If curLine = 0 And Length < MAX_LENGTH Then
            chatLines(curLine) = String((MAX_LENGTH - Length) \ 2 + 1, " ") & chatLines(curLine)
        End If
    End If
    
    'Resize array to fit
    ReDim Preserve chatLines(curLine) As String
    
    FormatChat = chatLines
End Function

''
'   Updates the position where a given dialog will be rendered.
'
'   @param    x           The X coord where to start rendering.
'   @param    y           The Y coord where to start rendering.
'   @param    charIndex   The character whose dialog's position is being updated.

Public Sub UpdateDialogPos(ByVal X As Integer, ByVal Y As Integer, ByVal CharIndex As Integer)
'**************************************************************
'Author: Juan Martín Sotuyo Dodero
'Last Modify Date: 07/28/07
'Updates a dialog's position.
'**************************************************************
    Dim Slot As Integer
    
    Slot = BinarySearch(CharIndex)
    
    If Slot >= 0 Then
        With dialogs(Slot)
            .X = IIf(Settings.Dialog_Align = 1, X - (MAX_LENGTH \ 2) * 2, X)
            .Y = Y - (UBound(.TextLine()) * 3)
            .renderable = True
        End With
    End If
End Sub

''
'   Displays all renderable dialogs on screen.

Public Sub Render()
'**************************************************************
'Author: Juan Martín Sotuyo Dodero
'Last Modification: 08/29/07 by Rapsodius
'Changed the way it gets rid of older dialogs.
'**************************************************************
    Dim i As Long
    Dim J As Long
    Dim Offset As Integer
    Dim removedDialogs As Long
    
    For i = 0 To dialogCount - 1
        'Decrease index to prevent jumping over a dialog
        'Crappy VB will cache the limit of the For loop, so even if it changed, it won't matter
        With dialogs(i - removedDialogs)

            If .renderable Then

                'Subir Letras
                If GetTickCount - .TickCountFadeIn > 9 Then
                    .Elevate = .Elevate - 1
                    .TickCountFadeIn = GetTickCount
                End If
                
                If .Elevate >= 0 Then .Y = .Y + .Elevate


                If Settings.Text_Effect = False Then
                    .Color.a = 255
                    
                    If GetTickCount - .TickCountFadeOut > 10000 Then
                        Call RemoveDialog(.CharIndex)
                        removedDialogs = removedDialogs + 1
                        Exit Sub
                    End If
                Else
                    'Alpha de Aparición
                    If .Color.a <= 255 And .Inicia = True And GetTickCount - .TickCOuntOutFinal > 9 Then
                        .TickCOuntOutFinal = GetTickCount
                        If Not .Color.a = 255 Then .Color.a = .Color.a + 5
                        If .Color.a >= 255 Then .Inicia = False: .Color.a = 255
                    End If
                                    
                    'Alpha de Desaparición
                    If Not .Inicia And GetTickCount - .TickCountFadeOut > 10000 Then 'Give 10 Seconds to read Dialog
                        If .Color.a >= 0 And Not .Inicia And GetTickCount - .TickCOuntOutFinal > 18 Then
                            .Color.a = .Color.a - 5
                            .TickCOuntOutFinal = GetTickCount
                                'Si el Alpha es 0, removemos el dialogo
                                If .Color.a <= 0 Then
                                    Call RemoveDialog(.CharIndex)
                                    removedDialogs = removedDialogs + 1
                                    Exit Sub
                                End If
                        End If
                    End If
                End If
                
                Offset = -10 * UBound(.TextLine())

                'Render it
                For J = 0 To UBound(.TextLine())
                    If Settings.Dialog_Align = 0 Then
                        Call Fonts_Render_String_RGBA(.TextLine(J), .X, .Y + Offset, .Color.r, .Color.g, .Color.b, .Color.a, True)
                    Else
                        Call Fonts_Render_String_RGBA(.TextLine(J), .X, .Y + Offset, .Color.r, .Color.g, .Color.b, .Color.a)
                    End If
                    
                    Offset = Offset + 12
                Next J
                
                'Won't render it unless it's updated in next frmae
                .renderable = False
            End If
        End With
    Next i
End Sub

'   Removes the dialog from the requested char.
'
'   @param    charIndex   The character whose dialog will be removed.

Public Sub RemoveDialog(ByVal CharIndex As Integer)
'**************************************************************
'Author: Juan Martín Sotuyo Dodero
'Last Modify Date: 07/28/07
'Removes the dialog from the requested char.
'**************************************************************
    Dim Slot As Integer
    Dim i As Long
    
    Slot = BinarySearch(CharIndex)
    
    If Slot < 0 Then Exit Sub
    
    For i = Slot To MAX_DIALOGS - 2
        dialogs(i) = dialogs(i + 1)
    Next i
    
    dialogCount = dialogCount - 1
End Sub

''
'   Removes all dialogs from all characters.

Public Sub RemoveAllDialogs()
'**************************************************************
'Author: Juan Martín Sotuyo Dodero
'Last Modify Date: 07/28/07
'Removes all dialogs from all chars.
'**************************************************************
    dialogCount = 0
End Sub

