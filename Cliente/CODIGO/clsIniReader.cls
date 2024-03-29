VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsIniReader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Option Base 0

Private Type ChildNode
    Key As String
    value As String
End Type



Private Type MainNode
    name As String
    values() As ChildNode
    numValues As Integer
End Type

''
'Containts all Main sections of the loaded INI file
'
Private fileData() As MainNode

''
'Stores the total number of main sections in the loaded INI file
'
Private MainNodes As Long

''
'Default constructor. Does nothing.
'

Private Sub Class_Initialize()
'**************************************************************
'Author: Juan Mart�n Sotuyo Dodero
'Last Modify Date: 5/01/2006
'
'**************************************************************
End Sub

''
'Destroy every array and deallocates al memory.
'

Private Sub Class_Terminate()
'**************************************************************
'Author: Juan Mart�n Sotuyo Dodero
'Last Modify Date: 5/01/2006
'
'**************************************************************
    Dim i As Long
    
    'Clean up
    If MainNodes Then
        For i = 1 To MainNodes - 1
            Erase fileData(i).values
        Next i
        
        Erase fileData
    End If
End Sub

''
'Loads a INI file so it's values can be read. Must be called before being able to use GetValue.
'
'   @param    file Complete path of the INI file to be loaded.
'   @see      GetValue
'

Public Sub Initialize(ByVal File As String)
'**************************************************************
'Author: Juan Mart�n Sotuyo Dodero
'Last Modify Date: 5/01/2006
'Opens the requested file and loads it's data into memory
'**************************************************************
    Dim handle As Integer
    Dim Text As String
    Dim Pos As Long
    
    'Get a free handle and start reading line by line until the end
    handle = FreeFile
    
    Open File For Input As handle
    
    Do Until EOF(handle)
        Line Input #handle, Text
        
        'Is it null??
        If Len(Text) Then
            'If it starts with '['   it is a main node or nothing (GetPrivateProfileStringA works this way), otherwise it's a value
            If Left$(Text, 1) = "[" Then
                'If it has an ending ']'   it's a main node, otherwise it's nothing
                Pos = InStr(2, Text, "]")
                If Pos Then
                    'Add a main node
                    ReDim Preserve fileData(MainNodes) As MainNode
                    
                    fileData(MainNodes).name = UCase$(Trim$(mid$(Text, 2, Pos - 2)))
                    
                    MainNodes = MainNodes + 1
                End If
            Else
                'So it's a value. Check if it has a '=', otherwise it's nothing
                Pos = InStr(2, Text, "=")
                If Pos Then
                    'Is it under any main node??
                    If MainNodes Then
                        With fileData(MainNodes - 1)
                            'Add it to the main node's value
                            ReDim Preserve .values(.numValues) As ChildNode
                            
                            .values(.numValues).value = Right$(Text, Len(Text) - Pos)
                            .values(.numValues).Key = UCase$(Left$(Text, Pos - 1))
                            
                            .numValues = .numValues + 1
                        End With
                    End If
                End If
            End If
        End If
    Loop
    
    Close handle
    
    Dim i As Long
    
    If MainNodes Then
        'Sort main nodes to allow binary search
        Call SortMainNodes(0, MainNodes - 1)
        
        'Sort values of each node to allow binary search
        For i = 0 To MainNodes - 1
            If fileData(i).numValues Then _
                Call SortChildNodes(fileData(i), 0, fileData(i).numValues - 1)
        Next i
    End If
End Sub

''
'Sorts all child nodes within the given MainNode alphabetically by their keys. Uses quicksort.
'
'   @param    Node The MainNode whose values are to be sorted.
'   @param    first The first index to consider when sorting.
'   @param    last The last index to be considered when sorting.
'

Private Sub SortChildNodes(ByRef Node As MainNode, ByVal first As Integer, ByVal last As Integer)
'**************************************************************
'Author: Juan Mart�n Sotuyo Dodero
'Last Modify Date: 5/01/2006
'Sorts the list of values in a given MainNode using quicksort,
'this allows the use of Binary Search for faster searches
'**************************************************************
    Dim min As Integer      'First item in the list
    Dim max As Integer      'Last item in the list
    Dim comp As String      'Item used to compare
    Dim Temp As ChildNode
    
    min = first
    max = last
    
    With Node
        comp = .values((min + max) \ 2).Key
    
        Do While min <= max
            Do While .values(min).Key < comp And min < last
                min = min + 1
            Loop
            Do While .values(max).Key > comp And max > first
                max = max - 1
            Loop
            If min <= max Then
                Temp = .values(min)
                .values(min) = .values(max)
                .values(max) = Temp
                min = min + 1
                max = max - 1
            End If
        Loop
    End With
    
    If first < max Then SortChildNodes Node, first, max
    If min < last Then SortChildNodes Node, min, last
End Sub

''
'Sorts all main nodes in the loaded INI file alphabetically by their names. Uses quicksort.
'
'   @param    first The first index to consider when sorting.
'   @param    last The last index to be considered when sorting.
'

Private Sub SortMainNodes(ByVal first As Integer, ByVal last As Integer)
'**************************************************************
'Author: Juan Mart�n Sotuyo Dodero
'Last Modify Date: 5/01/2006
'Sorts the MainNodes list using quicksort,
'this allows the use of Binary Search for faster searches
'**************************************************************
    Dim min As Integer      'First item in the list
    Dim max As Integer      'Last item in the list
    Dim comp As String      'Item used to compare
    Dim Temp As MainNode
    
    min = first
    max = last
    
    comp = fileData((min + max) \ 2).name
    
    Do While min <= max
        Do While fileData(min).name < comp And min < last
            min = min + 1
        Loop
        Do While fileData(max).name > comp And max > first
            max = max - 1
        Loop
        If min <= max Then
            Temp = fileData(min)
            fileData(min) = fileData(max)
            fileData(max) = Temp
            min = min + 1
            max = max - 1
        End If
    Loop
    
    If first < max Then SortMainNodes first, max
    If min < last Then SortMainNodes min, last
End Sub

''
'Searches for a given key within a given main section and if it exists retrieves it's value, otherwise a null string
'
'   @param    Main The name of the main section in which we will be searching.
'   @param    key The key of the value we are looking for.
'   @returns  The value asociated with the given key under the requeted main section of the INI file or a null string if it's not found.
'

Public Function GetValue(ByVal Main As String, ByVal Key As String) As String
'**************************************************************
'Author: Juan Mart�n Sotuyo Dodero
'Last Modify Date: 5/01/2006
'Returns a value if the key and main node exist, or a nullstring otherwise
'**************************************************************
    Dim i As Long
    Dim J As Long
    
    'Search for the main node
    i = FindMain(UCase$(Main))
    
    If i >= 0 Then
        'If valid, binary search among keys
        J = FindKey(fileData(i), UCase$(Key))
        
        'If we found it we return it
        If J >= 0 Then GetValue = fileData(i).values(J).value
    End If
End Function

''
'Searches for a given key within a given main node and returns the index in which it's stored or the negation of the index in which it should be if not found.
'
'   @param    Node The MainNode among whose value entries we will be searching.
'   @param    key The key of the value we are looking for.
'   @returns  The index in which the value with the key we are looking for is stored or the negation of the index in which it should be if not found.
'

Private Function FindKey(ByRef Node As MainNode, ByVal Key As String) As Long
'**************************************************************
'Author: Juan Mart�n Sotuyo Dodero
'Last Modify Date: 5/01/2006
'Returns the index of the value which key matches the requested one,
'or the negation of the position were it should be if not found
'**************************************************************
    Dim min As Long
    Dim max As Long
    Dim mid As Long
    
    min = 0
    max = Node.numValues - 1
    
    Do While min <= max
        mid = (min + max) \ 2
        
        If Node.values(mid).Key < Key Then
            min = mid + 1
        ElseIf Node.values(mid).Key > Key Then
            max = mid - 1
        Else
            'We found it
            FindKey = mid
            Exit Function
        End If
    Loop
    
    'Not found, return the negation of the position where it should be
    '(all higher values are to the right of the list and lower values are to the left)
    FindKey = Not mid
End Function

''
'Searches for a main section with the given name within the loaded INI file and returns the index in which it's stored or the negation of the index in which it should be if not found.
'
'   @param    name The name of the MainNode we are looking for.
'   @returns  The index in which the main section we are looking for is stored or the negation of the index in which it should be if not found.
'

Private Function FindMain(ByVal name As String) As Long
'**************************************************************
'Author: Juan Mart�n Sotuyo Dodero
'Last Modify Date: 5/01/2006
'Returns the index of the MainNode which name matches the requested one,
'or the negation of the position were it should be if not found
'**************************************************************
    Dim min As Long
    Dim max As Long
    Dim mid As Long
    
    min = 0
    max = MainNodes - 1
    
    Do While min <= max
        mid = (min + max) \ 2
        
        If fileData(mid).name < name Then
            min = mid + 1
        ElseIf fileData(mid).name > name Then
            max = mid - 1
        Else
            'We found it
            FindMain = mid
            Exit Function
        End If
    Loop
    
    'Not found, return the negation of the position where it should be
    '(all higher values are to the right of the list and lower values are to the left)
    FindMain = Not mid
End Function








