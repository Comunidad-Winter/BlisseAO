VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSurfaceManDyn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private Const BYTES_PER_MB As Long = 1048576                        '1Mb = 1024 Kb = 1024 * 1024 bytes = 1048576 bytes
Private Const MIN_MEMORY_TO_USE As Long = 256 * BYTES_PER_MB          '4 Mb
Private Const DEFAULT_MEMORY_TO_USE As Long = 512 * BYTES_PER_MB     '16 Mb

''Number of buckets in our hash table. Must be a nice prime number.
Const HASH_TABLE_SIZE As Long = 2000


Private Type SURFACE_ENTRY_DYN
    fileIndex As Long
    lastAccess As Long
    Surface As Direct3DTexture8
End Type

Private Type HashNode
    surfaceCount As Integer
    SurfaceEntry() As SURFACE_ENTRY_DYN
End Type

Private surfaceList(HASH_TABLE_SIZE - 1) As HashNode

Private DirectD3D As D3DX8

Private maxBytesToUse As Long
Private usedBytes As Long

Private useVideoMemory As Boolean

Private ResourcePath As String

Private Declare Function GetTickCount Lib "kernel32" () As Long

Private Const DIB_PAL_COLORS As Long = 1
Private Const DIB_RGB_COLORS As Long = 0


Private Sub Class_Initialize()
'**************************************************************
'Author: Juan Martín Sotuyo Dodero
'Last Modify Date: 3/06/2006
'
'**************************************************************
    usedBytes = 0
    maxBytesToUse = MIN_MEMORY_TO_USE
End Sub

Private Sub Class_Terminate()
'**************************************************************
'Author: Juan Martín Sotuyo Dodero
'Last Modify Date: 3/06/2006
'Clean up
'**************************************************************
    Dim i As Long
    Dim J As Long
    
    'Destroy every surface in memory
    For i = 0 To HASH_TABLE_SIZE - 1
        With surfaceList(i)
            For J = 1 To .surfaceCount
                Set .SurfaceEntry(J).Surface = Nothing
            Next J
            
            'Destroy the arrays
            Erase .SurfaceEntry
        End With
    Next i
End Sub

Public Sub Initialize(ByRef DD As D3DX8, ByVal videoMemory As Boolean, ByVal graphicPath As String, Optional ByVal maxMemoryUsageInMb As Long = -1)
'**************************************************************
'Author: Juan Martín Sotuyo Dodero
'Last Modify Date: 3/06/2006
'Initializes the manager
'**************************************************************
    Set DirectD3D = DD
    
    useVideoMemory = videoMemory
    ResourcePath = graphicPath
    
    If maxMemoryUsageInMb = -1 Then
        maxBytesToUse = DEFAULT_MEMORY_TO_USE   '   16 Mb by default
    ElseIf maxMemoryUsageInMb * BYTES_PER_MB < MIN_MEMORY_TO_USE Then
        maxBytesToUse = MIN_MEMORY_TO_USE       '   4 Mb is the minimum allowed
    Else
        maxBytesToUse = maxMemoryUsageInMb * BYTES_PER_MB
    End If
End Sub

Public Property Get Surface(ByVal fileIndex As Long) As Direct3DTexture8
'**************************************************************
'Author: Juan Martín Sotuyo Dodero
'Last Modify Date: 3/06/2006
'Retrieves the requested texture
'**************************************************************
    Dim i As Long
    
    '   Search the index on the list
    With surfaceList(fileIndex Mod HASH_TABLE_SIZE)
        For i = 1 To .surfaceCount
            If .SurfaceEntry(i).fileIndex = fileIndex Then
                .SurfaceEntry(i).lastAccess = GetTickCount
                Set Surface = .SurfaceEntry(i).Surface
                Exit Property
            End If
        Next i
    End With
    
    'Not in memory, load it!
    Set Surface = LoadSurface(fileIndex)
End Property

Private Function LoadSurface(ByVal fileIndex As Long) As Direct3DTexture8
'**************************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modify Date: 05/11/2009
'Loads the surface named fileIndex + ".bmp" and inserts it to the
'surface list in the listIndex position
'**************************************************************
On Error GoTo ErrHandler

    Dim newSurface As SURFACE_ENTRY_DYN
    Dim surface_desc As D3DSURFACE_DESC
    Dim texture_info As D3DXIMAGE_INFO
        
    With newSurface
        .fileIndex = fileIndex
        .lastAccess = GetTickCount
        
        If DX8_GreyScale = True Then
            If General_File_Exist(Resources.Graphics & CStr(fileIndex) & ".bmp", vbNormal) Then
                Set .Surface = DirectD3D.CreateTextureFromFileEx(DirectDevice, Resources.Graphics & CStr(fileIndex) & ".bmp", _
                        D3DX_DEFAULT, D3DX_DEFAULT, 3, 0, D3DFMT_A4L4, D3DPOOL_MANAGED, D3DX_FILTER_NONE, _
                        D3DX_FILTER_NONE, &HFF000000, texture_info, ByVal 0)
            Else
                Set .Surface = DirectD3D.CreateTextureFromFileEx(DirectDevice, Resources.Graphics & CStr(fileIndex) & ".png", _
                        D3DX_DEFAULT, D3DX_DEFAULT, 3, 0, D3DFMT_A4L4, D3DPOOL_MANAGED, D3DX_FILTER_NONE, _
                        D3DX_FILTER_NONE, &HFFFAFFAC, texture_info, ByVal 0)
            End If
        Else
            If General_File_Exist(Resources.Graphics & CStr(fileIndex) & ".bmp", vbNormal) Then
                Set .Surface = DirectD3D.CreateTextureFromFileEx(DirectDevice, Resources.Graphics & CStr(fileIndex) & ".bmp", _
                        D3DX_DEFAULT, D3DX_DEFAULT, 3, 0, D3DFMT_A8R8G8B8, D3DPOOL_MANAGED, D3DX_FILTER_NONE, _
                        D3DX_FILTER_NONE, &HFF000000, texture_info, ByVal 0)
            Else
                Set .Surface = DirectD3D.CreateTextureFromFileEx(DirectDevice, Resources.Graphics & CStr(fileIndex) & ".png", _
                        D3DX_DEFAULT, D3DX_DEFAULT, 3, 0, D3DFMT_A8R8G8B8, D3DPOOL_MANAGED, D3DX_FILTER_NONE, _
                        D3DX_FILTER_NONE, &HFFFAFFAC, texture_info, ByVal 0)
            End If
        End If

        newSurface.Surface.GetLevelDesc 0, surface_desc
    End With
    
    'Insert surface to the list
    With surfaceList(fileIndex Mod HASH_TABLE_SIZE)
        .surfaceCount = .surfaceCount + 1
        
        ReDim Preserve .SurfaceEntry(1 To .surfaceCount) As SURFACE_ENTRY_DYN
        
        .SurfaceEntry(.surfaceCount) = newSurface
        
        Set LoadSurface = newSurface.Surface
    End With
    
    'Update used bytes
    usedBytes = usedBytes + surface_desc.Size
    
    'Check if we have exceeded our allowed share of memory usage
    Do While usedBytes > maxBytesToUse
        'Remove a file. If no file could be removed we continue, if the file was previous to our surface we update the index
        If Not RemoveLRU() Then
            Exit Do
        End If
    Loop
    
Exit Function

ErrHandler:

End Function

Private Function RemoveLRU() As Boolean
'**************************************************************
'Author: Juan Martín Sotuyo Dodero
'Last Modify Date: 3/06/2006
'Removes the Least Recently Used surface to make some room for new ones
'**************************************************************
    Dim LRUi As Long
    Dim LRUj As Long
    Dim LRUtime As Long
    Dim i As Long
    Dim J As Long
    Dim ddsd As D3DSURFACE_DESC
    
    LRUtime = GetTickCount
    
    'Check out through the whole list for the least recently used
    For i = 0 To HASH_TABLE_SIZE - 1
        With surfaceList(i)
            For J = 1 To .surfaceCount
                If LRUtime > .SurfaceEntry(J).lastAccess Then
                    LRUi = i
                    LRUj = J
                    LRUtime = .SurfaceEntry(J).lastAccess
                End If
            Next J
        End With
    Next i
    
    If LRUj Then
        RemoveLRU = True
        'Retrieve the surface desc
        Call surfaceList(LRUi).SurfaceEntry(LRUj).Surface.GetLevelDesc(0, ddsd)
        
        'Remove it
        Set surfaceList(LRUi).SurfaceEntry(LRUj).Surface = Nothing
        surfaceList(LRUi).SurfaceEntry(LRUj).fileIndex = 0
        
        'Move back the list (if necessary)
        With surfaceList(LRUi)
            For J = LRUj To .surfaceCount - 1
                .SurfaceEntry(J) = .SurfaceEntry(J + 1)
            Next J
            
            .surfaceCount = .surfaceCount - 1
            If .surfaceCount Then
                ReDim Preserve .SurfaceEntry(1 To .surfaceCount) As SURFACE_ENTRY_DYN
            Else
                Erase .SurfaceEntry
            End If
        End With
        
        'Update the used bytes
        usedBytes = usedBytes - ddsd.Size
    End If
End Function

Public Function Reset_Surfaces()
Erase surfaceList

End Function
