Attribute VB_Name = "Mod_Security_Intervals"
'***************************************************
'Author: Ezequiel Juárez (Standelf)
'Last Modification: 15/05/10 | 21/05/10
'Blisse-AO | Security Intervals.
'***************************************************
Public Type typeInterval
    TimeCheck As Integer
    LastCheck As Long
    CanUse As Boolean
End Type

Public Enum dInter
    DragDrop = 1
    GlobalChat = 2
    Drop = 3
    Command = 4
    SeguroItems = 5
End Enum

Public Intervalos() As typeInterval
Private NumIntervalos As Byte
Private i As Byte

Public Sub Init_Intervals()
'***************************************************
'Author: Standelf
'Last Modification: 21/05/10
'Init and Create Intervals
'***************************************************
    NumIntervalos = 5
    
    If NumIntervalos <= 0 Then Exit Sub
    ReDim Preserve Intervalos(1 To NumIntervalos) As typeInterval
    
    Intervalos(dInter.DragDrop).TimeCheck = 400
    Intervalos(dInter.GlobalChat).TimeCheck = 600
    Intervalos(dInter.Drop).TimeCheck = 800
    Intervalos(dInter.Command).TimeCheck = 1000
    Intervalos(dInter.SeguroItems).TimeCheck = 6000
    
End Sub

Public Sub Updates_Intervals()
'***************************************************
'Author: Standelf
'Last Modification: 21/05/10
'Update and enabled functions
'***************************************************
    For i = 1 To NumIntervalos
        With Intervalos(i)
            If .CanUse = False Then
                If GetTickCount - .LastCheck > .TimeCheck Then
                    .CanUse = True
                    .LastCheck = GetTickCount
                    
                    If i = dInter.SeguroItems And Settings.SeguroItems Then
                        Items_Seg = .CanUse
                        Call frmMain.ControlSM(dInter.SeguroItems, Items_Seg)
                    End If
                End If
            End If
        End With
    Next i
End Sub

Public Function CanUse_Interval(ByVal Index As Byte) As Boolean
'***************************************************
'Author: Standelf
'Last Modification: 21/05/10
'Check the State
'***************************************************
        With Intervalos(Index)
            If .CanUse = False Then 'No puede usar
                With FontTypes(FontTypeNames.FONTTYPE_INFO)
                    Call ShowConsoleMsg("No puedes realizar esta acción tan rápido.", .Red, .Green, .Blue, .bold, .italic)
                    CanUse_Interval = False
                End With
            Else 'Puede Usar
                .CanUse = False
                .LastCheck = GetTickCount
                CanUse_Interval = True
            End If
        End With
End Function
