Attribute VB_Name = "Mod_Security_MAC"
'***************************************************
'Author: Ezequiel Juárez (Standelf)
'Last Modification: 26/05/10
'Blisse-AO | Black And White AO | Security General MAC Address.
'***************************************************

Option Explicit

Private Declare Sub CopyMemory Lib "Kernel32" Alias "RtlMoveMemory" (hpvDest As Any, hpvSource As Any, ByVal cbCopy As Long)

Private Declare Function GetNetworkParams Lib "IPHlpApi.dll" _
       (FixedInfo As Any, pOutBufLen As Long) As Long
Private Declare Function GetAdaptersInfo Lib "IPHlpApi.dll" _
       (IpAdapterInfo As Any, pOutBufLen As Long) As Long

Private Const MAX_ADAPTER_NAME_LENGTH = 260
Private Const MAX_ADAPTER_DESCRIPTION_LENGTH = 132
Private Const MAX_ADAPTER_ADDRESS_LENGTH = 8
Private Type IP_ADDR_STRING
    Next As Long
    IpAddress As String * 16
    IpMask As String * 16
    Context As Long
End Type

Private Const ERROR_BUFFER_OVERFLOW = 111

Private Const MIB_IF_TYPE_ETHERNET = 6
Private Const MIB_IF_TYPE_TOKENRING = 9
Private Const MIB_IF_TYPE_FDDI = 15
Private Const MIB_IF_TYPE_PPP = 23
Private Const MIB_IF_TYPE_LOOPBACK = 24
Private Const MIB_IF_TYPE_SLIP = 28

Private Type IP_ADAPTER_INFO
    Next As Long
    ComboIndex As Long
    AdapterName As String * MAX_ADAPTER_NAME_LENGTH
    Description As String * MAX_ADAPTER_DESCRIPTION_LENGTH
    AddressLength As Long
    Address(MAX_ADAPTER_ADDRESS_LENGTH - 1) As Byte
    Index As Long
    Type As Long
    DhcpEnabled As Long
    CurrentIpAddress As Long
    IpAddressList As IP_ADDR_STRING
    GatewayList As IP_ADDR_STRING
    DhcpServer As IP_ADDR_STRING
    HaveWins As Byte
    PrimaryWinsServer As IP_ADDR_STRING
    SecondaryWinsServer As IP_ADDR_STRING
    LeaseObtained As Long
    LeaseExpires As Long
End Type

Public Function Get_MAC_Address() As String
Dim AdapterInfoSize As Long
Dim Error As Long
Dim AdapterInfoBuffer() As Byte
Dim AdapterInfo As IP_ADAPTER_INFO
Dim pAdapt As Long
Dim Buffer2 As IP_ADAPTER_INFO
Dim i As Long
Dim tmpStr As String

AdapterInfoSize = 0
Error = GetAdaptersInfo(ByVal 0&, AdapterInfoSize)

If Error <> 0 Then
    If Error <> ERROR_BUFFER_OVERFLOW Then Exit Function
End If

ReDim AdapterInfoBuffer(AdapterInfoSize - 1)

' Get actual adapter information
Error = GetAdaptersInfo(AdapterInfoBuffer(0), AdapterInfoSize)
If Error <> 0 Then Exit Function

' Allocate memory
CopyMemory AdapterInfo, AdapterInfoBuffer(0), AdapterInfoSize
pAdapt = AdapterInfo.Next

Do
    CopyMemory Buffer2, AdapterInfo, AdapterInfoSize
    
    If Buffer2.Type = MIB_IF_TYPE_PPP Or Buffer2.Type = MIB_IF_TYPE_ETHERNET Then
        For i = 0 To Buffer2.AddressLength - 1
            tmpStr = tmpStr & Hex(Buffer2.Address(i))
            If i < Buffer2.AddressLength - 1 Then
               tmpStr = tmpStr & "-"
            End If
        Next
        
        Get_MAC_Address = tmpStr
        Exit Function
    End If
    
    pAdapt = Buffer2.Next
    If pAdapt <> 0 Then CopyMemory AdapterInfo, ByVal pAdapt, AdapterInfoSize

Loop Until pAdapt = 0
      
End Function
