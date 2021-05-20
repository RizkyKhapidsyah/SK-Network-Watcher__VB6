VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsAdaptors"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Option Explicit

Const MAX_IP = 15   'To make a buffer... i dont think you have more than 5 entries on your pc..

Private Type IPInfo
     dwAddr As Long   ' IP address
    dwIndex As Long '  interface index
    dwMask As Long ' subnet mask
    dwBCastAddr As Long ' broadcast address
    dwReasmSize  As Long ' assembly size
    unused1 As Integer ' not currently used
    unused2 As Integer '; not currently used
End Type

Private Type MIB_IPADDRTABLE
    dEntrys As Long   'number of entries in the table
    mIPInfo(MAX_IP) As IPInfo  'array of IP address entries
End Type

Dim realNIC As Integer ' count the total NIC
Private NewIPtable(MAX_IP) As IPInfo  'this table we shoud use in combination with te other class..
Dim realNIC_Table(MAX_IP) As Byte  'this table shows the ip to index..

Private IP_Table As MIB_IPADDRTABLE   'Table IP's settings data

 Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
 Private Declare Function GetIpAddrTable Lib "iphlpapi" (pIPAdrTable As Byte, pdwSize As Long, ByVal Sort As Long) As Long
'------------------------------------------------------------------------------------------------------------------
'converts a Long  to a string  =>  long to IP address
Private Function ConvertAddressToString(longAddr As Long) As String
    Dim myByte(3) As Byte
    Dim cnt As Long
    CopyMemory myByte(0), longAddr, 4
    For cnt = 0 To 3
        ConvertAddressToString = ConvertAddressToString + CStr(myByte(cnt)) + "."
    Next cnt
    ConvertAddressToString = Left$(ConvertAddressToString, Len(ConvertAddressToString) - 1)
End Function
'-------------------------------------------------------------------------------------------------------------------

Private Sub Class_Initialize()
Dim Ret As Long         'used for GetIPAddrTable
Dim bBytes() As Byte   ' array for retrieving the data from..
Dim tel As Integer  ' counter for loop


On Error GoTo END1
    GetIpAddrTable ByVal 0&, Ret, True

    If Ret <= 0 Then Exit Sub
    ReDim bBytes(0 To Ret - 1) As Byte
    'retrieve the data
    GetIpAddrTable bBytes(0), Ret, True
      
    'Get the first 4 bytes to get the entry's.. ip installed
    CopyMemory IP_Table.dEntrys, bBytes(0), 4
    'Now in listing stands the entry's to the table
    For tel = 0 To IP_Table.dEntrys - 1   'starts at index(0)
        'array starts a ZERO !!!
        'This part will copy the data to the IP defined array... (double indexes will be copied)
        CopyMemory IP_Table.mIPInfo(tel), bBytes(4 + (tel * Len(IP_Table.mIPInfo(0)))), Len(IP_Table.mIPInfo(tel))
    Next
    'Now there is the possibility that there are more IP's to one NIC coupled.. so..
    ' we have first check that part by searching to duplicate indexes..
    MakeRealNIC_Table
    'we have to redim our new IP table.. that we want to use in the application..
    
    
Exit Sub
END1:
MsgBox "ERROR reading IP Table"
End Sub

Public Property Get Count_adaptors() As Integer
'Check if there are multiple IP's pointed to a NIC
    
    Count_adaptors = UpdateAdaptorCount
End Property

Public Property Get Adaptor_IP(ByVal Index As Integer) As String
'Index = Index - 1
If Index < 0 Then
    MsgBox "Adaptor number(index) is too low ! ", vbCritical
    Exit Property
End If
If Index > IP_Table.dEntrys - 1 Then
    MsgBox "Adaptor number(index) is too big ! ", vbCritical
Else
     Adaptor_IP = ConvertAddressToString(NewIPtable(Index).dwAddr)
     
    End If
End Property

Public Property Get Adaptor_Subnet(ByVal Index As Integer) As String

If Index < 0 Then
    MsgBox "Adaptor number(index) is too low ! ", vbCritical
    Exit Property
End If
If Index > IP_Table.dEntrys - 1 Then
    MsgBox "Adaptor number(index) is too big ! ", vbCritical
Else
    Adaptor_Subnet = ConvertAddressToString(NewIPtable(Index).dwMask)
End If
End Property

Public Property Get Adaptor_BroadCast(ByVal Index As Integer) As String

If Index < 0 Then
    MsgBox "Adaptor number(index) is too low ! ", vbCritical
    Exit Property
End If
If Index > realNIC - 1 Then
    MsgBox "Adaptor number(index) is too big ! ", vbCritical
Else
    Adaptor_BroadCast = ConvertAddressToString(NewIPtable(Index).dwBCastAddr)
End If
End Property

Public Property Get T_index(ByVal index1 As Long) As Long

    T_index = NewIPtable(index1).dwIndex
End Property


'==================== NEW  detection for multiple IP's on 1 NIC ============
' thanx to Mike  Barber for helping this out .. testing person . !
'----------------------------------------------------------------------------------------------------------------
Private Sub MakeRealNIC_Table()
Dim cnt, cnt2 As Integer
Dim TotalIP As Integer
'idea..   (on pc you have 2 NIC's and you have 3 IP's coupled to one NIC)
'if you has example 3 NIC (2 real - 1 localhost) the IP_Table would be..
'|--------------------------------------------------------------------------|
'|196.12.2.1 | 212.12.1.1 | 196.12.2.2 | 196.12.2.4 | 127.0.0.1|  = IP
'|    1231            1111           1231           1231            1          |  => index
'|----------------------------------------------------------------------|

'This Table will be filled by the index and count so the real NIC's
'output to RealNIc table
'------------------------------------------------------------------------'
'|     1         |       2           |       1           |       1         |      3          |  => count real nic's & get now place dupl indexes
'-------------------------------------------------------------------------
'clears the input buffer to 0
For cnt = 0 To MAX_IP
    realNIC_Table(cnt) = 0
Next

realNIC = 1    'start with 1  (there is always localhost)
TotalIP = IP_Table.dEntrys - 1

For cnt = 0 To TotalIP
    'If cnt = TotalIP Then Exit For
    If realNIC_Table(cnt) = 0 Then   'check it the array() is zero.
        realNIC_Table(cnt) = realNIC
            For cnt2 = (cnt + 1) To TotalIP
                'This will check by index.. if the index is the same => Same NIC !
                    If IP_Table.mIPInfo(cnt).dwIndex = IP_Table.mIPInfo(cnt2).dwIndex Then
                        realNIC_Table(cnt2) = realNIC     'if second same index = found => put it there
                    End If
            Next
            realNIC = realNIC + 1
    End If
Next
realNIC = realNIC - 1
'now we have counted the real NIC's and we can find it in the table where it is..
For cnt = 1 To realNIC   'last one has counted
    For cnt2 = 0 To TotalIP
        If realNIC_Table(cnt2) = cnt Then
            NewIPtable(cnt - 1) = IP_Table.mIPInfo(cnt2)
            Exit For
        End If
    Next
Next
'Now the new table is filled correctly with the right true NIC's
End Sub

Function UpdateAdaptorCount() As Integer
Dim Ret As Long         'used for GetIPAddrTable
Dim bBytes() As Byte   ' array for retrieving the data from..
Dim tel As Integer  ' counter for loop


On Error GoTo END1
    GetIpAddrTable ByVal 0&, Ret, True

    If Ret <= 0 Then Exit Function
    ReDim bBytes(0 To Ret - 1) As Byte
    'retrieve the data
    GetIpAddrTable bBytes(0), Ret, True
      
    'Get the first 4 bytes to get the entry's.. ip installed
    CopyMemory IP_Table.dEntrys, bBytes(0), 4
    'Now in listing stands the entry's to the table
    'check the Real NIC's
    MakeRealNIC_Table
    UpdateAdaptorCount = realNIC
END1:
End Function
