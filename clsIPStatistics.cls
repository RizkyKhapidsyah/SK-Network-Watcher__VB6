VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsIPStatistics"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const ERROR_NOT_SUPPORTED = 50&
Private Const ERROR_SUCCESS = 0&

'------------------------------------ Status of the adaptor ---------------------------------------------
Public Enum OperationalStates
    MIB_IF_OPER_STATUS_NON_OPERATIONAL = 0
    MIB_IF_OPER_STATUS_UNREACHABLE = 1
    MIB_IF_OPER_STATUS_DISCONNECTED = 2
    MIB_IF_OPER_STATUS_CONNECTING = 3
    MIB_IF_OPER_STATUS_CONNECTED = 4
    MIB_IF_OPER_STATUS_OPERATIONAL = 5
End Enum
'--------------------------------------------------------------------------------------------------------------------
'----------------------------------- Type of interface ---------------------------------------------------------
Public Enum InterfaceTypes
    MIB_IF_TYPE_OTHER = 1
    MIB_IF_TYPE_ETHERNET = 6
    MIB_IF_TYPE_TOKENRING = 9
    MIB_IF_TYPE_FDDI = 15
    MIB_IF_TYPE_PPP = 23
    MIB_IF_TYPE_LOOPBACK = 24
    MIB_IF_TYPE_SLIP = 28
End Enum
'------------------------------------------ admin status interface ---------------------------------------------------------
Public Enum AdminStatuses
    MIB_IF_ADMIN_STATUS_UP = 1
    MIB_IF_ADMIN_STATUS_DOWN = 2
    MIB_IF_ADMIN_STATUS_TESTING = 3
End Enum

Private Const MAXLEN_IFDESCR = 256
Private Const MAXLEN_PHYSADDR = 8
Private Const MAX_INTERFACE_NAME_LEN = 256


Private Type MIB_IFROW
    wszName(0 To 511) As Byte
    dwIndex As Long             '// index of the interface
    dwType As Long              '// type of interface
    dwMtu As Long               '// max transmission unit
    dwSpeed As Long             '// speed of the interface
    dwPhysAddrLen As Long       '// length of physical address
    bPhysAddr(0 To 7) As Byte   '// physical address of adapter
    dwAdminStatus As Long       '// administrative status
    dwOperStatus As Long        '// operational status
    dwLastChange As Long        '// last time operational status changed
    dwInOctets As Long          '// octets received
    dwInUcastPkts As Long       '// unicast packets received
    dwInNUcastPkts As Long      '// non-unicast packets received
    dwInDiscards As Long        '// received packets discarded
    dwInErrors As Long          '// erroneous packets received
    dwInUnknownProtos As Long   '// unknown protocol packets received
    dwOutOctets As Long         '// octets sent
    dwOutUcastPkts As Long      '// unicast packets sent
    dwOutNUcastPkts As Long     '// non-unicast packets sent
    dwOutDiscards As Long       '// outgoing packets discarded
    dwOutErrors As Long         '// erroneous packets sent
    dwOutQLen As Long           '// output queue length
    dwDescrLen As Long          '// length of bDescr member
    bDescr(0 To 255) As Byte    '// interface description
End Type

Private Type Adaptorinfo
        TotRecieved As Long
        TotSends As Long
        ConnectionSpeed As Long
        This_IP As String
        This_Subnet As String
        This_BroadCast As String
        InterfaceType As String * 15
        ConnectionState As String * 20
        ErrorPacketsRecieved As Long
        ErrorPacketsSend As Long
        UnknownProtocolsRecieved As Long
        Index As Long
End Type

Private Declare Function GetIfTable Lib "iphlpapi" (ByRef pIfRowTable As Any, ByRef pdwSize As Long, ByVal bOrder As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByRef pDest As Any, ByRef pSource As Any, ByVal Length As Long)

Dim lngRetVal As Long
Dim lngSize As Long
Dim m_Adaptor As Long
Dim m_CntAdaptors As Long
Dim IfRowTable As MIB_IFROW

Enum Baudrate
    Bits = 1
    Kbits = 1000
    Mbits = 1000000
End Enum

Dim MyAdaptor() As Adaptorinfo
Dim Adaptorinfo As New clsAdaptors

Dim cntIndexedAdaptors As Integer 'count the adaptors found by index check..

Public Event AdaptorChanged(ByVal newCount As Integer)
Public Event AdaptorError(ByVal ErrorDiscription As Integer)

Private Sub Class_Initialize()
Dim X() As Byte
Dim IPRows As Long
Dim cnt As Long

lngRetVal = GetIfTable(ByVal 0&, lngSize, 0)
ReDim X(0 To lngSize - 1) As Byte
     If lngRetVal = ERROR_NOT_SUPPORTED Then
        '
        'This API works only on Win 98/2000 and NT4 with SP4
        MsgBox "IP Helper is not supported by this system."
  End If
 'Search for the table - rows ..adaptors founded in iplhapi..
  lngRetVal = GetIfTable(X(0), lngSize, True) 'True => Sorted..
  If lngRetVal = ERROR_SUCCESS Then
        '
        'The first 4 bytes (the Long value) contain the quantity of the table rows
        'Get that value into the lngRows variable
        CopyMemory IPRows, X(0), 4
        'Max IP is not NIC  found on local system..
        'm_CntAdaptors = IPRows ' - DetectMultipleIPtoNIC
        m_CntAdaptors = Adaptorinfo.Count_adaptors
        
       'Redimension MyAdaptor
        ReDim MyAdaptor(1 To m_CntAdaptors) As Adaptorinfo
        'Search for the IP - subnet - ect... thru indexes..
        'LinkIt
        Call Update_Adaptors_Stat  ' set the info about adaptor to the right adaptor ip,subnet....
        'm_Adaptor = 1 'Set the first adaptor .(local host..127.0.0.1?)
       
  End If
  End Sub
Public Sub Update_Adaptors_Stat()
'This will link the adaptors (ip.subnet..) by index to this class
Dim cnt As Long
Dim CntAd As Long
Dim X() As Byte
Dim IPRows As Long

ReDim X(0 To lngSize - 1) As Byte

lngRetVal = GetIfTable(X(0), lngSize, True)

'Look if new adaptors found !!!!!!!!!!!!!! ..............
        CopyMemory IPRows, X(0), 4
        'Max Adaptors found on local system...
        m_CntAdaptors = Adaptorinfo.Count_adaptors
         
         ReDim MyAdaptor(1 To m_CntAdaptors) As Adaptorinfo

For cnt = 1 To (m_CntAdaptors)
    CopyMemory IfRowTable, X(4 + (cnt - 1) * Len(IfRowTable)), Len(IfRowTable)
 ' there can be a problem with some NIC's.. 1 NIC & more IP going to it..
 ' After testing with persons that have this ...
 ' so now i have to detect that part also , the index remains the same value...
    For CntAd = 0 To (m_CntAdaptors - 1)
          If IfRowTable.dwIndex = Adaptorinfo.T_index(CntAd) Then
          'count 1 founded adaptor to list..
               MyAdaptor(cnt).This_IP = Adaptorinfo.Adaptor_IP(CntAd)
               'Check if  This ip = localhost (127.0.0.1)
               If MyAdaptor(cnt).This_IP = "127.0.0.1" Then
                    MyAdaptor(cnt).This_IP = MyAdaptor(cnt).This_IP & " (localhost)"
               End If
               MyAdaptor(cnt).This_Subnet = Adaptorinfo.Adaptor_Subnet(CntAd)
               MyAdaptor(cnt).This_BroadCast = Adaptorinfo.Adaptor_BroadCast(CntAd)
               MyAdaptor(cnt).ConnectionSpeed = IfRowTable.dwSpeed
               MyAdaptor(cnt).TotRecieved = IfRowTable.dwInOctets
               MyAdaptor(cnt).TotSends = IfRowTable.dwOutOctets
               MyAdaptor(cnt).InterfaceType = InterfaceType(IfRowTable.dwType)
               MyAdaptor(cnt).ConnectionState = OperationalState(IfRowTable.dwOperStatus)
               MyAdaptor(cnt).ErrorPacketsRecieved = IfRowTable.dwInErrors
               MyAdaptor(cnt).ErrorPacketsSend = IfRowTable.dwOutErrors
               MyAdaptor(cnt).UnknownProtocolsRecieved = IfRowTable.dwInUnknownProtos
          End If
          
    Next
       
Next

End Sub
  

Public Property Get Connection_Speed(Optional speed As Baudrate = 1) As Double
On Error GoTo Err1
            Connection_Speed = Round(((MyAdaptor(m_Adaptor).ConnectionSpeed / CLng(speed))), 2)
Err1:
End Property

Public Property Get BytesRecieved() As Long
On Error GoTo Err1
            BytesRecieved = MyAdaptor(m_Adaptor).TotRecieved
Err1:
End Property

Public Property Get BytesSends() As Long
On Error GoTo Err1
            BytesSends = MyAdaptor(m_Adaptor).TotSends
Err1:
End Property

Public Property Get ErrorPacketsRcv() As Long
On Error GoTo Err1
     ErrorPacketsRcv = MyAdaptor(m_Adaptor).ErrorPacketsRecieved
Err1:
End Property

Public Property Get Found_Adaptors() As Long
On Error GoTo Err1
    Found_Adaptors = m_CntAdaptors '- DetectMultipleIPtoNIC
Err1:
End Property

Public Property Get Choose_Adaptor() As Long
        Choose_Adaptor = m_Adaptor
End Property

Public Property Let Choose_Adaptor(ByVal Adaptor As Long)
    If Adaptor < 1 Then
        MsgBox "adaptor To low "
        m_Adaptor = 1
        Exit Property
    ElseIf Adaptor > m_CntAdaptors Then
        'Raise an ErrorEvent .. adaptor not found
        RaiseEvent AdaptorError(1)
        'Scan the register again ..count the adaptors ...
        Exit Property
        'm_Adaptor = m_CntAdaptors
        Exit Property
    End If
    
    m_Adaptor = Adaptor
End Property


Public Property Get Local_IP() As String
On Error GoTo Err1
    Local_IP = MyAdaptor(m_Adaptor).This_IP
Err1:
End Property

Public Property Get Subnet() As String
On Error GoTo Err1
    Subnet = MyAdaptor(m_Adaptor).This_Subnet
Err1:
End Property

Public Property Get BroadcastAddr() As String
On Error GoTo Err1
    BroadcastAddr = MyAdaptor(m_Adaptor).This_BroadCast
Err1:
End Property

Public Property Get Interface_Type() As String
On Error GoTo Err1
    Interface_Type = MyAdaptor(m_Adaptor).InterfaceType
Err1:
End Property

Private Function InterfaceType(X As InterfaceTypes) As String
Select Case X
    Case 1
        InterfaceType = "Other"
    Case 6
        InterfaceType = "Ethernet"
    Case 9
        InterfaceType = "Tokenring"
    Case 15
        InterfaceType = "FDDI"
    Case 23
        InterfaceType = "PPP"
    Case 24
        InterfaceType = "Loopback"
    Case 28
        InterfaceType = "Slip"
End Select
End Function

Private Function OperationalState(X As OperationalStates) As String
Select Case X
    Case 0
        OperationalState = "Not operational"
    Case 5
       OperationalState = "Unreachable"
    Case 2
       OperationalState = "Disconnected"
    Case 3
      OperationalState = "Connecting"
    Case 4
      OperationalState = "Connected"
    Case 1
      OperationalState = "Operational"
End Select
End Function

Public Property Get OperationStatus() As Variant
On Error GoTo Err1
       OperationStatus = MyAdaptor(m_Adaptor).ConnectionState
Err1:
End Property

Public Property Get UnknowProtocolsRvc() As Long
On Error GoTo Err1
    UnknowProtocolsRvc = MyAdaptor(m_Adaptor).UnknownProtocolsRecieved
Err1:
End Property


Public Property Get ErrorPacketsSnd() As Long
On Error GoTo Err1
    ErrorPacketsSnd = MyAdaptor(m_Adaptor).ErrorPacketsSend
Err1:
End Property


