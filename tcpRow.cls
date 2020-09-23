VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "tcpRow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
'This code is copyright 2000 Nick Johnson.
'This code may be reused and modified for non-commercial
'purposes only as long as credit is given to the author
'in the programmes about box and it's documentation.
'If you use this code, please email me at:
'arachnid@mad.scientist.com and let me know what you think
'and what you are doing with it.

'Winapi declarations
Private Declare Function SetTcpEntry Lib "IPhlpAPI" (pTcpRow As MIB_TCPROW) As Long

'property variables
Private pbLocalIP(0 To 3) As Byte
Private pbRemoteIP(0 To 3) As Byte
Private plngLocalPort As Long
Private plngRemotePort As Long
Private pState As tcpStates

'Winapi structures
Private Type MIB_TCPROW
  dwState As tcpStates
  dwLocalAddr(0 To 3) As Byte
  dwLocalPort As String * 4
  dwRemoteAddr(0 To 3) As Byte
  dwRemotePort As String * 4
End Type

'Enums
Public Enum tcpStates
    TCP_STATE_CLOSED = 1
    TCP_STATE_LISTEN = 2
    TCP_STATE_SYN_SENT = 3
    TCP_STATE_SYN_RCVD = 4
    TCP_STATE_ESTAB = 5
    TCP_STATE_FIN_WAIT1 = 6
    TCP_STATE_FIN_WAIT2 = 7
    TCP_STATE_CLOSE_WAIT = 8
    TCP_STATE_CLOSING = 9
    TCP_STATE_LAST_ACK = 10
    TCP_STATE_TIME_WAIT = 11
    TCP_STATE_DELETE_TCB = 12
End Enum

Friend Property Let State(ByVal vData As tcpStates)
    pState = vData
End Property
Public Property Get State() As tcpStates
    State = pState
End Property

Public Property Get StateText() As String
    Select Case pState
    Case TCP_STATE_CLOSED
        StateText = "Closed"
    Case TCP_STATE_LISTEN
        StateText = "Listening"
    Case TCP_STATE_SYN_SENT
        StateText = "SYN Sent"
    Case TCP_STATE_SYN_RCVD
        StateText = "SYN Recieved"
    Case TCP_STATE_ESTAB
        StateText = "Established"
    Case TCP_STATE_FIN_WAIT1
        StateText = "FIN Wait 1"
    Case TCP_STATE_FIN_WAIT2
        StateText = "FIN Wait 2"
    Case TCP_STATE_CLOSE_WAIT
        StateText = "Close Wait"
    Case TCP_STATE_CLOSING
        StateText = "Closing"
    Case TCP_STATE_LAST_ACK
        StateText = "Last ACK"
    Case TCP_STATE_TIME_WAIT
        StateText = "Time Wait"
    Case TCP_STATE_DELETE_TCB
        StateText = "PCB Deleted"
    End Select
End Property

Friend Property Let RemotePort(ByVal vData As Long)
    plngRemotePort = vData
End Property
Public Property Get RemotePort() As Long
    RemotePort = plngRemotePort
End Property

Friend Property Let LocalPort(ByVal vData As Long)
    plngLocalPort = vData
End Property
Public Property Get LocalPort() As Long
    LocalPort = plngLocalPort
End Property

Friend Property Let RemoteIP(vData As Variant)
    'vData is variant due to VB's stupid array rules.
    Dim a As Byte
    For a = 0 To 3
        pbRemoteIP(a) = vData(a)
    Next a
End Property

Public Property Get RemoteIPString() As String
    'Returns the remote IP address as a string
    RemoteIPString = bStr(pbRemoteIP(0)) & "." & bStr(pbRemoteIP(1)) & "." & bStr(pbRemoteIP(2)) & "." & bStr(pbRemoteIP(3))
End Property

Friend Property Let LocalIP(vData As Variant)
    'vData is variant due to VB's stupid array rules.
    Dim a As Byte
    For a = 0 To 3
        pbLocalIP(a) = vData(a)
    Next a
End Property

Public Property Get LocalIPString() As String
    LocalIPString = bStr(pbLocalIP(0)) & "." & bStr(pbLocalIP(1)) & "." & bStr(pbLocalIP(2)) & "." & bStr(pbLocalIP(3))
End Property

Public Sub Kill()
    'Kills this connection.
    Dim tcpThisRow As MIB_TCPROW
    
    tcpThisRow.dwLocalAddr(0) = pbLocalIP(0)
    tcpThisRow.dwLocalAddr(1) = pbLocalIP(1)
    tcpThisRow.dwLocalAddr(2) = pbLocalIP(2)
    tcpThisRow.dwLocalAddr(3) = pbLocalIP(3)
    
    tcpThisRow.dwRemoteAddr(0) = pbRemoteIP(0)
    tcpThisRow.dwRemoteAddr(1) = pbRemoteIP(1)
    tcpThisRow.dwRemoteAddr(2) = pbRemoteIP(2)
    tcpThisRow.dwRemoteAddr(3) = pbRemoteIP(3)
    
    tcpThisRow.dwState = TCP_STATE_DELETE_TCB
    
    tcpThisRow.dwLocalPort = lngToPort(plngLocalPort)
    tcpThisRow.dwRemotePort = lngToPort(plngRemotePort)
    
    SetTcpEntry tcpThisRow
End Sub

Private Function bStr(valIn As Byte) As String
    bStr = Right(Str(valIn), Len(Str(valIn)) - 1)
End Function

Private Function lngToPort(ByVal lngIn As Long) As String
    'Converts a long port number into a C style port number
    Dim lngTemp As Long
    
    lngTemp = Int(lngIn / 256)
    lngIn = lngIn - (256 * lngTemp)
    lngToPort = Chr(lngTemp) & Chr(lngIn) & Chr(0) & Chr(0)
End Function
