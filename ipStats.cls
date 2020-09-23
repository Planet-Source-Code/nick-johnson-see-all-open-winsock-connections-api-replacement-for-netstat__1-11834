VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ipStats"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
'Gives stats on IP connections - TCP, UDP and ICMP.
'This code is copyright 2000 Nick Johnson.
'This code may be reused and modified for non-commercial
'purposes only as long as credit is given to the author
'in the programmes about box and it's documentation.
'If you use this code, please email me at:
'arachnid@mad.scientist.com and let me know what you think
'and what you are doing with it.

'Winapi calls
Private Declare Function GetTcpTable Lib "IPhlpAPI" (pTcpTable As MIB_TCPTABLE, pdwSize As Long, bOrder As Long) As Long
Private Declare Function FormatMessage Lib "kernel32" Alias "FormatMessageA" (ByVal dwFlags As Long, lpSource As Any, ByVal dwMessageId As Long, ByVal dwLanguageId As Long, ByVal lpBuffer As String, ByVal nSize As Long, Arguments As Long) As Long

'Winapi structures
Private Type MIB_TCPROW
  dwState As tcpStates
  dwLocalAddr(0 To 3) As Byte
  dwLocalPort As String * 4
  dwRemoteAddr(0 To 3) As Byte
  dwRemotePort As String * 4
End Type

Private Type MIB_TCPTABLE
  dwNumEntries As Long    'number of entries in the table
  table(100) As MIB_TCPROW   'array of TCP connections
End Type

'Property variables
Private trRows() As tcpRow

'Function to get active & listening TCP connections
Public Function getTCPConnections() As Boolean
    Dim lngSize As Long
    Dim lngReturn As Long
    Dim tcpTable As MIB_TCPTABLE
    Dim intCount As Integer
    
    'The size of the tcpTable structure
    lngSize = 20 * 100 + 4
    lngReturn = GetTcpTable(tcpTable, lngSize, True)
    Select Case lngReturn
    Case 0
        'Success - copy structure to the array of ipStats
        ReDim trRows(0 To tcpTable.dwNumEntries - 1)
        For intCount = 0 To tcpTable.dwNumEntries - 1
            Set trRows(intCount) = New tcpRow
            trRows(intCount).LocalIP = tcpTable.table(intCount).dwLocalAddr
            trRows(intCount).RemoteIP = tcpTable.table(intCount).dwRemoteAddr
            trRows(intCount).LocalPort = c_port(tcpTable.table(intCount).dwLocalPort)
            If tcpTable.table(intCount).dwState = TCP_STATE_ESTAB Then
                trRows(intCount).RemotePort = c_port(tcpTable.table(intCount).dwRemotePort)
            Else
                trRows(intCount).RemotePort = 0
            End If
            trRows(intCount).State = tcpTable.table(intCount).dwState
        Next intCount
        getTCPConnections = True
    Case 122
        'Structure too small - can't gather. Solution on the way?
        getTCPConnections = False
    Case 232
        'Not connected to a network. Return an empty array and set connected to false.
        ReDim trRows(0 To 0)
        getTCPConnections = True
    Case Else
        'Unknown error. Get the message and report it
        Err.Raise 1 + vbObjectError, "ipStats", "Error getting connections:" & vbCrLf & "Number: " & Str(lngReturn) & vbCrLf & "Description: " & getErrorMessage(lngReturn)
        getTCPConnections = False
    End Select
End Function

Public Property Get RowData(index As Integer) As tcpRow
    Set RowData = trRows(index)
End Property

Public Property Get RowCount() As Integer
    RowCount = UBound(trRows) - LBound(trRows) + 1
End Property

'Retrieves the windows error message for a specific code
Private Function getErrorMessage(lngError As Long)
    Dim lngLen As Long
    Dim strOut As String
    
    strOut = Space(256)
    lngLen = FormatMessage(&H1000, 0, lngError, 0, strOut, 255, 0)
    getErrorMessage = Left(strOut, lngLen - 1)
End Function

'Extracts the port number
Private Function c_port(s) As Long
    c_port = Asc(Mid(s, 1, 1))
    c_port = c_port * 256
    c_port = c_port + Asc(Mid(s, 2, 1))
    'c_port = Asc(Mid(s, 1, 1)) * 256 + Asc(Mid(s, 2, 1))
End Function
