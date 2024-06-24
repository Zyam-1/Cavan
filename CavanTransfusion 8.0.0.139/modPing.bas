Attribute VB_Name = "modPing"
Option Explicit

Public Const SOCKET_ERROR = 0

Public Type WSAdata
    wVersion As Integer
    wHighVersion As Integer
    szDescription(0 To 255) As Byte
    szSystemStatus(0 To 128) As Byte
    iMaxSockets As Integer
    iMaxUdpDg As Integer
    lpVendorInfo As Long
End Type

Public Type Hostent
    h_name As Long
    h_aliases As Long
    h_addrtype As Integer
    h_length As Integer
    h_addr_list As Long
End Type

Public Type IP_OPTION_INFORMATION
    TTL As Byte
    Tos As Byte
    Flags As Byte
    OptionsSize As Long
    OptionsData As String * 128
End Type

Public Type IP_ECHO_REPLY
    Address(0 To 3) As Byte
    Status As Long
    RoundTripTime As Long
    DataSize As Integer
    Reserved As Integer
    data As Long
    Options As IP_OPTION_INFORMATION
End Type

Public Declare Function GetHostByName Lib "wsock32.dll" Alias "gethostbyname" (ByVal HostName As String) As Long
Public Declare Function WSAStartup Lib "wsock32.dll" (ByVal wVersionRequired&, lpWSAdata As WSAdata) As Long
Public Declare Function WSACleanup Lib "wsock32.dll" () As Long
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (hpvDest As Any, hpvSource As Any, ByVal cbCopy As Long)
Public Declare Function IcmpCreateFile Lib "icmp.dll" () As Long
Public Declare Function IcmpCloseHandle Lib "icmp.dll" (ByVal HANDLE As Long) As Boolean
Public Declare Function IcmpSendEcho Lib "ICMP" (ByVal IcmpHandle As Long, ByVal DestAddress As Long, ByVal RequestData As String, ByVal RequestSize As Integer, RequestOptns As IP_OPTION_INFORMATION, ReplyBuffer As IP_ECHO_REPLY, ByVal ReplySize As Long, ByVal TimeOut As Long) As Boolean

Public Function Ping(ByVal HostName As String) As Boolean
      'Returns True if success

      Dim hFile As Long, lpWSAdata As WSAdata
      Dim hHostent As Hostent, AddrList As Long
      Dim Address As Long, rIP As String
      Dim OptInfo As IP_OPTION_INFORMATION
      Dim EchoReply As IP_ECHO_REPLY

10    Ping = False

20    Call WSAStartup(&H101, lpWSAdata)

30    If GetHostByName(HostName + String$(64 - Len(HostName), 0)) <> SOCKET_ERROR Then
40        CopyMemory hHostent.h_name, ByVal GetHostByName(HostName + String$(64 - Len(HostName), 0)), Len(hHostent)
50        CopyMemory AddrList, ByVal hHostent.h_addr_list, 4
60        CopyMemory Address, ByVal AddrList, 4
70    End If

80    hFile = IcmpCreateFile()
90    If hFile = 0 Then
      '    MsgBox "Unable to Create File Handle"
100       Exit Function
110   End If

120   OptInfo.TTL = 255
130   If IcmpSendEcho(hFile, Address, String$(32, "A"), 32, OptInfo, EchoReply, Len(EchoReply) + 8, 2000) Then
140       rIP = CStr(EchoReply.Address(0)) + "." + CStr(EchoReply.Address(1)) + "." + CStr(EchoReply.Address(2)) + "." + CStr(EchoReply.Address(3))
150   Else
      '  MsgBox "Timeout"
160     Call IcmpCloseHandle(hFile)
170     Call WSACleanup
180     Exit Function
190   End If
200   If EchoReply.Status = 0 Then
      '    MsgBox "Reply from " + HostName + " (" + rIP + ") recieved after " + Trim$(CStr(EchoReply.RoundTripTime)) + "ms"
210   Else
      '    MsgBox "Failure ..."
220   End If

230   Call IcmpCloseHandle(hFile)
240   Call WSACleanup

250   Ping = True

End Function


