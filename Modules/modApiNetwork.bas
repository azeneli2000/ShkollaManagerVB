Attribute VB_Name = "modApiNetwork"
Option Explicit

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Copyright ©1996-2007 VBnet, Randy Birch, All Rights Reserved.
' Some pages may also contain other copyrights by the author.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Distribution: You can freely use this code in your own
'               applications, but you may not reproduce
'               or publish this code on any web site,
'               online service, or distribute as source
'               on any media without express permission.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'an array to hold the description strings for list3
Private desc() As String

Public Const MAX_PREFERRED_LENGTH As Long = -1
Public Const NERR_SUCCESS As Long = 0
Private Const ERROR_MORE_DATA As Long = 234
Private Const FORMAT_MESSAGE_MAX_WIDTH_MASK = &HFF

Private Const SV_TYPE_WORKSTATION As Long = &H1
Private Const SV_TYPE_SERVER As Long = &H2
Private Const SV_TYPE_SQLSERVER As Long = &H4
Private Const SV_TYPE_DOMAIN_CTRL As Long = &H8
Private Const SV_TYPE_DOMAIN_BAKCTRL As Long = &H10
Private Const SV_TYPE_TIME_SOURCE As Long = &H20
Private Const SV_TYPE_AFP As Long = &H40
Private Const SV_TYPE_NOVELL As Long = &H80
Private Const SV_TYPE_DOMAIN_MEMBER As Long = &H100
Private Const SV_TYPE_PRINTQ_SERVER As Long = &H200
Private Const SV_TYPE_DIALIN_SERVER As Long = &H400
Private Const SV_TYPE_XENIX_SERVER As Long = &H800
Private Const SV_TYPE_SERVER_UNIX As Long = SV_TYPE_XENIX_SERVER
Private Const SV_TYPE_NT As Long = &H1000
Private Const SV_TYPE_WFW As Long = &H2000
Private Const SV_TYPE_SERVER_MFPN As Long = &H4000
Private Const SV_TYPE_SERVER_NT As Long = &H8000&
Private Const SV_TYPE_POTENTIAL_BROWSER As Long = &H10000
Private Const SV_TYPE_BACKUP_BROWSER As Long = &H20000
Private Const SV_TYPE_MASTER_BROWSER As Long = &H40000
Private Const SV_TYPE_DOMAIN_MASTER As Long = &H80000
Private Const SV_TYPE_SERVER_OSF As Long = &H100000
Private Const SV_TYPE_SERVER_VMS As Long = &H200000
Private Const SV_TYPE_WINDOWS As Long = &H400000 'Windows95 and above
Private Const SV_TYPE_DFS As Long = &H800000 'Root of a DFS tree
Private Const SV_TYPE_CLUSTER_NT As Long = &H1000000 'NT Cluster
Private Const SV_TYPE_TERMINALSERVER As Long = &H2000000 'Terminal Server
Private Const SV_TYPE_DCE As Long = &H10000000 'IBM DSS (Directory and Security Services) or equivalent
Private Const SV_TYPE_ALTERNATE_XPORT As Long = &H20000000 'Return list for alternate transport
Private Const SV_TYPE_LOCAL_LIST_ONLY As Long = &H40000000 'Return local list only
Private Const SV_TYPE_DOMAIN_ENUM As Long = &H80000000 'Return domain list only
Private Const SV_TYPE_ALL As Long = &HFFFFFFFF

Private Const PLATFORM_ID_DOS As Long = 300
Private Const PLATFORM_ID_OS2 As Long = 400
Private Const PLATFORM_ID_NT As Long = 500
Private Const PLATFORM_ID_OSF As Long = 600
Private Const PLATFORM_ID_VMS As Long = 700

Private Const LB_SETTABSTOPS As Long = &H192
Private Const FORMAT_MESSAGE_FROM_SYSTEM As Long = &H1000
Private Const FORMAT_MESSAGE_IGNORE_INSERTS As Long = &H200

'Mask applied to sv*_version_major in
'order to obtain the major version number
Public Const MAJOR_VERSION_MASK As Long = &HF


 Public Const WS_VERSION_REQD = &H101
   Public Const WS_VERSION_MAJOR = WS_VERSION_REQD \ &H100 And &HFF&
   Public Const WS_VERSION_MINOR = WS_VERSION_REQD And &HFF&
   Public Const MIN_SOCKETS_REQD = 1
   Public Const SOCKET_ERROR = -1
   Public Const WSADescription_Len = 256
   Public Const WSASYS_Status_Len = 128
 
   Public Type HOSTENT
       hName As Long
       hAliases As Long
       hAddrType As Integer
       hLength As Integer
       hAddrList As Long
   End Type
 
   Public Type WSADATA
       wversion As Integer
       wHighVersion As Integer
       szDescription(0 To WSADescription_Len) As Byte
       szSystemStatus(0 To WSASYS_Status_Len) As Byte
       iMaxSockets As Integer
       iMaxUdpDg As Integer
       lpszVendorInfo As Long
   End Type
   
Public Type SERVER_INFO_100
  sv100_platform_id  As Long
  sv100_name As Long
End Type

Public Type SERVER_INFO_101
  sv101_platform_id  As Long
  sv101_name As Long
  sv101_version_major As Long
  sv101_version_minor As Long
  sv101_type As Long
  sv101_comment As Long
End Type

Public Declare Function NetServerEnum Lib "Netapi32" _
  (ByVal servername As Long, _
   ByVal level As Long, _
   buf As Any, _
   ByVal prefmaxlen As Long, _
   entriesread As Long, _
   totalentries As Long, _
   ByVal servertype As Long, _
   ByVal domain As Long, _
   resume_handle As Long) As Long

Public Declare Function NetServerGetInfo Lib "Netapi32" _
  (ByVal servername As Long, _
   ByVal level As Long, _
   bufptr As Any) As Long

Public Declare Function NetApiBufferFree Lib "netapi32.dll" _
   (ByVal Buffer As Long) As Long

Public Declare Sub CopyMemory Lib "KERNEL32" _
   Alias "RtlMoveMemory" _
  (pTo As Any, uFrom As Any, _
   ByVal lSize As Long)
   
Public Declare Function lstrlenW Lib "KERNEL32" _
  (ByVal lpString As Long) As Long

Public Declare Function SendMessage Lib "user32" _
   Alias "SendMessageA" _
   (ByVal hWnd As Long, _
    ByVal wMsg As Long, _
    ByVal wParam As Long, _
    lParam As Any) As Long

Public Declare Function FormatMessage Lib "KERNEL32" _
     Alias "FormatMessageA" _
    (ByVal dwFlags As Long, _
     lpSource As Long, _
     ByVal dwMessageId As Long, _
     ByVal dwLanguageId As Long, _
     ByVal lpBuffer As String, _
     ByVal nSize As Long, _
     Arguments As Any) As Long
     
Public Declare Function WSAGetLastError Lib "WSOCK32.DLL" () As Long
   Public Declare Function WSAStartup Lib "WSOCK32.DLL" (ByVal wVersionRequired&, lpWSAData As WSADATA) As Long
   Public Declare Function WSACleanup Lib "WSOCK32.DLL" () As Long
   Public Declare Function gethostname Lib "WSOCK32.DLL" (ByVal hostname$, ByVal HostLen&) As Long
   Public Declare Function gethostbyname Lib "WSOCK32.DLL" (ByVal hostname$) As Long
 
   Public Declare Sub RtlMoveMemory Lib "KERNEL32" (hpvDest As Any, ByVal hpvSource&, ByVal cbCopy&)


Public Function GetPointerToByteStringW(ByVal dwData As Long) As String
  
   Dim tmp() As Byte
   Dim tmplen As Long
   
   If dwData <> 0 Then
   
      tmplen = lstrlenW(dwData) * 2
      
      If tmplen <> 0 Then
      
         ReDim tmp(0 To (tmplen - 1)) As Byte
         CopyMemory tmp(0), ByVal dwData, tmplen
         GetPointerToByteStringW = tmp
         
     End If
     
   End If
    
End Function


  
 
   Public Function hibyte(ByVal wParam As Integer)
 
       hibyte = wParam \ &H100& And &HFF&
 
   End Function
 
   Public Function lobyte(ByVal wParam As Integer)
 
       lobyte = wParam And &HFF&
 
   End Function
 
   Public Sub SocketsInitialize()
   Dim WSAD As WSADATA
   Dim iReturn As Integer
   Dim sLowByte As String, sHighByte As String, sMsg As String
 
       iReturn = WSAStartup(WS_VERSION_REQD, WSAD)
 
       If iReturn <> 0 Then
           MsgBox "Winsock.dll nuk përgjigjet."
           End
       End If
 
       If lobyte(WSAD.wversion) < WS_VERSION_MAJOR Or (lobyte(WSAD.wversion) = WS_VERSION_MAJOR And hibyte(WSAD.wversion) < WS_VERSION_MINOR) Then
           sHighByte = Trim$(str$(hibyte(WSAD.wversion)))
           sLowByte = Trim$(str$(lobyte(WSAD.wversion)))
           sMsg = "Versioni i Windows Sockets " & sLowByte & "." & sHighByte
           sMsg = sMsg & " nuk suportohet winsock.dll "
           MsgBox sMsg
           End
       End If
 
       If WSAD.iMaxSockets < MIN_SOCKETS_REQD Then
           sMsg = "Ky program kërkon të paktën "
           sMsg = sMsg & Trim$(str$(MIN_SOCKETS_REQD)) & " socket për suport."
           MsgBox sMsg
           End
       End If
 
   End Sub
 
   Public Sub SocketsCleanup()
   Dim lReturn As Long
 
       lReturn = WSACleanup()
 
       If lReturn <> 0 Then
           MsgBox "Gabim në socket nr " & Trim$(str$(lReturn)) & " gjatë pastrimit"
           End
       End If
 
   End Sub

    Public Function GjejEmerWorkstation() As String
    Dim hostname As String * 256
    Dim hostent_addr As Long
    Dim host As HOSTENT
    Dim hostip_addr As Long
    Dim temp_ip_address() As Byte
    Dim i As Integer
    Dim ip_address As String
 
       If gethostname(hostname, 256) = SOCKET_ERROR Then
           MsgBox "Windows Sockets error " & str(WSAGetLastError())
       Else
           hostname = Trim$(hostname)
       End If
 
       hostent_addr = gethostbyname(hostname)
 
       If hostent_addr = 0 Then
           MsgBox "Winsock.dll is not responding."
       End If
 
       RtlMoveMemory host, hostent_addr, LenB(host)
       RtlMoveMemory hostip_addr, host.hAddrList, 4
 
       ReDim temp_ip_address(1 To host.hLength)
       RtlMoveMemory temp_ip_address(1), hostip_addr, host.hLength
 
       For i = 1 To host.hLength
           ip_address = ip_address & temp_ip_address(i) & "."
       Next
       ip_address = Mid$(ip_address, 1, Len(ip_address) - 1)
        Dim emri As String
        emri = ""
       For i = 1 To Len(hostname)
        If (Asc(Mid(hostname, i, 1)) <> 0) Then
            emri = emri + Mid(hostname, i, 1)
        Else
            Exit For
        End If
       Next i
       GjejEmerWorkstation = emri
    End Function

