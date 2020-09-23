Attribute VB_Name = "IsNetConneted"
Option Explicit
Public Declare Function RasEnumConnections Lib "rasapi32.dll" Alias "RasEnumConnectionsA" (lpRasConn As Any, lpcb As Long, lpcConnections As Long) As Long
Public Declare Function RasHangUp Lib "rasapi32.dll" Alias "RasHangUpA" (ByVal hRasConn As Long) As Long
Public Declare Function InternetGetConnectedState Lib "wininet.dll" (ByRef lpdwFlags As Long, ByVal dwReserved As Long) As Long

' Remote Access Services (RAS) APIs.
Public Const RAS_MAXENTRYNAME As Integer = 256
Public Const RAS_MAXDEVICETYPE As Integer = 16
Public Const RAS_MAXDEVICENAME As Integer = 128
Public Const RAS_RASCONNSIZE As Integer = 412

Public Type RasEntryName
    dwSize As Long
    szEntryName(RAS_MAXENTRYNAME) As Byte
End Type

Public Type RasConn
    dwSize As Long
    hRasConn As Long
    szEntryName(RAS_MAXENTRYNAME) As Byte
    szDeviceType(RAS_MAXDEVICETYPE) As Byte
    szDeviceName(RAS_MAXDEVICENAME) As Byte
End Type
' some handy little variables for dealing with the registry
Public Const HKEY_CLASSES_ROOT = &H80000000
Public Const HKEY_CURRENT_USER = &H80000001
Public Const HKEY_LOCAL_MACHINE = &H80000002
Public Const HKEY_USERS = &H80000003
Public Const HKEY_PERFORMANCE_DATA = &H80000004
Public Const HKEY_CURRENT_CONFIG = &H80000005
Public Const HKEY_DYN_DATA = &H80000006
Public Const ERROR_SUCCESS = 0&
Public Const APINULL = 0&
Public Const MAX_STRING_LENGTH As Integer = 256

Public ReturnCode As Long
'Local system uses a modem to connect to the Internet.
Public Const INTERNET_CONNECTION_MODEM As Long = &H1
'Local system uses a LAN to connect to the Internet.
Public Const INTERNET_CONNECTION_LAN As Long = &H2
'Local system uses a proxy server to connect to the Internet.
Public Const INTERNET_CONNECTION_PROXY As Long = &H4
'No longer used.
Public Const INTERNET_CONNECTION_MODEM_BUSY As Long = &H8
Public Const INTERNET_RAS_INSTALLED As Long = &H10
Public Const INTERNET_CONNECTION_OFFLINE As Long = &H20
Public Const INTERNET_CONNECTION_CONFIGURED As Long = &H40

Public Sub HangUp()
    'Submitted by: J. Gerard Olszowiec (modified) Source:Newsgroup/PSC search for "detect dial up"
    Dim i As Long
    Dim lpRasConn(255) As RasConn
    Dim lpcb As Long
    Dim lpcConnections As Long
    Dim hRasConn As Long
    lpRasConn(0).dwSize = RAS_RASCONNSIZE
    lpcb = RAS_MAXENTRYNAME * lpRasConn(0).dwSize
    lpcConnections = 0
    ReturnCode = RasEnumConnections(lpRasConn(0), lpcb, lpcConnections)
    ' Drop ALL the connections that match the currect connections name.
    If ReturnCode = ERROR_SUCCESS Then
        For i = 0 To lpcConnections - 1
            If Trim(ByteToString(lpRasConn(i).szEntryName)) = Trim(gstrISPName) Then
                hRasConn = lpRasConn(i).hRasConn
                ReturnCode = RasHangUp(ByVal hRasConn)
            End If
        Next i
    End If
    ' It usually takes about 3 seconds to drop the connection.
    While Connected_To_ISP
        DoEvents
    Wend
End Sub

Public Function IsNetConnectViaLAN() As Boolean
   Dim dwflags As Long
   Call InternetGetConnectedState(dwflags, 0&)
   IsNetConnectViaLAN = dwflags And INTERNET_CONNECTION_LAN
End Function

Public Function IsNetConnectViaModem() As Boolean
   Dim dwflags As Long
   Call InternetGetConnectedState(dwflags, 0&)
   IsNetConnectViaModem = dwflags And INTERNET_CONNECTION_MODEM
End Function

Public Function IsNetConnectViaProxy() As Boolean
   Dim dwflags As Long
   Call InternetGetConnectedState(dwflags, 0&)
   IsNetConnectViaProxy = dwflags And INTERNET_CONNECTION_PROXY
End Function

Public Function IsNetConnectOnline() As Boolean
   'IsNetConnectOnline = InternetGetConnectedState(0&, 0&)
   If Connected_To_ISP Then
        IsNetConnectOnline = True
   End If
End Function

Public Function IsNetRASInstalled() As Boolean
   Dim dwflags As Long
   Call InternetGetConnectedState(dwflags, 0&)
   IsNetRASInstalled = dwflags And INTERNET_RAS_INSTALLED
End Function

Public Function ByteToString(bytString() As Byte) As String
    Dim i As Integer
    ByteToString = ""
    i = 0
    While bytString(i) = 0&
        ByteToString = ByteToString & Chr(bytString(i))
        i = i + 1
    Wend
End Function

