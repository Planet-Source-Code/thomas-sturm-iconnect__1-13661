Attribute VB_Name = "RegistryRead"
Public Declare Function RegSetValue Lib "advapi32.dll" Alias "RegSetValueA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal dwType As Long, ByVal lpData As String, ByVal cbData As Long) As Long
Public Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Public Declare Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Public Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long         ' Note that if you declare the lpData parameter as String, you must pass it By Value.

Public gstrISPName As String
Public ReturnCode As Long
Public hKey As Long
Public lpSubKey As String
Public phkResult As Long
Public lpValueName As String
Public lpReserved As Long
Public lpType As Long
Public lpData As String
Public lpcbData As Long
Public Function Get_RegStringVal(key, subkey, valuename) As String
    Get_RegStringVal = ""
    lpSubKey = subkey
    ReturnCode = RegOpenKey(key, lpSubKey, phkResult)
    If ReturnCode = ERROR_SUCCESS Then
        hKey = phkResult
        lpValueName = valuename
        lpReserved = APINULL
        lpType = APINULL
        lpData = APINULL
        lpcbData = APINULL
        ReturnCode = RegQueryValueEx(hKey, lpValueName, lpReserved, lpType, ByVal lpData, lpcbData)
        lpData = String(lpcbData, 0)
        ReturnCode = RegQueryValueEx(hKey, lpValueName, lpReserved, lpType, ByVal lpData, lpcbData)
        If ReturnCode = ERROR_SUCCESS Then
            Get_RegStringVal = Left(lpData, lpcbData - 1)
        End If
        RegCloseKey (hKey)
    End If
End Function
Public Function Connected_To_ISP() As Boolean
    'Submitted by: J. Gerard Olszowiec Source:Newsgroup/PSC search for "detect dial up"
    Dim hKey As Long
    Dim lpSubKey As String
    Dim phkResult As Long
    Dim lpValueName As String
    Dim lpReserved As Long
    Dim lpType As Long
    Dim lpData As Long
    Dim lpcbData As Long
    Connected_To_ISP = False
    lpSubKey = "System\CurrentControlSet\Services\RemoteAccess"
    ReturnCode = RegOpenKey(HKEY_LOCAL_MACHINE, lpSubKey, phkResult)
    If ReturnCode = ERROR_SUCCESS Then
        hKey = phkResult
        lpValueName = "Remote Connection"
        lpReserved = APINULL
        lpType = APINULL
        lpData = APINULL
        lpcbData = APINULL
        ReturnCode = RegQueryValueEx(hKey, lpValueName, lpReserved, lpType, ByVal lpData, lpcbData)
        lpcbData = Len(lpData)
        ReturnCode = RegQueryValueEx(hKey, lpValueName, lpReserved, lpType, lpData, lpcbData)
        If ReturnCode = ERROR_SUCCESS Then
            If lpData = 0 Then
                ' Not Connected
            Else
                ' Connected
                Connected_To_ISP = True
            End If
        End If
        RegCloseKey (hKey)
    End If
End Function




