Attribute VB_Name = "Winfuncs"
Option Explicit
Public Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Public Declare Function WaitForSingleObject Lib "kernel32" (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long
Public Declare Function GetExitCodeProcess Lib "kernel32" (ByVal hProcess As Long, lpExitCode As Long) As Long
Public Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Public Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Public Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpString As Any, ByVal lpFileName As String) As Long
Public Declare Function CloseWindow Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Public Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long
Public Declare Function EnumWindows Lib "user32" (ByVal lpEnumFunc As Long, ByVal lParam As Long) As Long
Public Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Public Declare Function IsWindowVisible Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function GetParent Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hwnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Public Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, ByVal lpvParam As Any, ByVal fuWinIni As Long) As Long
Public Declare Function InternetAutodial Lib "wininet.dll" (ByVal dwflags As Long, ByVal dwReserved As Long) As Long
Public Declare Function InternetAutodialHangup Lib "wininet.dll" (ByVal dwReserved As Long) As Long

Global Const PROCESS_ALL_ACCESS& = &H1F0FFF 'some process handling stuff
Global Const STILL_ACTIVE& = &H103&
Global Const INFINITE& = &HFFFF
Global Const SWP_NOMOVE = 2                 'some window constants here
Global Const SWP_NOSIZE = 1
Global Const FLAGS = SWP_NOMOVE Or SWP_NOSIZE
Global Const HWND_TOPMOST = -1
Global Const HWND_NOTOPMOST = -2
Global Const WM_CLOSE = &H10
Global Const SW_SHOWDEFAULT = 10
Global Const SYNCHRONIZE = 1048576
Global Const NORMAL_PRIORITY_CLASS = &H20&

Global Const SPI_SETDESKWALLPAPER = 20      'settings for: wallpaper/font smoothing
Global Const SPI_SETDESKPATTERN = 21        '(re)sets desktop pattern
Global Const SPIF_UPDATEINIFILE = &H1       'used when setting On or Off
Global Const SPIF_SENDWININICHANGE = &H2    'used when setting On or Off
Global Const SPI_GETFONTSMOOTHING = 74      'used when reading if On or Off
Global Const SPI_SETFONTSMOOTHING = 75      'used when setting On or Off
Global Const SPI_GETWINDOWSEXTENSION = 92   'used to find out if Plus!-extensions are installed

Global Const INTERNET_AUTODIAL_FORCE_ONLINE = 1
Global Const INTERNET_AUTODIAL_FORCE_UNATTENDED = 2

Global lResult As Long
Global bol_fsmooth As Boolean               'stores state of font smoothing (on/off) before messing with it

Dim sPattern As String, hFind As Long
'Copyright: Arkadiy Olovyannikov Code: FindWindowWild Location: http://www.planetsourcecode.com
Public Function GetHWnd(sWild As String, Optional bMatchCase As Boolean = True) As Long
    sPattern = sWild
    If Not bMatchCase Then sPattern = UCase(sPattern)
    EnumWindows AddressOf EnumWinProc, bMatchCase
    GetHWnd = hFind
End Function

Function EnumWinProc(ByVal hwnd As Long, ByVal lParam As Long) As Long
    Dim k As Long, sName As String
    If IsWindowVisible(hwnd) And GetParent(hwnd) = 0 Then
        sName = Space$(128)
        k = GetWindowText(hwnd, sName, 128)
        If k > 0 Then
            sName = Left$(sName, k)
            If lParam = 0 Then sName = UCase(sName)
            If InStr(1, sName, sPattern, vbTextCompare) > 0 Then
                hFind = hwnd
                EnumWinProc = 0
                Exit Function
            End If
        End If
    End If
    EnumWinProc = 1
End Function

Public Function Open_Browser(strURL As String, lngHwnd As Long)
    Open_Browser = ShellExecute(lngHwnd, vbNullString, strURL, vbNullString, "c:\", SW_SHOWDEFAULT)
End Function

Public Function ShellWait(ByVal lProgID As Long) As Long
    ' Wait until proggie exit code <>  STILL_ACTIVE&
    Dim lExitCode As Long
    Dim hdlProg As Long
    ' Get proggie handle
    hdlProg = OpenProcess(PROCESS_ALL_ACCESS, False, lProgID)
    ' Get current proggie exit code
    GetExitCodeProcess hdlProg, lExitCode
    Do While lExitCode = STILL_ACTIVE&
        DoEvents
        GetExitCodeProcess hdlProg, lExitCode
    Loop
    CloseHandle hdlProg
    ShellWait = lExitCode
End Function
